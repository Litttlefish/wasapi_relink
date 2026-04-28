// #[cfg(test)]
// mod config_test;

use flexi_logger::*;
use log::*;
use retour::GenericDetour;
use rtrb::{Consumer, Producer, RingBuffer};
use serde::*;
use std::cell::{Cell, OnceCell, UnsafeCell};
use std::collections::HashMap;
use std::mem::transmute;
use std::num::NonZero;
use std::os::raw::c_void;
use std::path::Path;
use std::slice::from_raw_parts_mut;
use std::sync::{LazyLock, Once, OnceLock, atomic::*};

use windows::{
    Win32::{
        Foundation::*,
        Media::Audio::*,
        System::Com::{StructuredStorage::*, *},
        System::LibraryLoader::{GetModuleHandleW, GetProcAddress},
        System::Threading::*,
        UI::Shell::PropertiesSystem::IPropertyStore,
    },
    core::{Result as WinResult, *},
};

static LOGGER_HANDLE: OnceLock<LoggerHandle> = OnceLock::new();

#[derive(Debug, Default, Serialize, Deserialize, PartialEq, Eq, Clone, Copy)]
enum ConfigLogLevel {
    Trace,
    Debug,
    #[default]
    Info,
    Warn,
    Error,
    Never,
}

impl From<ConfigLogLevel> for LevelFilter {
    fn from(value: ConfigLogLevel) -> Self {
        match value {
            ConfigLogLevel::Trace => Self::Trace,
            ConfigLogLevel::Debug => Self::Debug,
            ConfigLogLevel::Info => Self::Info,
            ConfigLogLevel::Warn => Self::Warn,
            ConfigLogLevel::Error => Self::Error,
            ConfigLogLevel::Never => Self::Off,
        }
    }
}

#[derive(Debug, Default)]
enum ConfigSource {
    #[default]
    Success,
    NoParse,
    NoFile,
}

#[derive(Debug, Serialize, Deserialize, Default)]
#[serde(default)]
struct RedirectConfig {
    log_path: Option<Box<Path>>,
    log_level: ConfigLogLevel,
    only_log_stdout: bool,
    playback: ClientConfig,
    capture: ClientConfig,
    #[serde(skip)]
    source: ConfigSource,
}
impl RedirectConfig {
    fn load() -> Self {
        std::fs::read_to_string("redirect_config.toml").map_or_else(
            |_| Self::new_with_source(ConfigSource::NoFile),
            |str| {
                toml::from_str::<RedirectConfig>(&str)
                    .unwrap_or_else(|_| Self::new_with_source(ConfigSource::NoParse))
            },
        )
    }
    fn new_with_source(source: ConfigSource) -> Self {
        Self {
            source,
            ..Self::default()
        }
    }
    #[inline]
    fn get(&self, dataflow: DeviceDataFlow) -> &ClientConfig {
        match dataflow {
            DeviceDataFlow::Capture => &self.capture,
            DeviceDataFlow::Playback => &self.playback,
        }
    }
}

#[derive(Debug, Serialize, Deserialize, Default)]
#[serde(default)]
struct ClientConfig {
    target_period_hus: u32,
    ring_buffer_len: HashMap<u32, NonZero<u32>>,
    target_buffer_len: HashMap<u32, NonZero<u32>>,
    compat_buffer_dur_hns: HashMap<u32, i64>,
    force_period: bool,
    mode: ClientMode,
    raw: bool,
}
impl ClientConfig {
    fn target_buf_len(&self, info: &Shared3Info) -> Option<u32> {
        self.target_buffer_len
            .get(&info.samplerate)
            .map(|l| l.get().next_multiple_of(info.fundamental))
    }
    fn ring_buf_len(&self, info: &Shared3Info) -> Option<u32> {
        self.ring_buffer_len.get(&info.samplerate).map(|l| {
            info.current_period
                .max(l.get().next_multiple_of(info.fundamental))
        })
    }
    fn compat_buf_len(&self, info: &Shared3Info) -> Option<i64> {
        self.compat_buffer_dur_hns.get(&info.samplerate).copied()
    }
}

#[derive(Debug, Serialize, Deserialize, Default, PartialEq, Eq)]
enum ClientMode {
    #[default]
    Normal,
    Compat,
    Ringbuf,
    Bypass,
}
impl std::fmt::Display for ClientMode {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "{}",
            match self {
                Self::Normal => "normal",
                Self::Compat => "compat",
                Self::Ringbuf => "ringbuf",
                Self::Bypass => "bypass",
            }
        )
    }
}

#[derive(Clone, Copy)]
enum AudioAlign {
    Pow2(usize),
    Normal(usize),
}
impl AudioAlign {
    #[inline(always)]
    fn new(align: u16) -> Self {
        if align.is_power_of_two() {
            Self::Pow2(align.trailing_zeros() as usize)
        } else {
            Self::Normal(align as usize)
        }
    }
    #[inline(always)]
    fn bytes_to_frames(&self, bytes: usize) -> usize {
        match self {
            Self::Pow2(shift) => bytes >> *shift,
            Self::Normal(align) => bytes / *align,
        }
    }
    #[inline(always)]
    fn frames_to_bytes(&self, frames: usize) -> usize {
        match self {
            Self::Pow2(shift) => frames << *shift,
            Self::Normal(align) => frames * *align,
        }
    }
}

macro_rules! trace_tagged {
    (@$self:ident, $($arg:tt)+) => { trace_tagged!($self.info.tag, $($arg)+) };
    ($tag:expr, $($arg:tt)+) => { trace!(target: $tag.as_ref(), $($arg)+) };
}
macro_rules! debug_tagged {
    (@$self:ident, $($arg:tt)+) => { debug_tagged!($self.info.tag, $($arg)+) };
    ($tag:expr, $($arg:tt)+) => { debug!(target: $tag.as_ref(), $($arg)+) };
}
macro_rules! info_tagged {
    (@$self:ident, $($arg:tt)+) => { info_tagged!($self.info.tag, $($arg)+) };
    ($tag:expr, $($arg:tt)+) => { info!(target: $tag.as_ref(), $($arg)+) };
}
macro_rules! warn_tagged {
    (@$self:ident, $($arg:tt)+) => { warn_tagged!($self.info.tag, $($arg)+) };
    ($tag:expr, $($arg:tt)+) => { warn!(target: $tag.as_ref(), $($arg)+) };
}
macro_rules! error_tagged {
    (@$self:ident, $($arg:tt)+) => { error_tagged!($self.info.tag, $($arg)+) };
    ($tag:expr, $($arg:tt)+) => { error!(target: $tag.as_ref(), $($arg)+) };
}

static CONFIG: LazyLock<RedirectConfig> = LazyLock::new(RedirectConfig::load);

static CLIENT_ID: (AtomicU16, AtomicU16) = (AtomicU16::new(0), AtomicU16::new(0));

type FnCoCreateInstance = unsafe extern "system" fn(
    *const GUID,
    *mut c_void,
    CLSCTX,
    *const GUID,
    *mut *mut c_void,
) -> HRESULT;

type FnCoCreateInstanceEx = unsafe extern "system" fn(
    *const GUID,
    *mut c_void,
    CLSCTX,
    *const COSERVERINFO,
    u32,
    *mut MULTI_QI,
) -> HRESULT;

const KEYWORDS: &[&str] = &["[GAME]", "[SK]"];

static CO_CREATE: LazyLock<(
    GenericDetour<FnCoCreateInstance>,
    GenericDetour<FnCoCreateInstanceEx>,
)> = LazyLock::new(|| unsafe {
    link!("combase.dll" "system" fn CoCreateInstance(_ : *const GUID, _ : *mut c_void, _ : CLSCTX, _ : *const GUID, _ : *mut *mut c_void) -> HRESULT);
    link!("combase.dll" "system" fn CoCreateInstanceEx(_ : *const GUID, _ : *mut c_void, _ : CLSCTX, _ : *const COSERVERINFO, _ : u32, _ : *mut MULTI_QI) -> HRESULT);
    let (func, funcex): (FnCoCreateInstance, FnCoCreateInstanceEx) =
        transmute(GetModuleHandleW(w!("combase")).map_or(
            {
                (
                    CoCreateInstance as *mut c_void,
                    CoCreateInstanceEx as *mut c_void,
                )
            },
            |hmodule| {
                (
                    GetProcAddress(hmodule, s!("CoCreateInstance")).unwrap() as *mut c_void,
                    GetProcAddress(hmodule, s!("CoCreateInstanceEx")).unwrap() as *mut c_void,
                )
            },
        ));
    (
        GenericDetour::new(func, hooked_cocreateinstance).unwrap(),
        GenericDetour::new(funcex, hooked_cocreateinstanceex).unwrap(),
    )
});

unsafe extern "system" fn hooked_cocreateinstance(
    rclsid: *const GUID,
    punkouter: *mut c_void,
    dwclscontext: CLSCTX,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        let ret = CO_CREATE.0.call(rclsid, punkouter, dwclscontext, riid, ppv);
        if *riid == IMMDeviceEnumerator::IID && ret.is_ok() {
            LOGGER_HANDLE.get_or_init(setup);
            if let Ok(thread_desc) = GetThreadDescription(GetCurrentThread())
                && !thread_desc.is_empty()
                && let Ok(name) = thread_desc.to_string()
                && KEYWORDS.iter().any(|keyword| name.contains(keyword))
            {
                trace!("Skipping SpecialK CoCreateInstance calls, thread name: {name}");
            } else {
                debug!("Intercepted IMMDeviceEnumerator creation via CoCreateInstance");
                let proxy_enumerator: IMMDeviceEnumerator =
                    RedirectDeviceEnumerator::new(IMMDeviceEnumerator::from_raw(*ppv)).into();
                *ppv = proxy_enumerator.into_raw();
            }
        }
        ret
    }
}

unsafe extern "system" fn hooked_cocreateinstanceex(
    clsid: *const GUID,
    punkouter: *mut c_void,
    dwclsctx: CLSCTX,
    pserverinfo: *const COSERVERINFO,
    dwcount: u32,
    presults: *mut MULTI_QI,
) -> HRESULT {
    unsafe {
        let hr = CO_CREATE
            .1
            .call(clsid, punkouter, dwclsctx, pserverinfo, dwcount, presults);
        if *clsid == MMDeviceEnumerator && hr.is_ok() {
            LOGGER_HANDLE.get_or_init(setup);
            if let Ok(thread_desc) = GetThreadDescription(GetCurrentThread())
                && !thread_desc.is_empty()
                && let Ok(name) = thread_desc.to_string()
                && KEYWORDS.iter().any(|keyword| name.contains(keyword))
            {
                trace!("Skipping SpecialK CoCreateInstanceEx calls, thread name: {name}")
            } else {
                for qi in from_raw_parts_mut(presults, dwcount as usize) {
                    if *qi.pIID == IMMDeviceEnumerator::IID && qi.hr.is_ok() {
                        debug!("Intercepted IMMDeviceEnumerator via CoCreateInstanceEx");
                        let proxy_enumerator: IMMDeviceEnumerator =
                            RedirectDeviceEnumerator::new(IMMDeviceEnumerator::from_raw(
                                qi.pItf.take().unwrap_unchecked().into_raw(),
                            ))
                            .into();
                        _ = qi.pItf.insert(proxy_enumerator.into())
                    }
                }
            }
        }
        hr
    }
}

#[repr(transparent)]
#[implement(IMMDeviceEnumerator)]
struct RedirectDeviceEnumerator {
    inner: IMMDeviceEnumerator,
}
impl RedirectDeviceEnumerator {
    pub fn new(inner: IMMDeviceEnumerator) -> Self {
        Self { inner }
    }
}
impl IMMDeviceEnumerator_Impl for RedirectDeviceEnumerator_Impl {
    fn EnumAudioEndpoints(
        &self,
        dataflow: EDataFlow,
        dwstatemask: DEVICE_STATE,
    ) -> WinResult<IMMDeviceCollection> {
        debug!(
            "DeviceEnumerator::EnumAudioEndpoints requested on flow {}",
            dataflow.0
        );
        Ok(RedirectDeviceCollection {
            inner: unsafe { self.inner.EnumAudioEndpoints(dataflow, dwstatemask)? },
        }
        .into())
    }

    fn GetDefaultAudioEndpoint(&self, dataflow: EDataFlow, role: ERole) -> WinResult<IMMDevice> {
        debug!(
            "DeviceEnumerator::GetDefaultAudioEndpoint requested on flow {}",
            dataflow.0
        );
        Ok(
            RedirectDevice::new(unsafe { self.inner.GetDefaultAudioEndpoint(dataflow, role)? })
                .into(),
        )
    }

    fn GetDevice(&self, pwstrid: &PCWSTR) -> WinResult<IMMDevice> {
        info!("DeviceEnumerator::GetDevice called, wrapping");
        Ok(RedirectDevice::new(unsafe { self.inner.GetDevice(*pwstrid)? }).into())
    }

    fn RegisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> WinResult<()> {
        trace!("DeviceEnumerator::RegisterEndpointNotificationCallback called");
        unsafe {
            self.inner
                .RegisterEndpointNotificationCallback(pclient.as_ref())
        }
    }

    fn UnregisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> WinResult<()> {
        trace!("DeviceEnumerator::UnregisterEndpointNotificationCallback called");
        unsafe {
            self.inner
                .UnregisterEndpointNotificationCallback(pclient.as_ref())
        }
    }
}

#[repr(transparent)]
#[implement(IMMDeviceCollection)]
struct RedirectDeviceCollection {
    inner: IMMDeviceCollection,
}
impl IMMDeviceCollection_Impl for RedirectDeviceCollection_Impl {
    fn GetCount(&self) -> WinResult<u32> {
        trace!("DeviceCollection::GetCount called");
        unsafe { self.inner.GetCount() }
    }

    fn Item(&self, ndevice: u32) -> WinResult<IMMDevice> {
        debug!("DeviceCollection::Item retrieved device {ndevice}");
        Ok(RedirectDevice::new(unsafe { self.inner.Item(ndevice)? }).into())
    }
}

macro_rules! impl_boilerplate {
    (IAudioClient1) => {
        fn GetMixFormat(&self) -> WinResult<*mut WAVEFORMATEX> {
            info_tagged!(@self, "GetMixFormat called");
            unsafe { self.inner.GetMixFormat() }
        }
        fn IsFormatSupported(
            &self,
            sharemode: AUDCLNT_SHAREMODE,
            pformat: *const WAVEFORMATEX,
            ppclosestmatch: *mut *mut WAVEFORMATEX,
        ) -> HRESULT {
            debug_tagged!(@self, "IsFormatSupported called");
            unsafe {
                self.inner
                    .IsFormatSupported(sharemode, pformat, Some(ppclosestmatch))
            }
        }
        fn GetStreamLatency(&self) -> WinResult<i64> {
            info_tagged!(@self, "GetStreamLatency called");
            unsafe { self.inner.GetStreamLatency() }
        }
    };
    (IAudioClient2) =>
    {
        fn IsOffloadCapable(&self, category: AUDIO_STREAM_CATEGORY) -> WinResult<BOOL> {
            info_tagged!(@self, "IsOffloadCapable called");
            unsafe { self.inner.IsOffloadCapable(category) }
        }
        fn GetBufferSizeLimits(
            &self,
            pformat: *const WAVEFORMATEX,
            beventdriven: BOOL,
            phnsminbufferduration: *mut i64,
            phnsmaxbufferduration: *mut i64,
        ) -> WinResult<()> {
            info_tagged!(@self, "GetBufferSizeLimits called");
            unsafe {
                self.inner.GetBufferSizeLimits(
                    pformat,
                    beventdriven.into(),
                    phnsminbufferduration,
                    phnsmaxbufferduration,
                )
            }
        }
    };
    (IAudioClient3)=> {
        fn GetSharedModeEnginePeriod(
            &self,
            pformat: *const WAVEFORMATEX,
            pdefaultperiodinframes: *mut u32,
            pfundamentalperiodinframes: *mut u32,
            pminperiodinframes: *mut u32,
            pmaxperiodinframes: *mut u32,
        ) -> WinResult<()> {
            info_tagged!(@self, "GetSharedModeEnginePeriod called");
            unsafe {
                self.inner.GetSharedModeEnginePeriod(
                    pformat,
                    pdefaultperiodinframes,
                    pfundamentalperiodinframes,
                    pminperiodinframes,
                    pmaxperiodinframes,
                )
            }
        }
        fn GetCurrentSharedModeEnginePeriod(
            &self,
            ppformat: *mut *mut WAVEFORMATEX,
            pcurrentperiodinframes: *mut u32,
        ) -> WinResult<()> {
            info_tagged!(@self, "GetCurrentSharedModeEnginePeriod called");
            unsafe {
                self.inner
                    .GetCurrentSharedModeEnginePeriod(ppformat, pcurrentperiodinframes)
            }
        }
    }
}

macro_rules! drop_boilerplate {
    ($struct:ty) => {
        impl Drop for $struct {
            fn drop(&mut self) {
                debug_tagged!(@self, "Client dropped")
            }
        }
    };
}

#[repr(transparent)]
#[implement(IMMDevice, IMMEndpoint)]
struct RedirectDevice {
    inner: IMMDevice,
}

impl RedirectDevice {
    pub fn new(inner: IMMDevice) -> Self {
        Self { inner }
    }
}

impl IMMDevice_Impl for RedirectDevice_Impl {
    fn Activate(
        &self,
        riid: *const GUID,
        dwclsctx: CLSCTX,
        pactivationparams: *const PROPVARIANT,
        ppinterface: *mut *mut c_void,
    ) -> WinResult<()> {
        unsafe {
            let iid = *riid;
            info!("Device::Activate called, iid: {iid:?}");
            match iid {
                IAudioClient::IID | IAudioClient2::IID | IAudioClient3::IID => {
                    let inner: IAudioClient3 = self
                        .inner
                        .Activate::<IAudioClient3>(dwclsctx, Some(pactivationparams))?;
                    let dataflow = self.inner.cast::<IMMEndpoint>()?.GetDataFlow()?.into();
                    let config = CONFIG.get(dataflow);
                    let tag = format!(
                        "{dataflow}-{}::{}",
                        match dataflow {
                            DeviceDataFlow::Playback => CLIENT_ID.0.fetch_add(1, Ordering::Relaxed),
                            DeviceDataFlow::Capture => CLIENT_ID.1.fetch_add(1, Ordering::Relaxed),
                        },
                        config.mode
                    );
                    info_tagged!(tag, "Client created");
                    let info = RedirectClientInfo::new(config, tag.into());
                    let proxy: IAudioClient3 = match config.mode {
                        ClientMode::Normal => RedirectAudioClient::new(inner, info).into(),
                        ClientMode::Compat => RedirectCompatAudioClient::new(
                            inner,
                            self.inner
                                .Activate::<IAudioClient3>(dwclsctx, Some(pactivationparams))?,
                            info,
                        )
                        .into(),
                        ClientMode::Ringbuf => match dataflow {
                            DeviceDataFlow::Playback => {
                                RedirectRingbufAudioClient::new(inner, info).into()
                            }
                            DeviceDataFlow::Capture => {
                                warn_tagged!(
                                    info.tag,
                                    "Ringbuf mode doesn't work with capture client, switching to compat mode"
                                );
                                RedirectCompatAudioClient::new(
                                    inner,
                                    self.inner.Activate::<IAudioClient3>(
                                        dwclsctx,
                                        Some(pactivationparams),
                                    )?,
                                    info,
                                )
                                .into()
                            }
                        },
                        ClientMode::Bypass => inner,
                    };
                    proxy.query(riid, ppinterface).ok()
                }
                _ => (self.inner.vtable().Activate)(
                    self.inner.as_raw(),
                    riid,
                    dwclsctx,
                    pactivationparams,
                    ppinterface,
                )
                .ok(),
            }
        }
    }

    fn OpenPropertyStore(&self, stgmaccess: STGM) -> WinResult<IPropertyStore> {
        debug!("Device::OpenPropertyStore called");
        unsafe { self.inner.OpenPropertyStore(stgmaccess) }
    }

    fn GetId(&self) -> WinResult<PWSTR> {
        debug!("Device::GetId called");
        unsafe { self.inner.GetId() }
    }

    fn GetState(&self) -> WinResult<DEVICE_STATE> {
        trace!("Device::GetState called");
        unsafe { self.inner.GetState() }
    }
}
impl IMMEndpoint_Impl for RedirectDevice_Impl {
    fn GetDataFlow(&self) -> WinResult<EDataFlow> {
        trace!("Device::GetDataFlow called");
        unsafe { self.inner.cast::<IMMEndpoint>()?.GetDataFlow() }
    }
}

#[derive(Debug, Clone, Copy)]
enum DeviceDataFlow {
    Capture,
    Playback,
}
impl From<EDataFlow> for DeviceDataFlow {
    fn from(value: EDataFlow) -> Self {
        match value.0 {
            0 => Self::Playback,
            1 => Self::Capture,
            _ => unreachable!(),
        }
    }
}
impl std::fmt::Display for DeviceDataFlow {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "{}",
            match self {
                Self::Capture => "capture",
                Self::Playback => "playback",
            }
        )
    }
}

const fn calculate_buffer(sample_rate: u32, fundamental: u32, target: u32) -> u32 {
    sample_rate * target / 10000 / fundamental * fundamental
}

const fn calculate_period(sample_rate: u32, buffer_len: u32) -> i64 {
    (buffer_len * 100000 / (sample_rate / 100)) as i64
}

struct Shared3Info {
    current_period: u32,
    samplerate: u32,
    fundamental: u32,
}
impl Shared3Info {
    fn init(inner: &IAudioClient3, config: &ClientConfig, tag: &str) -> WinResult<Self> {
        let mut periods = [0; 4];
        let pformat = unsafe { inner.GetMixFormat()? };
        unsafe {
            inner.GetSharedModeEnginePeriod(
                pformat,
                &mut periods[0],
                &mut periods[1],
                &mut periods[2],
                &mut periods[3],
            )?
        };
        let samplerate = unsafe { *pformat }.nSamplesPerSec;
        unsafe { CoTaskMemFree(Some(pformat.cast())) };
        let current_period = if config.target_period_hus != 0 {
            calculate_buffer(samplerate, periods[1], config.target_period_hus)
                .clamp(periods[2], periods[3])
        } else {
            periods[2]
        };
        info_tagged!(
            tag,
            "Period: Current = {current_period}, Min = {}, Max = {}, Samplerate = {samplerate}",
            periods[2],
            periods[3]
        );
        Ok(Self {
            current_period,
            samplerate,
            fundamental: periods[1],
        })
    }
}

struct RedirectClientInfo {
    parameters: OnceCell<WinResult<Shared3Info>>,
    raw_flag: Once,
    config: &'static ClientConfig,
    tag: Box<str>,
    initialized: Cell<bool>,
}
impl RedirectClientInfo {
    fn new(config: &'static ClientConfig, tag: Box<str>) -> Self {
        Self {
            parameters: OnceCell::new(),
            raw_flag: Once::new(),
            config,
            tag,
            initialized: false.into(),
        }
    }
    fn param(&self, inner: &IAudioClient3) -> WinResult<&Shared3Info> {
        self.parameters
            .get_or_init(|| Shared3Info::init(inner, self.config, &self.tag))
            .as_ref()
            .map_err(|e| e.clone())
    }
    fn initialized(&self) -> bool {
        self.initialized.get()
    }
}

#[implement(IAudioClient3)]
struct RedirectAudioClient {
    inner: IAudioClient3,
    info: RedirectClientInfo,
}

impl RedirectAudioClient {
    fn new(inner: IAudioClient3, info: RedirectClientInfo) -> Self {
        Self { inner, info }
    }
}
impl IAudioClient_Impl for RedirectAudioClient_Impl {
    impl_boilerplate!(IAudioClient1);
    fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: u32,
        hnsbufferduration: i64,
        hnsperiodicity: i64,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> WinResult<()> {
        if sharemode == AUDCLNT_SHAREMODE_EXCLUSIVE {
            warn_tagged!(@self, "Rejected exclusive init");
            return AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED.ok();
        }
        info_tagged!(
            @self,
            "Initialize -> redirecting, original dur: {hnsbufferduration} * 100ns"
        );
        if streamflags & AUDCLNT_STREAMFLAGS_LOOPBACK == 0 {
            self.InitializeSharedAudioStream(streamflags, 0, pformat, audiosessionguid)
        } else {
            warn_tagged!(@self, "Bypassing loopback");
            unsafe {
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    hnsbufferduration,
                    hnsperiodicity,
                    pformat,
                    Some(audiosessionguid),
                )
            }
        }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> WinResult<()> {
        info_tagged!(@self, "GetDevicePeriod called");
        if self.info.initialized() || self.info.config.force_period {
            let mut minimumdeviceperiod = 0;
            unsafe {
                self.inner
                    .GetDevicePeriod(None, Some(&mut minimumdeviceperiod))?
            };
            if let Some(phnsdefaultdeviceperiod) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
                *phnsdefaultdeviceperiod = calculate_period(
                    self.info.param(&self.inner)?.samplerate,
                    self.info.param(&self.inner)?.current_period,
                )
                .max(minimumdeviceperiod)
            }
            if let Some(phnsminimumdeviceperiod) = unsafe { phnsminimumdeviceperiod.as_mut() } {
                *phnsminimumdeviceperiod = minimumdeviceperiod
            }
            // just assume no one will be silly here
            Ok(())
        } else {
            unsafe {
                self.inner
                    .GetDevicePeriod(Some(phnsdefaultdeviceperiod), Some(phnsminimumdeviceperiod))
            }
        }
    }

    fn GetBufferSize(&self) -> WinResult<u32> {
        if self.info.initialized() {
            let real_size = unsafe { self.inner.GetBufferSize()? };
            let param = self.info.param(&self.inner)?;
            let buf = self
                .info
                .config
                .target_buf_len(param)
                .map_or(real_size, |len| len.clamp(param.current_period, real_size));
            info_tagged!(@self, "GetBufferSize called, buffer length: {buf}");
            Ok(buf)
        } else {
            unsafe { self.inner.GetBufferSize() }
        }
    }

    fn GetCurrentPadding(&self) -> WinResult<u32> {
        trace_tagged!(@self, "GetCurrentPadding called");
        unsafe { self.inner.GetCurrentPadding() }
    }

    fn Start(&self) -> WinResult<()> {
        info_tagged!(@self, "Start called");
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> WinResult<()> {
        info_tagged!(@self, "Stop called");
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> WinResult<()> {
        info_tagged!(@self, "Reset called");
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> WinResult<()> {
        info_tagged!(@self, "SetEventHandle called");
        unsafe { self.inner.SetEventHandle(eventhandle) }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> WinResult<()> {
        debug_tagged!(@self, "GetService called, iid: {:?}", unsafe { *riid });
        unsafe {
            (self.inner.cast::<IAudioClient>()?.vtable().GetService)(self.inner.as_raw(), riid, ppv)
                .ok()
        }
    }
}

impl IAudioClient2_Impl for RedirectAudioClient_Impl {
    impl_boilerplate!(IAudioClient2);
    fn SetClientProperties(&self, pproperties: *const AudioClientProperties) -> WinResult<()> {
        info_tagged!(@self, "SetClientProperties called");
        if self.info.config.raw {
            self.info
                .raw_flag
                .call_once(|| info_tagged!(@self, "Applying raw flag"));
            let option = &mut unsafe { *pproperties }.Options;
            if option.contains(AUDCLNT_STREAMOPTIONS_RAW) {
                warn_tagged!(@self, "This stream already has raw flag!")
            } else {
                *option |= AUDCLNT_STREAMOPTIONS_RAW
            }
        }
        unsafe { self.inner.SetClientProperties(pproperties) }
    }
}

impl IAudioClient3_Impl for RedirectAudioClient_Impl {
    impl_boilerplate!(IAudioClient3);
    fn InitializeSharedAudioStream(
        &self,
        streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> WinResult<()> {
        if periodinframes != 0 {
            info_tagged!(
                @self,
                "InitializeSharedAudioStream -> replacing period, current period: {periodinframes}"
            );
        }
        self.info.initialized.set(true);
        unsafe {
            if self.info.config.raw && !self.info.raw_flag.is_completed() {
                info_tagged!(@self, "Applying raw flag");
                let properties = AudioClientProperties {
                    cbSize: size_of::<AudioClientProperties>() as u32,
                    Options: AUDCLNT_STREAMOPTIONS_RAW,
                    ..AudioClientProperties::default()
                };
                self.inner.SetClientProperties(&properties)?;
            }
            self.inner.InitializeSharedAudioStream(
                streamflags,
                self.info.param(&self.inner)?.current_period,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}

drop_boilerplate!(RedirectAudioClient);

#[implement(IAudioClient3)]
struct RedirectCompatAudioClient {
    inner: IAudioClient3,
    hooker: IAudioClient3,
    info: RedirectClientInfo,
    outer: OnceCell<IAudioRenderClient>,
    align: Cell<u16>,
}

impl RedirectCompatAudioClient {
    fn new(inner: IAudioClient3, hooker: IAudioClient3, info: RedirectClientInfo) -> Self {
        Self {
            inner,
            hooker,
            info,
            outer: OnceCell::new(),
            align: 0.into(),
        }
    }
}
impl IAudioClient_Impl for RedirectCompatAudioClient_Impl {
    impl_boilerplate!(IAudioClient1);
    fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: u32,
        hnsbufferduration: i64,
        hnsperiodicity: i64,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> WinResult<()> {
        if sharemode == AUDCLNT_SHAREMODE_EXCLUSIVE {
            warn_tagged!(@self, "Rejected exclusive init");
            return AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED.ok();
        }
        info_tagged!(
            @self,
            "Initialize -> setting hooker, original dur = {hnsbufferduration} * 100ns"
        );
        if streamflags & AUDCLNT_STREAMFLAGS_LOOPBACK == 0 {
            let calculated_dur = self
                .info
                .config
                .compat_buf_len(self.info.param(&self.inner)?)
                .unwrap_or_default();
            info_tagged!(@self, "Inner dur = {calculated_dur} * 100ns");
            unsafe {
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    calculated_dur,
                    hnsperiodicity,
                    pformat,
                    Some(audiosessionguid),
                )?
            }
            self.InitializeSharedAudioStream(streamflags, 0, pformat, audiosessionguid)
        } else {
            warn_tagged!(@self, "Bypassing loopback");
            unsafe {
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    hnsbufferduration,
                    hnsperiodicity,
                    pformat,
                    Some(audiosessionguid),
                )
            }
        }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> WinResult<()> {
        info_tagged!(@self, "GetDevicePeriod called");
        unsafe {
            self.inner
                .GetDevicePeriod(Some(phnsdefaultdeviceperiod), Some(phnsminimumdeviceperiod))
        }
    }

    fn GetBufferSize(&self) -> WinResult<u32> {
        let buf = unsafe { self.inner.GetBufferSize()? };
        info_tagged!(@self, "GetBufferSize called, buffer length: {buf}");
        Ok(buf)
    }

    fn GetCurrentPadding(&self) -> WinResult<u32> {
        trace_tagged!(@self, "GetCurrentPadding called");
        unsafe { self.inner.GetCurrentPadding() }
    }

    fn Start(&self) -> WinResult<()> {
        info_tagged!(@self, "Start called");
        if let Some(client) = self.outer.get() {
            let client: &RedirectCompatAudioRenderClient = unsafe { client.as_impl() };
            client.trick.set(false);
        }
        unsafe {
            _ = self.hooker.Start();
            self.inner.Start()
        }
    }

    fn Stop(&self) -> WinResult<()> {
        info_tagged!(@self, "Stop called");
        unsafe {
            _ = self.hooker.Stop();
            self.inner.Stop()
        }
    }

    fn Reset(&self) -> WinResult<()> {
        info_tagged!(@self, "Reset called");
        if let Some(client) = self.outer.get() {
            let client: &RedirectCompatAudioRenderClient = unsafe { client.as_impl() };
            client.trick.set(true);
        }
        unsafe {
            _ = self.hooker.Reset();
            self.inner.Reset()
        }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> WinResult<()> {
        info_tagged!(@self, "SetEventHandle called");
        unsafe {
            _ = self.hooker.SetEventHandle(eventhandle);
            self.inner.SetEventHandle(eventhandle)
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> WinResult<()> {
        let iid = unsafe { *riid };
        debug_tagged!(@self, "GetService called, iid: {iid:?}");
        match iid {
            IAudioRenderClient::IID if self.info.initialized() => {
                if let Some(client) = self.outer.get() {
                    unsafe { client.query(riid, ppv).ok() }
                } else {
                    let param = self.info.param(&self.inner)?;
                    let hooker_buffer_len = match (
                        unsafe { self.hooker.GetBufferSize().ok() },
                        self.info.config.target_buf_len(param),
                    ) {
                        (None, None) => param.current_period + param.fundamental,
                        (None, Some(len)) => {
                            len.clamp(param.current_period, unsafe { self.inner.GetBufferSize()? })
                        }
                        (Some(hooker_buf), None) => hooker_buf,
                        (Some(hooker_buf), Some(len)) => {
                            len.clamp(param.current_period, hooker_buf)
                        }
                    };
                    let inner_buffer_len = unsafe { self.inner.GetBufferSize()? };
                    let align = AudioAlign::new(self.align.get());
                    let service: IAudioRenderClient = RedirectCompatAudioRenderClient {
                        inner: unsafe { self.inner.GetService::<IAudioRenderClient>()? },
                        trick_buffer: vec![0; align.frames_to_bytes(inner_buffer_len as usize)]
                            .into_boxed_slice()
                            .into(),
                        trick: true.into(),
                        align,
                        buffer_len: (inner_buffer_len, hooker_buffer_len),
                        tag: format!("{}-client", self.info.tag).into(),
                    }
                    .into();
                    unsafe { service.query(riid, ppv).ok() }
                }
            }
            _ => unsafe {
                (self.inner.cast::<IAudioClient>()?.vtable().GetService)(
                    self.inner.as_raw(),
                    riid,
                    ppv,
                )
                .ok()
            },
        }
    }
}

impl IAudioClient2_Impl for RedirectCompatAudioClient_Impl {
    impl_boilerplate!(IAudioClient2);
    fn SetClientProperties(&self, pproperties: *const AudioClientProperties) -> WinResult<()> {
        info_tagged!(@self, "SetClientProperties called");
        if self.info.config.raw {
            self.info
                .raw_flag
                .call_once(|| info_tagged!(@self, "Applying raw flag"));
            let option = &mut unsafe { *pproperties }.Options;
            if option.contains(AUDCLNT_STREAMOPTIONS_RAW) {
                warn_tagged!(@self, "This stream already has raw flag!")
            } else {
                *option |= AUDCLNT_STREAMOPTIONS_RAW
            }
        }
        unsafe { self.inner.SetClientProperties(pproperties) }
    }
}

impl IAudioClient3_Impl for RedirectCompatAudioClient_Impl {
    impl_boilerplate!(IAudioClient3);
    fn InitializeSharedAudioStream(
        &self,
        streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> WinResult<()> {
        if self.info.config.raw && !self.info.raw_flag.is_completed() {
            info_tagged!(@self, "Applying raw flag");
            let properties = AudioClientProperties {
                cbSize: size_of::<AudioClientProperties>() as u32,
                Options: AUDCLNT_STREAMOPTIONS_RAW,
                ..AudioClientProperties::default()
            };
            unsafe { self.inner.SetClientProperties(&properties) }?;
        }
        self.align.set(unsafe { (*pformat).nBlockAlign });
        self.info.initialized.set(true);
        let client = if periodinframes != 0 {
            warn_tagged!(
                @self,
                "InitializeSharedAudioStream called, you shouldn't use this mode! original period: {periodinframes}"
            );
            &self.inner
        } else {
            &self.hooker
        };
        unsafe {
            client.InitializeSharedAudioStream(
                streamflags,
                self.info.param(&self.inner)?.current_period,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}
drop_boilerplate!(RedirectCompatAudioClient);

#[implement(IAudioRenderClient)]
struct RedirectCompatAudioRenderClient {
    inner: IAudioRenderClient,
    trick_buffer: UnsafeCell<Box<[u8]>>,
    trick: Cell<bool>,
    align: AudioAlign,
    buffer_len: (u32, u32),
    tag: Box<str>,
}
impl RedirectCompatAudioRenderClient {
    #[inline]
    fn apply_data(&self, len: u32, dwflags: u32) -> WinResult<()> {
        let read_len = self.align.frames_to_bytes(len as usize);
        unsafe {
            let slice_to_write = from_raw_parts_mut(self.inner.GetBuffer(len)?, read_len);
            let slice = &(&*self.trick_buffer.get())[..read_len];
            slice_to_write.copy_from_slice(slice);
            self.inner.ReleaseBuffer(len, dwflags)
        }
    }
}
impl IAudioRenderClient_Impl for RedirectCompatAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> WinResult<*mut u8> {
        if self.trick.get() {
            info_tagged!(
                self.tag,
                "GetBuffer called, requested: {numframesrequested}"
            );
            Ok(unsafe { *self.trick_buffer.get().cast() })
        } else {
            unsafe { self.inner.GetBuffer(numframesrequested) }
        }
    }

    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> WinResult<()> {
        if self.trick.get() {
            info_tagged!(
                self.tag,
                "ReleaseBuffer called, written: {numframeswritten}"
            );
            if dwflags == 2 {
                if numframeswritten == self.buffer_len.0 {
                    info_tagged!(
                        self.tag,
                        "filling silent buffer, {} frames filled",
                        self.buffer_len.1
                    );
                    self.apply_data(self.buffer_len.1, dwflags)
                } else {
                    info_tagged!(self.tag, "already filled, discarding");
                    Ok(())
                }
            } else {
                self.apply_data(numframeswritten, dwflags)
            }
        } else {
            if numframeswritten == 0 {
                warn_tagged!(
                    self.tag,
                    "no data written in this release call, overflow may happen!"
                );
            }
            unsafe { self.inner.ReleaseBuffer(numframeswritten, dwflags) }
        }
    }
}

#[implement(IAudioClient3)]
struct RedirectRingbufAudioClient {
    inner: IAudioClient3,
    info: RedirectClientInfo,
    buffer: Cell<u32>,
    align: Cell<AudioAlign>,
    outer: OnceCell<(IAudioRenderClient, IRtwqAsyncCallback)>,
    app_handle: Cell<HANDLE>,
}

impl RedirectRingbufAudioClient {
    fn new(inner: IAudioClient3, info: RedirectClientInfo) -> Self {
        Self {
            inner,
            info,
            buffer: 0.into(),
            align: AudioAlign::new(0).into(),
            outer: OnceCell::new(),
            app_handle: Cell::default(),
        }
    }
    fn set_buffer(&self, param: &Shared3Info) {
        self.buffer.update(|x| {
            if x != 0 {
                x
            } else {
                self.info
                    .config
                    .ring_buf_len(param)
                    .unwrap_or_else(|| param.current_period * 10)
            }
        })
    }
}

impl IAudioClient_Impl for RedirectRingbufAudioClient_Impl {
    impl_boilerplate!(IAudioClient1);
    fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: u32,
        hnsbufferduration: i64,
        hnsperiodicity: i64,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> WinResult<()> {
        if sharemode == AUDCLNT_SHAREMODE_EXCLUSIVE {
            warn_tagged!(@self, "Rejected exclusive init");
            return AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED.ok();
        }
        info_tagged!(
            @self,
            "Initialize -> adding ring buffer, original dur = {hnsbufferduration} * 100ns"
        );
        if streamflags & AUDCLNT_STREAMFLAGS_LOOPBACK == 0 {
            self.InitializeSharedAudioStream(streamflags, 0, pformat, audiosessionguid)
        } else {
            warn_tagged!(@self, "Bypassing loopback");
            unsafe {
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    hnsbufferduration,
                    hnsperiodicity,
                    pformat,
                    Some(audiosessionguid),
                )
            }
        }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> WinResult<()> {
        info_tagged!(@self, "GetDevicePeriod called");
        if self.info.initialized() || self.info.config.force_period {
            let mut minimumdeviceperiod = 0;
            unsafe {
                self.inner
                    .GetDevicePeriod(None, Some(&mut minimumdeviceperiod))?
            };
            if let Some(phnsdefaultdeviceperiod) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
                let param = self.info.param(&self.inner)?;
                self.set_buffer(param);
                *phnsdefaultdeviceperiod =
                    calculate_period(param.samplerate, self.buffer.get()).max(minimumdeviceperiod)
            }
            if let Some(phnsminimumdeviceperiod) = unsafe { phnsminimumdeviceperiod.as_mut() } {
                *phnsminimumdeviceperiod = minimumdeviceperiod
            }
            // just assume no one will be silly here
            Ok(())
        } else {
            unsafe {
                self.inner
                    .GetDevicePeriod(Some(phnsdefaultdeviceperiod), Some(phnsminimumdeviceperiod))
            }
        }
    }

    fn GetBufferSize(&self) -> WinResult<u32> {
        if self.info.initialized() {
            let buf = self.buffer.get();
            info_tagged!(@self, "GetBufferSize called, buffer length: {buf}");
            Ok(buf)
        } else {
            unsafe { self.inner.GetBufferSize() }
        }
    }

    fn GetCurrentPadding(&self) -> WinResult<u32> {
        trace_tagged!(@self, "GetCurrentPadding called");
        if let Some((outer, _)) = self.outer.get() {
            let outer: &RedirectRingbufAudioRenderClient = unsafe { outer.as_impl() };
            let buf = unsafe { &*outer.buffer.get() };
            let len = self.buffer.get();
            Ok(if buf.is_full() {
                len
            } else {
                len - self.align.get().bytes_to_frames(buf.slots()) as u32
            })
        } else {
            unsafe { self.inner.GetCurrentPadding() }
        }
    }

    fn Start(&self) -> WinResult<()> {
        info_tagged!(@self, "Start called");
        if let Some((outer, thread)) = self.outer.get() {
            let outer: &RedirectRingbufAudioRenderClient = unsafe { outer.as_impl() };
            outer.trick.set(false);
            let thread: &RedirectRingbufThread = unsafe { thread.as_impl() };
            (!self.app_handle.get().is_invalid())
                .then(|| thread.app_handle.get_or_init(|| self.app_handle.get()));
            thread.pause.store(false, Ordering::Relaxed);
        }
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> WinResult<()> {
        info_tagged!(@self, "Stop called");
        if let Some((_, thread)) = self.outer.get() {
            let thread: &RedirectRingbufThread = unsafe { thread.as_impl() };
            thread.pause.store(true, Ordering::Relaxed);
        }
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> WinResult<()> {
        info_tagged!(@self, "Reset called");
        if let Some((outer, thread)) = self.outer.get() {
            let thread: &RedirectRingbufThread = unsafe { thread.as_impl() };
            while unsafe { &mut *thread.buffer.get() }.pop().is_ok() {}
            let outer: &RedirectRingbufAudioRenderClient = unsafe { outer.as_impl() };
            outer.trick.set(true);
        }
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> WinResult<()> {
        info_tagged!(@self, "SetEventHandle called");
        if self.info.initialized() {
            self.app_handle.set(eventhandle);
            Ok(())
        } else {
            unsafe { self.inner.SetEventHandle(eventhandle) }
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> WinResult<()> {
        let iid = unsafe { *riid };
        debug_tagged!(@self, "GetService called, iid: {iid:?}");
        match iid {
            IAudioRenderClient::IID if self.info.initialized() => {
                if let Some((client, _)) = self.outer.get() {
                    unsafe { client.query(riid, ppv).ok() }
                } else {
                    let param = self.info.param(&self.inner)?;
                    let align = self.align.get();
                    let buffer = align.frames_to_bytes(self.buffer.get() as usize);
                    let event_handle = unsafe { CreateEventW(None, false, false, None)? };
                    unsafe { self.inner.SetEventHandle(event_handle)? }
                    let (producer, consumer) = RingBuffer::new(buffer);
                    unsafe { RtwqStartup()? };
                    let mut ids = [0; 2];
                    unsafe { RtwqLockSharedWorkQueue(w!("Audio"), 1, &mut ids[0], &mut ids[1])? };
                    let buf_size = unsafe { self.inner.GetBufferSize()? };
                    let real_size = self
                        .info
                        .config
                        .target_buf_len(param)
                        .map_or(buf_size, |len| len.clamp(param.current_period, buf_size));
                    info_tagged!(@self,"Creating thread");
                    let callback = RedirectRingbufThread {
                        buffer: consumer.into(),
                        client: self.inner.clone(),
                        inner: unsafe { self.inner.GetService::<IAudioRenderClient>()? },
                        align,
                        real_len: real_size,
                        event: unsafe { Owned::new(event_handle) },
                        app_handle: OnceCell::new(),
                        thread_id: ids[1],
                        tag: format!("{}-thread", self.info.tag).into(),
                        pause: false.into(),
                    };
                    let callback: IRtwqAsyncCallback = callback.into();
                    let result = unsafe { RtwqCreateAsyncResult(None, &callback, None)? };
                    unsafe { RtwqPutWaitingWorkItem(event_handle, 1, &result, None)? }
                    let client: IAudioRenderClient = RedirectRingbufAudioRenderClient {
                        buffer: producer.into(),
                        cache: vec![0u8; buffer].into_boxed_slice().into(),
                        align,
                        trick: true.into(),
                        tag: format!("{}-render", self.info.tag).into(),
                    }
                    .into();
                    let ret = unsafe { client.query(riid, ppv) }.ok();
                    _ = self.outer.set((client, callback));
                    ret
                }
            }
            _ => unsafe {
                (self.inner.cast::<IAudioClient>()?.vtable().GetService)(
                    self.inner.as_raw(),
                    riid,
                    ppv,
                )
                .ok()
            },
        }
    }
}

impl IAudioClient2_Impl for RedirectRingbufAudioClient_Impl {
    impl_boilerplate!(IAudioClient2);
    fn SetClientProperties(&self, pproperties: *const AudioClientProperties) -> WinResult<()> {
        info_tagged!(@self, "SetClientProperties called");
        if self.info.config.raw {
            self.info
                .raw_flag
                .call_once(|| info_tagged!(@self, "Applying raw flag"));
            let option = &mut unsafe { *pproperties }.Options;
            if option.contains(AUDCLNT_STREAMOPTIONS_RAW) {
                warn_tagged!(@self, "This stream already has raw flag!")
            } else {
                *option |= AUDCLNT_STREAMOPTIONS_RAW
            }
        }
        unsafe { self.inner.SetClientProperties(pproperties) }
    }
}

impl IAudioClient3_Impl for RedirectRingbufAudioClient_Impl {
    impl_boilerplate!(IAudioClient3);
    fn InitializeSharedAudioStream(
        &self,
        mut streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> WinResult<()> {
        if periodinframes != 0 {
            info_tagged!(
                @self,
                "InitializeSharedAudioStream -> adding ring buffer, current period: {periodinframes}"
            );
        }
        self.info.initialized.set(true);
        unsafe {
            let target_config = self.info.config;
            if target_config.raw && !self.info.raw_flag.is_completed() {
                info_tagged!(@self, "Applying raw flag");
                let properties = AudioClientProperties {
                    cbSize: size_of::<AudioClientProperties>() as u32,
                    Options: AUDCLNT_STREAMOPTIONS_RAW,
                    ..AudioClientProperties::default()
                };
                self.inner.SetClientProperties(&properties)?;
            }
            let param = self.info.param(&self.inner)?;
            self.set_buffer(param);
            self.align.set(AudioAlign::new((*pformat).nBlockAlign));
            if streamflags & AUDCLNT_STREAMFLAGS_EVENTCALLBACK == 0 {
                info_tagged!(@self, "Injecting event flag");
                streamflags |= AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
            } else {
                info_tagged!(@self, "Enabling inverse mode");
            }
            self.inner.InitializeSharedAudioStream(
                streamflags,
                param.current_period,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}
drop_boilerplate!(RedirectRingbufAudioClient);

#[implement(IRtwqAsyncCallback)]
struct RedirectRingbufThread {
    buffer: UnsafeCell<Consumer<u8>>,
    client: IAudioClient3,
    inner: IAudioRenderClient,
    align: AudioAlign,
    real_len: u32,
    event: Owned<HANDLE>,
    app_handle: OnceCell<HANDLE>,
    thread_id: u32,
    tag: Box<str>,
    pause: AtomicBool,
}
impl IRtwqAsyncCallback_Impl for RedirectRingbufThread_Impl {
    fn GetParameters(&self, _: *mut u32, pdwqueue: *mut u32) -> WinResult<()> {
        unsafe { *pdwqueue = self.thread_id }
        Ok(())
    }
    fn Invoke(&self, pasyncresult: Ref<IRtwqAsyncResult>) -> WinResult<()> {
        let buffer = unsafe { &mut *self.buffer.get() };
        if buffer.is_abandoned() {
            return Ok(());
        } else {
            unsafe { RtwqPutWaitingWorkItem(*self.event, 1, pasyncresult.as_ref(), None)? }
            if self.pause.load(Ordering::Relaxed) {
                return Ok(());
            }
        }
        if buffer.is_empty() {
            let pad = unsafe { self.client.GetCurrentPadding()? };
            if pad == 0 {
                warn_tagged!(self.tag, "buffer is empty, underflow may happen!")
            } else {
                debug_tagged!(self.tag, "mid-buffer empty, data in client buffer: {pad}")
            }
        } else {
            let read_len = self.align.bytes_to_frames(buffer.slots());
            let write_len = read_len
                .min((self.real_len - unsafe { self.client.GetCurrentPadding()? }) as usize);
            let slice = unsafe {
                from_raw_parts_mut(
                    self.inner.GetBuffer(write_len as u32)?,
                    self.align.frames_to_bytes(write_len),
                )
            };
            buffer
                .pop_entire_slice(slice)
                .unwrap_or_else(|e| warn_tagged!(self.tag, "pop overflow! {e}"));
            unsafe { self.inner.ReleaseBuffer(write_len as u32, 0)? };
            trace_tagged!(
                self.tag,
                "data in mid-buffer: {read_len}, written: {write_len}"
            )
        }
        unsafe { self.app_handle.get().map_or(Ok(()), |&h| SetEvent(h)) }
    }
}
impl Drop for RedirectRingbufThread {
    fn drop(&mut self) {
        unsafe {
            RtwqUnlockWorkQueue(self.thread_id)
                .and_then(|_| RtwqShutdown())
                .unwrap_or_else(|e| {
                    error_tagged!(self.tag, "Encountered error when closing thread: {e}")
                });
        }
        info_tagged!(self.tag, "Consumer thread stopped");
    }
}

#[implement(IAudioRenderClient)]
struct RedirectRingbufAudioRenderClient {
    buffer: UnsafeCell<Producer<u8>>,
    cache: UnsafeCell<Box<[u8]>>,
    align: AudioAlign,
    trick: Cell<bool>,
    tag: Box<str>,
}
impl IAudioRenderClient_Impl for RedirectRingbufAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> WinResult<*mut u8> {
        if self.trick.get() {
            info_tagged!(
                self.tag,
                "GetBuffer called, requested: {numframesrequested}"
            );
        } else {
            trace_tagged!(
                self.tag,
                "GetBuffer called, requested: {numframesrequested}"
            );
        }
        Ok(unsafe { &mut *self.cache.get() }.as_mut_ptr())
    }
    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> WinResult<()> {
        if numframeswritten == 0 {
            return Ok(());
        }
        if self.trick.get() {
            info_tagged!(
                self.tag,
                "ReleaseBuffer called, written: {numframeswritten}"
            );
            if dwflags == 2 {
                info_tagged!(self.tag, "discarding silent data");
                return Ok(());
            }
        }
        unsafe {
            let buffer = &mut *self.buffer.get();
            let slice = &mut (&mut *self.cache.get())
                [..(self.align.frames_to_bytes(numframeswritten as usize))];
            if dwflags == 2 {
                slice.fill(0);
            }
            buffer
                .push_entire_slice(slice)
                .unwrap_or_else(|e| warn_tagged!(self.tag, "push overflow! {e}"));
            debug_tagged!(
                self.tag,
                "ReleaseBuffer called, written: {numframeswritten}"
            );
            Ok(())
        }
    }
}
impl Drop for RedirectRingbufAudioRenderClient {
    fn drop(&mut self) {
        info_tagged!(self.tag, "Stopping consumer thread");
    }
}

fn formatter(
    w: &mut dyn std::io::Write,
    _: &mut DeferredNow,
    record: &Record,
) -> std::io::Result<()> {
    write!(w, "{} [{}] ", record.level(), record.target())?;
    write!(w, "{}", record.args())
}

fn setup() -> LoggerHandle {
    std::panic::set_hook(Box::new(|panic_info| {
        error!("{panic_info}");
    }));
    let logger = Logger::with(LevelFilter::from(CONFIG.log_level))
        .format(formatter)
        .write_mode(WriteMode::Async);
    let handle = if !CONFIG.only_log_stdout {
        logger
            .log_to_file(
                FileSpec::default()
                    .basename("wasapi_relink")
                    .suppress_timestamp()
                    .directory(
                        CONFIG
                            .log_path
                            .as_ref()
                            .filter(|path| path.is_dir())
                            .map_or(Path::new("."), |p| p.as_ref()),
                    ),
            )
            .duplicate_to_stdout(Duplicate::All)
    } else {
        logger.log_to_stdout()
    }
    .start()
    .expect("unable to setup logger");
    info!(
        "Attempting to load config from working directory: {}",
        std::env::current_dir().map_or_else(
            |e| format!("unknown, err: {e}"),
            |path| path.display().to_string()
        )
    );
    match CONFIG.source {
        ConfigSource::Success => info!("Config loaded!"),
        ConfigSource::NoParse => {
            warn!("Unable to parse config, using default values")
        }
        ConfigSource::NoFile => {
            info!("Config file not found, using default values")
        }
    }
    handle
}

#[unsafe(export_name = "proxy")]
extern "C" fn proxy_dummy() {}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(_: HINSTANCE, reason: u32, _: *mut c_void) -> BOOL {
    match reason {
        1 => unsafe { CO_CREATE.0.enable().is_ok() && CO_CREATE.1.enable().is_ok() },
        0 => unsafe { CO_CREATE.0.disable().is_ok() && CO_CREATE.1.disable().is_ok() },
        _ => true,
    }
    .into()
}
