// #[cfg(test)]
// mod config_test;

use flexi_logger::*;
use log::*;
use num_traits::PrimInt;
use retour::GenericDetour;
use ringbuf::traits::*;
use ringbuf::*;
use serde::*;
use std::cell::{Cell, OnceCell, UnsafeCell};
use std::collections::HashMap;
use std::hint::unreachable_unchecked;
use std::mem::transmute;
use std::os::raw::c_void;
use std::path::PathBuf;
use std::rc::Rc;
use std::slice::from_raw_parts_mut;
use std::sync::{Arc, LazyLock, Once, OnceLock, atomic::*};
use std::thread::spawn;
use windows::Win32::System::Threading::{
    AVRT_PRIORITY_HIGH, IRtwqAsyncCallback, IRtwqAsyncCallback_Impl, IRtwqAsyncResult,
    RtwqCreateAsyncResult, RtwqLockSharedWorkQueue, RtwqPutWaitingWorkItem, RtwqShutdown,
    RtwqStartup, RtwqUnlockWorkQueue,
};

use windows::{
    Win32::{
        Foundation::*,
        Media::Audio::*,
        System::Com::{StructuredStorage::*, *},
        System::LibraryLoader::{GetModuleHandleW, GetProcAddress, LoadLibraryW},
        System::SystemServices::{DLL_PROCESS_ATTACH, DLL_PROCESS_DETACH},
        System::Threading::{CreateEventW, GetCurrentThread, GetThreadDescription, SetEvent},
        UI::Shell::PropertiesSystem::IPropertyStore,
    },
    core::*,
};

static LOG_SETUP: Once = Once::new();
fn setup() {
    spawn(|| {
        let logger = Logger::with(<ConfigLogLevel as Into<LevelFilter>>::into(
            CONFIG.log_level,
        ))
        .format(formatter)
        .write_mode(WriteMode::Async);
        let handle = if !CONFIG.only_log_stdout {
            logger
                .log_to_file({
                    let spec = FileSpec::default()
                        .basename("wasapi_relink")
                        .suppress_timestamp();
                    if let Some(path) = &CONFIG.log_path
                        && path.is_dir()
                    {
                        spec.directory(path)
                    } else {
                        spec
                    }
                })
                .duplicate_to_stdout(Duplicate::All)
        } else {
            logger.log_to_stdout()
        }
        .start()
        .expect("unable to setup logger");
        LOGGER_HANDLE.set(handle).ok();
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
    });
}
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
    log_path: Option<PathBuf>,
    log_level: ConfigLogLevel,
    only_log_stdout: bool,
    playback: ClientConfig,
    capture: ClientConfig,
    #[serde(skip)]
    source: ConfigSource,
}
impl RedirectConfig {
    fn load() -> Self {
        if let Ok(str) = std::fs::read_to_string("redirect_config.toml") {
            if let Ok(cfg) = toml::from_str::<RedirectConfig>(&str) {
                cfg
            } else {
                Self::new_with_source(ConfigSource::NoParse)
            }
        } else {
            Self::new_with_source(ConfigSource::NoFile)
        }
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
    target_buffer_dur_ms: u32,
    ring_buffer_len: HashMap<u32, u32>,
    mode: ClientMode,
    compat_buffer_len: HashMap<u32, i64>,
    raw: bool,
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
enum AudioAlign<T: PrimInt> {
    Pow2(usize),
    Normal(T),
}
impl<T: PrimInt> AudioAlign<T> {
    #[inline(always)]
    fn new(align: T) -> Self {
        if (align & (align - T::one())) == T::zero() {
            Self::Pow2(align.trailing_zeros() as usize)
        } else {
            Self::Normal(align)
        }
    }
    #[inline(always)]
    fn bytes_to_frames(&self, bytes: T) -> T {
        match self {
            Self::Pow2(shift) => bytes >> *shift,
            Self::Normal(align) => bytes / *align,
        }
    }

    // 帧数转字节数：frames * align
    #[inline(always)]
    fn frames_to_bytes(&self, frames: T) -> T {
        match self {
            Self::Pow2(shift) => frames << *shift,
            Self::Normal(align) => frames * *align,
        }
    }
}
impl AudioAlign<u32> {
    #[inline(always)]
    fn as_usize(&self) -> AudioAlign<usize> {
        match self {
            Self::Pow2(shift) => AudioAlign::Pow2(*shift),
            Self::Normal(align) => AudioAlign::Normal(*align as usize),
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

static PLAYBACK_ID: AtomicU16 = AtomicU16::new(0);
static CAPTURE_ID: AtomicU16 = AtomicU16::new(0);

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

const LIB_NAME: PCWSTR = w!("ole32.dll");
const CO_CREATE: PCSTR = s!("CoCreateInstance");
const CO_CREATE_EX: PCSTR = s!("CoCreateInstanceEx");

const AUDIO_TASK: PCWSTR = w!("Audio");

const KEYWORDS: &[&str] = &["[GAME]", "[SK]"];

#[inline]
unsafe fn get_ole32() -> HMODULE {
    unsafe { GetModuleHandleW(LIB_NAME).unwrap_or_else(|_| LoadLibraryW(LIB_NAME).unwrap()) }
}

static HOOK_CO_CREATE_INSTANCE: LazyLock<GenericDetour<FnCoCreateInstance>> =
    LazyLock::new(|| unsafe {
        let func = GetProcAddress(get_ole32(), CO_CREATE).unwrap();
        let func: FnCoCreateInstance = transmute(func);
        GenericDetour::new(func, hooked_cocreateinstance).unwrap()
    });

static HOOK_CO_CREATE_INSTANCE_EX: LazyLock<GenericDetour<FnCoCreateInstanceEx>> =
    LazyLock::new(|| unsafe {
        let func = GetProcAddress(get_ole32(), CO_CREATE_EX).unwrap();
        let func: FnCoCreateInstanceEx = transmute(func);
        GenericDetour::new(func, hooked_cocreateinstanceex).unwrap()
    });

unsafe extern "system" fn hooked_cocreateinstance(
    rclsid: *const GUID,
    p_outer: *mut c_void,
    dwcls_context: CLSCTX,
    riid: *const GUID,
    ppv: *mut *mut c_void,
) -> HRESULT {
    unsafe {
        let ret = HOOK_CO_CREATE_INSTANCE.call(rclsid, p_outer, dwcls_context, riid, ppv);
        if *riid == IMMDeviceEnumerator::IID {
            LOG_SETUP.call_once(setup);
            if ret.is_ok() {
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
            } else {
                error!("CoCreateInstance failed with HRESULT: {ret}")
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
        let hr = HOOK_CO_CREATE_INSTANCE_EX.call(
            clsid,
            punkouter,
            dwclsctx,
            pserverinfo,
            dwcount,
            presults,
        );
        if *clsid == MMDeviceEnumerator {
            LOG_SETUP.call_once(setup);
            if hr.is_ok() {
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
            } else {
                error!("CoCreateInstanceEx failed with HRESULT: {hr}")
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
    ) -> windows::core::Result<IMMDeviceCollection> {
        debug!(
            "DeviceEnumerator::EnumAudioEndpoints requested on flow {}",
            dataflow.0
        );
        Ok(RedirectDeviceCollection {
            inner: unsafe { self.inner.EnumAudioEndpoints(dataflow, dwstatemask)? },
        }
        .into())
    }

    fn GetDefaultAudioEndpoint(
        &self,
        dataflow: EDataFlow,
        role: ERole,
    ) -> windows::core::Result<IMMDevice> {
        debug!(
            "DeviceEnumerator::GetDefaultAudioEndpoint requested on flow {}",
            dataflow.0
        );
        Ok(
            RedirectDevice::new(unsafe { self.inner.GetDefaultAudioEndpoint(dataflow, role)? })
                .into(),
        )
    }

    fn GetDevice(&self, pwstrid: &PCWSTR) -> windows::core::Result<IMMDevice> {
        info!("DeviceEnumerator::GetDevice called, wrapping");
        Ok(RedirectDevice::new(unsafe { self.inner.GetDevice(*pwstrid)? }).into())
    }

    fn RegisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> windows::core::Result<()> {
        trace!("DeviceEnumerator::RegisterEndpointNotificationCallback called");
        unsafe {
            self.inner
                .RegisterEndpointNotificationCallback(pclient.as_ref())
        }
    }

    fn UnregisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> windows::core::Result<()> {
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
    fn GetCount(&self) -> windows::core::Result<u32> {
        trace!("DeviceCollection::GetCount called");
        unsafe { self.inner.GetCount() }
    }

    fn Item(&self, ndevice: u32) -> windows::core::Result<IMMDevice> {
        debug!("DeviceCollection::Item retrieved device {ndevice}");
        Ok(RedirectDevice::new(unsafe { self.inner.Item(ndevice)? }).into())
    }
}

macro_rules! impl_boilerplate {
    (IAudioClient1) => {
        fn GetMixFormat(&self) -> windows::core::Result<*mut WAVEFORMATEX> {
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
        fn GetStreamLatency(&self) -> windows::core::Result<i64> {
            info_tagged!(@self, "GetStreamLatency called");
            unsafe { self.inner.GetStreamLatency() }
        }
        fn GetDevicePeriod(
            &self,
            phnsdefaultdeviceperiod: *mut i64,
            phnsminimumdeviceperiod: *mut i64,
        ) -> windows::core::Result<()> {
            info_tagged!(@self, "GetDevicePeriod called");
            let mut minimumdeviceperiod = 0;
            unsafe {
                self.inner
                    .GetDevicePeriod(None, Some(&mut minimumdeviceperiod))?
            };
            if let Some(phnsdefaultdeviceperiod) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
                *phnsdefaultdeviceperiod = calculate_period(
                    self.get_info().samplerate,
                    self.get_info().current_buffer_len,
                )
                .max(minimumdeviceperiod)
            }
            if let Some(phnsminimumdeviceperiod) = unsafe { phnsminimumdeviceperiod.as_mut() } {
                *phnsminimumdeviceperiod = minimumdeviceperiod
            }
            // just assume no one will be silly here
            Ok(())
        }
    };
    (IAudioClient2) =>
    {
        fn IsOffloadCapable(&self, category: AUDIO_STREAM_CATEGORY) -> windows::core::Result<BOOL> {
            info_tagged!(@self, "IsOffloadCapable called");
            unsafe { self.inner.IsOffloadCapable(category) }
        }
        fn GetBufferSizeLimits(
            &self,
            pformat: *const WAVEFORMATEX,
            beventdriven: BOOL,
            phnsminbufferduration: *mut i64,
            phnsmaxbufferduration: *mut i64,
        ) -> windows::core::Result<()> {
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
        ) -> windows::core::Result<()> {
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
        ) -> windows::core::Result<()> {
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
    ) -> windows::core::Result<()> {
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
                            DeviceDataFlow::Capture => CAPTURE_ID.fetch_add(1, Ordering::Relaxed),
                            DeviceDataFlow::Playback => PLAYBACK_ID.fetch_add(1, Ordering::Relaxed),
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

    fn OpenPropertyStore(&self, stgmaccess: STGM) -> windows::core::Result<IPropertyStore> {
        debug!("Device::OpenPropertyStore -> wrapping");
        unsafe { self.inner.OpenPropertyStore(stgmaccess) }
    }

    fn GetId(&self) -> windows::core::Result<PWSTR> {
        debug!("Device::GetId called");
        unsafe { self.inner.GetId() }
    }

    fn GetState(&self) -> windows::core::Result<DEVICE_STATE> {
        trace!("Device::GetState called");
        unsafe { self.inner.GetState() }
    }
}
impl IMMEndpoint_Impl for RedirectDevice_Impl {
    fn GetDataFlow(&self) -> windows::core::Result<EDataFlow> {
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
            _ => unsafe { unreachable_unchecked() },
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
    (buffer_len * 10000000 / sample_rate) as i64
}

struct Shared3Info {
    current_buffer_len: u32,
    samplerate: u32,
    fundamental: u32,
}
impl Shared3Info {
    fn init(inner: &IAudioClient3, info: &RedirectClientInfo) -> Self {
        let mut pdefaultperiodinframes = 0;
        let mut pfundamentalperiodinframes = 0;
        let mut pminperiodinframes = 0;
        let mut pmaxperiodinframes = 0;
        let pformat = unsafe { inner.GetMixFormat().unwrap() };
        unsafe {
            inner
                .GetSharedModeEnginePeriod(
                    pformat,
                    &mut pdefaultperiodinframes,
                    &mut pfundamentalperiodinframes,
                    &mut pminperiodinframes,
                    &mut pmaxperiodinframes,
                )
                .unwrap()
        };
        let samplerate = unsafe { *pformat }.nSamplesPerSec;
        let current_buffer_len = if info.config.target_buffer_dur_ms != 0 {
            calculate_buffer(
                samplerate,
                pfundamentalperiodinframes,
                info.config.target_buffer_dur_ms,
            )
            .clamp(pminperiodinframes, pmaxperiodinframes)
        } else {
            pminperiodinframes
        };
        info_tagged!(
            info.tag,
            "Period: Current = {current_buffer_len}, Min = {pminperiodinframes}, Max = {pmaxperiodinframes}, Samplerate = {samplerate}"
        );
        Self {
            current_buffer_len,
            samplerate,
            fundamental: pfundamentalperiodinframes,
        }
    }
}

struct RedirectClientInfo {
    parameters: OnceCell<Shared3Info>,
    raw_flag: Once,
    config: &'static ClientConfig,
    tag: Box<str>,
}
impl RedirectClientInfo {
    fn new(config: &'static ClientConfig, tag: Box<str>) -> Self {
        Self {
            parameters: OnceCell::new(),
            raw_flag: Once::new(),
            config,
            tag,
        }
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
    fn get_info(&self) -> &Shared3Info {
        self.info
            .parameters
            .get_or_init(|| Shared3Info::init(&self.inner, &self.info))
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
    ) -> windows::core::Result<()> {
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

    fn GetBufferSize(&self) -> windows::core::Result<u32> {
        let buf = unsafe { self.inner.GetBufferSize()? };
        info_tagged!(@self, "GetBufferSize called, buffer length: {buf}");
        Ok(buf)
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace_tagged!(@self, "GetCurrentPadding called");
        let pad = unsafe { self.inner.GetCurrentPadding()? };
        if pad == 0 {
            warn_tagged!(@self, "underflow may happen!");
        }
        Ok(pad)
    }

    fn Start(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Start called");
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Stop called");
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Reset called");
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info_tagged!(@self, "SetEventHandle called");
        unsafe { self.inner.SetEventHandle(eventhandle) }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        debug_tagged!(@self, "GetService called, iid: {:?}", unsafe { *riid });
        unsafe {
            (self.inner.cast::<IAudioClient>()?.vtable().GetService)(self.inner.as_raw(), riid, ppv)
                .ok()
        }
    }
}

impl IAudioClient2_Impl for RedirectAudioClient_Impl {
    impl_boilerplate!(IAudioClient2);
    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows::core::Result<()> {
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
    ) -> windows::core::Result<()> {
        if periodinframes != 0 {
            info_tagged!(
                @self,
                "InitializeSharedAudioStream -> replacing period, current period: {periodinframes}"
            );
        }
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
                self.get_info().current_buffer_len,
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
    trick: Rc<Cell<bool>>,
    align: Cell<u32>,
}

impl RedirectCompatAudioClient {
    fn new(inner: IAudioClient3, hooker: IAudioClient3, info: RedirectClientInfo) -> Self {
        Self {
            inner,
            hooker,
            info,
            trick: Cell::new(true).into(),
            align: 0.into(),
        }
    }
    fn get_info(&self) -> &Shared3Info {
        self.info
            .parameters
            .get_or_init(|| Shared3Info::init(&self.hooker, &self.info))
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
    ) -> windows::core::Result<()> {
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
                .compat_buffer_len
                .get(&self.get_info().samplerate)
                .copied()
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

    fn GetBufferSize(&self) -> windows::core::Result<u32> {
        let buf = unsafe { self.inner.GetBufferSize()? };
        info_tagged!(@self, "GetBufferSize called, buffer length: {buf}");
        Ok(buf)
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace_tagged!(@self, "GetCurrentPadding called");
        let pad = unsafe { self.inner.GetCurrentPadding()? };
        if pad == 0 {
            warn_tagged!(@self, "underflow may happen!");
        }
        Ok(pad)
    }

    fn Start(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Start called");
        unsafe {
            self.trick.set(false);
            _ = self.hooker.Start();
            self.inner.Start()
        }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Stop called");
        unsafe {
            _ = self.hooker.Stop();
            self.inner.Stop()
        }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Reset called");
        unsafe {
            self.trick.set(true);
            _ = self.hooker.Reset();
            self.inner.Reset()
        }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info_tagged!(@self, "SetEventHandle called");
        unsafe {
            _ = self.hooker.SetEventHandle(eventhandle);
            self.inner.SetEventHandle(eventhandle)
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        let iid = unsafe { *riid };
        debug_tagged!(@self, "GetService called, iid: {iid:?}");
        match iid {
            IAudioRenderClient::IID if self.align.get() != 0 => unsafe {
                let hooker_buffer_len = self.hooker.GetBufferSize().unwrap_or_else(|_| {
                    self.get_info().current_buffer_len + self.get_info().fundamental
                });
                let inner_buffer_len = self.inner.GetBufferSize()?;
                let service: IAudioRenderClient = RedirectCompatAudioRenderClient {
                    inner: self.inner.GetService::<IAudioRenderClient>()?,
                    trick_buffer: vec![0; inner_buffer_len as usize * self.align.get() as usize]
                        .into_boxed_slice()
                        .into(),
                    trick: self.trick.clone(),
                    align: AudioAlign::new(self.align.get()),
                    buffer_len: (inner_buffer_len, hooker_buffer_len),
                    tag: format!("{}-client", self.info.tag).into(),
                }
                .into();
                service.query(&IAudioRenderClient::IID, ppv).ok()
            },
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
    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows::core::Result<()> {
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
    ) -> windows::core::Result<()> {
        if self.info.config.raw && !self.info.raw_flag.is_completed() {
            info_tagged!(@self, "Applying raw flag");
            let properties = AudioClientProperties {
                cbSize: size_of::<AudioClientProperties>() as u32,
                Options: AUDCLNT_STREAMOPTIONS_RAW,
                ..AudioClientProperties::default()
            };
            unsafe { self.inner.SetClientProperties(&properties) }?;
        }
        self.align.set(unsafe { (*pformat).nBlockAlign as u32 });
        let client = if periodinframes != 0 {
            info_tagged!(
                @self,
                "InitializeSharedAudioStream called, original period: {periodinframes}"
            );
            &self.inner
        } else {
            &self.hooker
        };
        unsafe {
            client.InitializeSharedAudioStream(
                streamflags,
                self.get_info().current_buffer_len,
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
    trick: Rc<Cell<bool>>,
    align: AudioAlign<u32>,
    buffer_len: (u32, u32),
    tag: Box<str>,
}
impl RedirectCompatAudioRenderClient {
    #[inline]
    fn apply_data(&self, len: u32, dwflags: u32) -> windows::core::Result<()> {
        let read_len = self.align.frames_to_bytes(len) as usize;
        unsafe {
            let slice_to_write = from_raw_parts_mut(self.inner.GetBuffer(len)?, read_len);
            let slice = &(&*self.trick_buffer.get())[..read_len];
            slice_to_write.copy_from_slice(slice);
            self.inner.ReleaseBuffer(len, dwflags)
        }
    }
}
impl IAudioRenderClient_Impl for RedirectCompatAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows::core::Result<*mut u8> {
        if self.trick.get() {
            info_tagged!(
                self.tag,
                "GetBuffer called, requested: {numframesrequested}"
            );
            Ok(unsafe { &mut *self.trick_buffer.get() }.as_mut_ptr())
        } else {
            unsafe { self.inner.GetBuffer(numframesrequested) }
        }
    }

    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows::core::Result<()> {
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
    buffer: OnceCell<Arc<HeapRb<u8>>>,
    align: Cell<AudioAlign<u32>>,
    outer: OnceCell<IAudioRenderClient>,
    trick: Rc<Cell<bool>>,
    app_handle: Arc<AtomicPtr<c_void>>,
}

impl RedirectRingbufAudioClient {
    fn new(inner: IAudioClient3, info: RedirectClientInfo) -> Self {
        Self {
            inner,
            info,
            buffer: OnceCell::new(),
            align: AudioAlign::Normal(0).into(),
            outer: OnceCell::new(),
            trick: Cell::new(true).into(),
            app_handle: AtomicPtr::default().into(),
        }
    }
    fn get_info(&self) -> &Shared3Info {
        self.info
            .parameters
            .get_or_init(|| Shared3Info::init(&self.inner, &self.info))
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
    ) -> windows::core::Result<()> {
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

    fn GetBufferSize(&self) -> windows::core::Result<u32> {
        if let Some(buf) = self.buffer.get() {
            let buf = self
                .align
                .get()
                .bytes_to_frames(buf.capacity().get() as u32);
            info_tagged!(@self, "GetBufferSize called, buffer length: {buf}");
            Ok(buf)
        } else {
            unsafe { self.inner.GetBufferSize() }
        }
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace_tagged!(@self, "GetCurrentPadding called");
        if let Some(buf) = self.buffer.get() {
            Ok(self.align.get().bytes_to_frames(buf.occupied_len() as u32))
        } else {
            unsafe { self.inner.GetCurrentPadding() }
        }
    }

    fn Start(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Start called");
        self.trick.set(false);
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Stop called");
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info_tagged!(@self, "Reset called");
        self.trick.set(true);
        if let Some(buf) = self.buffer.get() {
            unsafe { &mut *(Arc::as_ptr(buf).cast_mut()) }.clear();
        }
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info_tagged!(@self, "SetEventHandle called");
        if self.buffer.get().is_some() {
            self.app_handle.store(eventhandle.0, Ordering::Relaxed);
            Ok(())
        } else {
            unsafe { self.inner.SetEventHandle(eventhandle) }
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        let iid = unsafe { *riid };
        debug_tagged!(@self, "GetService called, iid: {iid:?}");
        match iid {
            IAudioRenderClient::IID => {
                if let Some(client) = self.outer.get() {
                    unsafe { client.query(&IAudioRenderClient::IID, ppv).ok() }
                } else {
                    let align = self.align.get().as_usize();
                    let thread_tag = format!("{}-thread", self.info.tag).into_boxed_str();
                    let client_tag = format!("{}-render", self.info.tag).into();
                    let buffer = unsafe { self.buffer.get().unwrap_unchecked() };
                    let event_handle = unsafe { CreateEventW(None, false, false, None)? };
                    let stop_flag = Arc::new(AtomicBool::default());
                    unsafe { self.inner.SetEventHandle(event_handle)? }
                    let (producer, consumer) = buffer.clone().split();
                    unsafe { RtwqStartup()? };
                    let (mut _task_id, mut thread_id) = (0, 0);
                    unsafe {
                        RtwqLockSharedWorkQueue(
                            AUDIO_TASK,
                            AVRT_PRIORITY_HIGH.0,
                            &mut _task_id,
                            &mut thread_id,
                        )?
                    };
                    info_tagged!(@self,"Creating thread");
                    let callback = RedirectRingbufThread {
                        stop_flag: stop_flag.clone(),
                        thread_id,
                        client: self.inner.clone(),
                        app_handle: self.app_handle.clone(),
                        event_handle,
                        buffer: consumer.into(),
                        real_len: unsafe { self.inner.GetBufferSize()? },
                        inner: unsafe { self.inner.GetService::<IAudioRenderClient>()? },
                        align,
                        tag: thread_tag,
                    };
                    let callback: IRtwqAsyncCallback = callback.into();
                    let result = unsafe { RtwqCreateAsyncResult(None, &callback, None)? };
                    unsafe { RtwqPutWaitingWorkItem(event_handle, 1, &result, None)? }
                    let client: IAudioRenderClient = RedirectRingbufAudioRenderClient {
                        buffer: producer.into(),
                        cache: vec![0u8; buffer.capacity().get()].into_boxed_slice().into(),
                        align,
                        trick: self.trick.clone(),
                        tag: client_tag,
                        thread: (event_handle, stop_flag),
                    }
                    .into();
                    let ret = unsafe { client.query(&IAudioRenderClient::IID, ppv) }.ok();
                    _ = self.outer.set(client);
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
    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows::core::Result<()> {
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
    ) -> windows::core::Result<()> {
        if periodinframes != 0 {
            info_tagged!(
                @self,
                "InitializeSharedAudioStream -> replacing period, current period: {periodinframes}"
            );
        }
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
            let align = (*pformat).nBlockAlign;
            let buf_len = if let Some(buf) = target_config
                .ring_buffer_len
                .get(&self.get_info().samplerate)
                .copied()
                && buf != 0
            {
                self.get_info()
                    .current_buffer_len
                    .max(buf.div_ceil(self.get_info().fundamental) * self.get_info().fundamental)
            } else {
                self.get_info().current_buffer_len * 10
            };
            let buffer = Arc::new(HeapRb::new(buf_len as usize * align as usize));
            _ = self.buffer.set(buffer.clone());
            self.align.set(AudioAlign::new(align as u32));
            if streamflags & AUDCLNT_STREAMFLAGS_EVENTCALLBACK == 0 {
                info_tagged!(@self, "Injecting event flag");
                streamflags |= AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
            } else {
                info_tagged!(@self, "Enabling inverse mode");
            }
            self.inner.InitializeSharedAudioStream(
                streamflags,
                self.get_info().current_buffer_len,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}
drop_boilerplate!(RedirectRingbufAudioClient);

#[implement(IRtwqAsyncCallback)]
struct RedirectRingbufThread {
    stop_flag: Arc<AtomicBool>,
    thread_id: u32,
    event_handle: HANDLE,
    app_handle: Arc<AtomicPtr<c_void>>,
    buffer: UnsafeCell<CachingCons<Arc<HeapRb<u8>>>>,
    client: IAudioClient3,
    align: AudioAlign<usize>,
    real_len: u32,
    inner: IAudioRenderClient,
    tag: Box<str>,
}
impl IRtwqAsyncCallback_Impl for RedirectRingbufThread_Impl {
    fn GetParameters(&self, pdwflags: *mut u32, pdwqueue: *mut u32) -> windows_core::Result<()> {
        unsafe {
            *pdwflags = 1; // MFASYNC_FAST_IO_PROCESSING_CALLBACK
            *pdwqueue = self.thread_id
        };
        Ok(())
    }
    fn Invoke(
        &self,
        pasyncresult: windows_core::Ref<IRtwqAsyncResult>,
    ) -> windows_core::Result<()> {
        if self.stop_flag.load(Ordering::Relaxed) {
            return Ok(());
        }
        let buffer = unsafe { &mut *self.buffer.get() };
        if buffer.is_empty() {
            let pad = unsafe { self.client.GetCurrentPadding()? };
            if pad == 0 {
                warn_tagged!(self.tag, "buffer is empty, underflow may happen!");
            } else {
                debug_tagged!(self.tag, "mid-buffer empty, data in client buffer: {pad}");
            }
            unsafe {
                if let Some(handle) = self.app_handle.load(Ordering::Relaxed).as_mut() {
                    SetEvent(HANDLE(handle)).ok();
                }
                RtwqPutWaitingWorkItem(self.event_handle, 1, pasyncresult.as_ref(), None)?;
            }
            return Ok(());
        }
        let read_len = self.align.bytes_to_frames(buffer.occupied_len());
        let write_len =
            read_len.min((self.real_len - unsafe { self.client.GetCurrentPadding()? }) as usize);
        let slice = unsafe {
            from_raw_parts_mut(
                self.inner.GetBuffer(write_len as u32)?,
                self.align.frames_to_bytes(write_len),
            )
        };
        buffer.pop_slice(slice);
        unsafe {
            self.inner.ReleaseBuffer(write_len as u32, 0)?;
            trace_tagged!(
                self.tag,
                "data in mid-buffer: {read_len}, written: {write_len}"
            );
            if let Some(handle) = self.app_handle.load(Ordering::Relaxed).as_mut() {
                SetEvent(HANDLE(handle)).ok();
            }
            RtwqPutWaitingWorkItem(self.event_handle, 1, pasyncresult.as_ref(), None)?;
        };
        Ok(())
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
            self.event_handle.free();
        }
        info_tagged!(self.tag, "Consumer thread stopped");
    }
}

#[implement(IAudioRenderClient)]
struct RedirectRingbufAudioRenderClient {
    buffer: UnsafeCell<CachingProd<Arc<HeapRb<u8>>>>,
    cache: UnsafeCell<Box<[u8]>>,
    align: AudioAlign<usize>,
    trick: Rc<Cell<bool>>,
    tag: Box<str>,
    thread: (HANDLE, Arc<AtomicBool>),
}
impl IAudioRenderClient_Impl for RedirectRingbufAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows::core::Result<*mut u8> {
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
    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows::core::Result<()> {
        if numframeswritten == 0 {
            warn_tagged!(
                self.tag,
                "no data written in this release call, overflow may happen!"
            );
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
            let written_len = buffer.push_slice(slice);
            debug_tagged!(
                self.tag,
                "ReleaseBuffer called, written: {}",
                self.align.bytes_to_frames(written_len)
            )
        };
        Ok(())
    }
}
impl Drop for RedirectRingbufAudioRenderClient {
    fn drop(&mut self) {
        info_tagged!(self.tag, "Stopping consumer thread");
        self.thread.1.store(true, Ordering::Relaxed);
        unsafe {
            SetEvent(self.thread.0).ok();
        }
    }
}

fn formatter(
    w: &mut dyn std::io::Write,
    _now: &mut DeferredNow,
    record: &Record,
) -> std::result::Result<(), std::io::Error> {
    write!(w, "{} [{}] ", record.level(), record.target())?;

    write!(w, "{}", record.args())
}

#[unsafe(export_name = "proxy")]
extern "C" fn proxy_dummy() {}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(_hinst: HANDLE, reason: u32, _reserved: *mut c_void) -> BOOL {
    match reason {
        DLL_PROCESS_ATTACH => {
            std::panic::set_hook(Box::new(|panic_info| {
                error!("{panic_info}");
            }));
            unsafe {
                HOOK_CO_CREATE_INSTANCE.enable().unwrap();
                HOOK_CO_CREATE_INSTANCE_EX.enable().unwrap();
            };
        }
        DLL_PROCESS_DETACH => unsafe {
            HOOK_CO_CREATE_INSTANCE.disable().unwrap();
            HOOK_CO_CREATE_INSTANCE_EX.disable().unwrap();
        },
        _ => (),
    };
    true.into()
}
