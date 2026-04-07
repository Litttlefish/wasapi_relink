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
use std::hint::{spin_loop, unreachable_unchecked};
use std::mem::transmute;
use std::os::raw::c_void;
use std::path::PathBuf;
use std::ptr::with_exposed_provenance_mut;
use std::slice::from_raw_parts_mut;
use std::sync::{Arc, LazyLock, Once, OnceLock, atomic::*};
use std::thread::{JoinHandle, spawn};

use windows::{
    Win32::{
        Foundation::*,
        Media::Audio::{DirectSound::*, Endpoints::*, *},
        Media::DirectShow::*,
        System::Com::{StructuredStorage::*, *},
        System::LibraryLoader::{GetModuleHandleW, GetProcAddress, LoadLibraryW},
        System::SystemServices::{DLL_PROCESS_ATTACH, DLL_PROCESS_DETACH},
        System::Threading::{
            AvRevertMmThreadCharacteristics, AvSetMmThreadCharacteristicsW, CreateEventW,
            GetCurrentThread, GetThreadDescription, INFINITE, SetEvent, SetThreadPriority,
            THREAD_PRIORITY_TIME_CRITICAL, WaitForSingleObject,
        },
        UI::Shell::PropertiesSystem::{IPropertyStore, IPropertyStore_Impl},
    },
    core::*,
};

// static OLE32_LIB: std::sync::LazyLock<Library> = std::sync::LazyLock::new(|| unsafe {
//     Library::new("ole32.dll").expect("Failed to load original ole32.dll")
// });

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
    #[inline]
    fn new_with_source(source: ConfigSource) -> Self {
        Self {
            source,
            ..Self::default()
        }
    }
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

#[allow(clippy::redundant_closure)]
static CONFIG: LazyLock<RedirectConfig> = LazyLock::new(|| RedirectConfig::load());

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

// unsafe extern "system" fn local_cocreateinstance(
//     rclsid: *const windows::core::GUID,
//     punkouter: *mut core::ffi::c_void,
//     dwclscontext: CLSCTX,
//     riid: *const windows::core::GUID,
//     ppv: *mut *mut core::ffi::c_void,
// ) -> windows::core::HRESULT {
//     link!("ole32.dll" "system" fn CoCreateInstance(rclsid : *const windows::core::GUID, punkouter : * mut core::ffi::c_void, dwclscontext : CLSCTX, riid : *const windows::core::GUID, ppv : *mut *mut core::ffi::c_void) -> windows::core::HRESULT);
//     unsafe { CoCreateInstance(rclsid, punkouter.param().abi(), dwclscontext, riid, ppv) }
// }

const LIB_NAME: PCWSTR = w!("ole32.dll");
const CO_CREATE: PCSTR = s!("CoCreateInstance");
const CO_CREATE_EX: PCSTR = s!("CoCreateInstanceEx");

const EVENT_CALLBACK: PCWSTR = w!("wasapi_relink Audio Callback Thread");
const PRO_AUDIO: PCWSTR = w!("Pro Audio");

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
            if ret.is_ok() {
                debug!("CoCreateInstance CLSCTX: {dwcls_context:?}");
                if let Ok(thread_desc) = GetThreadDescription(GetCurrentThread())
                    && !thread_desc.is_empty()
                    && let Ok(name) = thread_desc.to_string()
                    && KEYWORDS.iter().any(|keyword| name.contains(keyword))
                {
                    info!("Skipping SpecialK CoCreateInstance calls, thread name: {name}");
                } else {
                    info!(
                        "!!! Intercepted IMMDeviceEnumerator creation via CoCreateInstance, returning proxy !!!"
                    );
                    let proxy_enumerator: IMMDeviceEnumerator =
                        RedirectDeviceEnumerator::new(IMMDeviceEnumerator::from_raw(*ppv)).into();
                    *ppv = proxy_enumerator.into_raw();
                }
            } else {
                error!("CoCreateInstance call failed with HRESULT: {ret}")
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
            if hr.is_ok() {
                debug!("CoCreateInstanceEx CLSCTX: {dwclsctx:?}");
                if let Ok(thread_desc) = GetThreadDescription(GetCurrentThread())
                    && !thread_desc.is_empty()
                    && let Ok(name) = thread_desc.to_string()
                    && KEYWORDS.iter().any(|keyword| name.contains(keyword))
                {
                    info!("Skipping SpecialK CoCreateInstanceEx calls, thread name: {name}")
                } else {
                    for qi in from_raw_parts_mut(presults, dwcount as usize) {
                        if *qi.pIID == IMMDeviceEnumerator::IID && qi.hr.is_ok() {
                            info!(
                                "!!! Intercepted IMMDeviceEnumerator via CoCreateInstanceEx, replacing with proxy !!!"
                            );
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
                error!("CoCreateInstanceEx call failed with HRESULT: {hr}")
            }
        }
        hr
    }
}

#[repr(transparent)]
#[implement(IMMDeviceCollection)]
struct RedirectDeviceCollection {
    inner: IMMDeviceCollection,
}
impl IMMDeviceCollection_Impl for RedirectDeviceCollection_Impl {
    fn GetCount(&self) -> windows::core::Result<u32> {
        trace!("RedirectDeviceCollection::GetCount() called");
        unsafe { self.inner.GetCount() }
    }

    fn Item(&self, ndevice: u32) -> windows::core::Result<IMMDevice> {
        debug!("RedirectDeviceCollection::Item() -> wrapping, device {ndevice}");
        Ok(RedirectDevice::new(unsafe { self.inner.Item(ndevice)? }).into())
    }
}

#[inline]
unsafe fn assign<I: Interface>(ptr: *mut *mut c_void, component: I) -> windows::core::Result<()> {
    unsafe { component.query(&I::IID, ptr).ok() }
}

macro_rules! boilerplate {
    (
        $iid:expr,
        $ptr:expr,
        $self:ident,
        [
            $($interface:ty),* $(,)?
        ]
    ) => {
        match $iid {
            $(
                <$interface>::IID => {
                    unsafe { assign($ptr, $self.inner.GetService::<$interface>()?) }
                },
            )*
            _ => {
                error!("Called unimplemented service!");
                Err(E_NOINTERFACE.into())
            },
        }
    };
    (
        $iid:expr,
        $ptr:expr,
        $self:ident,
        $dwclsctx:expr,
        $pactivationparams:expr,
        [
            $($interface:ty),* $(,)?
        ]
    ) => {
        match $iid {
            $(
                <$interface>::IID => {
                    assign($ptr, $self.inner.Activate::<$interface>($dwclsctx, $pactivationparams)?)
                },
            )*
            _ => {
                error!("Called unimplemented object!");
                Err(E_NOINTERFACE.into())
            },
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
            info!("RedirectDevice::Activate() called, iid: {iid:?}");
            match iid {
                IAudioClient::IID | IAudioClient2::IID | IAudioClient3::IID => {
                    let inner: IAudioClient3 = self
                        .inner
                        .Activate::<IAudioClient3>(dwclsctx, Some(pactivationparams))?;
                    let dataflow = self.inner.cast::<IMMEndpoint>()?.GetDataFlow()?.into();
                    let proxy_unknown: IAudioClient3 = match CONFIG.get(dataflow).mode {
                        ClientMode::Normal => RedirectAudioClient::new(inner, dataflow).into(),
                        ClientMode::Compat => RedirectCompatAudioClient::new(
                            inner,
                            self.inner
                                .Activate::<IAudioClient3>(dwclsctx, Some(pactivationparams))?,
                            dataflow,
                        )
                        .into(),
                        ClientMode::Ringbuf => match dataflow {
                            DeviceDataFlow::Playback => {
                                RedirectRingbufAudioClient::new(inner, dataflow).into()
                            }
                            DeviceDataFlow::Capture => {
                                warn!(
                                    "Ringbuf mode doesn't work with capture client, switching to compat mode"
                                );
                                RedirectCompatAudioClient::new(
                                    inner,
                                    self.inner.Activate::<IAudioClient3>(
                                        dwclsctx,
                                        Some(pactivationparams),
                                    )?,
                                    dataflow,
                                )
                                .into()
                            }
                        },
                        ClientMode::Bypass => inner,
                    };
                    assign(ppinterface, proxy_unknown)
                }
                IAudioSessionManager::IID | IAudioSessionManager2::IID => assign(
                    ppinterface,
                    self.inner
                        .Activate::<IAudioSessionManager2>(dwclsctx, Some(pactivationparams))?,
                ),
                IDirectSound::IID | IDirectSound8::IID => {
                    error!("The program is using DSound, tool won't work!");
                    assign(
                        ppinterface,
                        self.inner
                            .Activate::<IDirectSound8>(dwclsctx, Some(pactivationparams))?,
                    )
                }
                IDirectSoundCapture::IID => {
                    error!("The program is using DSound, tool won't work!");
                    assign(
                        ppinterface,
                        self.inner
                            .Activate::<IDirectSoundCapture>(dwclsctx, Some(pactivationparams))?,
                    )
                }
                iid => boilerplate!(
                    iid,
                    ppinterface,
                    self,
                    dwclsctx,
                    Some(pactivationparams),
                    [
                        IAudioEndpointVolume,
                        IAudioMeterInformation,
                        IBaseFilter,
                        IDeviceTopology
                    ]
                ),
            }
        }
    }

    fn OpenPropertyStore(&self, stgmaccess: STGM) -> windows::core::Result<IPropertyStore> {
        debug!("RedirectDevice::OpenPropertyStore() -> wrapping");
        Ok(RedirectPropertyStore {
            inner: unsafe { self.inner.OpenPropertyStore(stgmaccess)? },
        }
        .into())
    }

    fn GetId(&self) -> windows::core::Result<PWSTR> {
        debug!("RedirectDevice::GetId() called");
        unsafe { self.inner.GetId() }
    }

    fn GetState(&self) -> windows::core::Result<DEVICE_STATE> {
        trace!("RedirectDevice::GetState() called");
        unsafe { self.inner.GetState() }
    }
}
impl IMMEndpoint_Impl for RedirectDevice_Impl {
    fn GetDataFlow(&self) -> windows::core::Result<EDataFlow> {
        trace!("RedirectDevice::GetDataFlow() called");
        unsafe { self.inner.cast::<IMMEndpoint>()?.GetDataFlow() }
    }
}

#[repr(transparent)]
#[implement(IPropertyStore)]
pub struct RedirectPropertyStore {
    inner: IPropertyStore,
}
#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IPropertyStore_Impl for RedirectPropertyStore_Impl {
    fn GetValue(&self, key: *const PROPERTYKEY) -> windows::core::Result<PROPVARIANT> {
        debug!(
            "RedirectPropertyStore::GetValue() called, key: {:?}",
            unsafe { key.as_ref() }
        );
        unsafe { self.inner.GetValue(key) }
    }

    fn GetCount(&self) -> windows::core::Result<u32> {
        trace!("RedirectPropertyStore::GetCount() called");
        unsafe { self.inner.GetCount() }
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> windows::core::Result<()> {
        trace!("RedirectPropertyStore::GetAt() called");
        unsafe { self.inner.GetAt(iprop, pkey) }
    }

    fn SetValue(
        &self,
        key: *const PROPERTYKEY,
        propvar: *const PROPVARIANT,
    ) -> windows::core::Result<()> {
        trace!("RedirectPropertyStore::SetValue() called");
        unsafe { self.inner.SetValue(key, propvar) }
    }

    fn Commit(&self) -> windows::core::Result<()> {
        trace!("RedirectPropertyStore::Commit() called");
        unsafe { self.inner.Commit() }
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
        debug!("RedirectDeviceEnumerator::EnumAudioEndpoints() -> wrapping, flow: {dataflow:?}");
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
            "RedirectDeviceEnumerator::GetDefaultAudioEndpoint() -> wrapping, flow: {dataflow:?}"
        );
        Ok(
            RedirectDevice::new(unsafe { self.inner.GetDefaultAudioEndpoint(dataflow, role)? })
                .into(),
        )
    }

    fn GetDevice(&self, pwstrid: &PCWSTR) -> windows::core::Result<IMMDevice> {
        info!("RedirectDeviceEnumerator::GetDevice() -> wrapping");
        Ok(RedirectDevice::new(unsafe { self.inner.GetDevice(*pwstrid)? }).into())
    }

    fn RegisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> windows::core::Result<()> {
        trace!("RedirectDeviceEnumerator::RegisterEndpointNotificationCallback() called");
        unsafe {
            self.inner
                .RegisterEndpointNotificationCallback(pclient.as_ref())
        }
    }

    fn UnregisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> windows::core::Result<()> {
        trace!("RedirectDeviceEnumerator::UnregisterEndpointNotificationCallback() called");

        unsafe {
            self.inner
                .UnregisterEndpointNotificationCallback(pclient.as_ref())
        }
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
    fn init(inner: &IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        let target_cfg = CONFIG.get(dataflow);
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
        let current_buffer_len = if target_cfg.target_buffer_dur_ms != 0 {
            calculate_buffer(
                samplerate,
                pfundamentalperiodinframes,
                target_cfg.target_buffer_dur_ms,
            )
            .clamp(pminperiodinframes, pmaxperiodinframes)
        } else {
            pminperiodinframes
        };
        info!(
            "Current period = {current_buffer_len}, Min period = {pminperiodinframes}, Max period = {pmaxperiodinframes}, Samplerate = {samplerate}"
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
    dataflow: DeviceDataFlow,
    raw_flag: Once,
}
impl RedirectClientInfo {
    fn new(dataflow: DeviceDataFlow) -> Self {
        Self {
            parameters: OnceCell::new(),
            dataflow,
            raw_flag: Once::new(),
        }
    }
    fn config(&self) -> &ClientConfig {
        CONFIG.get(self.dataflow)
    }
}

#[implement(IAudioClient3)]
struct RedirectAudioClient {
    inner: IAudioClient3,
    info: RedirectClientInfo,
}

impl RedirectAudioClient {
    fn new(inner: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        Self {
            inner,
            info: RedirectClientInfo::new(dataflow),
        }
    }
}
impl RedirectAudioClient {
    fn inner_info(&self) -> &Shared3Info {
        self.info
            .parameters
            .get_or_init(|| Shared3Info::init(&self.inner, self.info.dataflow))
    }
}
impl IAudioClient_Impl for RedirectAudioClient_Impl {
    fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: u32,
        hnsbufferduration: i64,
        hnsperiodicity: i64,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Initialize() -> redirecting to Low Latency Shared, direction: {:?}",
            self.info.dataflow
        );
        if sharemode == AUDCLNT_SHAREMODE_EXCLUSIVE {
            warn!("Rejected exclusive init request");
            return AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED.ok();
        }
        if streamflags & AUDCLNT_STREAMFLAGS_LOOPBACK == 0 {
            info!("Original dur: {hnsbufferduration} * 100ns");
            self.InitializeSharedAudioStream(streamflags, 0, pformat, audiosessionguid)
        } else {
            warn!("Loopback flow detected, bypassing");
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
        info!("RedirectAudioClient::GetBufferSize() called, buffer length: {buf}");
        Ok(buf)
    }

    fn GetStreamLatency(&self) -> windows::core::Result<i64> {
        info!("RedirectAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace!("RedirectAudioClient::GetCurrentPadding() called");
        let pad = unsafe { self.inner.GetCurrentPadding()? };
        if pad < self.inner_info().current_buffer_len {
            warn!("underflow may happen! current padding: {pad}");
        }
        Ok(pad)
    }

    fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: *mut *mut WAVEFORMATEX,
    ) -> windows::core::HRESULT {
        debug!("RedirectAudioClient::IsFormatSupported() called");
        unsafe {
            self.inner
                .IsFormatSupported(sharemode, pformat, Some(ppclosestmatch))
        }
    }

    fn GetMixFormat(&self) -> windows::core::Result<*mut WAVEFORMATEX> {
        info!(
            "RedirectAudioClient::GetMixFormat() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.GetMixFormat() }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::GetDevicePeriod() called, direction: {:?}",
            self.info.dataflow
        );
        let mut minimumdeviceperiod = 0;
        unsafe {
            self.inner
                .GetDevicePeriod(None, Some(&mut minimumdeviceperiod))?
        };
        if let Some(phnsdefaultdeviceperiod) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
            *phnsdefaultdeviceperiod = calculate_period(
                self.inner_info().samplerate,
                self.inner_info().current_buffer_len,
            )
            .max(minimumdeviceperiod)
        }
        if let Some(phnsminimumdeviceperiod) = unsafe { phnsminimumdeviceperiod.as_mut() } {
            *phnsminimumdeviceperiod = minimumdeviceperiod
        }
        // just assume no one will be silly here
        Ok(())
    }

    fn Start(&self) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Start() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Stop() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Reset() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info!("RedirectAudioClient::SetEventHandle() called");
        unsafe { self.inner.SetEventHandle(eventhandle) }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        let iid = unsafe { *riid };
        debug!(
            "RedirectAudioClient::GetService() called, iid: {iid:?}, direction: {:?}",
            self.info.dataflow
        );
        boilerplate!(
            iid,
            ppv,
            self,
            [
                IAudioSessionControl,
                IAudioRenderClient,
                IAudioCaptureClient,
                IAudioClientDuckingControl,
                IAudioClock,
                IChannelAudioVolume,
                ISimpleAudioVolume,
                IAudioStreamVolume
            ]
        )
    }
}

impl IAudioClient2_Impl for RedirectAudioClient_Impl {
    fn IsOffloadCapable(&self, category: AUDIO_STREAM_CATEGORY) -> windows::core::Result<BOOL> {
        info!("RedirectAudioClient::IsOffloadCapable() called");
        unsafe { self.inner.IsOffloadCapable(category) }
    }

    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows::core::Result<()> {
        info!("RedirectAudioClient::SetClientProperties() called");
        if self.info.config().raw {
            self.info.raw_flag.call_once(|| info!("Applying raw flag"));
            let option = &mut unsafe { *pproperties }.Options;
            if option.contains(AUDCLNT_STREAMOPTIONS_RAW) {
                warn!("This stream already has raw flag!")
            } else {
                *option |= AUDCLNT_STREAMOPTIONS_RAW
            }
        }
        unsafe { self.inner.SetClientProperties(pproperties) }
    }

    fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: BOOL,
        phnsminbufferduration: *mut i64,
        phnsmaxbufferduration: *mut i64,
    ) -> windows::core::Result<()> {
        info!("RedirectAudioClient::GetBufferSizeLimits() called");
        unsafe {
            self.inner.GetBufferSizeLimits(
                pformat,
                beventdriven.into(),
                phnsminbufferduration,
                phnsmaxbufferduration,
            )
        }
    }
}

impl IAudioClient3_Impl for RedirectAudioClient_Impl {
    fn GetSharedModeEnginePeriod(
        &self,
        pformat: *const WAVEFORMATEX,
        pdefaultperiodinframes: *mut u32,
        pfundamentalperiodinframes: *mut u32,
        pminperiodinframes: *mut u32,
        pmaxperiodinframes: *mut u32,
    ) -> windows::core::Result<()> {
        info!("RedirectAudioClient::GetSharedModeEnginePeriod() called");
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
        info!("RedirectAudioClient::GetCurrentSharedModeEnginePeriod() called");
        unsafe {
            self.inner
                .GetCurrentSharedModeEnginePeriod(ppformat, pcurrentperiodinframes)
        }
    }

    fn InitializeSharedAudioStream(
        &self,
        streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> windows::core::Result<()> {
        if periodinframes != 0 {
            info!(
                "RedirectAudioClient::InitializeSharedAudioStream() -> replacing period, current period: {periodinframes}, direction: {:?}",
                self.info.dataflow
            );
        }
        unsafe {
            if self.info.config().raw && !self.info.raw_flag.is_completed() {
                info!("Applying raw flag");
                let properties = AudioClientProperties {
                    cbSize: size_of::<AudioClientProperties>() as u32,
                    Options: AUDCLNT_STREAMOPTIONS_RAW,
                    ..AudioClientProperties::default()
                };
                self.inner.SetClientProperties(&properties)?;
            }
            self.inner.InitializeSharedAudioStream(
                streamflags,
                self.inner_info().current_buffer_len,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}

#[repr(u8)]
enum TrickState {
    Tricking,
    Filled,
    Transparent,
}

#[implement(IAudioClient3)]
struct RedirectCompatAudioClient {
    inner: IAudioClient3,
    hooker: IAudioClient3,
    info: RedirectClientInfo,
    trick: Arc<AtomicU8>,
    align: Cell<u32>,
}

impl RedirectCompatAudioClient {
    fn new(inner: IAudioClient3, hooker: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        Self {
            inner,
            hooker,
            info: RedirectClientInfo::new(dataflow),
            trick: AtomicU8::default().into(),
            align: 0.into(),
        }
    }
}
impl RedirectCompatAudioClient {
    fn hooker_info(&self) -> &Shared3Info {
        self.info
            .parameters
            .get_or_init(|| Shared3Info::init(&self.hooker, self.info.dataflow))
    }
}
impl IAudioClient_Impl for RedirectCompatAudioClient_Impl {
    fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: u32,
        hnsbufferduration: i64,
        hnsperiodicity: i64,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Initialize() -> redirecting to hooker Shared with small buffer, direction: {:?}",
            self.info.dataflow
        );
        if sharemode == AUDCLNT_SHAREMODE_EXCLUSIVE {
            warn!("Rejected exclusive init request");
            return AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED.ok();
        }
        if streamflags & AUDCLNT_STREAMFLAGS_LOOPBACK == 0 {
            info!("Original duration: {hnsbufferduration} * 100ns");
            let calculated_dur = self
                .info
                .config()
                .compat_buffer_len
                .get(&self.hooker_info().samplerate)
                .copied()
                .unwrap_or_default();
            info!("Inner duration = {calculated_dur} * 100ns");
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
            warn!("Loopback flow detected, bypassing");
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
        info!("RedirectCompatAudioClient::GetBufferSize() called, buffer length: {buf}");
        Ok(buf)
    }

    fn GetStreamLatency(&self) -> windows::core::Result<i64> {
        info!("RedirectCompatAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace!("RedirectCompatAudioClient::GetCurrentPadding() called");
        let pad = unsafe { self.inner.GetCurrentPadding()? };
        if pad < self.hooker_info().current_buffer_len {
            warn!("underflow may happen! current padding: {pad}");
        }
        Ok(pad)
    }

    fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: *mut *mut WAVEFORMATEX,
    ) -> HRESULT {
        debug!("RedirectCompatAudioClient::IsFormatSupported() called");
        unsafe {
            self.inner
                .IsFormatSupported(sharemode, pformat, Some(ppclosestmatch))
        }
    }

    fn GetMixFormat(&self) -> windows::core::Result<*mut WAVEFORMATEX> {
        info!(
            "RedirectCompatAudioClient::GetMixFormat() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.GetMixFormat() }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::GetDevicePeriod() called, direction: {:?}",
            self.info.dataflow
        );
        let mut minimumdeviceperiod = 0;
        unsafe {
            self.inner
                .GetDevicePeriod(None, Some(&mut minimumdeviceperiod))?
        };
        if let Some(phnsdefaultdeviceperiod) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
            *phnsdefaultdeviceperiod = calculate_period(
                self.hooker_info().samplerate,
                self.hooker_info().current_buffer_len,
            )
            .max(minimumdeviceperiod)
        }
        if let Some(phnsminimumdeviceperiod) = unsafe { phnsminimumdeviceperiod.as_mut() } {
            *phnsminimumdeviceperiod = minimumdeviceperiod
        }
        // just assume no one will be silly here
        Ok(())
    }

    fn Start(&self) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Start() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe {
            self.trick
                .store(TrickState::Transparent as u8, Ordering::Relaxed);
            _ = self.hooker.Start();
            self.inner.Start()
        }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Stop() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe {
            _ = self.hooker.Stop();
            self.inner.Stop()
        }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Reset() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe {
            self.trick
                .store(TrickState::Tricking as u8, Ordering::Relaxed);
            _ = self.hooker.Reset();
            self.inner.Reset()
        }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info!("RedirectCompatAudioClient::SetEventHandle() called");
        unsafe {
            _ = self.hooker.SetEventHandle(eventhandle);
            self.inner.SetEventHandle(eventhandle)
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        let iid = unsafe { *riid };
        debug!(
            "RedirectCompatAudioClient::GetService() called, iid: {iid:?}, direction: {:?}",
            self.info.dataflow
        );

        match iid {
            IAudioRenderClient::IID => {
                debug!("Returned RedirectCompatAudioRenderClient");
                unsafe {
                    let hooker_buffer_len = self.hooker.GetBufferSize().unwrap_or_else(|_| {
                        self.hooker_info().current_buffer_len + self.hooker_info().fundamental
                    });
                    let service: IAudioRenderClient = RedirectCompatAudioRenderClient {
                        inner: self.inner.GetService::<IAudioRenderClient>()?,
                        trick_buffer: vec![
                            0;
                            hooker_buffer_len as usize * self.align.get() as usize
                        ]
                        .into_boxed_slice()
                        .into(),
                        trick: self.trick.clone(),
                        align: AudioAlign::new(self.align.get()),
                        hooker_buffer_len,
                    }
                    .into();

                    service.query(&IAudioRenderClient::IID, ppv).ok()
                }
            }

            iid => boilerplate!(
                iid,
                ppv,
                self,
                [
                    IAudioSessionControl,
                    IAudioCaptureClient,
                    IAudioClientDuckingControl,
                    IAudioClock,
                    IChannelAudioVolume,
                    ISimpleAudioVolume,
                    IAudioStreamVolume
                ]
            ),
        }
    }
}

impl IAudioClient2_Impl for RedirectCompatAudioClient_Impl {
    fn IsOffloadCapable(&self, category: AUDIO_STREAM_CATEGORY) -> windows::core::Result<BOOL> {
        info!("RedirectCompatAudioClient::IsOffloadCapable() called");
        unsafe { self.inner.IsOffloadCapable(category) }
    }

    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows::core::Result<()> {
        info!("RedirectCompatAudioClient::SetClientProperties() called");
        if self.info.config().raw {
            self.info.raw_flag.call_once(|| info!("Applying raw flag"));
            let option = &mut unsafe { *pproperties }.Options;
            if option.contains(AUDCLNT_STREAMOPTIONS_RAW) {
                warn!("This stream already has raw flag!")
            } else {
                *option |= AUDCLNT_STREAMOPTIONS_RAW
            }
        }
        unsafe { self.inner.SetClientProperties(pproperties) }
    }

    fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: BOOL,
        phnsminbufferduration: *mut i64,
        phnsmaxbufferduration: *mut i64,
    ) -> windows::core::Result<()> {
        info!("RedirectCompatAudioClient::GetBufferSizeLimits() called");
        unsafe {
            self.inner.GetBufferSizeLimits(
                pformat,
                beventdriven.into(),
                phnsminbufferduration,
                phnsmaxbufferduration,
            )
        }
    }
}

impl IAudioClient3_Impl for RedirectCompatAudioClient_Impl {
    fn GetSharedModeEnginePeriod(
        &self,
        pformat: *const WAVEFORMATEX,
        pdefaultperiodinframes: *mut u32,
        pfundamentalperiodinframes: *mut u32,
        pminperiodinframes: *mut u32,
        pmaxperiodinframes: *mut u32,
    ) -> windows::core::Result<()> {
        info!("RedirectCompatAudioClient::GetSharedModeEnginePeriod() called");
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
        info!("RedirectCompatAudioClient::GetCurrentSharedModeEnginePeriod() called");
        unsafe {
            self.inner
                .GetCurrentSharedModeEnginePeriod(ppformat, pcurrentperiodinframes)
        }
    }

    fn InitializeSharedAudioStream(
        &self,
        streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> windows::core::Result<()> {
        if self.info.config().raw && !self.info.raw_flag.is_completed() {
            info!("Applying raw flag");
            let properties = AudioClientProperties {
                cbSize: size_of::<AudioClientProperties>() as u32,
                Options: AUDCLNT_STREAMOPTIONS_RAW,
                ..AudioClientProperties::default()
            };
            unsafe { self.inner.SetClientProperties(&properties) }?;
        }
        self.align.set(unsafe { (*pformat).nBlockAlign as u32 });
        if periodinframes != 0 {
            info!(
                "RedirectCompatAudioClient::InitializeSharedAudioStream() called, direction: {:?}",
                self.info.dataflow
            );
            info!("Original period: {periodinframes}");
        }
        unsafe {
            self.inner.InitializeSharedAudioStream(
                streamflags,
                self.hooker_info().current_buffer_len,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}

#[implement(IAudioRenderClient)]
struct RedirectCompatAudioRenderClient {
    inner: IAudioRenderClient,
    trick_buffer: UnsafeCell<Box<[u8]>>,
    trick: Arc<AtomicU8>,
    align: AudioAlign<u32>,
    hooker_buffer_len: u32,
}
impl RedirectCompatAudioRenderClient {
    #[inline]
    fn apply_data(&self, len: u32, dwflags: u32) -> windows::core::Result<()> {
        let read_len = self.align.frames_to_bytes(len) as usize;
        unsafe {
            let slice_to_write = from_raw_parts_mut(self.inner.GetBuffer(len)?, read_len);
            let slice = &(&(*self.trick_buffer.get()))[..read_len];
            slice_to_write.copy_from_slice(slice);
            self.inner.ReleaseBuffer(len, dwflags)
        }
    }
}
impl IAudioRenderClient_Impl for RedirectCompatAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows::core::Result<*mut u8> {
        if self.trick.load(Ordering::Relaxed) != TrickState::Transparent as u8 {
            info!(
                "RedirectCompatAudioRenderClient::GetBuffer() called, requested: {numframesrequested}, tricking"
            );
            Ok(unsafe { &mut *self.trick_buffer.get() }.as_mut_ptr())
        } else {
            unsafe { self.inner.GetBuffer(numframesrequested) }
        }
    }

    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows::core::Result<()> {
        if numframeswritten == 0 {
            warn!("no data written in this release call, overflow may happen!");
            return unsafe { self.inner.ReleaseBuffer(0, 2) };
        }
        match unsafe { transmute::<u8, TrickState>(self.trick.load(Ordering::Relaxed)) } {
            TrickState::Tricking => {
                info!(
                    "RedirectCompatAudioRenderClient::ReleaseBuffer() called, written: {numframeswritten}, tricking"
                );
                info!(
                    "filling silent buffer, {} samples filled",
                    self.hooker_buffer_len
                );
                self.trick
                    .store(TrickState::Filled as u8, Ordering::Relaxed);
                self.apply_data(self.hooker_buffer_len, dwflags)
            }
            TrickState::Filled => {
                info!(
                    "RedirectCompatAudioRenderClient::ReleaseBuffer() called, written: {numframeswritten}"
                );
                if dwflags == 2 {
                    info!("already filled, discarding");
                    Ok(())
                } else {
                    self.apply_data(numframeswritten, dwflags)
                }
            }
            TrickState::Transparent => unsafe {
                self.inner.ReleaseBuffer(numframeswritten, dwflags)
            },
        }
    }
}

struct RenderClientBuildInfo {
    producer: CachingProd<Arc<HeapRb<u8>>>,
    consumer: CachingCons<Arc<HeapRb<u8>>>,
    buffer_length: usize,
    thread_handle: HANDLE,
}

enum BuildStatus {
    Building(Option<RenderClientBuildInfo>),
    Done(IAudioRenderClient),
}

#[implement(IAudioClient3)]
struct RedirectRingbufAudioClient {
    inner: IAudioClient3,
    info: RedirectClientInfo,
    buffer: OnceCell<Arc<HeapRb<u8>>>,
    align: Cell<AudioAlign<u32>>,
    build_info: UnsafeCell<BuildStatus>,
    trick: Arc<AtomicBool>,
    app_handle: Arc<AtomicPtr<c_void>>,
}

impl RedirectRingbufAudioClient {
    fn new(inner: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        Self {
            inner,
            info: RedirectClientInfo::new(dataflow),
            buffer: OnceCell::new(),
            align: AudioAlign::Normal(0).into(),
            build_info: BuildStatus::Building(None).into(),
            trick: AtomicBool::new(true).into(),
            app_handle: AtomicPtr::default().into(),
        }
    }
    fn inner_info(&self) -> &Shared3Info {
        self.info
            .parameters
            .get_or_init(|| Shared3Info::init(&self.inner, self.info.dataflow))
    }
}

impl IAudioClient_Impl for RedirectRingbufAudioClient_Impl {
    fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: u32,
        hnsbufferduration: i64,
        hnsperiodicity: i64,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> windows::core::Result<()> {
        info!(
            "RedirectRingbufAudioClient::Initialize() -> Adding ring buffer, direction: {:?}",
            self.info.dataflow
        );
        if sharemode == AUDCLNT_SHAREMODE_EXCLUSIVE {
            warn!("Rejected exclusive init request");
            return AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED.ok();
        }
        if streamflags & AUDCLNT_STREAMFLAGS_LOOPBACK == 0 {
            info!("Original dur: {hnsbufferduration} * 100ns");
            self.InitializeSharedAudioStream(streamflags, 0, pformat, audiosessionguid)
        } else {
            warn!("Loopback flow detected, bypassing");
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
            info!("RedirectRingbufAudioClient::GetBufferSize() called, buffer length: {buf}");
            Ok(buf)
        } else {
            unsafe { self.inner.GetBufferSize() }
        }
    }

    fn GetStreamLatency(&self) -> windows::core::Result<i64> {
        info!("RedirectRingbufAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace!("RedirectRingbufAudioClient::GetCurrentPadding() called");
        if let Some(buf) = self.buffer.get() {
            Ok(self.align.get().bytes_to_frames(buf.occupied_len() as u32))
        } else {
            unsafe { self.inner.GetCurrentPadding() }
        }
    }

    fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: *mut *mut WAVEFORMATEX,
    ) -> HRESULT {
        debug!("RedirectRingbufAudioClient::IsFormatSupported() called");
        unsafe {
            self.inner
                .IsFormatSupported(sharemode, pformat, Some(ppclosestmatch))
        }
    }

    fn GetMixFormat(&self) -> windows::core::Result<*mut WAVEFORMATEX> {
        info!(
            "RedirectRingbufAudioClient::GetMixFormat() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.GetMixFormat() }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> windows::core::Result<()> {
        info!(
            "RedirectRingbufAudioClient::GetDevicePeriod() called, direction: {:?}",
            self.info.dataflow
        );
        let mut minimumdeviceperiod = 0;
        unsafe {
            self.inner
                .GetDevicePeriod(None, Some(&mut minimumdeviceperiod))?
        };
        if let Some(phnsdefaultdeviceperiod) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
            *phnsdefaultdeviceperiod = calculate_period(
                self.inner_info().samplerate,
                self.inner_info().current_buffer_len,
            )
            .max(minimumdeviceperiod)
        }
        if let Some(phnsminimumdeviceperiod) = unsafe { phnsminimumdeviceperiod.as_mut() } {
            *phnsminimumdeviceperiod = minimumdeviceperiod
        }
        // just assume no one will be silly here
        Ok(())
    }

    fn Start(&self) -> windows::core::Result<()> {
        info!(
            "RedirectRingbufAudioClient::Start() called, direction: {:?}",
            self.info.dataflow
        );
        self.trick.store(false, Ordering::Relaxed);
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info!(
            "RedirectRingbufAudioClient::Stop() called, direction: {:?}",
            self.info.dataflow
        );
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info!(
            "RedirectRingbufAudioClient::Reset() called, direction: {:?}",
            self.info.dataflow
        );
        self.trick.store(true, Ordering::Relaxed);
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info!("RedirectRingbufAudioClient::SetEventHandle() called");
        if self.buffer.get().is_some() {
            self.app_handle.store(eventhandle.0, Ordering::Relaxed);
            Ok(())
        } else {
            unsafe { self.inner.SetEventHandle(eventhandle) }
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        let iid = unsafe { *riid };
        debug!(
            "RedirectRingbufAudioClient::GetService() called, iid: {iid:?}, direction: {:?}",
            self.info.dataflow
        );
        match iid {
            IAudioRenderClient::IID => {
                let build_info = unsafe { &mut *self.build_info.get() };
                match build_info {
                    BuildStatus::Building(None) => unreachable!("How did you do this?"),
                    BuildStatus::Building(info) => {
                        let info = info.take().unwrap();

                        let stop_flag: Arc<AtomicBool> = AtomicBool::default().into();

                        let raw = self.inner.as_raw().expose_provenance();

                        let app_thread = self.app_handle.clone();

                        let h = info.thread_handle.0.expose_provenance();

                        let thread_stop_flag = stop_flag.clone();

                        let align = self.align.get().as_usize();

                        let thread = spawn(move || {
                            let app_thread = app_thread;
                            let mut task_index = 0u32;
                            let client_ptr = with_exposed_provenance_mut::<c_void>(raw);
                            let client_ref = unsafe {
                                IAudioClient3::from_raw_borrowed(&client_ptr).unwrap_unchecked()
                            };
                            let mmcss_handle = unsafe {
                                info!("Registering MMCSS Pro Audio thread");
                                AvSetMmThreadCharacteristicsW(PRO_AUDIO, &mut task_index)
                            }
                            .inspect_err(|e| {
                                error!(
                                    "Failed to register MMCSS: {e}, falling back to TimeCritical"
                                );
                                unsafe {
                                    SetThreadPriority(
                                        GetCurrentThread(),
                                        THREAD_PRIORITY_TIME_CRITICAL,
                                    )
                                }
                                .unwrap_or_else(|e| {
                                    error!("Failed to set Critical: {e}");
                                });
                            })
                            .ok();
                            callback(
                                (
                                    HANDLE(with_exposed_provenance_mut::<c_void>(h)),
                                    thread_stop_flag,
                                    app_thread,
                                ),
                                info.consumer,
                                client_ref,
                                align,
                            );
                            if let Some(handle) = mmcss_handle {
                                unsafe { AvRevertMmThreadCharacteristics(handle) }.ok();
                            }
                        });
                        let client: IAudioRenderClient = RedirectRingbufAudioRenderClient {
                            buffer: (
                                info.producer,
                                vec![0u8; align.frames_to_bytes(info.buffer_length)]
                                    .into_boxed_slice(),
                            )
                                .into(),
                            align,
                            thread: (info.thread_handle, Some(thread), stop_flag),
                            trick: self.trick.clone(),
                        }
                        .into();
                        let ret = unsafe { client.query(&IAudioRenderClient::IID, ppv) }.ok();
                        *build_info = BuildStatus::Done(client);
                        ret
                    }
                    BuildStatus::Done(client) => unsafe {
                        client.query(&IAudioRenderClient::IID, ppv).ok()
                    },
                }
            }
            _ => boilerplate!(
                iid,
                ppv,
                self,
                [
                    IAudioSessionControl,
                    IAudioCaptureClient,
                    IAudioClientDuckingControl,
                    IAudioClock,
                    IChannelAudioVolume,
                    ISimpleAudioVolume,
                    IAudioStreamVolume
                ]
            ),
        }
    }
}

impl IAudioClient2_Impl for RedirectRingbufAudioClient_Impl {
    fn IsOffloadCapable(&self, category: AUDIO_STREAM_CATEGORY) -> windows::core::Result<BOOL> {
        info!("RedirectRingbufAudioClient::IsOffloadCapable() called");
        unsafe { self.inner.IsOffloadCapable(category) }
    }

    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows::core::Result<()> {
        info!("RedirectRingbufAudioClient::SetClientProperties() called");
        if self.info.config().raw {
            self.info.raw_flag.call_once(|| info!("Applying raw flag"));

            let option = &mut unsafe { *pproperties }.Options;
            if option.contains(AUDCLNT_STREAMOPTIONS_RAW) {
                warn!("This stream already has raw flag!")
            } else {
                *option |= AUDCLNT_STREAMOPTIONS_RAW
            }
        }
        unsafe { self.inner.SetClientProperties(pproperties) }
    }

    fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: BOOL,
        phnsminbufferduration: *mut i64,
        phnsmaxbufferduration: *mut i64,
    ) -> windows::core::Result<()> {
        info!("RedirectRingbufAudioClient::GetBufferSizeLimits() called");
        unsafe {
            self.inner.GetBufferSizeLimits(
                pformat,
                beventdriven.into(),
                phnsminbufferduration,
                phnsmaxbufferduration,
            )
        }
    }
}

impl IAudioClient3_Impl for RedirectRingbufAudioClient_Impl {
    fn GetSharedModeEnginePeriod(
        &self,
        pformat: *const WAVEFORMATEX,
        pdefaultperiodinframes: *mut u32,
        pfundamentalperiodinframes: *mut u32,
        pminperiodinframes: *mut u32,
        pmaxperiodinframes: *mut u32,
    ) -> windows::core::Result<()> {
        info!("RedirectRingbufAudioClient::GetSharedModeEnginePeriod() called");
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
        info!("RedirectRingbufAudioClient::GetCurrentSharedModeEnginePeriod() called");
        unsafe {
            self.inner
                .GetCurrentSharedModeEnginePeriod(ppformat, pcurrentperiodinframes)
        }
    }

    fn InitializeSharedAudioStream(
        &self,
        mut streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: *const GUID,
    ) -> windows::core::Result<()> {
        if periodinframes != 0 {
            info!(
                "RedirectRingbufAudioClient::InitializeSharedAudioStream() -> replacing period, current period: {periodinframes}, direction: {:?}",
                self.info.dataflow
            );
        }
        unsafe {
            let target_config = self.info.config();
            if target_config.raw && !self.info.raw_flag.is_completed() {
                info!("Applying raw flag");
                let properties = AudioClientProperties {
                    cbSize: size_of::<AudioClientProperties>() as u32,
                    Options: AUDCLNT_STREAMOPTIONS_RAW,
                    ..AudioClientProperties::default()
                };
                self.inner.SetClientProperties(&properties)?;
            }
            let align = (*pformat).nBlockAlign;
            let callback_handle = CreateEventW(None, false, false, EVENT_CALLBACK)?;
            let buf_len = if let Some(buf) = target_config
                .ring_buffer_len
                .get(&self.inner_info().samplerate)
                .copied()
                && buf != 0
            {
                self.inner_info().current_buffer_len.max(
                    buf.div_ceil(self.inner_info().fundamental) * self.inner_info().fundamental,
                )
            } else {
                self.inner_info().current_buffer_len * 10
            };
            let buffer = Arc::new(HeapRb::new(buf_len as usize * align as usize));
            let (producer, consumer) = buffer.clone().split();

            _ = self.buffer.set(buffer);
            self.align.set(AudioAlign::new(align as u32));
            let build_info = &mut *self.build_info.get();
            *build_info = BuildStatus::Building(Some(RenderClientBuildInfo {
                producer,
                consumer,
                buffer_length: buf_len as usize,
                thread_handle: callback_handle,
            }));

            if streamflags & AUDCLNT_STREAMFLAGS_EVENTCALLBACK == 0 {
                info!("Injecting event flag");
                streamflags |= AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
            } else {
                info!("Enabling inverse mode");
            }
            self.inner.InitializeSharedAudioStream(
                streamflags,
                self.inner_info().current_buffer_len,
                pformat,
                Some(audiosessionguid),
            )?;
            self.inner.SetEventHandle(callback_handle)
        }
    }
}

fn callback(
    thread_info: (HANDLE, Arc<AtomicBool>, Arc<AtomicPtr<c_void>>),
    mut buffer: CachingCons<Arc<HeapRb<u8>>>,
    client: &IAudioClient3,
    align: AudioAlign<usize>,
) {
    let real_len = unsafe { client.GetBufferSize().unwrap() } as usize;
    let inner = unsafe { client.GetService::<IAudioRenderClient>().unwrap() };
    while buffer.is_empty() {
        spin_loop();
    }
    while !thread_info.1.load(Ordering::Relaxed) {
        let (read_len, pad) = (buffer.occupied_len(), unsafe {
            client.GetCurrentPadding().unwrap()
        });
        if read_len == 0 {
            if pad == 0 {
                warn!("buffer is empty, underflow may happen!");
            } else {
                debug!("mid-buffer empty, data in client buffer: {pad}");
            }
        }
        let write_len = align.bytes_to_frames(read_len).min(real_len - pad as usize);
        let slice: &mut [u8] = unsafe {
            from_raw_parts_mut(
                inner.GetBuffer(write_len as u32).unwrap(),
                align.frames_to_bytes(write_len),
            )
        };
        buffer.pop_slice(slice);
        unsafe {
            inner.ReleaseBuffer(write_len as u32, 0).unwrap();
            trace!("data in mid-buffer: {read_len}, written: {write_len}");
            if let Some(handle) = thread_info.2.load(Ordering::Relaxed).as_mut() {
                SetEvent(HANDLE(handle)).ok();
            }
            WaitForSingleObject(thread_info.0, INFINITE);
        };
    }
}

type RbRenderBuffer = (CachingProd<Arc<HeapRb<u8>>>, Box<[u8]>);

#[implement(IAudioRenderClient)]
struct RedirectRingbufAudioRenderClient {
    buffer: UnsafeCell<RbRenderBuffer>,
    align: AudioAlign<usize>,
    thread: (
        HANDLE,
        Option<JoinHandle<()>>,
        Arc<AtomicBool>, /* stop flag */
    ),
    trick: Arc<AtomicBool>,
}
impl IAudioRenderClient_Impl for RedirectRingbufAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows::core::Result<*mut u8> {
        if self.trick.load(Ordering::Relaxed) {
            info!(
                "RedirectRingbufAudioRenderClient::GetBuffer() called, requested: {numframesrequested}"
            );
        }
        trace!("GetBuffer called, requested: {numframesrequested}");
        Ok(unsafe { &mut *self.buffer.get() }.1.as_mut_ptr())
    }
    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows::core::Result<()> {
        if numframeswritten == 0 {
            warn!("no data written in this release call, overflow may happen!");
            return Ok(());
        }
        if self.trick.load(Ordering::Relaxed) {
            info!(
                "RedirectRingbufAudioRenderClient::ReleaseBuffer() called, written: {numframeswritten}"
            );
            if dwflags == 2 {
                info!("discarding silent data");
                return Ok(());
            }
        }
        let (buffer, temp_buffer) = unsafe { &mut *self.buffer.get() };
        let slice = &mut temp_buffer[..(self.align.frames_to_bytes(numframeswritten as usize))];
        if dwflags == 2 {
            slice.fill(0);
        }
        let written_len = buffer.push_slice(slice);
        debug!(
            "ReleaseBuffer called, written: {}",
            self.align.bytes_to_frames(written_len)
        );
        Ok(())
    }
}

impl Drop for RedirectRingbufAudioRenderClient {
    fn drop(&mut self) {
        self.thread.2.store(true, Ordering::Relaxed);
        unsafe {
            SetEvent(self.thread.0).ok();
        }
        if let Some(thread) = self.thread.1.take() {
            let join_result = thread.join();
            match join_result {
                Ok(()) => {
                    info!("Consumer thread stopped");
                }
                Err(e) => {
                    error!("Consumer thread panicked on shutdown: {:?}", e);
                }
            }
        }
        unsafe { self.thread.0.free() };
    }
}

#[unsafe(export_name = "proxy")]
extern "C" fn proxy_dummy() {}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(_hinst: HANDLE, reason: u32, _reserved: *mut c_void) -> BOOL {
    match reason {
        DLL_PROCESS_ATTACH => {
            unsafe {
                HOOK_CO_CREATE_INSTANCE.enable().unwrap();
                HOOK_CO_CREATE_INSTANCE_EX.enable().unwrap();
            };
            spawn(|| {
                // let _logger = Logger::try_with_env_or_str("info")
                //     .unwrap()
                //     .log_to_stdout()
                //     .start();
                std::panic::set_hook(Box::new(|panic_info| {
                    error!("{panic_info}");
                }));
                let logger = Logger::with(<ConfigLogLevel as Into<LevelFilter>>::into(
                    CONFIG.log_level,
                ))
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
        DLL_PROCESS_DETACH => unsafe {
            HOOK_CO_CREATE_INSTANCE.disable().unwrap();
            HOOK_CO_CREATE_INSTANCE_EX.disable().unwrap();
        },
        _ => (),
    };
    true.into()
}
