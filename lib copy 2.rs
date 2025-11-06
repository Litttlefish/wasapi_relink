#[cfg(test)]
mod config_test;

#[allow(unused_imports)]
#[warn(clippy::single_component_path_imports)]
use auto_allocator;
// use openal_binds::*;
use flexi_logger::*;
use log::*;
use retour::GenericDetour;
use ringbuf::HeapRb;
use ringbuf::traits::*;
use serde::*;
use std::collections::VecDeque;
use std::os::raw::c_void;
use std::os::windows::io::AsRawHandle;
use std::os::windows::io::RawHandle;
use std::path::PathBuf;
use std::ptr::null_mut;
use std::sync::Arc;
use std::sync::LazyLock;
use std::sync::Mutex;
use std::sync::RwLock;
use std::sync::atomic::AtomicBool;
use std::sync::atomic::AtomicUsize;
use std::thread::JoinHandle;
use std::thread::spawn;
use std::{cell::UnsafeCell, fmt::Display};
use windows::Win32::System::Threading::CREATE_EVENT;
use windows::Win32::System::Threading::CreateEventA;
use windows::Win32::System::Threading::CreateEventExA;
use windows::Win32::System::Threading::EVENT_MODIFY_STATE;
use windows::Win32::System::Threading::INFINITE;
use windows::Win32::System::Threading::PROCESS_SYNCHRONIZE;
use windows::Win32::System::Threading::SYNCHRONIZATION_SYNCHRONIZE;
use windows::Win32::System::Threading::SetEvent;
use windows::Win32::System::Threading::WaitForSingleObject;

use windows::{
    Win32::{
        Foundation::*,
        Media::Audio::DirectSound::*,
        Media::Audio::*,
        System::Com::StructuredStorage::*,
        System::Com::*,
        System::LibraryLoader::{GetModuleHandleA, GetProcAddress, LoadLibraryA},
        System::SystemServices::{DLL_PROCESS_ATTACH, DLL_PROCESS_DETACH},
        System::Threading::{GetCurrentThread, GetThreadDescription},
        UI::Shell::PropertiesSystem::{IPropertyStore, IPropertyStore_Impl},
    },
    core::*,
};

// static OLE32_LIB: std::sync::LazyLock<Library> = std::sync::LazyLock::new(|| unsafe {
//     Library::new("ole32.dll").expect("Failed to load original ole32.dll")
// });

#[derive(Debug, Default)]
enum ConfigSource {
    #[default]
    Success,
    NoParse,
    NoFile,
}

#[derive(Debug, Default, Serialize, Deserialize)]
enum ConfigLogLevel {
    Trace,
    Debug,
    #[default]
    Info,
    Warn,
    Error,
}
impl Display for ConfigLogLevel {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "{}",
            match self {
                ConfigLogLevel::Trace => "trace",
                ConfigLogLevel::Debug => "debug",
                ConfigLogLevel::Info => "info",
                ConfigLogLevel::Warn => "warn",
                ConfigLogLevel::Error => "error",
            }
        )
    }
}

#[derive(Debug, Serialize, Deserialize, Default)]
#[serde(default)]
struct RedirectConfig {
    log_path: PathBuf,
    log_level: ConfigLogLevel,
    playback: ClientConfig,
    capture: ClientConfig,
    #[serde(skip)]
    source: ConfigSource,
}
impl RedirectConfig {
    fn load() -> Self {
        let source = if let Ok(str) = std::fs::read_to_string("redirect_config.toml") {
            if let Ok(cfg) = toml::from_str::<RedirectConfig>(&str) {
                return cfg;
            } else {
                ConfigSource::NoParse
            }
        } else {
            ConfigSource::NoFile
        };
        Self {
            log_path: PathBuf::new(),
            log_level: ConfigLogLevel::default(),
            playback: ClientConfig::default(),
            capture: ClientConfig::default(),
            source,
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
    target_buffer_dur_ms: u16,
    dur_modifier: u8,
    inverse: bool,
    compat: bool,
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
//     rclsid: *const windows_core::GUID,
//     punkouter: *mut core::ffi::c_void,
//     dwclscontext: CLSCTX,
//     riid: *const windows_core::GUID,
//     ppv: *mut *mut core::ffi::c_void,
// ) -> windows_core::HRESULT {
//     link!("ole32.dll" "system" fn CoCreateInstance(rclsid : *const windows_core::GUID, punkouter : * mut core::ffi::c_void, dwclscontext : CLSCTX, riid : *const windows_core::GUID, ppv : *mut *mut core::ffi::c_void) -> windows_core::HRESULT);
//     unsafe { CoCreateInstance(rclsid, punkouter.param().abi(), dwclscontext, riid, ppv) }
// }

// struct OpenALGlobal {
//     device: ALCdevice,
//     context: ALCcontext,
//     //mixer: Mutex<MixerState>,
// }
// static GLOBAL_AL: Mutex<Arc<OpenALGlobal>> = Mutex::new();

const LIB_NAME: PCSTR = s!("ole32.dll");
const CO_CREATE: PCSTR = s!("CoCreateInstance");
const CO_CREATE_EX: PCSTR = s!("CoCreateInstanceEx");

const KEYWORDS: &[&str] = &["[SK]", "[GAME]"];

#[allow(unused)]
static HOOK_CO_CREATE_INSTANCE: LazyLock<GenericDetour<FnCoCreateInstance>> =
    LazyLock::new(|| unsafe {
        let func = GetProcAddress(GetModuleHandleA(LIB_NAME).unwrap(), CO_CREATE).unwrap();
        let func: FnCoCreateInstance = std::mem::transmute(func);
        GenericDetour::new(func, hooked_cocreateinstance).unwrap()
    });

static HOOK_CO_CREATE_INSTANCE_EX: LazyLock<GenericDetour<FnCoCreateInstanceEx>> =
    LazyLock::new(|| unsafe {
        let func = GetProcAddress(
            GetModuleHandleA(LIB_NAME).unwrap_or(LoadLibraryA(LIB_NAME).unwrap()),
            CO_CREATE_EX,
        )
        .unwrap();
        let func: FnCoCreateInstanceEx = std::mem::transmute(func);
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
        if *riid == IMMDeviceEnumerator::IID {
            info!(
                "!!! Intercepted IMMDeviceEnumerator creation via CoCreateInstance, returning proxy !!!"
            );
            let mut inner_raw: *mut c_void = null_mut();
            let ret =
                HOOK_CO_CREATE_INSTANCE.call(rclsid, p_outer, dwcls_context, riid, &mut inner_raw);
            let inner_enumerator = IMMDeviceEnumerator::from_raw(inner_raw as _);
            let proxy_enumerator = RedirectDeviceEnumerator::new(inner_enumerator);
            let proxy_unknown: IMMDeviceEnumerator = proxy_enumerator.into();
            *ppv = proxy_unknown.into_raw() as _;
            ret
        } else {
            HOOK_CO_CREATE_INSTANCE.call(rclsid, p_outer, dwcls_context, riid, ppv)
        }
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
        if hr.is_ok() {
            for i in 0..dwcount {
                let p_qi = presults.add(i as usize);
                if (*p_qi).hr.is_ok()
                    && (*p_qi)
                        .pIID
                        .as_ref()
                        .is_some_and(|iid| *iid == IMMDeviceEnumerator::IID)
                    && (*p_qi).pItf.is_some()
                {
                    debug!("CoCreateInstanceEx CLSCTX: {:?}", dwclsctx);
                    if let Ok(thread_desc) = GetThreadDescription(GetCurrentThread())
                        && let Ok(name) = thread_desc.to_string()
                        && KEYWORDS.iter().any(|keyword| name.contains(keyword))
                    {
                        info!(
                            "Skipping SpecialK CoCreateInstanceEx calls, thread name: {}",
                            name
                        );
                        continue;
                    }
                    info!(
                        "!!! Intercepted IMMDeviceEnumerator via CoCreateInstanceEx, replacing with proxy !!!"
                    );
                    let inner_enumerator: IMMDeviceEnumerator =
                        (*p_qi).pItf.take().unwrap().cast().unwrap();

                    let proxy_enumerator = RedirectDeviceEnumerator::new(inner_enumerator);
                    let proxy_unknown: IMMDeviceEnumerator = proxy_enumerator.into();
                    _ = (*p_qi).pItf.insert(proxy_unknown.into());
                }
            }
        } else {
            error!("CoCreateInstanceEx call failed with HRESULT: {}", hr);
        }
        hr
    }
}

// struct OpenALState {
//     device: *mut ALCdevice,
//     context: *mut ALCcontext,
//     source: ALuint, // OpenAL 的“播放器”
// }

#[implement(IMMDeviceCollection)]
struct RedirectDeviceCollection {
    inner: IMMDeviceCollection,
}
impl RedirectDeviceCollection {
    fn new(inner: IMMDeviceCollection) -> Self {
        Self { inner }
    }
}

impl IMMDeviceCollection_Impl for RedirectDeviceCollection_Impl {
    fn GetCount(&self) -> windows_core::Result<u32> {
        trace!("RedirectDeviceCollection::GetCount() called");
        unsafe { self.inner.GetCount() }
    }

    fn Item(&self, ndevice: u32) -> windows_core::Result<IMMDevice> {
        debug!(
            "RedirectDeviceCollection::Item() -> wrapping, device {}",
            ndevice
        );
        Ok(RedirectDevice::new(unsafe { self.inner.Item(ndevice)? }).into())
    }
}

#[implement(IMMDevice, IMMEndpoint)]
#[derive(Clone)]
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
        iid: *const GUID,
        dwclsctx: CLSCTX,
        pactivationparams: *const PROPVARIANT,
        ppinterface: *mut *mut c_void,
    ) -> windows_core::Result<()> {
        unsafe {
            if matches!(
                *iid,
                IAudioClient::IID | IAudioClient2::IID | IAudioClient3::IID
            ) {
                debug!("RedirectDevice::Activate() -> wrapping, iid: {:?}", *iid);
                let inner: IAudioClient3 = self.inner.Activate::<IAudioClient3>(
                    dwclsctx,
                    (!pactivationparams.is_null()).then_some(pactivationparams),
                )?;
                let dataflow = self.inner.cast::<IMMEndpoint>()?.GetDataFlow()?.into();
                let target_cfg = CONFIG.get(dataflow);
                let proxy_unknown: IAudioClient3 = if !target_cfg.compat {
                    RedirectAudioClient::new(inner, dataflow).into()
                } else {
                    RedirectCompatAudioClient::new(inner, dataflow).into()
                };
                let ret = proxy_unknown.query(iid, ppinterface);
                if ret.is_ok() { Ok(()) } else { Err(ret.into()) }
            } else {
                debug!("RedirectDevice::Activate() called, iid: {:?}", *iid);
                if matches!(
                    *iid,
                    IDirectSound::IID | IDirectSound8::IID | IDirectSoundCapture::IID
                ) {
                    error!(
                        "Program is requesting DirectSound components, this means you should use other tools!"
                    )
                }
                let proxy_unknown = self.inner.Activate::<IUnknown>(
                    dwclsctx,
                    (!pactivationparams.is_null()).then_some(pactivationparams),
                )?;
                let ret = proxy_unknown.query(iid, ppinterface);
                if ret.is_ok() { Ok(()) } else { Err(ret.into()) }
            }
        }
    }

    fn OpenPropertyStore(&self, stgmaccess: STGM) -> windows_core::Result<IPropertyStore> {
        debug!("RedirectDevice::OpenPropertyStore() -> wrapping");
        Ok(RedirectPropertyStore::new(unsafe { self.inner.OpenPropertyStore(stgmaccess)? }).into())
    }

    fn GetId(&self) -> windows_core::Result<windows_core::PWSTR> {
        debug!("RedirectDevice::GetId() called");
        unsafe { self.inner.GetId() }
    }

    fn GetState(&self) -> windows_core::Result<DEVICE_STATE> {
        trace!("RedirectDevice::GetState() called");
        unsafe { self.inner.GetState() }
    }
}
impl IMMEndpoint_Impl for RedirectDevice_Impl {
    fn GetDataFlow(&self) -> windows_core::Result<EDataFlow> {
        trace!("RedirectDevice::GetDataFlow() called");
        unsafe { self.inner.cast::<IMMEndpoint>()?.GetDataFlow() }
    }
}

#[implement(IPropertyStore)]
pub struct RedirectPropertyStore {
    inner: IPropertyStore,
}

impl RedirectPropertyStore {
    pub fn new(inner: IPropertyStore) -> Self {
        Self { inner }
    }
}

#[allow(clippy::not_unsafe_ptr_arg_deref)]
impl IPropertyStore_Impl for RedirectPropertyStore_Impl {
    fn GetValue(&self, key: *const PROPERTYKEY) -> windows_core::Result<PROPVARIANT> {
        debug!(
            "RedirectPropertyStore::GetValue() called, key: {:?}",
            unsafe { key.as_ref() }
        );
        unsafe { self.inner.GetValue(key) }
    }

    fn GetCount(&self) -> windows_core::Result<u32> {
        trace!("RedirectPropertyStore::GetCount() called");
        unsafe { self.inner.GetCount() }
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> windows_core::Result<()> {
        trace!("RedirectPropertyStore::GetAt() called");
        unsafe { self.inner.GetAt(iprop, pkey) }
    }

    fn SetValue(
        &self,
        key: *const PROPERTYKEY,
        propvar: *const PROPVARIANT,
    ) -> windows_core::Result<()> {
        trace!("RedirectPropertyStore::SetValue() called");
        unsafe { self.inner.SetValue(key, propvar) }
    }

    fn Commit(&self) -> windows_core::Result<()> {
        trace!("RedirectPropertyStore::Commit() called");
        unsafe { self.inner.Commit() }
    }
}

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
    ) -> windows_core::Result<IMMDeviceCollection> {
        debug!(
            "RedirectDeviceEnumerator::EnumAudioEndpoints() -> wrapping, flow: {:?}",
            dataflow
        );
        Ok(RedirectDeviceCollection::new(unsafe {
            self.inner.EnumAudioEndpoints(dataflow, dwstatemask)?
        })
        .into())
    }

    fn GetDefaultAudioEndpoint(
        &self,
        dataflow: EDataFlow,
        role: ERole,
    ) -> windows_core::Result<IMMDevice> {
        debug!(
            "RedirectDeviceEnumerator::GetDefaultAudioEndpoint() -> wrapping, flow: {:?}",
            dataflow
        );
        Ok(
            RedirectDevice::new(unsafe { self.inner.GetDefaultAudioEndpoint(dataflow, role)? })
                .into(),
        )
    }

    fn GetDevice(&self, pwstrid: &PCWSTR) -> windows_core::Result<IMMDevice> {
        info!("RedirectDeviceEnumerator::GetDevice() -> wrapping");
        Ok(RedirectDevice::new(unsafe { self.inner.GetDevice(*pwstrid)? }).into())
    }

    fn RegisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> windows_core::Result<()> {
        trace!("RedirectDeviceEnumerator::RegisterEndpointNotificationCallback() called");
        unsafe {
            self.inner
                .RegisterEndpointNotificationCallback(pclient.as_ref())
        }
    }

    fn UnregisterEndpointNotificationCallback(
        &self,
        pclient: Ref<IMMNotificationClient>,
    ) -> windows_core::Result<()> {
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
            _ => unreachable!(),
        }
    }
}

const fn calculate_buffer(sample_rate: u32, fundamental: u32, target: u16) -> u32 {
    sample_rate * target as u32 / 10000 / fundamental * fundamental
}

fn calculate_period(sample_rate: u32, buffer_len: u32) -> i64 {
    (buffer_len as i64 * 10000000) / sample_rate as i64
}

#[implement(IAudioClient3)]
struct RedirectAudioClient {
    inner: IAudioClient3,
    current_buffer_len: RwLock<u32>,
    samplerate: RwLock<u32>,
    min_len: RwLock<u32>,
    dataflow: DeviceDataFlow,
    // oal_device: *mut ALCdevice,
    // oal_context: Option<ALCcontext>,
}

impl RedirectAudioClient {
    fn new(inner: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        // let oal_device = {
        //     let prop = unsafe { imm_device.GetId().unwrap() };
        //     let oal_name = CString::new(unsafe { prop.to_string() }.unwrap()).unwrap();
        //     unsafe { alcOpenDevice(oal_name.as_ptr() as _) }
        // };
        // Self {
        //     inner,
        //     device,
        //     oal_device,
        //     oal_context: None,
        // }
        Self {
            inner,
            current_buffer_len: 0.into(),
            samplerate: 0.into(),
            min_len: 0.into(),
            dataflow,
        }
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
    ) -> windows_core::Result<()> {
        info!(
            "RedirectAudioClient::Initialize() -> redirecting to Low Latency Shared, direction: {:?}",
            self.dataflow
        );
        // let mut attr = Vec::<ALCint>::new();
        // let format = unsafe { *pformat };
        // attr.push(ALC_FREQUENCY as i32);
        // attr.push(format.nSamplesPerSec as i32);
        // attr.push(ALC_FREQUENCY as i32);
        // attr.push(format.nSamplesPerSec as i32);
        // let oal_context = unsafe { alcCreateContext(self.oal_device, attrlist) };
        if sharemode == AUDCLNT_SHAREMODE_SHARED {
            info!("Original dur: {} * 100ns", hnsbufferduration);
            let target_cfg = CONFIG.get(self.dataflow);
            unsafe {
                let current_smp = (*pformat).nSamplesPerSec;
                if *self.samplerate.read().unwrap() != current_smp {
                    *self.samplerate.write().unwrap() = current_smp;
                    let mut pdefaultperiodinframes = 0;
                    let mut pfundamentalperiodinframes = 0;
                    let mut pminperiodinframes = 0;
                    let mut pmaxperiodinframes = 0;
                    self.inner.GetSharedModeEnginePeriod(
                        pformat,
                        &mut pdefaultperiodinframes,
                        &mut pfundamentalperiodinframes,
                        &mut pminperiodinframes,
                        &mut pmaxperiodinframes,
                    )?;

                    let calculated_len = if target_cfg.target_buffer_dur_ms != 0 {
                        calculate_buffer(
                            current_smp,
                            pfundamentalperiodinframes,
                            target_cfg.target_buffer_dur_ms,
                        )
                        .clamp(pminperiodinframes, pmaxperiodinframes)
                    } else {
                        pminperiodinframes
                    };

                    *self.current_buffer_len.write().unwrap() = calculated_len;
                    *self.min_len.write().unwrap() = pminperiodinframes;

                    info!(
                        "Current period = {}, Min period = {}",
                        calculated_len, pminperiodinframes
                    );
                    self.inner.InitializeSharedAudioStream(
                        streamflags,
                        calculated_len,
                        pformat,
                        (!audiosessionguid.is_null()).then_some(audiosessionguid),
                    )
                } else {
                    info!(
                        "Current period = {}, Min period = {}",
                        self.current_buffer_len.read().unwrap(),
                        self.min_len.read().unwrap()
                    );
                    self.inner.InitializeSharedAudioStream(
                        streamflags,
                        *self.current_buffer_len.read().unwrap(),
                        pformat,
                        (!audiosessionguid.is_null()).then_some(audiosessionguid),
                    )
                }
            }
        } else {
            unsafe {
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    hnsbufferduration,
                    hnsperiodicity,
                    pformat,
                    (!audiosessionguid.is_null()).then_some(audiosessionguid),
                )
            }
        }
    }

    fn GetBufferSize(&self) -> windows_core::Result<u32> {
        let buf = unsafe { self.inner.GetBufferSize()? };
        info!(
            "RedirectAudioClient::GetBufferSize() called, buffer length: {}",
            buf
        );
        Ok(buf)
    }

    fn GetStreamLatency(&self) -> windows_core::Result<i64> {
        info!("RedirectAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows_core::Result<u32> {
        trace!("RedirectAudioClient::GetCurrentPadding() called");
        unsafe { self.inner.GetCurrentPadding() }
    }

    fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT {
        debug!("RedirectAudioClient::IsFormatSupported() called");
        unsafe {
            self.inner.IsFormatSupported(
                sharemode,
                pformat,
                (!ppclosestmatch.is_null()).then_some(ppclosestmatch),
            )
        }
    }

    fn GetMixFormat(&self) -> windows_core::Result<*mut WAVEFORMATEX> {
        info!(
            "RedirectAudioClient::GetMixFormat() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.GetMixFormat() }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> windows_core::Result<()> {
        info!(
            "RedirectAudioClient::GetDevicePeriod() called, direction: {:?}",
            self.dataflow
        );
        let mut returned_default = 0;
        unsafe {
            self.inner.GetDevicePeriod(
                Some(&mut returned_default),
                (!phnsminimumdeviceperiod.is_null()).then_some(phnsminimumdeviceperiod),
            )?
        }
        let target_cfg = CONFIG.get(self.dataflow);
        if (*self.samplerate.read().unwrap()) == 0 {
            warn!("Called before initialize, inserting parameters");

            unsafe {
                let pformat = self.inner.GetMixFormat()?;
                *self.samplerate.write().unwrap() = (*pformat).nSamplesPerSec;

                let mut pdefaultperiodinframes = 0;
                let mut pfundamentalperiodinframes = 0;
                let mut pminperiodinframes = 0;
                let mut pmaxperiodinframes = 0;
                self.inner.GetSharedModeEnginePeriod(
                    pformat,
                    &mut pdefaultperiodinframes,
                    &mut pfundamentalperiodinframes,
                    &mut pminperiodinframes,
                    &mut pmaxperiodinframes,
                )?;

                let calculated_len = if target_cfg.target_buffer_dur_ms != 0 {
                    calculate_buffer(
                        (*pformat).nSamplesPerSec,
                        pfundamentalperiodinframes,
                        target_cfg.target_buffer_dur_ms,
                    )
                    .clamp(pminperiodinframes, pmaxperiodinframes)
                } else {
                    pminperiodinframes
                };
                *self.current_buffer_len.write().unwrap() = calculated_len;
                *self.min_len.write().unwrap() = pminperiodinframes;
            };
        }
        if let Some(ptr) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
            info!("original phnsdefaultdeviceperiod: {}", returned_default);
            let mut dur = calculate_period(
                *self.samplerate.read().unwrap(),
                *self.current_buffer_len.read().unwrap(),
            );
            if target_cfg.dur_modifier > 1 {
                if target_cfg.inverse {
                    dur *= target_cfg.dur_modifier as i64
                } else {
                    dur /= target_cfg.dur_modifier as i64
                }
            }
            *ptr = dur;
            info!("phnsdefaultdeviceperiod: {}", ptr);
        }
        if let Some(ptr) = unsafe { phnsminimumdeviceperiod.as_ref() } {
            info!("phnsminimumdeviceperiod: {}", ptr);
        }
        Ok(())
    }

    fn Start(&self) -> windows_core::Result<()> {
        info!(
            "RedirectAudioClient::Start() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows_core::Result<()> {
        info!(
            "RedirectAudioClient::Stop() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        info!(
            "RedirectAudioClient::Reset() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows_core::Result<()> {
        info!("RedirectAudioClient::SetEventHandle() called");
        unsafe { self.inner.SetEventHandle(eventhandle) }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows_core::Result<()> {
        let iid = unsafe { *riid };
        debug!(
            "RedirectAudioClient::GetService() called, iid: {iid:?}, direction: {:?}",
            self.dataflow
        );
        match iid {
            IAudioSessionControl::IID => {
                debug!("Returned IAudioSessionControl");
                unsafe { *ppv = self.inner.GetService::<IAudioSessionControl>()?.into_raw() as _ };
                Ok(())
            }
            IAudioRenderClient::IID => {
                debug!("Returned IAudioRenderClient");
                unsafe { *ppv = self.inner.GetService::<IAudioRenderClient>()?.into_raw() as _ };
                Ok(())
            }
            IAudioCaptureClient::IID => {
                debug!("Returned IAudioCaptureClient");
                unsafe { *ppv = self.inner.GetService::<IAudioCaptureClient>()?.into_raw() as _ };
                Ok(())
            }
            IAudioClientDuckingControl::IID => {
                debug!("Returned IAudioClientDuckingControl");
                unsafe {
                    *ppv = self
                        .inner
                        .GetService::<IAudioClientDuckingControl>()?
                        .into_raw() as _
                };
                Ok(())
            }
            IAudioClock::IID => {
                debug!("Returned IAudioClock");
                unsafe { *ppv = self.inner.GetService::<IAudioClock>()?.into_raw() as _ };
                Ok(())
            }
            IChannelAudioVolume::IID => {
                debug!("Returned IChannelAudioVolume");
                unsafe { *ppv = self.inner.GetService::<IChannelAudioVolume>()?.into_raw() as _ };
                Ok(())
            }
            ISimpleAudioVolume::IID => {
                debug!("Returned ISimpleAudioVolume");
                unsafe { *ppv = self.inner.GetService::<ISimpleAudioVolume>()?.into_raw() as _ };
                Ok(())
            }
            IAudioStreamVolume::IID => {
                debug!("Returned IAudioStreamVolume");
                unsafe { *ppv = self.inner.GetService::<IAudioStreamVolume>()?.into_raw() as _ };
                Ok(())
            }
            _ => {
                error!("Called unimplemented service!");
                Err(Error::from(E_NOINTERFACE))
            }
        }
    }
}

impl IAudioClient2_Impl for RedirectAudioClient_Impl {
    fn IsOffloadCapable(
        &self,
        category: AUDIO_STREAM_CATEGORY,
    ) -> windows_core::Result<windows_core::BOOL> {
        info!("RedirectAudioClient::IsOffloadCapable() called");
        unsafe { self.inner.IsOffloadCapable(category) }
    }

    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows_core::Result<()> {
        info!("RedirectAudioClient::SetClientProperties() called");
        unsafe { self.inner.SetClientProperties(pproperties) }
    }

    fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: windows_core::BOOL,
        phnsminbufferduration: *mut i64,
        phnsmaxbufferduration: *mut i64,
    ) -> windows_core::Result<()> {
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
    ) -> windows_core::Result<()> {
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
    ) -> windows_core::Result<()> {
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
    ) -> windows_core::Result<()> {
        info!(
            "RedirectAudioClient::InitializeSharedAudioStream() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.inner.InitializeSharedAudioStream(
                streamflags,
                periodinframes,
                pformat,
                (!audiosessionguid.is_null()).then_some(audiosessionguid),
            )
        }
    }
}
struct RawAudioBuffer {
    buffer: HeapRb<u8>,
    align: u16,
    dwflag: u32,
}
impl RawAudioBuffer {
    fn new(align: u16, buffer_len: u32) -> (Self, usize) {
        let len = buffer_len as usize * align as usize;
        (
            Self {
                buffer: HeapRb::new(len),
                align,
                dwflag: 0,
            },
            len,
        )
    }
    fn write(&mut self, data: &[u8]) {
        self.buffer.push_slice_overwrite(data)
    }
    fn read(&mut self, buf: &mut [u8]) -> usize {
        self.buffer.pop_slice(buf) / self.align as usize
    }
    fn available_frames(&self) -> usize {
        self.buffer.occupied_len() / self.align as usize
    }
}

#[derive(Debug)]
struct Communicator {
    stop_flag: AtomicBool,
    handle: Arc<CallbackHandle>,
    raw_buf_len: usize,
    align: u16,
    pad: Arc<AtomicUsize>,
}
impl Communicator {
    fn new(handle: Arc<CallbackHandle>, raw_buf_len: usize, align: u16) -> Self {
        Self {
            stop_flag: AtomicBool::new(false),
            handle,
            raw_buf_len,
            align,
            pad: AtomicUsize::new(0).into(),
        }
    }
}

#[derive(Debug)]
#[repr(transparent)]
struct CallbackHandle(HANDLE);
impl CallbackHandle {
    fn new() -> Result<Self> {
        unsafe {
            Ok(Self(CreateEventExA(
                None,
                None,
                CREATE_EVENT::default(),
                (EVENT_MODIFY_STATE | SYNCHRONIZATION_SYNCHRONIZE).0,
            )?))
        }
    }
    fn inner(&self) -> HANDLE {
        self.0
    }
}
unsafe impl Send for CallbackHandle {}
unsafe impl Sync for CallbackHandle {}
impl Drop for CallbackHandle {
    fn drop(&mut self) {
        unsafe { CloseHandle(self.0).unwrap() }
    }
}

struct CallbackThread {
    handle: Option<JoinHandle<()>>,
    inner: Arc<Communicator>,
    buffer: Arc<Mutex<RawAudioBuffer>>,
}
impl CallbackThread {
    fn new(
        handle: JoinHandle<()>,
        inner: Arc<Communicator>,
        buffer: Arc<Mutex<RawAudioBuffer>>,
    ) -> Self {
        Self {
            handle: Some(handle),
            inner,
            buffer,
        }
    }
}
impl Drop for CallbackThread {
    fn drop(&mut self) {
        let callback = self.inner.handle.inner();
        self.inner
            .stop_flag
            .store(true, std::sync::atomic::Ordering::Release);

        if let Some(handle) = self.handle.take() {
            unsafe {
                SetEvent(callback).ok();
            }
            let join_result = std::thread::spawn(move || handle.join()).join();
            match join_result {
                Ok(Ok(())) => {
                    info!("Consumer thread shut down cleanly.");
                }
                Ok(Err(e)) => {
                    error!("Consumer thread panicked on shutdown: {:?}", e);
                }
                Err(e) => {
                    error!("Failed to join the joiner thread: {:?}", e);
                }
            }
        }
    }
}

#[repr(transparent)]
struct RenderClientWrapper(IAudioRenderClient);
impl RenderClientWrapper {
    fn into_inner(self) -> IAudioRenderClient {
        self.0
    }
}

unsafe impl Send for RenderClientWrapper {}

#[implement(IAudioClient3)]
struct RedirectCompatAudioClient {
    inner: IAudioClient3,
    current_buffer_len: UnsafeCell<u32>,
    samplerate: UnsafeCell<u32>,
    min_len: UnsafeCell<u32>,
    dataflow: DeviceDataFlow,
    thread: RwLock<Option<CallbackThread>>,
}

impl RedirectCompatAudioClient {
    fn new(inner: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        Self {
            inner,
            current_buffer_len: 0.into(),
            samplerate: 0.into(),
            min_len: 0.into(),
            dataflow,
            thread: None.into(),
        }
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
    ) -> windows_core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Initialize() -> redirecting to Low Latency Shared with event call, direction: {:?}",
            self.dataflow
        );
        if sharemode == AUDCLNT_SHAREMODE_SHARED {
            info!("Original dur: {} * 100ns", hnsbufferduration);
            let target_cfg = CONFIG.get(self.dataflow);
            unsafe {
                let inner_smp = UnsafeCell::raw_get(&self.samplerate);
                if *inner_smp != (*pformat).nSamplesPerSec {
                    *inner_smp = (*pformat).nSamplesPerSec;
                    let mut pdefaultperiodinframes = 0;
                    let mut pfundamentalperiodinframes = 0;
                    let mut pminperiodinframes = 0;
                    let mut pmaxperiodinframes = 0;
                    self.inner.GetSharedModeEnginePeriod(
                        pformat,
                        &mut pdefaultperiodinframes,
                        &mut pfundamentalperiodinframes,
                        &mut pminperiodinframes,
                        &mut pmaxperiodinframes,
                    )?;

                    let calculated_len = if target_cfg.target_buffer_dur_ms != 0 {
                        calculate_buffer(
                            *UnsafeCell::get(&self.samplerate),
                            pfundamentalperiodinframes,
                            target_cfg.target_buffer_dur_ms,
                        )
                        .clamp(pminperiodinframes, pmaxperiodinframes)
                    } else {
                        pminperiodinframes
                    };

                    *UnsafeCell::raw_get(&self.current_buffer_len) = calculated_len;
                    *UnsafeCell::raw_get(&self.min_len) = pminperiodinframes;

                    info!(
                        "Current period = {}, Min period = {}",
                        calculated_len, pminperiodinframes
                    );
                    self.inner.InitializeSharedAudioStream(
                        streamflags | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                        *self.current_buffer_len.get(),
                        pformat,
                        (!audiosessionguid.is_null()).then_some(audiosessionguid),
                    )?;
                    let callback: Arc<CallbackHandle> = CallbackHandle::new()?.into();
                    let inthread_callback = callback.clone();
                    let (buffer, len) =
                        RawAudioBuffer::new((*pformat).nBlockAlign, self.inner.GetBufferSize()?);
                    let communicator: Arc<Communicator> =
                        Communicator::new(callback.clone(), len, (*pformat).nBlockAlign).into();
                    let buffer = Arc::new(Mutex::new(buffer));
                    let inthread_buffer = buffer.clone();
                    let inthread_communicator = communicator.clone();
                    let render =
                        RenderClientWrapper(self.inner.GetService::<IAudioRenderClient>()?);
                    let consumer = spawn(move || {
                        consumer(
                            inthread_callback,
                            inthread_communicator,
                            render,
                            inthread_buffer,
                        )
                    });
                    _ = self.thread.write().unwrap().insert(CallbackThread::new(
                        consumer,
                        communicator,
                        buffer,
                    ));
                    self.inner.SetEventHandle(
                        (*self.thread.read().unwrap())
                            .as_ref()
                            .unwrap()
                            .inner
                            .handle
                            .inner(),
                    )?;
                    Ok(())
                } else {
                    info!(
                        "Current period = {}, Min period = {}",
                        *self.current_buffer_len.get(),
                        *self.min_len.get()
                    );

                    self.inner.InitializeSharedAudioStream(
                        streamflags | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                        *self.current_buffer_len.get(),
                        pformat,
                        (!audiosessionguid.is_null()).then_some(audiosessionguid),
                    )?;
                    let callback: Arc<CallbackHandle> = CallbackHandle::new()?.into();
                    let inthread_callback = callback.clone();
                    let (buffer, len) =
                        RawAudioBuffer::new((*pformat).nBlockAlign, self.inner.GetBufferSize()?);
                    let communicator: Arc<Communicator> =
                        Communicator::new(callback.clone(), len, (*pformat).nBlockAlign).into();
                    let buffer = Arc::new(Mutex::new(buffer));
                    let inthread_buffer = buffer.clone();
                    let inthread_communicator = communicator.clone();
                    let render =
                        RenderClientWrapper(self.inner.GetService::<IAudioRenderClient>()?);
                    let consumer = spawn(move || {
                        consumer(
                            inthread_callback,
                            inthread_communicator,
                            render,
                            inthread_buffer,
                        )
                    });
                    _ = self.thread.write().unwrap().insert(CallbackThread::new(
                        consumer,
                        communicator,
                        buffer,
                    ));
                    self.inner.SetEventHandle(
                        (*self.thread.read().unwrap())
                            .as_ref()
                            .unwrap()
                            .inner
                            .handle
                            .inner(),
                    )?;
                    Ok(())
                }
            }
        } else {
            unsafe {
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    hnsbufferduration,
                    hnsperiodicity,
                    pformat,
                    (!audiosessionguid.is_null()).then_some(audiosessionguid),
                )
            }
        }
    }

    fn GetBufferSize(&self) -> windows_core::Result<u32> {
        let buf = unsafe { self.inner.GetBufferSize()? };
        info!(
            "RedirectCompatAudioClient::GetBufferSize() called, buffer length: {}",
            buf
        );
        Ok(buf)
    }

    fn GetStreamLatency(&self) -> windows_core::Result<i64> {
        info!("RedirectCompatAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows_core::Result<u32> {
        let pad = self
            .thread
            .read()
            .unwrap()
            .as_ref()
            .unwrap()
            .inner
            .pad
            .load(std::sync::atomic::Ordering::Acquire);
        trace!(
            "RedirectCompatAudioClient::GetCurrentPadding() called, pad: {}",
            pad
        );
        Ok(pad as u32)
    }

    fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: *mut *mut WAVEFORMATEX,
    ) -> HRESULT {
        debug!("RedirectCompatAudioClient::IsFormatSupported() called");
        unsafe {
            self.inner.IsFormatSupported(
                sharemode,
                pformat,
                (!ppclosestmatch.is_null()).then_some(ppclosestmatch),
            )
        }
    }

    fn GetMixFormat(&self) -> windows_core::Result<*mut WAVEFORMATEX> {
        info!(
            "RedirectCompatAudioClient::GetMixFormat() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.GetMixFormat() }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> windows_core::Result<()> {
        info!(
            "RedirectCompatAudioClient::GetDevicePeriod() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.inner.GetDevicePeriod(
                (!phnsdefaultdeviceperiod.is_null()).then_some(phnsdefaultdeviceperiod),
                (!phnsminimumdeviceperiod.is_null()).then_some(phnsminimumdeviceperiod),
            )
        }
    }

    fn Start(&self) -> windows_core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Start() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows_core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Stop() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Reset() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows_core::Result<()> {
        info!("RedirectCompatAudioClient::SetEventHandle() called");
        unsafe { self.inner.SetEventHandle(eventhandle) }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows_core::Result<()> {
        let iid = unsafe { *riid };
        debug!(
            "RedirectCompatAudioClient::GetService() called, iid: {iid:?}, direction: {:?}",
            self.dataflow
        );
        match iid {
            IAudioSessionControl::IID => {
                debug!("Returned IAudioSessionControl");
                unsafe { *ppv = self.inner.GetService::<IAudioSessionControl>()?.into_raw() as _ };
                Ok(())
            }
            IAudioRenderClient::IID => {
                debug!("Returned DummyAudioRenderClient");
                let (buf, len, align, pad) = {
                    let guard = self.thread.read().unwrap();
                    let thread = guard.as_ref();
                    let thread = thread.unwrap();
                    (
                        thread.buffer.clone(),
                        thread.inner.raw_buf_len,
                        thread.inner.align as usize,
                        thread.inner.pad.clone(),
                    )
                };
                let dummy = DummyAudioRenderClient::new(buf, len, align, pad);
                let renderclient: IAudioRenderClient = dummy.into();
                let ret = unsafe { renderclient.query(&IAudioRenderClient::IID, ppv) };
                if ret.is_ok() { Ok(()) } else { Err(ret.into()) }
            }
            IAudioCaptureClient::IID => {
                debug!("Returned IAudioCaptureClient");
                unsafe { *ppv = self.inner.GetService::<IAudioCaptureClient>()?.into_raw() as _ };
                Ok(())
            }
            IAudioClientDuckingControl::IID => {
                debug!("Returned IAudioClientDuckingControl");
                unsafe {
                    *ppv = self
                        .inner
                        .GetService::<IAudioClientDuckingControl>()?
                        .into_raw() as _
                };
                Ok(())
            }
            IAudioClock::IID => {
                debug!("Returned IAudioClock");
                unsafe { *ppv = self.inner.GetService::<IAudioClock>()?.into_raw() as _ };
                Ok(())
            }
            IChannelAudioVolume::IID => {
                debug!("Returned IChannelAudioVolume");
                unsafe { *ppv = self.inner.GetService::<IChannelAudioVolume>()?.into_raw() as _ };
                Ok(())
            }
            ISimpleAudioVolume::IID => {
                debug!("Returned ISimpleAudioVolume");
                unsafe { *ppv = self.inner.GetService::<ISimpleAudioVolume>()?.into_raw() as _ };
                Ok(())
            }
            IAudioStreamVolume::IID => {
                debug!("Returned IAudioStreamVolume");
                unsafe { *ppv = self.inner.GetService::<IAudioStreamVolume>()?.into_raw() as _ };
                Ok(())
            }
            _ => {
                error!("Called unimplemented service!");
                Err(Error::from(E_NOINTERFACE))
            }
        }
    }
}

fn consumer(
    handle: Arc<CallbackHandle>,
    data: Arc<Communicator>,
    client: RenderClientWrapper,
    buffer: Arc<Mutex<RawAudioBuffer>>,
) {
    let client = client.into_inner();
    loop {
        unsafe {
            WaitForSingleObject(handle.inner(), INFINITE);
        }
        if data.stop_flag.load(std::sync::atomic::Ordering::Acquire) {
            break;
        }
        let mut guard = buffer.lock().unwrap();
        let read_len = guard.available_frames();
        if read_len == 0 {
            continue;
        }
        let slice_ptr = unsafe { client.GetBuffer(read_len as u32).unwrap() };
        info!("callback!");
        let slice_len = read_len * guard.align as usize;
        let slice = unsafe { std::slice::from_raw_parts_mut(slice_ptr, slice_len) };
        let write_len = dbg!(guard.read(slice));
        data.pad.store(
            guard.available_frames(),
            std::sync::atomic::Ordering::Release,
        );
        unsafe {
            client
                .ReleaseBuffer(write_len as u32, guard.dwflag)
                .unwrap()
        };
    }
}

impl IAudioClient2_Impl for RedirectCompatAudioClient_Impl {
    fn IsOffloadCapable(
        &self,
        category: AUDIO_STREAM_CATEGORY,
    ) -> windows_core::Result<windows_core::BOOL> {
        info!("RedirectCompatAudioClient::IsOffloadCapable() called");
        unsafe { self.inner.IsOffloadCapable(category) }
    }

    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows_core::Result<()> {
        info!("RedirectCompatAudioClient::SetClientProperties() called");
        unsafe { self.inner.SetClientProperties(pproperties) }
    }

    fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: windows_core::BOOL,
        phnsminbufferduration: *mut i64,
        phnsmaxbufferduration: *mut i64,
    ) -> windows_core::Result<()> {
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
    ) -> windows_core::Result<()> {
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
    ) -> windows_core::Result<()> {
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
    ) -> windows_core::Result<()> {
        info!(
            "RedirectCompatAudioClient::InitializeSharedAudioStream() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.inner.InitializeSharedAudioStream(
                streamflags,
                periodinframes,
                pformat,
                (!audiosessionguid.is_null()).then_some(audiosessionguid),
            )
        }
    }
}

#[implement(IAudioRenderClient)]
struct DummyAudioRenderClient {
    buf: Arc<Mutex<RawAudioBuffer>>,
    align: usize,
    pad: Arc<AtomicUsize>,
    temp_buffer: UnsafeCell<Vec<u8>>,
}
impl DummyAudioRenderClient {
    fn new(
        buf: Arc<Mutex<RawAudioBuffer>>,
        len: usize,
        align: usize,
        pad: Arc<AtomicUsize>,
    ) -> Self {
        Self {
            buf,
            align,
            pad,
            temp_buffer: vec![0u8; len].into(),
        }
    }
}
impl IAudioRenderClient_Impl for DummyAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows_core::Result<*mut u8> {
        // 希望脏写不会影响
        let vec_ptr = self.temp_buffer.get();
        let buffer_ref: &mut Vec<u8> = unsafe { &mut *vec_ptr };
        info!("GetBuffer called, write buffer: {}", numframesrequested);
        Ok(buffer_ref.as_mut_ptr())
    }

    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows_core::Result<()> {
        let buffer_ptr = self.temp_buffer.get();
        let slice = unsafe { &(&(*buffer_ptr))[0..(numframeswritten as usize * self.align)] };
        match self.buf.lock() {
            Ok(mut buf) => {
                info!("ReleaseBuffer called!");
                buf.write(slice);
                // we'll pass everything happened to driver as a combined mess
                buf.dwflag |= dwflags;
                self.pad
                    .store(buf.available_frames(), std::sync::atomic::Ordering::Release);
                Ok(())
            }
            Err(e) => {
                error!("posioned?!?!?");
                panic!()
            }
        }
        //let mut buf = self.buf.lock().unwrap();
    }
}

#[unsafe(export_name = "proxy")]
unsafe extern "system" fn proxy_dummy() {}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(_hinst: HANDLE, reason: u32, _reserved: *mut c_void) -> BOOL {
    match reason {
        DLL_PROCESS_ATTACH => {
            unsafe {
                // HOOK_CO_CREATE_INSTANCE.enable().unwrap();
                HOOK_CO_CREATE_INSTANCE_EX.enable().unwrap();
            };
            std::thread::spawn(|| {
                // let _logger = Logger::try_with_env_or_str("info")
                //     .unwrap()
                //     .log_to_stdout()
                //     .start();

                let _logger = Logger::try_with_str(CONFIG.log_level.to_string())
                    .unwrap()
                    .log_to_file({
                        let spec = FileSpec::default()
                            .basename("redirect")
                            .suffix("log")
                            .suppress_timestamp();
                        if CONFIG.log_path.is_dir() {
                            spec.directory(CONFIG.log_path.clone())
                        } else {
                            spec
                        }
                    })
                    .duplicate_to_stdout(Duplicate::All)
                    .start();
                info!(
                    "Attempting to load config from working directory: {:?}",
                    std::env::current_dir()
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
            // HOOK_CO_CREATE_INSTANCE.disable().unwrap();
            HOOK_CO_CREATE_INSTANCE_EX.disable().unwrap();
        },
        _ => {}
    };
    BOOL::from(true)
}
