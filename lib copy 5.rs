#[cfg(test)]
mod config_test;

#[allow(unused_imports)]
#[warn(clippy::single_component_path_imports)]
use auto_allocator;
// use openal_binds::*;
use flexi_logger::*;
use log::*;
use parking_lot::Mutex;
use parking_lot::MutexGuard;
use parking_lot::RwLock;
use parking_lot::RwLockReadGuard;
use parking_lot::RwLockWriteGuard;
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
use std::sync::atomic::AtomicBool;
use std::sync::atomic::AtomicU32;
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
    aux_buf_len: u32,
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
                if *self.samplerate.read() != current_smp {
                    *self.samplerate.write() = current_smp;
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

                    *self.current_buffer_len.write() = calculated_len;
                    *self.min_len.write() = pminperiodinframes;

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
                        self.current_buffer_len.read(),
                        self.min_len.read()
                    );
                    self.inner.InitializeSharedAudioStream(
                        streamflags,
                        *self.current_buffer_len.read(),
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
        if (*self.samplerate.read()) == 0 {
            warn!("Called before initialize, inserting parameters");

            unsafe {
                let pformat = self.inner.GetMixFormat()?;
                *self.samplerate.write() = (*pformat).nSamplesPerSec;

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
                *self.current_buffer_len.write() = calculated_len;
                *self.min_len.write() = pminperiodinframes;
            };
        }
        if let Some(ptr) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
            info!("original phnsdefaultdeviceperiod: {}", returned_default);
            let mut dur =
                calculate_period(*self.samplerate.read(), *self.current_buffer_len.read());
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
    align: usize,
    dwflag: u32,
}
impl RawAudioBuffer {
    fn new(align: u16, buffer_len: u32) -> Self {
        let len = buffer_len as usize * align as usize;
        Self {
            buffer: HeapRb::new(len * 2),
            align: align as usize,
            dwflag: 0,
        }
    }
    fn write(&mut self, data: &[u8]) {
        self.buffer.push_slice_overwrite(data)
    }
    fn read(&mut self, mut buf: &mut [u8], len: usize) -> usize {
        self.buffer
            .write_into(&mut buf, Some(len * self.align))
            .unwrap()
            .unwrap()
            / self.align
    }
    fn available_frames(&self) -> usize {
        self.buffer.occupied_len() / self.align
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

#[implement(IAudioClient3)]
struct RedirectCompatAudioClient {
    inner: IAudioClient3,
    current_buffer_len: UnsafeCell<u32>,
    samplerate: UnsafeCell<u32>,
    min_len: UnsafeCell<u32>,
    padding: Arc<AtomicU32>,
    buf_len: u32,
    init_params: UnsafeCell<Option<(CallbackHandle, u16 /* align */)>>,
    dataflow: DeviceDataFlow,
    trick: Arc<AtomicBool>,
}

impl RedirectCompatAudioClient {
    fn new(inner: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        let target_cfg = CONFIG.get(dataflow);
        Self {
            inner,
            current_buffer_len: 0.into(),
            samplerate: 0.into(),
            min_len: 0.into(),
            padding: AtomicU32::new(0).into(),
            buf_len: if target_cfg.aux_buf_len > 0 {
                target_cfg.aux_buf_len
            } else {
                1052
            },
            init_params: None.into(),
            dataflow,
            trick: AtomicBool::new(true).into(),
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
                    let callback: CallbackHandle = CallbackHandle::new()?;
                    self.inner.SetEventHandle(callback.inner())?;
                    *UnsafeCell::raw_get(&self.init_params) =
                        Some((callback, (*pformat).nBlockAlign));
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
                    let callback: CallbackHandle = CallbackHandle::new()?;
                    self.inner.SetEventHandle(callback.inner())?;
                    *UnsafeCell::raw_get(&self.init_params) =
                        Some((callback, (*pformat).nBlockAlign));
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
        info!(
            "RedirectCompatAudioClient::GetBufferSize() called, buffer length: {}",
            self.buf_len
        );
        Ok(self.buf_len)
    }

    fn GetStreamLatency(&self) -> windows_core::Result<i64> {
        info!("RedirectCompatAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows_core::Result<u32> {
        trace!("RedirectCompatAudioClient::GetCurrentPadding() called");
        // unsafe { Ok(self.inner.GetCurrentPadding()? - 422) }

        Ok(self.padding.load(std::sync::atomic::Ordering::Acquire))
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
        unsafe {
            self.trick
                .store(false, std::sync::atomic::Ordering::Release);
            self.inner.Start()
        }
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
                debug!("Returned RedirectAudioRenderClient");
                let (handle, align) = (unsafe { &mut *UnsafeCell::raw_get(&self.init_params) })
                    .take()
                    .unwrap();
                let redirected = RedirectAudioRenderClient::new(
                    align,
                    self.buf_len,
                    handle,
                    self.padding.clone(),
                    self.inner.downgrade()?,
                    self.trick.clone(),
                );
                let renderclient: IAudioRenderClient = redirected.into();
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

fn callback(
    handle: Arc<CallbackHandle>,
    padding: Arc<AtomicU32>,
    buffer: Arc<RwLock<RawAudioBuffer>>,
    client: Weak<IAudioClient3>,
    align: usize,
    stop_flag: Arc<AtomicBool>,
) {
    let client = client.upgrade().unwrap();
    let inner = unsafe { client.GetService::<IAudioRenderClient>().unwrap() };
    let real_len = unsafe { client.GetBufferSize().unwrap() } as usize;
    loop {
        unsafe {
            WaitForSingleObject(handle.inner(), INFINITE);
        }
        if stop_flag.load(std::sync::atomic::Ordering::Acquire) {
            break;
        }
        let mut guard = buffer.write();
        let read_len = guard.available_frames();
        let pad = unsafe { client.GetCurrentPadding().unwrap() } as usize;
        let write_len = read_len.min(real_len - pad);
        info!("callback! data in mid-buffer:{}", read_len);
        info!("callback! data should write:{}", real_len - pad);
        if write_len == 0 {
            warn!("mid-buffer has been emptied, underflow may happen!");
            continue;
        }
        let slice_ptr = unsafe { inner.GetBuffer(write_len as u32).unwrap() };
        let slice_len = read_len * align;
        let slice = unsafe { std::slice::from_raw_parts_mut(slice_ptr, slice_len) };
        let replaced = std::mem::take(&mut guard.dwflag);
        let write_len = guard.read(slice, write_len);
        let guard = RwLockWriteGuard::downgrade(guard);
        padding.store(
            guard.available_frames() as u32,
            std::sync::atomic::Ordering::Release,
        );
        unsafe { inner.ReleaseBuffer(write_len as u32, replaced).unwrap() };
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
struct RedirectAudioRenderClient {
    buffer: Arc<RwLock<RawAudioBuffer>>,
    align: usize,
    temp_buffer: UnsafeCell<Vec<u8>>,
    padding: Arc<AtomicU32>,
    stop_flag: Arc<AtomicBool>,
    thread: Option<JoinHandle<()>>,
    handle: Arc<CallbackHandle>,
    trick: Arc<AtomicBool>,
}
impl RedirectAudioRenderClient {
    fn new(
        align: u16,
        buffer_len: u32,
        handle: CallbackHandle,
        padding: Arc<AtomicU32>,
        client: Weak<IAudioClient3>,
        trick: Arc<AtomicBool>,
    ) -> Self {
        let buffer = Arc::new(RwLock::new(RawAudioBuffer::new(align, buffer_len)));
        let thread_buffer = buffer.clone();
        let handle: Arc<CallbackHandle> = handle.into();
        let thread_handle = handle.clone();
        let thread_padding = padding.clone();
        let stop_flag: Arc<AtomicBool> = AtomicBool::new(false).into();
        let thread_stop_flag = stop_flag.clone();

        let thread = spawn(move || {
            callback(
                thread_handle,
                thread_padding,
                thread_buffer,
                client,
                align as usize,
                thread_stop_flag,
            );
        });
        Self {
            buffer,
            align: align as usize,
            temp_buffer: vec![0u8; buffer_len as usize * align as usize].into(),
            padding,
            stop_flag,
            thread: Some(thread),
            handle,
            trick,
        }
    }
}
impl IAudioRenderClient_Impl for RedirectAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows_core::Result<*mut u8> {
        // 希望脏写不会影响
        let vec_ptr = self.temp_buffer.get();
        info!("GetBuffer called, requested: {numframesrequested}");
        Ok(unsafe { &mut *vec_ptr }.as_mut_ptr())
    }

    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows_core::Result<()> {
        if numframeswritten == 0 {
            warn!("no data written in this release call, overflow may happen!");
            return Ok(());
        }
        let slice = unsafe {
            &(&(*self.temp_buffer.get()))[0..(if !self
                .trick
                .load(std::sync::atomic::Ordering::Acquire)
            {
                numframeswritten as usize
            } else {
                280
            } * self.align)]
        };
        let mut buf = self.buffer.write();
        self.padding
            .fetch_add(numframeswritten, std::sync::atomic::Ordering::SeqCst);
        buf.dwflag |= dwflags;
        info!("ReleaseBuffer called, written: {numframeswritten}");
        buf.write(slice);
        // we'll pass everything happened to driver as a combined mess
        Ok(())
        //let mut buf = self.buf.lock().unwrap();
    }
}
impl Drop for RedirectAudioRenderClient {
    fn drop(&mut self) {
        let callback = self.handle.inner();
        self.stop_flag
            .store(true, std::sync::atomic::Ordering::Release);

        if let Some(handle) = self.thread.take() {
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
