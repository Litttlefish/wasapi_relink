#[cfg(test)]
mod config_test;

#[allow(unused_imports)]
#[allow(clippy::single_component_path_imports)]
use auto_allocator;
// use openal_binds::*;
use flexi_logger::*;
use log::*;
use retour::GenericDetour;
use serde::*;
use std::cell::UnsafeCell;
use std::hint::unreachable_unchecked;
use std::ops::Add;
use std::os::raw::c_void;
use std::path::PathBuf;
use std::ptr::null_mut;
use std::slice::from_raw_parts_mut;
use std::sync::atomic::{AtomicU8, Ordering};
use std::sync::{Arc, LazyLock};

use windows::{
    Win32::{
        Foundation::*,
        Media::Audio::{DirectSound::*, Endpoints::*, *},
        Media::DirectShow::*,
        System::Com::{StructuredStorage::*, *},
        System::LibraryLoader::{GetModuleHandleW, GetProcAddress, LoadLibraryW},
        System::SystemServices::{DLL_PROCESS_ATTACH, DLL_PROCESS_DETACH},
        System::Threading::{GetCurrentThread, GetThreadDescription},
        UI::Shell::PropertiesSystem::{IPropertyStore, IPropertyStore_Impl},
    },
    core::*,
};

// static OLE32_LIB: std::sync::LazyLock<Library> = std::sync::LazyLock::new(|| unsafe {
//     Library::new("ole32.dll").expect("Failed to load original ole32.dll")
// });

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
    target_buffer_dur_ms: u16,
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
//     rclsid: *const windows::core::GUID,
//     punkouter: *mut core::ffi::c_void,
//     dwclscontext: CLSCTX,
//     riid: *const windows::core::GUID,
//     ppv: *mut *mut core::ffi::c_void,
// ) -> windows::core::HRESULT {
//     link!("ole32.dll" "system" fn CoCreateInstance(rclsid : *const windows::core::GUID, punkouter : * mut core::ffi::c_void, dwclscontext : CLSCTX, riid : *const windows::core::GUID, ppv : *mut *mut core::ffi::c_void) -> windows::core::HRESULT);
//     unsafe { CoCreateInstance(rclsid, punkouter.param().abi(), dwclscontext, riid, ppv) }
// }

// struct OpenALGlobal {
//     device: ALCdevice,
//     context: ALCcontext,
//     //mixer: Mutex<MixerState>,
// }
// static GLOBAL_AL: Mutex<Arc<OpenALGlobal>> = Mutex::new();

const LIB_NAME: PCWSTR = w!("ole32.dll");
const CO_CREATE: PCSTR = s!("CoCreateInstance");
const CO_CREATE_EX: PCSTR = s!("CoCreateInstanceEx");

const KEYWORDS: &[&str] = &["[GAME]", "[SK]"];

#[allow(unused)]
static HOOK_CO_CREATE_INSTANCE: LazyLock<GenericDetour<FnCoCreateInstance>> =
    LazyLock::new(|| unsafe {
        let func = GetProcAddress(
            GetModuleHandleW(LIB_NAME).unwrap_or_else(|_| LoadLibraryW(LIB_NAME).unwrap()),
            CO_CREATE,
        )
        .unwrap();
        let func: FnCoCreateInstance = std::mem::transmute(func);
        GenericDetour::new(func, hooked_cocreateinstance).unwrap()
    });

static HOOK_CO_CREATE_INSTANCE_EX: LazyLock<GenericDetour<FnCoCreateInstanceEx>> =
    LazyLock::new(|| unsafe {
        let func = GetProcAddress(
            GetModuleHandleW(LIB_NAME).unwrap_or_else(|e| {
                warn!("Unable to find ole32.dll handle, err: {e}, loading");
                LoadLibraryW(LIB_NAME).unwrap()
            }),
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
            if ret.is_ok() {
                let proxy_enumerator: IMMDeviceEnumerator =
                    RedirectDeviceEnumerator::new(IMMDeviceEnumerator::from_raw(inner_raw)).into();
                *ppv = proxy_enumerator.into_raw();
            }
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
        if *clsid == MMDeviceEnumerator && hr.is_ok() {
            debug!("CoCreateInstanceEx CLSCTX: {:?}", dwclsctx);
            if let Ok(thread_desc) = GetThreadDescription(GetCurrentThread())
                && !thread_desc.is_empty()
                && let Ok(name) = thread_desc.to_string()
                && KEYWORDS.iter().any(|keyword| name.contains(keyword))
            {
                info!(
                    "Skipping SpecialK CoCreateInstanceEx calls, thread name: {}",
                    name
                )
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
            error!("CoCreateInstanceEx call failed with HRESULT: {}", hr)
        }
        hr
    }
}

// struct OpenALState {
//     device: *mut ALCdevice,
//     context: *mut ALCcontext,
//     source: ALuint, // OpenAL 的“播放器”
// }

#[repr(transparent)]
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
    fn GetCount(&self) -> windows::core::Result<u32> {
        trace!("RedirectDeviceCollection::GetCount() called");
        unsafe { self.inner.GetCount() }
    }

    fn Item(&self, ndevice: u32) -> windows::core::Result<IMMDevice> {
        debug!(
            "RedirectDeviceCollection::Item() -> wrapping, device {}",
            ndevice
        );
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
                    let proxy_unknown: IAudioClient3 = if !CONFIG.get(dataflow).compat {
                        RedirectAudioClient::new(inner, dataflow).into()
                    } else {
                        RedirectCompatAudioClient::new(
                            inner,
                            self.inner
                                .Activate::<IAudioClient3>(dwclsctx, Some(pactivationparams))?,
                            dataflow,
                        )
                        .into()
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
        Ok(RedirectPropertyStore::new(unsafe { self.inner.OpenPropertyStore(stgmaccess)? }).into())
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

impl RedirectPropertyStore {
    pub fn new(inner: IPropertyStore) -> Self {
        Self { inner }
    }
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
    ) -> windows::core::Result<IMMDevice> {
        debug!(
            "RedirectDeviceEnumerator::GetDefaultAudioEndpoint() -> wrapping, flow: {:?}",
            dataflow
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

const fn calculate_buffer(sample_rate: u32, fundamental: u32, target: u16) -> u32 {
    sample_rate * target as u32 / 10000 / fundamental * fundamental
}

// fn calculate_period(sample_rate: u32, buffer_len: u32) -> i64 {
//     (buffer_len as i64 * 10000000) / sample_rate as i64
// }

// #[derive(Default)]
// struct InnerInfo {
//     current_buffer_len: u32,
//     samplerate: u32,
//     min_len: u32,
// }

#[implement(IAudioClient3)]
struct RedirectAudioClient {
    inner: IAudioClient3,
    // inner_info: UnsafeCell<InnerInfo>,
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
            // inner_info: InnerInfo::default().into(),
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
    ) -> windows::core::Result<()> {
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
                // let info_ref = &mut *self.inner_info.get();
                // info_ref.samplerate = (*pformat).nSamplesPerSec;
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
                // info_ref.current_buffer_len = calculated_len;
                // info_ref.min_len = pminperiodinframes;
                info!(
                    "Current period = {}, Min period = {}",
                    calculated_len, pminperiodinframes
                );
                self.inner.InitializeSharedAudioStream(
                    streamflags,
                    calculated_len,
                    pformat,
                    Some(audiosessionguid),
                )
            }
        } else {
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
        info!(
            "RedirectAudioClient::GetBufferSize() called, buffer length: {}",
            buf
        );
        Ok(buf)
    }

    fn GetStreamLatency(&self) -> windows::core::Result<i64> {
        info!("RedirectAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        trace!("RedirectAudioClient::GetCurrentPadding() called");
        unsafe { self.inner.GetCurrentPadding() }
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
            self.dataflow
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
            self.dataflow
        );
        unsafe {
            self.inner
                .GetDevicePeriod(Some(phnsdefaultdeviceperiod), Some(phnsminimumdeviceperiod))
        }
        // let mut returned_default = 0;
        // unsafe {
        //     self.inner.GetDevicePeriod(
        //         Some(&mut returned_default),
        //         (!phnsminimumdeviceperiod.is_null()).then_some(phnsminimumdeviceperiod),
        //     )?
        // }
        // let target_cfg = CONFIG.get(self.dataflow);
        // if (unsafe { *self.samplerate.get() }) == 0 {
        //     warn!("Called before initialize, inserting parameters");

        //     unsafe {
        //         let pformat = self.inner.GetMixFormat()?;
        //         *UnsafeCell::raw_get(&self.samplerate) = (*pformat).nSamplesPerSec;

        //         let mut pdefaultperiodinframes = 0;
        //         let mut pfundamentalperiodinframes = 0;
        //         let mut pminperiodinframes = 0;
        //         let mut pmaxperiodinframes = 0;
        //         self.inner.GetSharedModeEnginePeriod(
        //             pformat,
        //             &mut pdefaultperiodinframes,
        //             &mut pfundamentalperiodinframes,
        //             &mut pminperiodinframes,
        //             &mut pmaxperiodinframes,
        //         )?;

        //         let calculated_len = if target_cfg.target_buffer_dur_ms != 0 {
        //             calculate_buffer(
        //                 *UnsafeCell::get(&self.samplerate),
        //                 pfundamentalperiodinframes,
        //                 target_cfg.target_buffer_dur_ms,
        //             )
        //             .clamp(pminperiodinframes, pmaxperiodinframes)
        //         } else {
        //             pminperiodinframes
        //         };
        //         *UnsafeCell::raw_get(&self.current_buffer_len) = calculated_len;
        //         *UnsafeCell::raw_get(&self.min_len) = pminperiodinframes;
        //     };
        // }
        // if let Some(ptr) = unsafe { phnsdefaultdeviceperiod.as_mut() } {
        //     info!("original phnsdefaultdeviceperiod: {}", returned_default);
        //     let dur =
        //         unsafe { calculate_period(*self.samplerate.get(), *self.current_buffer_len.get()) };
        //     // if target_cfg.dur_modifier > 1 {
        //     //     if target_cfg.inverse {
        //     //         dur *= target_cfg.dur_modifier as i64
        //     //     } else {
        //     //         dur /= target_cfg.dur_modifier as i64
        //     //     }
        //     // }
        //     *ptr = dur;
        //     info!("phnsdefaultdeviceperiod: {}", ptr);
        // }
        // if let Some(ptr) = unsafe { phnsminimumdeviceperiod.as_ref() } {
        //     info!("phnsminimumdeviceperiod: {}", ptr);
        // }
        // Ok(())
    }

    fn Start(&self) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Start() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Stop() called, direction: {:?}",
            self.dataflow
        );
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info!(
            "RedirectAudioClient::Reset() called, direction: {:?}",
            self.dataflow
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
            self.dataflow
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
        info!(
            "RedirectAudioClient::InitializeSharedAudioStream() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.inner.InitializeSharedAudioStream(
                streamflags,
                periodinframes,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}

#[derive(Default)]
struct HookerInfo {
    hooker_padding: u32,
    align: u16,
    inner_buffer_len: u32,
    hooker_buffer_len: u32,
}
impl HookerInfo {
    #[inline]
    fn init(&mut self, inner: u32, hooker: u32, align: u16) {
        self.hooker_buffer_len = hooker;
        self.inner_buffer_len = inner;
        self.align = align;
        self.hooker_padding = inner - hooker;
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
    dataflow: DeviceDataFlow,
    trick: Arc<AtomicU8>,
    hooker_info: UnsafeCell<HookerInfo>,
}

impl RedirectCompatAudioClient {
    fn new(inner: IAudioClient3, hooker: IAudioClient3, dataflow: DeviceDataFlow) -> Self {
        Self {
            inner,
            hooker,
            dataflow,
            trick: AtomicU8::default().into(),
            hooker_info: HookerInfo::default().into(),
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
    ) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Initialize() -> redirecting to hooker Shared with small buffer, direction: {:?}",
            self.dataflow
        );
        if sharemode == AUDCLNT_SHAREMODE_SHARED {
            info!("Original dur: {} * 100ns", hnsbufferduration);
            let target_cfg = CONFIG.get(self.dataflow);
            unsafe {
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
                info!(
                    "Hooker period = {}, Min period = {}",
                    calculated_len, pminperiodinframes
                );
                self.hooker.InitializeSharedAudioStream(
                    streamflags,
                    calculated_len,
                    pformat,
                    Some(audiosessionguid),
                )?;
                self.inner.Initialize(
                    sharemode,
                    streamflags,
                    0,
                    hnsperiodicity,
                    pformat,
                    Some(audiosessionguid),
                )?;
                (&mut *self.hooker_info.get()).init(
                    self.inner.GetBufferSize()?,
                    self.hooker.GetBufferSize()?,
                    (*pformat).nBlockAlign,
                );
                Ok(())
            }
        } else {
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
        info!(
            "RedirectCompatAudioClient::GetBufferSize() called, buffer length: {}",
            buf
        );
        Ok(buf)
    }

    fn GetStreamLatency(&self) -> windows::core::Result<i64> {
        info!("RedirectCompatAudioClient::GetStreamLatency() called");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows::core::Result<u32> {
        if self.trick.load(Ordering::Acquire) != TrickState::Transparent as u8 {
            info!("RedirectCompatAudioClient::GetCurrentPadding() called, tricking");
            let info_ref = unsafe { &*self.hooker_info.get() };
            Ok(info_ref
                .hooker_padding
                .add(unsafe { self.inner.GetCurrentPadding()? })
                .min(info_ref.inner_buffer_len))
        } else {
            trace!("RedirectCompatAudioClient::GetCurrentPadding() called");
            unsafe { self.inner.GetCurrentPadding() }
        }
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
            self.dataflow
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
            self.dataflow
        );
        unsafe {
            self.inner
                .GetDevicePeriod(Some(phnsdefaultdeviceperiod), Some(phnsminimumdeviceperiod))
        }
    }

    fn Start(&self) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Start() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.trick
                .store(TrickState::Transparent as u8, Ordering::Release);
            self.hooker.Start()?;
            self.inner.Start()
        }
    }

    fn Stop(&self) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Stop() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.hooker.Stop()?;
            self.inner.Stop()
        }
    }

    fn Reset(&self) -> windows::core::Result<()> {
        info!(
            "RedirectCompatAudioClient::Reset() called, direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.trick
                .store(TrickState::Tricking as u8, Ordering::Release);
            self.hooker.Reset()?;
            self.inner.Reset()
        }
    }

    fn SetEventHandle(&self, eventhandle: HANDLE) -> windows::core::Result<()> {
        info!("RedirectCompatAudioClient::SetEventHandle() called");
        unsafe {
            self.hooker.SetEventHandle(eventhandle)?;
            self.inner.SetEventHandle(eventhandle)
        }
    }

    fn GetService(&self, riid: *const GUID, ppv: *mut *mut c_void) -> windows::core::Result<()> {
        let iid = unsafe { *riid };
        debug!(
            "RedirectCompatAudioClient::GetService() called, iid: {iid:?}, direction: {:?}",
            self.dataflow
        );

        match iid {
            IAudioRenderClient::IID => {
                debug!("Returned RedirectAudioRenderClient");
                let info_ref = unsafe { &*self.hooker_info.get() };
                let redirected = RedirectAudioRenderClient::new(
                    self.trick.clone(),
                    unsafe { self.inner.GetService::<IAudioRenderClient>()? },
                    info_ref.align,
                    info_ref.inner_buffer_len,
                    info_ref.hooker_buffer_len,
                );
                let renderclient: IAudioRenderClient = redirected.into();
                unsafe { assign(ppv, renderclient) }
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
        warn!(
            "RedirectCompatAudioClient::InitializeSharedAudioStream() called, this shouldn't happen on compat mode! direction: {:?}",
            self.dataflow
        );
        unsafe {
            self.hooker.InitializeSharedAudioStream(
                streamflags,
                periodinframes,
                pformat,
                Some(audiosessionguid),
            )?;
            self.inner.InitializeSharedAudioStream(
                streamflags,
                periodinframes,
                pformat,
                Some(audiosessionguid),
            )
        }
    }
}

#[implement(IAudioRenderClient)]
struct RedirectAudioRenderClient {
    inner: IAudioRenderClient,
    trick_buffer: UnsafeCell<Vec<u8>>,
    trick: Arc<AtomicU8>,
    raw_hooker_len: usize,
    hooker_buffer_len: u32,
}
impl RedirectAudioRenderClient {
    fn new(
        trick: Arc<AtomicU8>,
        inner: IAudioRenderClient,
        align: u16,
        inner_buffer_len: u32,
        hooker_buffer_len: u32,
    ) -> Self {
        Self {
            inner,
            trick_buffer: vec![0; inner_buffer_len as usize * align as usize].into(),
            trick,
            raw_hooker_len: hooker_buffer_len as usize * align as usize,
            hooker_buffer_len,
        }
    }
}
impl IAudioRenderClient_Impl for RedirectAudioRenderClient_Impl {
    fn GetBuffer(&self, numframesrequested: u32) -> windows::core::Result<*mut u8> {
        if self.trick.load(Ordering::Acquire) != TrickState::Transparent as u8 {
            info!(
                "RedirectAudioRenderClient::GetBuffer() called, requested: {numframesrequested}, tricking"
            );
            Ok(unsafe { &mut *self.trick_buffer.get() }.as_mut_ptr())
        } else {
            unsafe { self.inner.GetBuffer(numframesrequested) }
        }
    }

    fn ReleaseBuffer(&self, numframeswritten: u32, dwflags: u32) -> windows::core::Result<()> {
        match unsafe { std::mem::transmute::<u8, TrickState>(self.trick.load(Ordering::Acquire)) } {
            TrickState::Tricking => {
                info!(
                    "RedirectAudioRenderClient::ReleaseBuffer() called, written: {numframeswritten}, tricking"
                );
                unsafe {
                    info!(
                        "filling silent buffer, {} samples filled",
                        self.hooker_buffer_len
                    );
                    self.trick
                        .store(TrickState::Filled as u8, Ordering::Release);
                    let slice_to_write = from_raw_parts_mut(
                        self.inner.GetBuffer(self.hooker_buffer_len)?,
                        self.raw_hooker_len,
                    );
                    let slice = &(&(*self.trick_buffer.get()))[0..self.raw_hooker_len];
                    slice_to_write.copy_from_slice(slice);
                    self.inner.ReleaseBuffer(self.hooker_buffer_len, dwflags)
                }
            }
            TrickState::Filled => {
                info!("already filled, discarding");
                Ok(())
            }
            TrickState::Transparent => unsafe {
                self.inner.ReleaseBuffer(numframeswritten, dwflags)
            },
        }
    }
}

#[unsafe(export_name = "proxy")]
extern "C" fn proxy_dummy() {}

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
                let logger = Logger::with(<ConfigLogLevel as Into<LevelFilter>>::into(
                    CONFIG.log_level,
                ));
                let _handle = if !CONFIG.only_log_stdout {
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
                .start();

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
            // HOOK_CO_CREATE_INSTANCE.disable().unwrap();
            HOOK_CO_CREATE_INSTANCE_EX.disable().unwrap();
        },
        _ => (),
    };
    true.into()
}
