mod config_test;

use libloading::{Library, Symbol};
use openal_binds::*;
use retour::GenericDetour;
use std::ffi::CString;
use std::os::raw::c_void;
use std::os::windows::ffi::OsStrExt;
use std::ptr::null_mut;
use std::sync::atomic::{AtomicBool, Ordering};
use std::{
    collections::HashMap,
    sync::{Arc, LazyLock, Mutex},
};
use windows::Win32::System::LibraryLoader::{GetModuleHandleW, GetProcAddress, LoadLibraryW};
use windows::Win32::System::Threading::{GetCurrentProcessId, GetCurrentThreadId};
use windows::Win32::System::Variant::VT_LPWSTR;
use windows::Win32::UI::WindowsAndMessaging::{
    DBT_DEVICEARRIVAL, DBT_DEVTYP_DEVICEINTERFACE, DEV_BROADCAST_DEVICEINTERFACE_W, HWND_BROADCAST,
    SendMessageA, WM_DEVICECHANGE,
};
use windows::{
    Win32::{
        Devices::FunctionDiscovery::PKEY_Device_FriendlyName,
        Foundation::*,
        Media::Audio::*,
        System::Com::StructuredStorage::*,
        System::Com::*,
        System::SystemServices::{
            DLL_PROCESS_ATTACH, DLL_PROCESS_DETACH, DLL_THREAD_ATTACH, DLL_THREAD_DETACH,
        },
        UI::Shell::PropertiesSystem::{IPropertyStore, IPropertyStore_Impl},
    },
    core::*,
};

static OLE32_LIB: std::sync::LazyLock<Library> = std::sync::LazyLock::new(|| unsafe {
    Library::new("ole32.dll").expect("Failed to load original ole32.dll")
});

type fn_CoCreateInstance = unsafe extern "system" fn(
    *const GUID,
    *mut core::ffi::c_void,
    CLSCTX,
    *const GUID,
    *mut *mut core::ffi::c_void,
) -> HRESULT;

type fn_CoCreateInstanceEx = unsafe extern "system" fn(
    *const GUID,
    *mut core::ffi::c_void,
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

static hook_CoCreateInstance: LazyLock<GenericDetour<fn_CoCreateInstance>> =
    LazyLock::new(|| unsafe {
        let func = {
            let module = w!("ole32.dll");
            let symbol = CString::new("CoCreateInstance").unwrap();
            let handle = GetModuleHandleW(module).unwrap();
            GetProcAddress(handle, PCSTR(symbol.as_ptr() as _)).unwrap()
        };
        let func: fn_CoCreateInstance = std::mem::transmute(func);
        GenericDetour::new(func, hooked_cocreateinstance).unwrap()
    });

static hook_CoCreateInstanceEx: LazyLock<GenericDetour<fn_CoCreateInstanceEx>> =
    LazyLock::new(|| unsafe {
        let func = {
            let module = w!("ole32.dll");
            let symbol = CString::new("CoCreateInstanceEx").unwrap();
            let handle = GetModuleHandleW(module).unwrap();
            GetProcAddress(handle, PCSTR(symbol.as_ptr() as _)).unwrap()
        };
        let func: fn_CoCreateInstanceEx = std::mem::transmute(func);
        GenericDetour::new(func, hooked_cocreateinstanceex).unwrap()
    });

unsafe extern "system" fn hooked_cocreateinstance(
    rclsid: *const GUID,
    p_outer: *mut core::ffi::c_void,
    dwcls_context: CLSCTX,
    riid: *const GUID,
    ppv: *mut *mut core::ffi::c_void,
) -> HRESULT {
    unsafe {
        let ret = if *riid == IMMDeviceEnumerator::IID {
            log::info!(
                "!!! Intercepted IMMDeviceEnumerator creation via CoCreateInstance, returning proxy !!!"
            );
            let mut inner_raw: *mut c_void = null_mut();
            let ret =
                hook_CoCreateInstance.call(rclsid, p_outer, dwcls_context, riid, &mut inner_raw);
            let inner_enumerator = IMMDeviceEnumerator::from_raw(inner_raw as _);
            if let Ok(_) = inner_enumerator.cast::<RedirectDeviceEnumerator>() {
                log::info!("recieved RedirectDeviceEnumerator, skipping");
                *ppv = inner_raw;
            } else {
                let proxy_enumerator = RedirectDeviceEnumerator::new(inner_enumerator);
                let proxy_unknown: IMMDeviceEnumerator = proxy_enumerator.into();
                *ppv = proxy_unknown.into_raw() as _;
            }
            ret
        } else {
            hook_CoCreateInstance.call(rclsid, p_outer, dwcls_context, riid, ppv)
        };
        ret
    }
}

unsafe extern "system" fn hooked_cocreateinstanceex(
    clsid: *const windows_core::GUID,
    punkouter: *mut core::ffi::c_void,
    dwclsctx: CLSCTX,
    pserverinfo: *const COSERVERINFO,
    dwcount: u32,
    presults: *mut MULTI_QI,
) -> HRESULT {
    unsafe {
        let hr = hook_CoCreateInstanceEx.call(
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
                    log::info!(
                        "!!! Intercepted IMMDeviceEnumerator via CoCreateInstanceEx, replacing with proxy !!!"
                    );
                    let inner_enumerator: IMMDeviceEnumerator = (*p_qi)
                        .pItf
                        .take()
                        .unwrap()
                        .cast::<IMMDeviceEnumerator>()
                        .unwrap();
                    if let Ok(_) = inner_enumerator.cast::<RedirectDeviceEnumerator>() {
                        log::info!("recieved RedirectDeviceEnumerator, skipping");
                        _ = (*p_qi).pItf.insert(inner_enumerator.into());
                    } else {
                        let proxy_enumerator = RedirectDeviceEnumerator::new(inner_enumerator);
                        let proxy_unknown: IMMDeviceEnumerator = proxy_enumerator.into();
                        _ = (*p_qi).pItf.insert(proxy_unknown.into());
                    }
                }
            }
        } else {
            log::error!("CoCreateInstanceEx call failed with HRESULT: {}", hr);
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
        unsafe { self.inner.GetCount() }
    }

    fn Item(&self, ndevice: u32) -> windows_core::Result<IMMDevice> {
        log::info!("[PROXY] RedirectDevice::Item called for device {}", ndevice);
        Ok(RedirectDevice::new(unsafe { self.inner.Item(ndevice)? }, Some(ndevice)).into())
    }
}

#[implement(IMMDevice)]
#[derive(Clone)]
struct RedirectDevice {
    inner: IMMDevice,
    device_enum: Option<u32>,
}

impl RedirectDevice {
    pub fn new(inner: IMMDevice, device_enum: Option<u32>) -> Self {
        Self { inner, device_enum }
    }
}

impl IMMDevice_Impl for RedirectDevice_Impl {
    fn Activate(
        &self,
        iid: *const GUID,
        dwclsctx: CLSCTX,
        pactivationparams: *const PROPVARIANT,
        ppinterface: *mut *mut core::ffi::c_void,
    ) -> windows_core::Result<()> {
        unsafe {
            let riid = iid;
            let iid = *iid;
            log::info!(
                "[PROXY] RedirectDevice::Activate called, caller: {:?}",
                self.device_enum
            );
            if matches!(
                iid,
                IAudioClient::IID | IAudioClient2::IID | IAudioClient3::IID
            ) {
                log::info!("[PROXY] RedirectDevice::Activate called for IID: {:?}", iid);
                let inner: IAudioClient3 = dbg!(self.inner.Activate::<IAudioClient3>(
                    dwclsctx,
                    (!pactivationparams.is_null()).then(|| pactivationparams),
                )?);
                let proxy_client = RedirectAudioClient::new(inner, self.device_enum);
                let proxy_unknown: IAudioClient3 = proxy_client.into();
                let ret = proxy_unknown.query(riid, ppinterface);
                if ret.is_ok() { Ok(()) } else { Err(ret.into()) }
            } else {
                let ret = self.inner.Activate::<IUnknown>(
                    dwclsctx,
                    (!pactivationparams.is_null()).then(|| pactivationparams),
                )?;
                let ret = ret.query(riid, ppinterface);
                if ret.is_ok() { Ok(()) } else { Err(ret.into()) }
            }
        }
    }

    fn OpenPropertyStore(&self, stgmaccess: STGM) -> windows_core::Result<IPropertyStore> {
        // let real_store = unsafe { self.inner.OpenPropertyStore(stgmaccess)? };
        log::info!("[PROXY] RedirectDevice::OpenPropertyStore called");
        // let proxy_store = RedirectPropertyStore::new(real_store);
        // Ok(proxy_store.into())
        unsafe { self.inner.OpenPropertyStore(stgmaccess) }
    }

    fn GetId(&self) -> windows_core::Result<windows_core::PWSTR> {
        log::info!("[PROXY] RedirectDevice::GetId called");
        log::info!(
            "[PROXY] RedirectDevice::GetId called, caller: {:?}",
            self.device_enum
        );
        unsafe { self.inner.GetId() }
    }

    fn GetState(&self) -> windows_core::Result<DEVICE_STATE> {
        log::info!("[PROXY] RedirectDevice::GetState called");
        let state = unsafe { self.inner.GetState()? };
        log::info!(
            "[PROXY] RedirectDevice::GetState called, state: {:?}",
            state
        );
        Ok(state)
    }
}

#[implement(IMMDeviceEnumerator)]
#[derive(Clone)]
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
        log::info!(
            "[PROXY] EnumAudioEndpoints called, creating redirected collection, flow: {:?}",
            dataflow
        );
        let inner = unsafe { self.inner.EnumAudioEndpoints(dataflow, dwstatemask) }?;
        Ok(RedirectDeviceCollection::new(inner).into())
    }

    fn GetDefaultAudioEndpoint(
        &self,
        dataflow: EDataFlow,
        role: ERole,
    ) -> windows_core::Result<IMMDevice> {
        log::info!("[PROXY] GetDefaultAudioEndpoint() -> wrapping");
        let device = unsafe { self.inner.GetDefaultAudioEndpoint(dataflow, role)? };
        let redirected = RedirectDevice::new(device, None);
        Ok(redirected.into())
    }

    fn GetDevice(&self, pwstrid: &windows_core::PCWSTR) -> windows_core::Result<IMMDevice> {
        log::info!("[PROXY] GetDevice() -> wrapping");
        let device = unsafe { self.inner.GetDevice(*pwstrid)? };
        let redirected = RedirectDevice::new(device, None);
        Ok(redirected.into())
    }

    fn RegisterEndpointNotificationCallback(
        &self,
        pclient: windows_core::Ref<IMMNotificationClient>,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] RegisterEndpointNotificationCallback called");

        unsafe {
            self.inner
                .RegisterEndpointNotificationCallback(pclient.as_ref())
        }
    }

    fn UnregisterEndpointNotificationCallback(
        &self,
        pclient: windows_core::Ref<IMMNotificationClient>,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] UnregisterEndpointNotificationCallback called");

        unsafe {
            self.inner
                .UnregisterEndpointNotificationCallback(pclient.as_ref())
        }
    }
}
unsafe impl Interface for RedirectDeviceEnumerator {
    type Vtable = IUnknown_Vtbl;
    const IID: GUID = GUID::from_u128(0xBB1F0F23_C073_F2D7_4139_CF4D17978270);
}

#[implement(IAudioClient3)]
struct RedirectAudioClient {
    inner: IAudioClient3,
    device: Option<u32>,
}

impl RedirectAudioClient {
    fn new(inner: IAudioClient3, device: Option<u32>) -> Self {
        Self { inner, device }
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
        audiosessionguid: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] Initialize called, device: {:?}", self.device);
        dbg!(unsafe {
            self.inner.Initialize(
                sharemode,
                streamflags,
                hnsbufferduration,
                hnsperiodicity,
                pformat,
                (!audiosessionguid.is_null()).then(|| audiosessionguid),
            )
        })
    }

    fn GetBufferSize(&self) -> windows_core::Result<u32> {
        log::info!("[PROXY] GetBufferSize() -> Forwarding to real client");
        unsafe { self.inner.GetBufferSize() }
    }

    fn GetStreamLatency(&self) -> windows_core::Result<i64> {
        log::info!("[PROXY] GetStreamLatency() -> Forwarding to real client");
        unsafe { self.inner.GetStreamLatency() }
    }

    fn GetCurrentPadding(&self) -> windows_core::Result<u32> {
        // log::info!("[PROXY] GetCurrentPadding() -> Forwarding to real client");
        unsafe { self.inner.GetCurrentPadding() }
    }

    fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT {
        // log::info!("[PROXY] IsFormatSupported() -> Forwarding to real client");
        unsafe {
            self.inner.IsFormatSupported(
                sharemode,
                pformat,
                (!ppclosestmatch.is_null()).then(|| ppclosestmatch),
            )
        }
    }

    fn GetMixFormat(&self) -> windows_core::Result<*mut WAVEFORMATEX> {
        log::info!("[PROXY] GetMixFormat() -> Forwarding to real client");
        unsafe { self.inner.GetMixFormat() }
    }

    fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: *mut i64,
        phnsminimumdeviceperiod: *mut i64,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] GetDevicePeriod() -> Forwarding to real client");
        unsafe {
            self.inner.GetDevicePeriod(
                (!phnsdefaultdeviceperiod.is_null()).then(|| phnsdefaultdeviceperiod),
                (!phnsminimumdeviceperiod.is_null()).then(|| phnsminimumdeviceperiod),
            )
        }
    }

    fn Start(&self) -> windows_core::Result<()> {
        log::info!("[PROXY] Start() -> Forwarding to real client");
        unsafe { self.inner.Start() }
    }

    fn Stop(&self) -> windows_core::Result<()> {
        log::info!("[PROXY] Stop() -> Forwarding to real client");
        unsafe { self.inner.Stop() }
    }

    fn Reset(&self) -> windows_core::Result<()> {
        log::info!("[PROXY] Reset() -> Forwarding to real client");
        unsafe { self.inner.Reset() }
    }

    fn SetEventHandle(
        &self,
        eventhandle: windows::Win32::Foundation::HANDLE,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] SetEventHandle() -> Forwarding to real client");
        unsafe { self.inner.SetEventHandle(eventhandle) }
    }

    fn GetService(
        &self,
        riid: *const windows_core::GUID,
        ppv: *mut *mut core::ffi::c_void,
    ) -> windows_core::Result<()> {
        let iid = unsafe { *riid };
        log::info!(
            "[PROXY] GetService() -> Forwarding to real client, iid: {:?}",
            iid
        );
        log::info!("[PROXY] GetService called, device: {:?}", self.device);
        match iid {
            IAudioSessionControl::IID => {
                log::info!("[PROXY] GetService(IAudioSessionControl)",);
                let control = unsafe { self.inner.GetService::<IAudioSessionControl>()? };
                unsafe { *ppv = control.into_raw() as _ };
                Ok(())
            }
            IAudioRenderClient::IID => {
                log::info!("[PROXY] GetService(IAudioRenderClient)",);
                let render = unsafe { self.inner.GetService::<IAudioRenderClient>()? };
                unsafe { *ppv = render.into_raw() as _ };
                Ok(())
            }
            IAudioCaptureClient::IID => {
                log::info!("[PROXY] GetService(IAudioCaptureClient)",);
                let capture = unsafe { self.inner.GetService::<IAudioCaptureClient>()? };
                unsafe { *ppv = capture.into_raw() as _ };
                Ok(())
            }
            IAudioClientDuckingControl::IID => {
                let capture = unsafe { self.inner.GetService::<IAudioClientDuckingControl>()? };
                unsafe { *ppv = capture.into_raw() as _ };
                Ok(())
            }
            IAudioClock::IID => {
                let capture = unsafe { self.inner.GetService::<IAudioClock>()? };
                unsafe { *ppv = capture.into_raw() as _ };
                Ok(())
            }
            IChannelAudioVolume::IID => {
                let capture = unsafe { self.inner.GetService::<IChannelAudioVolume>()? };
                unsafe { *ppv = capture.into_raw() as _ };
                Ok(())
            }
            ISimpleAudioVolume::IID => {
                let capture = unsafe { self.inner.GetService::<ISimpleAudioVolume>()? };
                unsafe { *ppv = capture.into_raw() as _ };
                Ok(())
            }
            IAudioStreamVolume::IID => {
                let capture = unsafe { self.inner.GetService::<IAudioStreamVolume>()? };
                unsafe { *ppv = capture.into_raw() as _ };
                Ok(())
            }
            _ => Err(Error::from(E_NOINTERFACE)),
        }
    }
}

impl IAudioClient2_Impl for RedirectAudioClient_Impl {
    fn IsOffloadCapable(
        &self,
        category: AUDIO_STREAM_CATEGORY,
    ) -> windows_core::Result<windows_core::BOOL> {
        log::info!("[PROXY] IsOffloadCapable() -> Forwarding to real client");
        unsafe { self.inner.IsOffloadCapable(category) }
    }

    fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] SetClientProperties() -> Forwarding to real client");
        unsafe { self.inner.SetClientProperties(pproperties) }
    }

    fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: windows_core::BOOL,
        phnsminbufferduration: *mut i64,
        phnsmaxbufferduration: *mut i64,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] GetBufferSizeLimits() -> Forwarding to real client");
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
        log::info!("[PROXY] GetSharedModeEnginePeriod() -> Forwarding to real client");
        log::info!(
            "[PROXY] GetSharedModeEnginePeriod called, device: {:?}",
            self.device
        );
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
        log::info!("[PROXY] GetCurrentSharedModeEnginePeriod() -> Forwarding to real client");
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
        audiosessionguid: *const windows_core::GUID,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] InitializeSharedAudioStream() -> Forwarding to real client");
        unsafe {
            self.inner.InitializeSharedAudioStream(
                streamflags,
                periodinframes,
                pformat,
                (!audiosessionguid.is_null()).then(|| audiosessionguid),
            )
        }
    }
}

fn get_current_ids() -> (u32, u32) {
    let pid = unsafe { GetCurrentProcessId() };
    let tid = unsafe { GetCurrentThreadId() };
    (pid, tid)
}

#[unsafe(no_mangle)]
unsafe extern "system" fn DllMain(_hinst: HANDLE, reason: u32, _reserved: *mut c_void) -> BOOL {
    match reason {
        DLL_PROCESS_ATTACH => {
            std::thread::spawn(|| {
                let file = std::fs::OpenOptions::new()
                    .create(true)
                    .write(true)
                    .open("redirect.log")
                    .unwrap();
                structured_logger::Builder::new()
                    .with_target_writer("*", structured_logger::json::new_writer(file))
                    .init();
            });
            unsafe {
                hook_CoCreateInstance.enable().unwrap();
                hook_CoCreateInstanceEx.enable().unwrap();
            };
        }
        DLL_PROCESS_DETACH => unsafe {
            hook_CoCreateInstance.disable().unwrap();
            hook_CoCreateInstanceEx.enable().unwrap();
        },
        DLL_THREAD_ATTACH => {}
        DLL_THREAD_DETACH => {}
        _ => {}
    };
    return BOOL::from(true);
}
