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
        if *riid == IMMDeviceEnumerator::IID {
            log::info!("!!! Intercepted IMMDeviceEnumerator creation, returning proxy !!!");
            let mut inner_raw: *mut c_void = null_mut();
            println!("Calling original CoCreateInstance...");
            let ret =
                hook_CoCreateInstance.call(rclsid, p_outer, dwcls_context, riid, &mut inner_raw);
            println!("Translating...");
            let inner_enumerator = IMMDeviceEnumerator::from_raw(inner_raw as _);
            println!("Redirecting...");
            let proxy_enumerator = RedirectDeviceEnumerator::new(inner_enumerator);
            let proxy_unknown: IMMDeviceEnumerator = proxy_enumerator.into();
            *ppv = proxy_unknown.into_raw() as _;
            ret
        } else {
            hook_CoCreateInstance.call(rclsid, p_outer, dwcls_context, riid, ppv)
        }
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
        if presults.is_null() || dwcount == 0 {
            return E_INVALIDARG;
        }
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

                    let proxy_enumerator = RedirectDeviceEnumerator::new(inner_enumerator);
                    let proxy_unknown: IMMDeviceEnumerator = proxy_enumerator.into();

                    _ = (*p_qi).pItf.insert(proxy_unknown.into());
                }
            }
        } else {
            log::error!("CoCreateInstanceEx call failed with HRESULT: {:?}", hr);
        }
        hr
    }
}

// /// 向整个系统广播一个“音频设备已更改”的事件
// fn broadcast_device_change() {
//     unsafe {
//         // 为了发送一个看起来真实的消息，我们最好先注册一个设备通知。
//         // 这会让我们的广播看起来更合法。
//         // 注意：这个注册需要一个窗口句柄。由于我们是DLL，我们可以创建一个隐藏的消息窗口。
//         // 但为了简化，我们可以尝试直接广播，通常也足够有效。

//         // 准备一个假的设备变更广播结构
//         // let mut dbi: DEV_BROADCAST_DEVICEINTERFACE_W = unsafe { std::mem::zeroed() };
//         // dbi.dbcc_size = std::mem::size_of::<DEV_BROADCAST_DEVICEINTERFACE_W>() as u32;
//         // dbi.dbcc_devicetype = DBT_DEVTYP_DEVICEINTERFACE;
//         // // 使用音频设备的GUID类别
//         // dbi.dbcc_classguid = windows::Win32::Media::Audio::DEVINTERFACE_AUDIO_RENDER; // 或 DEVINTERFACE_AUDIO_CAPTURE

//         // 向所有顶级窗口广播消息
//         // HWND_BROADCAST 是一个特殊值，表示发送给所有窗口
//         SendMessageA(
//             HWND_BROADCAST,
//             WM_DEVICECHANGE,
//             WPARAM(DBT_DEVICEARRIVAL),
//             LPARAM(()),
//         );

//         log::info!("Successfully broadcasted WM_DEVICECHANGE (DBT_DEVICEARRIVAL) to all windows.");
//     }
// }

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
        Ok(RedirectDevice::new(unsafe { self.inner.Item(ndevice)? }).into())
    }
}

#[implement(IMMDevice)]
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
        iid: *const windows_core::GUID,
        dwclsctx: windows::Win32::System::Com::CLSCTX,
        pactivationparams: *const windows::Win32::System::Com::StructuredStorage::PROPVARIANT,
        ppinterface: *mut *mut core::ffi::c_void,
    ) -> windows_core::Result<()> {
        if ppinterface.is_null() {
            return Err(Error::from(E_POINTER));
        }
        unsafe {
            let iid = *iid;
            if matches!(
                iid,
                IAudioClient::IID | IAudioClient2::IID | IAudioClient3::IID
            ) {
                log::info!("[PROXY] RedirectDevice::Activate called for IID: {:?}", iid);
                log::info!(
                    "[PROXY] RedirectDevice::Activate called, ppinterface: {:?}",
                    pactivationparams.as_ref()
                );
                let inner: IAudioClient3 = self.inner.Activate::<IAudioClient3>(
                    dwclsctx,
                    (!pactivationparams.is_null()).then(|| pactivationparams),
                )?;
                let proxy_client = RedirectAudioClient::new(inner);
                let proxy_unknown: IAudioClient3 = proxy_client.into();
                let requested_interface = proxy_unknown.cast::<IAudioClient3>()?;
                *ppinterface = requested_interface.into_raw();
                Ok(())
            } else {
                log::info!(
                    "[PROXY] RedirectDevice::Activate called for invalid IID: {:?}",
                    iid
                );
                Err(Error::from(E_NOINTERFACE))
            }
        }
    }

    fn OpenPropertyStore(
        &self,
        stgmaccess: windows::Win32::System::Com::STGM,
    ) -> windows_core::Result<IPropertyStore> {
        let real_store = unsafe { self.inner.OpenPropertyStore(stgmaccess)? };
        log::info!("[PROXY] RedirectDevice::OpenPropertyStore called");
        let proxy_store = RedirectPropertyStore::new(real_store);
        Ok(proxy_store.into())
    }

    fn GetId(&self) -> windows_core::Result<windows_core::PWSTR> {
        log::info!("[PROXY] RedirectDevice::GetId called");
        unsafe { self.inner.GetId() }
    }

    fn GetState(&self) -> windows_core::Result<DEVICE_STATE> {
        log::info!("[PROXY] RedirectDevice::GetState called");
        unsafe { self.inner.GetState() }
    }
}

#[implement(IPropertyStore)]
pub struct RedirectPropertyStore {
    inner: IPropertyStore,
    name: PWSTR,
}

impl RedirectPropertyStore {
    pub fn new(inner: IPropertyStore) -> Self {
        let mut pv = unsafe { inner.GetValue(&PKEY_Device_FriendlyName).unwrap() };
        let name = if pv.vt() == VT_LPWSTR {
            format!("{} (Redirected)", pv.to_string())
        } else {
            "Default (Redirected)".to_string()
        };
        _ = unsafe { PropVariantClear(&mut pv) };
        let wide_name: Vec<u16> = std::ffi::OsString::from(name)
            .encode_wide()
            .chain(Some(0))
            .collect();
        let byte_len = wide_name.len() * std::mem::size_of::<u16>();
        let pwsz = unsafe { CoTaskMemAlloc(byte_len) as *mut u16 };
        if pwsz.is_null() {
            panic!("null pointer");
        }
        unsafe {
            std::ptr::copy_nonoverlapping(wide_name.as_ptr(), pwsz, wide_name.len());
        }
        let name = PWSTR(pwsz);
        Self { inner, name }
    }
}

impl IPropertyStore_Impl for RedirectPropertyStore_Impl {
    fn GetValue(
        &self,
        key: *const windows::Win32::Foundation::PROPERTYKEY,
    ) -> windows_core::Result<PROPVARIANT> {
        // 检查是否是设备友好名称
        let mut pv = unsafe { self.inner.GetValue(key)? };
        if unsafe { *key } == PKEY_Device_FriendlyName {
            unsafe {
                PropVariantClear(&mut pv)?;
                (*pv.Anonymous.Anonymous).vt = VT_LPWSTR;
                (*pv.Anonymous.Anonymous).Anonymous.pwszVal = self.name;
            };
        }
        Ok(pv)
    }

    fn GetCount(&self) -> windows_core::Result<u32> {
        unsafe { self.inner.GetCount() }
    }

    fn GetAt(
        &self,
        iprop: u32,
        pkey: *mut windows::Win32::Foundation::PROPERTYKEY,
    ) -> windows_core::Result<()> {
        unsafe { self.inner.GetAt(iprop, pkey) }
    }

    fn SetValue(
        &self,
        key: *const windows::Win32::Foundation::PROPERTYKEY,
        propvar: *const windows::Win32::System::Com::StructuredStorage::PROPVARIANT,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] SetValue: Redirected.");
        unsafe { self.inner.SetValue(key, propvar) }
    }

    fn Commit(&self) -> windows_core::Result<()> {
        log::info!("[PROXY] Commit: Redirected.");
        unsafe { self.inner.Commit() }
    }
}

static NOTIFICATION_MANAGER: LazyLock<Arc<NotificationManager>> =
    LazyLock::new(|| Arc::new(NotificationManager::new()));

pub struct NotificationManager {
    callbacks: Mutex<HashMap<usize, Weak<IMMNotificationClient>>>,
}

impl NotificationManager {
    pub fn new() -> Self {
        Self {
            callbacks: Mutex::new(HashMap::new()),
        }
    }

    pub fn register(&self, client: &IMMNotificationClient) {
        // 将 COM 对象指针作为 key
        let ptr = client.as_raw() as usize;
        // 使用 Weak 引用，防止我们阻止程序释放回调对象
        let weak_client = client.downgrade().expect("client is missing?!?!");
        if let Ok(mut cb_map) = self.callbacks.lock() {
            cb_map.insert(ptr, weak_client);
            log::info!(
                "[PROXY] Registered IMMNotificationClient at {:p}",
                client.as_raw()
            );
        }
    }

    pub fn unregister(&self, client: &IMMNotificationClient) {
        let ptr = client.as_raw() as usize;
        if let Ok(mut cb_map) = self.callbacks.lock() {
            cb_map.remove(&ptr);
            log::info!(
                "[PROXY] Unregistered IMMNotificationClient at {:p}",
                client.as_raw()
            );
        }
    }

    // 这个方法将由我们自己的系统通知客户端调用
    pub fn forward_device_state_changed(&self, device_id: &str, new_state: DEVICE_STATE) {
        if let Ok(cb_map) = self.callbacks.lock() {
            // 收集所有有效的回调，避免在锁内调用外部代码
            let mut valid_clients = Vec::new();
            for weak_client in cb_map.values() {
                if let Some(client) = weak_client.upgrade() {
                    valid_clients.push(client);
                }
            }

            // 在锁外转发事件
            for client in valid_clients {
                log::info!(
                    "[PROXY] Forwarding OnDeviceStateChanged to {:p}",
                    client.as_raw()
                );
                // 将设备ID转换为 PWSTR
                let wide_id: Vec<u16> =
                    device_id.encode_utf16().chain(std::iter::once(0)).collect();
                unsafe {
                    let _ = client.OnDeviceStateChanged(PCWSTR(wide_id.as_ptr()), new_state);
                }
            }
        }
    }

    pub fn forward_device_added(&self, device_id: &str) {
        if let Ok(cb_map) = self.callbacks.lock() {
            // 收集所有有效的回调，避免在锁内调用外部代码
            let mut valid_clients = Vec::new();
            for weak_client in cb_map.values() {
                if let Some(client) = weak_client.upgrade() {
                    valid_clients.push(client);
                }
            }

            // 在锁外转发事件
            for client in valid_clients {
                log::info!("[PROXY] Forwarding OnDeviceAdded to {:p}", client.as_raw());
                // 将设备ID转换为 PWSTR
                let wide_id: Vec<u16> =
                    device_id.encode_utf16().chain(std::iter::once(0)).collect();
                unsafe {
                    let _ = client.OnDeviceAdded(PCWSTR(wide_id.as_ptr()));
                }
            }
        }
    }

    pub fn forward_device_removed(&self, device_id: &str) {
        if let Ok(cb_map) = self.callbacks.lock() {
            // 收集所有有效的回调，避免在锁内调用外部代码
            let mut valid_clients = Vec::new();
            for weak_client in cb_map.values() {
                if let Some(client) = weak_client.upgrade() {
                    valid_clients.push(client);
                }
            }

            // 在锁外转发事件
            for client in valid_clients {
                log::info!(
                    "[PROXY] Forwarding OnDeviceRemoved to {:p}",
                    client.as_raw()
                );
                // 将设备ID转换为 PWSTR
                let wide_id: Vec<u16> =
                    device_id.encode_utf16().chain(std::iter::once(0)).collect();
                unsafe {
                    let _ = client.OnDeviceRemoved(PCWSTR(wide_id.as_ptr()));
                }
            }
        }
    }

    pub fn forward_default_device_changed(&self, flow: EDataFlow, role: ERole, device_id: &str) {
        if let Ok(cb_map) = self.callbacks.lock() {
            // 收集所有有效的回调，避免在锁内调用外部代码
            let mut valid_clients = Vec::new();
            for weak_client in cb_map.values() {
                if let Some(client) = weak_client.upgrade() {
                    valid_clients.push(client);
                }
            }

            // 在锁外转发事件
            for client in valid_clients {
                log::info!(
                    "[PROXY] Forwarding OnDefaultDeviceChanged to {:p}",
                    client.as_raw()
                );
                // 将设备ID转换为 PWSTR
                let wide_id: Vec<u16> =
                    device_id.encode_utf16().chain(std::iter::once(0)).collect();
                unsafe {
                    let _ = client.OnDefaultDeviceChanged(flow, role, PCWSTR(wide_id.as_ptr()));
                }
            }
        }
    }

    pub fn forward_property_value_changed(&self, device_id: &str, key: &PROPERTYKEY) {
        if let Ok(cb_map) = self.callbacks.lock() {
            // 收集所有有效的回调，避免在锁内调用外部代码
            let mut valid_clients = Vec::new();
            for weak_client in cb_map.values() {
                if let Some(client) = weak_client.upgrade() {
                    valid_clients.push(client);
                }
            }

            // 在锁外转发事件
            for client in valid_clients {
                log::info!(
                    "[PROXY] Forwarding OnPropertyValueChanged to {:p}",
                    client.as_raw()
                );
                // 将设备ID转换为 PWSTR
                let wide_id: Vec<u16> =
                    device_id.encode_utf16().chain(std::iter::once(0)).collect();
                unsafe {
                    let _ = client.OnPropertyValueChanged(PCWSTR(wide_id.as_ptr()), *key);
                }
            }
        }
    }
}

#[implement(IMMNotificationClient)]
pub struct SystemNotificationClient;
impl SystemNotificationClient {
    pub fn new() -> Self {
        Self
    }
}
impl IMMNotificationClient_Impl for SystemNotificationClient_Impl {
    fn OnDeviceStateChanged(
        &self,
        device_id: &windows_core::PCWSTR,
        new_state: DEVICE_STATE,
    ) -> windows_core::Result<()> {
        let id_str = unsafe { device_id.to_string().unwrap_or_default() };
        log::info!(
            "[PROXY] System event: OnDeviceStateChanged for '{}'",
            id_str
        );
        NOTIFICATION_MANAGER.forward_device_state_changed(&id_str, new_state);
        Ok(())
    }

    fn OnDeviceAdded(&self, device_id: &windows_core::PCWSTR) -> windows_core::Result<()> {
        let id_str = unsafe { device_id.to_string().unwrap_or_default() };
        log::info!("[PROXY] System event: OnDeviceAdded for '{}'", id_str);
        NOTIFICATION_MANAGER.forward_device_added(&id_str);
        Ok(())
    }

    fn OnDeviceRemoved(&self, device_id: &windows_core::PCWSTR) -> windows_core::Result<()> {
        let id_str = unsafe { device_id.to_string().unwrap_or_default() };
        log::info!("[PROXY] System event: OnDeviceRemoved for '{}'", id_str);
        NOTIFICATION_MANAGER.forward_device_removed(&id_str);
        Ok(())
    }

    fn OnDefaultDeviceChanged(
        &self,
        flow: EDataFlow,
        role: ERole,
        default_device_id: &windows_core::PCWSTR,
    ) -> windows_core::Result<()> {
        let id_str = unsafe { default_device_id.to_string().unwrap_or_default() };
        log::info!(
            "[PROXY] System event: OnDefaultDeviceChanged for '{}' (flow: {:?}, role: {:?})",
            id_str,
            flow,
            role
        );
        NOTIFICATION_MANAGER.forward_default_device_changed(flow, role, &id_str);
        Ok(())
    }

    fn OnPropertyValueChanged(
        &self,
        device_id: &windows_core::PCWSTR,
        key: &PROPERTYKEY,
    ) -> windows_core::Result<()> {
        let id_str = unsafe { device_id.to_string().unwrap_or_default() };
        log::info!(
            "[PROXY] System event: OnPropertyValueChanged for '{}'",
            id_str
        );
        NOTIFICATION_MANAGER.forward_property_value_changed(&id_str, key);
        Ok(())
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
        log::info!("[PROXY] EnumAudioEndpoints called, creating redirected collection");
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
        let redirected = RedirectDevice::new(device);
        Ok(redirected.into())
    }

    fn GetDevice(&self, pwstrid: &windows_core::PCWSTR) -> windows_core::Result<IMMDevice> {
        log::info!("[PROXY] GetDevice() -> wrapping");
        let device = unsafe { self.inner.GetDevice(*pwstrid)? };
        let redirected = RedirectDevice::new(device);
        Ok(redirected.into())
    }

    fn RegisterEndpointNotificationCallback(
        &self,
        pclient: windows_core::Ref<IMMNotificationClient>,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] RegisterEndpointNotificationCallback called");

        // 1. 将程序的回调添加到我们的管理器
        NOTIFICATION_MANAGER.register(pclient.as_ref().unwrap());

        // 2. 确保我们自己的系统客户端已经注册到真实的系统中
        // (这里需要一个标志位来避免重复注册)
        static IS_OUR_CLIENT_REGISTERED: AtomicBool = AtomicBool::new(false);
        unsafe {
            if !IS_OUR_CLIENT_REGISTERED.load(Ordering::Acquire) {
                let our_client: IMMNotificationClient = SystemNotificationClient::new().into();
                self.inner
                    .RegisterEndpointNotificationCallback(&our_client)?;
                log::info!(
                    "[PROXY] Our SystemNotificationClient has been registered with the system."
                );
                IS_OUR_CLIENT_REGISTERED.store(true, Ordering::Release);
            }
        }
        Ok(())
    }

    fn UnregisterEndpointNotificationCallback(
        &self,
        pclient: windows_core::Ref<IMMNotificationClient>,
    ) -> windows_core::Result<()> {
        log::info!("[PROXY] UnregisterEndpointNotificationCallback called");

        // 1. 从我们的管理器中移除程序的回调
        NOTIFICATION_MANAGER.unregister(pclient.as_ref().unwrap());

        // 注意：我们不应该注销我们自己的系统客户端，因为其他程序可能也需要它。
        // 让它一直存在直到进程结束是更安全的选择。
        Ok(())
    }
}

#[implement(IAudioClient, IAudioClient2, IAudioClient3)]
struct RedirectAudioClient {
    inner: IAudioClient3,
}

impl RedirectAudioClient {
    fn new(inner: IAudioClient3) -> Self {
        Self { inner }
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
        log::info!(
            "[PROXY] Initialize() -> Forwarding to real client, sharemode: {}",
            sharemode.0
        );
        log::info!(
            "[PROXY] Initialize() -> Forwarding to real client, streamflags: {}",
            streamflags
        );
        log::info!(
            "[PROXY] Initialize() -> Forwarding to real client, hnsbufferduration: {}",
            hnsbufferduration
        );
        log::info!(
            "[PROXY] Initialize() -> Forwarding to real client, hnsperiodicity: {}",
            hnsperiodicity
        );
        log::info!(
            "[PROXY] Initialize() -> Forwarding to real client, pformat: {:?}",
            unsafe { (*pformat).nChannels }
        );
        log::info!(
            "[PROXY] Initialize() -> Forwarding to real client, guid: {:?}",
            unsafe { audiosessionguid.as_ref() }
        );
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
        log::info!("[PROXY] GetCurrentPadding() -> Forwarding to real client");
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
            simple_logger::SimpleLogger::new().init().unwrap();
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
