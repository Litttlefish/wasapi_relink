#[cfg(test)]
use crate::RedirectDeviceEnumerator; // 替换成你的模块路径
use crate::{HOOK_CO_CREATE_INSTANCE_EX, RedirectConfig};
// use openal_binds::*;
#[cfg(test)]
use windows::{
    Win32::{
        Devices::FunctionDiscovery::PKEY_Device_FriendlyName, Media::Audio::*, System::Com::*,
    },
    core::*,
};

#[cfg(test)]
/// 辅助函数：查找我们修改过的设备
fn find_redirected_device(collection: &IMMDeviceCollection) -> Result<Vec<IMMDevice>> {
    let count = unsafe { collection.GetCount()? };
    println!("Found {} audio devices.", count);

    // 4. 遍历设备，找到 ID 以 "(redirected)" 结尾的那个
    let device = (0..count)
        .map(|i| unsafe { collection.Item(i).unwrap() })
        .collect();
    return Ok(device);
}

#[test]
#[inline(never)]
fn test_property_store_and_audio_client() -> Result<()> {
    // 初始化 COM
    _ = unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) };

    // 1. 创建真实的设备枚举器
    let enumerator: IMMDeviceEnumerator =
        unsafe { CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)? };

    // 2. 使用我们的代理枚举器来包装真实的枚举器
    println!("\n--- Testing NewEnum ---");
    let proxy_enumerator = RedirectDeviceEnumerator::new(enumerator);

    // 3. 枚举所有活动设备
    println!("\n--- Testing Enum ---");
    let collection = proxy_enumerator
        .into_outer()
        .EnumAudioEndpoints(windows::Win32::Media::Audio::eAll, DEVICE_STATE_ACTIVE)?;

    // --- 测试 1: IPropertyStore ---
    println!("\n--- Testing IPropertyStore ---");
    let devices = find_redirected_device(&collection)?;
    for device in devices {
        // 激活 IPropertyStore
        let prop_store = unsafe { device.OpenPropertyStore(STGM_READ)? };

        // 获取设备友好名称
        let friendly_name_key = PKEY_Device_FriendlyName;
        let name_variant = unsafe { prop_store.GetValue(&friendly_name_key)? };

        let name = name_variant.to_string();
        println!("Device Friendly Name from Proxy: {}", name);

        // --- 测试 2: IAudioClient ---
        println!("\n--- Testing IAudioClient ---");

        // 激活 IAudioClient
        let audio_client = unsafe { device.Activate::<IAudioClient>(CLSCTX_ALL, None)? };

        // 获取混音格式（这个调用应该被透明转发）
        let mix_format = unsafe { audio_client.GetMixFormat()? };
        println!("Successfully got mix format via proxy: {:?}", mix_format);

        // 尝试初始化客户端（这个调用应该被我们的代理处理）
        // 注意：这里的参数可能需要根据你的系统调整
        let result = unsafe {
            audio_client.Initialize(
                windows::Win32::Media::Audio::AUDCLNT_SHAREMODE_SHARED,
                windows::Win32::Media::Audio::AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                10000000, // 1 second
                0,
                mix_format,
                None,
            )
        };

        assert!(
            result.is_ok(),
            "IAudioClient::Initialize failed via proxy: {:?}",
            result
        );
        println!("Successfully initialized IAudioClient via proxy.");
    }

    // 清理 COM
    unsafe { windows::Win32::System::Com::CoUninitialize() };
    Ok(())
}

// #[test]
// fn test_cocreateinstance_hook() {
//     unsafe {
//         CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
//         simple_logger::SimpleLogger::new().init().unwrap();
//         hook_CoCreateInstance.enable().unwrap();
//         let enumerator: IMMDeviceEnumerator =
//             CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL).unwrap();

//         // 这里可以添加更多检查来验证它确实是我们的代理
//         println!("\n--- Testing Enum ---");
//         let collection = unsafe {
//             enumerator
//                 .EnumAudioEndpoints(windows::Win32::Media::Audio::eAll, DEVICE_STATE_ACTIVE)
//                 .unwrap()
//         };

//         // --- 测试 1: IPropertyStore ---
//         println!("\n--- Testing IPropertyStore ---");
//         let devices = find_redirected_device(&collection).unwrap();
//         for device in devices {
//             // 激活 IPropertyStore
//             let prop_store = unsafe { device.OpenPropertyStore(STGM_READ).unwrap() };

//             // 获取设备友好名称
//             let friendly_name_key = PKEY_Device_FriendlyName;
//             let name_variant = unsafe { prop_store.GetValue(&friendly_name_key).unwrap() };

//             // 验证名称是否被修改
//             let name = unsafe { name_variant.to_string() };
//             println!("Device Friendly Name from Proxy: {}", name);
//             assert!(
//                 name.ends_with(" (Redirected)"),
//                 "Friendly name was not modified by proxy!"
//             );
//         }

//         CoUninitialize();
//     }
// }

#[test]
fn test_cocreateinstance_ex_hook() {
    unsafe {
        CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
        simple_logger::SimpleLogger::new().init().unwrap();
        HOOK_CO_CREATE_INSTANCE_EX.enable().unwrap();
        let mut result = [MULTI_QI::default()];
        result[0].pIID = &IMMDeviceEnumerator::IID;
        dbg!(RedirectConfig::load());
        CoCreateInstanceEx(&MMDeviceEnumerator, None, CLSCTX_ALL, None, &mut result).unwrap();
        let enumerator = result[0]
            .pItf
            .take()
            .unwrap()
            .cast::<IMMDeviceEnumerator>()
            .unwrap();

        // 这里可以添加更多检查来验证它确实是我们的代理
        println!("\n--- Testing Enum ---");
        let collection = enumerator
            .EnumAudioEndpoints(windows::Win32::Media::Audio::eAll, DEVICE_STATE_ACTIVE)
            .unwrap();

        // --- 测试 1: IPropertyStore ---
        println!("\n--- Testing IPropertyStore ---");
        let devices = find_redirected_device(&collection).unwrap();
        for device in devices {
            // 激活 IPropertyStore
            let prop_store = device.OpenPropertyStore(STGM_READ).unwrap();

            // 获取设备友好名称
            let friendly_name_key = PKEY_Device_FriendlyName;
            let name_variant = prop_store.GetValue(&friendly_name_key).unwrap();

            let name = name_variant.to_string();
            println!("Device Friendly Name from Proxy: {}", name);
        }

        CoUninitialize();
    }
}
