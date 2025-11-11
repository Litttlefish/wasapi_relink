# WASAPI Relink
[English](./readme.md) | 中文

`wasapi_relink` 是一个 hook 库，它修改 WASAPI (Windows Audio Session API) 共享流的行为，将其“嫁接”到现代的低延迟接口上，以最大限度地减少音频延迟。

## 🎯目的
许多应用程序和游戏使用 WASAPI 共享模式播放音频。虽然兼容性好，但此模式默认 10ms 的大缓冲区，导致了明显的高延迟。

`wasapi_relink` 拦截应用程序的音频请求，并利用 `IAudioClient3` 接口强制它们使用更小的缓冲区，在不牺牲共享模式便利性的前提下，实现接近独占模式的低延迟体验。

### ❗注意：你的驱动非常重要
本工具完全依赖于你的音频驱动程序。配置中的 target_buffer_dur_ms 是一个***请求值***，而非*强制命令*。

你的音频驱动会报告一个支持的缓冲区范围 (即 `GetSharedModeEnginePeriod`)，`wasapi_relink` 会将请求限制在此范围内。这个问题在 Realtek 板载声卡上尤为明显，详见下文。

#### Realtek 的问题
大多数 Realtek 原厂驱动会将其缓冲区范围**完全无效**为 10ms。

如果你的驱动如此， `wasapi_relink` 将**毫无效果**，因为它无法请求小于 10ms 的缓冲区。

#### 解决方案
你必须**卸载 Realtek 驱动**，并强制 Windows 安装通用的 **"High Definition Audio Device" 驱动程序**。这个驱动通常允许 2ms-10ms 的范围，使得 `wasapi_relink` 能成功请求 2ms 的缓冲区。

条件允许的话，你也可以使用支持低延迟的专业声卡或外置音频接口。

## 🚀工作原理：“移花接木”
`wasapi_relink` 以两种不同模式运行，以处理不同类型的应用程序。

### 普通模式 (适用于事件驱动型应用)
**目标：** 使用事件驱动 (基于回调) 音频流的现代应用程序和游戏。

**方法：**

1. Hook `IAudioClient::Initialize` 。

2. 忽略应用程序的缓冲区请求，转而调用 `IAudioClient3::InitializeSharedAudioStream` 来初始化一个低延迟流。

3. 这为应用程序提供了一个更小的缓冲区 (例如 2ms)，从而显著降低延迟。

### 兼容模式 (适用于轮询驱动型应用)
    这是本项目的核心，专为使用轮询 (基于循环) 音频流的应用程序设计。

**目标：** 使用轮询驱动流的应用程序和游戏。

**方法：**

1. Hook `IMMDevice::Activate` 并创建**两个** `IAudioClient3` 实例。

2. **实例 A (诱饵)：** 以应用程序期望的常规共享模式初始化。

3. **实例 B (触发器)：** 以 `IAudioClient3` 的 __低延迟模式__初始化。这是关键：它会[**触发 Windows 的特殊行为**](https://learn.microsoft.com/en-us/windows-hardware/drivers/audio/low-latency-audio#faq)，强制实际的硬件音频引擎缓冲区与此低延迟请求 (例如 2ms) 对齐。

4. **预填充欺骗：** `wasapi_relink` 接着拦截应用程序的静默预填充 (Prefill) 过程，**欺骗**它填充一个与新的硬件缓冲区大小相匹配的、更小的缓冲区。

**结果：** 最差情况下，延迟是普通模式的两倍 (预填充缓冲区 + 设备缓冲区)。最佳情况下，与普通模式一致。

## 🔧如何使用 (注入)
本工具**不包含**自己的 DLL 注入器。你必须使用外部工具或[stub生成器](https://github.com/namazso/dll-proxy-generator)。

### 推荐方法：Special K
本工具设计为与 **Special K** 完美协作。

推荐将 `wasapi_relink.dll` 作为 Special K 的“插件” (plug-in) ，以 lazy mode 加载。

**内置兼容性：** `wasapi_relink` 会主动检测并 **跳过** 对 Special K 自身音频请求的 hook ， 确保 SK 的 OSD 声音和其他功能不受影响。

### 关于使用stub生成器的提醒
**请勿使用 `ole32.dll` 。** 这样做可能在 `LoadLibrary`/COM 激活期间导致**无限递归**，并导致不稳定或难以调试/预测的行为 (例如，重复的加载器调用、崩溃或目标进程挂起)。

**原因：** `ole32.dll` 深度参与 COM 组件的操作过程。创建转发或重新导出 COM/ole32 入口点的 stub 很容易创建循环调用，导致加载器重复调用相同的激活路径。

## ⚙️️配置（`redirect_config.toml`）
配置文件 `redirect_config.toml` **必须**放置在 DLL 所在的目录（或工作目录）下。如果未指定条目，将使用默认值。
```toml
# 日志文件路径。"" (空字符串) 默认为当前工作目录。
log_path = ""
# 日志级别: Trace, Debug, Info, Warn, Error
log_level = "Info"

[playback]
# 目标缓冲区时长，单位 0.1ms (u16)。
# 工具将计算最接近但不超过此持续时间、且在驱动允许范围内的缓冲区大小。
# 例如：10 = 1ms。
target_buffer_dur_ms = 10
# 强制此流进入兼容模式 (bool)
compat = false

[capture]
target_buffer_dur_ms = 10
compat = false
```

### 配置详情
- `log_path` (string): 保存日志文件的位置。默认或 `""` 表示当前工作目录。

- `log_level` (string): `Trace`, `Debug`, `Info`, `Warn`, `Error` 。 默认是 `Info`。  
**请参阅下面的性能警告。**

- `[playback]`/`[capture]`: 分别配置输出和输入。

  - `target_buffer_dur_ms` (u16): 目标缓冲区大小，单位为 **0.1 毫秒**。如果设置过低或未指定，将默认为驱动的最小值。**除非遇到音频爆音，否则通常不应更改此值。**

  - `compat` (bool): 强制此流使用**兼容**模式。

## 🩺故障排查
使用本指南诊断和修复常见的音频问题。

### 音频“切片”或“慢放”
**现象：** 声音严重失真、拉伸，或听起来像是被“切片”后缓慢播放。

**原因：** 这是**轮询驱动**的应用程序在**普通模式** (`compat = false`) 下运行的典型标志。应用的轮询逻辑与事件驱动的缓冲区发生冲突。这在 **Unity** 引擎的游戏中很常见。

**解决方案：** 在 `[playback]` 部分设置 `compat = true`。

### 音频正常，但偶尔有“爆音”或“噼啪声”
**现象：** 音频播放速度和音高都正确，但你会听到断断续续的爆音、咔嗒声或轻微的撕裂声。

**原因：** 缓冲区对于你的系统或软件来说**太小了**。你的 CPU 或应用程序无法足够快地“喂给”音频驱动程序新数据，导致缓冲区欠载 (buffer underrun)。

**解决方案（按顺序尝试）：** 
1. **增加缓冲区：** 略微增加 `target_buffer_dur_ms` 。如果驱动的最小值是 2ms (即 `20`)，请尝试 `30` (3ms) 或 `40` (4ms)，直到爆音消失。

2. **尝试兼容模式：** 设置 `compat = true` 。这相当于为程序增加了一个缓冲层（即普通Shared模式给出的buffer），结合Compat模式的改变可能能解决问题。

### 程序无法启动
**现象：** 加载 DLL 后目标程序启动失败。

**原因：** 这通常是**文件权限问题** —— DLL 尝试在没有写入权限的位置写入日志文件。这通常发生在目标程序位于 `Program Files` 或其他受保护目录时。

**解决方案：** 在 `redirect_config.toml` 中为 `log_path` 指定一个有写入权限的位置。  
例如：
```toml
log_path = "C:\\Users\\Public\\wasapi_relink.log"
```

## ⚠️性能警告：日志与音频撕裂
Windows 音频引擎对时序 _极其_ 敏感。

 - **请勿** 在正常使用中将 `log_level` 设置为 `Debug` 或 `Trace` 。

 - 详细日志记录产生的高频磁盘操作**会**干扰音频线程，尤其是在缓冲区比较小的时候。

 - 这种干扰**完全有可能**导致音频撕裂、爆音和卡顿。

 - 只有在您明确知道自己在做什么，并主动为本工具进行开发调试时，才应该使用 `Debug`/`Trace`。

## 🔬深度解析：技术实现
### “外部低延迟”的误解
以前的工具([REAL](https://github.com/miniant-git/REAL), [LowAudioLatency](https://github.com/spddl/LowAudioLatency) 及类似项目) 试图通过在后台运行一个单独的低延迟 `IAudioClient3` 应用程序/线程来解决此问题。

 - **没问题的部分：** 正如 Windows FAQ 所述，这*确实*会触发 Windows 的特殊行为，并*强制*硬件缓冲区为所有应用切换到更小尺寸。

 - **被忽视的部分：** 他们假设降低硬件缓冲区就够用了。**但事实并非如此。** 一个轮询驱动的应用程序*仍然*会尝试静默预填充 (Prefill) 其旧的 10ms 缓冲区。延迟瓶颈只是从硬件转移到了应用程序自己的预填充上。

`wasapi_relink` 的**兼容模式***同时*解决了这两个问题：它触发了硬件变更，并从内部 hook 应用程序，以*欺骗*和*修改*其静默预填充行为来匹配新的、更小的硬件缓冲区。

## 🧩实现：COM 包装链
`wasapi_relink` 不使用简单的 VTable 钩子。它从 COM 音频系统的根部开始hook过程，以确保不会漏过音频行为。

1. Hook `CoCreateInstanceEx`.

2. 当应用请求 `CLSID_MMDeviceEnumerator` (主音频设备枚举器) `时，wasapi_relink` 创建真实对象，并向应用返回一个**带有包装层**的 `IMMDeviceEnumerator`。

3. 当应用在这个组件上请求设备时， `wasapi_relink` 返回一个包装的 `IMMDevice` 。

    3.1. **兼容模式：** 当应用在包装的 `IMMDevice` 上调用 `Activate` `时，wasapi_relink` 执行其核心的兼容模式逻辑 (创建两个客户端) 并返回*一个*包装的 `IAudioClient` 。

    3.2.**普通模式：** 当应用在包装的 `IAudioClient` 上调用 `Initialize` 时，将执行普通模式的逻辑。

4. 这条包装链一直延续到 `IAudioRenderClient`，使 `wasapi_relink` 能够完全、透明地控制整个音频流的生命周期。

### 关于捕获 (麦克风) 的说明
在**兼容模式**下，捕获 (麦克风) 流只会简单转发。这是*有意为之*：

1. 在共享模式下，与效果或网络延迟相比，麦克风延迟的重要性要低得多，并且基于捕获设备驱动和应用的交互原理（无需填充静默区），触发 Windows 特殊模式本身就已经能降低延迟了。
2. 真正需要低延迟输入的用户都已经在用 ASIO 或 WASAPI 独占模式了，用不到 `wasapi_relink` 。

---
### **免责声明：**
该项目仅供实验、学习和交流使用。 

请注意：
- hook 行为可能会在某些游戏中触发反作弊检测。
- WASAPI 的内部机制和缓冲区行为可能因 Windows 版本或驱动程序而异。
- 请负责任地使用 —— 风险自负。