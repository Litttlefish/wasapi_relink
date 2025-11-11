# WASAPI Relink
English | [中文](./readme_zh-CN.md)

`wasapi_relink` is a hook library that modifies the behavior of WASAPI (Windows Audio Session API) Shared streams, "grafting" them onto modern, low-latency interfaces to drastically minimize audio latency.

## 🎯Purpose
Many applications and games use WASAPI's Shared mode for audio playback. While compatible, this mode defaults to a large 10ms buffer, creating noticeable, high-latency audio.

`wasapi_relink` intercepts the application's audio requests and forces them to use much smaller buffers by leveraging the `IAudioClient3` interface, achieving a low-latency experience similar to Exclusive mode without sacrificing Shared mode's convenience.

### ❗Reminder: Your Driver is Everything
This tool is 100% dependent on your audio driver. The target_buffer_dur_ms setting in config is a ***request***, not a *command*.

Your audio driver reports a supported buffer range (i.e. `GetSharedModeEnginePeriod`), and `wasapi_relink` will clamp the request to fit within that range, this problem is most obvious when using Realtek onboard soundcards, see below.

#### The Realtek Problem
Most stock Realtek audio drivers (the ones from the manufacturer) lock the buffer range to 10ms _exact_.

If your driver does this, wasapi_relink will have no effect, as it cannot request a buffer smaller than 10ms.

#### The Solution
You must uninstall the Realtek drivers and force Windows to install the generic "High Definition Audio Device" driver. This standard Microsoft driver typically allows a 2ms-10ms range, enabling `wasapi_relink` to successfully request a 2ms buffer.

Or, use a professional soundcard or external audio interface that supports low-latency.

## 🚀How It Works: The "Grafting"
`wasapi_relink` operates in two distinct modes to handle different types of applications.

### Normal Mode (For Event-Driven Apps)
**Target:** Modern applications and games that use event-driven (callback-based) audio streaming.

**Method:**

1. Hooks `IAudioClient::Initialize`.

2. Ignores the application's buffer request and instead calls `IAudioClient3::InitializeSharedAudioStream` to initialize a low-latency stream.

3. This provides the application with a much smaller buffer (e.g., 2ms), dramatically reducing latency.

### Compat Mode (For Poll-Driven Apps)
    This is the core of the project, designed for applications that uses polling (loop-based) audio stream.

**Target:** Applications and games that use poll-driven streaming.

**Method:**

1. Hooks `IMMDevice::Activate` and creates **two** `IAudioClient3` instances.

2. **Instance A (The "Decoy"):** Is initialized in a regular Shared mode that the application expects.

3. **Instance B (The "Hooker"):** Is initialized in `IAudioClient3`'s __low-latency mode__. This is the key: it [**triggers a special Windows behavior**](https://learn.microsoft.com/en-us/windows-hardware/drivers/audio/low-latency-audio#faq) that forces the actual hardware audio engine buffer to align with this low-latency request (e.g. 2ms).

4. **Prefill Deception:** `wasapi_relink` then intercepts the application's silent prefill process, **deceiving it** into filling a much smaller buffer that matches the new hardware buffer size.

**Result:** At worst, latency is double that of Normal Mode (Prefill Buffer + Device Buffer). At best, it's identical to Normal Mode.

## 🔧How to Use (Injection)
This library **does not** include its own DLL injector. You must use an external tool or [stub generator](https://github.com/namazso/dll-proxy-generator).

### Recommended Method: Special K
This library is designed to work seamlessly with **Special K**.

It is recommended to load `wasapi_relink.dll` as a Special K "plug-in" (lazy mode).

**Built-in Compatibility:** `wasapi_relink` actively detects and **skips** hooking Special K's own audio requests, ensuring SK's OSD sounds and other features are not affected.

### Reminder on using stub generator
**DO NOT use `ole32.dll` when generating DLL stubs.** Doing so can cause infinite recursion during `LoadLibrary`/COM activation and lead to unstable or hard-to-debug behavior (e.g., repeated loader calls, crashes, or the target process hanging).

**Why:** `ole32.dll` is heavily involved in COM activation and in-process marshaling. Creating stubs that forward or re-export COM/ole32 entry points can easily create circular calls where the loader repeatedly invokes the same activation paths.

## ⚙️Configuration (`redirect_config.toml`)
The configuration file `redirect_config.toml` **must** be placed in the same directory(or working directory) as the DLL. If an entry is not specified, default value will be used.
```toml
# Path for the log file. "" (empty string) defaults to the working directory.
log_path = ""
# Log level: Trace, Debug, Info, Warn, Error
log_level = "Info"

[playback]
# Target buffer duration in 0.1ms units (u16).
# The tool will calculate the closest buffer size *not exceeding* this duration while clamped within driver range.
# e.g., 10 = 1ms.
target_buffer_dur_ms = 10
# Force this stream into compat mode (bool)
compat = false

[capture]
target_buffer_dur_ms = 10
compat = false
```

### Config Details
- `log_path` (string): Where to save the log file. Default or `""` means current working directory.

- `log_level` (string): `Trace`, `Debug`, `Info`, `Warn`, `Error`. Default is `Info`.  
**See performance warning below.**

- `[playback]`/`[capture]`: Separate configs for output and input.

  - `target_buffer_dur_ms` (u16): The target buffer size in **units of 0.1 milliseconds**. The tool will default to the driver's minimum if this is set too low or not specified. **You should generally not change this from the default low value unless you experience audio pops.**

  - `compat` (bool): Forces this stream to use **Compat Mode**.

## 🩺Troubleshooting
Use this guide to diagnose and fix common audio issues.

### Audio is "Sliced" or in "Slow-Motion"
**Phenomenon:** Sound is heavily distorted, stretched, or sounds like it's being "sliced" and played back slowly.

**Cause:** This is the classic sign of a **Poll-driven** application running in **Normal Mode** (`compat = false`). The app's polling logic is fighting the event-driven buffer. This is commonly seen in **Unity** games.

**Solution:** Set `compat = true` for the `[playback]` section.

### Good Audio with Occasional "Pops" or "Crackles"
**Phenomenon:** Audio playback is at the correct speed and pitch, but you hear intermittent pops, clicks, or small tearing sounds.

**Cause:** The buffer is **too small** for your system/software. Your CPU or the application cannot "feed" the audio driver new data fast enough, resulting in a buffer underrun.

**Solutions (Try in order):** 
1. **Increase Buffer:** Slightly increase `target_buffer_dur_ms`. If your driver's minimum is 2ms (e.g., `20`), try `30` (3ms) or `40` (4ms) until the pops disappear.

2. **Try Compat Mode:** Set `compat = true`. This effectively adds a buffer layer for the application (i.e., the standard buffer from normal Shared mode). In combination with the changes made by Compat mode, this can potentially resolve the issue.

### Program won't start
**Phenomenon:** The target program fails to start after loading the DLL.

**Cause:** This is usually a **file permission issue** — the DLL attempts to write its log file to a location where it lacks write access.
It commonly occurs when the target executable resides under `Program Files` or other protected directories.

**Solution:** Specify a writable location for `log_path` in `redirect_config.toml`.  
For example:
```toml
log_path = "C:\\Users\\Public\\wasapi_relink.log"
```

## ⚠️Performance Warning: Logging & Audio Tearing
The Windows audio engine is _extremely_ sensitive to timing.

 - **DO NOT** set `log_level` to `Debug` or `Trace` for normal use.

 - The high-frequency I/O (disk writing) from detailed logging **will** interfere with the audio thread, especially at low buffer sizes.

 - This interference ***will* cause audio tearing, pops, and stuttering**.

 - ONLY use `Debug`/`Trace` if you are actively debugging the tool itself and you know what you are doing.

## 🔬Deep Dive: Technical Implementation
### The "Myth" of External Low-Latency Activators
Previous tools([REAL](https://github.com/miniant-git/REAL), [LowAudioLatency](https://github.com/spddl/LowAudioLatency) and similar projects) attempted to solve this by running a _separate_ low-latency `IAudioClient3` application in the background.

 - **What they got right:** As said the Windows FAQ, this _does_ trigger the special Windows behavior and forces the _hardware_ buffer to a smaller size for all apps.

 - **What they missed (The "Myth"):** They assumed lowering the hardware buffer was enough. **It isn't.** A poll-driven application will _still_ try to **prefill** its old 10ms buffer. The latency bottleneck just moves from the hardware to the application's own prefill.

 `wasapi_relink`'s **Compat Mode** solves _both_ problems: it triggers the hardware change and internally hooks the application to _deceive and modify its prefill_ to match the new, smaller hardware buffer.

## 🧩Implementation: The COM Wrapper Chain

`wasapi_relink` does not use simple VTable hooking. It hooks at the root of the COM audio system to ensure 100% capture.

1. **Hook:** `CoCreateInstanceEx`.

2. **Intercept:** When the app requests `CLSID_MMDeviceEnumerator` (the main audio device enumerator), `wasapi_relink` creates the real one and returns a custom **wrapped** `IMMDeviceEnumerator` to the app.

3. **Propagate:** When the app requests device on this wrapper, `wasapi_relink` returns a wrapped `IMMDevice`.

    3.1. **Execute (Compat Mode):** When the app calls `Activate` on the wrapped `IMMDevice`, wasapi_relink executes its core Compat Mode logic (creating the two clients) and returns _one_ wrapped `IAudioClient`.

    3.2.**Execute (Normal Mode):** When the app calls `Initialize` on the wrapped `IAudioClient`, the Normal Mode logic is executed.

4. This wrapper chain continues all the way down to `IAudioRenderClient`, giving `wasapi_relink` full, transparent control over the entire audio stream lifecycle.

### A Note on Capture (Microphone)
In **Compat Mode**, capture (mic) streams are simply forwarded. This *is* intentional:

1. Mic latency is far less critical in Shared mode compared to effects or network latency. Due to the nature of capture streams, which don’t require a silent pre-fill buffer in their interaction with the driver, triggering Windows’ special behavior is by itself sufficient to reduce latency.
2. Users with true low-latency input needs are already using `ASIO` or `WASAPI Exclusive` mode, which this tool does not target.

---
### **Disclaimer:**
This is an **experimental** and **educational** project.  

Keep in mind:
- Hooking may trigger anti-cheat detections in some games.
- WASAPI internals and buffer behavior can differ between Windows builds or drivers.
- Use it responsibly — at your own risk.