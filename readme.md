# WASAPI Relink

English | [中文](./readme_zh-CN.md)

`wasapi_relink` is a hook library that modifies the behavior of WASAPI (Windows Audio Session API) Shared streams, "grafting" them onto modern, low-latency interfaces to drastically minimize audio latency.

## Purpose

Many applications and games use WASAPI's Shared mode for audio playback. While it offers good compatibility, this mode only allows for a minimum buffer size of 10ms, preventing further improvement in latency.

`wasapi_relink` intercepts the application's audio requests and forces them to use much smaller buffers by leveraging the `IAudioClient3` interface, achieving a low-latency experience similar to Exclusive mode without sacrificing Shared mode's convenience.

### Reminder: Your Driver is Everything

This tool is 100% dependent on your audio driver. The `target_buffer_dur_ms` setting in config is a ***request***, not a *command*.

Your audio driver reports a supported buffer range (i.e. `GetSharedModeEnginePeriod`), and `wasapi_relink` will clamp the request to fit within that range, this problem is most obvious when using Realtek onboard soundcards, see below.

#### The Realtek Problem

Most stock Realtek audio drivers (the ones from the manufacturer) lock the buffer range to 10ms *exact*.

If your driver does this, wasapi_relink will have no effect, as it cannot request a buffer smaller than 10ms.

#### The Solution

You must replace the Realtek driver with Windows generic "High Definition Audio Device" driver. This standard Microsoft driver typically allows a 2ms-10ms range, enabling `wasapi_relink` to successfully request a 2ms buffer.

Or, use a professional soundcard or external audio interface that supports low-latency.

P.S.: Currently, there don't seem to be many sound cards that are compatible with Low Latency Shared mode. The best option right now is probably to use the USB dongle that comes with Windows' built-in USB 2.0 driver.

## How It Works: The "Grafting"

`wasapi_relink` operates in four distinct modes to handle different types of applications.

### Normal Mode (For modern apps)

**Target:** Applications and games that use event‑driven (callback) WASAPI audio and behave well.

**Method:**

1. Hooks `IAudioClient::Initialize`.

2. Ignores the application's buffer request and instead calls `IAudioClient3::InitializeSharedAudioStream` to initialize a low-latency stream.

3. This provides the application with a much smaller buffer (e.g., 2ms), dramatically reducing latency.

### Compat Mode (For compatibility with legacy/poorly coded apps)

**Target:** Programs that break if you only shrink their buffer, and Ringbuf mode doesn't work well with them.

**Method:**

1. Hooks `IMMDevice::Activate` and creates **two** `IAudioClient3` instances.

- **Instance A (App‑facing):** Is initialized in a regular Shared mode that the program expects.

- **Instance B (Low‑latency):** Is initialized with `IAudioClient3::InitializeSharedAudioStream`using the tool’s low‑latency period. This client may not be used by the app at all; its main job is to influence the Windows audio engine.

1. **Engine‑side effect:** According to [**Microsoft’s low‑latency documentation**](https://learn.microsoft.com/en-us/windows-hardware/drivers/audio/low-latency-audio#faq), when any application on an endpoint requests small buffers via `IAudioClient3`, the audio engine switches to that small period for all shared‑mode streams on the same endpoint. By keeping the low‑latency client alive, `wasapi_relink` forces the engine to run at the small period even though the “main” app client is using a larger buffer.

2. **Prefill Deception:** Before `IAudioClient::Start`, many apps will write “silent” prefill data. `wasapi_relink` intercepts this via the wrapped `IAudioRenderClient`, then writes a smaller prefill (sized to the real engine period) into the hardware client, non-silent data won't be affected.'

**Result:** The app continues to work as if it had its original large buffer.

- **Worst-case:** Latency is doubled comparing to normal mode (the sum of the pre-filled buffer and the device buffer).
- **Best-case:** Latency is close to the normal mode.

### Ringbuf Mode (Most powerful)

**Target:** Programs that use polling or fixed‑size mixing and/or have problematic timing behavior.

**Method:**

1. Intercepts `Initialize` / `InitializeSharedAudioStream` and always calls `InitializeSharedAudioStream` with the tool’s low‑latency period.

2. Creates an internal ring buffer whose capacity is exposed to the app as the reported `IAudioClient::GetBufferSize`, and `IAudioClient::GetCurrentPadding` reports how much data is in the ring buffer, not the real engine buffer. The app feels like it’s writing into a large, traditional buffer, while the engine actually runs at the small period.

3. Injection and inverse modes

- If the application doesn't use event-driven flag, `wasapi_relink` injects `AUDCLNT_STREAMFLAGS_EVENTCALLBACK` and runs its own consumer thread, which is based on Real-Time Work Queue API, that waits on real engine event and move the data from the ring buffer to `IAudioRenderClient` on the real client.
- If the application requested event-driven flag, `wasapi_relink` takes over the app’s event handle and decides when to call `SetEvent` to wake the app’s callback. Internally it still uses its own thread to consume from the ring buffer. In effect, the chain becomes:

```text
hardware engine → ringbuf consumer thread → app callback
```

All silent data before `Start()` will be discarded.

**Result:** The app just sees a large, friendly WASAPI client, fully isolated from the engine’s real timing and buffer size, which works even with “broken” timing patterns (fixed‑size blocks, sleep‑based loops, etc.).

### Bypass Mode (On demand)

**Target:** Streams that don't need this tool, such as capture stream, etc..

**Method:** `wasapi_relink` no longer wraps client. Instead, it directly returns `IAudioClient3`.

## How to Use (Injection)

This library **does not** include its own DLL injector. You must use an external tool or [stub generator](https://github.com/namazso/dll-proxy-generator).

### Recommended Method: Special K

It is recommended to load `wasapi_relink.dll` as a **Special K** "plug-in" (lazy mode).

**Built-in Compatibility:** `wasapi_relink` actively detects and **skips** hooking Special K's own audio requests, ensuring SK's sound related features are not affected.

### Use as a Developer Library (Advanced Usage)

An additional purpose of this tool is to act as a "auxiliary" library if the audio library you are using does not support `IAudioClient3` low-latency initialization.

No complex injection tools are needed. Simply, early in your program’s startup, explicitly load this DLL with `LoadLibrary("wasapi_relink.dll")`. Its hooking behavior is automatically performed when the DLL is loaded, thereby "transparently" providing low-latency features to your existing audio library.

### Reminder on using stub generator

To facilitate use with stub generators, `wasapi_relink.dll` additionally exports a blank C function named `proxy`.

This function does nothing on its own; it exists only as a definitive "entry point".

**DO NOT use `ole32.dll` when generating DLL stubs.** Doing so can cause infinite recursion during `LoadLibrary`/COM activation and lead to unstable or hard-to-debug behavior (e.g., repeated loader calls, crashes, or the target process hanging).

## Configuration (`redirect_config.toml`)

The configuration file `redirect_config.toml` **must** be placed in the same directory(or working directory) as the DLL. If an entry is not specified, default value will be used.

```toml
# Path for the log file. "" (empty string) defaults to the working directory.
log_path = ""
# Log level: Trace, Debug, Info, Warn, Error, Never
log_level = "Info"
# Log only to stdout (true) or to both stdout and file (false).
only_log_stdout = false

[capture]
# (General) Target buffer duration in 0.1ms units (u16).
# The tool will calculate the closest buffer size *not exceeding* this duration while clamped within driver range.
# e.g., 20 = 2.0ms.
target_buffer_dur_ms = 20
# (General) Enable raw process for this stream (bool).
raw = true

[playback]

# Tool mode, available mode: Normal, Compat, Ringbuf, Bypass
mode = "Ringbuf"

# (Ringbuf mode exclusive, Optional) Assign a ring buffer length (in audio frames) to the corresponding samplerate.
# The number will be automatically rounded UP to align with the driver's fundamental period for optimal performance.
ring_buffer_len.48000 = 340 # about 7ms buffer

# (Compat mode exclusive, Optional) Assign a shared stream buffer duration (in 100-nanosecond units) to the corresponding samplerate.
# The number will be directly used as the inner shared buffer, and will be clamped by Windows if set too low.
compat_buffer_len.48000 = 0 # this will clamp to minimum allowed value
compat_buffer_len.96000 = 238350


```

### Config Details

- `log_path` (string): Where to save the log file. Non-directory value means current working directory.

- `log_level` (string): `Trace`, `Debug`, `Info`, `Warn`, `Error`, `Never`. Default is `Info`.  
**See performance warning below.**

- `only_log_stdout` (bool): Controls logging targets.
  - `true`: Logs *only* to the standard output (stdout). No log file will be created.
  - `false` (Default): Logs to **both** the standard output (stdout) and the file specified by `log_path`.
    - This option is particularly useful for developers who want to monitor logs in real-time in a terminal or for applications running in containerized environments (like Docker) where capturing stdout is the standard practice.

- `[playback]`/`[capture]`: Separate configs for output and input.

  - `mode` (string): `Normal`, `Compat`, `Ringbuf`, `Bypass`. Default is `Normal`.

  - `target_buffer_dur_ms` (u16): The target buffer size for all created low latency shared stream in **units of 0.1 milliseconds**. The tool will default to the driver's minimum if this is set too low or not specified. **You should generally not change this from the default value unless you experience audio pops.**

  - `raw` (bool): Indicates this stream to use **raw processing**, which bypasses most APO. Does nothing when mode is `Bypass`.

  - `ring_buffer_len.<samplerate>` (u32): The target buffer length for the ring buffer in **audio frames** (not samples). For example, 340 means 340 frames (680 samples in 2-channel audio). **It's recommended to set a proper value in Ringbuf mode.**
    - Note: The tool will automatically round this value *UP* to the nearest multiple of the driver’s fundamental period to ensure smooth streaming and prevent micro-glitches.

  - `compat_buffer_len.<samplerate>` (i64): The target buffer size for shared stream in **100-nanosecond units**. This controls the size of the shared buffer the program actually sees in Compat mode. The tool/Windows will default to the driver’s minimum if this is set too low or not specified. **This can help fix audio pops that occur after changing the audio sample rate in Compat mode.**

## Troubleshooting

Use this guide to diagnose and fix common audio issues.

### Audio is "Sliced" or in "Slow-Motion"

**Phenomenon:** Sound is heavily distorted, stretched, or sounds like it's being "sliced" and played back slowly.

**Cause:** This is the classic sign of a **Poll-driven** application running in **small buffer**. The app's polling logic is fighting the event-driven buffer. This is commonly seen in **Unity** games.

**Solution:** Try Compat or Ringbuf mode for the `[playback]` section.

### No sound or crashes with a log

**Phenomenon:** Completely silent, or simply crashes with `wasapi_relink` log provided.

**Cause:** This is mostly because of a **Fixed-size** application running in **small buffer**. The app's mixing logic is waiting indefinitely or encountering math errors on the buffer. This is usually seen in **Rhythm** games.

**Solution:** Try Ringbuf mode for the `[playback]` section.

### Good Audio with Occasional "Pops" or "Crackles"

**Phenomenon:** Audio playback is at the correct speed and pitch, but you hear intermittent pops, clicks, or small tearing sounds.

**Cause:** The buffer is **too small** for your system/software. Your CPU or the application cannot "feed" the audio driver new data fast enough, resulting in a buffer underrun.

**Solutions (Try either one):**

1. **Increase Buffer:** Slightly increase `target_buffer_dur_ms`. If your driver's minimum is 2ms (e.g., `20`), try `30` (3ms) or `40` (4ms) until the pops disappear.

2. **Try Compat or Ringbuf Mode:** This effectively adds a buffer layer for the application (i.e., the standard buffer from normal Shared mode). In combination with the changes made by those modes, this can potentially resolve the issue.

### Audio Pop after changing samplerate in Windows settings

**Phenomenon:** Audio playback is normal at samplerate A (e.g. 48000Hz), but pops at samplerate B (e.g. 96000 Hz) in compat mode.

**Cause:** In Compat mode, `compat_buffer_len` is mapped specifically to the active samplerate. If you switch the system samplerate but haven’t configured a buffer length for the new one, it falls back to a default value that might be too low for that specific rate.

**Solution:** Explicitly add the new samplerate to your config. For example:

```toml
[playback]
compat_buffer_len.48000 = 238350
compat_buffer_len.96000 = 250000  # Add a specific value for 96kHz
```

### Program won't start at all

**Phenomenon:** The target program fails to start after loading the DLL, and no log is given.

**Cause:** This is usually a **file permission issue** — the DLL attempts to write its log file to a location where it lacks write access.
It commonly occurs when the target executable resides under `Program Files` or other protected directories.

**Solution:** Specify a writable location for `log_path` in `redirect_config.toml`.  
For example:

```toml
log_path = "C:\\Users\\Public\\wasapi_relink.log"
```

If you don't need logging to file, you can also disable it with `only_log_stdout` :

```toml
only_log_stdout = true
```

As a last resort, you can entirely disable logging (not recommended):

```toml
log_level = "Never"
```

## Performance Warning: Logging & Audio Tearing

The Windows audio engine is *extremely* sensitive to timing.

- **DO NOT** set `log_level` to `Debug` or `Trace` during normal use.

- The high-frequency I/O (disk writing) from detailed logging **will** interfere with the audio thread, especially at low buffer sizes.

- This interference *can* cause **audio tearing, pops, and stuttering**.

- While asynchronous logging systems reduce the probability of popping sounds, your logs will be flooded with numerous function calls, consuming storage space.

- **ONLY** use `Debug`/`Trace` if you are actively debugging the application and you know what you are doing.

## Deep Dive: Technical Implementation

### The "Myth" of External Low-Latency Activators

Previous tools([REAL](https://github.com/miniant-git/REAL), [LowAudioLatency](https://github.com/spddl/LowAudioLatency) and similar projects) attempted to solve this by running a *separate* low-latency `IAudioClient3` application in the background.

- **What they got right:** As said the Windows FAQ, this *does* trigger the special Windows behavior and forces the *hardware* buffer to a smaller size for all apps.

- **What they missed (The "Myth"):** They assumed lowering the hardware buffer was enough. **It isn't.** A poll-driven application will *still* try to **prefill** its old 10ms buffer. The latency bottleneck just moves from the hardware to the application's own prefill.

 `wasapi_relink`'s **Compat Mode** solves *both* problems: it triggers the hardware change and internally hooks the application to *deceive and modify its prefill* to match the new, smaller hardware buffer.

### The Size Mismatch: Why Ringbuf Mode Exists

Even if you successfully force the audio engine into a low-latency period (e.g., 128 frames per cycle, roughly every 2.67ms at 48kHz), a fundamental scheduling mismatch remains between Windows and the application.

**The Engine’s Rhythm:**

In low-latency event-driven mode, Windows becomes an extremely fast, strict “pump”. It fires an event exactly every 128 frames (the requested period), and expects you to deliver exactly 128 frames of new data each time it wakes up. It has zero tolerance for variability.

**The Application’s Rhythm:**

Games, however, do not think in 128-frame chunks:

- Rhythm Games / Fixed-block mixers: They mix audio in large, rigid chunks (e.g., 256, 512, or 1024 frames) dictated by their internal logic frame.

- Poll-driven Engines (Unity, etc.): They operate in a tight while loop, calling `GetCurrentPadding` with fixed duration to see if *any* space is available, then dumping whatever they have.

**The Direct Consequence:**

If you let a game like these talk directly to a 128-frame low-latency engine buffer:

**For Poll-driven games:** If the buffer length is shorter than the poll intreval, the game will unable to send samples, causing severe underrun.

**For Fixed-block games:** If the buffer length is shorter than the mix chunk length, the audio engine may refuse to start.

**The `wasapi_relink` Ringbuf Solution:**

Ringbuf mode inserts a high-performance software decoupler between the audio engine and the application:

**The Engine Side:** `wasapi_relink` completely takes over the Windows Event callback. It wakes up exactly every 128 frames, quietly transfers the corresponding amount of data from the ring buffer to the hardware, and notifies the program to replenish the data via callback (if any). The hardware never sees the game’s messy behavior.

**The App Side:** The game is tricked into thinking it is talking to a much larger, traditional buffer. It writes its large fixed-frame chunks or polls at its leisure into this safe zone.

## Implementation: The COM Wrapper Chain

`wasapi_relink` does not use simple VTable hooking. It hooks at the root of the COM audio system to ensure 100% capture.

1. **Hook:** `CoCreateInstanceEx`.

2. **Intercept:** When the app requests `CLSID_MMDeviceEnumerator` (the main audio device enumerator), `wasapi_relink` creates the real one and returns a custom **wrapped** `IMMDeviceEnumerator` to the app.

3. **Propagate:** When the app requests device on this wrapper, `wasapi_relink` returns a wrapped `IMMDevice`.

    3.1. **Compat Mode:** When the app calls `Activate` on the wrapped `IMMDevice`, `wasapi_relink` executes its core Compat Mode logic (creating the two clients) and returns *one* wrapped `IAudioClient`.

    3.2. **Normal Mode:** When the app calls `Initialize` on the wrapped `IAudioClient`, the Normal Mode logic is executed.

    3.3. **Ringbuf Mode:** When the app calls `Initialize`, it will enable event callback, setup low-latency stream, choose modes based on program behavior, and creates consumer thread.

4. This wrapper chain continues all the way down to `IAudioRenderClient`, giving `wasapi_relink` full, transparent control over the entire audio stream lifecycle.

### A Note on Capture (Microphone)

In **Compat/Ringbuf Mode**, capture (mic) streams are simply forwarded. This *is* intentional:

1. Mic latency is far less critical in Shared mode compared to effects or network latency. Due to the nature of capture streams, which don’t require a silent pre-fill buffer in their interaction with the driver, triggering Windows’ special behavior is by itself sufficient to reduce latency.
2. Users with true low-latency input needs are already using `ASIO` or `WASAPI Exclusive` mode, which this tool does not target.

---

### **Disclaimer:**

This is an **experimental** and **educational** project.  

Keep in mind:

- Hooking may trigger anti-cheat detections in some games.
- WASAPI internals and buffer behavior can differ between Windows builds or drivers.
- Use it responsibly — at your own risk.
