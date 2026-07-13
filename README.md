# Local Text Formatting Assistant for llama.cpp

This is a local-first Windows helper for formatting selected editable text with a local `llama.cpp` model. It works in standard Windows Notepad and many other Windows apps that support normal `Ctrl+C` and `Ctrl+V` editing. It copies the current selection, sends it to a locally running `llama-server`, previews the proposed replacement, pastes the approved Markdown back over the selection, then restores the previous clipboard contents.

No cloud AI APIs, Microsoft AI credits, Copilot, OpenAI API credits, or paid remote services are used.

## Files

- `dist/LocalTextFormattingAssistant.exe` - primary compiled tray application; the adjacent `.dll` and runtime metadata files are part of the lightweight .NET 8 build.
- `src/LocalTextFormattingAssistant/` - dependency-free .NET 8 WinForms source.
- `Build-App.ps1` - rebuilds and publishes the compiled app when a .NET 8 SDK is available.
- `NotepadMarkdownAssistant.ps1` - legacy PowerShell implementation kept as a fallback.
- `Launch-Assistant.vbs` - compatibility launcher; prefers the compiled app and falls back to PowerShell.
- `Install-AssistantShortcuts.ps1` - creates Desktop or Startup shortcuts directly to the compiled app.
- `Uninstall-AssistantShortcuts.ps1` - optional helper to remove those shortcuts.
- `Start-LlamaServer.ps1` - starts the local `llama-server.exe` with the configured GGUF model.
- `Test-Assistant.ps1` - validates paths and optionally checks the local server.
- `Diagnose-LlamaGpu.ps1` - checks whether NVIDIA/CUDA is visible to llama.cpp.
- `Unblock-LlamaCpp.ps1` - optional helper for removing Windows download-blocking marks from the local llama.cpp install.
- `config.json` - active settings.
- `config.example.json` - reset/reference settings.
- `models/` - local-only folder for GGUF model files. It is ignored by Git.
- `llama-b9977-bin-win-cuda-12.4-x64/` - local-only llama.cpp Windows install folder. It is ignored by Git.

## Setup

Open PowerShell in this folder:

```powershell
cd C:\Users\shli8\Documents\LLM_Overlay_Project
```

If PowerShell blocks local scripts, allow scripts for only this process:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Install local requirements before running the assistant:

1. Download the official [llama.cpp b9977 Windows CUDA release](https://github.com/ggml-org/llama.cpp/releases/tag/b9977). This config expects the CUDA 12.4 Windows x64 build folder named `llama-b9977-bin-win-cuda-12.4-x64/`; either extract that build into this project folder or update `llama.cpp_dir` in `config.json`.
2. Create `models/` in this project folder.
3. Download the normal profile model from [unsloth/gemma-4-E4B-it-GGUF](https://huggingface.co/unsloth/gemma-4-E4B-it-GGUF/blob/main/gemma-4-E4B-it-Q5_K_M.gguf) and save it as `models/gemma-4-E4B-it-Q5_K_M.gguf`.
4. Download the fast profile model from [ruygar/gemma-4-E2B-it-GGUF](https://huggingface.co/ruygar/gemma-4-E2B-it-GGUF/blob/main/gemma-4-E2B-it-Q4_K_M.gguf) and save it as `models/gemma-4-E2B-it-Q4_K_M.gguf`.
5. For speculative decoding, download the current-format [`E4B Q4_K_M` assistant head](https://huggingface.co/cascade-tech/gemma-4-E4B-it-qat-q4_0-unquantized-assistant-gguf) and [`E2B Q4_0` assistant head](https://huggingface.co/unsloth/gemma-4-E2B-it-qat-GGUF/tree/main/MTP). Save them as `models/gemma-4-E4B-it-assistant-Q4_K_M-current.gguf` and `models/gemma-4-E2B-it-assistant-Q4_0-current.gguf` respectively.

The published app needs the [Microsoft .NET 8 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0). Install it if Windows does not already provide it, then validate the local files and compiled app:

```powershell
.\Test-Assistant.ps1
```

If validation warns that `llama-server.exe` has a `Zone.Identifier` download mark, unblock the local llama.cpp install once:

```powershell
.\Unblock-LlamaCpp.ps1
```

## Run

Option A, primary compiled app: double-click:

```text
dist\LocalTextFormattingAssistant.exe
```

The app starts directly in the system tray with no PowerShell process. It finds `config.json` in the parent project folder.

Option B, create a normal Windows shortcut:

```powershell
.\Install-AssistantShortcuts.ps1 -Desktop
```

Then launch it from the `Local Text Formatting Assistant` shortcut on your Desktop.

To also start the assistant automatically when you sign in:

```powershell
.\Install-AssistantShortcuts.ps1 -Startup
```

Option C, validate the compiled app without opening the tray UI:

```powershell
.\Test-Assistant.ps1
```

Option D, legacy diagnostic mode with a visible PowerShell console:

```powershell
.\NotepadMarkdownAssistant.ps1
```

Option E, more visible server debugging: start the server yourself in one PowerShell window:

```powershell
.\Start-LlamaServer.ps1
```

Then start the assistant in another PowerShell window:

```powershell
.\NotepadMarkdownAssistant.ps1
```

The assistant appears as a small tray icon. Double-click it or press `Ctrl+Alt+Space` for the command palette. Right-click it for model, preview, server, settings, diagnostics, and exit commands. Auto-started `llama-server` runs hidden; tray notifications remain off.

## Usage

1. Open Notepad or another app where selected text can be edited.
2. Select text.
3. Press the command-palette hotkey, or press a direct mode hotkey.

Popup menu:

| Action | Hotkey |
| --- | --- |
| Open keyboard-navigable formatter at cursor | `Ctrl+Alt+Space` |

Model profile:

| Profile | Model | Best for |
| --- | --- | --- |
| `Normal (E4B quality)` | [`models/gemma-4-E4B-it-Q5_K_M.gguf`](https://huggingface.co/unsloth/gemma-4-E4B-it-GGUF/blob/main/gemma-4-E4B-it-Q5_K_M.gguf) | Better formatting quality |
| `Fast (E2B)` | [`models/gemma-4-E2B-it-Q4_K_M.gguf`](https://huggingface.co/ruygar/gemma-4-E2B-it-GGUF/blob/main/gemma-4-E2B-it-Q4_K_M.gguf) | Lower latency |

Right-click the tray icon to switch profiles. If a different configured model is already running, the assistant stops that local `llama-server` process and starts the selected profile on the next request.

Direct mode hotkeys:

| Mode | Hotkey |
| --- | --- |
| Format as Markdown | `Ctrl+Alt+W` in the current `config.json` |
| Convert to bullet points | `Ctrl+Alt+B` |
| Convert to table when appropriate | `Ctrl+Alt+T` |
| Clean up without changing meaning | `Ctrl+Alt+C` |
| Summarize into concise Markdown | `Ctrl+Alt+S` |

Direct hotkeys show a compact cancellable progress window while copying, starting the model, and generating. After generation, the output-only preview shows the editable replacement plus local telemetry. `Replace` pastes it, `Copy` temporarily places it on the clipboard while the preview stays open, and `Cancel` leaves the target unchanged. Use `Ctrl+Enter` to replace or Escape to cancel. Disable preview from the palette, tray, or Settings for immediate replacement.

The command palette follows the Windows light/dark preference, remembers the last mode during the session, and exposes model and preview controls without opening Settings.

For apps other than Notepad, the target field must support ordinary keyboard copy and paste. This includes many editors, browsers, email clients, chat boxes, and document fields. It will not replace text in read-only views, protected admin windows, password fields, or apps that block simulated keyboard input.

## Configuration

Edit `config.json`.

Important settings:

```json
{
  "llama": {
    "cpp_dir": "llama-b9977-bin-win-cuda-12.4-x64",
    "model_path": "models/gemma-4-E4B-it-Q5_K_M.gguf",
    "active_profile": "normal",
    "profiles": {
      "normal": {
        "label": "Normal (E4B quality)",
        "model_path": "models/gemma-4-E4B-it-Q5_K_M.gguf",
        "model_name": "local-gemma-e4b",
        "context_size": 8192,
        "speculative": {
          "enabled": true,
          "type": "gemma4_mtp",
          "fallback_without_support": true,
          "draft_model_path": "models/gemma-4-E4B-it-assistant-Q4_K_M-current.gguf",
          "draft_gpu_layers": 999,
          "draft_context_size": 8192,
          "draft_block_size": 2,
          "draft_n_max": 2,
          "draft_n_min": 1,
          "draft_p_min": 0.75
        }
      },
      "fast": {
        "label": "Fast (E2B)",
        "model_path": "models/gemma-4-E2B-it-Q4_K_M.gguf",
        "model_name": "local-gemma-e2b",
        "context_size": 8192,
        "speculative": {
          "enabled": true,
          "type": "gemma4_mtp",
          "fallback_without_support": true,
          "draft_model_path": "models/gemma-4-E2B-it-assistant-Q4_0-current.gguf",
          "draft_gpu_layers": 999,
          "draft_context_size": 8192,
          "draft_block_size": 2,
          "draft_n_max": 2,
          "draft_n_min": 1,
          "draft_p_min": 0.75
        }
      }
    },
    "port": 8080,
    "server_url": "http://127.0.0.1:8080",
    "context_size": 8192,
    "gpu_layers": 999,
    "prefer_gpu": true,
    "require_gpu": false,
    "gpu_device": "CUDA0",
    "auto_start_server": true,
    "prewarm_on_launch": false,
    "health_cache_sec": 30,
    "server_args": [
      "--flash-attn",
      "auto",
      "--cache-prompt",
      "--threads",
      "8",
      "--threads-batch",
      "8",
      "--parallel",
      "1"
    ]
  },
  "generation": {
    "prefer_completion": true,
    "temperature": 0.2,
    "top_p": 0.9,
    "max_tokens": 2048,
    "timeout_sec": 180
  },
  "ui": {
    "theme": "system",
    "preview_enabled": true,
    "progress_overlay_enabled": true,
    "show_notifications": false,
    "show_timing_notifications": false,
    "menu_hotkey": "Ctrl+Alt+Space",
    "copy_wait_ms": 180,
    "paste_wait_ms": 220
  }
}
```

To use a different model, put the `.gguf` file in `models/` or another local folder, then change the relevant `llama.profiles.<profile>.model_path`. `llama.model_path` remains as a compatibility fallback for older configs.

[Gemma 4 multi-token prediction](https://blog.google/innovation-and-ai/technology/developers-tools/multi-token-prediction-gemma-4/) is configured per profile under `llama.profiles.<profile>.speculative`. The launcher auto-detects current llama.cpp's [`draft-mtp` interface](https://github.com/ggml-org/llama.cpp/pull/23398) and older Atomic-style runtimes (`--mtp-head ... --spec-type mtp`). Keep E4B paired with its E4B assistant and E2B paired with its E2B assistant. To disable MTP for a profile, set `speculative.enabled` to `false`. With `fallback_without_support: true`, a runtime or head that fails during startup is retried with standard decoding.

To change the default startup profile, edit:

```json
"active_profile": "fast"
```

To change the server port, change both `port` and `server_url`, for example:

```json
"port": 8081,
"server_url": "http://127.0.0.1:8081"
```

To reduce GPU memory use, lower `gpu_layers`. To force CPU-only behavior, set `prefer_gpu` to `false` or set `gpu_layers` to `0`.

`prefer_gpu` is enabled by default. The launcher checks `llama-server --list-devices`: when CUDA is visible it starts with `--device CUDA0` and the configured `gpu_layers`; when CUDA is not visible it starts with `--n-gpu-layers 0` as a CPU fallback. Set `require_gpu` to `true` only if you want the assistant to fail instead of falling back to CPU.

For larger inputs and outputs, the default context is `8192`, the global generation cap is `2048`, and each mode has its own `max_tokens` cap. Lower these values if latency or VRAM use becomes uncomfortable.

`generation.prefer_completion` uses llama.cpp's fast `/completion` endpoint first and keeps `/v1/chat/completions` as fallback.

`llama.health_cache_sec` skips repeated server health probes after a recent successful check. `llama.server_args` are appended to the `llama-server.exe` command for both manual and auto-started servers. The default `--threads 8` and `--threads-batch 8` keep llama.cpp's CPU-side scheduling pool moderate while the model layers run on CUDA.

Manual MTP startup example:

```powershell
.\llama-b9977-bin-win-cuda-12.4-x64\llama-server.exe `
  --model .\models\gemma-4-E4B-it-Q5_K_M.gguf `
  --spec-draft-model .\models\gemma-4-E4B-it-assistant-Q4_K_M-current.gguf `
  --spec-type draft-mtp `
  --spec-draft-n-max 2 `
  --spec-draft-n-min 1 `
  --spec-draft-p-min 0.75 `
  --alias local-gemma-e4b `
  --host 127.0.0.1 `
  --port 8080 `
  --ctx-size 8192 `
  --n-gpu-layers 999 `
  --n-gpu-layers-draft 999 `
  --device CUDA0 `
  --spec-draft-device CUDA0 `
  --flash-attn auto `
  --cache-prompt `
  --parallel 1
```

Most settings can be changed from the tray menu's `Settings` window. It validates model paths, llama.cpp, ports, context size, generation limits, and duplicate or malformed hotkeys before saving.

To turn off the preview dialog and replace immediately, uncheck `Preview before replacing` in the palette, tray menu, or Settings, or set `ui.preview_enabled` to `false`.

`ui.theme` accepts `system`, `light`, or `dark`. `ui.progress_overlay_enabled` controls the cancellable working window. `llama.prewarm_on_launch` remains `false`, so the model starts only when first needed.

The compiled app does not use tray balloons. Errors and server state appear in the palette, progress window, or Settings. The legacy PowerShell app still honors its notification settings.

To change the popup menu hotkey, edit `ui.menu_hotkey`.

To change a direct mode hotkey, edit the relevant value under `modes`.

## Troubleshooting

### No selected text

Select editable text before pressing the hotkey. The app uses the target app's normal copy/paste behavior, so the selection must be active and the app must allow `Ctrl+C` and `Ctrl+V`.

### llama.cpp server not running

Run:

```powershell
.\Start-LlamaServer.ps1
```

Or set `"auto_start_server": true` in `config.json`.

### GPU not used / CPU-only fallback

Run:

```powershell
.\llama-b9977-bin-win-cuda-12.4-x64\llama-server.exe --list-devices
```

You should see a CUDA device. If it only shows CPU/RPC backends, llama.cpp cannot load its CUDA backend yet. `nvidia-smi` can still see your GPU while llama.cpp cannot use it if CUDA runtime/cuBLAS DLLs are missing from your PATH or from the llama.cpp folder.

This project sets:

```json
"gpu_layers": 999,
"prefer_gpu": true,
"require_gpu": false,
"gpu_device": "CUDA0"
```

With the default settings, the assistant uses GPU first and falls back to CPU if llama.cpp cannot see CUDA. If you want to refuse CPU fallback, set `require_gpu` to `true`. Install the missing CUDA runtime/cuBLAS DLLs or replace the llama.cpp folder with a CUDA Windows build that includes the required runtime DLLs, then re-run `.\Test-Assistant.ps1`.

For a fuller local diagnosis, run:

```powershell
.\Diagnose-LlamaGpu.ps1
```

The common missing CUDA 12 DLLs are `cudart64_12.dll`, `cublas64_12.dll`, and `cublasLt64_12.dll`. Install them from NVIDIA's official CUDA Toolkit, then make sure they are either on PATH or copied into the same folder as `llama-server.exe`.

If Windows cancels the launch or shows a download/security prompt for `llama-server.exe`, run:

```powershell
.\Unblock-LlamaCpp.ps1
```

### Model unavailable

Check that `model_path`, every `llama.profiles.*.model_path`, and every enabled `llama.profiles.*.speculative.draft_model_path` point to existing `.gguf` files. Run:

```powershell
.\Test-Assistant.ps1
```

### Request timeout

Increase `generation.timeout_sec`, reduce `generation.max_tokens`, or start with a smaller model/context size.

### Latency tuning

The compiled preview reports endpoint, generation time, token count, prompt TPS, decode TPS, MTP state, and the accepted/generated draft-token ratio. Use `Copy diagnostics` from the tray for the bounded in-memory server/app log. For faster responses, keep `llama-server` running, lower mode-specific `max_tokens`, or disable preview if confirmation is unnecessary.

If decode TPS is around CPU speeds, fully exit the tray assistant and relaunch it. The launcher validates the server alias/model and restarts the configured port so it does not silently reuse an old CPU-only server.

With MTP enabled, speedup depends on draft acceptance. The tested consistency-first default is `draft_n_max: 2`; increase it to `3` or `4` for experimentation, or disable MTP for a workload with poor acceptance.

The pre-merge F16 heads use obsolete architecture metadata and do not load in current llama.cpp. Use the `*-current.gguf` heads and b9977 or newer. `Test-Assistant.ps1` reports the detected MTP runtime style.

### Invalid or empty response

Make sure your `llama-server.exe` supports either `/v1/chat/completions` or `/completion`. This assistant tries both.

### Hotkey conflict or does nothing

Another app may already own the hotkey, or another copy of the assistant may still be running. Check the tray and exit any old assistant instance. If the conflict remains, change the hotkey in `config.json`, restart the assistant, and try again.

If you accidentally launch the assistant twice, the second copy shows an "already running" message and exits.

The assistant continues running when only some hotkeys fail. Startup warnings list which shortcuts could not be registered.

### Compiled app starts but no tray icon appears

Run compiled validation first:

```powershell
.\Test-Assistant.ps1
```

If validation passes, run legacy diagnostic mode once for comparison:

```powershell
.\NotepadMarkdownAssistant.ps1
```

The console will show startup errors such as a blocked script, missing model, missing `llama-server.exe`, or a hotkey conflict.

### Clipboard behavior

The assistant snapshots the clipboard before copying the selection and restores it after pasting. Most normal clipboard contents are preserved. Some apps with unusual delayed-render clipboard data may not restore perfectly, which is the main tradeoff of supporting standard Windows apps without app-specific plugins.

### Preview window

The preview window appears after generation and before paste. It only shows the replacement text, not the original selection. `Replace` pastes the current preview text, including any manual edits. `Cancel` restores the clipboard snapshot and does not paste anything into the target app. The status line shows local telemetry such as generation time, output tokens, prompt tokens, endpoint, decode TPS, wall TPS, and prompt TPS when available.

## Uninstall

1. Right-click the tray icon and choose `Exit`.
2. Close the `llama-server` PowerShell window if you started it manually.
3. If you created shortcuts, run:

```powershell
.\Uninstall-AssistantShortcuts.ps1
```

4. Delete this project folder, or delete only the assistant files listed above.

## Chosen approach

Standard Windows Notepad does not expose a practical plugin API for this workflow, and most Windows apps do not share one common editing API. The most reliable simple approach is therefore a small Windows hotkey helper that uses normal copy/paste automation and a local HTTP call to `llama.cpp`. This keeps dependencies minimal, works with ordinary Notepad, also works with many other editable text fields, stays fully local, and avoids installing a keyboard macro framework or editor-specific plugin.

## Primary references

- [Google: Multi-token prediction in Gemma 4](https://blog.google/innovation-and-ai/technology/developers-tools/multi-token-prediction-gemma-4/)
- [llama.cpp PR #23398: Gemma 4 MTP implementation](https://github.com/ggml-org/llama.cpp/pull/23398)
- [llama.cpp b9977 Windows CUDA release](https://github.com/ggml-org/llama.cpp/releases/tag/b9977)
- [Cascade Tech: current-format Gemma 4 E4B assistant GGUFs](https://huggingface.co/cascade-tech/gemma-4-E4B-it-qat-q4_0-unquantized-assistant-gguf)
- [Unsloth: current-format Gemma 4 E2B MTP GGUFs](https://huggingface.co/unsloth/gemma-4-E2B-it-qat-GGUF/tree/main/MTP)
