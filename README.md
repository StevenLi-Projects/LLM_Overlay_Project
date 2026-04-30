# Local Text Formatting Assistant for llama.cpp

This is a local-first Windows helper for formatting selected editable text with a local `llama.cpp` model. It works in standard Windows Notepad and many other Windows apps that support normal `Ctrl+C` and `Ctrl+V` editing. It copies the current selection, sends it to a locally running `llama-server`, previews the proposed replacement, pastes the approved Markdown back over the selection, then restores the previous clipboard contents.

No cloud AI APIs, Microsoft AI credits, Copilot, OpenAI API credits, or paid remote services are used.

## Files

- `NotepadMarkdownAssistant.ps1` - tray app, global hotkeys, and popup mode menu.
- `Launch-Assistant.vbs` - double-click launcher that starts the tray app with no visible PowerShell window.
- `Install-AssistantShortcuts.ps1` - optional helper to create Desktop or Startup shortcuts to the hidden launcher.
- `Uninstall-AssistantShortcuts.ps1` - optional helper to remove those shortcuts.
- `Start-LlamaServer.ps1` - starts the local `llama-server.exe` with the configured GGUF model.
- `Test-Assistant.ps1` - validates paths and optionally checks the local server.
- `Unblock-LlamaCpp.ps1` - optional helper for removing Windows download-blocking marks from the local llama.cpp install.
- `config.json` - active settings.
- `config.example.json` - reset/reference settings.
- `models/` - local-only folder for GGUF model files. It is ignored by Git.
- `llama-b8987-bin-win-cuda-12.4-x64/` - local-only llama.cpp Windows install folder. It is ignored by Git.

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

1. Download a llama.cpp Windows release from [ggml-org/llama.cpp releases](https://github.com/ggml-org/llama.cpp/releases). This config expects the CUDA 12.4 Windows x64 build folder named `llama-b8987-bin-win-cuda-12.4-x64/`; either extract that build into this project folder or update `llama.cpp_dir` in `config.json`.
2. Create `models/` in this project folder.
3. Download the normal profile model from [unsloth/gemma-4-E4B-it-GGUF](https://huggingface.co/unsloth/gemma-4-E4B-it-GGUF/blob/main/gemma-4-E4B-it-Q5_K_M.gguf) and save it as `models/gemma-4-E4B-it-Q5_K_M.gguf`.
4. Download the fast profile model from [ruygar/gemma-4-E2B-it-GGUF](https://huggingface.co/ruygar/gemma-4-E2B-it-GGUF/blob/main/gemma-4-E2B-it-Q4_K_M.gguf) and save it as `models/gemma-4-E2B-it-Q4_K_M.gguf`.

Validate the local files:

```powershell
.\Test-Assistant.ps1
```

If validation warns that `llama-server.exe` has a `Zone.Identifier` download mark, unblock the local llama.cpp install once:

```powershell
.\Unblock-LlamaCpp.ps1
```

## Run

Option A, tray-only: double-click:

```text
Launch-Assistant.vbs
```

This starts the assistant hidden in the background. No PowerShell window stays open. The assistant appears only as a tray icon.

Option B, create a normal Windows shortcut:

```powershell
.\Install-AssistantShortcuts.ps1 -Desktop
```

Then launch it from the `Local Text Formatting Assistant` shortcut on your Desktop.

To also start the assistant automatically when you sign in:

```powershell
.\Install-AssistantShortcuts.ps1 -Startup
```

Option C, diagnostic mode: run with a visible PowerShell console so you can see timing logs and errors:

```powershell
.\NotepadMarkdownAssistant.ps1
```

Option D, more visible server debugging: start the server yourself in one PowerShell window:

```powershell
.\Start-LlamaServer.ps1
```

Then start the assistant in another PowerShell window:

```powershell
.\NotepadMarkdownAssistant.ps1
```

The assistant appears as a small tray icon. Right-click it to switch between `Normal (E4B quality)` and `Fast (E2B)`, start/check the server, or exit. Auto-started `llama-server` runs hidden; tray notifications are off by default.

## Usage

1. Open Notepad or another app where selected text can be edited.
2. Select text.
3. Press the popup menu hotkey, or press a direct mode hotkey.

Popup menu:

| Action | Hotkey |
| --- | --- |
| Show formatting menu at cursor | `Ctrl+Alt+Space` |

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

After the model returns text, a compact preview window opens with only the editable replacement plus local telemetry such as endpoint, generation time, output tokens, and TPS when llama.cpp reports token counts. Choose `Replace` to paste it over the selected text or `Cancel` to leave the target app unchanged.

For apps other than Notepad, the target field must support ordinary keyboard copy and paste. This includes many editors, browsers, email clients, chat boxes, and document fields. It will not replace text in read-only views, protected admin windows, password fields, or apps that block simulated keyboard input.

## Configuration

Edit `config.json`.

Important settings:

```json
{
  "llama": {
    "cpp_dir": "llama-b8987-bin-win-cuda-12.4-x64",
    "model_path": "models/gemma-4-E4B-it-Q5_K_M.gguf",
    "active_profile": "normal",
    "profiles": {
      "normal": {
        "label": "Normal (E4B quality)",
        "model_path": "models/gemma-4-E4B-it-Q5_K_M.gguf",
        "model_name": "local-gemma-e4b",
        "context_size": 8192
      },
      "fast": {
        "label": "Fast (E2B)",
        "model_path": "models/gemma-4-E2B-it-Q4_K_M.gguf",
        "model_name": "local-gemma-e2b",
        "context_size": 8192
      }
    },
    "port": 8080,
    "server_url": "http://127.0.0.1:8080",
    "context_size": 8192,
    "gpu_layers": 999,
    "auto_start_server": true,
    "health_cache_sec": 30,
    "server_args": [
      "--flash-attn",
      "auto",
      "--cache-prompt",
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
    "preview_enabled": true,
    "show_notifications": false,
    "show_timing_notifications": false,
    "menu_hotkey": "Ctrl+Alt+Space",
    "copy_wait_ms": 180,
    "paste_wait_ms": 220
  }
}
```

To use a different model, put the `.gguf` file in `models/` or another local folder, then change the relevant `llama.profiles.<profile>.model_path`. `llama.model_path` remains as a compatibility fallback for older configs.

To change the default startup profile, edit:

```json
"active_profile": "fast"
```

To change the server port, change both `port` and `server_url`, for example:

```json
"port": 8081,
"server_url": "http://127.0.0.1:8081"
```

To reduce GPU memory use, lower `gpu_layers`. To force CPU-only behavior, set it to `0`.

For larger inputs and outputs, the default context is `8192`, the global generation cap is `2048`, and each mode has its own `max_tokens` cap. Lower these values if latency or VRAM use becomes uncomfortable.

`generation.prefer_completion` uses llama.cpp's fast `/completion` endpoint first and keeps `/v1/chat/completions` as fallback.

`llama.health_cache_sec` skips repeated server health probes after a recent successful check. `llama.server_args` are appended to the `llama-server.exe` command for both manual and auto-started servers.

To turn off the preview dialog and replace immediately, set `ui.preview_enabled` to `false`.

Notifications are off by default so the assistant stays quiet in the system tray. To enable tray balloons, set `ui.show_notifications` to `true`. To show timing details in completion balloons, also set `ui.show_timing_notifications` to `true`.

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

If Windows cancels the launch or shows a download/security prompt for `llama-server.exe`, run:

```powershell
.\Unblock-LlamaCpp.ps1
```

### Model unavailable

Check that `model_path` and every `llama.profiles.*.model_path` point to existing `.gguf` files. Run:

```powershell
.\Test-Assistant.ps1
```

### Request timeout

Increase `generation.timeout_sec`, reduce `generation.max_tokens`, or start with a smaller model/context size.

### Latency tuning

The assistant prints per-run timings to the PowerShell console, including server check, copy, generation, preview, paste, total milliseconds, endpoint, token counts, and TPS when available. Generation is usually the largest part. For faster responses, keep `llama-server` running, lower mode-specific `max_tokens`, or disable preview if you do not need confirmation.

### Invalid or empty response

Make sure your `llama-server.exe` supports either `/v1/chat/completions` or `/completion`. This assistant tries both.

### Hotkey conflict or does nothing

Another app may already own the hotkey, or another copy of the assistant may still be running. Check the tray and exit any old assistant instance. If the conflict remains, change the hotkey in `config.json`, restart the assistant, and try again.

If you accidentally launch the assistant twice, the second copy shows an "already running" message and exits.

The assistant continues running when only some hotkeys fail. Startup warnings list which shortcuts could not be registered.

### Hidden launcher started but no tray icon appears

Run diagnostic mode once:

```powershell
.\NotepadMarkdownAssistant.ps1
```

The console will show startup errors such as a blocked script, missing model, missing `llama-server.exe`, or a hotkey conflict.

### Clipboard behavior

The assistant snapshots the clipboard before copying the selection and restores it after pasting. Most normal clipboard contents are preserved. Some apps with unusual delayed-render clipboard data may not restore perfectly, which is the main tradeoff of supporting standard Windows apps without app-specific plugins.

### Preview window

The preview window appears after generation and before paste. It only shows the replacement text, not the original selection. `Replace` pastes the current preview text, including any manual edits. `Cancel` restores the clipboard snapshot and does not paste anything into the target app. The status line shows local telemetry such as generation time, output tokens, prompt tokens, endpoint, and TPS when available.

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
