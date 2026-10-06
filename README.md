**English** | [简体中文](./README.zh-CN.md)

<div align="center">

<img src="public/icons/icon.svg" width="112" alt="Slilot logo"/>

# Slilot

**A self-hosted PowerPoint AI Agent add-in — chat, and it edits the presentation you have open**

![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20PowerPoint%20desktop-blue)
![Node](https://img.shields.io/badge/Node.js-%E2%89%A5%2018-green)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-pilot--release-orange)
![PRs](https://img.shields.io/badge/PRs-welcome-orange)

Slilot is a **PowerPoint AI Agent add-in** (pilot release): in a sidebar chat, the model
directly builds and edits your currently open presentation — slides, text, tables, and
generated illustrations — then reviews its own work from real per-slide screenshots.
It runs entirely on your machine against the LLM provider of your choice; your
documents and API keys never leave it.

<!-- TODO: record a 10-second demo GIF (chat → live edit) into docs/demo.gif, then uncomment
![Demo](docs/demo.gif)
-->

</div>

## Features

- **Edits the live document**: the model drives Office.js / PowerPoint COM to operate on the open presentation — changes land in the real file, fully undoable and savable in PowerPoint
- **11 fine-grained tools**: read overview / read slide / edit text / style text / add textbox / add slides / delete shape / generate image / insert image / screenshot / per-slide visual review
- **Real screenshot review (RIP loop)**: after each slide, the add-in renders it to an image so the model can actually *look* at the result — checking overlap, overflow, distortion, style consistency — and fix issues on the spot; a deterministic numeric audit (text truncation/overflow, out-of-canvas) must be cleared before moving on
- **No distorted images**: the model specifies a placement box; the add-in scales the picture to its true aspect ratio and centers it inside the box
- **Bilingual UI**: Chinese / English toggle in settings (Chinese by default)
- **Bring your own model**: the panel speaks the Anthropic tool protocol; a local translation layer bridges any provider offering `/v1/messages` or `/v1/responses` — no vendor lock-in
- **Long tasks survive**: context over a threshold is auto-summarized into a handoff note (key data preserved), screenshots pruned to the latest one, with a hard stop as backstop

## Prerequisites

| Item | Requirement |
|---|---|
| OS | Windows 10 / 11 (sideloading, COM rendering, and the autostart script are Windows-only) |
| Office | **Desktop PowerPoint** (Office 2021 / Microsoft 365 verified). PowerPoint for the web and Mac are not supported |
| Node.js | ≥ 18 ([nodejs.org](https://nodejs.org) LTS; check with `node -v`). Zero dependencies — **no npm install needed** |
| Model service | Your own LLM API key (image generation requires a key with that capability enabled) |
| Admin rights | Not required (except for the optional shared-catalog fallback) |

## Install

### Option 1 (recommended): let an AI agent install it

Clone this repository, then tell the AI agent on your computer:

> Install this add-in for my PowerPoint: https://github.com/x1han/slilot

The agent follows the README and scripts to:

1. Check Node ≥ 18 and desktop PowerPoint;
2. Run `npx office-addin-dev-certs install --days 3650` and trust the localhost dev certificate (a system dialog appears — choose trust). `--days` matters: the default validity is only 30 days;
3. Run `install-sideload.ps1` (writes the HKCU developer registry, no admin needed);
4. Start `node server.js` and verify `https://localhost:3010/healthz` returns ok;
5. Optional: create a shortcut to `autostart-hidden.vbs` and drop the shortcut into the Startup folder (Win+R → `shell:startup`) for autostart.

### Option 2: manual install

```powershell
# 1. Trust the dev certificate (must precede the server, or server.js exits immediately;
#    default validity is 30 days — --days 3650 avoids monthly expiry)
npx office-addin-dev-certs install --days 3650

# 2. Start the local service (or double-click start-addin.bat)
node server.js

# 3. Sideload (no admin), then restart PowerPoint
powershell -ExecutionPolicy Bypass -File install-sideload.ps1

# 4. (Optional) autostart: create a shortcut to autostart-hidden.vbs and move the
#    shortcut into the Startup folder (do not move the vbs itself — it locates the
#    repo relative to its own path)
```

After restarting PowerPoint, a **Slilot** button appears at the right end of the Home tab.

## First run

Open ⚙ in the bottom-right of the pane and fill in your model service:

| Protocol | Full chat URL |
|---|---|
| Anthropic — /v1/messages | `<base>/v1/messages` |
| OpenAI — /v1/responses | `<base>/v1/responses` |

(`<base>` = the upstream base you fill in above.)

- The dialog has two blocks. **Text & vision model** (used for chat and screenshot review): upstream base / protocol / API key / text & vision model id. **Image model** (used for illustrations): image base / image API key — leave both empty to follow the text block's — plus a required image model id, separate from the text one (e.g. MiniMax needs `image-01`).
- The pane UI itself is bilingual: toggle with the button next to "Settings" (Chinese by default).
- Every input shows a grey example placeholder (api.example.com / your-model-id / your-image-model); nothing is preconfigured — fill your provider before testing.
- **Test** runs three checks in parallel — **chat / vision / image** — each line shows a spinner, then a green dot (pass) or red dot (fail) as it completes; all three must pass before **Save** unlocks (any field change requires re-testing). The vision check draws a random shape/color/position and compares the model's description against ground truth, so it cannot be passed by guessing.
- The image endpoint is auto-probed: `<base>/v1/images/generations` (OpenAI standard) first, then `<base>/v1/image_generation` (MiniMax-style); the working one is remembered. The image base may also be a full endpoint URL — it is used as-is. Slow image models are supported (5-minute test window, 10-minute generation window).

Then just describe what you want, for example:

> Read this deck and add a closing summary slide in the same style
> Generate a 16:9 cover image for slide 1, minimal style
> Turn the three bullets on slide 3 into a 2×3 grid with an illustration

## How it works

```
PowerPoint task pane (public/taskpane.html/js)
    │  same-origin requests to https://localhost:3010/api/forward (relative URL)
    ▼
Local service server.js (Node ≥ 18, zero dependencies, HTTPS + dev cert)
    │  generic forwarding + protocol translation (Anthropic ↔ OpenAI, both ways)
    ▼
Any https upstream
    ├─ Anthropic /v1/messages        passthrough
    └─ OpenAI   /v1/responses        request/response translation
```

- **Tool loop**: the pane runs an Anthropic-protocol agent loop, decoupled from the upstream format; tools execute through two channels — Office.js (text, slides, reading shapes) and COM (precise image placement, full-deck screenshots).
- **Screenshot review**: `scripts/export-slides.ps1` exports the open deck slide-by-slide to JPG; the model reviews the real rendered result and fixes problems (layout placeholders that refuse to be deleted degrade safely).
- **Long tasks don't blow the context**: only the latest screenshot is kept; history over the threshold is folded into a handoff summary; a circuit breaker backstops.

<details>
<summary>Repository layout</summary>

| File | Purpose |
|---|---|
| `manifest.xml` | Standard Office add-in manifest (TaskPaneApp) |
| `server.js` | Local static server + generic reverse proxy + protocol translation (port 3010) |
| `public/taskpane.*` | Chat UI, 11 tool implementations, settings with three-way test |
| `public/blank.pptx` | Minimal template for `add_slides` (insertSlidesFromBase64) |
| `public/office.js etc.` | Local fallback copies of the Office.js CDN files (official Microsoft files) |
| `scripts/*.ps1` | Screenshot rendering / precise image insertion / icon generation (PowerPoint COM) |
| `install-*.ps1` | Sideload registration (HKCU, no admin; optional shared-catalog fallback needs admin) |
| `autostart-hidden.vbs` | Autostart the local service in a hidden window |
| `start-addin.bat` | Manual start (visible window, for debugging) |
| `tests/` | Manual verification payloads for the protocol translation layer |
| `logs/` | Diagnostic logs (panel errors are reported here, gitignored) |

</details>

## Security & privacy

- The API key lives only in the pane's localStorage (browser storage) — **never in the repo, never uploaded**; the local service forwards in memory only, writes nothing, no telemetry.
- Your presentation content is sent only to the upstream you configure (https enforced); switching models means switching providers.
- `server.js` is a few hundred lines, zero dependencies — audit it yourself; it listens on `127.0.0.1` only and rejects cross-origin requests.
- Use a reputable model provider; the optional shared-catalog script creates an SMB share — remove it afterwards with `net share addincatalog /delete`.

## Troubleshooting

| Symptom | Fix |
|---|---|
| No Slilot button in PowerPoint | Check the service is up (`https://localhost:3010/healthz` returns ok) → restart PowerPoint → still missing: run `install-shared-catalog-admin.ps1` as admin and restart |
| Pane won't load / stuck initializing | Check `logs/client-log.txt`; usually the local service isn't running or the certificate isn't trusted |
| No Slilot button on the Home tab at startup | Known limitation of developer sideloading on perpetual Office 2021 — ribbon commands register on first launch of each session; open it once from the Add-ins menu. (Icons are served from GitHub Pages and need internet access to x1han.github.io.) |
| Certificate expired (pane suddenly won't load) | The dev certificate defaults to 30-day validity. Refresh with `npx office-addin-dev-certs install --days 3650`, then restart the local service |
| GitHub unreachable (clone fails) | Direct github.com access is blocked in some networks — use SSH clone (`git clone git@github.com:x1han/slilot.git`, requires an SSH key on your GitHub account) or a proxy. Everything else runs without GitHub: icons/FunctionFile live on x1han.github.io (reachable directly), office.js has a local fallback, and the model upstream is your own provider |
| Tests fail | Chat: check upstream URL / protocol / key. Vision: the model doesn't accept images — switch models. Image: the key likely lacks image-generation access |
| Code changes not taking effect | Restart the local service and reopen the pane (×); if `manifest.xml` changed, bump its `<Version>` and restart PowerPoint |
| COM errors on image/screenshot | Make sure PowerPoint has the target deck open (COM attaches to the active presentation); close any blocking dialogs and retry |

## Uninstall

- Delete the value named after the manifest Id (`a7f3d9e2-…`) under `HKCU\Software\Microsoft\Office\16.0\Wef\Developer`, then restart PowerPoint;
- Remove the Slilot shortcut from the Startup folder and stop the node process;
- If you used the shared-catalog option, run `net share addincatalog /delete` and delete `C:\addin-catalog`.

## Capability limits

- **Windows desktop PowerPoint only**; PowerPoint for the web / Mac / WPS are not supported.
- Operations the PowerPoint API doesn't cover (animations, transitions, master design, SmartArt internals, chart data editing) are out of scope.
- The pane probes host capabilities and trims its toolset automatically; per-slide review uses real screenshots and requires a vision-capable model (testable in settings).

## License

Released under the [MIT License](./LICENSE).

---

If this is useful, give it a Star ⭐ ; issues and PRs are welcome.
