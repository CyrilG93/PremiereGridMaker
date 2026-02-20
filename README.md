# Grid Maker (CEP) - v1.2.5

Premiere Pro 2025+ extension to place timeline clips into a video grid fast.

## Features

- Grid size with two sliders: rows x columns (1 to 10)
- Ratio preset selector: `16:9`, `1:1`, `9:16`, `4:5`, `3:2`
- Global margin slider (px) applied to outer margins and spacing between cells/blocks
- Clickable live grid preview
- One-click placement to a target cell using `Transform` + `Crop`
- Batch apply: map selected timeline clips to cells in one click (ordered by track from bottom to top)
- No manual position presets required
- Grid Designer mode (10x10 canvas): irregular layouts with draggable/resizable blocks
- Designer presets per ratio, saved locally with gallery preview and quick load/delete
- Designer presets also store their own Global Margin value
- Designer preset import/export in JSON (team sharing + backup)
- UI localization: English (default), French, Spanish, German, Portuguese (Brazil), Japanese, Italian, Chinese (Simplified), Russian
- Language selector with flag dropdown
- Collapsible debug panel (collapsed by default)
- Built-in update check against latest GitHub release (with direct ZIP download button when newer version exists)

## Installation (Recommended)

The easiest and safest method is to use the installer scripts included in this repo.

### Windows

```powershell
cd C:\path\to\PremiereGridMaker
.\install-win.bat
```

### macOS

```bash
cd /path/to/PremiereGridMaker
chmod +x ./install-macos.sh
./install-macos.sh
```

### Optional flags

- Install for all users/system scope:
  - Windows: `.\install-win.bat --scope System`
  - macOS: `sudo ./install-macos.sh --scope system`
- Skip debug mode changes:
  - Windows: `.\install-win.bat --skip-debug`
  - macOS: `./install-macos.sh --skip-debug`

## Manual Installation (Fallback)

Use this only if installer scripts are not available.

1. Copy extension files to one CEP extensions path:

- Windows user scope: `C:\Users\<your-user>\AppData\Roaming\Adobe\CEP\extensions\PremiereGridMaker`
- Windows system scope: `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\PremiereGridMaker`
- macOS user scope: `~/Library/Application Support/Adobe/CEP/extensions/PremiereGridMaker`
- macOS system scope: `/Library/Application Support/Adobe/CEP/extensions/PremiereGridMaker`

2. Enable CEP debug mode:

Windows:
```powershell
8..15 | ForEach-Object { reg add "HKCU\Software\Adobe\CSXS.$_" /v PlayerDebugMode /t REG_SZ /d 1 /f }
```

macOS:
```bash
for v in {8..15}; do defaults write "com.adobe.CSXS.$v" PlayerDebugMode 1; done
```

3. Open Premiere Pro:

- `Window > Extensions > Grid Maker`

## Usage

1. Open a sequence.
2. Select one video clip on timeline.
3. Set rows, columns, and ratio.
4. Optional: adjust global margin (px).
5. Click a target cell in the preview.

The extension adds/uses `Transform` and `Crop`, then positions/scales the selected clip for the target grid cell.

### Batch apply

1. Select multiple video clips in the timeline.
2. Keep the desired grid/designer layout active.
3. Click `Apply batch`.

Batch order is deterministic:

- clips are sorted by track from bottom to top (`V1`, then `V2`, etc.),
- then by timeline start time inside each track.

### Designer import/export

- `Import Configuration`: merge configs from a `.json` file into local Designer presets.
- `Export Configuration`: save all local Designer presets into a shareable `.json` file.

## Premiere Preference Recommendation (Important)

To avoid scale interpretation mismatches (especially in mixed 4K/HD sequences), set:

- `Premiere Pro > Preferences > Media > Default Media Scaling = None` (recommended for deterministic native behavior)

If your team uses a Fit workflow, keep it consistent on all clips before using Grid Maker (do not mix clips with and without frame-fit behavior in the same operation).

## Compatibility

- Host: Adobe Premiere Pro 2025+ (`PPRO 25.0+`)
- OS: Windows and macOS
- Technology: CEP panel + ExtendScript/QE API

## Project Structure

- `CSXS/manifest.xml`: CEP extension manifest
- `index.html`: panel markup
- `css/style.css`: panel styling
- `js/main.js`: UI + CEP bridge + i18n runtime
- `js/i18n-registry.js`: language registry
- `js/locales/*.js`: locale dictionaries
- `jsx/hostscript.jsx`: Premiere ExtendScript logic

## Localization

Localization files are modular and easy to extend:

- `js/i18n-registry.js`: locale registry
- `js/locales/en.js`: English strings
- `js/locales/fr.js`: French strings
- `js/locales/es.js`: Spanish strings
- `js/locales/de.js`: German strings
- `js/locales/pt-BR.js`: Portuguese (Brazil) strings
- `js/locales/ja.js`: Japanese strings
- `js/locales/it.js`: Italian strings
- `js/locales/zh-CN.js`: Chinese (Simplified) strings
- `js/locales/ru.js`: Russian strings

To add a new language, add a new file in `js/locales/` and call:

```js
window.PGM_I18N.registerLocale({
  code: "es",
  flag: "🇪🇸",
  label: "Espanol",
  strings: {
    "app.title": "Grid Maker"
  }
});
```

Then include it in `index.html` before `js/main.js`.

## Changelog

### v1.2.5

- Added batch apply for selected clips (ordered by track from bottom to top, then by clip start time).
- Added a global margin control (px) for classic grid and designer layouts, with persistence inside designer presets.
- Added designer preset import/export (`Configuration`) to share and back up layouts in JSON.
- Improved localization coverage for the new actions across all supported languages (including batch apply label translations).
- Added high-level inline comments in main panel and host scripts to simplify long-term maintenance.

### v1.2.0

- Added 7 new UI localizations: Spanish, German, Portuguese (Brazil), Japanese, Italian, Chinese (Simplified), Russian.
- Improved small-panel usability with a minimum layout height and vertical scrolling instead of shrinking controls out of view.
- Designer presets gallery is now collapsible (like Debug), with better compact layout and cleaner header behavior.
- Fixed designer preset thumbnail aspect ratio so each saved config preview respects its own ratio (for example `9:16` no longer shown as `16:9`).
- Reduced gallery thumbnail size range and fixed gallery size persistence across Premiere restarts.
- Removed duplicate Debug title in expanded panel (single title source in collapse header).

### v1.1.6

- Switched placement scale computation to strict native mode (no implicit fit-to-frame assumption).
- Added documentation note for `Default Media Scaling` preference to avoid scale mismatches on mixed-resolution timelines.

### v1.1.5

- Added Designer overlap visualization on true intersection areas only (instead of styling the whole block), with clearer overlap badges.
- Improved overlapping-block interaction by keeping selected block on top during edit actions.
- Hardened QE clip targeting by prioritizing selected QE items to avoid cross-track effect insertion mismatches.
- Stabilized effect ensure flow by requiring real clip component visibility before continuing placement.
- Enforced `Transform > Uniform Scale` activation with explicit toggle fallback to improve default Transform behavior after automatic insertion.

### v1.1.4

- Fixed QE clip targeting across duplicated clips on multiple tracks by strongly prioritizing the selected track during QE matching.
- Improved track index resolution using parent track, clip identity, and nodeId fallback checks.
- Reduced host-side insertion latency by removing long settle waits while keeping duplicate-add protection.
- Disabled UI auto-retry loop for apply requests to keep interactions immediate and deterministic.

### v1.1.3

- Stabilized effect insertion flow to prevent duplicate `Transform`/`Crop` additions when Premiere/QE effect visibility is delayed.
- Added queued apply requests in panel UI to avoid overlapping host calls during rapid cell clicks.
- Added short auto-retry for transient `transform_effect_unavailable` / `crop_effect_unavailable` responses.
- Improved debug traces with QE effect counts before/after insertion attempts.

### v1.1.2

- Enforced strict placement pipeline: Motion for size/position, Transform effect always required (neutral, no parameter writes).
- Improved universal Transform detection and stabilization after effect insertion to reduce QE timing issues.
- Kept crop path strictly on Crop effect components (never Motion-integrated crop behavior).

### v1.1.1

- Fixed intermittent `crop_effect_unavailable` failures caused by delayed QE effect visibility after `addVideoEffect`.
- Improved effect detection stabilization with short retries before concluding that Crop is unavailable.
- Enforced crop handling to use Crop effect components only (never Motion-integrated crop behavior).
- Avoided hard failure when computed crop is zero and Crop effect is unavailable (placement can still succeed).

### v1.1.0

- Added full `Grid Designer` mode with irregular block layouts on a `10x10` canvas.
- Added drag/drop + resize editing flow with `Edit ON/OFF` state and stronger visual feedback.
- Added local designer preset storage by ratio, with gallery preview, quick load, delete, and size slider.
- Added custom-cell host placement API for designer blocks with normalized bounds.
- Increased classic grid limits from `8x8` to `10x10` in UI and placement validation.
- Improved classic preview responsiveness in small panels (compact rendering and tighter spacing).

### v1.0.6

- Fixed scale computation when source and sequence ratios differ (notably `9:16` and `1:1` timelines).
- Grid cell fill is now consistent with crop direction rules, including Motion normalized workflows.
- Added stronger QE/native-size diagnostics to stabilize placement behavior.

### v1.0.5

- Fixed update banner download action by aligning URL opening with the same CEP method used in other Premiere panels.
- Added stronger fallback chain for opening update download URLs.

### v1.0.4

- Fixed update banner click behavior so direct ZIP download URLs open reliably.
- Improved update URL handling for GitHub release download variants (including query string forms).

### v1.0.3

- Improved clip fill behavior by grid cell orientation:
  - Portrait and square cells prioritize full-height fill (crop left/right)
  - Landscape cells prioritize full-width fill (crop top/bottom)
- Improved preview responsiveness in the panel:
  - Preview always fits inside available space
  - No overflow in vertical ratios like `9:16` (letterbox/pillarbox when needed)
