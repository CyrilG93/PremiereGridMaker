# Grid Maker (CEP) - v1.1.0

Premiere Pro 2025+ extension to place timeline clips into a video grid fast.

## Features

- Grid size with two sliders: rows x columns (1 to 10)
- Ratio preset selector: `16:9`, `1:1`, `9:16`, `4:5`, `3:2`
- Clickable live grid preview
- One-click placement to a target cell using `Transform` + `Crop`
- No manual position presets required
- Grid Designer mode (10x10 canvas): irregular layouts with draggable/resizable blocks
- Designer presets per ratio, saved locally with gallery preview and quick load/delete
- UI localization: English (default) and French
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
reg add "HKCU\Software\Adobe\CSXS.11" /v PlayerDebugMode /t REG_SZ /d 1 /f
```

macOS:
```bash
defaults write com.adobe.CSXS.11 PlayerDebugMode 1
```

3. Open Premiere Pro:

- `Window > Extensions > Grid Maker`

## Usage

1. Open a sequence.
2. Select one video clip on timeline.
3. Set rows, columns, and ratio.
4. Click a target cell in the preview.

The extension adds/uses `Transform` and `Crop`, then positions/scales the selected clip for the target grid cell.

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
