# Grid Maker (CEP)

Premiere Pro 2025+ extension to place timeline clips into a video grid fast.

## Features

- Grid size with two sliders: rows x columns (1 to 8)
- Ratio preset selector: `16:9`, `1:1`, `9:16`, `4:5`, `3:2`
- Clickable live grid preview
- One-click placement to a target cell using `Transform` + `Crop`
- No manual position presets required
- UI localization: English (default) and French
- Language selector with flag dropdown

## Localization system

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

## Project structure

- `CSXS/manifest.xml`: CEP extension manifest
- `index.html`: panel markup
- `css/style.css`: panel styling
- `js/main.js`: UI + CEP bridge + i18n runtime
- `js/i18n-registry.js`: language registry
- `js/locales/*.js`: locale dictionaries
- `jsx/hostscript.jsx`: Premiere ExtendScript logic

## Compatibility

- Host: Adobe Premiere Pro 2025+ (`PPRO 25.0+`)
- OS: Windows and macOS
- Technology: CEP panel + ExtendScript/QE API

## Install scripts

Windows (PowerShell):

```powershell
cd C:\Users\Admin\CascadeProjects\PremiereGridMaker
.\install-win.bat
```

macOS (Terminal):

```bash
cd /path/to/PremiereGridMaker
chmod +x ./install-macos.sh
./install-macos.sh
```

Options:

- Install for all users/system scope:
  - Windows: `.\install-win.bat --scope System`
  - macOS: `sudo ./install-macos.sh --scope system`
- Skip debug mode changes:
  - Windows: `.\install-win.bat --skip-debug`
  - macOS: `./install-macos.sh --skip-debug`

## Manual install (fallback)

1. Copy this folder as `PremiereGridMaker` into CEP extensions directory.

Windows:

- `C:\Program Files (x86)\Common Files\Adobe\CEP\extensions\PremiereGridMaker`
- `C:\Users\<your-user>\AppData\Roaming\Adobe\CEP\extensions\PremiereGridMaker`

macOS:

- `/Library/Application Support/Adobe/CEP/extensions/PremiereGridMaker`
- `~/Library/Application Support/Adobe/CEP/extensions/PremiereGridMaker`

2. Enable CEP debug mode.

Windows (PowerShell):

```powershell
reg add "HKCU\Software\Adobe\CSXS.11" /v PlayerDebugMode /t REG_SZ /d 1 /f
```

macOS (Terminal):

```bash
defaults write com.adobe.CSXS.11 PlayerDebugMode 1
```

3. Open Premiere Pro and launch the panel:

- `Window > Extensions (Legacy) > Grid Maker`

## Usage

1. Open a sequence.
2. Select one video clip on timeline.
3. Set rows, columns, and ratio.
4. Click a target cell in the preview.

The extension adds/uses `Transform` and `Crop`, then positions/scales the selected clip for the target grid cell.
