(function () {
  "use strict";

  window.PGM_I18N.registerLocale({
    code: "en",
    flag: "🇺🇸",
    label: "English",
    strings: {
      "app.title": "Grid Maker",
      "label.language": "Language",
      "label.rows": "Rows (height)",
      "label.cols": "Columns (width)",
      "label.ratio": "Target source ratio",
      "label.preview": "Preview",
      "label.debug": "Debug",
      "action.copyDebug": "Copy",
      "action.downloadUpdate": "Download update",
      "help.cellClick": "Click a cell below: the selected timeline video clip will be placed automatically.",
      "summary.format": "{rows} x {cols} | ratio {ratio}",
      "cell.label": "{row},{col}",
      "update.available": "Update available: v{latest} (installed v{current})",
      "status.ready": "Ready. Select a timeline video clip, then click a grid cell.",
      "status.applying": "Applying to selected clip...",
      "status.err.cep_unavailable": "CEP runtime is unavailable.",
      "status.err.cep_bridge_unavailable": "CEP bridge is unavailable.",
      "status.err.no_active_sequence": "No active sequence.",
      "status.err.invalid_grid": "Invalid grid settings.",
      "status.err.cell_out_of_bounds": "Cell index is out of bounds.",
      "status.err.invalid_ratio": "Invalid ratio.",
      "status.err.no_selection": "Select a clip in the timeline.",
      "status.err.no_video_selected": "No selected video clip found.",
      "status.err.multiple_video_selected": "Select exactly one video clip (audio can stay linked).",
      "status.err.qe_unavailable": "QE sequence is unavailable.",
      "status.err.qe_clip_not_found": "Unable to find selected clip in QE API.",
      "status.err.invalid_sequence_size": "Unable to read sequence frame size.",
      "status.err.placement_apply_failed": "Failed to apply clip placement values.",
      "status.err.motion_effect_unavailable": "Motion component unavailable on selected clip.",
      "status.err.transform_effect_unavailable": "Transform effect unavailable or failed to add.",
      "status.err.crop_effect_unavailable": "Crop effect unavailable or failed to add.",
      "status.err.exception": "Unexpected error: {message}",
      "status.err.empty_response": "No response returned by host script.",
      "status.err.copy_failed": "Failed to copy debug log.",
      "status.err.unknown": "Unknown error.",
      "status.ok.generic": "Done.",
      "status.ok.debug_copied": "Debug log copied to clipboard.",
      "status.info.host": "Host: {message}"
    }
  });
})();
