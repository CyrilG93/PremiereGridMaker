(function () {
  "use strict";

  var cep = window.__adobe_cep__ || null;
  var cepBridge = window.cep || null;
  var csInterface = (typeof CSInterface !== "undefined") ? new CSInterface() : null;
  var i18n = window.PGM_I18N || { defaultLocale: "en", locales: {} };
  var APP_VERSION = "1.2.0";
  var RELEASE_API_URL = "https://api.github.com/repos/CyrilG93/PremiereGridMaker/releases/latest";
  var DESIGNER_GRID_SIZE = 10;
  var DESIGNER_FREE_SUBDIVISION = 10;
  var DESIGNER_MIN_BLOCK_SIZE = 1 / DESIGNER_FREE_SUBDIVISION;
  var DESIGNER_GALLERY_MIN = 56;
  var DESIGNER_GALLERY_MAX = 140;
  var DESIGNER_GALLERY_DEFAULT = 64;

  // Central runtime state for UI controls, active ratio, locale and designer mode.
  var state = {
    rows: 2,
    cols: 2,
    ratioW: 16,
    ratioH: 9,
    marginPx: resolveInitialMarginPx(),
    selectedCell: { row: 0, col: 0 },
    locale: resolveInitialLocale(),
    designer: {
      enabled: false,
      editMode: false,
      freeMode: false,
      blocks: [],
      selectedBlockId: "",
      configs: [],
      activeConfigId: "",
      nextBlockSeq: 1,
      loaded: false,
      gallerySize: resolveInitialGallerySize()
    }
  };

  var statusState = {
    mode: "key",
    key: "status.ready",
    vars: {},
    kind: ""
  };

  var updateState = {
    visible: false,
    current: APP_VERSION,
    latest: "",
    downloadUrl: ""
  };
  var applyRequestState = {
    inFlight: false,
    queued: null
  };
  var APPLY_RETRY_DELAY_MS = 450;
  var APPLY_MAX_RETRIES = 0;

  // Cache DOM references once to keep rendering and event handlers simple.
  var rowsRange = document.getElementById("rows");
  var rowsNumber = document.getElementById("rowsNumber");
  var colsRange = document.getElementById("cols");
  var colsNumber = document.getElementById("colsNumber");
  var ratio = document.getElementById("ratio");
  var marginRange = document.getElementById("marginRange");
  var summary = document.getElementById("summary");
  var gridStage = document.getElementById("gridStage");
  var grid = document.getElementById("gridPreview");
  var status = document.getElementById("status");
  var languageSelect = document.getElementById("languageSelect");
  var copyDebugBtn = document.getElementById("copyDebugBtn");
  var debugPanel = document.getElementById("debugPanel");
  var debugLog = document.getElementById("debugLog");
  var appVersion = document.getElementById("appVersion");
  var updateBanner = document.getElementById("updateBanner");
  var updateLink = document.getElementById("updateLink");

  var helpText = document.querySelector(".help");
  var designerModeBtn = document.getElementById("designerModeBtn");
  var classicGridControls = document.getElementById("classicGridControls");
  var designerControls = document.getElementById("designerControls");
  var designerEditBtn = document.getElementById("designerEditBtn");
  var designerFreeBtn = document.getElementById("designerFreeBtn");
  var designerAddBtn = document.getElementById("designerAddBtn");
  var designerRemoveBtn = document.getElementById("designerRemoveBtn");
  var designerNewBtn = document.getElementById("designerNewBtn");
  var designerSaveBtn = document.getElementById("designerSaveBtn");
  var designerImportBtn = document.getElementById("designerImportBtn");
  var designerExportBtn = document.getElementById("designerExportBtn");
  var designerNameInput = document.getElementById("designerNameInput");
  var designerGalleryPanel = document.getElementById("designerGalleryPanel");
  var designerGallery = document.getElementById("designerGallery");
  var designerGallerySize = document.getElementById("designerGallerySize");
  var designerGalleryTools = document.getElementById("designerGalleryTools");
  var applyBatchBtn = document.getElementById("applyBatchBtn");

  var designerDrag = null;

  // Debug helpers: timestamped logs in the collapsible debug panel.
  function getClockStamp() {
    var now = new Date();
    var hh = String(now.getHours()).padStart(2, "0");
    var mm = String(now.getMinutes()).padStart(2, "0");
    var ss = String(now.getSeconds()).padStart(2, "0");
    return hh + ":" + mm + ":" + ss;
  }

  function appendDebug(message) {
    if (!debugLog) {
      return;
    }
    debugLog.value += "[" + getClockStamp() + "] " + String(message) + "\n";
    if (debugLog.value.length > 60000) {
      debugLog.value = debugLog.value.slice(debugLog.value.length - 60000);
    }
    debugLog.scrollTop = debugLog.scrollHeight;
  }

  function appendHostDebug(rawDebugText) {
    if (!rawDebugText) {
      return;
    }
    var lines = String(rawDebugText).split(/\r?\n/);
    for (var i = 0; i < lines.length; i += 1) {
      if (lines[i]) {
        appendDebug("HOST> " + lines[i]);
      }
    }
  }

  // Locale and i18n bootstrapping from localStorage + dictionary registry.
  function resolveInitialLocale() {
    var fallback = i18n.defaultLocale || "en";
    var stored = null;

    try {
      stored = window.localStorage.getItem("pgm.locale");
    } catch (e1) {
      stored = null;
    }

    if (stored && i18n.locales[stored]) {
      return stored;
    }
    if (i18n.locales[fallback]) {
      return fallback;
    }

    var keys = Object.keys(i18n.locales);
    if (keys.length > 0) {
      return keys[0];
    }
    return "en";
  }

  function resolveInitialGallerySize() {
    var fallback = DESIGNER_GALLERY_DEFAULT;
    var stored = null;

    try {
      stored = window.localStorage.getItem("pgm.designer.gallerySize");
    } catch (e1) {
      stored = null;
    }

    var parsed = parseInt(stored, 10);
    if (isNaN(parsed)) {
      return fallback;
    }
    return clampInt(parsed, DESIGNER_GALLERY_MIN, DESIGNER_GALLERY_MAX, fallback);
  }

  // Load the global margin (in sequence pixels) used for classic/designer placement.
  function resolveInitialMarginPx() {
    var fallback = 0;
    var stored = null;

    try {
      stored = window.localStorage.getItem("pgm.marginPx");
    } catch (e1) {
      stored = null;
    }

    var parsed = parseInt(stored, 10);
    if (isNaN(parsed)) {
      return fallback;
    }
    return clampInt(parsed, 0, 200, fallback);
  }

  function getLocaleStrings() {
    var locale = i18n.locales[state.locale] || i18n.locales[i18n.defaultLocale];
    return locale ? locale.strings : {};
  }

  function getDefaultLocaleStrings() {
    var fallback = i18n.locales[i18n.defaultLocale];
    return fallback ? fallback.strings : {};
  }

  function format(template, vars) {
    var safeVars = vars || {};
    return String(template).replace(/\{([a-zA-Z0-9_]+)\}/g, function (_, key) {
      return (safeVars[key] !== undefined) ? String(safeVars[key]) : "";
    });
  }

  function hasText(key) {
    var strings = getLocaleStrings();
    if (strings[key] !== undefined) {
      return true;
    }
    var fallbackStrings = getDefaultLocaleStrings();
    return fallbackStrings[key] !== undefined;
  }

  function t(key, vars) {
    var strings = getLocaleStrings();
    var template = strings[key];
    if (template === undefined) {
      var fallbackStrings = getDefaultLocaleStrings();
      template = fallbackStrings[key];
    }
    if (template === undefined) {
      return key;
    }
    return format(template, vars);
  }

  function setStatusText(text, kind) {
    if (!status) {
      return;
    }
    status.className = "status" + (kind ? " " + kind : "");
    status.textContent = text;
  }

  function setStatusKey(key, vars, kind) {
    statusState.mode = "key";
    statusState.key = key;
    statusState.vars = vars || {};
    statusState.kind = kind || "";
    setStatusText(t(key, vars), kind);
  }

  function setStatusRaw(text, kind) {
    statusState.mode = "raw";
    statusState.key = "";
    statusState.vars = { message: text };
    statusState.kind = kind || "";
    setStatusText(text, kind);
  }

  function refreshStatus() {
    if (statusState.mode === "key") {
      setStatusText(t(statusState.key, statusState.vars), statusState.kind);
      return;
    }

    if (hasText("status.info.host")) {
      setStatusText(t("status.info.host", { message: statusState.vars.message || "" }), statusState.kind);
      return;
    }

    setStatusText(statusState.vars.message || "", statusState.kind);
  }

  // Release/update helpers (version compare + safe download URL filtering).
  function normalizeVersion(raw) {
    var clean = String(raw || "").trim().replace(/^v/i, "");
    var match = clean.match(/^(\d+)\.(\d+)\.(\d+)/);
    if (!match) {
      return "";
    }
    return match[1] + "." + match[2] + "." + match[3];
  }

  function compareVersions(a, b) {
    var pa = normalizeVersion(a).split(".");
    var pb = normalizeVersion(b).split(".");
    if (pa.length !== 3 || pb.length !== 3) {
      return 0;
    }
    for (var i = 0; i < 3; i += 1) {
      var da = parseInt(pa[i], 10);
      var db = parseInt(pb[i], 10);
      if (da > db) {
        return 1;
      }
      if (da < db) {
        return -1;
      }
    }
    return 0;
  }

  function resolveReleaseZipUrl(release) {
    if (release && release.assets && release.assets.length) {
      for (var i = 0; i < release.assets.length; i += 1) {
        var asset = release.assets[i];
        var name = String(asset.name || "").toLowerCase();
        var url = asset.browser_download_url || "";
        if (name.indexOf(".zip") !== -1 && url) {
          return url;
        }
      }
    }
    return "";
  }

  function isTrustedReleaseZipUrl(url) {
    var raw = String(url || "");
    return /^https:\/\/github\.com\/CyrilG93\/PremiereGridMaker\/releases\/(?:download\/v[0-9]+\.[0-9]+\.[0-9]+|latest\/download)\/[^?#]+\.zip(?:[?#].*)?$/i.test(raw);
  }

  function quoteForEvalScript(value) {
    var s = String(value || "");
    return "'" + s.replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'";
  }

  function refreshUpdateBanner() {
    if (!updateBanner || !updateLink) {
      return;
    }

    if (!updateState.visible || !updateState.latest || !updateState.downloadUrl) {
      updateBanner.hidden = true;
      updateLink.href = "#";
      schedulePreviewFit();
      return;
    }

    updateBanner.hidden = false;
    updateLink.textContent = hasText("update.download_notice")
      ? t("update.download_notice", { latest: updateState.latest, current: updateState.current })
      : ("New update available (v" + updateState.latest + "), click here to download.");
    updateLink.href = updateState.downloadUrl;
    schedulePreviewFit();
  }

  function checkForUpdates() {
    if (!window.fetch) {
      appendDebug("UPDATE> fetch unavailable in CEP runtime");
      updateState.visible = false;
      refreshUpdateBanner();
      return;
    }

    updateState.visible = false;
    updateState.latest = "";
    updateState.downloadUrl = "";
    refreshUpdateBanner();

    appendDebug("UPDATE> checking latest release");
    window.fetch(RELEASE_API_URL, { cache: "no-store" }).then(function (response) {
      if (!response.ok) {
        throw new Error("HTTP " + response.status);
      }
      return response.json();
    }).then(function (release) {
      var latest = normalizeVersion(release && release.tag_name);
      var current = normalizeVersion(APP_VERSION);
      if (!latest) {
        appendDebug("UPDATE> latest release tag missing/invalid");
        updateState.visible = false;
        refreshUpdateBanner();
        return;
      }

      if (!current) {
        appendDebug("UPDATE> current app version invalid");
        updateState.visible = false;
        refreshUpdateBanner();
        return;
      }

      if (latest === current || compareVersions(latest, current) !== 1) {
        appendDebug("UPDATE> up to date (local v" + current + ", latest v" + latest + ")");
        updateState.visible = false;
        refreshUpdateBanner();
        return;
      }

      var zipUrl = resolveReleaseZipUrl(release);
      if (!zipUrl) {
        appendDebug("UPDATE> newer version found but no zip URL available");
        updateState.visible = false;
        refreshUpdateBanner();
        return;
      }

      updateState.visible = true;
      updateState.latest = latest;
      updateState.current = current;
      updateState.downloadUrl = zipUrl;
      refreshUpdateBanner();
      appendDebug("UPDATE> update available v" + latest + " zip=" + zipUrl);
    }).catch(function (err) {
      updateState.visible = false;
      refreshUpdateBanner();
      appendDebug("UPDATE> check failed: " + err);
    });
  }

  function openExternalUrl(url) {
    if (!url || !isTrustedReleaseZipUrl(url)) {
      appendDebug("UI> rejected untrusted update URL: " + url);
      return false;
    }

    if (openExternalUrlFallback(url)) {
      return true;
    }

    var script = "gridMaker_openExternalUrl(" + quoteForEvalScript(url) + ")";
    callHost(script, function (result) {
      appendDebug("HOST< raw(open-url): " + (result || "<empty>"));
      var parsed = parseHostResponse(result);
      appendHostDebug(parsed.hostDebug);
      if (parsed.kind === "ok") {
        appendDebug("UI> update URL opened via host");
        return;
      }
      appendDebug("UI> host open-url failed, trying CEP fallback");
      openExternalUrlFallback(url);
    });
    return false;
  }

  function openExternalUrlFallback(url) {
    try {
      if (csInterface && typeof csInterface.openURLInDefaultBrowser === "function") {
        csInterface.openURLInDefaultBrowser(url);
        appendDebug("UI> update URL opened via CSInterface");
        return true;
      }
    } catch (e1) {}

    try {
      if (cepBridge && cepBridge.util && typeof cepBridge.util.openURLInDefaultBrowser === "function") {
        cepBridge.util.openURLInDefaultBrowser(url);
        appendDebug("UI> update URL opened via window.cep.util");
        return true;
      }
    } catch (e2) {}

    try {
      var popup = window.open(url, "_blank");
      if (popup) {
        appendDebug("UI> update URL opened via window.open");
        return true;
      }
    } catch (e3) {}

    try {
      window.location.href = url;
      appendDebug("UI> update URL opened via window.location fallback");
      return true;
    } catch (e4) {}

    appendDebug("UI> failed to open update URL");
    return false;
  }

  // Generic numeric/string helpers shared by grid + designer flows.
  function parseRatio(value) {
    var parts = value.split(":");
    if (parts.length !== 2) {
      return { w: 16, h: 9 };
    }

    var w = parseFloat(parts[0]);
    var h = parseFloat(parts[1]);
    if (!w || !h) {
      return { w: 16, h: 9 };
    }

    return { w: w, h: h };
  }

  function clampInt(value, min, max, fallback) {
    var raw = parseInt(value, 10);
    if (isNaN(raw)) {
      raw = fallback;
    }
    if (raw < min) {
      return min;
    }
    if (raw > max) {
      return max;
    }
    return raw;
  }

  function clampNumber(value, min, max) {
    var n = parseFloat(value);
    if (isNaN(n)) {
      n = min;
    }
    if (n < min) {
      return min;
    }
    if (n > max) {
      return max;
    }
    return n;
  }

  function roundToStep(value, step) {
    if (!(step > 0)) {
      return value;
    }
    return Math.round(value / step) * step;
  }

  function clampStep(value, min, max, step, fallback) {
    var raw = parseFloat(value);
    if (isNaN(raw)) {
      raw = fallback;
    }
    if (step > 0) {
      raw = roundToStep(raw, step);
    }
    if (raw < min) {
      raw = min;
    }
    if (raw > max) {
      raw = max;
    }
    return Math.round(raw * 1000) / 1000;
  }

  function getDesignerStep() {
    return state.designer.freeMode ? (1 / DESIGNER_FREE_SUBDIVISION) : 1;
  }

  function getDesignerGridUnits() {
    return DESIGNER_GRID_SIZE / getDesignerStep();
  }

  function applyDesignerGallerySize(nextSize, persist) {
    var size = clampInt(nextSize, DESIGNER_GALLERY_MIN, DESIGNER_GALLERY_MAX, DESIGNER_GALLERY_DEFAULT);
    state.designer.gallerySize = size;

    if (designerGallery) {
      designerGallery.style.setProperty("--gallery-card-min", size + "px");
    }
    if (designerGallerySize) {
      designerGallerySize.value = String(size);
    }

    if (!persist) {
      return;
    }
    try {
      window.localStorage.setItem("pgm.designer.gallerySize", String(size));
    } catch (e1) {
      // Ignore localStorage issues in CEP hosts.
    }
  }

  // Persist and normalize the global margin slider value.
  function applyGlobalMarginPx(nextMargin, persist) {
    var marginPx = clampInt(nextMargin, 0, 200, 0);
    state.marginPx = marginPx;

    if (marginRange) {
      marginRange.value = String(marginPx);
    }

    if (!persist) {
      return;
    }
    try {
      window.localStorage.setItem("pgm.marginPx", String(marginPx));
    } catch (e1) {
      // Ignore localStorage issues in CEP hosts.
    }
  }

  function syncValue(inputA, inputB, callback) {
    inputA.addEventListener("input", function () {
      var v = clampInt(inputA.value, parseInt(inputA.min, 10), parseInt(inputA.max, 10), parseInt(inputB.value, 10));
      inputA.value = String(v);
      inputB.value = String(v);
      callback(v);
    });

    inputB.addEventListener("input", function () {
      var v = clampInt(inputB.value, parseInt(inputB.min, 10), parseInt(inputB.max, 10), parseInt(inputA.value, 10));
      inputA.value = String(v);
      inputB.value = String(v);
      callback(v);
    });
  }

  function getRatioText() {
    return state.ratioW + ":" + state.ratioH;
  }

  function formatPercent(v) {
    return clampNumber(v, 0, 1).toFixed(6);
  }

  function formatDesignerSizePercent(gridUnits) {
    var pct = (clampNumber(gridUnits, 0, DESIGNER_GRID_SIZE) / DESIGNER_GRID_SIZE) * 100;
    var rounded = Math.round(pct * 10) / 10;
    if (Math.abs(rounded - Math.round(rounded)) < 0.01) {
      return String(Math.round(rounded)) + "%";
    }
    return rounded.toFixed(1) + "%";
  }

  // Designer block normalization and geometry validation helpers.
  function normalizeDesignerBlock(raw, fallbackId) {
    if (!raw) {
      return null;
    }
    var x = parseFloat(raw.x);
    var y = parseFloat(raw.y);
    var w = parseFloat(raw.w);
    var h = parseFloat(raw.h);
    if (isNaN(x) || isNaN(y) || isNaN(w) || isNaN(h)) {
      return null;
    }
    if (w < DESIGNER_MIN_BLOCK_SIZE || h < DESIGNER_MIN_BLOCK_SIZE) {
      return null;
    }
    if (x < 0 || y < 0) {
      return null;
    }
    if (x + w > DESIGNER_GRID_SIZE + 0.000001 || y + h > DESIGNER_GRID_SIZE + 0.000001) {
      return null;
    }
    return {
      id: String(raw.id || fallbackId || ("cell_" + Date.now())),
      x: Math.round(x * 1000) / 1000,
      y: Math.round(y * 1000) / 1000,
      w: Math.round(w * 1000) / 1000,
      h: Math.round(h * 1000) / 1000
    };
  }

  function cloneDesignerBlocks(rawBlocks) {
    var out = [];
    if (!rawBlocks || !rawBlocks.length) {
      return out;
    }
    for (var i = 0; i < rawBlocks.length; i += 1) {
      var block = normalizeDesignerBlock(rawBlocks[i], "cell_" + i);
      if (block) {
        out.push(block);
      }
    }
    return out;
  }

  function findDesignerBlockById(id) {
    for (var i = 0; i < state.designer.blocks.length; i += 1) {
      if (state.designer.blocks[i].id === id) {
        return state.designer.blocks[i];
      }
    }
    return null;
  }

  function ensureDesignerSelection() {
    if (findDesignerBlockById(state.designer.selectedBlockId)) {
      return;
    }
    if (state.designer.blocks.length > 0) {
      state.designer.selectedBlockId = state.designer.blocks[0].id;
    } else {
      state.designer.selectedBlockId = "";
    }
  }

  function nextDesignerBlockId() {
    var next;
    do {
      next = "cell_" + Date.now() + "_" + state.designer.nextBlockSeq;
      state.designer.nextBlockSeq += 1;
    } while (findDesignerBlockById(next));
    return next;
  }

  function makeDefaultDesignerBlocks() {
    return [{ id: nextDesignerBlockId(), x: 0, y: 0, w: DESIGNER_GRID_SIZE, h: DESIGNER_GRID_SIZE }];
  }

  function adoptDesignerBlocks(blocks) {
    var next = cloneDesignerBlocks(blocks);
    if (!next.length) {
      next = makeDefaultDesignerBlocks();
    }
    state.designer.blocks = next;
    ensureDesignerSelection();
  }

  function designerBlocksOverlap(a, b) {
    if (!a || !b) {
      return false;
    }
    return (a.x < b.x + b.w) &&
      (a.x + a.w > b.x) &&
      (a.y < b.y + b.h) &&
      (a.y + a.h > b.y);
  }

  function designerCanPlace(candidate, ignoreId, allowOverlap) {
    if (!candidate) {
      return false;
    }
    if (candidate.x < 0 || candidate.y < 0 || candidate.w < DESIGNER_MIN_BLOCK_SIZE || candidate.h < DESIGNER_MIN_BLOCK_SIZE) {
      return false;
    }
    if (candidate.x + candidate.w > DESIGNER_GRID_SIZE + 0.000001 || candidate.y + candidate.h > DESIGNER_GRID_SIZE + 0.000001) {
      return false;
    }
    if (allowOverlap === false) {
      for (var i = 0; i < state.designer.blocks.length; i += 1) {
        var other = state.designer.blocks[i];
        if (ignoreId && other.id === ignoreId) {
          continue;
        }
        if (designerBlocksOverlap(candidate, other)) {
          return false;
        }
      }
    }
    return true;
  }

  function designerOverlapCountForBlock(block) {
    if (!block) {
      return 0;
    }
    var count = 0;
    for (var i = 0; i < state.designer.blocks.length; i += 1) {
      var other = state.designer.blocks[i];
      if (!other || other.id === block.id) {
        continue;
      }
      if (designerBlocksOverlap(block, other)) {
        count += 1;
      }
    }
    return count;
  }

  function designerOverlapRegionsForBlock(block) {
    var regions = [];
    if (!block || !(block.w > 0) || !(block.h > 0)) {
      return regions;
    }

    for (var i = 0; i < state.designer.blocks.length; i += 1) {
      var other = state.designer.blocks[i];
      if (!other || other.id === block.id) {
        continue;
      }
      if (!designerBlocksOverlap(block, other)) {
        continue;
      }

      var ix = Math.max(block.x, other.x);
      var iy = Math.max(block.y, other.y);
      var ix2 = Math.min(block.x + block.w, other.x + other.w);
      var iy2 = Math.min(block.y + block.h, other.y + other.h);
      var iw = ix2 - ix;
      var ih = iy2 - iy;
      if (!(iw > 0) || !(ih > 0)) {
        continue;
      }

      regions.push({
        left: ((ix - block.x) * 100) / block.w,
        top: ((iy - block.y) * 100) / block.h,
        width: (iw * 100) / block.w,
        height: (ih * 100) / block.h
      });
    }

    return regions;
  }

  function designerBringBlockToFront(blockId) {
    if (!blockId) {
      return;
    }
    var index = -1;
    for (var i = 0; i < state.designer.blocks.length; i += 1) {
      if (state.designer.blocks[i].id === blockId) {
        index = i;
        break;
      }
    }
    if (index < 0 || index === state.designer.blocks.length - 1) {
      return;
    }
    var moved = state.designer.blocks.splice(index, 1)[0];
    state.designer.blocks.push(moved);
  }

  // Preview sizing helpers keep the visual grid fitted to available panel space.
  function updateSummary() {
    var ratioText = getRatioText();
    if (state.designer.enabled) {
      var activeName = designerNameInput && designerNameInput.value ? designerNameInput.value : t("designer.config_untitled");
      summary.textContent = t("summary.designer_format", {
        cells: state.designer.blocks.length,
        name: activeName,
        ratio: ratioText
      });
    } else {
      summary.textContent = t("summary.format", {
        rows: state.rows,
        cols: state.cols,
        ratio: ratioText
      });
    }
    grid.style.aspectRatio = state.ratioW + " / " + state.ratioH;
  }

  function updateClassicGridDensity(gridWidth, gridHeight) {
    if (!grid || state.designer.enabled) {
      return;
    }

    var safeW = Math.max(1, gridWidth || grid.clientWidth || 1);
    var safeH = Math.max(1, gridHeight || grid.clientHeight || 1);
    var minCell = Math.min(safeW / Math.max(1, state.cols), safeH / Math.max(1, state.rows));
    var gap = 2;

    if (minCell < 22) {
      gap = 1;
    }
    if (minCell < 14) {
      gap = 0;
    }

    grid.style.gap = gap + "px";
    grid.classList.toggle("classic-compact", minCell < 15);
  }

  function fitGridPreview() {
    if (!grid || !gridStage) {
      return;
    }

    var stageWidth = gridStage.clientWidth;
    var stageHeight = gridStage.clientHeight;
    if (!(stageWidth > 0) || !(stageHeight > 0)) {
      return;
    }

    var ratioValue = state.ratioW / state.ratioH;
    if (!(ratioValue > 0)) {
      ratioValue = 16 / 9;
    }

    var nextWidth = stageWidth;
    var nextHeight = nextWidth / ratioValue;
    if (nextHeight > stageHeight) {
      nextHeight = stageHeight;
      nextWidth = nextHeight * ratioValue;
    }

    grid.style.width = Math.max(1, Math.floor(nextWidth)) + "px";
    grid.style.height = Math.max(1, Math.floor(nextHeight)) + "px";
    updateClassicGridDensity(nextWidth, nextHeight);
  }

  function schedulePreviewFit() {
    if (typeof window.requestAnimationFrame === "function") {
      window.requestAnimationFrame(fitGridPreview);
      return;
    }
    window.setTimeout(fitGridPreview, 0);
  }

  // CEP/CSInterface host bridge + parser for hostscript responses.
  function callHost(fnCall, onDone) {
    if (csInterface) {
      csInterface.evalScript(fnCall, function (result) {
        onDone(result || "");
      });
      return;
    }

    if (cep && typeof cep.evalScript === "function") {
      cep.evalScript(fnCall, function (result) {
        onDone(result || "");
      });
      return;
    }

    if (!cep) {
      onDone("ERR|cep_unavailable");
      return;
    }

    onDone("ERR|cep_bridge_unavailable");
  }

  function parseHostDetails(raw) {
    var out = {};
    if (!raw) {
      return out;
    }

    var chunks = raw.split("&");
    for (var i = 0; i < chunks.length; i += 1) {
      var part = chunks[i];
      if (!part) {
        continue;
      }

      var idx = part.indexOf("=");
      if (idx < 0) {
        out[decodeURIComponent(part)] = "";
        continue;
      }

      var key = decodeURIComponent(part.substring(0, idx));
      var value = decodeURIComponent(part.substring(idx + 1));
      out[key] = value;
    }

    return out;
  }

  function parseHostResponse(result) {
    if (!result) {
      return { kind: "err", key: "status.err.empty_response", vars: {}, hostDebug: "", code: "" };
    }

    var parts = result.split("|");
    if (parts.length >= 2 && (parts[0] === "OK" || parts[0] === "ERR")) {
      var statusCode = parts[0];
      var code = parts[1];
      var details = parseHostDetails(parts.slice(2).join("|"));
      var hostDebug = details.debug || "";

      if (statusCode === "OK") {
        if (code === "cell_applied") {
          return {
            kind: "ok",
            key: "status.ok.generic",
            vars: {},
            code: code,
            details: {
              row: details.row || "",
              col: details.col || "",
              scale: details.scale || ""
            },
            hostDebug: hostDebug
          };
        }
        if (code === "batch_applied") {
          return {
            kind: "ok",
            key: "status.batch_applied",
            vars: {
              applied: details.applied || "0",
              total: details.total || "0",
              failed: details.failed || "0",
              skipped: details.skipped || "0"
            },
            code: code,
            details: details,
            hostDebug: hostDebug
          };
        }
        return { kind: "ok", key: "status.ok.generic", vars: {}, hostDebug: hostDebug, code: code, details: details };
      }

      if (code === "exception") {
        return {
          kind: "err",
          key: "status.err.exception",
          vars: { message: details.message || "" },
          hostDebug: hostDebug,
          code: code
        };
      }

      return {
        kind: "err",
        key: hasText("status.err." + code) ? ("status.err." + code) : "status.err.unknown",
        vars: details,
        hostDebug: hostDebug,
        code: code
      };
    }

    if (result.indexOf("ERROR:") === 0) {
      return { kind: "err", raw: result, hostDebug: "", code: "" };
    }

    if (result.indexOf("OK:") === 0) {
      return { kind: "ok", raw: result, hostDebug: "", code: "" };
    }

    return { kind: "", raw: result, hostDebug: "", code: "" };
  }

  // Apply queue ensures only one host placement request runs at a time.
  function isRetryableApplyError(code) {
    return code === "transform_effect_unavailable" || code === "crop_effect_unavailable";
  }

  function runApplyRequest(request, attempt) {
    if (!request || !request.script) {
      return;
    }
    applyRequestState.inFlight = true;

    callHost(request.script, function (result) {
      appendDebug("HOST< raw: " + (result || "<empty>"));
      var parsed = parseHostResponse(result);
      appendHostDebug(parsed.hostDebug);

      if (parsed.kind === "ok" && typeof request.onSuccess === "function") {
        request.onSuccess(parsed);
      }

      if (parsed.kind === "err" && isRetryableApplyError(parsed.code) && attempt < APPLY_MAX_RETRIES) {
        appendDebug(
          "UI> host reported " + parsed.code +
          ", auto-retry " + (attempt + 2) + "/" + (APPLY_MAX_RETRIES + 1) +
          " in " + APPLY_RETRY_DELAY_MS + "ms"
        );
        window.setTimeout(function () {
          runApplyRequest(request, attempt + 1);
        }, APPLY_RETRY_DELAY_MS);
        return;
      }

      if (typeof request.onDone === "function") {
        request.onDone(parsed);
      }

      applyRequestState.inFlight = false;
      if (applyRequestState.queued) {
        var queued = applyRequestState.queued;
        applyRequestState.queued = null;
        appendDebug("UI> apply busy queue: running latest pending request");
        runApplyRequest(queued, 0);
      }
    });
  }

  function enqueueApplyRequest(request) {
    if (!request || !request.script) {
      return;
    }

    if (applyRequestState.inFlight) {
      applyRequestState.queued = request;
      appendDebug("UI> apply busy queue: updated pending request");
      return;
    }
    runApplyRequest(request, 0);
  }

  function parseJsonSafe(raw) {
    if (!raw) {
      return null;
    }
    try {
      return JSON.parse(raw);
    } catch (e1) {
      return null;
    }
  }

  // Build classic grid batch targets in row-major order (top-left to bottom-right).
  function buildClassicBatchCells() {
    var cells = [];
    for (var r = 0; r < state.rows; r += 1) {
      for (var c = 0; c < state.cols; c += 1) {
        cells.push({
          leftNorm: c / state.cols,
          topNorm: r / state.rows,
          widthNorm: 1 / state.cols,
          heightNorm: 1 / state.rows,
          label: "r" + (r + 1) + "c" + (c + 1)
        });
      }
    }
    return cells;
  }

  // Build designer batch targets using visual order: top-to-bottom then left-to-right.
  function buildDesignerBatchCells() {
    var blocks = (state.designer.blocks || []).slice();
    blocks.sort(function (a, b) {
      if (Math.abs(a.y - b.y) > 0.0001) {
        return a.y - b.y;
      }
      if (Math.abs(a.x - b.x) > 0.0001) {
        return a.x - b.x;
      }
      return String(a.id || "").localeCompare(String(b.id || ""));
    });

    var cells = [];
    for (var i = 0; i < blocks.length; i += 1) {
      var block = blocks[i];
      cells.push({
        leftNorm: block.x / DESIGNER_GRID_SIZE,
        topNorm: block.y / DESIGNER_GRID_SIZE,
        widthNorm: block.w / DESIGNER_GRID_SIZE,
        heightNorm: block.h / DESIGNER_GRID_SIZE,
        label: String(block.id || ("block_" + i))
      });
    }
    return cells;
  }

  // Resolve the current mode targets for batch placement.
  function buildBatchCells() {
    if (state.designer.enabled) {
      return buildDesignerBatchCells();
    }
    return buildClassicBatchCells();
  }

  // Batch apply maps selected clips to cells, ordered by track (bottom->top) in hostscript.
  function applyBatchSelection() {
    var cells = buildBatchCells();
    if (!cells.length) {
      setStatusKey("status.batch_no_cells", {}, "err");
      return;
    }

    var script = "gridMaker_applyBatchToSelectedClips(" +
      quoteForEvalScript(JSON.stringify(cells)) + "," +
      state.ratioW + "," +
      state.ratioH + "," +
      state.marginPx +
      ")";

    appendDebug("UI> batch click cells=" + cells.length + " ratio=" + getRatioText() + " marginPx=" + state.marginPx);
    appendDebug("UI> evalScript: gridMaker_applyBatchToSelectedClips(<cells>," + state.ratioW + "," + state.ratioH + "," + state.marginPx + ")");
    setStatusKey("status.batch_applying", {}, "");

    enqueueApplyRequest({
      script: script,
      onSuccess: function (parsed) {
        var applied = (parsed.details && parsed.details.applied) ? parsed.details.applied : "0";
        var total = (parsed.details && parsed.details.total) ? parsed.details.total : "0";
        var failed = (parsed.details && parsed.details.failed) ? parsed.details.failed : "0";
        var skipped = (parsed.details && parsed.details.skipped) ? parsed.details.skipped : "0";
        appendDebug("UI> batch applied=" + applied + "/" + total + " failed=" + failed + " skipped=" + skipped);
      },
      onDone: function (parsed) {
        if (parsed.raw) {
          setStatusRaw(parsed.raw, parsed.kind);
          return;
        }
        setStatusKey(parsed.key, parsed.vars, parsed.kind);
      }
    });
  }

  // Import designer configs from JSON, then refresh gallery for the current ratio.
  function importDesignerConfigs() {
    appendDebug("UI> evalScript: gridMaker_designerImportConfigs()");
    setStatusKey("status.designer_importing", {}, "");

    callHost("gridMaker_designerImportConfigs()", function (result) {
      appendDebug("HOST< raw(designer-import): " + (result || "<empty>"));
      var parsed = parseJsonSafe(result);
      if (!parsed || !parsed.ok) {
        if (parsed && parsed.cancelled) {
          setStatusKey("status.designer_io_cancelled", {}, "");
          return;
        }
        setStatusKey("status.designer_import_failed", {}, "err");
        return;
      }

      setStatusKey("status.designer_imported", { count: parsed.count || 0 }, "ok");
      loadDesignerConfigs(state.designer.activeConfigId, true);
    });
  }

  // Export all designer configs to a JSON file for backup/team sharing.
  function exportDesignerConfigs() {
    appendDebug("UI> evalScript: gridMaker_designerExportConfigs()");
    setStatusKey("status.designer_exporting", {}, "");

    callHost("gridMaker_designerExportConfigs()", function (result) {
      appendDebug("HOST< raw(designer-export): " + (result || "<empty>"));
      var parsed = parseJsonSafe(result);
      if (!parsed || !parsed.ok) {
        if (parsed && parsed.cancelled) {
          setStatusKey("status.designer_io_cancelled", {}, "");
          return;
        }
        setStatusKey("status.designer_export_failed", {}, "err");
        return;
      }

      setStatusKey("status.designer_exported", { count: parsed.count || 0 }, "ok");
    });
  }

  // User actions: apply classic cell or designer custom block to timeline selection.
  function applyClassicCell(row, col) {
    state.selectedCell = { row: row, col: col };
    renderPreview();

    var script = "gridMaker_applyToSelectedClip(" +
      row + "," +
      col + "," +
      state.rows + "," +
      state.cols + "," +
      state.ratioW + "," +
      state.ratioH + "," +
      state.marginPx +
      ")";

    appendDebug("UI> click cell row=" + (row + 1) + " col=" + (col + 1) + " grid=" + state.rows + "x" + state.cols + " ratio=" + getRatioText() + " marginPx=" + state.marginPx);
    appendDebug("UI> evalScript: " + script);
    setStatusKey("status.applying", {}, "");

    enqueueApplyRequest({
      script: script,
      onSuccess: function (parsed) {
        if (parsed.code === "cell_applied" && parsed.details) {
          appendDebug(
            "UI> applied cell row=" + parsed.details.row +
            " col=" + parsed.details.col +
            " scale=" + parsed.details.scale + "%"
          );
        }
      },
      onDone: function (parsed) {
        if (parsed.raw) {
          setStatusRaw(parsed.raw, parsed.kind);
          return;
        }
        setStatusKey(parsed.key, parsed.vars, parsed.kind);
      }
    });
  }

  function applyDesignerBlock(blockId) {
    var block = findDesignerBlockById(blockId);
    if (!block) {
      setStatusKey("status.err.unknown", {}, "err");
      return;
    }

    state.designer.selectedBlockId = blockId;
    renderPreview();

    var leftNorm = formatPercent(block.x / DESIGNER_GRID_SIZE);
    var topNorm = formatPercent(block.y / DESIGNER_GRID_SIZE);
    var widthNorm = formatPercent(block.w / DESIGNER_GRID_SIZE);
    var heightNorm = formatPercent(block.h / DESIGNER_GRID_SIZE);

    var script = "gridMaker_applyToSelectedCustomCell(" +
      leftNorm + "," +
      topNorm + "," +
      widthNorm + "," +
      heightNorm + "," +
      state.ratioW + "," +
      state.ratioH + "," +
      state.marginPx +
      ")";

    appendDebug("UI> click designer cell id=" + block.id + " x=" + block.x + " y=" + block.y + " w=" + block.w + " h=" + block.h + " ratio=" + getRatioText() + " marginPx=" + state.marginPx);
    appendDebug("UI> evalScript: " + script);
    setStatusKey("status.applying", {}, "");

    enqueueApplyRequest({
      script: script,
      onSuccess: function (parsed) {
        if (parsed.code === "cell_applied") {
          appendDebug("UI> applied designer cell id=" + block.id + " scale=" + ((parsed.details && parsed.details.scale) || "") + "%");
        }
      },
      onDone: function (parsed) {
        if (parsed.raw) {
          setStatusRaw(parsed.raw, parsed.kind);
          return;
        }
        setStatusKey(parsed.key, parsed.vars, parsed.kind);
      }
    });
  }

  // Drag/resize interactions used in Designer edit mode.
  function startDesignerDrag(event, blockId, action) {
    if (!state.designer.enabled || !state.designer.editMode) {
      return;
    }
    if (!event || event.button !== 0) {
      return;
    }

    var block = findDesignerBlockById(blockId);
    if (!block) {
      return;
    }

    var rect = grid.getBoundingClientRect();
    if (!(rect.width > 0) || !(rect.height > 0)) {
      return;
    }

    state.designer.selectedBlockId = blockId;
    designerBringBlockToFront(blockId);
    designerDrag = {
      id: blockId,
      action: action,
      startClientX: event.clientX,
      startClientY: event.clientY,
      stageRect: rect,
      startBlock: {
        id: block.id,
        x: block.x,
        y: block.y,
        w: block.w,
        h: block.h
      }
    };

    document.body.classList.add("designer-dragging");
    window.addEventListener("mousemove", onDesignerDragMove);
    window.addEventListener("mouseup", onDesignerDragEnd);
    event.preventDefault();
    event.stopPropagation();
    renderPreview();
  }

  function onDesignerDragMove(event) {
    if (!designerDrag) {
      return;
    }

    var base = designerDrag.startBlock;
    var rect = designerDrag.stageRect;
    var dragStep = getDesignerStep();
    var gridUnits = getDesignerGridUnits();
    var stepX = rect.width / gridUnits;
    var stepY = rect.height / gridUnits;
    if (!(stepX > 0) || !(stepY > 0)) {
      return;
    }

    var dxCells = Math.round((event.clientX - designerDrag.startClientX) / stepX) * dragStep;
    var dyCells = Math.round((event.clientY - designerDrag.startClientY) / stepY) * dragStep;
    var candidate = {
      id: base.id,
      x: base.x,
      y: base.y,
      w: base.w,
      h: base.h
    };

    if (designerDrag.action === "move") {
      candidate.x = clampStep(base.x + dxCells, 0, DESIGNER_GRID_SIZE - base.w, dragStep, base.x);
      candidate.y = clampStep(base.y + dyCells, 0, DESIGNER_GRID_SIZE - base.h, dragStep, base.y);
    } else {
      candidate.w = clampStep(base.w + dxCells, dragStep, DESIGNER_GRID_SIZE - base.x, dragStep, base.w);
      candidate.h = clampStep(base.h + dyCells, dragStep, DESIGNER_GRID_SIZE - base.y, dragStep, base.h);
    }

    if (!designerCanPlace(candidate, candidate.id)) {
      return;
    }

    var live = findDesignerBlockById(candidate.id);
    if (!live) {
      return;
    }

    if (live.x === candidate.x && live.y === candidate.y && live.w === candidate.w && live.h === candidate.h) {
      return;
    }

    live.x = candidate.x;
    live.y = candidate.y;
    live.w = candidate.w;
    live.h = candidate.h;
    renderPreview();
  }

  function onDesignerDragEnd() {
    if (!designerDrag) {
      return;
    }
    designerDrag = null;
    document.body.classList.remove("designer-dragging");
    window.removeEventListener("mousemove", onDesignerDragMove);
    window.removeEventListener("mouseup", onDesignerDragEnd);
  }

  function stopDesignerDrag() {
    if (designerDrag) {
      onDesignerDragEnd();
    }
  }

  // Render pipeline for classic grid, designer grid and shared panel state.
  function renderClassicGrid() {
    grid.classList.remove("designer-mode");
    grid.innerHTML = "";
    grid.style.gridTemplateRows = "repeat(" + state.rows + ", 1fr)";
    grid.style.gridTemplateColumns = "repeat(" + state.cols + ", 1fr)";
    grid.classList.remove("classic-compact");

    for (var r = 0; r < state.rows; r += 1) {
      for (var c = 0; c < state.cols; c += 1) {
        var cell = document.createElement("button");
        cell.type = "button";
        cell.className = "cell";

        if (state.selectedCell.row === r && state.selectedCell.col === c) {
          cell.className += " active";
        }

        cell.textContent = t("cell.label", { row: r + 1, col: c + 1 });

        (function (row, col) {
          cell.addEventListener("click", function () {
            applyClassicCell(row, col);
          });
        })(r, c);

        grid.appendChild(cell);
      }
    }
    updateClassicGridDensity();
  }

  function renderDesignerGrid() {
    ensureDesignerSelection();
    grid.classList.add("designer-mode");
    grid.innerHTML = "";
    grid.style.gridTemplateRows = "";
    grid.style.gridTemplateColumns = "";

    if (!state.designer.blocks.length) {
      adoptDesignerBlocks(makeDefaultDesignerBlocks());
    }

    var blocks = state.designer.blocks.slice();

    for (var i = 0; i < blocks.length; i += 1) {
      (function (block, index) {
        var cell = document.createElement("button");
        cell.type = "button";
        cell.className = "designer-cell";
        if (state.designer.selectedBlockId === block.id) {
          cell.className += " active";
        }
        if (state.designer.editMode) {
          cell.className += " editable";
        }

        cell.style.left = (block.x * 100 / DESIGNER_GRID_SIZE) + "%";
        cell.style.top = (block.y * 100 / DESIGNER_GRID_SIZE) + "%";
        cell.style.width = (block.w * 100 / DESIGNER_GRID_SIZE) + "%";
        cell.style.height = (block.h * 100 / DESIGNER_GRID_SIZE) + "%";
        cell.style.zIndex = state.designer.selectedBlockId === block.id
          ? "1000"
          : String(10 + index);
        cell.setAttribute("data-id", block.id);

        var overlapCount = designerOverlapCountForBlock(block);
        if (overlapCount > 0) {
          cell.className += " overlap";
          var overlapRegions = designerOverlapRegionsForBlock(block);
          for (var r = 0; r < overlapRegions.length; r += 1) {
            var region = overlapRegions[r];
            var overlapRegion = document.createElement("span");
            overlapRegion.className = "designer-overlap-region";
            overlapRegion.style.left = region.left + "%";
            overlapRegion.style.top = region.top + "%";
            overlapRegion.style.width = region.width + "%";
            overlapRegion.style.height = region.height + "%";
            cell.appendChild(overlapRegion);
          }

          var overlapBadge = document.createElement("span");
          overlapBadge.className = "designer-overlap-badge";
          overlapBadge.textContent = "+" + String(overlapCount);
          cell.appendChild(overlapBadge);
        }

        var label = document.createElement("span");
        label.className = "designer-cell-label";
        label.textContent = String(index + 1);
        cell.appendChild(label);

        var widthLabel = document.createElement("span");
        widthLabel.className = "designer-cell-size designer-cell-size-x";
        widthLabel.textContent = formatDesignerSizePercent(block.w);
        cell.appendChild(widthLabel);

        var heightLabel = document.createElement("span");
        heightLabel.className = "designer-cell-size designer-cell-size-y";
        heightLabel.textContent = formatDesignerSizePercent(block.h);
        cell.appendChild(heightLabel);

        if (state.designer.editMode) {
          var handle = document.createElement("span");
          handle.className = "designer-resize-handle";
          handle.title = t("designer.resize_hint");
          cell.appendChild(handle);

          cell.addEventListener("mousedown", function (event) {
            var target = event.target;
            var action = (target && target.className && String(target.className).indexOf("designer-resize-handle") !== -1)
              ? "resize"
              : "move";
            startDesignerDrag(event, block.id, action);
          });
        }

        cell.addEventListener("click", function (event) {
          event.stopPropagation();
          designerBringBlockToFront(block.id);
          state.designer.selectedBlockId = block.id;
          if (state.designer.editMode) {
            renderPreview();
            return;
          }
          applyDesignerBlock(block.id);
        });

        grid.appendChild(cell);
      })(blocks[i], i);
    }
  }

  function renderPreview() {
    updateSummary();
    if (state.designer.enabled) {
      renderDesignerGrid();
    } else {
      renderClassicGrid();
    }
    renderDesignerControlsState();
    renderDesignerGallery();
    schedulePreviewFit();
  }

  function renderDesignerControlsState() {
    if (classicGridControls) {
      classicGridControls.hidden = !!state.designer.enabled;
      classicGridControls.style.display = state.designer.enabled ? "none" : "";
    }
    if (designerControls) {
      designerControls.hidden = !state.designer.enabled;
      designerControls.style.display = state.designer.enabled ? "" : "none";
    }
    if (designerGalleryPanel) {
      designerGalleryPanel.hidden = !state.designer.enabled;
      designerGalleryPanel.style.display = state.designer.enabled ? "" : "none";
    }

    if (!designerEditBtn) {
      return;
    }

    designerEditBtn.textContent = state.designer.editMode
      ? t("designer.edit_on")
      : t("designer.edit_off");
    designerEditBtn.classList.toggle("edit-on", !!state.designer.editMode);
    designerEditBtn.classList.toggle("edit-off", !state.designer.editMode);
    if (designerFreeBtn) {
      designerFreeBtn.textContent = state.designer.freeMode
        ? t("designer.free_on")
        : t("designer.free_off");
      designerFreeBtn.classList.toggle("free-on", !!state.designer.freeMode);
      designerFreeBtn.classList.toggle("free-off", !state.designer.freeMode);
      designerFreeBtn.disabled = !state.designer.enabled || !state.designer.editMode;
    }

    if (designerModeBtn) {
      designerModeBtn.classList.toggle("active", !!state.designer.enabled);
      designerModeBtn.textContent = state.designer.enabled
        ? t("designer.mode_on")
        : t("designer.mode_off");
    }

    if (helpText) {
      helpText.textContent = state.designer.enabled
        ? t("help.designer")
        : t("help.cellClick");
    }
  }

  function renderDesignerGallery() {
    if (!designerGallery) {
      return;
    }

    if (!state.designer.enabled) {
      designerGallery.innerHTML = "";
      return;
    }

    designerGallery.innerHTML = "";
    var configs = state.designer.configs || [];

    if (!configs.length) {
      var empty = document.createElement("div");
      empty.className = "designer-gallery-empty";
      empty.textContent = t("designer.gallery_empty");
      designerGallery.appendChild(empty);
      return;
    }

    for (var i = 0; i < configs.length; i += 1) {
      (function (cfg) {
        var item = document.createElement("button");
        item.type = "button";
        item.className = "designer-gallery-item";
        if (cfg.id === state.designer.activeConfigId) {
          item.className += " active";
        }

        var thumb = document.createElement("div");
        thumb.className = "designer-thumb";
        var thumbRatioW = Number(cfg.ratioW);
        var thumbRatioH = Number(cfg.ratioH);
        if (!(thumbRatioW > 0) || !(thumbRatioH > 0)) {
          thumbRatioW = state.ratioW;
          thumbRatioH = state.ratioH;
        }
        thumb.style.aspectRatio = thumbRatioW + " / " + thumbRatioH;
        var blocks = cloneDesignerBlocks(cfg.blocks || []);
        for (var bi = 0; bi < blocks.length; bi += 1) {
          var block = blocks[bi];
          var piece = document.createElement("span");
          piece.className = "designer-thumb-block";
          piece.style.left = (block.x * 100 / DESIGNER_GRID_SIZE) + "%";
          piece.style.top = (block.y * 100 / DESIGNER_GRID_SIZE) + "%";
          piece.style.width = (block.w * 100 / DESIGNER_GRID_SIZE) + "%";
          piece.style.height = (block.h * 100 / DESIGNER_GRID_SIZE) + "%";
          thumb.appendChild(piece);
        }

        var meta = document.createElement("div");
        meta.className = "designer-gallery-meta";

        var name = document.createElement("strong");
        name.textContent = cfg.name || cfg.id;

        meta.appendChild(name);

        item.appendChild(thumb);
        item.appendChild(meta);

        item.addEventListener("click", function () {
          loadDesignerConfig(cfg.id);
        });

        item.addEventListener("contextmenu", function (event) {
          event.preventDefault();
          deleteDesignerConfig(cfg.id, cfg.name || cfg.id);
        });

        designerGallery.appendChild(item);
      })(configs[i]);
    }
  }

  // Designer config CRUD (load/save/delete) persisted by hostscript per ratio.
  function setDesignerMode(enabled) {
    var next = !!enabled;
    if (state.designer.enabled === next) {
      return;
    }

    state.designer.enabled = next;
    if (!next) {
      state.designer.freeMode = false;
    }
    stopDesignerDrag();

    if (classicGridControls) {
      classicGridControls.hidden = next;
    }
    if (designerControls) {
      designerControls.hidden = !next;
    }
    if (designerGalleryPanel) {
      designerGalleryPanel.hidden = !next;
    }

    appendDebug("UI> designer mode " + (next ? "enabled" : "disabled"));

    if (next) {
      if (!state.designer.blocks.length) {
        adoptDesignerBlocks(makeDefaultDesignerBlocks());
      }
      loadDesignerConfigs(state.designer.activeConfigId, true);
      setStatusKey("status.designer_ready", {}, "ok");
    } else {
      setStatusKey("status.ready", {}, "");
    }

    renderPreview();
  }

  function loadDesignerConfig(configId) {
    var list = state.designer.configs || [];
    for (var i = 0; i < list.length; i += 1) {
      if (list[i].id === configId) {
        state.designer.activeConfigId = list[i].id;
        adoptDesignerBlocks(list[i].blocks || []);
        if (designerNameInput) {
          designerNameInput.value = list[i].name || "";
        }
        appendDebug("UI> loaded designer config id=" + list[i].id + " name=" + (list[i].name || ""));
        setStatusKey("status.designer_config_loaded", { name: list[i].name || list[i].id }, "ok");
        renderPreview();
        return;
      }
    }
  }

  function adoptDesignerConfigList(configs, preferredId, autoPick) {
    state.designer.configs = configs || [];

    if (!state.designer.configs.length) {
      state.designer.activeConfigId = "";
      if (designerNameInput && !designerNameInput.value) {
        designerNameInput.value = t("designer.default_name");
      }
      renderPreview();
      return;
    }

    var targetId = preferredId || state.designer.activeConfigId;
    var found = null;

    if (targetId) {
      for (var i = 0; i < state.designer.configs.length; i += 1) {
        if (state.designer.configs[i].id === targetId) {
          found = state.designer.configs[i];
          break;
        }
      }
    }

    if (!found && autoPick) {
      found = state.designer.configs[0];
    }

    if (found) {
      state.designer.activeConfigId = found.id;
      adoptDesignerBlocks(found.blocks || []);
      if (designerNameInput) {
        designerNameInput.value = found.name || "";
      }
    }

    renderPreview();
  }

  function loadDesignerConfigs(preferredId, autoPick) {
    var script = "gridMaker_designerListConfigs(" + state.ratioW + "," + state.ratioH + ")";
    appendDebug("UI> evalScript: " + script);

    callHost(script, function (result) {
      appendDebug("HOST< raw(designer-list): " + (result || "<empty>"));
      var payload = parseJsonSafe(result);
      if (!payload || !payload.ok) {
        appendDebug("UI> designer list failed");
        setStatusKey("status.designer_load_failed", {}, "err");
        return;
      }

      var configs = [];
      var rawConfigs = payload.configs || [];
      for (var i = 0; i < rawConfigs.length; i += 1) {
        var cfg = rawConfigs[i] || {};
        configs.push({
          id: String(cfg.id || ("cfg_" + i)),
          name: String(cfg.name || cfg.id || ""),
          ratioW: cfg.ratioW,
          ratioH: cfg.ratioH,
          blocks: cloneDesignerBlocks(cfg.blocks || []),
          updatedAt: String(cfg.updatedAt || "")
        });
      }

      state.designer.loaded = true;
      appendDebug("UI> designer configs loaded count=" + configs.length + " ratio=" + getRatioText());
      adoptDesignerConfigList(configs, preferredId, !!autoPick);
    });
  }

  function saveDesignerConfig() {
    if (!state.designer.blocks.length) {
      setStatusKey("status.designer_empty", {}, "err");
      return;
    }

    var name = (designerNameInput && designerNameInput.value) ? designerNameInput.value.trim() : "";
    if (!name) {
      name = t("designer.default_name");
      if (designerNameInput) {
        designerNameInput.value = name;
      }
    }

    var payload = {
      id: state.designer.activeConfigId || "",
      name: name,
      ratioW: state.ratioW,
      ratioH: state.ratioH,
      blocks: cloneDesignerBlocks(state.designer.blocks)
    };

    var script = "gridMaker_designerSaveConfig(" + quoteForEvalScript(JSON.stringify(payload)) + ")";
    appendDebug("UI> evalScript: gridMaker_designerSaveConfig(<payload>)");
    setStatusKey("status.designer_saving", {}, "");

    callHost(script, function (result) {
      appendDebug("HOST< raw(designer-save): " + (result || "<empty>"));
      var parsed = parseJsonSafe(result);
      if (!parsed || !parsed.ok || !parsed.id) {
        appendDebug("UI> designer save failed");
        setStatusKey("status.designer_save_failed", {}, "err");
        return;
      }

      state.designer.activeConfigId = String(parsed.id);
      appendDebug("UI> designer config saved id=" + state.designer.activeConfigId);
      setStatusKey("status.designer_saved", { name: name }, "ok");
      loadDesignerConfigs(state.designer.activeConfigId, true);
    });
  }

  function deleteDesignerConfig(configId, displayName) {
    if (!configId) {
      return;
    }

    var question = t("designer.confirm_delete", { name: displayName || configId });
    if (!window.confirm(question)) {
      return;
    }

    var script = "gridMaker_designerDeleteConfig(" + quoteForEvalScript(configId) + ")";
    appendDebug("UI> evalScript: " + script);

    callHost(script, function (result) {
      appendDebug("HOST< raw(designer-delete): " + (result || "<empty>"));
      var parsed = parseJsonSafe(result);
      if (!parsed || !parsed.ok) {
        appendDebug("UI> designer delete failed");
        setStatusKey("status.designer_delete_failed", {}, "err");
        return;
      }

      appendDebug("UI> designer config deleted id=" + configId);
      if (state.designer.activeConfigId === configId) {
        state.designer.activeConfigId = "";
        adoptDesignerBlocks(makeDefaultDesignerBlocks());
        if (designerNameInput) {
          designerNameInput.value = t("designer.default_name");
        }
      }
      setStatusKey("status.designer_deleted", {}, "ok");
      loadDesignerConfigs(state.designer.activeConfigId, true);
    });
  }

  function createNewDesignerConfig() {
    state.designer.activeConfigId = "";
    state.designer.editMode = true;
    state.designer.freeMode = false;
    stopDesignerDrag();
    adoptDesignerBlocks(makeDefaultDesignerBlocks());
    if (designerNameInput) {
      designerNameInput.value = t("designer.default_name");
    }
    appendDebug("UI> new designer config draft");
    setStatusKey("status.designer_new", {}, "ok");
    renderPreview();
  }

  function findFirstDesignerFreePosition(width, height) {
    for (var y = 0; y <= DESIGNER_GRID_SIZE - height; y += 1) {
      for (var x = 0; x <= DESIGNER_GRID_SIZE - width; x += 1) {
        var candidate = { id: "", x: x, y: y, w: width, h: height };
        if (designerCanPlace(candidate, "", false)) {
          return { x: x, y: y, w: width, h: height };
        }
      }
    }
    return null;
  }

  function addDesignerBlock() {
    var preferredSizes = [
      { w: 2, h: 2 },
      { w: 2, h: 1 },
      { w: 1, h: 2 },
      { w: 1, h: 1 }
    ];

    var slot = null;
    for (var i = 0; i < preferredSizes.length; i += 1) {
      slot = findFirstDesignerFreePosition(preferredSizes[i].w, preferredSizes[i].h);
      if (slot) {
        break;
      }
    }

    if (!slot) {
      setStatusKey("status.designer_no_space", {}, "err");
      return;
    }

    var block = {
      id: nextDesignerBlockId(),
      x: slot.x,
      y: slot.y,
      w: slot.w,
      h: slot.h
    };

    state.designer.blocks.push(block);
    state.designer.selectedBlockId = block.id;
    appendDebug("UI> designer block added id=" + block.id + " x=" + block.x + " y=" + block.y + " w=" + block.w + " h=" + block.h);
    setStatusKey("status.designer_block_added", {}, "ok");
    renderPreview();
  }

  function removeSelectedDesignerBlock() {
    var selectedId = state.designer.selectedBlockId;
    if (!selectedId) {
      setStatusKey("status.designer_block_select", {}, "err");
      return;
    }

    var next = [];
    var removed = false;
    for (var i = 0; i < state.designer.blocks.length; i += 1) {
      if (state.designer.blocks[i].id === selectedId) {
        removed = true;
        continue;
      }
      next.push(state.designer.blocks[i]);
    }

    if (!removed) {
      setStatusKey("status.designer_block_select", {}, "err");
      return;
    }

    state.designer.blocks = next;
    ensureDesignerSelection();
    appendDebug("UI> designer block removed id=" + selectedId);
    setStatusKey("status.designer_block_removed", {}, "ok");
    renderPreview();
  }

  function toggleDesignerEditMode() {
    state.designer.editMode = !state.designer.editMode;
    if (!state.designer.editMode) {
      state.designer.freeMode = false;
    }
    stopDesignerDrag();
    appendDebug("UI> designer edit mode " + (state.designer.editMode ? "ON" : "OFF"));
    setStatusKey(state.designer.editMode ? "status.designer_edit_on" : "status.designer_edit_off", {}, "ok");
    renderPreview();
  }

  function toggleDesignerFreeMode() {
    if (!state.designer.enabled || !state.designer.editMode) {
      return;
    }
    state.designer.freeMode = !state.designer.freeMode;
    stopDesignerDrag();
    appendDebug("UI> designer free mode " + (state.designer.freeMode ? "ON" : "OFF"));
    setStatusKey(state.designer.freeMode ? "status.designer_free_on" : "status.designer_free_off", {}, "ok");
    renderPreview();
  }

  // Static text and language selector rendering.
  function renderStaticTexts() {
    var nodes = document.querySelectorAll("[data-i18n]");
    for (var i = 0; i < nodes.length; i += 1) {
      var node = nodes[i];
      var key = node.getAttribute("data-i18n");
      node.textContent = t(key);
    }

    var placeholders = document.querySelectorAll("[data-i18n-placeholder]");
    for (var j = 0; j < placeholders.length; j += 1) {
      var input = placeholders[j];
      var pkey = input.getAttribute("data-i18n-placeholder");
      input.setAttribute("placeholder", t(pkey));
    }
  }

  function renderLanguageOptions() {
    languageSelect.innerHTML = "";

    var localeCodes = Object.keys(i18n.locales);
    localeCodes.sort(function (a, b) {
      if (a === i18n.defaultLocale) {
        return -1;
      }
      if (b === i18n.defaultLocale) {
        return 1;
      }
      return a.localeCompare(b);
    });

    for (var i = 0; i < localeCodes.length; i += 1) {
      var code = localeCodes[i];
      var locale = i18n.locales[code];
      var option = document.createElement("option");
      option.value = code;
      option.textContent = (locale.flag ? locale.flag + " " : "") + (locale.label || code);
      languageSelect.appendChild(option);
    }

    languageSelect.value = state.locale;
  }

  function setLocale(nextLocale) {
    if (!i18n.locales[nextLocale]) {
      return;
    }

    state.locale = nextLocale;
    document.documentElement.lang = nextLocale;

    try {
      window.localStorage.setItem("pgm.locale", nextLocale);
    } catch (e1) {
      // Ignore storage issues in restricted CEP hosts.
    }

    renderStaticTexts();
    renderPreview();
    refreshStatus();
    refreshUpdateBanner();
  }

  // Event wiring: input sync, panel toggles, designer commands and clipboard copy.
  syncValue(rowsRange, rowsNumber, function (v) {
    state.rows = v;
    if (state.selectedCell.row >= v) {
      state.selectedCell.row = v - 1;
    }
    renderPreview();
  });

  syncValue(colsRange, colsNumber, function (v) {
    state.cols = v;
    if (state.selectedCell.col >= v) {
      state.selectedCell.col = v - 1;
    }
    renderPreview();
  });

  ratio.addEventListener("change", function () {
    var next = parseRatio(ratio.value);
    state.ratioW = next.w;
    state.ratioH = next.h;
    appendDebug("UI> ratio changed to " + ratio.value);
    if (state.designer.enabled) {
      loadDesignerConfigs("", true);
    }
    renderPreview();
  });

  if (marginRange) {
    marginRange.addEventListener("input", function () {
      applyGlobalMarginPx(marginRange.value, true);
      appendDebug("UI> margin changed to " + state.marginPx + "px");
    });
  }

  window.addEventListener("resize", function () {
    schedulePreviewFit();
  });

  if (debugPanel) {
    debugPanel.addEventListener("toggle", function () {
      schedulePreviewFit();
    });
  }

  if (designerGalleryPanel) {
    designerGalleryPanel.addEventListener("toggle", function () {
      schedulePreviewFit();
    });
  }

  if (designerGalleryTools) {
    ["click", "mousedown", "pointerdown", "touchstart"].forEach(function (evtName) {
      designerGalleryTools.addEventListener(evtName, function (event) {
        event.stopPropagation();
      });
    });
  }

  languageSelect.addEventListener("change", function () {
    appendDebug("UI> locale changed to " + languageSelect.value);
    setLocale(languageSelect.value);
  });

  rowsNumber.addEventListener("change", function () {
    appendDebug("UI> rows changed to " + rowsNumber.value);
  });

  colsNumber.addEventListener("change", function () {
    appendDebug("UI> cols changed to " + colsNumber.value);
  });

  if (grid) {
    grid.addEventListener("click", function (event) {
      if (!state.designer.enabled || !state.designer.editMode) {
        return;
      }
      if (event.target === grid) {
        state.designer.selectedBlockId = "";
        renderPreview();
      }
    });
  }

  if (designerModeBtn) {
    designerModeBtn.addEventListener("click", function () {
      setDesignerMode(!state.designer.enabled);
    });
  }

  if (designerEditBtn) {
    designerEditBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      toggleDesignerEditMode();
    });
  }

  if (designerFreeBtn) {
    designerFreeBtn.addEventListener("click", function () {
      if (!state.designer.enabled || !state.designer.editMode) {
        return;
      }
      toggleDesignerFreeMode();
    });
  }

  if (designerAddBtn) {
    designerAddBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      addDesignerBlock();
    });
  }

  if (designerRemoveBtn) {
    designerRemoveBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      removeSelectedDesignerBlock();
    });
  }

  if (designerNewBtn) {
    designerNewBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      createNewDesignerConfig();
    });
  }

  if (designerSaveBtn) {
    designerSaveBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      saveDesignerConfig();
    });
  }

  if (designerImportBtn) {
    designerImportBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      importDesignerConfigs();
    });
  }

  if (designerExportBtn) {
    designerExportBtn.addEventListener("click", function () {
      if (!state.designer.enabled) {
        return;
      }
      exportDesignerConfigs();
    });
  }

  if (designerNameInput) {
    designerNameInput.addEventListener("input", function () {
      updateSummary();
    });
  }

  if (designerGallerySize) {
    designerGallerySize.addEventListener("input", function () {
      applyDesignerGallerySize(designerGallerySize.value, true);
      appendDebug("UI> designer gallery size changed to " + designerGallerySize.value);
    });
  }

  if (applyBatchBtn) {
    applyBatchBtn.addEventListener("click", function () {
      applyBatchSelection();
    });
  }

  copyDebugBtn.addEventListener("click", function () {
    var text = debugLog ? debugLog.value : "";
    if (!text) {
      appendDebug("UI> copy debug requested (log empty)");
      setStatusKey("status.ok.debug_copied", {}, "ok");
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () {
        appendDebug("UI> debug copied to clipboard");
        setStatusKey("status.ok.debug_copied", {}, "ok");
      }).catch(function (err) {
        appendDebug("UI> clipboard API failed: " + err);
        try {
          debugLog.focus();
          debugLog.select();
          var ok = document.execCommand("copy");
          if (ok) {
            appendDebug("UI> debug copied via execCommand");
            setStatusKey("status.ok.debug_copied", {}, "ok");
          } else {
            setStatusKey("status.err.copy_failed", {}, "err");
          }
        } catch (e1) {
          setStatusKey("status.err.copy_failed", {}, "err");
        }
      });
      return;
    }

    try {
      debugLog.focus();
      debugLog.select();
      var copied = document.execCommand("copy");
      if (copied) {
        appendDebug("UI> debug copied via legacy clipboard");
        setStatusKey("status.ok.debug_copied", {}, "ok");
      } else {
        setStatusKey("status.err.copy_failed", {}, "err");
      }
    } catch (e2) {
      setStatusKey("status.err.copy_failed", {}, "err");
    }
  });

  if (appVersion) {
    appVersion.textContent = "V" + APP_VERSION;
  }

  if (updateLink) {
    updateLink.addEventListener("click", function (event) {
      var url = updateState.downloadUrl || updateLink.href || "";
      appendDebug("UI> update download clicked: " + url);
      if (!url || !isTrustedReleaseZipUrl(url)) {
        event.preventDefault();
        appendDebug("UI> click blocked (untrusted/empty URL)");
        return;
      }

      updateLink.href = url;

      openExternalUrl(url);
      appendDebug("UI> native anchor fallback remains enabled");
    });
  }

  // Initial boot sequence for UI state, i18n, preview and update check.
  applyDesignerGallerySize(state.designer.gallerySize, false);
  applyGlobalMarginPx(state.marginPx, false);
  if (classicGridControls) {
    classicGridControls.style.display = "";
  }
  if (designerControls) {
    designerControls.style.display = "none";
  }
  if (designerGalleryPanel) {
    designerGalleryPanel.style.display = "none";
  }

  adoptDesignerBlocks(makeDefaultDesignerBlocks());
  if (designerNameInput) {
    designerNameInput.value = t("designer.default_name");
  }

  renderLanguageOptions();
  setLocale(state.locale);
  setStatusKey("status.ready", {}, "");
  appendDebug("INIT> panel ready");
  checkForUpdates();
})();
