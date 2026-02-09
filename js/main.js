(function () {
  "use strict";

  var cep = window.__adobe_cep__ || null;
  var csInterface = (typeof CSInterface !== "undefined") ? new CSInterface() : null;
  var i18n = window.PGM_I18N || { defaultLocale: "en", locales: {} };
  var APP_VERSION = "1.0.3";
  var RELEASE_API_URL = "https://api.github.com/repos/CyrilG93/PremiereGridMaker/releases/latest";

  var state = {
    rows: 2,
    cols: 2,
    ratioW: 16,
    ratioH: 9,
    selectedCell: { row: 0, col: 0 },
    locale: resolveInitialLocale()
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

  var rowsRange = document.getElementById("rows");
  var rowsNumber = document.getElementById("rowsNumber");
  var colsRange = document.getElementById("cols");
  var colsNumber = document.getElementById("colsNumber");
  var ratio = document.getElementById("ratio");
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

  function getLocaleStrings() {
    var locale = i18n.locales[state.locale] || i18n.locales[i18n.defaultLocale];
    return locale ? locale.strings : {};
  }

  function format(template, vars) {
    var safeVars = vars || {};
    return String(template).replace(/\{([a-zA-Z0-9_]+)\}/g, function (_, key) {
      return (safeVars[key] !== undefined) ? String(safeVars[key]) : "";
    });
  }

  function hasText(key) {
    var strings = getLocaleStrings();
    return strings[key] !== undefined;
  }

  function t(key, vars) {
    var strings = getLocaleStrings();
    var template = strings[key];
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
    return /^https:\/\/github\.com\/CyrilG93\/PremiereGridMaker\/releases\/download\/v[0-9]+\.[0-9]+\.[0-9]+\/[^?#]+\.zip$/i.test(raw);
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
      return;
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
  }

  function openExternalUrlFallback(url) {
    try {
      if (csInterface && typeof csInterface.openURLInDefaultBrowser === "function") {
        csInterface.openURLInDefaultBrowser(url);
        appendDebug("UI> update URL opened via CSInterface");
        return;
      }
    } catch (e1) {}

    try {
      if (cep && cep.util && typeof cep.util.openURLInDefaultBrowser === "function") {
        cep.util.openURLInDefaultBrowser(url);
        appendDebug("UI> update URL opened via cep.util");
        return;
      }
    } catch (e2) {}

    appendDebug("UI> failed to open update URL");
  }

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

  function updateSummary() {
    var ratioText = state.ratioW + ":" + state.ratioH;
    summary.textContent = t("summary.format", {
      rows: state.rows,
      cols: state.cols,
      ratio: ratioText
    });
    grid.style.gridTemplateRows = "repeat(" + state.rows + ", 1fr)";
    grid.style.gridTemplateColumns = "repeat(" + state.cols + ", 1fr)";
    grid.style.aspectRatio = state.ratioW + " / " + state.ratioH;
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
  }

  function schedulePreviewFit() {
    if (typeof window.requestAnimationFrame === "function") {
      window.requestAnimationFrame(fitGridPreview);
      return;
    }
    window.setTimeout(fitGridPreview, 0);
  }

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
        return { kind: "ok", key: "status.ok.generic", vars: {}, hostDebug: hostDebug, code: code };
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

  function applyCell(row, col) {
    state.selectedCell = { row: row, col: col };
    renderGrid();

    var script = "gridMaker_applyToSelectedClip(" +
      row + "," +
      col + "," +
      state.rows + "," +
      state.cols + "," +
      state.ratioW + "," +
      state.ratioH +
      ")";

    appendDebug("UI> click cell row=" + (row + 1) + " col=" + (col + 1) + " grid=" + state.rows + "x" + state.cols + " ratio=" + state.ratioW + ":" + state.ratioH);
    appendDebug("UI> evalScript: " + script);
    setStatusKey("status.applying", {}, "");

    callHost(script, function (result) {
      appendDebug("HOST< raw: " + (result || "<empty>"));
      var parsed = parseHostResponse(result);
      appendHostDebug(parsed.hostDebug);

      if (parsed.kind === "ok" && parsed.code === "cell_applied" && parsed.details) {
        appendDebug(
          "UI> applied cell row=" + parsed.details.row +
          " col=" + parsed.details.col +
          " scale=" + parsed.details.scale + "%"
        );
      }

      if (parsed.raw) {
        setStatusRaw(parsed.raw, parsed.kind);
        return;
      }

      setStatusKey(parsed.key, parsed.vars, parsed.kind);
    });
  }

  function renderGrid() {
    grid.innerHTML = "";
    updateSummary();

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
            applyCell(row, col);
          });
        })(r, c);

        grid.appendChild(cell);
      }
    }
    schedulePreviewFit();
  }

  function renderStaticTexts() {
    var nodes = document.querySelectorAll("[data-i18n]");
    for (var i = 0; i < nodes.length; i += 1) {
      var node = nodes[i];
      var key = node.getAttribute("data-i18n");
      node.textContent = t(key);
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
    renderGrid();
    refreshStatus();
    refreshUpdateBanner();
  }

  syncValue(rowsRange, rowsNumber, function (v) {
    state.rows = v;
    if (state.selectedCell.row >= v) {
      state.selectedCell.row = v - 1;
    }
    renderGrid();
  });

  syncValue(colsRange, colsNumber, function (v) {
    state.cols = v;
    if (state.selectedCell.col >= v) {
      state.selectedCell.col = v - 1;
    }
    renderGrid();
  });

  ratio.addEventListener("change", function () {
    var next = parseRatio(ratio.value);
    state.ratioW = next.w;
    state.ratioH = next.h;
    renderGrid();
  });

  window.addEventListener("resize", function () {
    schedulePreviewFit();
  });

  if (debugPanel) {
    debugPanel.addEventListener("toggle", function () {
      schedulePreviewFit();
    });
  }

  languageSelect.addEventListener("change", function () {
    appendDebug("UI> locale changed to " + languageSelect.value);
    setLocale(languageSelect.value);
  });

  ratio.addEventListener("change", function () {
    appendDebug("UI> ratio changed to " + ratio.value);
  });

  rowsNumber.addEventListener("change", function () {
    appendDebug("UI> rows changed to " + rowsNumber.value);
  });

  colsNumber.addEventListener("change", function () {
    appendDebug("UI> cols changed to " + colsNumber.value);
  });

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
      event.preventDefault();
      var url = updateState.downloadUrl || updateLink.href || "";
      appendDebug("UI> update download clicked: " + url);
      openExternalUrl(url);
    });
  }

  renderLanguageOptions();
  setLocale(state.locale);
  setStatusKey("status.ready", {}, "");
  appendDebug("INIT> panel ready");
  checkForUpdates();
})();
