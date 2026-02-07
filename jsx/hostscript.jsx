function gridMaker_applyToSelectedClip(row, col, rows, cols, ratioW, ratioH) {
    try {
        app.enableQE();

        var seq = app.project.activeSequence;
        if (!seq) {
            return _gridMaker_result("ERR", "no_active_sequence");
        }

        row = parseInt(row, 10);
        col = parseInt(col, 10);
        rows = parseInt(rows, 10);
        cols = parseInt(cols, 10);
        ratioW = parseFloat(ratioW);
        ratioH = parseFloat(ratioH);

        if (rows < 1 || cols < 1) {
            return _gridMaker_result("ERR", "invalid_grid");
        }
        if (row < 0 || col < 0 || row >= rows || col >= cols) {
            return _gridMaker_result("ERR", "cell_out_of_bounds");
        }
        if (!(ratioW > 0) || !(ratioH > 0)) {
            return _gridMaker_result("ERR", "invalid_ratio");
        }

        var selection = seq.getSelection();
        if (!selection || selection.length < 1) {
            return _gridMaker_result("ERR", "no_selection");
        }

        var clip = null;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].mediaType === "Video") {
                clip = selection[i];
                break;
            }
        }
        if (!clip) {
            return _gridMaker_result("ERR", "no_video_selected");
        }

        var transformComp = _gridMaker_findTransformComponent(clip);
        var motionComp = _gridMaker_findMotionComponent(clip);
        var placementComp = motionComp || transformComp;
        var cropComp = _gridMaker_findCropComponent(clip);

        var qSeq = null;
        if (!placementComp || !cropComp) {
            qSeq = qe.project.getActiveSequence();
            if (!qSeq) {
                return _gridMaker_result("ERR", "qe_unavailable");
            }

            var qClip = _gridMaker_findQEClip(qSeq, seq, clip);
            if (!qClip) {
                return _gridMaker_result("ERR", "qe_clip_not_found");
            }

            if (!placementComp) {
                transformComp = _gridMaker_ensureEffect(clip, qClip, _gridMaker_transformEffectLookupNames(), _gridMaker_findTransformComponent);
                motionComp = _gridMaker_findMotionComponent(clip);
                placementComp = motionComp || transformComp;
            }
            if (!cropComp) {
                cropComp = _gridMaker_ensureEffect(clip, qClip, _gridMaker_cropEffectLookupNames(), _gridMaker_findCropComponent);
            }
        }

        if (!placementComp) {
            return _gridMaker_result("ERR", "transform_effect_unavailable");
        }
        if (!cropComp) {
            return _gridMaker_result("ERR", "crop_effect_unavailable");
        }

        var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
        if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
            return _gridMaker_result("ERR", "invalid_sequence_size");
        }
        var frameW = frameSize.width;
        var frameH = frameSize.height;
        var frameAspect = frameW / frameH;

        var cropL = 0.0;
        var cropR = 0.0;
        var cropT = 0.0;
        var cropB = 0.0;

        var targetAspect = ratioW / ratioH;
        if (frameAspect > targetAspect) {
            var visibleX = targetAspect / frameAspect;
            var lossX = 1.0 - visibleX;
            cropL += lossX * 0.5;
            cropR += lossX * 0.5;
        } else {
            var visibleY = frameAspect / targetAspect;
            var lossY = 1.0 - visibleY;
            cropT += lossY * 0.5;
            cropB += lossY * 0.5;
        }

        var visX = 1.0 - cropL - cropR;
        var visY = 1.0 - cropT - cropB;

        var neededVisRatio = rows / cols;
        var currentVisRatio = visX / visY;

        if (currentVisRatio > neededVisRatio) {
            var finalVisX = visY * neededVisRatio;
            var extraX = (visX - finalVisX) * 0.5;
            cropL += extraX;
            cropR += extraX;
            visX = finalVisX;
        } else {
            var finalVisY = visX / neededVisRatio;
            var extraY = (visY - finalVisY) * 0.5;
            cropT += extraY;
            cropB += extraY;
            visY = finalVisY;
        }

        cropL = _gridMaker_clamp(cropL * 100.0, 0, 49.5);
        cropR = _gridMaker_clamp(cropR * 100.0, 0, 49.5);
        cropT = _gridMaker_clamp(cropT * 100.0, 0, 49.5);
        cropB = _gridMaker_clamp(cropB * 100.0, 0, 49.5);

        visX = 1.0 - (cropL + cropR) / 100.0;
        var baseScale = _gridMaker_getCurrentScalePercent(placementComp);
        if (!(baseScale > 0)) {
            baseScale = 100.0;
        }
        var scale = baseScale / (cols * visX);

        var x = frameW * ((col + 0.5) / cols);
        var y = frameH * ((row + 0.5) / rows);

        if (!_gridMaker_setPlacement(placementComp, scale, x, y, frameW, frameH)) {
            return _gridMaker_result("ERR", "placement_apply_failed", {
                x: x.toFixed(3),
                y: y.toFixed(3)
            });
        }
        _gridMaker_setCrop(cropComp, cropL, cropR, cropT, cropB);

        return _gridMaker_result("OK", "cell_applied", {
            row: row + 1,
            col: col + 1,
            scale: scale.toFixed(2)
        });
    } catch (e) {
        return _gridMaker_result("ERR", "exception", { message: e });
    }
}

function _gridMaker_findQEClip(qSeq, seq, clip) {
    var targetStart = _gridMaker_timeToSeconds(clip.start);
    var targetEnd = _gridMaker_timeToSeconds(clip.end);
    var targetDuration = targetEnd - targetStart;
    var targetTrack = _gridMaker_findTrackIndex(seq, clip);
    var targetName = _gridMaker_clipName(clip);

    var bestItem = null;
    var bestScore = Number.MAX_VALUE;

    for (var vt = 0; vt < qSeq.numVideoTracks; vt++) {
        var track = qSeq.getVideoTrackAt(vt);
        if (!track) {
            continue;
        }

        for (var ci = 0; ci < track.numItems; ci++) {
            var item = track.getItemAt(ci);
            if (!item) {
                continue;
            }

            var score = _gridMaker_matchScore(item, vt, targetStart, targetEnd, targetDuration, targetTrack, targetName);
            if (score < bestScore) {
                bestScore = score;
                bestItem = item;
            }
        }
    }

    if (!bestItem) {
        return null;
    }

    return bestItem;
}

function _gridMaker_matchScore(item, itemTrackIndex, targetStart, targetEnd, targetDuration, targetTrack, targetName) {
    var itemStart = _gridMaker_timeToSeconds(item.start);
    var itemEnd = _gridMaker_timeToSeconds(item.end);
    var itemDuration = itemEnd - itemStart;

    if (isNaN(itemStart) || isNaN(itemEnd)) {
        return Number.MAX_VALUE;
    }

    var score = 0.0;
    score += Math.abs(itemStart - targetStart);
    score += Math.abs(itemEnd - targetEnd);
    score += Math.abs(itemDuration - targetDuration) * 0.25;

    if (targetTrack >= 0) {
        score += Math.abs(itemTrackIndex - targetTrack) * 0.05;
    }

    if (_gridMaker_isQEItemSelected(item)) {
        score -= 0.2;
    }

    var itemName = _gridMaker_clipName(item);
    if (targetName && itemName && targetName === itemName) {
        score -= 0.1;
    }

    return score;
}

function _gridMaker_findTrackIndex(seq, clip) {
    try {
        if (!seq || !clip || !clip.parentTrack || !seq.videoTracks) {
            return -1;
        }

        for (var i = 0; i < seq.videoTracks.numTracks; i++) {
            if (seq.videoTracks[i] === clip.parentTrack) {
                return i;
            }
        }
    } catch (e) {}

    return -1;
}

function _gridMaker_clipName(clipLike) {
    if (!clipLike) {
        return "";
    }

    var name = "";
    try {
        name = clipLike.name || "";
    } catch (e1) {}

    if (!name) {
        try {
            if (clipLike.projectItem) {
                name = clipLike.projectItem.name || "";
            }
        } catch (e2) {}
    }

    return String(name).toLowerCase();
}

function _gridMaker_isQEItemSelected(item) {
    try {
        if (item && typeof item.isSelected === "function") {
            return !!item.isSelected();
        }
    } catch (e1) {}

    try {
        if (item && typeof item.getSelected === "function") {
            return !!item.getSelected();
        }
    } catch (e2) {}

    return false;
}

function _gridMaker_timeToSeconds(timeLike) {
    var ticksPerSecond = 254016000000.0;

    if (timeLike === null || timeLike === undefined) {
        return NaN;
    }

    if (typeof timeLike === "number") {
        return timeLike;
    }

    if (typeof timeLike === "string") {
        var parsed = parseFloat(timeLike);
        return isNaN(parsed) ? NaN : parsed;
    }

    try {
        if (timeLike.seconds !== undefined) {
            var s = parseFloat(timeLike.seconds);
            if (!isNaN(s)) {
                return s;
            }
        }
    } catch (e1) {}

    try {
        if (timeLike.ticks !== undefined) {
            var t = parseFloat(timeLike.ticks);
            if (!isNaN(t)) {
                return t / ticksPerSecond;
            }
        }
    } catch (e2) {}

    return NaN;
}

function _gridMaker_ensureEffect(clip, qClip, lookupNames, resolverFn) {
    var comp = resolverFn(clip);
    if (comp) {
        return comp;
    }

    for (var i = 0; i < lookupNames.length; i++) {
        try {
            var fx = qe.project.getVideoEffectByName(lookupNames[i]);
            if (fx) {
                qClip.addVideoEffect(fx);
                comp = resolverFn(clip);
                if (comp) {
                    return comp;
                }
            }
        } catch (e1) {}
    }

    return resolverFn(clip);
}

function _gridMaker_transformEffectLookupNames() {
    return [
        "Transform",
        "Transformation",
        "Trasformazione",
        "Transformar",
        "Transformieren",
        "ADBE Transform",
        "ADBE Geometry2",
        "AE.ADBE Geometry2"
    ];
}

function _gridMaker_cropEffectLookupNames() {
    return [
        "Crop",
        "Recadrage",
        "Recortar",
        "Ritaglia",
        "Freistellen",
        "ADBE Crop",
        "AE.ADBE Crop"
    ];
}

function _gridMaker_findTransformComponent(clip) {
    return _gridMaker_findComponentByHints(
        clip,
        ["transform", "transformation", "trasform", "transformar", "transformier"],
        ["adbe transform", "adbe geometry2", "ae.adbe geometry2"],
        false
    );
}

function _gridMaker_findMotionComponent(clip) {
    return _gridMaker_findComponentByHints(
        clip,
        ["motion", "mouvement", "movimiento", "movimento", "beweg"],
        ["adbe motion"],
        true
    );
}

function _gridMaker_findCropComponent(clip) {
    return _gridMaker_findComponentByHints(
        clip,
        ["crop", "recadr", "recortar", "ritagli", "freistell"],
        ["adbe crop", "ae.adbe crop"],
        false
    );
}

function _gridMaker_findComponentByHints(clip, displayHints, matchHints, requirePlacementProps) {
    if (!clip || !clip.components) {
        return null;
    }

    var bestComp = null;
    var bestScore = -1;

    for (var c = 0; c < clip.components.numItems; c++) {
        var comp = clip.components[c];
        if (!comp) {
            continue;
        }

        var displayName = comp.displayName ? comp.displayName.toLowerCase() : "";
        var matchName = comp.matchName ? comp.matchName.toLowerCase() : "";

        if (requirePlacementProps && !_gridMaker_componentHasPlacementProps(comp)) {
            continue;
        }

        var score = _gridMaker_componentMatchScore(displayName, matchName, displayHints, matchHints);
        if (score > bestScore) {
            bestScore = score;
            bestComp = comp;
        }
    }

    return bestComp;
}

function _gridMaker_componentMatchScore(displayName, matchName, displayHints, matchHints) {
    var score = -1;

    for (var i = 0; i < matchHints.length; i++) {
        if (matchName === matchHints[i]) {
            score = Math.max(score, 100);
        } else if (matchName.indexOf(matchHints[i]) !== -1) {
            score = Math.max(score, 90);
        }
    }

    for (var j = 0; j < displayHints.length; j++) {
        if (displayName === displayHints[j]) {
            score = Math.max(score, 80);
        } else if (displayName.indexOf(displayHints[j]) !== -1) {
            score = Math.max(score, 70);
        }
    }

    return score;
}

function _gridMaker_componentHasPlacementProps(component) {
    var pos = _gridMaker_findProperty(component, [
        "position",
        "adbe transform position",
        "adbe position",
        "adbe motion position"
    ], "point2d");
    var scale = _gridMaker_findProperty(component, [
        "scale",
        "echelle",
        "escala",
        "scala",
        "adbe transform scale",
        "adbe scale",
        "adbe motion scale"
    ], "number");

    return !!pos && !!scale;
}

function _gridMaker_containsAny(source, hints) {
    if (!source || !hints || hints.length < 1) {
        return false;
    }

    for (var i = 0; i < hints.length; i++) {
        if (source.indexOf(hints[i]) !== -1) {
            return true;
        }
    }

    return false;
}

function _gridMaker_findProperty(component, names, expectedKind) {
    if (!component || !component.properties) {
        return null;
    }

    var targets = [];
    for (var i = 0; i < names.length; i++) {
        targets.push(names[i].toLowerCase());
    }

    var bestProp = null;
    var bestScore = -1;

    for (var p = 0; p < component.properties.numItems; p++) {
        var prop = component.properties[p];
        if (!prop) {
            continue;
        }

        var displayName = prop.displayName ? prop.displayName.toLowerCase() : "";
        var matchName = prop.matchName ? prop.matchName.toLowerCase() : "";
        for (var t = 0; t < targets.length; t++) {
            var score = -1;
            if (displayName === targets[t] || matchName === targets[t]) {
                score = 3;
            } else if (displayName.indexOf(targets[t]) !== -1 || matchName.indexOf(targets[t]) !== -1) {
                score = 2;
            }

            if (score > bestScore && _gridMaker_propertyMatchesExpected(prop, expectedKind)) {
                bestScore = score;
                bestProp = prop;
            }
        }
    }

    return bestProp;
}

function _gridMaker_propertyMatchesExpected(prop, expectedKind) {
    if (!expectedKind) {
        return true;
    }

    var value = _gridMaker_readPropertyValue(prop);
    if (expectedKind === "number") {
        return _gridMaker_isFiniteNumber(value);
    }
    if (expectedKind === "point2d") {
        return _gridMaker_isPointValue(value);
    }

    return true;
}

function _gridMaker_readPropertyValue(prop) {
    if (!prop || typeof prop.getValue !== "function") {
        return null;
    }

    try {
        return prop.getValue();
    } catch (e1) {}

    return null;
}

function _gridMaker_isFiniteNumber(value) {
    return (typeof value === "number" && isFinite(value));
}

function _gridMaker_isPointValue(value) {
    if (!value || value.length === undefined || value.length !== 2) {
        return false;
    }
    return _gridMaker_isFiniteNumber(parseFloat(value[0])) && _gridMaker_isFiniteNumber(parseFloat(value[1]));
}

function _gridMaker_disableTimeVarying(prop) {
    if (!prop || typeof prop.isTimeVarying !== "function" || typeof prop.setTimeVarying !== "function") {
        return;
    }

    try {
        if (prop.isTimeVarying()) {
            prop.setTimeVarying(false);
        }
    } catch (e1) {}
}

function _gridMaker_trySetNumberProperty(prop, value) {
    if (!prop || !_gridMaker_isFiniteNumber(value)) {
        return false;
    }

    try {
        prop.setValue(value, true);
        var readback = _gridMaker_readPropertyValue(prop);
        if (_gridMaker_isFiniteNumber(readback)) {
            return true;
        }
    } catch (e1) {}

    _gridMaker_disableTimeVarying(prop);

    try {
        prop.setValue(value, true);
        var readback2 = _gridMaker_readPropertyValue(prop);
        return _gridMaker_isFiniteNumber(readback2);
    } catch (e2) {}

    return false;
}

function _gridMaker_trySetPointProperty(prop, point) {
    if (!prop || !point || point.length < 2) {
        return false;
    }

    var x = parseFloat(point[0]);
    var y = parseFloat(point[1]);
    if (!_gridMaker_isFiniteNumber(x) || !_gridMaker_isFiniteNumber(y)) {
        return false;
    }

    try {
        prop.setValue([x, y], true);
        var readback = _gridMaker_readPropertyValue(prop);
        if (_gridMaker_isPointValue(readback)) {
            return true;
        }
    } catch (e1) {}

    _gridMaker_disableTimeVarying(prop);

    try {
        prop.setValue([x, y], true);
        var readback2 = _gridMaker_readPropertyValue(prop);
        return _gridMaker_isPointValue(readback2);
    } catch (e2) {}

    return false;
}

function _gridMaker_getCurrentScalePercent(component) {
    var scaleProp = _gridMaker_findProperty(component, [
        "scale",
        "echelle",
        "escala",
        "scala",
        "adbe transform scale",
        "adbe scale",
        "adbe motion scale"
    ], "number");
    if (!scaleProp) {
        return NaN;
    }

    try {
        var v = scaleProp.getValue();
        if (typeof v === "number") {
            return v;
        }
    } catch (e1) {}

    return NaN;
}

function _gridMaker_setPlacement(component, scale, x, y, frameW, frameH) {
    var kind = _gridMaker_componentKind(component);

    var uniform = _gridMaker_findProperty(component, ["uniform scale", "echelle uniforme", "uniform"]);
    if (uniform) {
        _gridMaker_trySetNumberProperty(uniform, 1);
    }

    var scaleProp = _gridMaker_findProperty(component, [
        "scale",
        "echelle",
        "escala",
        "scala",
        "adbe transform scale",
        "adbe scale",
        "adbe motion scale"
    ], "number");
    var scaleOk = _gridMaker_trySetNumberProperty(scaleProp, scale);

    var position = _gridMaker_findProperty(component, [
        "position",
        "adbe transform position",
        "adbe position",
        "adbe motion position"
    ], "point2d");
    if (!position) {
        return false;
    }

    var candidates = _gridMaker_buildPositionCandidates(kind, x, y, frameW, frameH);
    var positionOk = false;
    for (var i = 0; i < candidates.length; i++) {
        if (_gridMaker_trySetPointProperty(position, candidates[i])) {
            var readback = _gridMaker_readPropertyValue(position);
            if (_gridMaker_isPointValue(readback)) {
                var rx = parseFloat(readback[0]);
                var ry = parseFloat(readback[1]);
                if (_gridMaker_isFiniteNumber(rx) && _gridMaker_isFiniteNumber(ry) && Math.abs(rx) < 100000 && Math.abs(ry) < 100000) {
                    positionOk = true;
                    break;
                }
            } else {
                positionOk = true;
                break;
            }
        }
    }

    return !!scaleOk && !!positionOk;
}

function _gridMaker_buildPositionCandidates(kind, x, y, frameW, frameH) {
    var absolute = [x, y];
    var centered = [x - (frameW * 0.5), y - (frameH * 0.5)];

    if (kind === "transform") {
        return [centered, absolute];
    }
    if (kind === "motion") {
        return [absolute, centered];
    }
    return [absolute, centered];
}

function _gridMaker_componentKind(component) {
    if (!component) {
        return "unknown";
    }

    var displayName = "";
    var matchName = "";

    try {
        displayName = component.displayName ? component.displayName.toLowerCase() : "";
    } catch (e1) {}

    try {
        matchName = component.matchName ? component.matchName.toLowerCase() : "";
    } catch (e2) {}

    if (_gridMaker_containsAny(matchName, ["adbe motion"]) || _gridMaker_containsAny(displayName, ["motion", "mouvement", "movimiento", "movimento", "beweg"])) {
        return "motion";
    }
    if (_gridMaker_containsAny(matchName, ["adbe transform", "adbe geometry2", "ae.adbe geometry2"]) || _gridMaker_containsAny(displayName, ["transform", "transformation", "trasform", "transformar", "transformier"])) {
        return "transform";
    }

    return "unknown";
}

function _gridMaker_getSequenceFrameSize(seq, qSeq) {
    var candidates = [];

    try {
        candidates.push({
            width: _gridMaker_toNumber(seq.frameSizeHorizontal),
            height: _gridMaker_toNumber(seq.frameSizeVertical)
        });
    } catch (e1) {}

    try {
        var settings = seq.getSettings();
        if (settings) {
            candidates.push({
                width: _gridMaker_toNumber(settings.videoFrameWidth),
                height: _gridMaker_toNumber(settings.videoFrameHeight)
            });
        }
    } catch (e2) {}

    if (qSeq) {
        try {
            candidates.push({
                width: _gridMaker_toNumber(qSeq.videoFrameWidth),
                height: _gridMaker_toNumber(qSeq.videoFrameHeight)
            });
        } catch (e3) {}

        try {
            if (qSeq.sequence) {
                candidates.push({
                    width: _gridMaker_toNumber(qSeq.sequence.videoFrameWidth),
                    height: _gridMaker_toNumber(qSeq.sequence.videoFrameHeight)
                });
            }
        } catch (e4) {}
    }

    for (var i = 0; i < candidates.length; i++) {
        if (_gridMaker_isReasonableFrameSize(candidates[i].width, candidates[i].height)) {
            return candidates[i];
        }
    }

    return null;
}

function _gridMaker_toNumber(value) {
    if (value === null || value === undefined) {
        return NaN;
    }

    if (typeof value === "number") {
        return value;
    }

    if (typeof value === "string") {
        var normalized = value.replace(",", ".");
        var parsed = parseFloat(normalized);
        return isNaN(parsed) ? NaN : parsed;
    }

    try {
        if (value.value !== undefined) {
            return _gridMaker_toNumber(value.value);
        }
    } catch (e1) {}

    try {
        if (value.seconds !== undefined) {
            return _gridMaker_toNumber(value.seconds);
        }
    } catch (e2) {}

    return NaN;
}

function _gridMaker_isReasonableFrameSize(width, height) {
    if (!_gridMaker_isFiniteNumber(width) || !_gridMaker_isFiniteNumber(height)) {
        return false;
    }
    if (width < 16 || height < 16) {
        return false;
    }
    if (width > 20000 || height > 20000) {
        return false;
    }
    return true;
}

function _gridMaker_setCrop(component, left, right, top, bottom) {
    var pLeft = _gridMaker_findProperty(component, ["left", "gauche", "adbe crop left"]);
    var pRight = _gridMaker_findProperty(component, ["right", "droite", "adbe crop right"]);
    var pTop = _gridMaker_findProperty(component, ["top", "haut", "adbe crop top"]);
    var pBottom = _gridMaker_findProperty(component, ["bottom", "bas", "adbe crop bottom"]);

    if (pLeft) {
        try {
            pLeft.setValue(left, true);
        } catch (e1) {}
    }
    if (pRight) {
        try {
            pRight.setValue(right, true);
        } catch (e2) {}
    }
    if (pTop) {
        try {
            pTop.setValue(top, true);
        } catch (e3) {}
    }
    if (pBottom) {
        try {
            pBottom.setValue(bottom, true);
        } catch (e4) {}
    }
}

function _gridMaker_clamp(v, min, max) {
    if (v < min) {
        return min;
    }
    if (v > max) {
        return max;
    }
    return v;
}

function _gridMaker_result(status, code, details) {
    var payload = status + "|" + code;
    var detailString = _gridMaker_serializeDetails(details);
    if (detailString) {
        payload += "|" + detailString;
    }
    return payload;
}

function _gridMaker_serializeDetails(details) {
    if (!details) {
        return "";
    }

    var parts = [];
    for (var key in details) {
        if (!details.hasOwnProperty(key)) {
            continue;
        }
        parts.push(encodeURIComponent(key) + "=" + encodeURIComponent(String(details[key])));
    }

    return parts.join("&");
}

function gridMaker_ping() {
    return _gridMaker_result("OK", "pong");
}
