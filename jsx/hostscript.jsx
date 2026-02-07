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

        var qSeq = qe.project.getActiveSequence();
        if (!qSeq) {
            return _gridMaker_result("ERR", "qe_unavailable");
        }

        var qClip = _gridMaker_findQEClip(qSeq, clip.start.seconds, clip.end.seconds);
        if (!qClip) {
            return _gridMaker_result("ERR", "qe_clip_not_found");
        }

        var transformComp = _gridMaker_ensureEffect(clip, qClip, ["Transform", "Transformation"]);
        var cropComp = _gridMaker_ensureEffect(clip, qClip, ["Crop", "Recadrage"]);

        if (!transformComp) {
            return _gridMaker_result("ERR", "transform_effect_unavailable");
        }
        if (!cropComp) {
            return _gridMaker_result("ERR", "crop_effect_unavailable");
        }

        var frameW = parseFloat(seq.frameSizeHorizontal);
        var frameH = parseFloat(seq.frameSizeVertical);
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
        var scale = 100.0 / (cols * visX);

        var x = frameW * ((col + 0.5) / cols);
        var y = frameH * ((row + 0.5) / rows);

        _gridMaker_setTransform(transformComp, scale, x, y);
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

function _gridMaker_findQEClip(qSeq, startSeconds, endSeconds) {
    var tolerance = 0.05;
    var start = parseFloat(startSeconds);
    var end = parseFloat(endSeconds);

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

            var s = parseFloat(item.start.seconds);
            var e = parseFloat(item.end.seconds);
            if (Math.abs(s - start) <= tolerance && Math.abs(e - end) <= tolerance) {
                return item;
            }
        }
    }

    return null;
}

function _gridMaker_ensureEffect(clip, qClip, names) {
    var comp = _gridMaker_findComponent(clip, names);
    if (comp) {
        return comp;
    }

    for (var i = 0; i < names.length; i++) {
        var fx = qe.project.getVideoEffectByName(names[i]);
        if (fx) {
            qClip.addVideoEffect(fx);
            break;
        }
    }

    return _gridMaker_findComponent(clip, names);
}

function _gridMaker_findComponent(clip, names) {
    var lowerNames = [];
    for (var i = 0; i < names.length; i++) {
        lowerNames.push(names[i].toLowerCase());
    }

    for (var c = 0; c < clip.components.numItems; c++) {
        var comp = clip.components[c];
        if (!comp) {
            continue;
        }

        var displayName = comp.displayName ? comp.displayName.toLowerCase() : "";
        var matchName = comp.matchName ? comp.matchName.toLowerCase() : "";

        for (var n = 0; n < lowerNames.length; n++) {
            if (displayName.indexOf(lowerNames[n]) !== -1 || matchName.indexOf(lowerNames[n]) !== -1) {
                return comp;
            }
        }

        if (matchName.indexOf("adbe transform") !== -1 && (lowerNames[0].indexOf("transform") !== -1 || lowerNames[0].indexOf("transformation") !== -1)) {
            return comp;
        }
        if (matchName.indexOf("adbe crop") !== -1 && (lowerNames[0].indexOf("crop") !== -1 || lowerNames[0].indexOf("recadr") !== -1)) {
            return comp;
        }
    }

    return null;
}

function _gridMaker_findProperty(component, names) {
    var targets = [];
    for (var i = 0; i < names.length; i++) {
        targets.push(names[i].toLowerCase());
    }

    for (var p = 0; p < component.properties.numItems; p++) {
        var prop = component.properties[p];
        if (!prop) {
            continue;
        }

        var displayName = prop.displayName ? prop.displayName.toLowerCase() : "";
        var matchName = prop.matchName ? prop.matchName.toLowerCase() : "";
        for (var t = 0; t < targets.length; t++) {
            if (displayName === targets[t] || displayName.indexOf(targets[t]) !== -1 || matchName.indexOf(targets[t]) !== -1) {
                return prop;
            }
        }
    }

    return null;
}

function _gridMaker_setTransform(component, scale, x, y) {
    var uniform = _gridMaker_findProperty(component, ["uniform scale", "echelle uniforme", "uniform"]);
    if (uniform) {
        try {
            uniform.setValue(1, true);
        } catch (e1) {}
    }

    var scaleProp = _gridMaker_findProperty(component, ["scale", "echelle", "adbe transform scale"]);
    if (scaleProp) {
        try {
            scaleProp.setValue(scale, true);
        } catch (e2) {}
    }

    var position = _gridMaker_findProperty(component, ["position", "adbe transform position"]);
    if (position) {
        try {
            position.setValue([x, y], true);
        } catch (e3) {}
    }
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
