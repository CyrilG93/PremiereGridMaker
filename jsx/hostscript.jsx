// Expose host-side capability flags so the panel can adapt UI controls to this host.
function gridMaker_getHostCapabilities() {
    try {
        return _gridMaker_jsonStringify({
            ok: true,
            hostVersion: _gridMaker_getHostVersionString(),
            supportsRoundedCrop: _gridMaker_supportsRoundedCropEffect()
        });
    } catch (e) {
        return _gridMaker_jsonStringify({
            ok: false,
            hostVersion: "",
            supportsRoundedCrop: false,
            message: String(e)
        });
    }
}

// Main entry for classic grid placement (row/col in an N x M layout).
function gridMaker_applyToSelectedClip(row, col, rows, cols, ratioW, ratioH, marginPx, roundnessPct) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT row=" + row + " col=" + col + " rows=" + rows + " cols=" + cols + " ratio=" + ratioW + ":" + ratioH + " marginPx=" + marginPx + " roundnessPct=" + roundnessPct);
        app.enableQE();
        dbg("QE enabled");

        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_result("ERR", "no_active_sequence", null, debugLines);
        }
        dbg("Sequence name=" + (seq.name || "<unknown>"));

        row = parseInt(row, 10);
        col = parseInt(col, 10);
        rows = parseInt(rows, 10);
        cols = parseInt(cols, 10);
        ratioW = parseFloat(ratioW);
        ratioH = parseFloat(ratioH);
        marginPx = _gridMaker_parseMarginPx(marginPx);
        roundnessPct = _gridMaker_parseRoundnessPercent(roundnessPct);
        dbg("PARSED row=" + row + " col=" + col + " rows=" + rows + " cols=" + cols + " ratio=" + ratioW + ":" + ratioH + " marginPx=" + marginPx + " roundnessPct=" + roundnessPct);

        if (rows < 1 || cols < 1) {
            dbg("Invalid grid");
            return _gridMaker_result("ERR", "invalid_grid", null, debugLines);
        }
        if (row < 0 || col < 0 || row >= rows || col >= cols) {
            dbg("Cell out of bounds");
            return _gridMaker_result("ERR", "cell_out_of_bounds", null, debugLines);
        }
        if (!(ratioW > 0) || !(ratioH > 0)) {
            dbg("Invalid ratio");
            return _gridMaker_result("ERR", "invalid_ratio", null, debugLines);
        }

        var selection = seq.getSelection();
        if (!selection || selection.length < 1) {
            dbg("No timeline selection");
            return _gridMaker_result("ERR", "no_selection", null, debugLines);
        }
        dbg("Selection length=" + selection.length);

        var videoClips = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].mediaType === "Video") {
                videoClips.push(selection[i]);
                dbg("Selected video #" + videoClips.length + " name=" + _gridMaker_clipName(selection[i]) + " start=" + _gridMaker_timeToSeconds(selection[i].start) + " end=" + _gridMaker_timeToSeconds(selection[i].end));
            }
        }
        dbg("Video clips in selection=" + videoClips.length);
        if (videoClips.length < 1) {
            dbg("No selected video clip found in selection");
            return _gridMaker_result("ERR", "no_video_selected", null, debugLines);
        }
        if (videoClips.length > 1) {
            dbg("Multiple selected video clips; abort for deterministic behavior");
            return _gridMaker_result("ERR", "multiple_video_selected", null, debugLines);
        }
        var clip = videoClips[0];
        dbg("Clip name=" + _gridMaker_clipName(clip) + " start=" + _gridMaker_timeToSeconds(clip.start) + " end=" + _gridMaker_timeToSeconds(clip.end));

        // Prefer Rounded Crop on supported hosts, with safe fallback to classic Crop.
        var preferRoundedCrop = _gridMaker_supportsRoundedCropEffect();
        dbg("Rounded Crop supported=" + preferRoundedCrop);

        var transformComp = _gridMaker_findManagedTransformComponent(clip);
        var motionComp = _gridMaker_findMotionComponent(clip);
        var placementComp = motionComp;
        var cropComp = _gridMaker_findManagedCropComponent(clip, preferRoundedCrop);
        dbg("Components pre-check placement=" + _gridMaker_componentLabel(placementComp) + " transform=" + _gridMaker_componentLabel(transformComp) + " motion=" + _gridMaker_componentLabel(motionComp) + " crop=" + _gridMaker_componentLabel(cropComp));
        _gridMaker_dumpPlacementComponents(clip, debugLines, "BEFORE");

        var qSeq = null;
        var qClip = null;
        try {
            qSeq = qe.project.getActiveSequence();
        } catch (eQeSeq) {
            qSeq = null;
            dbg("QE sequence lookup exception=" + eQeSeq);
        }
        if (qSeq) {
            dbg("QE sequence acquired");
            qClip = _gridMaker_findQEClip(qSeq, seq, clip);
            if (qClip) {
                dbg("QE clip found");
            } else {
                dbg("QE clip not found (non-blocking unless effect ensure is required)");
            }
        } else {
            dbg("QE sequence unavailable (non-blocking unless effect ensure is required)");
        }

        transformComp = _gridMaker_tryEnsureOptionalTransform(clip, qSeq, qClip, debugLines);

        if (!cropComp || !placementComp) {
            if (!qSeq) {
                dbg("QE sequence unavailable");
                return _gridMaker_result("ERR", "qe_unavailable", null, debugLines);
            }
            if (!qClip) {
                dbg("QE clip not found");
                return _gridMaker_result("ERR", "qe_clip_not_found", null, debugLines);
            }

            if (!cropComp) {
                cropComp = _gridMaker_ensureManagedCropComponent(
                    clip,
                    qClip,
                    preferRoundedCrop,
                    debugLines
                );
                dbg("Crop component after ensure=" + _gridMaker_componentLabel(cropComp));
            }

            motionComp = _gridMaker_findMotionComponent(clip);
            placementComp = motionComp;
            dbg("Placement component after ensure=" + _gridMaker_componentLabel(placementComp));
        }

        if (!placementComp) {
            dbg("Motion component unavailable");
            return _gridMaker_result("ERR", "motion_effect_unavailable", null, debugLines);
        }
        dbg("Transform strategy: optional neutral effect; placement does not require it");
        dbg("Placement strategy: Motion only");

        var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
        if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
            dbg("Invalid sequence frame size");
            return _gridMaker_result("ERR", "invalid_sequence_size", null, debugLines);
        }
        var frameW = frameSize.width;
        var frameH = frameSize.height;
        var frameAspect = frameW / frameH;
        var cellRect = _gridMaker_computePaddedCellRect(
            frameW,
            frameH,
            col / cols,
            row / rows,
            1 / cols,
            1 / rows,
            marginPx,
            debugLines
        );
        if (!cellRect) {
            dbg("Invalid effective cell size after margin");
            return _gridMaker_result("ERR", "invalid_grid", null, debugLines);
        }
        var cellW = cellRect.width;
        var cellH = cellRect.height;
        var cellAspect = cellW / cellH;
        var preferHeightAxis = cellAspect <= 1.0;
        dbg("Frame size " + frameW + "x" + frameH + " aspect=" + frameAspect + " marginPx=" + marginPx);
        dbg("Cell size " + cellW.toFixed(3) + "x" + cellH.toFixed(3) + " aspect=" + cellAspect.toFixed(6) + " preferHeightAxis=" + preferHeightAxis + " center=[" + cellRect.centerX.toFixed(3) + "," + cellRect.centerY.toFixed(3) + "]");

        var cropL = 0.0;
        var cropR = 0.0;
        var cropT = 0.0;
        var cropB = 0.0;

        var nativeSize = _gridMaker_getClipNativeFrameSize(clip, qClip, debugLines);
        var sourceW = frameW;
        var sourceH = frameH;
        if (nativeSize && _gridMaker_isReasonableFrameSize(nativeSize.width, nativeSize.height)) {
            sourceW = nativeSize.width;
            sourceH = nativeSize.height;
        }

        var placementKind = _gridMaker_componentKind(placementComp);
        var currentPlacementPos = _gridMaker_getCurrentPosition(placementComp);
        var placementModeHint = _gridMaker_detectPositionMode(placementKind, currentPlacementPos, frameW, frameH);

        var intrinsicScaleFactor = 1.0;
        var assumeFrameFit = false;
        // Native mode: always compute Motion scale from the clip's native pixels.
        // Do not assume any implicit "fit/set to frame" factor.
        if (!_gridMaker_isFiniteNumber(intrinsicScaleFactor) || !(intrinsicScaleFactor > 0)) {
            intrinsicScaleFactor = 1.0;
            assumeFrameFit = false;
        }

        var baseDisplayW = sourceW * intrinsicScaleFactor;
        var baseDisplayH = sourceH * intrinsicScaleFactor;
        if (!(baseDisplayW > 0) || !(baseDisplayH > 0)) {
            baseDisplayW = frameW;
            baseDisplayH = frameH;
        }

        var scaleForWidth = 100.0 * (cellW / baseDisplayW);
        var scaleForHeight = 100.0 * (cellH / baseDisplayH);
        var scale = preferHeightAxis ? scaleForHeight : scaleForWidth;

        if (preferHeightAxis) {
            var prefScaledW = baseDisplayW * (scale / 100.0);
            if (prefScaledW + 0.0001 < cellW) {
                scale = scaleForWidth;
                dbg("Scale fallback to width to ensure full cell fill");
            }
        } else {
            var prefScaledH = baseDisplayH * (scale / 100.0);
            if (prefScaledH + 0.0001 < cellH) {
                scale = scaleForHeight;
                dbg("Scale fallback to height to ensure full cell fill");
            }
        }

        if (!_gridMaker_isFiniteNumber(scale) || !(scale > 0)) {
            scale = 100.0;
            dbg("Scale fallback to 100 due to invalid computed scale");
        }

        var scaledW = baseDisplayW * (scale / 100.0);
        var scaledH = baseDisplayH * (scale / 100.0);

        var visX = cellW / scaledW;
        var visY = cellH / scaledH;
        if (!_gridMaker_isFiniteNumber(visX) || !(visX > 0)) {
            visX = 1.0;
        }
        if (!_gridMaker_isFiniteNumber(visY) || !(visY > 0)) {
            visY = 1.0;
        }
        if (visX > 1.0) {
            visX = 1.0;
        }
        if (visY > 1.0) {
            visY = 1.0;
        }

        cropL = (1.0 - visX) * 0.5;
        cropR = cropL;
        cropT = (1.0 - visY) * 0.5;
        cropB = cropT;

        cropL = _gridMaker_clamp(cropL * 100.0, 0, 49.5);
        cropR = _gridMaker_clamp(cropR * 100.0, 0, 49.5);
        cropT = _gridMaker_clamp(cropT * 100.0, 0, 49.5);
        cropB = _gridMaker_clamp(cropB * 100.0, 0, 49.5);

        visX = 1.0 - (cropL + cropR) / 100.0;
        visY = 1.0 - (cropT + cropB) / 100.0;

        var x = cellRect.centerX;
        var y = cellRect.centerY;
        dbg("Computed placement mode kind=" + placementKind + " modeHint=" + placementModeHint + " assumeFrameFit=" + assumeFrameFit + " intrinsicScaleFactor=" + intrinsicScaleFactor.toFixed(6));
        dbg("Computed source size " + sourceW.toFixed(3) + "x" + sourceH.toFixed(3) + " baseDisplayAt100=" + baseDisplayW.toFixed(3) + "x" + baseDisplayH.toFixed(3));
        dbg("Computed target scale width=" + scaleForWidth.toFixed(3) + " height=" + scaleForHeight.toFixed(3) + " chosen=" + scale.toFixed(3));
        dbg("Computed scaled size " + scaledW.toFixed(3) + "x" + scaledH.toFixed(3) + " targetCell=" + cellW.toFixed(3) + "x" + cellH.toFixed(3));
        dbg("Computed crop LRTB=" + cropL.toFixed(3) + "," + cropR.toFixed(3) + "," + cropT.toFixed(3) + "," + cropB.toFixed(3) + " visX=" + visX.toFixed(6) + " visY=" + visY.toFixed(6));
        dbg("Computed position x=" + x.toFixed(3) + " y=" + y.toFixed(3));

        if (!_gridMaker_setPlacement(placementComp, scale, x, y, frameW, frameH, debugLines)) {
            dbg("Placement write failed");
            return _gridMaker_result("ERR", "placement_apply_failed", {
                x: x.toFixed(3),
                y: y.toFixed(3)
            }, debugLines);
        }
        var cropRequired = (cropL > 0.0001 || cropR > 0.0001 || cropT > 0.0001 || cropB > 0.0001);
        if (cropComp) {
            _gridMaker_setCrop(cropComp, cropL, cropR, cropT, cropB);
            _gridMaker_setCropRoundness(cropComp, roundnessPct, debugLines);
        } else if (cropRequired) {
            dbg("Crop component unavailable and crop required");
            return _gridMaker_result("ERR", "crop_effect_unavailable", null, debugLines);
        } else {
            dbg("Crop component unavailable but crop not required; skipping crop write");
        }
        dbg("Placement write succeeded");
        dbg("Readback scale=" + _gridMaker_getCurrentScalePercent(placementComp));
        dbg("Readback position=" + _gridMaker_pointToString(_gridMaker_getCurrentPosition(placementComp)));
        _gridMaker_dumpPlacementComponents(clip, debugLines, "AFTER");

        return _gridMaker_result("OK", "cell_applied", {
            row: row + 1,
            col: col + 1,
            scale: scale.toFixed(2)
        }, debugLines);
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_result("ERR", "exception", { message: e }, debugLines);
    }
}

// Main entry for designer placement using normalized bounds (0..1).
function gridMaker_applyToSelectedCustomCell(leftNorm, topNorm, widthNorm, heightNorm, ratioW, ratioH, marginPx, roundnessPct) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT customCell left=" + leftNorm + " top=" + topNorm + " width=" + widthNorm + " height=" + heightNorm + " ratio=" + ratioW + ":" + ratioH + " marginPx=" + marginPx + " roundnessPct=" + roundnessPct);
        app.enableQE();
        dbg("QE enabled");

        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_result("ERR", "no_active_sequence", null, debugLines);
        }
        dbg("Sequence name=" + (seq.name || "<unknown>"));

        leftNorm = parseFloat(leftNorm);
        topNorm = parseFloat(topNorm);
        widthNorm = parseFloat(widthNorm);
        heightNorm = parseFloat(heightNorm);
        ratioW = parseFloat(ratioW);
        ratioH = parseFloat(ratioH);
        marginPx = _gridMaker_parseMarginPx(marginPx);
        roundnessPct = _gridMaker_parseRoundnessPercent(roundnessPct);
        dbg("PARSED customCell left=" + leftNorm + " top=" + topNorm + " width=" + widthNorm + " height=" + heightNorm + " ratio=" + ratioW + ":" + ratioH + " marginPx=" + marginPx + " roundnessPct=" + roundnessPct);

        if (!(ratioW > 0) || !(ratioH > 0)) {
            dbg("Invalid ratio");
            return _gridMaker_result("ERR", "invalid_ratio", null, debugLines);
        }
        if (!(widthNorm > 0) || !(heightNorm > 0)) {
            dbg("Invalid custom cell size");
            return _gridMaker_result("ERR", "invalid_grid", null, debugLines);
        }
        if (!_gridMaker_isFiniteNumber(leftNorm) || !_gridMaker_isFiniteNumber(topNorm) || !_gridMaker_isFiniteNumber(widthNorm) || !_gridMaker_isFiniteNumber(heightNorm)) {
            dbg("Invalid custom cell values");
            return _gridMaker_result("ERR", "invalid_grid", null, debugLines);
        }
        if (leftNorm < 0 || topNorm < 0 || leftNorm + widthNorm > 1.000001 || topNorm + heightNorm > 1.000001) {
            dbg("Custom cell out of bounds");
            return _gridMaker_result("ERR", "cell_out_of_bounds", null, debugLines);
        }

        var selection = seq.getSelection();
        if (!selection || selection.length < 1) {
            dbg("No timeline selection");
            return _gridMaker_result("ERR", "no_selection", null, debugLines);
        }
        dbg("Selection length=" + selection.length);

        var videoClips = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].mediaType === "Video") {
                videoClips.push(selection[i]);
                dbg("Selected video #" + videoClips.length + " name=" + _gridMaker_clipName(selection[i]) + " start=" + _gridMaker_timeToSeconds(selection[i].start) + " end=" + _gridMaker_timeToSeconds(selection[i].end));
            }
        }
        dbg("Video clips in selection=" + videoClips.length);
        if (videoClips.length < 1) {
            dbg("No selected video clip found in selection");
            return _gridMaker_result("ERR", "no_video_selected", null, debugLines);
        }
        if (videoClips.length > 1) {
            dbg("Multiple selected video clips; abort for deterministic behavior");
            return _gridMaker_result("ERR", "multiple_video_selected", null, debugLines);
        }
        var clip = videoClips[0];
        dbg("Clip name=" + _gridMaker_clipName(clip) + " start=" + _gridMaker_timeToSeconds(clip.start) + " end=" + _gridMaker_timeToSeconds(clip.end));

        // Prefer Rounded Crop on supported hosts, with safe fallback to classic Crop.
        var preferRoundedCrop = _gridMaker_supportsRoundedCropEffect();
        dbg("Rounded Crop supported=" + preferRoundedCrop);

        var transformComp = _gridMaker_findManagedTransformComponent(clip);
        var motionComp = _gridMaker_findMotionComponent(clip);
        var placementComp = motionComp;
        var cropComp = _gridMaker_findManagedCropComponent(clip, preferRoundedCrop);
        dbg("Components pre-check placement=" + _gridMaker_componentLabel(placementComp) + " transform=" + _gridMaker_componentLabel(transformComp) + " motion=" + _gridMaker_componentLabel(motionComp) + " crop=" + _gridMaker_componentLabel(cropComp));
        _gridMaker_dumpPlacementComponents(clip, debugLines, "BEFORE");

        var qSeq = null;
        var qClip = null;
        try {
            qSeq = qe.project.getActiveSequence();
        } catch (eQeSeq) {
            qSeq = null;
            dbg("QE sequence lookup exception=" + eQeSeq);
        }
        if (qSeq) {
            dbg("QE sequence acquired");
            qClip = _gridMaker_findQEClip(qSeq, seq, clip);
            if (qClip) {
                dbg("QE clip found");
            } else {
                dbg("QE clip not found (non-blocking unless effect ensure is required)");
            }
        } else {
            dbg("QE sequence unavailable (non-blocking unless effect ensure is required)");
        }

        transformComp = _gridMaker_tryEnsureOptionalTransform(clip, qSeq, qClip, debugLines);

        if (!cropComp || !placementComp) {
            if (!qSeq) {
                dbg("QE sequence unavailable");
                return _gridMaker_result("ERR", "qe_unavailable", null, debugLines);
            }
            if (!qClip) {
                dbg("QE clip not found");
                return _gridMaker_result("ERR", "qe_clip_not_found", null, debugLines);
            }

            if (!cropComp) {
                cropComp = _gridMaker_ensureManagedCropComponent(
                    clip,
                    qClip,
                    preferRoundedCrop,
                    debugLines
                );
                dbg("Crop component after ensure=" + _gridMaker_componentLabel(cropComp));
            }

            motionComp = _gridMaker_findMotionComponent(clip);
            placementComp = motionComp;
            dbg("Placement component after ensure=" + _gridMaker_componentLabel(placementComp));
        }

        if (!placementComp) {
            dbg("Motion component unavailable");
            return _gridMaker_result("ERR", "motion_effect_unavailable", null, debugLines);
        }
        dbg("Transform strategy: optional neutral effect; placement does not require it");
        dbg("Placement strategy: Motion only");

        var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
        if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
            dbg("Invalid sequence frame size");
            return _gridMaker_result("ERR", "invalid_sequence_size", null, debugLines);
        }
        var frameW = frameSize.width;
        var frameH = frameSize.height;
        var frameAspect = frameW / frameH;
        var cellRect = _gridMaker_computePaddedCellRect(
            frameW,
            frameH,
            leftNorm,
            topNorm,
            widthNorm,
            heightNorm,
            marginPx,
            debugLines
        );
        if (!cellRect) {
            dbg("Invalid effective custom cell size after margin");
            return _gridMaker_result("ERR", "invalid_grid", null, debugLines);
        }
        var cellW = cellRect.width;
        var cellH = cellRect.height;
        var cellAspect = cellW / cellH;
        var preferHeightAxis = cellAspect <= 1.0;
        dbg("Frame size " + frameW + "x" + frameH + " aspect=" + frameAspect + " marginPx=" + marginPx);
        dbg("Custom cell size " + cellW.toFixed(3) + "x" + cellH.toFixed(3) + " aspect=" + cellAspect.toFixed(6) + " preferHeightAxis=" + preferHeightAxis + " center=[" + cellRect.centerX.toFixed(3) + "," + cellRect.centerY.toFixed(3) + "]");

        var cropL = 0.0;
        var cropR = 0.0;
        var cropT = 0.0;
        var cropB = 0.0;

        var nativeSize = _gridMaker_getClipNativeFrameSize(clip, qClip, debugLines);
        var sourceW = frameW;
        var sourceH = frameH;
        if (nativeSize && _gridMaker_isReasonableFrameSize(nativeSize.width, nativeSize.height)) {
            sourceW = nativeSize.width;
            sourceH = nativeSize.height;
        }

        var placementKind = _gridMaker_componentKind(placementComp);
        var currentPlacementPos = _gridMaker_getCurrentPosition(placementComp);
        var placementModeHint = _gridMaker_detectPositionMode(placementKind, currentPlacementPos, frameW, frameH);

        var intrinsicScaleFactor = 1.0;
        var assumeFrameFit = false;
        // Native mode: always compute Motion scale from the clip's native pixels.
        // Do not assume any implicit "fit/set to frame" factor.
        if (!_gridMaker_isFiniteNumber(intrinsicScaleFactor) || !(intrinsicScaleFactor > 0)) {
            intrinsicScaleFactor = 1.0;
            assumeFrameFit = false;
        }

        var baseDisplayW = sourceW * intrinsicScaleFactor;
        var baseDisplayH = sourceH * intrinsicScaleFactor;
        if (!(baseDisplayW > 0) || !(baseDisplayH > 0)) {
            baseDisplayW = frameW;
            baseDisplayH = frameH;
        }

        var scaleForWidth = 100.0 * (cellW / baseDisplayW);
        var scaleForHeight = 100.0 * (cellH / baseDisplayH);
        var scale = preferHeightAxis ? scaleForHeight : scaleForWidth;

        if (preferHeightAxis) {
            var prefScaledW = baseDisplayW * (scale / 100.0);
            if (prefScaledW + 0.0001 < cellW) {
                scale = scaleForWidth;
                dbg("Scale fallback to width to ensure full cell fill");
            }
        } else {
            var prefScaledH = baseDisplayH * (scale / 100.0);
            if (prefScaledH + 0.0001 < cellH) {
                scale = scaleForHeight;
                dbg("Scale fallback to height to ensure full cell fill");
            }
        }

        if (!_gridMaker_isFiniteNumber(scale) || !(scale > 0)) {
            scale = 100.0;
            dbg("Scale fallback to 100 due to invalid computed scale");
        }

        var scaledW = baseDisplayW * (scale / 100.0);
        var scaledH = baseDisplayH * (scale / 100.0);

        var visX = cellW / scaledW;
        var visY = cellH / scaledH;
        if (!_gridMaker_isFiniteNumber(visX) || !(visX > 0)) {
            visX = 1.0;
        }
        if (!_gridMaker_isFiniteNumber(visY) || !(visY > 0)) {
            visY = 1.0;
        }
        if (visX > 1.0) {
            visX = 1.0;
        }
        if (visY > 1.0) {
            visY = 1.0;
        }

        cropL = (1.0 - visX) * 0.5;
        cropR = cropL;
        cropT = (1.0 - visY) * 0.5;
        cropB = cropT;

        cropL = _gridMaker_clamp(cropL * 100.0, 0, 49.5);
        cropR = _gridMaker_clamp(cropR * 100.0, 0, 49.5);
        cropT = _gridMaker_clamp(cropT * 100.0, 0, 49.5);
        cropB = _gridMaker_clamp(cropB * 100.0, 0, 49.5);

        visX = 1.0 - (cropL + cropR) / 100.0;
        visY = 1.0 - (cropT + cropB) / 100.0;

        var x = cellRect.centerX;
        var y = cellRect.centerY;
        dbg("Computed placement mode kind=" + placementKind + " modeHint=" + placementModeHint + " assumeFrameFit=" + assumeFrameFit + " intrinsicScaleFactor=" + intrinsicScaleFactor.toFixed(6));
        dbg("Computed source size " + sourceW.toFixed(3) + "x" + sourceH.toFixed(3) + " baseDisplayAt100=" + baseDisplayW.toFixed(3) + "x" + baseDisplayH.toFixed(3));
        dbg("Computed target scale width=" + scaleForWidth.toFixed(3) + " height=" + scaleForHeight.toFixed(3) + " chosen=" + scale.toFixed(3));
        dbg("Computed scaled size " + scaledW.toFixed(3) + "x" + scaledH.toFixed(3) + " targetCell=" + cellW.toFixed(3) + "x" + cellH.toFixed(3));
        dbg("Computed crop LRTB=" + cropL.toFixed(3) + "," + cropR.toFixed(3) + "," + cropT.toFixed(3) + "," + cropB.toFixed(3) + " visX=" + visX.toFixed(6) + " visY=" + visY.toFixed(6));
        dbg("Computed position x=" + x.toFixed(3) + " y=" + y.toFixed(3));

        if (!_gridMaker_setPlacement(placementComp, scale, x, y, frameW, frameH, debugLines)) {
            dbg("Placement write failed");
            return _gridMaker_result("ERR", "placement_apply_failed", {
                x: x.toFixed(3),
                y: y.toFixed(3)
            }, debugLines);
        }
        var cropRequired = (cropL > 0.0001 || cropR > 0.0001 || cropT > 0.0001 || cropB > 0.0001);
        if (cropComp) {
            _gridMaker_setCrop(cropComp, cropL, cropR, cropT, cropB);
            _gridMaker_setCropRoundness(cropComp, roundnessPct, debugLines);
        } else if (cropRequired) {
            dbg("Crop component unavailable and crop required");
            return _gridMaker_result("ERR", "crop_effect_unavailable", null, debugLines);
        } else {
            dbg("Crop component unavailable but crop not required; skipping crop write");
        }
        dbg("Placement write succeeded");
        dbg("Readback scale=" + _gridMaker_getCurrentScalePercent(placementComp));
        dbg("Readback position=" + _gridMaker_pointToString(_gridMaker_getCurrentPosition(placementComp)));
        _gridMaker_dumpPlacementComponents(clip, debugLines, "AFTER");

        return _gridMaker_result("OK", "cell_applied", {
            scale: scale.toFixed(2)
        }, debugLines);
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_result("ERR", "exception", { message: e }, debugLines);
    }
}

// Designer capture endpoint: read selected clip visible rectangles (Motion + Crop)
// and return normalized bounds so the panel can create designer blocks from them.
function gridMaker_designerCaptureSelectedClipToBlock() {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT designerCapture");
        app.enableQE();
        dbg("QE enabled");

        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_result("ERR", "no_active_sequence", null, debugLines);
        }
        dbg("Sequence name=" + (seq.name || "<unknown>"));

        var selection = seq.getSelection();
        if (!selection || selection.length < 1) {
            dbg("No timeline selection");
            return _gridMaker_result("ERR", "no_selection", null, debugLines);
        }
        dbg("Selection length=" + selection.length);

        var videoClips = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].mediaType === "Video") {
                videoClips.push(selection[i]);
                dbg("Selected video #" + videoClips.length + " name=" + _gridMaker_clipName(selection[i]) + " start=" + _gridMaker_timeToSeconds(selection[i].start) + " end=" + _gridMaker_timeToSeconds(selection[i].end));
            }
        }
        dbg("Video clips in selection=" + videoClips.length);
        if (videoClips.length < 1) {
            dbg("No selected video clip found in selection");
            return _gridMaker_result("ERR", "no_video_selected", null, debugLines);
        }

        var qSeq = null;
        try {
            qSeq = qe.project.getActiveSequence();
        } catch (eQeSeq) {
            qSeq = null;
            dbg("QE sequence lookup exception=" + eQeSeq);
        }
        if (qSeq) {
            dbg("QE sequence acquired");
        } else {
            dbg("QE sequence unavailable (non-blocking for capture)");
        }

        var orderedClips = _gridMaker_sortClipsBottomToTop(seq, videoClips);
        var bounds = [];
        for (var ci = 0; ci < orderedClips.length; ci++) {
            var entry = orderedClips[ci];
            var capture = _gridMaker_captureDesignerClipBounds(seq, qSeq, entry.clip, debugLines, ci + 1);
            if (!capture.ok) {
                dbg("Designer capture failed clipIndex=" + (ci + 1) + " code=" + capture.code + " clip=" + _gridMaker_clipName(entry.clip));
                return _gridMaker_result("ERR", capture.code || "designer_capture_failed", {
                    failedClip: _gridMaker_clipName(entry.clip)
                }, debugLines);
            }
            bounds.push(capture.bounds);
        }

        if (bounds.length < 1) {
            dbg("Designer capture produced no bounds");
            return _gridMaker_result("ERR", "designer_capture_failed", null, debugLines);
        }

        dbg("Designer capture summary captured=" + bounds.length);
        return _gridMaker_result("OK", "designer_blocks_captured", {
            count: bounds.length,
            boundsJson: _gridMaker_jsonStringify(bounds)
        }, debugLines);
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_result("ERR", "exception", { message: e }, debugLines);
    }
}

function _gridMaker_captureDesignerClipBounds(seq, qSeq, clip, debugLines, index) {
    var prefix = "Capture #" + index + " ";
    _gridMaker_debugPush(debugLines, prefix + "clip name=" + _gridMaker_clipName(clip) + " start=" + _gridMaker_timeToSeconds(clip.start) + " end=" + _gridMaker_timeToSeconds(clip.end));

    var qClip = null;
    if (qSeq) {
        qClip = _gridMaker_findQEClip(qSeq, seq, clip);
        if (qClip) {
            _gridMaker_debugPush(debugLines, prefix + "QE clip found");
        } else {
            _gridMaker_debugPush(debugLines, prefix + "QE clip not found (non-blocking for capture)");
        }
    }

    var motionComp = _gridMaker_findMotionComponent(clip);
    var cropComp = _gridMaker_findManagedCropComponent(clip);
    _gridMaker_debugPush(debugLines, prefix + "components motion=" + _gridMaker_componentLabel(motionComp) + " crop=" + _gridMaker_componentLabel(cropComp));
    _gridMaker_dumpPlacementComponents(clip, debugLines, "CAPTURE #" + index);

    if (!motionComp) {
        _gridMaker_debugPush(debugLines, prefix + "Motion component unavailable");
        return { ok: false, code: "motion_effect_unavailable" };
    }

    var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
    if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
        _gridMaker_debugPush(debugLines, prefix + "invalid sequence frame size");
        return { ok: false, code: "invalid_sequence_size" };
    }
    var frameW = frameSize.width;
    var frameH = frameSize.height;
    _gridMaker_debugPush(debugLines, prefix + "frame size " + frameW + "x" + frameH);

    var nativeSize = _gridMaker_getClipNativeFrameSize(clip, qClip, debugLines);
    var sourceW = frameW;
    var sourceH = frameH;
    if (nativeSize && _gridMaker_isReasonableFrameSize(nativeSize.width, nativeSize.height)) {
        sourceW = nativeSize.width;
        sourceH = nativeSize.height;
    }

    // Capture uses the current Motion scale and current Crop values exactly as read.
    var scale = _gridMaker_getCurrentScalePercent(motionComp);
    if (!_gridMaker_isFiniteNumber(scale) || !(scale > 0)) {
        scale = 100.0;
        _gridMaker_debugPush(debugLines, prefix + "scale fallback to 100 due to invalid Motion scale readback");
    }

    var currentPos = _gridMaker_getCurrentPosition(motionComp);
    var positionMode = _gridMaker_detectPositionMode("motion", currentPos, frameW, frameH);
    var absPos = _gridMaker_positionToAbsolute(currentPos, positionMode, frameW, frameH);
    if (!absPos) {
        _gridMaker_debugPush(debugLines, prefix + "unable to convert Motion position to absolute frame coordinates");
        return { ok: false, code: "placement_apply_failed" };
    }

    var cropValues = _gridMaker_getCropValues(cropComp);
    _gridMaker_debugPush(debugLines, prefix + "readback scale=" + scale + " pos=" + _gridMaker_pointToString(currentPos) + " mode=" + positionMode + " absPos=[" + absPos[0] + "," + absPos[1] + "]");
    _gridMaker_debugPush(debugLines, prefix + "readback crop LRTB=" + cropValues.left.toFixed(3) + "," + cropValues.right.toFixed(3) + "," + cropValues.top.toFixed(3) + "," + cropValues.bottom.toFixed(3));

    var baseDisplayW = sourceW;
    var baseDisplayH = sourceH;
    if (!(baseDisplayW > 0) || !(baseDisplayH > 0)) {
        baseDisplayW = frameW;
        baseDisplayH = frameH;
    }

    var scaledW = baseDisplayW * (scale / 100.0);
    var scaledH = baseDisplayH * (scale / 100.0);
    if (!(scaledW > 0) || !(scaledH > 0)) {
        _gridMaker_debugPush(debugLines, prefix + "invalid scaled size from captured Motion scale");
        return { ok: false, code: "invalid_grid" };
    }

    var visX = 1.0 - ((cropValues.left + cropValues.right) / 100.0);
    var visY = 1.0 - ((cropValues.top + cropValues.bottom) / 100.0);
    visX = _gridMaker_clamp(visX, 0.0, 1.0);
    visY = _gridMaker_clamp(visY, 0.0, 1.0);

    var visibleW = scaledW * visX;
    var visibleH = scaledH * visY;
    if (!(visibleW > 0.5) || !(visibleH > 0.5)) {
        _gridMaker_debugPush(debugLines, prefix + "visible rectangle collapsed after crop");
        return { ok: false, code: "invalid_grid" };
    }

    var leftPx = absPos[0] - (visibleW * 0.5);
    var topPx = absPos[1] - (visibleH * 0.5);
    var rightPx = leftPx + visibleW;
    var bottomPx = topPx + visibleH;
    _gridMaker_debugPush(debugLines, prefix + "visible rect px before clip left=" + leftPx.toFixed(3) + " top=" + topPx.toFixed(3) + " right=" + rightPx.toFixed(3) + " bottom=" + bottomPx.toFixed(3));

    // Clip to the sequence frame because Designer blocks live only inside the visible canvas.
    if (leftPx < 0) {
        leftPx = 0;
    }
    if (topPx < 0) {
        topPx = 0;
    }
    if (rightPx > frameW) {
        rightPx = frameW;
    }
    if (bottomPx > frameH) {
        bottomPx = frameH;
    }
    var clippedW = rightPx - leftPx;
    var clippedH = bottomPx - topPx;
    _gridMaker_debugPush(debugLines, prefix + "visible rect px after clip left=" + leftPx.toFixed(3) + " top=" + topPx.toFixed(3) + " width=" + clippedW.toFixed(3) + " height=" + clippedH.toFixed(3));

    if (!(clippedW > 0.5) || !(clippedH > 0.5)) {
        _gridMaker_debugPush(debugLines, prefix + "visible rectangle is outside frame after clipping");
        return { ok: false, code: "designer_capture_out_of_bounds" };
    }

    var leftNorm = leftPx / frameW;
    var topNorm = topPx / frameH;
    var widthNorm = clippedW / frameW;
    var heightNorm = clippedH / frameH;

    _gridMaker_debugPush(debugLines, prefix + "visible rect norm left=" + leftNorm.toFixed(6) + " top=" + topNorm.toFixed(6) + " width=" + widthNorm.toFixed(6) + " height=" + heightNorm.toFixed(6));
    return {
        ok: true,
        bounds: {
            leftNorm: leftNorm.toFixed(6),
            topNorm: topNorm.toFixed(6),
            widthNorm: widthNorm.toFixed(6),
            heightNorm: heightNorm.toFixed(6),
            label: "clip_" + index,
            clipName: _gridMaker_clipName(clip)
        }
    };
}

// Batch endpoint: apply a list of normalized cells to selected clips (track order: bottom -> top).
function gridMaker_applyBatchToSelectedClips(cellsJson, ratioW, ratioH, marginPx, roundnessPct) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT batch ratio=" + ratioW + ":" + ratioH + " marginPx=" + marginPx + " roundnessPct=" + roundnessPct);
        app.enableQE();
        dbg("QE enabled");

        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_result("ERR", "no_active_sequence", null, debugLines);
        }
        dbg("Sequence name=" + (seq.name || "<unknown>"));

        ratioW = parseFloat(ratioW);
        ratioH = parseFloat(ratioH);
        marginPx = _gridMaker_parseMarginPx(marginPx);
        roundnessPct = _gridMaker_parseRoundnessPercent(roundnessPct);
        if (!(ratioW > 0) || !(ratioH > 0)) {
            dbg("Invalid ratio");
            return _gridMaker_result("ERR", "invalid_ratio", null, debugLines);
        }

        var parsedCells = _gridMaker_jsonParse(String(cellsJson || ""));
        var cells = _gridMaker_normalizeBatchCells(parsedCells);
        dbg("Batch cells parsed=" + cells.length);
        if (cells.length < 1) {
            dbg("No valid batch cells");
            return _gridMaker_result("ERR", "invalid_grid", null, debugLines);
        }

        var selection = seq.getSelection();
        if (!selection || selection.length < 1) {
            dbg("No timeline selection");
            return _gridMaker_result("ERR", "no_selection", null, debugLines);
        }
        dbg("Selection length=" + selection.length);

        var selectedVideoClips = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].mediaType === "Video") {
                selectedVideoClips.push(selection[i]);
            }
        }
        if (selectedVideoClips.length < 1) {
            dbg("No selected video clip found in selection");
            return _gridMaker_result("ERR", "no_video_selected", null, debugLines);
        }

        var orderedClips = _gridMaker_sortClipsBottomToTop(seq, selectedVideoClips);
        dbg("Batch selected video clips=" + orderedClips.length);
        for (var ci = 0; ci < orderedClips.length; ci++) {
            var oc = orderedClips[ci];
            dbg("Batch order #" + (ci + 1) + " track=" + oc.trackIndex + " clip=" + _gridMaker_clipName(oc.clip) + " start=" + _gridMaker_timeToSeconds(oc.clip.start));
        }

        var processCount = Math.min(orderedClips.length, cells.length);
        var applied = 0;
        var failed = 0;
        var firstError = "";
        var firstErrorClip = "";

        for (var bi = 0; bi < processCount; bi++) {
            var batchClip = orderedClips[bi].clip;
            var batchCell = cells[bi];
            dbg("Batch apply #" + (bi + 1) + " clip=" + _gridMaker_clipName(batchClip) + " cell=" + batchCell.label);

            var perClip = _gridMaker_applyNormalizedCellToClip(
                batchClip,
                seq,
                batchCell.leftNorm,
                batchCell.topNorm,
                batchCell.widthNorm,
                batchCell.heightNorm,
                ratioW,
                ratioH,
                marginPx,
                roundnessPct,
                debugLines
            );

            if (perClip.ok) {
                applied += 1;
                continue;
            }

            failed += 1;
            dbg("Batch clip failed code=" + perClip.code + " clip=" + _gridMaker_clipName(batchClip));
            if (!firstError) {
                firstError = perClip.code || "batch_apply_failed";
                firstErrorClip = _gridMaker_clipName(batchClip);
            }
        }

        var skipped = orderedClips.length - processCount;
        dbg("Batch summary applied=" + applied + " failed=" + failed + " skipped=" + skipped + " selected=" + orderedClips.length + " cells=" + cells.length);

        if (applied < 1) {
            return _gridMaker_result("ERR", firstError || "batch_apply_failed", {
                applied: applied,
                failed: failed,
                skipped: skipped,
                total: orderedClips.length,
                firstErrorClip: firstErrorClip
            }, debugLines);
        }

        return _gridMaker_result("OK", "batch_applied", {
            applied: applied,
            failed: failed,
            skipped: skipped,
            total: orderedClips.length
        }, debugLines);
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_result("ERR", "exception", { message: e }, debugLines);
    }
}

// Capture the selected video clips before a plugin mutation so the panel can build a multi-step undo stack.
function gridMaker_captureSelectedClipsState() {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT captureSelectedClipsState");
        app.enableQE();
        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_jsonStringify({ ok: false, code: "no_active_sequence", debug: debugLines.join("\n"), clips: [] });
        }

        var clips = _gridMaker_getSelectedVideoClips(seq, debugLines);
        if (clips.length < 1) {
            dbg("No selected video clips");
            return _gridMaker_jsonStringify({ ok: false, code: "no_video_selected", debug: debugLines.join("\n"), clips: [] });
        }

        var captured = [];
        for (var i = 0; i < clips.length; i++) {
            captured.push(_gridMaker_captureClipState(seq, clips[i], debugLines));
        }
        dbg("Captured clips=" + captured.length);
        return _gridMaker_jsonStringify({ ok: true, clips: captured, debug: debugLines.join("\n") });
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_jsonStringify({ ok: false, code: "exception", message: String(e), debug: debugLines.join("\n"), clips: [] });
    }
}

// Restore clips captured by gridMaker_captureSelectedClipsState; used by the panel Undo button.
function gridMaker_restoreClipsState(snapshotJson) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT restoreClipsState");
        app.enableQE();
        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_result("ERR", "no_active_sequence", null, debugLines);
        }

        var snapshot = _gridMaker_jsonParse(String(snapshotJson || ""));
        var clips = snapshot && snapshot.clips ? snapshot.clips : [];
        if (!(clips instanceof Array) || clips.length < 1) {
            dbg("Invalid undo snapshot");
            return _gridMaker_result("ERR", "undo_empty", null, debugLines);
        }

        var restored = 0;
        var missing = 0;
        for (var i = 0; i < clips.length; i++) {
            var item = clips[i];
            var clip = _gridMaker_findClipByReference(seq, item.ref, debugLines);
            if (!clip) {
                missing += 1;
                dbg("Undo target missing index=" + i);
                continue;
            }
            if (_gridMaker_restoreClipState(seq, clip, item, debugLines)) {
                restored += 1;
            }
        }

        dbg("Undo summary restored=" + restored + " missing=" + missing + " total=" + clips.length);
        if (restored < 1) {
            return _gridMaker_result("ERR", "undo_failed", { restored: restored, missing: missing, total: clips.length }, debugLines);
        }
        return _gridMaker_result("OK", "undo_applied", { restored: restored, missing: missing, total: clips.length }, debugLines);
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_result("ERR", "exception", { message: e }, debugLines);
    }
}

// Reset selected clips to full-frame Motion and default Grid Maker-managed Crop/Transform values.
function gridMaker_resetSelectedClips() {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT resetSelectedClips");
        app.enableQE();
        var seq = app.project.activeSequence;
        if (!seq) {
            dbg("No active sequence");
            return _gridMaker_result("ERR", "no_active_sequence", null, debugLines);
        }

        var clips = _gridMaker_getSelectedVideoClips(seq, debugLines);
        if (clips.length < 1) {
            dbg("No selected video clips");
            return _gridMaker_result("ERR", "no_video_selected", null, debugLines);
        }

        var reset = 0;
        for (var i = 0; i < clips.length; i++) {
            if (_gridMaker_resetClipState(seq, clips[i], debugLines)) {
                reset += 1;
            }
        }

        dbg("Reset summary reset=" + reset + " total=" + clips.length);
        if (reset < 1) {
            return _gridMaker_result("ERR", "reset_failed", { reset: reset, total: clips.length }, debugLines);
        }
        return _gridMaker_result("OK", "reset_applied", { reset: reset, total: clips.length }, debugLines);
    } catch (e) {
        dbg("EXCEPTION " + e);
        return _gridMaker_result("ERR", "exception", { message: e }, debugLines);
    }
}

function _gridMaker_getSelectedVideoClips(seq, debugLines) {
    // Collect only selected video TrackItems; linked audio selections are ignored.
    var out = [];
    var selection = seq && typeof seq.getSelection === "function" ? seq.getSelection() : null;
    if (!selection || selection.length < 1) {
        return out;
    }
    for (var i = 0; i < selection.length; i++) {
        if (selection[i] && selection[i].mediaType === "Video") {
            out.push(selection[i]);
            _gridMaker_debugPush(debugLines, "Selected video #" + out.length + " name=" + _gridMaker_clipName(selection[i]) + " track=" + _gridMaker_findTrackIndex(seq, selection[i]));
        }
    }
    return out;
}

function _gridMaker_captureClipState(seq, clip, debugLines) {
    // Store enough clip identity and effect values to restore this item later.
    return {
        ref: _gridMaker_buildClipReference(seq, clip),
        motion: _gridMaker_captureMotionState(_gridMaker_findMotionComponent(clip)),
        crop: _gridMaker_captureCropState(_gridMaker_findManagedCropComponent(clip, true)),
        transform: _gridMaker_captureTransformState(_gridMaker_findManagedTransformComponent(clip))
    };
}

function _gridMaker_buildClipReference(seq, clip) {
    return {
        trackIndex: _gridMaker_findTrackIndex(seq, clip),
        start: _gridMaker_timeToSeconds(clip.start),
        end: _gridMaker_timeToSeconds(clip.end),
        name: _gridMaker_clipName(clip),
        nodeId: _gridMaker_clipNodeId(clip)
    };
}

function _gridMaker_captureMotionState(component) {
    if (!component) {
        return { exists: false };
    }
    return {
        exists: true,
        position: _gridMaker_clonePoint(_gridMaker_getCurrentPosition(component)),
        scale: _gridMaker_getCurrentScalePercent(component)
    };
}

function _gridMaker_captureCropState(component) {
    if (!component) {
        return { exists: false };
    }
    return {
        exists: true,
        values: _gridMaker_getCropValues(component),
        roundness: _gridMaker_getCropRoundnessValue(component),
        rounded: _gridMaker_isRoundedCropComponent(component),
        label: _gridMaker_componentLabel(component)
    };
}

function _gridMaker_captureTransformState(component) {
    if (!component) {
        return { exists: false };
    }
    return {
        exists: true,
        uniformScale: _gridMaker_getTransformUniformScaleValue(component),
        label: _gridMaker_componentLabel(component)
    };
}

function _gridMaker_findClipByReference(seq, ref, debugLines) {
    if (!seq || !ref || !seq.videoTracks) {
        return null;
    }

    var targetTrack = parseInt(ref.trackIndex, 10);
    if (!isNaN(targetTrack) && targetTrack >= 0 && targetTrack < seq.videoTracks.numTracks) {
        var inTrack = _gridMaker_findClipInTrack(seq.videoTracks[targetTrack], ref);
        if (inTrack) {
            return inTrack;
        }
    }

    var best = null;
    var bestScore = Number.MAX_VALUE;
    for (var vt = 0; vt < seq.videoTracks.numTracks; vt++) {
        var track = seq.videoTracks[vt];
        if (!track || !track.clips) {
            continue;
        }
        for (var ci = 0; ci < track.clips.numItems; ci++) {
            var clip = track.clips[ci];
            var score = _gridMaker_clipReferenceScore(clip, vt, ref);
            if (score < bestScore) {
                bestScore = score;
                best = clip;
            }
        }
    }

    if (best && bestScore < 0.25) {
        _gridMaker_debugPush(debugLines, "Clip reference fallback matched score=" + bestScore + " name=" + _gridMaker_clipName(best));
        return best;
    }
    return null;
}

function _gridMaker_findClipInTrack(track, ref) {
    if (!track || !track.clips) {
        return null;
    }
    var best = null;
    var bestScore = Number.MAX_VALUE;
    for (var i = 0; i < track.clips.numItems; i++) {
        var clip = track.clips[i];
        var score = _gridMaker_clipReferenceScore(clip, ref.trackIndex, ref);
        if (score < bestScore) {
            bestScore = score;
            best = clip;
        }
    }
    return bestScore < 0.25 ? best : null;
}

function _gridMaker_clipReferenceScore(clip, trackIndex, ref) {
    if (!clip || !ref) {
        return Number.MAX_VALUE;
    }
    var score = 0;
    var start = _gridMaker_timeToSeconds(clip.start);
    var end = _gridMaker_timeToSeconds(clip.end);
    score += Math.abs(start - _gridMaker_toNumber(ref.start));
    score += Math.abs(end - _gridMaker_toNumber(ref.end));
    if (parseInt(ref.trackIndex, 10) !== trackIndex) {
        score += 10;
    }
    var nodeId = _gridMaker_clipNodeId(clip);
    if (ref.nodeId && nodeId && ref.nodeId !== nodeId) {
        score += 5;
    }
    var name = _gridMaker_clipName(clip);
    if (ref.name && name && ref.name !== name) {
        score += 1;
    }
    return score;
}

function _gridMaker_restoreClipState(seq, clip, snapshot, debugLines) {
    // Restore Motion first, then put Crop/Transform back to their captured state.
    var ok = _gridMaker_restoreMotionState(seq, clip, snapshot.motion, debugLines);
    ok = _gridMaker_restoreCropState(clip, snapshot.crop, debugLines) && ok;
    ok = _gridMaker_restoreTransformState(clip, snapshot.transform, debugLines) && ok;
    return ok;
}

function _gridMaker_resetClipState(seq, clip, debugLines) {
    // Reset only the placement/effects that Grid Maker manages.
    var ok = _gridMaker_resetMotionState(seq, clip, debugLines);
    ok = _gridMaker_removeOrNeutralizeManagedEffect(clip, "crop", debugLines) && ok;
    ok = _gridMaker_resetManagedTransformDefaults(seq, clip, debugLines) && ok;
    return ok;
}

function _gridMaker_restoreMotionState(seq, clip, motionState, debugLines) {
    var motion = _gridMaker_findMotionComponent(clip);
    if (!motion) {
        _gridMaker_debugPush(debugLines, "Restore Motion failed: missing component");
        return false;
    }
    if (!motionState || !motionState.exists) {
        return _gridMaker_resetMotionState(seq, clip, debugLines);
    }
    var scaleProp = _gridMaker_findProperty(motion, ["scale", "echelle", "escala", "scala", "adbe motion scale"], "number");
    var positionProp = _gridMaker_findProperty(motion, ["position", "adbe motion position"], "point2d");
    var scaleOk = _gridMaker_trySetNumberProperty(scaleProp, _gridMaker_toNumber(motionState.scale), debugLines, "undo.motion.scale");
    var positionOk = _gridMaker_trySetPointProperty(positionProp, motionState.position, debugLines, "undo.motion.position");
    return !!scaleOk && !!positionOk;
}

function _gridMaker_resetMotionState(seq, clip, debugLines) {
    var motion = _gridMaker_findMotionComponent(clip);
    if (!motion) {
        _gridMaker_debugPush(debugLines, "Reset Motion failed: missing component");
        return false;
    }
    var qSeq = null;
    try {
        qSeq = qe.project.getActiveSequence();
    } catch (e1) {
        qSeq = null;
    }
    var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
    if (!frameSize || !(frameSize.width > 0) || !(frameSize.height > 0)) {
        _gridMaker_debugPush(debugLines, "Reset Motion failed: invalid sequence size");
        return false;
    }

    var qClip = null;
    try {
        qClip = _gridMaker_findQEClip(qSeq, seq, clip);
    } catch (e2) {
        qClip = null;
    }

    // Reset uses the same fill logic as a single 1x1 Grid Maker cell covering the whole frame.
    var nativeSize = _gridMaker_getClipNativeFrameSize(clip, qClip, debugLines);
    var sourceW = frameSize.width;
    var sourceH = frameSize.height;
    if (nativeSize && _gridMaker_isReasonableFrameSize(nativeSize.width, nativeSize.height)) {
        sourceW = nativeSize.width;
        sourceH = nativeSize.height;
    }

    var scaleForWidth = 100.0 * (frameSize.width / sourceW);
    var scaleForHeight = 100.0 * (frameSize.height / sourceH);
    var scale = Math.max(scaleForWidth, scaleForHeight);
    if (!_gridMaker_isFiniteNumber(scale) || !(scale > 0)) {
        scale = 100.0;
    }

    _gridMaker_debugPush(debugLines, "Reset full-frame target scale=" + scale + " source=" + sourceW + "x" + sourceH + " frame=" + frameSize.width + "x" + frameSize.height);
    return _gridMaker_setPlacement(motion, scale, frameSize.width * 0.5, frameSize.height * 0.5, frameSize.width, frameSize.height, debugLines);
}

function _gridMaker_restoreCropState(clip, cropState, debugLines) {
    var crop = _gridMaker_findManagedCropComponent(clip, cropState && cropState.rounded);
    if (!cropState || !cropState.exists) {
        return _gridMaker_removeOrNeutralizeManagedEffect(clip, "crop", debugLines);
    }
    if (!crop) {
        _gridMaker_debugPush(debugLines, "Restore Crop skipped: original crop existed but no component is available");
        return false;
    }
    var values = cropState.values || {};
    _gridMaker_setCrop(crop, _gridMaker_toNumber(values.left), _gridMaker_toNumber(values.right), _gridMaker_toNumber(values.top), _gridMaker_toNumber(values.bottom));
    if (_gridMaker_isFiniteNumber(_gridMaker_toNumber(cropState.roundness))) {
        _gridMaker_setCropRoundness(crop, cropState.roundness, debugLines);
    }
    return true;
}

function _gridMaker_restoreTransformState(clip, transformState, debugLines) {
    var transform = _gridMaker_findManagedTransformComponent(clip);
    if (!transformState || !transformState.exists) {
        return _gridMaker_removeOrNeutralizeManagedEffect(clip, "transform", debugLines);
    }
    if (!transform) {
        _gridMaker_debugPush(debugLines, "Restore Transform skipped: original transform existed but no component is available");
        return true;
    }
    if (transformState.uniformScale !== null && transformState.uniformScale !== undefined) {
        var uniform = _gridMaker_findTransformUniformScaleProperty(transform);
        _gridMaker_trySetRawPropertyValue(uniform, transformState.uniformScale, debugLines, "undo.transform.uniformScale");
    }
    return true;
}

function _gridMaker_resetManagedTransformDefaults(seq, clip, debugLines) {
    var transform = _gridMaker_findManagedTransformComponent(clip);
    if (!transform) {
        _gridMaker_debugPush(debugLines, "Reset Transform skipped: no managed Transform component");
        return true;
    }

    var qSeq = null;
    try {
        qSeq = qe.project.getActiveSequence();
    } catch (e1) {
        qSeq = null;
    }

    var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
    if (!frameSize || !(frameSize.width > 0) || !(frameSize.height > 0)) {
        _gridMaker_debugPush(debugLines, "Reset Transform failed: invalid sequence size");
        return false;
    }

    _gridMaker_resetTransformDefaults(transform, frameSize.width, frameSize.height, debugLines);
    return true;
}

function _gridMaker_resetTransformDefaults(component, frameW, frameH, debugLines) {
    // Transform point values are normalized in the scripting API; Premiere displays them as sequence pixels.
    var center = [0.5, 0.5];
    _gridMaker_trySetPointProperty(_gridMaker_findTransformAnchorPointProperty(component), center, debugLines, "reset.transform.anchorPoint");
    _gridMaker_trySetPointProperty(_gridMaker_findTransformPositionProperty(component), center, debugLines, "reset.transform.position");
    _gridMaker_trySetToggleProperty(_gridMaker_findTransformUniformScaleProperty(component), true, debugLines, "reset.transform.uniformScale");
    _gridMaker_trySetNumericProperty(_gridMaker_findTransformScaleHeightProperty(component), 100, debugLines, "reset.transform.scaleHeight");
    _gridMaker_trySetNumericProperty(_gridMaker_findTransformScaleWidthProperty(component), 100, debugLines, "reset.transform.scaleWidth");
    _gridMaker_trySetNumberProperty(_gridMaker_findTransformSkewProperty(component), 0, debugLines, "reset.transform.skew");
    _gridMaker_trySetNumberProperty(_gridMaker_findTransformSkewAxisProperty(component), 0, debugLines, "reset.transform.skewAxis");
    _gridMaker_trySetNumberProperty(_gridMaker_findTransformRotationProperty(component), 0, debugLines, "reset.transform.rotation");
    _gridMaker_trySetNumberProperty(_gridMaker_findTransformOpacityProperty(component), 100, debugLines, "reset.transform.opacity");
    _gridMaker_trySetToggleProperty(_gridMaker_findTransformUseCompShutterProperty(component), true, debugLines, "reset.transform.useCompositionShutterAngle");
    _gridMaker_trySetNumberProperty(_gridMaker_findTransformShutterAngleProperty(component), 0, debugLines, "reset.transform.shutterAngle");
}

function _gridMaker_removeOrNeutralizeManagedEffect(clip, type, debugLines) {
    var comp = (type === "crop") ? _gridMaker_findManagedCropComponent(clip, true) : _gridMaker_findManagedTransformComponent(clip);
    if (!comp) {
        return true;
    }
    var qClip = null;
    try {
        qClip = _gridMaker_findQEClip(qe.project.getActiveSequence(), app.project.activeSequence, clip);
    } catch (e1) {
        qClip = null;
    }
    if (_gridMaker_tryRemoveComponent(clip, qClip, comp, type, debugLines)) {
        _gridMaker_debugPush(debugLines, "Removed managed " + type + " component");
        return true;
    }
    if (type === "crop") {
        _gridMaker_setCrop(comp, 0, 0, 0, 0);
        _gridMaker_setCropRoundness(comp, 0, debugLines);
    } else {
        _gridMaker_neutralizeTransform(comp, debugLines);
    }
    _gridMaker_debugPush(debugLines, "Neutralized managed " + type + " component because removal API was unavailable");
    return true;
}

function _gridMaker_tryRemoveComponent(clip, qClip, component, type, debugLines) {
    // Premiere's public ExtendScript docs do not expose effect deletion; try hidden runtime methods defensively.
    var beforeCount = _gridMaker_getTypeComponents(clip, type).length;
    var attempts = [
        { target: component, method: "remove", arg: null },
        { target: component, method: "delete", arg: null },
        { target: component, method: "removeFromClip", arg: null }
    ];

    var qeIndex = _gridMaker_findQEComponentIndex(qClip, type);
    if (qClip && qeIndex >= 0) {
        attempts.push({ target: qClip, method: "removeComponent", arg: qeIndex });
        attempts.push({ target: qClip, method: "removeVideoEffect", arg: qeIndex });
        var qeComp = _gridMaker_qeGetComponentAt(qClip, qeIndex);
        attempts.push({ target: qClip, method: "removeComponent", arg: qeComp });
        attempts.push({ target: qClip, method: "removeVideoEffect", arg: qeComp });
    }

    for (var i = 0; i < attempts.length; i++) {
        var attempt = attempts[i];
        if (!attempt.target || typeof attempt.target[attempt.method] !== "function") {
            continue;
        }
        try {
            if (attempt.arg === null) {
                attempt.target[attempt.method]();
            } else {
                attempt.target[attempt.method](attempt.arg);
            }
            _gridMaker_refreshHostUI();
            _gridMaker_sleepSafe(80);
            if (_gridMaker_getTypeComponents(clip, type).length < beforeCount || !_gridMaker_componentStillExists(clip, component)) {
                _gridMaker_debugPush(debugLines, "Effect remove succeeded method=" + attempt.method + " type=" + type);
                return true;
            }
        } catch (e1) {
            _gridMaker_debugPush(debugLines, "Effect remove attempt failed method=" + attempt.method + " type=" + type + " error=" + e1);
        }
    }
    return false;
}

function _gridMaker_findQEComponentIndex(qClip, type) {
    var count = _gridMaker_qeGetComponentCount(qClip);
    for (var i = count - 1; i >= 0; i--) {
        var qeComp = _gridMaker_qeGetComponentAt(qClip, i);
        if (_gridMaker_qeComponentMatchesType(qeComp, type)) {
            return i;
        }
    }
    return -1;
}

function _gridMaker_componentStillExists(clip, component) {
    if (!clip || !clip.components || !component) {
        return false;
    }
    for (var i = 0; i < clip.components.numItems; i++) {
        if (clip.components[i] === component) {
            return true;
        }
    }
    return false;
}

function _gridMaker_neutralizeTransform(component, debugLines) {
    if (!component) {
        return;
    }
    // Effect-removal fallback stays conservative and does not rewrite Transform Position.
    _gridMaker_trySetToggleProperty(_gridMaker_findTransformUniformScaleProperty(component), true, debugLines, "neutralize.transform.uniformScale");
    _gridMaker_trySetNumberProperty(_gridMaker_findProperty(component, ["scale", "scale height", "height", "hauteur"], "number"), 100, debugLines, "neutralize.transform.scaleHeight");
    _gridMaker_trySetNumberProperty(_gridMaker_findProperty(component, ["scale width", "width", "largeur"], "number"), 100, debugLines, "neutralize.transform.scaleWidth");
    _gridMaker_trySetNumberProperty(_gridMaker_findProperty(component, ["rotation", "adbe transform rotation"], "number"), 0, debugLines, "neutralize.transform.rotation");
}

function _gridMaker_clonePoint(point) {
    if (!_gridMaker_isPointValue(point)) {
        return null;
    }
    return [_gridMaker_toNumber(point[0]), _gridMaker_toNumber(point[1])];
}

function _gridMaker_findTransformUniformScaleProperty(component) {
    return _gridMaker_findProperty(component, [
        "uniform scale",
        "echelle uniforme",
        "échelle uniforme",
        "escala uniforme",
        "scala uniforme",
        "adbe geometry2-0003"
    ]);
}

function _gridMaker_findTransformAnchorPointProperty(component) {
    return _gridMaker_findProperty(component, [
        "anchor point",
        "point d'ancrage",
        "point d’ancrage",
        "punto de ancla",
        "punto di ancoraggio",
        "ankerpunkt",
        "ponto de ancoragem",
        "adbe geometry2-0001"
    ], "point2d");
}

function _gridMaker_findTransformPositionProperty(component) {
    return _gridMaker_findProperty(component, [
        "position",
        "adbe transform position",
        "adbe geometry2-0002"
    ], "point2d");
}

function _gridMaker_findTransformScaleHeightProperty(component) {
    return _gridMaker_findProperty(component, [
        "scale height",
        "hauteur d'echelle",
        "hauteur d’échelle",
        "hauteur",
        "alto de escala",
        "altezza scala",
        "skalierung hoehe",
        "skalierung höhe",
        "adbe geometry2-0004"
    ]) || _gridMaker_findProperty(component, [
        "scale",
        "echelle",
        "échelle",
        "adbe transform scale"
    ]);
}

function _gridMaker_findTransformScaleWidthProperty(component) {
    return _gridMaker_findProperty(component, [
        "scale width",
        "largeur d'echelle",
        "largeur d’échelle",
        "largeur",
        "ancho de escala",
        "larghezza scala",
        "skalierung breite",
        "adbe geometry2-0005"
    ]);
}

function _gridMaker_findTransformSkewProperty(component) {
    return _gridMaker_findProperty(component, [
        "skew",
        "inclinaison",
        "sesgar",
        "inclinazione",
        "neigung",
        "adbe geometry2-0006"
    ], "number");
}

function _gridMaker_findTransformSkewAxisProperty(component) {
    return _gridMaker_findProperty(component, [
        "skew axis",
        "axe d'inclinaison",
        "axe d’inclinaison",
        "eje de sesgo",
        "asse inclinazione",
        "neigungsachse",
        "adbe geometry2-0007"
    ], "number");
}

function _gridMaker_findTransformRotationProperty(component) {
    return _gridMaker_findProperty(component, [
        "rotation",
        "adbe transform rotation",
        "adbe geometry2-0008"
    ], "number");
}

function _gridMaker_findTransformOpacityProperty(component) {
    return _gridMaker_findProperty(component, [
        "opacity",
        "opacite",
        "opacité",
        "opacidad",
        "opacita",
        "opacità",
        "deckkraft",
        "adbe geometry2-0009"
    ], "number");
}

function _gridMaker_findTransformUseCompShutterProperty(component) {
    return _gridMaker_findProperty(component, [
        "use composition's shutter angle",
        "use composition shutter angle",
        "composition shutter",
        "utiliser l'angle d'obturation",
        "utiliser l’angle d’obturation",
        "usar angulo de obturador de composicion",
        "adbe geometry2-0010"
    ]);
}

function _gridMaker_findTransformShutterAngleProperty(component) {
    return _gridMaker_findProperty(component, [
        "shutter angle",
        "angle d'obturation",
        "angle d’obturation",
        "angle d'obturateur",
        "angle d’obturateur",
        "angulo de obturador",
        "angolo otturatore",
        "verschlusswinkel",
        "adbe geometry2-0011"
    ], "number");
}

function _gridMaker_findTransformSamplingProperty(component) {
    return _gridMaker_findProperty(component, [
        "sampling",
        "echantillonnage",
        "échantillonnage",
        "muestreo",
        "campionamento",
        "abtastung",
        "adbe geometry2-0012"
    ]);
}

function _gridMaker_trySetTransformSamplingBilinear(component, debugLines) {
    var sampling = _gridMaker_findTransformSamplingProperty(component);
    if (!sampling) {
        _gridMaker_debugPush(debugLines, "reset.transform.sampling skipped missing prop");
        return false;
    }

    // CEP exposes popup parameters inconsistently; string write is the safest non-destructive default attempt.
    return _gridMaker_trySetRawPropertyValue(sampling, "Bilinear", debugLines, "reset.transform.sampling");
}

function _gridMaker_trySetNumericProperty(prop, value, debugLines, label) {
    // Some Transform properties return numeric strings or locale-wrapped values; accept any finite numeric readback.
    if (!prop || !_gridMaker_isFiniteNumber(value)) {
        _gridMaker_debugPush(debugLines, "set-numeric skip " + (label || "?") + " invalid prop/value");
        return false;
    }

    _gridMaker_disableTimeVarying(prop);
    try {
        prop.setValue(value, true);
        var readback = _gridMaker_readPropertyValue(prop);
        var ok = _gridMaker_isNumericLikeFinite(readback);
        _gridMaker_debugPush(debugLines, "set-numeric attempt " + (label || "?") + " value=" + value + " readback=" + readback + " ok=" + ok);
        return ok;
    } catch (e1) {
        _gridMaker_debugPush(debugLines, "set-numeric failed " + (label || "?") + " error=" + e1);
    }

    return false;
}

function _gridMaker_isNumericLikeFinite(value) {
    if (_gridMaker_isFiniteNumber(value)) {
        return true;
    }
    if (typeof value === "string") {
        var parsed = parseFloat(value.replace(",", "."));
        return _gridMaker_isFiniteNumber(parsed);
    }
    return false;
}

function _gridMaker_trySetToggleOffProperty(prop, debugLines, label) {
    // Some Premiere toggles read back as booleans, others as 0/1 numeric values.
    if (!prop) {
        _gridMaker_debugPush(debugLines, "set-toggle-off skip " + (label || "?") + " missing prop");
        return false;
    }

    _gridMaker_disableTimeVarying(prop);
    var values = [false, 0];
    for (var i = 0; i < values.length; i++) {
        try {
            prop.setValue(values[i], true);
            var readback = _gridMaker_readPropertyValue(prop);
            var ok = !_gridMaker_isTruthyToggleValue(readback);
            _gridMaker_debugPush(debugLines, "set-toggle-off attempt " + (label || "?") + " value=" + values[i] + " readback=" + readback + " ok=" + ok);
            if (ok) {
                return true;
            }
        } catch (e1) {}
    }

    _gridMaker_debugPush(debugLines, "set-toggle-off failed " + (label || "?"));
    return false;
}

function _gridMaker_getTransformUniformScaleValue(component) {
    var prop = _gridMaker_findTransformUniformScaleProperty(component);
    return prop ? _gridMaker_readPropertyValue(prop) : null;
}

function _gridMaker_getCropRoundnessValue(component) {
    var prop = _gridMaker_findCropRoundnessProperty(component);
    var value = _gridMaker_toNumber(_gridMaker_readPropertyValue(prop));
    return _gridMaker_isFiniteNumber(value) ? value : null;
}

function _gridMaker_trySetRawPropertyValue(prop, value, debugLines, label) {
    // Restore captured non-numeric values such as boolean toggles without forcing a truthy readback.
    if (!prop || value === null || value === undefined) {
        return false;
    }
    _gridMaker_disableTimeVarying(prop);
    try {
        prop.setValue(value, true);
        _gridMaker_debugPush(debugLines, "set-raw ok " + (label || "?") + " value=" + value + " readback=" + _gridMaker_readPropertyValue(prop));
        return true;
    } catch (e1) {
        _gridMaker_debugPush(debugLines, "set-raw failed " + (label || "?") + " error=" + e1);
    }
    return false;
}

// Normalize UI batch cells payload to safe 0..1 rectangles.
function _gridMaker_normalizeBatchCells(rawCells) {
    var out = [];
    if (!(rawCells instanceof Array)) {
        return out;
    }

    for (var i = 0; i < rawCells.length; i++) {
        var raw = rawCells[i];
        if (!raw) {
            continue;
        }

        var leftNorm = _gridMaker_toNumber(raw.leftNorm);
        var topNorm = _gridMaker_toNumber(raw.topNorm);
        var widthNorm = _gridMaker_toNumber(raw.widthNorm);
        var heightNorm = _gridMaker_toNumber(raw.heightNorm);
        if (!_gridMaker_isFiniteNumber(leftNorm) || !_gridMaker_isFiniteNumber(topNorm) || !_gridMaker_isFiniteNumber(widthNorm) || !_gridMaker_isFiniteNumber(heightNorm)) {
            continue;
        }
        if (!(widthNorm > 0) || !(heightNorm > 0)) {
            continue;
        }
        if (leftNorm < 0 || topNorm < 0 || leftNorm + widthNorm > 1.000001 || topNorm + heightNorm > 1.000001) {
            continue;
        }

        out.push({
            leftNorm: leftNorm,
            topNorm: topNorm,
            widthNorm: widthNorm,
            heightNorm: heightNorm,
            label: String(raw.label || ("cell_" + i))
        });
    }
    return out;
}

// Sort clips by V track index ascending (V1, V2, V3...) and by timeline start.
function _gridMaker_sortClipsBottomToTop(seq, clips) {
    var entries = [];
    for (var i = 0; i < clips.length; i++) {
        var clip = clips[i];
        entries.push({
            clip: clip,
            trackIndex: _gridMaker_findTrackIndex(seq, clip),
            start: _gridMaker_timeToSeconds(clip.start),
            end: _gridMaker_timeToSeconds(clip.end),
            name: _gridMaker_clipName(clip)
        });
    }

    entries.sort(function (a, b) {
        var at = (a.trackIndex >= 0) ? a.trackIndex : 99999;
        var bt = (b.trackIndex >= 0) ? b.trackIndex : 99999;
        if (at !== bt) {
            return at - bt;
        }
        if (a.start !== b.start) {
            return a.start - b.start;
        }
        if (a.end !== b.end) {
            return a.end - b.end;
        }
        if (a.name > b.name) {
            return 1;
        }
        if (a.name < b.name) {
            return -1;
        }
        return 0;
    });

    return entries;
}

// Shared clip placement for batch mode (same placement rules as single custom cells).
function _gridMaker_applyNormalizedCellToClip(clip, seq, leftNorm, topNorm, widthNorm, heightNorm, ratioW, ratioH, marginPx, roundnessPct, debugLines) {
    var qSeq = null;
    var qClip = null;

    try {
        qSeq = qe.project.getActiveSequence();
    } catch (eQeSeq) {
        qSeq = null;
        _gridMaker_debugPush(debugLines, "QE sequence lookup exception=" + eQeSeq);
    }
    if (qSeq) {
        _gridMaker_debugPush(debugLines, "QE sequence acquired");
        qClip = _gridMaker_findQEClip(qSeq, seq, clip);
        if (qClip) {
            _gridMaker_debugPush(debugLines, "QE clip found");
        } else {
            _gridMaker_debugPush(debugLines, "QE clip not found (non-blocking unless effect ensure is required)");
        }
    } else {
        _gridMaker_debugPush(debugLines, "QE sequence unavailable (non-blocking unless effect ensure is required)");
    }

    // Prefer Rounded Crop on supported hosts, with safe fallback to classic Crop.
    var preferRoundedCrop = _gridMaker_supportsRoundedCropEffect();
    _gridMaker_debugPush(debugLines, "Rounded Crop supported=" + preferRoundedCrop);

    var transformComp = _gridMaker_findManagedTransformComponent(clip);
    var motionComp = _gridMaker_findMotionComponent(clip);
    var placementComp = motionComp;
    var cropComp = _gridMaker_findManagedCropComponent(clip, preferRoundedCrop);
    _gridMaker_debugPush(debugLines, "Components pre-check placement=" + _gridMaker_componentLabel(placementComp) + " transform=" + _gridMaker_componentLabel(transformComp) + " motion=" + _gridMaker_componentLabel(motionComp) + " crop=" + _gridMaker_componentLabel(cropComp));
    _gridMaker_dumpPlacementComponents(clip, debugLines, "BEFORE");

    transformComp = _gridMaker_tryEnsureOptionalTransform(clip, qSeq, qClip, debugLines);

    if (!cropComp || !placementComp) {
        if (!qSeq) {
            _gridMaker_debugPush(debugLines, "QE sequence unavailable");
            return { ok: false, code: "qe_unavailable" };
        }
        if (!qClip) {
            _gridMaker_debugPush(debugLines, "QE clip not found");
            return { ok: false, code: "qe_clip_not_found" };
        }

        if (!cropComp) {
            cropComp = _gridMaker_ensureManagedCropComponent(
                clip,
                qClip,
                preferRoundedCrop,
                debugLines
            );
            _gridMaker_debugPush(debugLines, "Crop component after ensure=" + _gridMaker_componentLabel(cropComp));
        }

        motionComp = _gridMaker_findMotionComponent(clip);
        placementComp = motionComp;
        _gridMaker_debugPush(debugLines, "Placement component after ensure=" + _gridMaker_componentLabel(placementComp));
    }

    if (!placementComp) {
        _gridMaker_debugPush(debugLines, "Motion component unavailable");
        return { ok: false, code: "motion_effect_unavailable" };
    }
    _gridMaker_debugPush(debugLines, "Transform strategy: optional neutral effect; placement does not require it");
    _gridMaker_debugPush(debugLines, "Placement strategy: Motion only");

    var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
    if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
        _gridMaker_debugPush(debugLines, "Invalid sequence frame size");
        return { ok: false, code: "invalid_sequence_size" };
    }
    var frameW = frameSize.width;
    var frameH = frameSize.height;
    var frameAspect = frameW / frameH;
    var cellRect = _gridMaker_computePaddedCellRect(frameW, frameH, leftNorm, topNorm, widthNorm, heightNorm, marginPx, debugLines);
    if (!cellRect) {
        _gridMaker_debugPush(debugLines, "Invalid effective custom cell size after margin");
        return { ok: false, code: "invalid_grid" };
    }
    var cellW = cellRect.width;
    var cellH = cellRect.height;
    var cellAspect = cellW / cellH;
    var preferHeightAxis = cellAspect <= 1.0;
    _gridMaker_debugPush(debugLines, "Frame size " + frameW + "x" + frameH + " aspect=" + frameAspect + " marginPx=" + marginPx);
    _gridMaker_debugPush(debugLines, "Custom cell size " + cellW.toFixed(3) + "x" + cellH.toFixed(3) + " aspect=" + cellAspect.toFixed(6) + " preferHeightAxis=" + preferHeightAxis + " center=[" + cellRect.centerX.toFixed(3) + "," + cellRect.centerY.toFixed(3) + "]");

    var cropL = 0.0;
    var cropR = 0.0;
    var cropT = 0.0;
    var cropB = 0.0;

    var nativeSize = _gridMaker_getClipNativeFrameSize(clip, qClip, debugLines);
    var sourceW = frameW;
    var sourceH = frameH;
    if (nativeSize && _gridMaker_isReasonableFrameSize(nativeSize.width, nativeSize.height)) {
        sourceW = nativeSize.width;
        sourceH = nativeSize.height;
    }

    var placementKind = _gridMaker_componentKind(placementComp);
    var currentPlacementPos = _gridMaker_getCurrentPosition(placementComp);
    var placementModeHint = _gridMaker_detectPositionMode(placementKind, currentPlacementPos, frameW, frameH);

    var intrinsicScaleFactor = 1.0;
    var assumeFrameFit = false;
    if (!_gridMaker_isFiniteNumber(intrinsicScaleFactor) || !(intrinsicScaleFactor > 0)) {
        intrinsicScaleFactor = 1.0;
        assumeFrameFit = false;
    }

    var baseDisplayW = sourceW * intrinsicScaleFactor;
    var baseDisplayH = sourceH * intrinsicScaleFactor;
    if (!(baseDisplayW > 0) || !(baseDisplayH > 0)) {
        baseDisplayW = frameW;
        baseDisplayH = frameH;
    }

    var scaleForWidth = 100.0 * (cellW / baseDisplayW);
    var scaleForHeight = 100.0 * (cellH / baseDisplayH);
    var scale = preferHeightAxis ? scaleForHeight : scaleForWidth;

    if (preferHeightAxis) {
        var prefScaledW = baseDisplayW * (scale / 100.0);
        if (prefScaledW + 0.0001 < cellW) {
            scale = scaleForWidth;
            _gridMaker_debugPush(debugLines, "Scale fallback to width to ensure full cell fill");
        }
    } else {
        var prefScaledH = baseDisplayH * (scale / 100.0);
        if (prefScaledH + 0.0001 < cellH) {
            scale = scaleForHeight;
            _gridMaker_debugPush(debugLines, "Scale fallback to height to ensure full cell fill");
        }
    }

    if (!_gridMaker_isFiniteNumber(scale) || !(scale > 0)) {
        scale = 100.0;
        _gridMaker_debugPush(debugLines, "Scale fallback to 100 due to invalid computed scale");
    }

    var scaledW = baseDisplayW * (scale / 100.0);
    var scaledH = baseDisplayH * (scale / 100.0);

    var visX = cellW / scaledW;
    var visY = cellH / scaledH;
    if (!_gridMaker_isFiniteNumber(visX) || !(visX > 0)) {
        visX = 1.0;
    }
    if (!_gridMaker_isFiniteNumber(visY) || !(visY > 0)) {
        visY = 1.0;
    }
    if (visX > 1.0) {
        visX = 1.0;
    }
    if (visY > 1.0) {
        visY = 1.0;
    }

    cropL = (1.0 - visX) * 0.5;
    cropR = cropL;
    cropT = (1.0 - visY) * 0.5;
    cropB = cropT;

    cropL = _gridMaker_clamp(cropL * 100.0, 0, 49.5);
    cropR = _gridMaker_clamp(cropR * 100.0, 0, 49.5);
    cropT = _gridMaker_clamp(cropT * 100.0, 0, 49.5);
    cropB = _gridMaker_clamp(cropB * 100.0, 0, 49.5);

    visX = 1.0 - (cropL + cropR) / 100.0;
    visY = 1.0 - (cropT + cropB) / 100.0;

    var x = cellRect.centerX;
    var y = cellRect.centerY;
    _gridMaker_debugPush(debugLines, "Computed placement mode kind=" + placementKind + " modeHint=" + placementModeHint + " assumeFrameFit=" + assumeFrameFit + " intrinsicScaleFactor=" + intrinsicScaleFactor.toFixed(6));
    _gridMaker_debugPush(debugLines, "Computed source size " + sourceW.toFixed(3) + "x" + sourceH.toFixed(3) + " baseDisplayAt100=" + baseDisplayW.toFixed(3) + "x" + baseDisplayH.toFixed(3));
    _gridMaker_debugPush(debugLines, "Computed target scale width=" + scaleForWidth.toFixed(3) + " height=" + scaleForHeight.toFixed(3) + " chosen=" + scale.toFixed(3));
    _gridMaker_debugPush(debugLines, "Computed scaled size " + scaledW.toFixed(3) + "x" + scaledH.toFixed(3) + " targetCell=" + cellW.toFixed(3) + "x" + cellH.toFixed(3));
    _gridMaker_debugPush(debugLines, "Computed crop LRTB=" + cropL.toFixed(3) + "," + cropR.toFixed(3) + "," + cropT.toFixed(3) + "," + cropB.toFixed(3) + " visX=" + visX.toFixed(6) + " visY=" + visY.toFixed(6));
    _gridMaker_debugPush(debugLines, "Computed position x=" + x.toFixed(3) + " y=" + y.toFixed(3));

    if (!_gridMaker_setPlacement(placementComp, scale, x, y, frameW, frameH, debugLines)) {
        _gridMaker_debugPush(debugLines, "Placement write failed");
        return { ok: false, code: "placement_apply_failed" };
    }
    var cropRequired = (cropL > 0.0001 || cropR > 0.0001 || cropT > 0.0001 || cropB > 0.0001);
    if (cropComp) {
        _gridMaker_setCrop(cropComp, cropL, cropR, cropT, cropB);
        _gridMaker_setCropRoundness(cropComp, roundnessPct, debugLines);
    } else if (cropRequired) {
        _gridMaker_debugPush(debugLines, "Crop component unavailable and crop required");
        return { ok: false, code: "crop_effect_unavailable" };
    } else {
        _gridMaker_debugPush(debugLines, "Crop component unavailable but crop not required; skipping crop write");
    }

    _gridMaker_debugPush(debugLines, "Placement write succeeded");
    _gridMaker_debugPush(debugLines, "Readback scale=" + _gridMaker_getCurrentScalePercent(placementComp));
    _gridMaker_debugPush(debugLines, "Readback position=" + _gridMaker_pointToString(_gridMaker_getCurrentPosition(placementComp)));
    _gridMaker_dumpPlacementComponents(clip, debugLines, "AFTER");

    return { ok: true, scale: scale.toFixed(2) };
}

// QE clip targeting helpers: find the timeline item that best matches selection.
function _gridMaker_findQEClip(qSeq, seq, clip) {
    var targetStart = _gridMaker_timeToSeconds(clip.start);
    var targetEnd = _gridMaker_timeToSeconds(clip.end);
    var targetDuration = targetEnd - targetStart;
    var targetTrack = _gridMaker_findTrackIndex(seq, clip);
    var targetName = _gridMaker_clipName(clip);

    // Priority 0: use currently selected QE video clip when possible.
    // This best matches manual drag-and-drop behavior on the selected timeline item.
    var selectedItem = _gridMaker_findBestSelectedQEClip(
        qSeq,
        targetStart,
        targetEnd,
        targetDuration,
        targetTrack,
        targetName
    );
    if (selectedItem) {
        return selectedItem;
    }

    // Priority 1: if we can resolve a target track, pick the best candidate on that track first.
    if (targetTrack >= 0 && targetTrack < qSeq.numVideoTracks) {
        var sameTrackItem = _gridMaker_findBestQEClipInTrack(
            qSeq,
            targetTrack,
            targetStart,
            targetEnd,
            targetDuration,
            targetTrack,
            targetName
        );
        if (sameTrackItem) {
            return sameTrackItem;
        }
    }

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

function _gridMaker_findBestSelectedQEClip(qSeq, targetStart, targetEnd, targetDuration, targetTrack, targetName) {
    if (!qSeq) {
        return null;
    }

    var bestItem = null;
    var bestScore = Number.MAX_VALUE;

    for (var vt = 0; vt < qSeq.numVideoTracks; vt++) {
        var track = qSeq.getVideoTrackAt(vt);
        if (!track) {
            continue;
        }

        for (var ci = 0; ci < track.numItems; ci++) {
            var item = track.getItemAt(ci);
            if (!item || !_gridMaker_isQEItemSelected(item)) {
                continue;
            }

            var score = _gridMaker_matchScore(item, vt, targetStart, targetEnd, targetDuration, targetTrack, targetName);
            if (score < bestScore) {
                bestScore = score;
                bestItem = item;
            }
        }
    }

    return bestItem;
}

function _gridMaker_findBestQEClipInTrack(qSeq, trackIndex, targetStart, targetEnd, targetDuration, targetTrack, targetName) {
    if (!qSeq || trackIndex < 0 || trackIndex >= qSeq.numVideoTracks) {
        return null;
    }
    var track = qSeq.getVideoTrackAt(trackIndex);
    if (!track) {
        return null;
    }

    var bestItem = null;
    var bestScore = Number.MAX_VALUE;
    for (var ci = 0; ci < track.numItems; ci++) {
        var item = track.getItemAt(ci);
        if (!item) {
            continue;
        }
        var score = _gridMaker_matchScore(item, trackIndex, targetStart, targetEnd, targetDuration, targetTrack, targetName);
        if (score < bestScore) {
            bestScore = score;
            bestItem = item;
        }
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
        // Strongly prefer the expected track; timing/name collisions across tracks are common.
        score += Math.abs(itemTrackIndex - targetTrack) * 25.0;
    }

    if (_gridMaker_isQEItemSelected(item)) {
        score -= 3.0;
    }

    var itemName = _gridMaker_clipName(item);
    if (targetName && itemName && targetName === itemName) {
        score -= 0.1;
    }

    return score;
}

function _gridMaker_findTrackIndex(seq, clip) {
    try {
        if (!seq || !clip || !seq.videoTracks) {
            return -1;
        }

        var targetNodeId = _gridMaker_clipNodeId(clip);
        var clipParentTrack = null;
        try {
            clipParentTrack = clip.parentTrack || null;
        } catch (e1) {
            clipParentTrack = null;
        }

        for (var i = 0; i < seq.videoTracks.numTracks; i++) {
            var track = seq.videoTracks[i];
            if (!track) {
                continue;
            }
            if (clipParentTrack && track === clipParentTrack) {
                return i;
            }

            // Fallback: compare clips in track by object identity and nodeId.
            try {
                if (track.clips && track.clips.numItems > 0) {
                    for (var j = 0; j < track.clips.numItems; j++) {
                        var trackClip = track.clips[j];
                        if (!trackClip) {
                            continue;
                        }
                        if (trackClip === clip) {
                            return i;
                        }
                        var trackClipNodeId = _gridMaker_clipNodeId(trackClip);
                        if (targetNodeId && trackClipNodeId && targetNodeId === trackClipNodeId) {
                            return i;
                        }
                    }
                }
            } catch (e2) {}
        }
    } catch (e) {}

    return -1;
}

function _gridMaker_clipNodeId(clipLike) {
    if (!clipLike) {
        return "";
    }
    try {
        if (clipLike.nodeId !== undefined && clipLike.nodeId !== null) {
            return String(clipLike.nodeId);
        }
    } catch (e1) {}
    try {
        if (clipLike.projectItem && clipLike.projectItem.nodeId !== undefined && clipLike.projectItem.nodeId !== null) {
            return String(clipLike.projectItem.nodeId);
        }
    } catch (e2) {}
    return "";
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

// Effect ensure helpers: locate/add Transform & Crop without creating duplicates.
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

function _gridMaker_findManagedEffectComponent(clip, type) {
    var list = _gridMaker_getTypeComponents(clip, type);
    if (!list || list.length < 1) {
        return null;
    }
    // Keep a stable "managed" pick without relying on rename APIs.
    return list[list.length - 1];
}

function _gridMaker_findManagedTransformComponent(clip) {
    return _gridMaker_findManagedEffectComponent(clip, "transform");
}

function _gridMaker_findManagedCropComponent(clip, preferRounded) {
    var list = _gridMaker_getTypeComponents(clip, "crop");
    if (!list || list.length < 1) {
        return null;
    }

    // When Rounded Crop is supported, always prioritize that effect variant.
    if (preferRounded) {
        for (var i = list.length - 1; i >= 0; i--) {
            if (_gridMaker_isRoundedCropComponent(list[i])) {
                return list[i];
            }
        }
    }

    // Keep the previous deterministic pick for non-rounded hosts.
    return list[list.length - 1];
}

// Ensure a managed crop component, preferring Rounded Crop on newer Premiere hosts.
function _gridMaker_ensureManagedCropComponent(clip, qClip, preferRounded, debugLines) {
    var existingPreferred = _gridMaker_findManagedCropComponent(clip, preferRounded);
    if (existingPreferred) {
        _gridMaker_debugPush(debugLines, "crop managed component found: " + _gridMaker_componentLabel(existingPreferred));
        return existingPreferred;
    }

    var existingAny = _gridMaker_findManagedCropComponent(clip, false);
    if (preferRounded && existingAny && !_gridMaker_isRoundedCropComponent(existingAny)) {
        _gridMaker_debugPush(debugLines, "Rounded Crop preferred; attempting rounded-only insertion");
        var ensuredRounded = _gridMaker_ensureManagedEffect(
            clip,
            qClip,
            "crop",
            _gridMaker_roundedCropEffectLookupNames(),
            debugLines,
            true
        );
        if (ensuredRounded && _gridMaker_isRoundedCropComponent(ensuredRounded)) {
            return ensuredRounded;
        }
        _gridMaker_debugPush(debugLines, "Rounded Crop unavailable; falling back to existing classic Crop");
        return existingAny;
    }

    return _gridMaker_ensureManagedEffect(
        clip,
        qClip,
        "crop",
        _gridMaker_cropEffectLookupNames(preferRounded),
        debugLines,
        false
    );
}

// Transform is useful as a neutral managed marker, but Motion performs the placement.
function _gridMaker_tryEnsureOptionalTransform(clip, qSeq, qClip, debugLines) {
    var transformComp = _gridMaker_findManagedTransformComponent(clip);
    if (transformComp) {
        _gridMaker_debugPush(debugLines, "Transform component found: " + _gridMaker_componentLabel(transformComp));
        _gridMaker_enableTransformUniformScale(transformComp, debugLines);
        return transformComp;
    }

    if (!qSeq || !qClip) {
        _gridMaker_debugPush(debugLines, "Transform optional ensure skipped because QE clip is unavailable");
        return null;
    }

    transformComp = _gridMaker_ensureManagedEffect(
        clip,
        qClip,
        "transform",
        _gridMaker_transformEffectLookupNames(),
        debugLines
    );
    _gridMaker_debugPush(debugLines, "Transform component after optional ensure=" + _gridMaker_componentLabel(transformComp));
    if (transformComp) {
        _gridMaker_enableTransformUniformScale(transformComp, debugLines);
    } else {
        _gridMaker_debugPush(debugLines, "Transform still unavailable; continuing with Motion-only placement");
    }
    return transformComp;
}

function _gridMaker_addVideoEffectToQEClip(qClip, fx, effectName, type, debugLines) {
    // Some Premiere QE builds only mutate the real selected timeline item.
    var selectedOk = _gridMaker_trySetQEItemSelected(qClip, true, true, debugLines);
    var modes = ["direct-boolean", "direct"];
    for (var i = 0; i < modes.length; i++) {
        try {
            var addReturn = null;
            if (modes[i] === "direct-boolean") {
                addReturn = qClip.addVideoEffect(fx, true);
            } else {
                addReturn = qClip.addVideoEffect(fx);
            }
            if (addReturn === false) {
                _gridMaker_debugPush(
                    debugLines,
                    "QE addVideoEffect returned false type=" + type +
                    " effect='" + effectName + "' mode=" + modes[i]
                );
                continue;
            }
            _gridMaker_refreshHostUI();
            _gridMaker_sleepSafe(120);
            _gridMaker_debugPush(
                debugLines,
                "QE addVideoEffect ok type=" + type +
                " effect='" + effectName + "' mode=" + modes[i] +
                " selectedOk=" + selectedOk +
                " return=" + addReturn
            );
            return { ok: true, mode: modes[i], qClip: qClip };
        } catch (e1) {
            _gridMaker_debugPush(
                debugLines,
                "QE addVideoEffect failed type=" + type +
                " effect='" + effectName + "' mode=" + modes[i] +
                " error=" + e1
            );
        }
    }
    return { ok: false, mode: "failed", qClip: qClip };
}

function _gridMaker_trySetQEItemSelected(item, selected, deselectOthers, debugLines) {
    if (!item || typeof item.setSelected !== "function") {
        _gridMaker_debugPush(debugLines, "QE setSelected unavailable for item=" + _gridMaker_qeClipDebugLabel(item));
        return false;
    }
    try {
        item.setSelected(selected === true, deselectOthers === true);
        return true;
    } catch (e1) {}
    try {
        item.setSelected(selected === true ? 1 : 0, deselectOthers === true ? 1 : 0);
        return true;
    } catch (e2) {}
    try {
        item.setSelected(selected === true);
        return true;
    } catch (e3) {}
    _gridMaker_debugPush(debugLines, "QE setSelected failed for item=" + _gridMaker_qeClipDebugLabel(item));
    return false;
}

function _gridMaker_refreshQEClipHandle(clip, fallbackQClip, debugLines) {
    // Reacquire the QE item after addVideoEffect because some builds expose stale wrappers.
    try {
        var seq = app.project.activeSequence;
        var qSeq = qe.project.getActiveSequence();
        var refreshed = _gridMaker_findQEClip(qSeq, seq, clip);
        if (refreshed) {
            _gridMaker_debugPush(debugLines, "QE clip refreshed after add: " + _gridMaker_qeClipDebugLabel(refreshed));
            return refreshed;
        }
    } catch (e1) {
        _gridMaker_debugPush(debugLines, "QE clip refresh failed: " + e1);
    }
    return fallbackQClip;
}

function _gridMaker_refreshHostUI() {
    try {
        if (app && typeof app.refresh === "function") {
            app.refresh();
        }
    } catch (e1) {}
}

function _gridMaker_sleepSafe(ms) {
    try {
        if (typeof $ !== "undefined" && $.sleep) {
            $.sleep(ms);
        }
    } catch (e1) {}
}

function _gridMaker_waitForManagedEffect(clip, type, attempts, sleepMs, debugLines) {
    var tries = parseInt(attempts, 10);
    if (!(tries > 0)) {
        tries = 1;
    }
    var delay = parseInt(sleepMs, 10);
    if (!(delay >= 0)) {
        delay = 0;
    }

    for (var i = 0; i < tries; i++) {
        var found = _gridMaker_findManagedEffectComponent(clip, type);
        if (found) {
            if (i > 0) {
                _gridMaker_debugPush(debugLines, type + " appeared after settle retry #" + (i + 1) + ": " + _gridMaker_componentLabel(found));
            }
            return found;
        }
        if (i + 1 < tries && delay > 0) {
            _gridMaker_sleepSafe(delay);
        }
    }
    return null;
}

function _gridMaker_ensureManagedEffect(clip, qClip, type, lookupNames, debugLines, forceInsert) {
    var byTypeNow = _gridMaker_getTypeComponents(clip, type);
    if (byTypeNow.length > 1) {
        _gridMaker_debugPush(debugLines, type + " duplicates detected before ensure=" + byTypeNow.length + " (will reuse existing, no new insert)");
    }

    var shouldForceInsert = !!forceInsert;
    var existing = _gridMaker_findManagedEffectComponent(clip, type);
    if (existing && !shouldForceInsert) {
        _gridMaker_debugPush(debugLines, type + " managed component found: " + _gridMaker_componentLabel(existing));
        return existing;
    }
    if (existing && shouldForceInsert) {
        _gridMaker_debugPush(debugLines, type + " forceInsert enabled; attempting insertion despite existing component=" + _gridMaker_componentLabel(existing));
    }

    var beforeType = byTypeNow;
    _gridMaker_debugPush(debugLines, type + " components before ensure=" + beforeType.length);

    var qeTypeCountBefore = _gridMaker_qeCountTypeComponents(qClip, type);
    _gridMaker_debugPush(debugLines, type + " QE components before ensure=" + qeTypeCountBefore);
    _gridMaker_debugPush(debugLines, type + " QE clip before ensure=" + _gridMaker_qeClipDebugLabel(qClip));

    for (var i = 0; i < lookupNames.length; i++) {
        var effectName = lookupNames[i];
        var fx = null;
        try {
            fx = qe.project.getVideoEffectByName(effectName);
        } catch (e1) {}
        if (!fx) {
            continue;
        }

        var addInfo = _gridMaker_addVideoEffectToQEClip(qClip, fx, effectName, type, debugLines);
        if (!addInfo.ok) {
            continue;
        }
        if (addInfo.qClip) {
            qClip = addInfo.qClip;
        }

        var afterType = _gridMaker_getTypeComponents(clip, type);
        var candidate = _gridMaker_pickNewComponent(beforeType, afterType);
        if (!candidate && afterType.length > beforeType.length) {
            candidate = afterType[afterType.length - 1];
        }
        qClip = _gridMaker_refreshQEClipHandle(clip, qClip, debugLines);
        var qeTypeCountAfter = _gridMaker_qeCountTypeComponents(qClip, type);
        _gridMaker_debugPush(
            debugLines,
            "Added " + type + " via '" + effectName + "' candidate=" + _gridMaker_componentLabel(candidate) +
            " qeCountBefore=" + qeTypeCountBefore + " qeCountAfter=" + qeTypeCountAfter +
            " addMode=" + addInfo.mode
        );

        if (candidate) {
            return candidate;
        }

        var settled = _gridMaker_waitForManagedEffect(clip, type, 4, 60, debugLines);
        if (settled) {
            _gridMaker_debugPush(debugLines, type + " became visible after short settle window: " + _gridMaker_componentLabel(settled));
            return settled;
        }

        // Prevent stacking duplicate adds across localized lookup names.
        _gridMaker_debugPush(debugLines, type + " add acknowledged without immediate candidate; stopping additional insertions");
        break;
    }

    var fallback = _gridMaker_findManagedEffectComponent(clip, type);
    if (fallback) {
        _gridMaker_debugPush(debugLines, type + " managed fallback found: " + _gridMaker_componentLabel(fallback));
    } else {
        _gridMaker_debugPush(debugLines, type + " still missing in clip.components after ensure pass");
    }
    return fallback;
}

function _gridMaker_effectTagSuffix() {
    return " (Grid Maker)";
}

function _gridMaker_hasGridMakerTag(name) {
    if (!name) {
        return false;
    }
    var suffix = _gridMaker_effectTagSuffix().toLowerCase();
    var lower = String(name).toLowerCase();
    return lower.indexOf(suffix) !== -1;
}

function _gridMaker_stripGridMakerTag(name) {
    if (!name) {
        return "";
    }
    var suffix = _gridMaker_effectTagSuffix();
    var trimmed = String(name);
    var idx = trimmed.lastIndexOf(suffix);
    if (idx === -1) {
        return trimmed;
    }
    return trimmed.substring(0, idx);
}

function _gridMaker_componentDisplayName(component) {
    if (!component) {
        return "";
    }

    try {
        if (component.displayName) {
            return String(component.displayName);
        }
    } catch (e1) {}

    try {
        if (component.name) {
            return String(component.name);
        }
    } catch (e2) {}

    return "";
}

function _gridMaker_transformLabelBase() {
    return "Transform";
}

function _gridMaker_cropLabelBase() {
    return "Crop";
}

function _gridMaker_getHostVersionString() {
    try {
        if (app && app.version !== undefined && app.version !== null) {
            return String(app.version);
        }
    } catch (e1) {}
    return "";
}

function _gridMaker_parseVersionTuple(rawVersion) {
    var text = String(rawVersion || "");
    var m = /([0-9]+)(?:\.([0-9]+))?/.exec(text);
    if (!m) {
        return null;
    }
    var major = parseInt(m[1], 10);
    var minor = parseInt(m[2] || "0", 10);
    if (isNaN(major) || isNaN(minor)) {
        return null;
    }
    return { major: major, minor: minor };
}

function _gridMaker_supportsRoundedCropEffect() {
    // Rounded Crop exists from Premiere Pro 25.5+, then QE must resolve it by name.
    var tuple = _gridMaker_parseVersionTuple(_gridMaker_getHostVersionString());
    if (!tuple) {
        return false;
    }
    if (tuple.major < 25) {
        return false;
    }
    if (tuple.major === 25 && tuple.minor < 5) {
        return false;
    }
    return _gridMaker_canResolveAnyVideoEffect(_gridMaker_roundedCropEffectLookupNames());
}

function _gridMaker_canResolveAnyVideoEffect(lookupNames) {
    // Resolve through QE so capability checks reflect the effects this host can actually add.
    try {
        app.enableQE();
    } catch (e0) {}
    if (!lookupNames || lookupNames.length < 1) {
        return false;
    }
    for (var i = 0; i < lookupNames.length; i++) {
        try {
            if (qe.project.getVideoEffectByName(lookupNames[i])) {
                return true;
            }
        } catch (e1) {}
    }
    return false;
}

function _gridMaker_isRoundedCropComponent(component) {
    if (!component) {
        return false;
    }
    var displayName = "";
    var matchName = "";
    try {
        displayName = component.displayName ? component.displayName.toLowerCase() : "";
    } catch (e1) {}
    try {
        matchName = component.matchName ? component.matchName.toLowerCase() : "";
    } catch (e2) {}

    return _gridMaker_containsAny(displayName, ["rounded crop", "recadrage arrondi"]) || _gridMaker_containsAny(matchName, ["ae.impact_crop_fx", "impact_crop_fx"]);
}

function _gridMaker_defaultLabelForType(type) {
    if (type === "transform") {
        return _gridMaker_transformLabelBase();
    }
    if (type === "crop") {
        return _gridMaker_cropLabelBase();
    }
    return "Effect";
}

function _gridMaker_componentMatchesType(component, type) {
    if (!component) {
        return false;
    }

    var displayName = "";
    var matchName = "";
    try {
        displayName = component.displayName ? component.displayName.toLowerCase() : "";
    } catch (e1) {}
    try {
        matchName = component.matchName ? component.matchName.toLowerCase() : "";
    } catch (e2) {}

    if (type === "transform") {
        if (_gridMaker_isMotionComponentExplicit(component)) {
            return false;
        }

        var transformScore = _gridMaker_componentMatchScore(
            displayName,
            matchName,
            ["transform", "transformation", "trasform", "transformar", "transformier"],
            ["adbe transform", "adbe geometry2", "ae.adbe geometry2"]
        );
        if (transformScore >= 0) {
            return true;
        }

        if (_gridMaker_componentHasTransformProps(component)) {
            return true;
        }
        return false;
    }

    if (type === "crop") {
        if (_gridMaker_isMotionComponentExplicit(component)) {
            return false;
        }

        var cropScore = _gridMaker_componentMatchScore(
            displayName,
            matchName,
            ["crop", "rounded crop", "recadr", "recortar", "ritagli", "freistell"],
            ["adbe crop", "adbe aecrop", "ae.adbe crop", "ae.adbe aecrop", "ae.impact_crop_fx", "impact_crop_fx"]
        );
        if (cropScore >= 0) {
            return true;
        }

        // Universal fallback: some Premiere locales/builds expose localized names
        // that are not covered by static effect-name lookups. Crop always has
        // the four directional numeric properties.
        if (_gridMaker_componentHasCropProps(component)) {
            return true;
        }
        return false;
    }

    return false;
}

function _gridMaker_componentHasTransformProps(component) {
    if (!component || !component.properties || component.properties.numItems < 3) {
        return false;
    }
    if (_gridMaker_isMotionComponentExplicit(component)) {
        return false;
    }

    var pos = _gridMaker_findProperty(component, [
        "position",
        "adbe transform position",
        "adbe position",
        "adbe geometry2-0001"
    ], "point2d");
    var scaleW = _gridMaker_findProperty(component, [
        "scale width",
        "largeur echelle",
        "anchura de escala",
        "larghezza scala",
        "breitenskalierung",
        "adbe geometry2-0004"
    ], "number");
    var scaleH = _gridMaker_findProperty(component, [
        "scale height",
        "hauteur echelle",
        "altura de escala",
        "altezza scala",
        "hoehenskalierung",
        "höhenskalierung",
        "adbe geometry2-0005"
    ], "number");

    return !!pos && !!scaleW && !!scaleH;
}

function _gridMaker_componentHasCropProps(component) {
    if (!component || !component.properties || component.properties.numItems < 4) {
        return false;
    }
    if (_gridMaker_isMotionComponentExplicit(component)) {
        return false;
    }

    var pLeft = _gridMaker_findProperty(component, [
        "left",
        "gauche",
        "izquierda",
        "sinistra",
        "links",
        "esquerda",
        "adbe crop left"
    ], "number");
    var pRight = _gridMaker_findProperty(component, [
        "right",
        "droite",
        "derecha",
        "destra",
        "rechts",
        "direita",
        "adbe crop right"
    ], "number");
    var pTop = _gridMaker_findProperty(component, [
        "top",
        "haut",
        "superior",
        "alto",
        "oben",
        "topo",
        "adbe crop top"
    ], "number");
    var pBottom = _gridMaker_findProperty(component, [
        "bottom",
        "bas",
        "inferior",
        "basso",
        "unten",
        "baixo",
        "adbe crop bottom"
    ], "number");

    return !!pLeft && !!pRight && !!pTop && !!pBottom;
}

function _gridMaker_isMotionComponentExplicit(component) {
    if (!component) {
        return false;
    }

    var displayName = "";
    var matchName = "";
    try {
        displayName = component.displayName ? component.displayName.toLowerCase() : "";
    } catch (e1) {}
    try {
        matchName = component.matchName ? component.matchName.toLowerCase() : "";
    } catch (e2) {}

    return _gridMaker_containsAny(matchName, ["adbe motion"]) || _gridMaker_containsAny(displayName, ["motion", "mouvement", "movimiento", "movimento", "beweg"]);
}

function _gridMaker_getTypeComponents(clip, type) {
    var out = [];
    if (!clip || !clip.components) {
        return out;
    }

    for (var i = 0; i < clip.components.numItems; i++) {
        var comp = clip.components[i];
        if (!comp) {
            continue;
        }
        if (_gridMaker_componentMatchesType(comp, type)) {
            out.push(comp);
        }
    }

    return out;
}

function _gridMaker_findTaggedComponentByType(clip, type) {
    var comps = _gridMaker_getTypeComponents(clip, type);
    for (var i = 0; i < comps.length; i++) {
        var name = _gridMaker_componentDisplayName(comps[i]);
        if (_gridMaker_hasGridMakerTag(name)) {
            return comps[i];
        }
    }
    return null;
}

function _gridMaker_findTaggedTransformComponent(clip) {
    return _gridMaker_findTaggedComponentByType(clip, "transform");
}

function _gridMaker_findTaggedCropComponent(clip) {
    return _gridMaker_findTaggedComponentByType(clip, "crop");
}

function _gridMaker_tryTagComponent(component, type, debugLines) {
    if (!component) {
        return false;
    }

    var currentName = _gridMaker_componentDisplayName(component);
    if (_gridMaker_hasGridMakerTag(currentName)) {
        _gridMaker_debugPush(debugLines, "Tag already present on " + type + " component: " + _gridMaker_componentLabel(component));
        return true;
    }

    var baseName = _gridMaker_stripGridMakerTag(currentName);
    if (!baseName) {
        baseName = _gridMaker_defaultLabelForType(type);
    }
    var targetName = baseName + _gridMaker_effectTagSuffix();

    try {
        component.displayName = targetName;
    } catch (e1) {}
    if (_gridMaker_hasGridMakerTag(_gridMaker_componentDisplayName(component))) {
        _gridMaker_debugPush(debugLines, "Tag set via displayName: " + _gridMaker_componentLabel(component));
        return true;
    }

    try {
        component.name = targetName;
    } catch (e2) {}
    if (_gridMaker_hasGridMakerTag(_gridMaker_componentDisplayName(component))) {
        _gridMaker_debugPush(debugLines, "Tag set via name: " + _gridMaker_componentLabel(component));
        return true;
    }

    try {
        if (typeof component.setName === "function") {
            component.setName(targetName);
        }
    } catch (e3) {}
    if (_gridMaker_hasGridMakerTag(_gridMaker_componentDisplayName(component))) {
        _gridMaker_debugPush(debugLines, "Tag set via setName: " + _gridMaker_componentLabel(component));
        return true;
    }

    _gridMaker_debugPush(debugLines, "Unable to tag " + type + " component name (API limitation likely): " + _gridMaker_componentLabel(component));
    return false;
}

function _gridMaker_qeGetComponentCount(qClip) {
    if (!qClip) {
        return 0;
    }
    try {
        if (qClip.numComponents !== undefined) {
            return parseInt(qClip.numComponents, 10) || 0;
        }
    } catch (e1) {}
    return 0;
}

function _gridMaker_qeGetComponentAt(qClip, index) {
    if (!qClip) {
        return null;
    }
    try {
        if (typeof qClip.getComponentAt === "function") {
            return qClip.getComponentAt(index);
        }
    } catch (e1) {}
    return null;
}

function _gridMaker_qeComponentLabel(component) {
    if (!component) {
        return "<none>";
    }
    var name = "";
    var matchName = "";
    try {
        name = component.name || component.displayName || "";
    } catch (e1) {}
    try {
        matchName = component.matchName || "";
    } catch (e2) {}
    return "[" + name + "|" + matchName + "]";
}

function _gridMaker_qeClipDebugLabel(qClip) {
    if (!qClip) {
        return "<none>";
    }
    var name = "";
    var start = "";
    var end = "";
    var selected = false;
    try {
        name = qClip.name || "";
    } catch (e1) {}
    try {
        start = String(_gridMaker_timeToSeconds(qClip.start));
    } catch (e2) {}
    try {
        end = String(_gridMaker_timeToSeconds(qClip.end));
    } catch (e3) {}
    selected = _gridMaker_isQEItemSelected(qClip);
    return "[name=" + name + " start=" + start + " end=" + end +
        " selected=" + selected + " components=" + _gridMaker_qeGetComponentCount(qClip) + "]";
}

function _gridMaker_qeComponentMatchesType(component, type) {
    if (!component) {
        return false;
    }
    var name = "";
    var matchName = "";
    try {
        name = (component.name || component.displayName || "").toLowerCase();
    } catch (e1) {}
    try {
        matchName = (component.matchName || "").toLowerCase();
    } catch (e2) {}

    if (type === "transform") {
        return _gridMaker_containsAny(matchName, ["adbe geometry", "adbe transform"]) || _gridMaker_containsAny(name, ["transform", "transformation", "trasform"]);
    }
    if (type === "crop") {
        return _gridMaker_containsAny(matchName, ["adbe crop", "adbe aecrop", "ae.impact_crop_fx", "impact_crop_fx"]) || _gridMaker_containsAny(name, ["crop", "rounded crop", "recadr", "recortar", "ritagli"]);
    }
    return false;
}

function _gridMaker_qeCountTypeComponents(qClip, type) {
    var count = _gridMaker_qeGetComponentCount(qClip);
    if (count < 1) {
        return 0;
    }
    var out = 0;
    for (var i = 0; i < count; i++) {
        var qeComp = _gridMaker_qeGetComponentAt(qClip, i);
        if (!qeComp) {
            continue;
        }
        if (_gridMaker_qeComponentMatchesType(qeComp, type)) {
            out += 1;
        }
    }
    return out;
}

function _gridMaker_tryTagSingleTypeViaQE(qClip, type, targetName, debugLines) {
    var count = _gridMaker_qeGetComponentCount(qClip);
    if (count < 1) {
        return false;
    }

    var matches = [];
    for (var i = 0; i < count; i++) {
        var qeComp = _gridMaker_qeGetComponentAt(qClip, i);
        if (!qeComp) {
            continue;
        }
        if (_gridMaker_qeComponentMatchesType(qeComp, type)) {
            matches.push(qeComp);
        }
    }

    if (matches.length !== 1) {
        _gridMaker_debugPush(debugLines, "QE rename skipped for " + type + ": ambiguous components=" + matches.length);
        return false;
    }

    var c = matches[0];
    try {
        c.displayName = targetName;
    } catch (e1) {}
    try {
        c.name = targetName;
    } catch (e2) {}
    try {
        if (typeof c.setName === "function") {
            c.setName(targetName);
        }
    } catch (e3) {}

    var label = _gridMaker_qeComponentLabel(c);
    var ok = _gridMaker_hasGridMakerTag(label);
    _gridMaker_debugPush(debugLines, "QE rename attempt for " + type + " => " + label + " ok=" + ok);
    return ok;
}

function _gridMaker_pickNewComponent(beforeList, afterList) {
    if (!afterList || afterList.length < 1) {
        return null;
    }
    if (!beforeList || beforeList.length < 1) {
        return afterList[afterList.length - 1];
    }

    for (var a = 0; a < afterList.length; a++) {
        var found = false;
        for (var b = 0; b < beforeList.length; b++) {
            if (afterList[a] === beforeList[b]) {
                found = true;
                break;
            }
        }
        if (!found) {
            return afterList[a];
        }
    }

    return afterList[afterList.length - 1];
}

// Tagged effect strategy keeps a stable component identity across localized hosts.
function _gridMaker_ensureTaggedEffect(clip, qClip, type, lookupNames, findAnyFn, findTaggedFn, debugLines) {
    var tagged = findTaggedFn(clip);
    if (tagged) {
        _gridMaker_debugPush(debugLines, type + " tagged component found: " + _gridMaker_componentLabel(tagged));
        return tagged;
    }

    var beforeType = _gridMaker_getTypeComponents(clip, type);
    _gridMaker_debugPush(debugLines, type + " components before ensure=" + beforeType.length);

    var addedOnce = false;
    for (var i = 0; i < lookupNames.length; i++) {
        var effectName = lookupNames[i];
        var fx = null;
        try {
            fx = qe.project.getVideoEffectByName(effectName);
        } catch (e1) {}
        if (!fx) {
            continue;
        }

        try {
            qClip.addVideoEffect(fx);
        } catch (e2) {
            continue;
        }

        var afterType = _gridMaker_getTypeComponents(clip, type);
        var candidate = _gridMaker_pickNewComponent(beforeType, afterType);
        if (!candidate) {
            candidate = findAnyFn(clip);
        }
        addedOnce = true;
        _gridMaker_debugPush(debugLines, "Added " + type + " via '" + effectName + "' candidate=" + _gridMaker_componentLabel(candidate));
        if (candidate) {
            var taggedOk = _gridMaker_tryTagComponent(candidate, type, debugLines);
            if (!taggedOk) {
                var baseName = _gridMaker_stripGridMakerTag(_gridMaker_componentDisplayName(candidate));
                if (!baseName) {
                    baseName = _gridMaker_defaultLabelForType(type);
                }
                _gridMaker_tryTagSingleTypeViaQE(qClip, type, baseName + _gridMaker_effectTagSuffix(), debugLines);
            }

            tagged = findTaggedFn(clip);
            if (tagged) {
                return tagged;
            }

            // Prevent multiple duplicate insertions across localized lookup names.
            break;
        }
        beforeType = afterType;
    }

    tagged = findTaggedFn(clip);
    if (tagged) {
        return tagged;
    }

    if (addedOnce) {
        var anyNow = findAnyFn(clip);
        if (anyNow) {
            _gridMaker_debugPush(debugLines, type + " added but untaggable; using detected component: " + _gridMaker_componentLabel(anyNow));
            return anyNow;
        }
    }

    var byType = _gridMaker_getTypeComponents(clip, type);
    if (byType.length === 1) {
        _gridMaker_debugPush(debugLines, type + " tagging unavailable; using sole " + type + " component safely: " + _gridMaker_componentLabel(byType[0]));
        return byType[0];
    }

    _gridMaker_debugPush(debugLines, type + " tagged component unavailable and ambiguous count=" + byType.length);
    return null;
}

function _gridMaker_isTruthyToggleValue(value) {
    if (value === true) {
        return true;
    }
    if (value === false || value === null || value === undefined) {
        return false;
    }
    var n = _gridMaker_toNumber(value);
    if (_gridMaker_isFiniteNumber(n)) {
        return n !== 0;
    }
    return false;
}

function _gridMaker_trySetToggleProperty(prop, value, debugLines, label) {
    if (!prop) {
        _gridMaker_debugPush(debugLines, "set-toggle skip " + (label || "?") + " missing prop");
        return false;
    }

    try {
        prop.setValue(value, true);
        var readback = _gridMaker_readPropertyValue(prop);
        var ok = _gridMaker_isTruthyToggleValue(readback);
        _gridMaker_debugPush(debugLines, "set-toggle attempt " + (label || "?") + " value=" + value + " readback=" + readback + " ok=" + ok);
        if (ok) {
            return true;
        }
    } catch (e1) {}

    _gridMaker_disableTimeVarying(prop);
    try {
        prop.setValue(value, true);
        var readback2 = _gridMaker_readPropertyValue(prop);
        var ok2 = _gridMaker_isTruthyToggleValue(readback2);
        _gridMaker_debugPush(debugLines, "set-toggle retry " + (label || "?") + " value=" + value + " readback=" + readback2 + " ok=" + ok2);
        return ok2;
    } catch (e2) {}

    _gridMaker_debugPush(debugLines, "set-toggle failed " + (label || "?"));
    return false;
}

function _gridMaker_enableTransformUniformScale(transformComp, debugLines) {
    if (!transformComp) {
        return;
    }

    var uniform = _gridMaker_findProperty(
        transformComp,
        [
            "uniform scale",
            "echelle uniforme",
            "échelle uniforme",
            "escala uniforme",
            "scala uniforme",
            "adbe geometry2-0003"
        ]
    );

    if (!uniform) {
        _gridMaker_debugPush(debugLines, "Transform uniform scale property not found");
        return;
    }

    if (_gridMaker_isTruthyToggleValue(_gridMaker_readPropertyValue(uniform))) {
        _gridMaker_debugPush(debugLines, "Transform uniform scale already enabled");
        return;
    }

    var ok = _gridMaker_trySetToggleProperty(uniform, true, debugLines, "transform.uniformScale=true");
    if (!ok) {
        ok = _gridMaker_trySetToggleProperty(uniform, 1, debugLines, "transform.uniformScale=1");
    }
    if (ok) {
        _gridMaker_debugPush(debugLines, "Transform uniform scale enabled readback=" + _gridMaker_readPropertyValue(uniform));
        return;
    }

    // Some hosts require a hard toggle cycle.
    _gridMaker_debugPush(debugLines, "Transform uniform initial enable failed; trying explicit off/on cycle");
    _gridMaker_trySetToggleProperty(uniform, 0, debugLines, "transform.uniformScale=0");
    _gridMaker_sleepSafe(40);
    ok = _gridMaker_trySetToggleProperty(uniform, true, debugLines, "transform.uniformScale=true.retry");
    if (!ok) {
        ok = _gridMaker_trySetToggleProperty(uniform, 1, debugLines, "transform.uniformScale=1.retry");
    }
    if (ok) {
        _gridMaker_debugPush(debugLines, "Transform uniform scale retry enabled readback=" + _gridMaker_readPropertyValue(uniform));
        return;
    }

    _gridMaker_debugPush(debugLines, "Transform uniform scale enforcement failed readback=" + _gridMaker_readPropertyValue(uniform));
}

function _gridMaker_transformSyncScaleAxes(widthProp, heightProp, debugLines) {
    if (!widthProp || !heightProp) {
        return;
    }
    var w = _gridMaker_toNumber(_gridMaker_readPropertyValue(widthProp));
    var h = _gridMaker_toNumber(_gridMaker_readPropertyValue(heightProp));
    var base = 100;
    if (_gridMaker_isFiniteNumber(w)) {
        base = w;
    } else if (_gridMaker_isFiniteNumber(h)) {
        base = h;
    }
    _gridMaker_trySetNumberProperty(widthProp, base, debugLines, "transform.scaleWidth.sync");
    _gridMaker_trySetNumberProperty(heightProp, base, debugLines, "transform.scaleHeight.sync");
}

function _gridMaker_transformUniformLinkIsEffective(uniformProp, widthProp, heightProp, debugLines) {
    if (!uniformProp || !widthProp || !heightProp) {
        _gridMaker_debugPush(debugLines, "Transform uniform linkage test skipped (missing props)");
        return _gridMaker_isTruthyToggleValue(_gridMaker_readPropertyValue(uniformProp));
    }

    var w0 = _gridMaker_toNumber(_gridMaker_readPropertyValue(widthProp));
    var h0 = _gridMaker_toNumber(_gridMaker_readPropertyValue(heightProp));
    if (!_gridMaker_isFiniteNumber(w0) || !_gridMaker_isFiniteNumber(h0)) {
        _gridMaker_debugPush(debugLines, "Transform uniform linkage test skipped (invalid numeric readback)");
        return _gridMaker_isTruthyToggleValue(_gridMaker_readPropertyValue(uniformProp));
    }

    var testVal = w0 + 0.1234;
    _gridMaker_trySetNumberProperty(widthProp, testVal, debugLines, "transform.scaleWidth.linkTest");
    var h1 = _gridMaker_toNumber(_gridMaker_readPropertyValue(heightProp));
    var w1 = _gridMaker_toNumber(_gridMaker_readPropertyValue(widthProp));
    var linked = _gridMaker_isFiniteNumber(h1) && _gridMaker_isFiniteNumber(w1) && Math.abs(h1 - w1) < 0.0001;

    _gridMaker_trySetNumberProperty(widthProp, w0, debugLines, "transform.scaleWidth.restore");
    _gridMaker_trySetNumberProperty(heightProp, h0, debugLines, "transform.scaleHeight.restore");

    _gridMaker_debugPush(debugLines, "Transform uniform linkage test linked=" + linked + " w1=" + w1 + " h1=" + h1);
    return linked;
}

// Effect/component lookup tables and matching helpers.
function _gridMaker_transformEffectLookupNames() {
    return [
        "Transformer",
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

function _gridMaker_roundedCropEffectLookupNames() {
    return [
        "Recadrage arrondi",
        "Rounded Crop",
        "AE.Impact_Crop_FX",
        "Impact Crop"
    ];
}

function _gridMaker_classicCropEffectLookupNames() {
    return [
        "Crop",
        "Recorte",
        "Recadrage",
        "Recortar",
        "Ritaglia",
        "Freistellen",
        "ADBE Crop",
        "AE.ADBE Crop",
        "AE.ADBE AECrop"
    ];
}

function _gridMaker_cropEffectLookupNames(preferRounded) {
    var useRounded = (preferRounded === true) || (preferRounded === undefined && _gridMaker_supportsRoundedCropEffect());
    if (!useRounded) {
        return _gridMaker_classicCropEffectLookupNames();
    }

    // Prefer Rounded Crop first, then fallback to classic Crop in case lookup fails.
    return _gridMaker_roundedCropEffectLookupNames().concat(_gridMaker_classicCropEffectLookupNames());
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
        ["crop", "rounded crop", "recadr", "recortar", "recorte", "ritagli", "freistell"],
        ["adbe crop", "ae.adbe crop", "ae.impact_crop_fx", "impact_crop_fx"],
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

// Property read/write helpers with defensive checks and readback validation.
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

function _gridMaker_trySetNumberProperty(prop, value, debugLines, label) {
    if (!prop || !_gridMaker_isFiniteNumber(value)) {
        _gridMaker_debugPush(debugLines, "set-number skip " + (label || "?") + " invalid prop/value");
        return false;
    }

    try {
        prop.setValue(value, true);
        var readback = _gridMaker_readPropertyValue(prop);
        if (_gridMaker_isFiniteNumber(readback)) {
            _gridMaker_debugPush(debugLines, "set-number ok " + (label || "?") + " value=" + value + " readback=" + readback);
            return true;
        }
    } catch (e1) {}

    _gridMaker_disableTimeVarying(prop);

    try {
        prop.setValue(value, true);
        var readback2 = _gridMaker_readPropertyValue(prop);
        var ok = _gridMaker_isFiniteNumber(readback2);
        _gridMaker_debugPush(debugLines, "set-number retry " + (label || "?") + " value=" + value + " readback=" + readback2 + " ok=" + ok);
        return ok;
    } catch (e2) {}

    _gridMaker_debugPush(debugLines, "set-number failed " + (label || "?"));
    return false;
}

function _gridMaker_trySetPointProperty(prop, point, debugLines, label) {
    if (!prop || !point || point.length < 2) {
        _gridMaker_debugPush(debugLines, "set-point skip " + (label || "?") + " invalid prop/value");
        return false;
    }

    var x = parseFloat(point[0]);
    var y = parseFloat(point[1]);
    if (!_gridMaker_isFiniteNumber(x) || !_gridMaker_isFiniteNumber(y)) {
        _gridMaker_debugPush(debugLines, "set-point skip " + (label || "?") + " non-finite point");
        return false;
    }

    try {
        prop.setValue([x, y], true);
        var readback = _gridMaker_readPropertyValue(prop);
        if (_gridMaker_isPointValue(readback)) {
            _gridMaker_debugPush(debugLines, "set-point ok " + (label || "?") + " value=[" + x + "," + y + "] readback=" + _gridMaker_pointToString(readback));
            return true;
        }
    } catch (e1) {}

    _gridMaker_disableTimeVarying(prop);

    try {
        prop.setValue([x, y], true);
        var readback2 = _gridMaker_readPropertyValue(prop);
        var ok = _gridMaker_isPointValue(readback2);
        _gridMaker_debugPush(debugLines, "set-point retry " + (label || "?") + " value=[" + x + "," + y + "] readback=" + _gridMaker_pointToString(readback2) + " ok=" + ok);
        return ok;
    } catch (e2) {}

    _gridMaker_debugPush(debugLines, "set-point failed " + (label || "?"));
    return false;
}

// Source metadata + sizing helpers used to compute target scale/crop/position.
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

function _gridMaker_getClipNativeFrameSize(clip, qClip, debugLines) {
    var candidates = [];
    var projectItem = null;

    // Prefer QE clip dimensions first: they often reflect timeline-effective media sizing
    // (for example when Premiere applies internal frame-scaling behavior).
    if (qClip) {
        _gridMaker_pushSizeCandidate(candidates, _gridMaker_toNumber(qClip.width), _gridMaker_toNumber(qClip.height), "qClip.size", debugLines);
        _gridMaker_pushSizeCandidate(candidates, _gridMaker_toNumber(qClip.videoFrameWidth), _gridMaker_toNumber(qClip.videoFrameHeight), "qClip.videoFrame", debugLines);
        _gridMaker_pushSizeCandidate(candidates, _gridMaker_toNumber(qClip.sourceWidth), _gridMaker_toNumber(qClip.sourceHeight), "qClip.source", debugLines);
    }

    try {
        projectItem = clip.projectItem;
    } catch (e1) {}

    if (projectItem) {
        _gridMaker_pushSizeCandidate(candidates, _gridMaker_toNumber(projectItem.videoFrameWidth), _gridMaker_toNumber(projectItem.videoFrameHeight), "projectItem.videoFrame", debugLines);
        _gridMaker_pushSizeCandidate(candidates, _gridMaker_toNumber(projectItem.mediaWidth), _gridMaker_toNumber(projectItem.mediaHeight), "projectItem.media", debugLines);
        _gridMaker_pushSizeCandidate(candidates, _gridMaker_toNumber(projectItem.width), _gridMaker_toNumber(projectItem.height), "projectItem.size", debugLines);

        try {
            var projectMetadata = projectItem.getProjectMetadata && projectItem.getProjectMetadata();
            var metaSize = _gridMaker_extractSizeFromMetadata(projectMetadata);
            if (metaSize) {
                _gridMaker_pushSizeCandidate(candidates, metaSize.width, metaSize.height, "projectItem.projectMetadata", debugLines);
            }
        } catch (e2) {}

        try {
            var xmpMetadata = projectItem.getXMPMetadata && projectItem.getXMPMetadata();
            var xmpSize = _gridMaker_extractSizeFromMetadata(xmpMetadata);
            if (xmpSize) {
                _gridMaker_pushSizeCandidate(candidates, xmpSize.width, xmpSize.height, "projectItem.xmpMetadata", debugLines);
            }
        } catch (e3) {}
    }

    if (candidates.length > 0) {
        _gridMaker_debugPush(debugLines, "Native clip size selected " + candidates[0].width + "x" + candidates[0].height + " from " + candidates[0].source);
        return { width: candidates[0].width, height: candidates[0].height };
    }

    _gridMaker_debugPush(debugLines, "Native clip size unavailable, fallback baseScale=100");
    return null;
}

function _gridMaker_computeBaseFitScale(frameW, frameH, nativeSize) {
    if (!nativeSize || !_gridMaker_isReasonableFrameSize(nativeSize.width, nativeSize.height)) {
        return 100.0;
    }

    var sx = frameW / nativeSize.width;
    var sy = frameH / nativeSize.height;
    var fitScale = Math.max(sx, sy) * 100.0;
    if (!_gridMaker_isFiniteNumber(fitScale) || fitScale <= 0) {
        return 100.0;
    }
    return fitScale;
}

function _gridMaker_pushSizeCandidate(candidates, width, height, source, debugLines) {
    if (!_gridMaker_isReasonableFrameSize(width, height)) {
        return;
    }
    candidates.push({ width: width, height: height, source: source });
    _gridMaker_debugPush(debugLines, "Native size candidate " + width + "x" + height + " from " + source);
}

function _gridMaker_extractSizeFromMetadata(text) {
    if (!text || typeof text !== "string") {
        return null;
    }

    var inlinePair = /([0-9]{2,5})\s*[xX]\s*([0-9]{2,5})/.exec(text);
    if (inlinePair) {
        var pW = _gridMaker_toNumber(inlinePair[1]);
        var pH = _gridMaker_toNumber(inlinePair[2]);
        if (_gridMaker_isReasonableFrameSize(pW, pH)) {
            return { width: pW, height: pH };
        }
    }

    var width = _gridMaker_findFirstNumericMatch(text, [
        /Column\.Intrinsic\.(?:Media|Video)Width[^0-9]{0,24}([0-9]{2,5})/i,
        /(?:Media|Video|Frame|Source)Width[^0-9]{0,24}([0-9]{2,5})/i,
        /\bwidth\b[^0-9]{0,24}([0-9]{2,5})/i
    ]);
    var height = _gridMaker_findFirstNumericMatch(text, [
        /Column\.Intrinsic\.(?:Media|Video)Height[^0-9]{0,24}([0-9]{2,5})/i,
        /(?:Media|Video|Frame|Source)Height[^0-9]{0,24}([0-9]{2,5})/i,
        /\bheight\b[^0-9]{0,24}([0-9]{2,5})/i
    ]);

    if (_gridMaker_isReasonableFrameSize(width, height)) {
        return { width: width, height: height };
    }

    return null;
}

function _gridMaker_findFirstNumericMatch(text, patterns) {
    for (var i = 0; i < patterns.length; i++) {
        var match = patterns[i].exec(text);
        if (!match || match.length < 2) {
            continue;
        }
        var value = _gridMaker_toNumber(match[1]);
        if (_gridMaker_isFiniteNumber(value)) {
            return value;
        }
    }
    return NaN;
}

function _gridMaker_setPlacement(component, scale, x, y, frameW, frameH, debugLines) {
    var kind = _gridMaker_componentKind(component);
    _gridMaker_debugPush(debugLines, "Placement component=" + _gridMaker_componentLabel(component) + " kind=" + kind);

    if (kind === "transform") {
        var uniform = _gridMaker_findProperty(component, ["uniform scale", "echelle uniforme", "uniform"]);
        if (uniform) {
            _gridMaker_trySetNumberProperty(uniform, 1, debugLines, "uniform");
        }
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
    _gridMaker_debugPush(debugLines, "Scale property=" + _gridMaker_propertyLabel(scaleProp) + " target=" + scale);
    _gridMaker_disableTimeVarying(scaleProp);
    var scaleOk = _gridMaker_trySetNumberProperty(scaleProp, scale, debugLines, "scale");

    var position = _gridMaker_findProperty(component, [
        "position",
        "adbe transform position",
        "adbe position",
        "adbe motion position"
    ], "point2d");
    if (!position) {
        _gridMaker_debugPush(debugLines, "No 2D position property found");
        return false;
    }
    var currentPos = _gridMaker_readPropertyValue(position);
    var positionMode = _gridMaker_detectPositionMode(kind, currentPos, frameW, frameH);
    _gridMaker_debugPush(debugLines, "Position property=" + _gridMaker_propertyLabel(position) + " current=" + _gridMaker_pointToString(currentPos) + " mode=" + positionMode + " targetAbs=[" + x + "," + y + "]");
    _gridMaker_disableTimeVarying(position);

    var candidates = _gridMaker_buildPositionCandidates(positionMode, x, y, frameW, frameH);
    var positionOk = false;
    for (var i = 0; i < candidates.length; i++) {
        _gridMaker_debugPush(debugLines, "Position candidate #" + (i + 1) + "=" + _gridMaker_pointToString(candidates[i]));
        if (_gridMaker_trySetPointProperty(position, candidates[i], debugLines, "position#" + (i + 1))) {
            var readback = _gridMaker_readPropertyValue(position);
            if (_gridMaker_isPointValue(readback)) {
                var rx = parseFloat(readback[0]);
                var ry = parseFloat(readback[1]);
                if (_gridMaker_isFiniteNumber(rx) && _gridMaker_isFiniteNumber(ry) && Math.abs(rx) < 100000 && Math.abs(ry) < 100000) {
                    _gridMaker_debugPush(debugLines, "Position accepted readback=" + _gridMaker_pointToString(readback));
                    positionOk = true;
                    break;
                }
            } else {
                _gridMaker_debugPush(debugLines, "Position accepted without readable point value");
                positionOk = true;
                break;
            }
        }
    }

    _gridMaker_debugPush(debugLines, "Placement result scaleOk=" + scaleOk + " positionOk=" + positionOk);
    return !!scaleOk && !!positionOk;
}

function _gridMaker_buildPositionCandidates(mode, x, y, frameW, frameH) {
    var absolute = [x, y];
    var centered = [x - (frameW * 0.5), y - (frameH * 0.5)];
    var normalizedAbs = [x / frameW, y / frameH];
    var normalizedCentered = [centered[0] / frameW, centered[1] / frameH];

    if (mode === "transform_centered") {
        return [centered, absolute, normalizedCentered, normalizedAbs];
    }
    if (mode === "motion_normalized") {
        return [normalizedAbs, absolute, centered, normalizedCentered];
    }
    if (mode === "motion_pixels") {
        return [absolute, normalizedAbs, centered, normalizedCentered];
    }
    if (mode === "unknown_normalized") {
        return [normalizedAbs, absolute, centered, normalizedCentered];
    }
    return [absolute, centered, normalizedAbs, normalizedCentered];
}

function _gridMaker_detectPositionMode(kind, currentPos, frameW, frameH) {
    if (kind === "transform") {
        return "transform_centered";
    }

    if (kind === "motion") {
        if (_gridMaker_isLikelyNormalizedPoint(currentPos)) {
            return "motion_normalized";
        }
        return "motion_pixels";
    }

    if (_gridMaker_isLikelyNormalizedPoint(currentPos)) {
        return "unknown_normalized";
    }
    return "unknown_pixels";
}

function _gridMaker_isLikelyNormalizedPoint(point) {
    if (!_gridMaker_isPointValue(point)) {
        return false;
    }

    var x = parseFloat(point[0]);
    var y = parseFloat(point[1]);

    if (!_gridMaker_isFiniteNumber(x) || !_gridMaker_isFiniteNumber(y)) {
        return false;
    }

    // Premiere variants can use normalized coordinates around 0..1 for motion.
    return x >= -0.1 && x <= 1.1 && y >= -0.1 && y <= 1.1;
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

function _gridMaker_componentLabel(component) {
    if (!component) {
        return "<none>";
    }

    var displayName = "";
    var matchName = "";
    try {
        displayName = component.displayName || "";
    } catch (e1) {}
    try {
        matchName = component.matchName || "";
    } catch (e2) {}

    return "[" + displayName + "|" + matchName + "]";
}

function _gridMaker_propertyLabel(prop) {
    if (!prop) {
        return "<none>";
    }

    var displayName = "";
    var matchName = "";
    try {
        displayName = prop.displayName || "";
    } catch (e1) {}
    try {
        matchName = prop.matchName || "";
    } catch (e2) {}
    return "[" + displayName + "|" + matchName + "]";
}

function _gridMaker_getCurrentPosition(component) {
    var position = _gridMaker_findProperty(component, [
        "position",
        "adbe transform position",
        "adbe position",
        "adbe motion position"
    ], "point2d");
    if (!position) {
        return null;
    }
    return _gridMaker_readPropertyValue(position);
}

function _gridMaker_pointToString(point) {
    if (!_gridMaker_isPointValue(point)) {
        return "<invalid>";
    }
    return "[" + parseFloat(point[0]) + "," + parseFloat(point[1]) + "]";
}

function _gridMaker_debugPush(debugLines, message) {
    if (!debugLines) {
        return;
    }
    debugLines.push(String(message));
}

function _gridMaker_dumpPlacementComponents(clip, debugLines, phase) {
    if (!clip || !clip.components) {
        return;
    }

    _gridMaker_debugPush(debugLines, phase + " COMPONENT SNAPSHOT START");
    for (var c = 0; c < clip.components.numItems; c++) {
        var comp = clip.components[c];
        if (!comp) {
            continue;
        }

        var compLabel = _gridMaker_componentLabel(comp);
        var posProp = _gridMaker_findProperty(comp, [
            "position",
            "adbe transform position",
            "adbe position",
            "adbe motion position"
        ], "point2d");
        var scaleProp = _gridMaker_findProperty(comp, [
            "scale",
            "echelle",
            "escala",
            "scala",
            "adbe transform scale",
            "adbe scale",
            "adbe motion scale"
        ], "number");

        var hasPlacement = !!posProp || !!scaleProp;
        if (!hasPlacement) {
            continue;
        }

        var posVal = posProp ? _gridMaker_pointToString(_gridMaker_readPropertyValue(posProp)) : "<none>";
        var scaleVal = scaleProp ? _gridMaker_readPropertyValue(scaleProp) : "<none>";
        _gridMaker_debugPush(
            debugLines,
            phase + " comp#" + c + " " + compLabel +
            " posProp=" + _gridMaker_propertyLabel(posProp) + " pos=" + posVal +
            " scaleProp=" + _gridMaker_propertyLabel(scaleProp) + " scale=" + scaleVal
        );
    }
    _gridMaker_debugPush(debugLines, phase + " COMPONENT SNAPSHOT END");
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

// Parse and clamp global margin (in sequence pixels) from UI payload.
function _gridMaker_parseMarginPx(rawMargin) {
    var margin = _gridMaker_toNumber(rawMargin);
    if (!_gridMaker_isFiniteNumber(margin) || margin < 0) {
        return 0;
    }
    if (margin > 10000) {
        margin = 10000;
    }
    return margin;
}

// Parse and clamp roundness percentage from UI payload.
function _gridMaker_parseRoundnessPercent(rawRoundness) {
    var value = _gridMaker_toNumber(rawRoundness);
    if (!_gridMaker_isFiniteNumber(value) || value < 0) {
        return 0;
    }
    if (value > 100) {
        value = 100;
    }
    return value;
}

// Compute final cell rectangle with a single slider controlling outer margin + inter-cell gap.
function _gridMaker_computePaddedCellRect(frameW, frameH, leftNorm, topNorm, widthNorm, heightNorm, marginPx, debugLines) {
    if (!(frameW > 0) || !(frameH > 0)) {
        return null;
    }
    if (!_gridMaker_isFiniteNumber(leftNorm) || !_gridMaker_isFiniteNumber(topNorm) || !_gridMaker_isFiniteNumber(widthNorm) || !_gridMaker_isFiniteNumber(heightNorm)) {
        return null;
    }
    if (!(widthNorm > 0) || !(heightNorm > 0)) {
        return null;
    }
    if (leftNorm < 0 || topNorm < 0 || leftNorm + widthNorm > 1.000001 || topNorm + heightNorm > 1.000001) {
        return null;
    }

    var safeMargin = _gridMaker_parseMarginPx(marginPx);
    var halfMargin = safeMargin * 0.5;
    var maxHalfMargin = Math.min(frameW, frameH) * 0.49;
    if (halfMargin > maxHalfMargin) {
        halfMargin = maxHalfMargin;
    }

    var contentLeft = halfMargin;
    var contentTop = halfMargin;
    var contentW = frameW - (halfMargin * 2.0);
    var contentH = frameH - (halfMargin * 2.0);
    if (!(contentW > 1) || !(contentH > 1)) {
        contentLeft = 0;
        contentTop = 0;
        contentW = frameW;
        contentH = frameH;
        halfMargin = 0;
    }

    var cellLeft = contentLeft + (contentW * leftNorm);
    var cellTop = contentTop + (contentH * topNorm);
    var cellW = contentW * widthNorm;
    var cellH = contentH * heightNorm;

    var localInset = halfMargin;
    var maxInset = Math.min(cellW, cellH) * 0.49;
    if (localInset > maxInset) {
        localInset = maxInset;
    }
    if (!(localInset >= 0)) {
        localInset = 0;
    }

    cellLeft += localInset;
    cellTop += localInset;
    cellW -= localInset * 2.0;
    cellH -= localInset * 2.0;

    if (!(cellW > 1) || !(cellH > 1)) {
        _gridMaker_debugPush(debugLines, "Margin collapsed cell dimensions (cellW=" + cellW + " cellH=" + cellH + ")");
        return null;
    }

    return {
        left: cellLeft,
        top: cellTop,
        width: cellW,
        height: cellH,
        centerX: cellLeft + (cellW * 0.5),
        centerY: cellTop + (cellH * 0.5)
    };
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

// Placement and crop write helpers for Motion/Crop components.
function _gridMaker_setCrop(component, left, right, top, bottom) {
    if (!component || !_gridMaker_componentMatchesType(component, "crop")) {
        return;
    }

    var pLeft = _gridMaker_findProperty(component, [
        "left",
        "gauche",
        "izquierda",
        "sinistra",
        "links",
        "esquerda",
        "adbe crop left"
    ]);
    var pRight = _gridMaker_findProperty(component, [
        "right",
        "droite",
        "derecha",
        "destra",
        "rechts",
        "direita",
        "adbe crop right"
    ]);
    var pTop = _gridMaker_findProperty(component, [
        "top",
        "haut",
        "superior",
        "alto",
        "oben",
        "topo",
        "adbe crop top"
    ]);
    var pBottom = _gridMaker_findProperty(component, [
        "bottom",
        "bas",
        "inferior",
        "basso",
        "unten",
        "baixo",
        "adbe crop bottom"
    ]);

    if (pLeft) {
        _gridMaker_disableTimeVarying(pLeft);
        try {
            pLeft.setValue(left, true);
        } catch (e1) {}
    }
    if (pRight) {
        _gridMaker_disableTimeVarying(pRight);
        try {
            pRight.setValue(right, true);
        } catch (e2) {}
    }
    if (pTop) {
        _gridMaker_disableTimeVarying(pTop);
        try {
            pTop.setValue(top, true);
        } catch (e3) {}
    }
    if (pBottom) {
        _gridMaker_disableTimeVarying(pBottom);
        try {
            pBottom.setValue(bottom, true);
        } catch (e4) {}
    }
}

function _gridMaker_findCropRoundnessProperty(component) {
    return _gridMaker_findProperty(component, [
        "roundness",
        "arrondi",
        "redondez",
        "abrundung",
        "arredondamento",
        "arrotondamento",
        "roundness top left",
        "arrondi coin haut gauche"
    ], "number");
}

// Apply Rounded Crop roundness when the property exists. No-op on classic Crop.
function _gridMaker_setCropRoundness(component, roundnessPct, debugLines) {
    if (!component || !_gridMaker_componentMatchesType(component, "crop")) {
        return;
    }

    var target = _gridMaker_parseRoundnessPercent(roundnessPct);
    var roundnessProp = _gridMaker_findCropRoundnessProperty(component);
    if (!roundnessProp) {
        _gridMaker_debugPush(debugLines, "Roundness property not found on crop component; skipping");
        return;
    }

    _gridMaker_disableTimeVarying(roundnessProp);
    var ok = _gridMaker_trySetNumberProperty(roundnessProp, target, debugLines, "crop.roundness");
    if (!ok) {
        _gridMaker_debugPush(debugLines, "Roundness write failed");
        return;
    }

    // Mirror corner roundness controls when available so the effect stays uniformly rounded.
    var cornerProps = [
        _gridMaker_findProperty(component, ["roundness top left"], "number"),
        _gridMaker_findProperty(component, ["roundness top right"], "number"),
        _gridMaker_findProperty(component, ["roundness bottom left"], "number"),
        _gridMaker_findProperty(component, ["roundness bottom right"], "number")
    ];
    for (var i = 0; i < cornerProps.length; i++) {
        if (!cornerProps[i]) {
            continue;
        }
        _gridMaker_disableTimeVarying(cornerProps[i]);
        _gridMaker_trySetNumberProperty(cornerProps[i], target, debugLines, "crop.roundness.corner#" + (i + 1));
    }
}

// Read current Crop values (percent) from a Crop component. Missing component means no crop.
function _gridMaker_getCropValues(component) {
    var out = { left: 0, right: 0, top: 0, bottom: 0 };
    if (!component || !_gridMaker_componentMatchesType(component, "crop")) {
        return out;
    }

    var pLeft = _gridMaker_findProperty(component, [
        "left",
        "gauche",
        "izquierda",
        "sinistra",
        "links",
        "esquerda",
        "adbe crop left"
    ], "number");
    var pRight = _gridMaker_findProperty(component, [
        "right",
        "droite",
        "derecha",
        "destra",
        "rechts",
        "direita",
        "adbe crop right"
    ], "number");
    var pTop = _gridMaker_findProperty(component, [
        "top",
        "haut",
        "superior",
        "alto",
        "oben",
        "topo",
        "adbe crop top"
    ], "number");
    var pBottom = _gridMaker_findProperty(component, [
        "bottom",
        "bas",
        "inferior",
        "basso",
        "unten",
        "baixo",
        "adbe crop bottom"
    ], "number");

    var left = _gridMaker_toNumber(_gridMaker_readPropertyValue(pLeft));
    var right = _gridMaker_toNumber(_gridMaker_readPropertyValue(pRight));
    var top = _gridMaker_toNumber(_gridMaker_readPropertyValue(pTop));
    var bottom = _gridMaker_toNumber(_gridMaker_readPropertyValue(pBottom));

    if (_gridMaker_isFiniteNumber(left)) {
        out.left = _gridMaker_clamp(left, 0, 99);
    }
    if (_gridMaker_isFiniteNumber(right)) {
        out.right = _gridMaker_clamp(right, 0, 99);
    }
    if (_gridMaker_isFiniteNumber(top)) {
        out.top = _gridMaker_clamp(top, 0, 99);
    }
    if (_gridMaker_isFiniteNumber(bottom)) {
        out.bottom = _gridMaker_clamp(bottom, 0, 99);
    }

    return out;
}

// Convert the current Position property readback to absolute sequence pixels.
function _gridMaker_positionToAbsolute(point, mode, frameW, frameH) {
    if (!_gridMaker_isPointValue(point)) {
        return null;
    }

    var x = _gridMaker_toNumber(point[0]);
    var y = _gridMaker_toNumber(point[1]);
    if (!_gridMaker_isFiniteNumber(x) || !_gridMaker_isFiniteNumber(y)) {
        return null;
    }

    if (mode === "motion_normalized" || mode === "unknown_normalized") {
        return [x * frameW, y * frameH];
    }
    if (mode === "transform_centered") {
        return [x + (frameW * 0.5), y + (frameH * 0.5)];
    }
    return [x, y];
}

// Result + JSON helpers used by all hostscript public endpoints.
function _gridMaker_clamp(v, min, max) {
    if (v < min) {
        return min;
    }
    if (v > max) {
        return max;
    }
    return v;
}

function _gridMaker_result(status, code, details, debugLines) {
    var payload = status + "|" + code;
    var mergedDetails = {};
    var hasDetail = false;

    if (details) {
        for (var key in details) {
            if (!details.hasOwnProperty(key)) {
                continue;
            }
            mergedDetails[key] = details[key];
            hasDetail = true;
        }
    }

    if (debugLines && debugLines.length > 0) {
        mergedDetails.debug = debugLines.join("\n");
        hasDetail = true;
    }

    var detailString = _gridMaker_serializeDetails(hasDetail ? mergedDetails : null);
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

function _gridMaker_jsonStringify(value) {
    try {
        return JSON.stringify(value);
    } catch (e1) {
        return _gridMaker_jsonStringifyFallback(value);
    }
}

function _gridMaker_jsonStringifyFallback(value) {
    // Keep host JSON responses working on ExtendScript engines without a JSON global.
    if (value === null) {
        return "null";
    }
    if (value === undefined) {
        return "null";
    }
    var t = typeof value;
    if (t === "boolean") {
        return value ? "true" : "false";
    }
    if (t === "number") {
        return isFinite(value) ? String(value) : "null";
    }
    if (t === "string") {
        return "\"" + _gridMaker_jsonEscapeString(value) + "\"";
    }
    if (value instanceof Array) {
        var arr = [];
        for (var i = 0; i < value.length; i++) {
            arr.push(_gridMaker_jsonStringifyFallback(value[i]));
        }
        return "[" + arr.join(",") + "]";
    }
    if (t === "object") {
        var parts = [];
        for (var key in value) {
            if (!value.hasOwnProperty(key)) {
                continue;
            }
            var item = value[key];
            if (typeof item === "undefined" || typeof item === "function") {
                continue;
            }
            parts.push("\"" + _gridMaker_jsonEscapeString(key) + "\":" + _gridMaker_jsonStringifyFallback(item));
        }
        return "{" + parts.join(",") + "}";
    }
    return "null";
}

function _gridMaker_jsonEscapeString(value) {
    // Escape only JSON-sensitive characters so simple capability payloads parse reliably.
    return String(value)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, "\\\"")
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n")
        .replace(/\t/g, "\\t");
}

function _gridMaker_jsonParse(text) {
    if (!text) {
        return null;
    }
    try {
        return JSON.parse(text);
    } catch (e1) {}
    try {
        return eval("(" + text + ")");
    } catch (e2) {}
    return null;
}

// Designer config storage in userData (per-ratio list/save/delete).
function _gridMaker_designerSettingsFolder() {
    var base = Folder.userData;
    if (!base) {
        return null;
    }
    var root = new Folder(base.fsName + "/PremiereGridMaker/settings");
    if (!root.exists) {
        root.create();
    }
    return root;
}

function _gridMaker_designerStoreFile() {
    var folder = _gridMaker_designerSettingsFolder();
    if (!folder) {
        return null;
    }
    return new File(folder.fsName + "/designer-configs.json");
}

function _gridMaker_designerRatioKey(ratioW, ratioH) {
    ratioW = _gridMaker_toNumber(ratioW);
    ratioH = _gridMaker_toNumber(ratioH);
    if (!_gridMaker_isFiniteNumber(ratioW) || !_gridMaker_isFiniteNumber(ratioH) || !(ratioW > 0) || !(ratioH > 0)) {
        return "16x9";
    }
    return String(Math.round(ratioW * 1000)) + "x" + String(Math.round(ratioH * 1000));
}

function _gridMaker_designerReadStore() {
    var file = _gridMaker_designerStoreFile();
    if (!file) {
        return { configs: [] };
    }
    if (!file.exists) {
        return { configs: [] };
    }
    try {
        file.encoding = "UTF-8";
        file.open("r");
        var raw = file.read();
        file.close();
        var parsed = _gridMaker_jsonParse(raw);
        if (!parsed || !parsed.configs || !(parsed.configs instanceof Array)) {
            return { configs: [] };
        }
        return parsed;
    } catch (e1) {
        try { file.close(); } catch (e2) {}
        return { configs: [] };
    }
}

function _gridMaker_designerWriteStore(store) {
    var file = _gridMaker_designerStoreFile();
    if (!file) {
        return false;
    }
    try {
        file.encoding = "UTF-8";
        file.open("w");
        file.write(_gridMaker_jsonStringify(store));
        file.close();
        return true;
    } catch (e1) {
        try { file.close(); } catch (e2) {}
        return false;
    }
}

function _gridMaker_designerSanitizeId(raw) {
    var base = String(raw || "").replace(/[^a-zA-Z0-9_-]/g, "");
    if (!base) {
        base = "cfg_" + (new Date().getTime());
    }
    return base;
}

function _gridMaker_designerNormalizeBlocks(rawBlocks) {
    var out = [];
    var minSize = 0.1;
    if (!(rawBlocks instanceof Array)) {
        return out;
    }
    for (var i = 0; i < rawBlocks.length; i++) {
        var b = rawBlocks[i];
        if (!b) {
            continue;
        }
        var x = _gridMaker_toNumber(b.x);
        var y = _gridMaker_toNumber(b.y);
        var w = _gridMaker_toNumber(b.w);
        var h = _gridMaker_toNumber(b.h);
        var id = _gridMaker_designerSanitizeId(b.id || ("cell_" + i));
        if (isNaN(x) || isNaN(y) || isNaN(w) || isNaN(h)) {
            continue;
        }
        if (w < minSize || h < minSize) {
            continue;
        }
        if (x < 0 || y < 0) {
            continue;
        }
        if (x + w > 10.000001 || y + h > 10.000001) {
            continue;
        }
        x = Math.round(x * 1000) / 1000;
        y = Math.round(y * 1000) / 1000;
        w = Math.round(w * 1000) / 1000;
        h = Math.round(h * 1000) / 1000;
        out.push({
            id: id,
            x: x,
            y: y,
            w: w,
            h: h
        });
    }
    return out;
}

function gridMaker_designerListConfigs(ratioW, ratioH) {
    try {
        var ratioKey = _gridMaker_designerRatioKey(ratioW, ratioH);
        var store = _gridMaker_designerReadStore();
        var list = [];
        for (var i = 0; i < store.configs.length; i++) {
            var cfg = store.configs[i];
            if (!cfg || cfg.ratioKey !== ratioKey) {
                continue;
            }
            list.push({
                id: cfg.id,
                name: cfg.name || cfg.id,
                ratioW: cfg.ratioW,
                ratioH: cfg.ratioH,
                // Return per-config margin so UI can restore spacing when loading presets.
                marginPx: _gridMaker_parseMarginPx(cfg.marginPx),
                // Return per-config roundness so UI can restore rounded crop value with each preset.
                roundness: _gridMaker_parseRoundnessPercent(cfg.roundness),
                blocks: cfg.blocks || [],
                updatedAt: cfg.updatedAt || ""
            });
        }
        // Keep user-defined visual order from store (do not auto-sort by updatedAt).
        return _gridMaker_jsonStringify({ ok: true, configs: list });
    } catch (e) {
        return _gridMaker_jsonStringify({ ok: false, message: String(e), configs: [] });
    }
}

// Persist a manual config order (drag & drop from panel preview cards) for one ratio.
function gridMaker_designerReorderConfigs(payloadJson) {
    try {
        var payload = _gridMaker_jsonParse(String(payloadJson || ""));
        if (!payload) {
            return _gridMaker_jsonStringify({ ok: false, message: "invalid_payload" });
        }

        var ratioW = _gridMaker_toNumber(payload.ratioW);
        var ratioH = _gridMaker_toNumber(payload.ratioH);
        if (!_gridMaker_isFiniteNumber(ratioW) || !_gridMaker_isFiniteNumber(ratioH) || !(ratioW > 0) || !(ratioH > 0)) {
            return _gridMaker_jsonStringify({ ok: false, message: "invalid_ratio" });
        }
        var ratioKey = _gridMaker_designerRatioKey(ratioW, ratioH);

        var orderedIds = (payload.orderedIds instanceof Array) ? payload.orderedIds : [];
        var uniqueIds = [];
        var seen = {};
        for (var i = 0; i < orderedIds.length; i++) {
            var id = _gridMaker_designerSanitizeId(orderedIds[i]);
            if (!id || seen[id]) {
                continue;
            }
            seen[id] = true;
            uniqueIds.push(id);
        }

        var store = _gridMaker_designerReadStore();
        var targetConfigs = [];
        var byId = {};
        for (i = 0; i < store.configs.length; i++) {
            var cfg = store.configs[i];
            if (!cfg || cfg.ratioKey !== ratioKey) {
                continue;
            }
            targetConfigs.push(cfg);
            byId[cfg.id] = cfg;
        }
        if (targetConfigs.length < 1) {
            return _gridMaker_jsonStringify({ ok: true, count: 0 });
        }

        var reordered = [];
        var used = {};
        for (i = 0; i < uniqueIds.length; i++) {
            var picked = byId[uniqueIds[i]];
            if (!picked) {
                continue;
            }
            reordered.push(picked);
            used[picked.id] = true;
        }
        // Append missing configs so accidental partial payloads cannot drop entries.
        for (i = 0; i < targetConfigs.length; i++) {
            if (used[targetConfigs[i].id]) {
                continue;
            }
            reordered.push(targetConfigs[i]);
        }

        // Replace only ratio-matching slots to preserve global store structure.
        var cursor = 0;
        for (i = 0; i < store.configs.length; i++) {
            if (!store.configs[i] || store.configs[i].ratioKey !== ratioKey) {
                continue;
            }
            store.configs[i] = reordered[cursor];
            cursor += 1;
        }

        if (!_gridMaker_designerWriteStore(store)) {
            return _gridMaker_jsonStringify({ ok: false, message: "write_failed" });
        }
        return _gridMaker_jsonStringify({ ok: true, count: reordered.length });
    } catch (e) {
        return _gridMaker_jsonStringify({ ok: false, message: String(e) });
    }
}

function gridMaker_designerSaveConfig(payloadJson) {
    try {
        var payload = _gridMaker_jsonParse(String(payloadJson || ""));
        if (!payload) {
            return _gridMaker_jsonStringify({ ok: false, message: "invalid_payload" });
        }

        var ratioW = _gridMaker_toNumber(payload.ratioW);
        var ratioH = _gridMaker_toNumber(payload.ratioH);
        if (!_gridMaker_isFiniteNumber(ratioW) || !_gridMaker_isFiniteNumber(ratioH) || !(ratioW > 0) || !(ratioH > 0)) {
            return _gridMaker_jsonStringify({ ok: false, message: "invalid_ratio" });
        }

        // Allow empty designer presets so the UI can persist a blank draft/grid intentionally.
        var blocks = _gridMaker_designerNormalizeBlocks(payload.blocks);

        var id = _gridMaker_designerSanitizeId(payload.id);
        var ratioKey = _gridMaker_designerRatioKey(ratioW, ratioH);
        var marginPx = _gridMaker_parseMarginPx(payload.marginPx);
        var roundness = _gridMaker_parseRoundnessPercent(payload.roundness);
        var now = (new Date()).toISOString ? (new Date()).toISOString() : String(new Date().getTime());
        var name = String(payload.name || "").replace(/^\s+|\s+$/g, "");
        if (!name) {
            name = "Config " + now;
        }

        var store = _gridMaker_designerReadStore();
        var found = false;
        for (var i = 0; i < store.configs.length; i++) {
            if (store.configs[i] && store.configs[i].id === id) {
                store.configs[i].name = name;
                store.configs[i].ratioW = ratioW;
                store.configs[i].ratioH = ratioH;
                store.configs[i].ratioKey = ratioKey;
                store.configs[i].marginPx = marginPx;
                store.configs[i].roundness = roundness;
                store.configs[i].blocks = blocks;
                store.configs[i].updatedAt = now;
                found = true;
                break;
            }
        }
        if (!found) {
            store.configs.push({
                id: id,
                name: name,
                ratioW: ratioW,
                ratioH: ratioH,
                ratioKey: ratioKey,
                marginPx: marginPx,
                roundness: roundness,
                blocks: blocks,
                createdAt: now,
                updatedAt: now
            });
        }

        if (!_gridMaker_designerWriteStore(store)) {
            return _gridMaker_jsonStringify({ ok: false, message: "write_failed" });
        }
        return _gridMaker_jsonStringify({ ok: true, id: id });
    } catch (e) {
        return _gridMaker_jsonStringify({ ok: false, message: String(e) });
    }
}

function gridMaker_designerDeleteConfig(configId) {
    try {
        var id = _gridMaker_designerSanitizeId(configId);
        var store = _gridMaker_designerReadStore();
        var next = [];
        var removed = false;
        for (var i = 0; i < store.configs.length; i++) {
            var cfg = store.configs[i];
            if (cfg && cfg.id === id) {
                removed = true;
                continue;
            }
            next.push(cfg);
        }
        store.configs = next;
        if (!_gridMaker_designerWriteStore(store)) {
            return _gridMaker_jsonStringify({ ok: false, message: "write_failed" });
        }
        return _gridMaker_jsonStringify({ ok: true, removed: removed });
    } catch (e) {
        return _gridMaker_jsonStringify({ ok: false, message: String(e) });
    }
}

// Export all saved designer configs to a JSON file for backup/sharing.
function gridMaker_designerExportConfigs() {
    try {
        var store = _gridMaker_designerReadStore();
        var exportPayload = {
            exportedAt: (new Date()).toISOString ? (new Date()).toISOString() : String(new Date().getTime()),
            source: "PremiereGridMaker",
            version: "1.2.6",
            configs: (store && store.configs instanceof Array) ? store.configs : []
        };

        var defaultFileName = "PremiereGridMaker-designer-configs.json";
        var target = File.saveDialog("Export Grid Maker designer configs", "*.json");
        if (!target) {
            return _gridMaker_jsonStringify({ ok: false, cancelled: true });
        }

        if (!(target instanceof File)) {
            target = new File(String(target || ""));
        }
        if (!target.name || target.name.indexOf(".json") === -1) {
            target = new File(target.fsName + ".json");
        }

        target.encoding = "UTF-8";
        target.open("w");
        target.write(_gridMaker_jsonStringify(exportPayload));
        target.close();

        return _gridMaker_jsonStringify({
            ok: true,
            count: exportPayload.configs.length,
            path: target.fsName || target.fullName || defaultFileName
        });
    } catch (e) {
        try { if (target) { target.close(); } } catch (e2) {}
        return _gridMaker_jsonStringify({ ok: false, message: String(e) });
    }
}

// Import designer configs from JSON and merge them into local storage by config id.
function gridMaker_designerImportConfigs() {
    try {
        var source = File.openDialog("Import Grid Maker designer configs", "*.json");
        if (!source) {
            return _gridMaker_jsonStringify({ ok: false, cancelled: true });
        }

        source.encoding = "UTF-8";
        source.open("r");
        var raw = source.read();
        source.close();

        var parsed = _gridMaker_jsonParse(raw);
        if (!parsed) {
            return _gridMaker_jsonStringify({ ok: false, message: "invalid_json" });
        }

        var rawConfigs = null;
        if (parsed instanceof Array) {
            rawConfigs = parsed;
        } else if (parsed.configs instanceof Array) {
            rawConfigs = parsed.configs;
        }
        if (!(rawConfigs instanceof Array)) {
            return _gridMaker_jsonStringify({ ok: false, message: "invalid_payload" });
        }

        var now = (new Date()).toISOString ? (new Date()).toISOString() : String(new Date().getTime());
        var normalized = [];
        for (var i = 0; i < rawConfigs.length; i++) {
            var cfg = rawConfigs[i] || {};
            var ratioW = _gridMaker_toNumber(cfg.ratioW);
            var ratioH = _gridMaker_toNumber(cfg.ratioH);
            if (!_gridMaker_isFiniteNumber(ratioW) || !_gridMaker_isFiniteNumber(ratioH) || !(ratioW > 0) || !(ratioH > 0)) {
                continue;
            }
            // Keep empty configs during import so shared preset packs can include blank drafts/templates.
            var blocks = _gridMaker_designerNormalizeBlocks(cfg.blocks);

            normalized.push({
                id: _gridMaker_designerSanitizeId(cfg.id || ("cfg_" + now + "_" + i)),
                name: String(cfg.name || ("Config " + (i + 1))),
                ratioW: ratioW,
                ratioH: ratioH,
                ratioKey: _gridMaker_designerRatioKey(ratioW, ratioH),
                // Keep per-config spacing during import/export cycles.
                marginPx: _gridMaker_parseMarginPx(cfg.marginPx),
                // Keep per-config roundness during import/export cycles.
                roundness: _gridMaker_parseRoundnessPercent(cfg.roundness),
                blocks: blocks,
                createdAt: String(cfg.createdAt || now),
                updatedAt: now
            });
        }

        if (normalized.length < 1) {
            return _gridMaker_jsonStringify({ ok: false, message: "no_valid_configs" });
        }

        var store = _gridMaker_designerReadStore();
        var nextById = {};
        var si;
        for (si = 0; si < store.configs.length; si++) {
            var existing = store.configs[si];
            if (!existing || !existing.id) {
                continue;
            }
            nextById[String(existing.id)] = existing;
        }
        for (si = 0; si < normalized.length; si++) {
            var incoming = normalized[si];
            var previous = nextById[incoming.id];
            if (previous && previous.createdAt) {
                incoming.createdAt = previous.createdAt;
            }
            nextById[incoming.id] = incoming;
        }

        var merged = [];
        for (var id in nextById) {
            if (!nextById.hasOwnProperty(id)) {
                continue;
            }
            merged.push(nextById[id]);
        }
        store.configs = merged;

        if (!_gridMaker_designerWriteStore(store)) {
            return _gridMaker_jsonStringify({ ok: false, message: "write_failed" });
        }

        return _gridMaker_jsonStringify({ ok: true, count: normalized.length });
    } catch (e) {
        try { if (source) { source.close(); } } catch (e2) {}
        return _gridMaker_jsonStringify({ ok: false, message: String(e) });
    }
}

// Safe external URL opener for trusted release ZIP links only.
function _gridMaker_isTrustedReleaseZipUrl(url) {
    if (!url) {
        return false;
    }
    var raw = String(url);
    return /^https:\/\/github\.com\/CyrilG93\/PremiereGridMaker\/releases\/(?:download\/v[0-9]+\.[0-9]+\.[0-9]+|latest\/download)\/[^?#]+\.zip(?:[?#].*)?$/i.test(raw);
}

function _gridMaker_escapeWindowsArg(value) {
    return String(value).replace(/"/g, "\"\"");
}

function _gridMaker_escapePosixArg(value) {
    return String(value).replace(/(["\\`$])/g, "\\$1");
}

function gridMaker_openExternalUrl(url) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        var targetUrl = String(url || "");
        dbg("OPEN-URL input=" + targetUrl);

        if (!_gridMaker_isTrustedReleaseZipUrl(targetUrl)) {
            dbg("OPEN-URL rejected (untrusted url)");
            return _gridMaker_result("ERR", "invalid_update_url", { url: targetUrl }, debugLines);
        }

        var os = "";
        try {
            os = $.os ? String($.os).toLowerCase() : "";
        } catch (eOs) {}
        dbg("OPEN-URL os=" + os);

        var cmd = "";
        if (os.indexOf("windows") !== -1) {
            cmd = 'cmd.exe /c start "" "' + _gridMaker_escapeWindowsArg(targetUrl) + '"';
        } else if (os.indexOf("mac") !== -1) {
            cmd = 'open "' + _gridMaker_escapePosixArg(targetUrl) + '"';
        } else {
            cmd = 'xdg-open "' + _gridMaker_escapePosixArg(targetUrl) + '"';
        }
        dbg("OPEN-URL cmd=" + cmd);

        var output = "";
        try {
            output = system.callSystem(cmd);
        } catch (eCall) {
            dbg("OPEN-URL callSystem exception=" + eCall);
            return _gridMaker_result("ERR", "open_url_failed", { message: eCall }, debugLines);
        }
        dbg("OPEN-URL result=" + output);

        return _gridMaker_result("OK", "url_opened", { url: targetUrl }, debugLines);
    } catch (e) {
        dbg("OPEN-URL exception=" + e);
        return _gridMaker_result("ERR", "open_url_exception", { message: e }, debugLines);
    }
}

function gridMaker_ping() {
    return _gridMaker_result("OK", "pong");
}
