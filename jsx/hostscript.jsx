function gridMaker_applyToSelectedClip(row, col, rows, cols, ratioW, ratioH) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT row=" + row + " col=" + col + " rows=" + rows + " cols=" + cols + " ratio=" + ratioW + ":" + ratioH);
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
        dbg("PARSED row=" + row + " col=" + col + " rows=" + rows + " cols=" + cols + " ratio=" + ratioW + ":" + ratioH);

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

        var transformComp = _gridMaker_findManagedTransformComponent(clip);
        var motionComp = _gridMaker_findMotionComponent(clip);
        var placementComp = motionComp;
        var cropComp = _gridMaker_findManagedCropComponent(clip);
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

        if (!transformComp) {
            if (!qSeq) {
                dbg("QE sequence unavailable (transform required)");
                return _gridMaker_result("ERR", "qe_unavailable", null, debugLines);
            }
            if (!qClip) {
                dbg("QE clip not found (transform required)");
                return _gridMaker_result("ERR", "qe_clip_not_found", null, debugLines);
            }
            transformComp = _gridMaker_ensureManagedEffect(
                clip,
                qClip,
                "transform",
                _gridMaker_transformEffectLookupNames(),
                debugLines
            );
            dbg("Transform component after ensure=" + _gridMaker_componentLabel(transformComp));
        }

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
                cropComp = _gridMaker_ensureManagedEffect(
                    clip,
                    qClip,
                    "crop",
                    _gridMaker_cropEffectLookupNames(),
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
        if (!transformComp) {
            dbg("Transform effect unavailable (required)");
            return _gridMaker_result("ERR", "transform_effect_unavailable", null, debugLines);
        }
        dbg("Transform strategy: required and kept as neutral effect (no parameter writes)");
        dbg("Placement strategy: Motion only");

        var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
        if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
            dbg("Invalid sequence frame size");
            return _gridMaker_result("ERR", "invalid_sequence_size", null, debugLines);
        }
        var frameW = frameSize.width;
        var frameH = frameSize.height;
        var frameAspect = frameW / frameH;
        var cellW = frameW / cols;
        var cellH = frameH / rows;
        var cellAspect = cellW / cellH;
        var preferHeightAxis = cellAspect <= 1.0;
        dbg("Frame size " + frameW + "x" + frameH + " aspect=" + frameAspect);
        dbg("Cell size " + cellW.toFixed(3) + "x" + cellH.toFixed(3) + " aspect=" + cellAspect.toFixed(6) + " preferHeightAxis=" + preferHeightAxis);

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

        // Some Premiere setups apply implicit "fit to frame" behavior before Motion scale.
        // In that case, Motion scale=100 maps to a frame-fitted size, not the raw native size.
        var intrinsicScaleFactor = 1.0;
        var assumeFrameFit = false;
        if (
            placementKind === "motion" &&
            placementModeHint === "motion_normalized" &&
            _gridMaker_isReasonableFrameSize(sourceW, sourceH) &&
            Math.abs((sourceW / sourceH) - frameAspect) > 0.0001
        ) {
            assumeFrameFit = true;
            intrinsicScaleFactor = Math.min(frameW / sourceW, frameH / sourceH);
        }
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

        var x = frameW * ((col + 0.5) / cols);
        var y = frameH * ((row + 0.5) / rows);
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

function gridMaker_applyToSelectedCustomCell(leftNorm, topNorm, widthNorm, heightNorm, ratioW, ratioH) {
    var debugLines = [];
    function dbg(message) {
        _gridMaker_debugPush(debugLines, message);
    }

    try {
        dbg("INPUT customCell left=" + leftNorm + " top=" + topNorm + " width=" + widthNorm + " height=" + heightNorm + " ratio=" + ratioW + ":" + ratioH);
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
        dbg("PARSED customCell left=" + leftNorm + " top=" + topNorm + " width=" + widthNorm + " height=" + heightNorm + " ratio=" + ratioW + ":" + ratioH);

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

        var transformComp = _gridMaker_findManagedTransformComponent(clip);
        var motionComp = _gridMaker_findMotionComponent(clip);
        var placementComp = motionComp;
        var cropComp = _gridMaker_findManagedCropComponent(clip);
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

        if (!transformComp) {
            if (!qSeq) {
                dbg("QE sequence unavailable (transform required)");
                return _gridMaker_result("ERR", "qe_unavailable", null, debugLines);
            }
            if (!qClip) {
                dbg("QE clip not found (transform required)");
                return _gridMaker_result("ERR", "qe_clip_not_found", null, debugLines);
            }
            transformComp = _gridMaker_ensureManagedEffect(
                clip,
                qClip,
                "transform",
                _gridMaker_transformEffectLookupNames(),
                debugLines
            );
            dbg("Transform component after ensure=" + _gridMaker_componentLabel(transformComp));
        }

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
                cropComp = _gridMaker_ensureManagedEffect(
                    clip,
                    qClip,
                    "crop",
                    _gridMaker_cropEffectLookupNames(),
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
        if (!transformComp) {
            dbg("Transform effect unavailable (required)");
            return _gridMaker_result("ERR", "transform_effect_unavailable", null, debugLines);
        }
        dbg("Transform strategy: required and kept as neutral effect (no parameter writes)");
        dbg("Placement strategy: Motion only");

        var frameSize = _gridMaker_getSequenceFrameSize(seq, qSeq);
        if (!frameSize || !_gridMaker_isFiniteNumber(frameSize.width) || !_gridMaker_isFiniteNumber(frameSize.height) || !(frameSize.width > 0) || !(frameSize.height > 0)) {
            dbg("Invalid sequence frame size");
            return _gridMaker_result("ERR", "invalid_sequence_size", null, debugLines);
        }
        var frameW = frameSize.width;
        var frameH = frameSize.height;
        var frameAspect = frameW / frameH;
        var cellW = frameW * widthNorm;
        var cellH = frameH * heightNorm;
        var cellAspect = cellW / cellH;
        var preferHeightAxis = cellAspect <= 1.0;
        dbg("Frame size " + frameW + "x" + frameH + " aspect=" + frameAspect);
        dbg("Custom cell size " + cellW.toFixed(3) + "x" + cellH.toFixed(3) + " aspect=" + cellAspect.toFixed(6) + " preferHeightAxis=" + preferHeightAxis);

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
        if (
            placementKind === "motion" &&
            placementModeHint === "motion_normalized" &&
            _gridMaker_isReasonableFrameSize(sourceW, sourceH) &&
            Math.abs((sourceW / sourceH) - frameAspect) > 0.0001
        ) {
            assumeFrameFit = true;
            intrinsicScaleFactor = Math.min(frameW / sourceW, frameH / sourceH);
        }
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

        var x = frameW * (leftNorm + widthNorm * 0.5);
        var y = frameH * (topNorm + heightNorm * 0.5);
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

function _gridMaker_findManagedCropComponent(clip) {
    return _gridMaker_findManagedEffectComponent(clip, "crop");
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

function _gridMaker_ensureManagedEffect(clip, qClip, type, lookupNames, debugLines) {
    var existing = _gridMaker_findManagedEffectComponent(clip, type);
    if (existing) {
        _gridMaker_debugPush(debugLines, type + " managed component found: " + _gridMaker_componentLabel(existing));
        return existing;
    }

    var beforeType = _gridMaker_getTypeComponents(clip, type);
    _gridMaker_debugPush(debugLines, type + " components before ensure=" + beforeType.length);

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
        if (!candidate && afterType.length > beforeType.length) {
            candidate = afterType[afterType.length - 1];
        }
        if (!candidate) {
            candidate = _gridMaker_waitForManagedEffect(clip, type, 4, 20, debugLines);
        }
        _gridMaker_debugPush(debugLines, "Added " + type + " via '" + effectName + "' candidate=" + _gridMaker_componentLabel(candidate));

        if (candidate) {
            return candidate;
        }

        // Effect insertion can be async/laggy in some Premiere builds.
        // Stop stacking duplicate adds and wait for component refresh below.
        _gridMaker_debugPush(debugLines, type + " add acknowledged without immediate candidate; waiting before additional insertions");
        break;
    }

    var fallback = _gridMaker_waitForManagedEffect(clip, type, 5, 20, debugLines);
    if (fallback) {
        _gridMaker_debugPush(debugLines, type + " managed fallback found: " + _gridMaker_componentLabel(fallback));
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
            ["crop", "recadr", "recortar", "ritagli", "freistell"],
            ["adbe crop", "adbe aecrop", "ae.adbe crop", "ae.adbe aecrop"]
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
        return _gridMaker_containsAny(matchName, ["adbe crop", "adbe aecrop"]) || _gridMaker_containsAny(name, ["crop", "recadr", "recortar", "ritagli"]);
    }
    return false;
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

    var widthProp = _gridMaker_findProperty(
        transformComp,
        [
            "scale width",
            "echelle largeur",
            "échelle largeur",
            "largeur",
            "width",
            "adbe geometry-0002",
            "adbe geometry2-0002"
        ],
        "number"
    );
    var heightProp = _gridMaker_findProperty(
        transformComp,
        [
            "scale height",
            "echelle hauteur",
            "échelle hauteur",
            "hauteur",
            "height",
            "adbe geometry-0001",
            "adbe geometry2-0001"
        ],
        "number"
    );

    if (_gridMaker_isTruthyToggleValue(_gridMaker_readPropertyValue(uniform))) {
        _gridMaker_debugPush(debugLines, "Transform uniform scale already checked; validating linkage");
        if (_gridMaker_transformUniformLinkIsEffective(uniform, widthProp, heightProp, debugLines)) {
            return;
        }
        _gridMaker_debugPush(debugLines, "Transform uniform linkage not effective despite checked state");
    }

    var ok = _gridMaker_trySetToggleProperty(uniform, 1, debugLines, "transform.uniformScale=1");
    if (!ok) {
        ok = _gridMaker_trySetToggleProperty(uniform, true, debugLines, "transform.uniformScale=true");
    }
    if (ok) {
        // Force a no-op sync pass so both scale axes stay aligned with default-like behavior.
        _gridMaker_transformSyncScaleAxes(widthProp, heightProp, debugLines);
        var linked = _gridMaker_transformUniformLinkIsEffective(uniform, widthProp, heightProp, debugLines);
        _gridMaker_debugPush(debugLines, "Transform uniform scale enforced=" + linked + " readback=" + _gridMaker_readPropertyValue(uniform));
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
        ["crop", "recadr", "recortar", "recorte", "ritagli", "freistell"],
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
        return "{}";
    }
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
    if (!(rawBlocks instanceof Array)) {
        return out;
    }
    for (var i = 0; i < rawBlocks.length; i++) {
        var b = rawBlocks[i];
        if (!b) {
            continue;
        }
        var x = parseInt(b.x, 10);
        var y = parseInt(b.y, 10);
        var w = parseInt(b.w, 10);
        var h = parseInt(b.h, 10);
        var id = _gridMaker_designerSanitizeId(b.id || ("cell_" + i));
        if (isNaN(x) || isNaN(y) || isNaN(w) || isNaN(h)) {
            continue;
        }
        if (w < 1 || h < 1) {
            continue;
        }
        if (x < 0 || y < 0) {
            continue;
        }
        if (x + w > 10 || y + h > 10) {
            continue;
        }
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
                blocks: cfg.blocks || [],
                updatedAt: cfg.updatedAt || ""
            });
        }
        list.sort(function (a, b) {
            var ta = String(a.updatedAt || "");
            var tb = String(b.updatedAt || "");
            if (ta > tb) {
                return -1;
            }
            if (ta < tb) {
                return 1;
            }
            return 0;
        });
        return _gridMaker_jsonStringify({ ok: true, configs: list });
    } catch (e) {
        return _gridMaker_jsonStringify({ ok: false, message: String(e), configs: [] });
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

        var blocks = _gridMaker_designerNormalizeBlocks(payload.blocks);
        if (blocks.length < 1) {
            return _gridMaker_jsonStringify({ ok: false, message: "empty_blocks" });
        }

        var id = _gridMaker_designerSanitizeId(payload.id);
        var ratioKey = _gridMaker_designerRatioKey(ratioW, ratioH);
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
