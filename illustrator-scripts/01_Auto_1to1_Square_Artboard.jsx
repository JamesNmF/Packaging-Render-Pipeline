#target illustrator

/**
 * 包装渲染专用工具 01：一键自适应 1:1 正方形画板
 * 功能：
 * 1. 框选图稿时：自动抓取最大边长，瞬间生成居中 1:1 正方形画板（不拉伸、不改变原有物理比例）。
 * 2. 未选图稿时：直接将当前长方形画板以中心为基准扩展为 1:1 正方形。
 */
(function () {
    if (app.documents.length === 0) {
        alert("请先在 Illustrator 中打开包装平面文件！", "提示");
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    function ptToMm(pt) {
        return (pt * 0.352778).toFixed(1);
    }

    // 模式 A：框选了图稿 -> 自动根据选区生成居中 1:1 正方形画板
    if (selection && selection.length > 0) {
        var bounds = selection[0].visibleBounds;
        for (var i = 1; i < selection.length; i++) {
            var b = selection[i].visibleBounds;
            bounds[0] = Math.min(bounds[0], b[0]); // Left
            bounds[1] = Math.max(bounds[1], b[1]); // Top
            bounds[2] = Math.max(bounds[2], b[2]); // Right
            bounds[3] = Math.min(bounds[3], b[3]); // Bottom
        }

        var w = bounds[2] - bounds[0];
        var h = bounds[1] - bounds[3];
        var maxSide = Math.max(w, h);

        var centerX = bounds[0] + w / 2;
        var centerY = bounds[1] - h / 2;

        var rect = [
            centerX - maxSide / 2,
            centerY + maxSide / 2,
            centerX + maxSide / 2,
            centerY - maxSide / 2
        ];

        var newAb = doc.artboards.add(rect);
        var abIdx = doc.artboards.length;
        newAb.name = "Texture_1to1_SKU_" + abIdx;
        doc.artboards.setActiveArtboardIndex(abIdx - 1);
        app.redraw();
        return;
    }

    // 模式 B：没有选中图稿 -> 直接将当前长方形画板扩展为 1:1 正方形
    var activeIdx = doc.artboards.getActiveArtboardIndex();
    var curAb = doc.artboards[activeIdx];
    var rect = curAb.artboardRect;

    var curW = rect[2] - rect[0];
    var curH = rect[1] - rect[3];

    if (Math.abs(curW - curH) < 0.01) {
        alert("当前画板已经是 1:1 正方形 (" + ptToMm(curW) + " mm)！", "提示");
        return;
    }

    var maxLen = Math.max(curW, curH);
    var midX = rect[0] + curW / 2;
    var midY = rect[1] - curH / 2;

    curAb.artboardRect = [
        midX - maxLen / 2,
        midY + maxLen / 2,
        midX + maxLen / 2,
        midY - maxLen / 2
    ];

    app.redraw();
})();