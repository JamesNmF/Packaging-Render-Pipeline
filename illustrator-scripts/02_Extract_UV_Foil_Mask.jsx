#target illustrator

/**
 * 包装渲染专用工具 02：一键提取工艺为黑白遮罩 (Mask) - 修复版
 * 修复说明：
 * 1. 严格使用新画板的实时边界 (artboardRect) 锁定黑底坐标，彻底解决黑底偏移问题；
 * 2. 将黑底自动置于底层 (Send to Back)，确保白色工艺图层在上；
 * 3. 使用 translate(deltaX, deltaY) 平移复制图元，完美保留图元群组与复合路径相对位置。
 */
(function () {
    if (app.documents.length === 0) {
        alert("请先打开包装设计文件！", "提示");
        return;
    }

    var doc = app.activeDocument;
    var selection = doc.selection;

    if (!selection || selection.length === 0) {
        alert("请先选中需要制作工艺遮罩的图元（如 UV 文字或烫金 Logo）！\n提示：可用【选择 -> 相同 -> 填充颜色】一秒全选。", "提示");
        return;
    }

    var activeIdx = doc.artboards.getActiveArtboardIndex();
    var curAb = doc.artboards[activeIdx];
    var abRect = curAb.artboardRect;
    
    var curLeft = abRect[0];
    var curTop = abRect[1];
    var curRight = abRect[2];
    var curBottom = abRect[3];

    var abW = Math.abs(curRight - curLeft);
    var abH = Math.abs(curTop - curBottom);

    // 新画板在右侧生成，水平偏移量
    var offsetX = abW + 50;

    var newRect = [
        curLeft + offsetX,
        curTop,
        curRight + offsetX,
        curBottom
    ];

    // 1. 创建新画板
    var newAb = doc.artboards.add(newRect);
    newAb.name = curAb.name + "_Mask_UV";
    var newIdx = doc.artboards.length - 1;
    doc.artboards.setActiveArtboardIndex(newIdx);

    // 2. 在新画板位置创建 100% 贴合的纯黑背景
    var blackColor = new RGBColor();
    blackColor.red = 0;
    blackColor.green = 0;
    blackColor.blue = 0;

    var whiteColor = new RGBColor();
    whiteColor.red = 255;
    whiteColor.green = 255;
    whiteColor.blue = 255;

    // 获取新画板的精确坐标
    var newAbRect = newAb.artboardRect;
    var bgTop = newAbRect[1];
    var bgLeft = newAbRect[0];
    var bgW = Math.abs(newAbRect[2] - newAbRect[0]);
    var bgH = Math.abs(newAbRect[1] - newAbRect[3]);

    // 创建矩形：PathItems.rectangle(top, left, width, height)
    var bgRect = doc.pathItems.rectangle(bgTop, bgLeft, bgW, bgH);
    bgRect.filled = true;
    bgRect.fillColor = blackColor;
    bgRect.stroked = false;
    // 将黑底置于当前图层最底层
    bgRect.zOrder(ZOrderMethod.SENDTOBACK);

    // 3. 复制选中的工艺图元并平移到新画板
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i].duplicate();
        // 采用 translate 平移，防止群组/复合路径内部坐标错乱
        item.translate(offsetX, 0);
        // 置于顶层
        item.zOrder(ZOrderMethod.BRINGTOFRONT);
        // 递归变纯白
        recolorToWhite(item, whiteColor);
    }

    function recolorToWhite(obj, color) {
        try {
            if (obj.typename === "PathItem") {
                if (obj.filled) obj.fillColor = color;
                if (obj.stroked) obj.strokeColor = color;
            } else if (obj.typename === "CompoundPathItem") {
                for (var j = 0; j < obj.pathItems.length; j++) {
                    recolorToWhite(obj.pathItems[j], color);
                }
            } else if (obj.typename === "GroupItem") {
                for (var k = 0; k < obj.pageItems.length; k++) {
                    recolorToWhite(obj.pageItems[k], color);
                }
            } else if (obj.typename === "TextFrame") {
                obj.textRange.characterAttributes.fillColor = color;
            }
        } catch (e) {
            // 忽略非着色图元错误
        }
    }

    app.redraw();
})();