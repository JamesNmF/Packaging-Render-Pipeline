#target illustrator

/**
 * 包装渲染专用工具 03：一键批量导出所有画板为 300DPI 透明 PNG 贴图
 * 功能：
 * 1. 自动以画板名称作为文件名；
 * 2. 自动导出为 300 PPI 高清透明 PNG（带 Alpha 通道）；
 * 3. 自动保存在 AI 文件同级目录下的【Textures】文件夹中。
 */
(function () {
    if (app.documents.length === 0) {
        alert("请先打开包装文件！", "提示");
        return;
    }

    var doc = app.activeDocument;
    var docPath;

    try {
        docPath = doc.path;
    } catch (e) {
        alert("请先保存当前 AI 文件，以便确定贴图导出目录！", "提示");
        return;
    }

    var targetFolder = new Folder(docPath + "/Textures");
    if (!targetFolder.exists) {
        targetFolder.create();
    }

    var options = new ExportOptionsPNG24();
    options.antiAliasing = true;
    options.transparency = true;
    options.horizontalScale = 416.66; // 72 to 300 PPI (约 4.16x)
    options.verticalScale = 416.66;
    options.artBoardClipping = true;

    for (var i = 0; i < doc.artboards.length; i++) {
        doc.artboards.setActiveArtboardIndex(i);
        var ab = doc.artboards[i];
        var file = new File(targetFolder.fsName + "/" + ab.name + ".png");
        
        doc.exportFile(file, ExportType.PNG24, options);
    }

    alert("批量导出完成！\n所有贴图已保存在：\n" + targetFolder.fsName, "导出成功");
})();