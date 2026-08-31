#target illustrator
/**
 * 🎨 美术资产中枢 (Art Asset Hub) · Adobe Illustrator 专属贴图导出脚本
 * ----------------------------------------------------------------------
 * 【核心功能】：
 * 1. 智能溯源：自动获取当前打开的 .ai 文件路径；
 * 2. 自动定位：自动回溯到项目根目录，精准锁定同级的「02_Textures_贴图资产」目录（不存在则自动创建）；
 * 3. 高清直出：将当前文档所有画板以 300 DPI / PNG-24 高清输出，文件名格式：[画板名称].png；
 * 4. 免弹窗秒导：执行完毕后弹出完成提示，支持一键打开贴图文件夹。
 */

function exportArtboardsToTextures() {
    if (app.documents.length === 0) {
        alert("⚠️ 当前没有打开任何 Illustrator 文档！\n请先打开需要导出贴图的 AI 文件。");
        return;
    }

    var doc = app.activeDocument;
    var aiFile;
    try {
        aiFile = doc.fullName;
    } catch (e) {
        alert("⚠️ 当前文档尚未保存到硬盘！\n请先保存该 AI 文件后再执行贴图导出。");
        return;
    }

    var currentFolder = aiFile.parent; // 如 E:\zjc\柏缇\水乳\01_Design_平面原稿
    var projectFolder = currentFolder.parent; // 如 E:\zjc\柏缇\水乳
    
    // 寻找贴图目录（优先 02_Textures_贴图资产，其次 02_Textures，再其次 Textures）
    var textureCandidates = [
        "02_Textures_贴图资产",
        "02_Textures",
        "02_贴图资产",
        "Textures",
        "贴图"
    ];

    var targetTextureFolder = null;
    for (var i = 0; i < textureCandidates.length; i++) {
        var testFolder = new Folder(projectFolder.fsName + "/" + textureCandidates[i]);
        if (testFolder.exists) {
            targetTextureFolder = testFolder;
            break;
        }
    }

    // 如果未找到，默认在项目根目录下自动创建「02_Textures_贴图资产」
    if (!targetTextureFolder) {
        targetTextureFolder = new Folder(projectFolder.fsName + "/02_Textures_贴图资产");
        if (!targetTextureFolder.exists) {
            targetTextureFolder.create();
        }
    }

    var artboardCount = doc.artboards.length;
    var baseDocName = doc.name.replace(/\.[^\.]+$/, "");

    // 导图参数配置：PNG-24, 300 DPI 高清, 抗锯齿
    var exportOptions = new ExportOptionsPNG24();
    exportOptions.antiAliasing = true;
    exportOptions.transparency = true;
    exportOptions.artBoardClipping = true;
    exportOptions.horizontalScale = 416.666; // 72 DPI -> 300 DPI (300 / 72 * 100)
    exportOptions.verticalScale = 416.666;

    var exportedFiles = [];

    // 逐画板导出
    for (var a = 0; a < artboardCount; a++) {
        doc.artboards.setActiveArtboardIndex(a);
        var ab = doc.artboards[a];
        var abName = ab.name ? ab.name.replace(/[:\\/*?\"<>|]/g, "_") : ("画板_" + (a + 1));
        
        // 规整文件名：如果画板名包含通用名则加前缀，否则直接用画板名
        var outFileName = baseDocName + "_" + abName + ".png";
        if (abName !== "画板 1" && abName !== "Artboard 1" && abName !== baseDocName) {
            outFileName = abName + ".png";
        }

        var destFile = new File(targetTextureFolder.fsName + "/" + outFileName);
        doc.exportFile(destFile, ExportType.PNG24, exportOptions);
        exportedFiles.push(outFileName);
    }

    var successMsg = "🎉 【美术资产中枢 · 贴图直出成功】\n" +
                     "------------------------------------\n" +
                     "📁 导出目标：" + targetTextureFolder.fsName + "\n" +
                     "🖼️ 导出画板数量：" + artboardCount + " 张 (300 DPI 高保真 PNG)\n\n" +
                     "是否立即打开贴图文件夹检视？";

    if (confirm(successMsg)) {
        targetTextureFolder.execute();
    }
}

// 启动执行
exportArtboardsToTextures();
