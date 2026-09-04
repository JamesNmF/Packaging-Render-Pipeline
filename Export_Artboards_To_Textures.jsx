#target illustrator
/**
 * 🎨 美术资产中枢 (Art Asset Hub) · Adobe Illustrator 专属贴图导出脚本 (v1.0 正式版)
 * ----------------------------------------------------------------------
 * 【核心功能】：
 * 1. 智能溯源：自动获取当前打开的 .ai 文件路径；
 * 2. 自动定位：自动回溯到项目根目录，精准锁定同级的「02_Textures_贴图资产」或「png」目录（不存在则自动创建）；
 * 3. 高清直出：将当前文档所有画板以 300 DPI / PNG-24 高清输出，文件名格式：[画板名称].png；
 * 4. 查重与容错：自动识别同名画板并添加序号防止覆盖，支持单画板异常隔离；
 * 5. 免弹窗秒导：执行完毕后弹出完成提示，支持一键打开贴图文件夹。
 */

function sanitizeFileName(name) {
    if (!name) return "";
    return name.replace(/[:\\/*?\"<>|]/g, "_").replace(/^\.+|\.+$/g, "").replace(/\s+$/g, "");
}

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
    var projectFolder = currentFolder.parent ? currentFolder.parent : currentFolder; // 安全兜底：防止根目录导致 null
    
    // 寻找贴图目录（优先 02_Textures_贴图资产，其次 02_Textures，再其次老工程 png，再其次 Textures）
    var textureCandidates = [
        "02_Textures_贴图资产",
        "02_Textures",
        "png",
        "PNG",
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
    var baseDocName = sanitizeFileName(doc.name.replace(/\.[^\.]+$/, ""));

    // 导图参数配置：PNG-24, 300 DPI 高清, 抗锯齿
    var exportOptions = new ExportOptionsPNG24();
    exportOptions.antiAliasing = true;
    exportOptions.transparency = true;
    exportOptions.artBoardClipping = true;
    exportOptions.horizontalScale = 416.666; // 72 DPI -> 300 DPI (300 / 72 * 100)
    exportOptions.verticalScale = 416.666;

    var usedNames = {};
    var successCount = 0;
    var failedList = [];

    // 逐画板导出
    for (var a = 0; a < artboardCount; a++) {
        doc.artboards.setActiveArtboardIndex(a);
        var ab = doc.artboards[a];
        var rawAbName = ab.name ? sanitizeFileName(ab.name) : ("画板_" + (a + 1));
        
        // 规整基础文件名
        var baseName = (rawAbName !== "画板 1" && rawAbName !== "Artboard 1" && rawAbName !== baseDocName) ? rawAbName : (baseDocName + "_" + rawAbName);
        
        // 同名画板查重与递增编号，防止先导出的被后导出的静默覆盖
        var finalName = baseName;
        if (usedNames[baseName]) {
            usedNames[baseName]++;
            finalName = baseName + "_" + usedNames[baseName];
        } else {
            usedNames[baseName] = 1;
        }

        var destFile = new File(targetTextureFolder.fsName + "/" + finalName + ".png");
        try {
            doc.exportFile(destFile, ExportType.PNG24, exportOptions);
            successCount++;
        } catch (err) {
            var errMsg = err && err.message ? err.message : String(err);
            failedList.push(finalName + " (" + errMsg + ")");
        }
    }

    var successMsg = "🎉 【美术资产中枢 · 贴图直出完成】\n" +
                     "------------------------------------\n" +
                     "📁 导出目标：" + targetTextureFolder.fsName + "\n" +
                     "🖼️ 成功导出画板：" + successCount + " / " + artboardCount + " 张 (300 DPI 高保真 PNG)\n";

    if (failedList.length > 0) {
        successMsg += "\n⚠️ 导出失败画板 (" + failedList.length + " 个):\n - " + failedList.join("\n - ") + "\n\n";
    } else {
        successMsg += "\n";
    }
    successMsg += "是否立即打开贴图文件夹检视？";

    if (confirm(successMsg)) {
        targetTextureFolder.execute();
    }
}

// 启动执行
exportArtboardsToTextures();
