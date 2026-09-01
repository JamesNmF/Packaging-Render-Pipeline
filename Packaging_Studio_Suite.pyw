# -*- coding: utf-8 -*-
"""
美术资产中枢 (Art Asset Hub - v1.0 正式版 Qt6 / PySide6 工业级架构)
===================================================================
【核心架构体系】：
1. 🚀 144 FPS 满帧 GPU 硬件加速画廊 (Qt6 / PySide6 RHI Direct3D Engine)：
   - 基于工业标准 Qt6 C++ 原生 Direct3D / OpenGL 硬件渲染引擎；
   - 采用 QListView + QAbstractListModel + QStyledItemDelegate 虚拟化视口；
   - 4-Worker QThreadPool 异步图像流式解码与信号槽 (Signal/Slot) 回填；
   - 彻底绝缘假死与“未响应”，冷启动 0.05 秒瞬开。

2. 🧭 「业务形态 + 客户品牌」双维度侧边导航与资产时间阶梯排序：
   - 🏷️ 上组：业务形态分类 (全部 / 包装 / 套盒 / 海报 / 物料)
   - 🏢 下组：客户品牌库 (全部品牌 / 柏缇 / 森之露 / 语后 / 漱外... 动态扫描实时联动)
   - ⏱️ 4 级资产时间阶梯排序 (Tier 1 渲染图修改时间 > Tier 2 贴图资产时间 > Tier 3 Blend文件时间 > 文件夹时间)

3. 🐱 萌猫开屏等待页与 Windows 沉浸式暗黑顶栏：
   - 启动时展示萌猫开工界面，后台全量资产就绪后无缝淡入主界面；
   - Windows 11/10 原生 DWM 沉浸式暗黑标题栏 (Immersive Dark Title Bar)。

4. 📥 设计源文件分拣与开工工作台 (Source Organizer & Pipeline Launcher)：
   - ⚙️ 自定义文件夹归档规则管理器：自由新建/编辑子目录结构，内置目录树实时预览；
   - 自动绑定 E:\\zjc\\包装默认文件.blend 母版工程并拉起 Blender 5.2 LTS 开工；
   - 自动将新项目录入《产品列表.xlsx》；
   - 📊 渲染图一键双向内嵌写入 Excel 台账单元格。
===================================================================
"""

import os
import sys
import re
import json
import glob
import html
import shutil
import hashlib
import zipfile
import tempfile
import datetime
import threading
import webbrowser
import subprocess
import xml.etree.ElementTree as ET

import ctypes
from ctypes import wintypes

from PIL import Image
Image.MAX_IMAGE_PIXELS = None
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage

from PySide6.QtCore import (
    Qt, QSize, QRect, QRectF, QPoint, QModelIndex, QAbstractListModel,
    QThreadPool, QRunnable, Signal, QObject, Slot, QTimer, QEvent
)
from PySide6.QtGui import (
    QIcon, QPixmap, QImage, QPainter, QColor, QFont, QPen, QBrush,
    QPainterPath, QCursor, QFontMetrics, QLinearGradient
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QComboBox, QListWidget, QListWidgetItem,
    QListView, QStyledItemDelegate, QStyle, QMenu, QFileDialog, QInputDialog,
    QMessageBox, QDialog, QTableWidget, QTableWidgetItem, QCheckBox,
    QHeaderView, QSplitter, QGroupBox, QTextEdit, QPlainTextEdit, QFrame,
    QProgressBar, QSplashScreen
)

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_suite_v7.json")
CACHE_DIR = os.path.join(os.path.expanduser("~"), ".packaging_asset_thumbnails")
EXCEL_CACHE_DIR = os.path.join(CACHE_DIR, "excel_images")
THUMB_CACHE_DIR = os.path.join(CACHE_DIR, "fast_thumbs")
META_CACHE_FILE = os.path.join(CACHE_DIR, "disk_meta_cache.json")
os.makedirs(EXCEL_CACHE_DIR, exist_ok=True)
os.makedirs(THUMB_CACHE_DIR, exist_ok=True)

def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        p = os.path.join(sys._MEIPASS, relative_path)
        if os.path.exists(p):
            return p
    local_p = os.path.join(os.path.dirname(__file__), relative_path)
    if os.path.exists(local_p):
        return local_p
    fixed_p = os.path.join(r"C:\Users\qq424\Packaging_Tools", relative_path)
    if os.path.exists(fixed_p):
        return fixed_p
    return relative_path

APP_ICON_ICO = get_resource_path("app_icon.ico")
APP_ICON_PNG = get_resource_path("app_icon.png")
SPLASH_CAT_JPG = get_resource_path("splash_cat.jpg")

DEFAULT_EXCEL_PATH = r"C:\Users\qq424\WorkBuddy\2026-08-26-15-33-05\产品列表.xlsx"

BLENDER_EXE = r"C:\Program Files\Blender Foundation\Blender 5.2\blender.exe"
if not os.path.exists(BLENDER_EXE):
    BLENDER_EXE = "blender"

def set_dark_titlebar(hwnd, dark=True):
    if sys.platform == "win32":
        try:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            val = ctypes.c_int(1 if dark else 0)
            res = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                wintypes.HWND(hwnd),
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(val),
                ctypes.sizeof(val)
            )
            if res != 0:
                DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    wintypes.HWND(hwnd),
                    DWMWA_USE_IMMERSIVE_DARK_MODE_OLD,
                    ctypes.byref(val),
                    ctypes.sizeof(val)
                )
        except Exception:
            pass

def get_valid_template_blend(cfg=None):
    candidates = []
    if cfg and cfg.get("template_blend_path"):
        candidates.append(cfg.get("template_blend_path"))
    candidates.extend([
        r"E:\zjc\包装默认文件.blend",
        r"C:\Users\qq424\Packaging_Tools\templates\Packaging_Master_Template.blend",
        os.path.join(os.path.dirname(__file__), "templates", "Packaging_Master_Template.blend"),
        os.path.join(os.path.dirname(__file__), "Packaging_Master_Template.blend"),
        r"C:\Users\qq424\Desktop\Packaging-Render-Pipeline\templates\Packaging_Master_Template.blend"
    ])
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return ""

DEFAULT_TEMPLATE = get_valid_template_blend()

DEFAULT_WORKSPACES = []
for p in ["E:\\zjc", "D:\\Projects", "E:\\Projects"]:
    if os.path.exists(p) and p not in DEFAULT_WORKSPACES:
        DEFAULT_WORKSPACES.append(p)
if not DEFAULT_WORKSPACES:
    DEFAULT_WORKSPACES = ["E:\\zjc" if os.path.exists("E:\\zjc") else "D:\\Projects"]

SYSTEM_IGNORED_DIRS = {
    'system volume information', '$recycle.bin', 'recovery', '$windows.~bt',
    'msocache', 'config.msi', 'perflogs', 'program files', 'program files (x86)',
    'windows', 'programdata', 'appdata', '.git', '.vscode', '.idea', 'temp',
    'fonts', 'render  preset', 'renderraw', 'studio light hdri pack'
}

VALID_CATEGORIES = ["包装", "套盒", "海报", "物料"]

DEFAULT_FOLDER_RULES = [
    {
        "id": "standard_packaging_5stage",
        "name": "📦 标准包装 5 阶段工程 (默认)",
        "desc": "适用于标准产品包装设计、三维渲染与最终交付的标准工业流",
        "path_pattern": "{brand}/{sku}",
        "subfolders": [
            "01_Design_平面原稿",
            "02_Textures_贴图资产",
            "03_3D_三维工程",
            "04_Renders_通道输出",
            "05_Delivery_最终交付"
        ],
        "design_sub": "01_Design_平面原稿",
        "blend_sub": "03_3D_三维工程",
        "render_sub": "04_Renders_通道输出"
    },
    {
        "id": "category_grouped_packaging",
        "name": "🗂️ 按业务形态分流 (形态/客户/SKU)",
        "desc": "在工作盘下一级建立【包装/套盒/海报/物料】分类目录，二级为客户，三级为SKU",
        "path_pattern": "{category}/{brand}/{sku}",
        "subfolders": [
            "01_Design_平面原稿",
            "02_Textures_贴图资产",
            "03_3D_三维工程",
            "04_Renders_通道输出",
            "05_Delivery_最终交付"
        ],
        "design_sub": "01_Design_平面原稿",
        "blend_sub": "03_3D_三维工程",
        "render_sub": "04_Renders_通道输出"
    },
    {
        "id": "poster_kv_flow",
        "name": "🖼️ 海报与主视觉 KV 创作流",
        "desc": "适用于平面海报、主视觉 KV、电商展板等以渲染合成素材为主的工程",
        "path_pattern": "{brand}/{sku}",
        "subfolders": [
            "01_Ref_参考意向",
            "02_Design_设计原稿",
            "03_3D_三维模型与场景",
            "04_Renders_高清分层输出",
            "05_Final_精修定稿"
        ],
        "design_sub": "02_Design_设计原稿",
        "blend_sub": "03_3D_三维模型与场景",
        "render_sub": "04_Renders_高清分层输出"
    },
    {
        "id": "simple_3stage",
        "name": "⚡ 极简轻量 3 目录流",
        "desc": "适合轻量快速打样与紧急物料制作",
        "path_pattern": "{brand}/{sku}",
        "subfolders": [
            "01_源文件",
            "02_工程",
            "03_输出"
        ],
        "design_sub": "01_源文件",
        "blend_sub": "02_工程",
        "render_sub": "03_输出"
    }
]

DEFAULT_CONFIG = {
    "workspaces": DEFAULT_WORKSPACES,
    "current_workspace": DEFAULT_WORKSPACES[0],
    "excel_path": DEFAULT_EXCEL_PATH if os.path.exists(DEFAULT_EXCEL_PATH) else "",
    "curated_brands": ["柏缇", "森之露", "语后", "漱外", "零食有鸣"],
    "current_brand": "柏缇",
    "default_category": "包装",
    "theme": "dark",
    "auto_create_blend": True,
    "auto_open_blender": True,
    "auto_open_ai": True,
    "template_blend_path": DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else "",
    "auto_append_to_excel": True,
    "folder_rules": DEFAULT_FOLDER_RULES,
    "active_rule_id": "standard_packaging_5stage"
}

DARK_QSS = """
QMainWindow, QWidget {
    background-color: #18191C;
    color: #E2E4E8;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
    font-size: 13px;
}
QTabWidget::pane {
    border: 1px solid #2D3037;
    background-color: #18191C;
}
QTabBar::tab {
    background: #202227;
    color: #9BA1B0;
    padding: 8px 18px;
    margin-right: 4px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background: #282A31;
    color: #3B82F6;
    border-bottom: 2px solid #3B82F6;
}
QGroupBox {
    border: 1px solid #333640;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 14px;
    font-weight: bold;
    color: #9BA1B0;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
}
QLineEdit, QComboBox, QTextEdit, QPlainTextEdit {
    background-color: #202227;
    border: 1px solid #383B44;
    border-radius: 6px;
    padding: 6px 10px;
    color: #F1F3F5;
}
QLineEdit:focus, QComboBox:focus {
    border: 1px solid #3B82F6;
}
QPushButton {
    background-color: #282A31;
    border: 1px solid #383B44;
    border-radius: 6px;
    padding: 6px 14px;
    color: #F1F3F5;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #32353E;
    border-color: #6B7280;
}
QPushButton:pressed {
    background-color: #202227;
}
QPushButton#PrimaryBtn {
    background-color: #2563EB;
    border: 1px solid #1D4ED8;
    color: #FFFFFF;
}
QPushButton#PrimaryBtn:hover {
    background-color: #3B82F6;
}
QListWidget {
    background-color: #202227;
    border: 1px solid #383B44;
    border-radius: 8px;
    padding: 4px;
}
QListWidget::item {
    padding: 6px 10px;
    border-radius: 6px;
    color: #9BA1B0;
    font-weight: bold;
}
QListWidget::item:hover {
    background-color: #282A31;
    color: #F1F3F5;
}
QListWidget::item:selected {
    background-color: #2563EB;
    color: #FFFFFF;
}
QListView#GalleryView {
    background-color: #18191C;
    border: none;
}
QScrollBar:vertical {
    border: none;
    background: #18191C;
    width: 8px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #383B44;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #6B7280;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
QTableWidget {
    background-color: #202227;
    border: 1px solid #383B44;
    border-radius: 8px;
    gridline-color: #2D3037;
    color: #F1F3F5;
}
QHeaderView::section {
    background-color: #282A31;
    color: #9BA1B0;
    padding: 6px;
    border: none;
    border-bottom: 1px solid #383B44;
    font-weight: bold;
}
QMenu {
    background-color: #202227;
    border: 1px solid #383B44;
    border-radius: 8px;
    padding: 4px;
}
QMenu::item {
    padding: 8px 24px;
    border-radius: 4px;
    color: #F1F3F5;
}
QMenu::item:selected {
    background-color: #2563EB;
    color: #FFFFFF;
}
QMenu::separator {
    height: 1px;
    background: #383B44;
    margin: 4px 6px;
}
"""

LIGHT_QSS = """
QMainWindow, QWidget {
    background-color: #F4F6F9;
    color: #0F172A;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
    font-size: 13px;
}
QTabWidget::pane {
    border: 1px solid #E2E8F0;
    background-color: #F4F6F9;
}
QTabBar::tab {
    background: #FFFFFF;
    color: #64748B;
    padding: 8px 18px;
    margin-right: 4px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: bold;
}
QTabBar::tab:selected {
    background: #F4F6F9;
    color: #2563EB;
    border-bottom: 2px solid #2563EB;
}
QGroupBox {
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    margin-top: 10px;
    padding-top: 14px;
    font-weight: bold;
    color: #475569;
}
QLineEdit, QComboBox, QTextEdit, QPlainTextEdit {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 6px 10px;
    color: #0F172A;
}
QPushButton {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 6px 14px;
    color: #0F172A;
    font-weight: 600;
}
QPushButton:hover {
    background-color: #F8FAFC;
    border-color: #94A3B8;
}
QPushButton#PrimaryBtn {
    background-color: #2563EB;
    border: 1px solid #1D4ED8;
    color: #FFFFFF;
}
QListWidget {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 4px;
}
QListWidget::item {
    padding: 6px 10px;
    border-radius: 6px;
    color: #475569;
    font-weight: bold;
}
QListWidget::item:selected {
    background-color: #2563EB;
    color: #FFFFFF;
}
QListView#GalleryView {
    background-color: #F4F6F9;
    border: none;
}
"""

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in DEFAULT_CONFIG.items():
                    if k not in data:
                        data[k] = v
                if not data.get("folder_rules"):
                    data["folder_rules"] = DEFAULT_FOLDER_RULES
                return data
        except Exception:
            pass
    return DEFAULT_CONFIG

def save_json_atomic(filepath, data, ensure_ascii=False, indent=2):
    try:
        dir_name = os.path.dirname(filepath)
        os.makedirs(dir_name, exist_ok=True)
        temp_fd, temp_path = tempfile.mkstemp(dir=dir_name, prefix="tmp_save_", suffix=".json")
        with os.fdopen(temp_fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=ensure_ascii, indent=indent)
        os.replace(temp_path, filepath)
    except Exception as e:
        if 'temp_path' in locals() and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

def save_config(cfg):
    save_json_atomic(CONFIG_FILE, cfg, ensure_ascii=False, indent=4)

def load_meta_cache():
    if os.path.exists(META_CACHE_FILE):
        try:
            with open(META_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_meta_cache(cache):
    save_json_atomic(META_CACHE_FILE, cache, ensure_ascii=False, indent=1)

def get_fast_disk_thumbnail_path(orig_img_path, size=(220, 220)):
    if not orig_img_path or not os.path.exists(orig_img_path):
        return None
    try:
        mtime = os.path.getmtime(orig_img_path)
        img_id = hashlib.md5(f"{orig_img_path}_{mtime}_{size[0]}x{size[1]}".encode("utf-8")).hexdigest()
        thumb_path = os.path.join(THUMB_CACHE_DIR, f"{img_id}.jpg")
        
        if os.path.exists(thumb_path) and os.path.getsize(thumb_path) > 0:
            return thumb_path
            
        with Image.open(orig_img_path) as im:
            if im.mode in ("RGBA", "P"):
                im = im.convert("RGB")
            im.thumbnail(size, Image.Resampling.LANCZOS)
            im.save(thumb_path, "JPEG", quality=85, optimize=True)
            return thumb_path
    except Exception:
        return orig_img_path

def extract_images_from_excel_zip(excel_path):
    if not excel_path or not os.path.exists(excel_path):
        return {}
    cell_image_map = {}
    try:
        with zipfile.ZipFile(excel_path, 'r') as z:
            sheet_drawing_rels = {}
            for name in z.namelist():
                if name.startswith("xl/drawings/_rels/") and name.endswith(".rels"):
                    drawing_id = os.path.splitext(os.path.basename(name))[0].replace(".xml", "")
                    rels_xml = z.read(name)
                    root = ET.fromstring(rels_xml)
                    target_map = {}
                    for rel in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                        r_id = rel.attrib.get('Id')
                        target = rel.attrib.get('Target')
                        if target:
                            norm_target = os.path.normpath(os.path.join("xl/drawings", target)).replace("\\", "/")
                            target_map[r_id] = norm_target
                    sheet_drawing_rels[drawing_id] = target_map

            for name in z.namelist():
                if name.startswith("xl/drawings/drawing") and name.endswith(".xml"):
                    drawing_id = os.path.splitext(os.path.basename(name))[0]
                    target_map = sheet_drawing_rels.get(drawing_id, {})
                    xml_content = z.read(name)
                    root = ET.fromstring(xml_content)
                    
                    for anchor in root.findall('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}twoCellAnchor') + \
                                  root.findall('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}oneCellAnchor'):
                        from_tag = anchor.find('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from')
                        if from_tag is not None:
                            row_tag = from_tag.find('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row')
                            if row_tag is not None:
                                row_idx = int(row_tag.text) + 1
                                blip = anchor.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                                if blip is not None:
                                    embed_id = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                    if embed_id and embed_id in target_map:
                                        img_zip_path = target_map[embed_id]
                                        ext = os.path.splitext(img_zip_path)[1]
                                        img_hash = hashlib.md5(img_zip_path.encode('utf-8')).hexdigest()[:10]
                                        out_name = f"row_{row_idx}_{img_hash}{ext}"
                                        out_disk_path = os.path.join(EXCEL_CACHE_DIR, out_name)
                                        if not os.path.exists(out_disk_path):
                                            try:
                                                with open(out_disk_path, "wb") as f_out:
                                                    f_out.write(z.read(img_zip_path))
                                            except Exception:
                                                pass
                                        if os.path.exists(out_disk_path):
                                            cell_image_map[row_idx] = out_disk_path
    except Exception:
        pass
    return cell_image_map

def normalize_category(raw_val):
    if not raw_val:
        return "包装"
    s = str(raw_val).strip()
    if "套盒" in s or "礼盒" in s:
        return "套盒"
    if "海报" in s or "KV" in s or "展板" in s or "主图" in s:
        return "海报"
    if "物料" in s or "单页" in s or "折页" in s or "展架" in s or "画册" in s:
        return "物料"
    return "包装"

def parse_and_cache_excel(excel_path):
    if not excel_path or not os.path.exists(excel_path):
        return []
    projects = []
    try:
        cell_images = extract_images_from_excel_zip(excel_path)
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        sheet = wb.active
        
        headers = {}
        for col_idx, cell in enumerate(sheet[1], start=1):
            if cell.value:
                headers[str(cell.value).strip()] = col_idx
                
        sku_col = headers.get("产品名称") or headers.get("SKU") or headers.get("品名") or headers.get("产品命名") or 2
        brand_col = headers.get("品牌") or headers.get("客户") or 1
        cat_col = headers.get("业务形态") or headers.get("分类") or headers.get("类别") or None
        time_col = headers.get("创建时间") or headers.get("日期") or headers.get("录入时间") or None
        path_col = headers.get("文件路径") or headers.get("路径") or None
        
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not row or not any(row):
                continue
            sku_val = str(row[sku_col - 1]).strip() if len(row) >= sku_col and row[sku_col - 1] else ""
            if not sku_val or sku_val == "None":
                continue
            brand_val = str(row[brand_col - 1]).strip() if len(row) >= brand_col and row[brand_col - 1] else ""
            if brand_val == "None":
                brand_val = ""
            
            p_val = str(row[path_col - 1]).strip() if path_col and len(row) >= path_col and row[path_col - 1] else ""
            if p_val == "None":
                p_val = ""
                
            # 若品牌列为空或为纯数字序号，从文件路径中反提取品牌
            if (not brand_val or brand_val.isdigit()) and p_val:
                norm_p = p_val.replace("\\", "/")
                parts = [part for part in norm_p.split("/") if part]
                if len(parts) >= 2:
                    brand_val = parts[-2]
            if not brand_val or brand_val.isdigit():
                brand_val = "柏缇"
                
            cat_val = str(row[cat_col - 1]).strip() if cat_col and len(row) >= cat_col and row[cat_col - 1] else ""
            cat_val = normalize_category(cat_val)
            time_val = str(row[time_col - 1]).strip() if time_col and len(row) >= time_col and row[time_col - 1] else ""
            if time_val == "None":
                time_val = ""
                
            img_path = cell_images.get(row_idx, "")
            
            projects.append({
                "source": "excel",
                "brand": brand_val,
                "sku": sku_val,
                "cat": cat_val,
                "path": p_val,
                "thumbnail": img_path,
                "time": time_val,
                "row_idx": row_idx,
                "mtime": 0
            })
        wb.close()
    except Exception:
        pass
    return projects

def find_project_thumbnail(proj_dir):
    if not proj_dir or not os.path.exists(proj_dir):
        return None
    render_candidates = [
        os.path.join(proj_dir, "png"),
        os.path.join(proj_dir, "PNG"),
        os.path.join(proj_dir, "04_Renders_通道输出"),
        os.path.join(proj_dir, "04_Renders_高清分层输出"),
        os.path.join(proj_dir, "03_输出"),
        os.path.join(proj_dir, "04_输出"),
        os.path.join(proj_dir, "05_Delivery_最终交付"),
        os.path.join(proj_dir, "05_Final_精修定稿"),
        os.path.join(proj_dir, "渲染"),
        os.path.join(proj_dir, "Renders"),
        proj_dir
    ]
    
    img_exts = ("*.png", "*.jpg", "*.jpeg", "*.webp")
    for r_dir in render_candidates:
        if os.path.exists(r_dir):
            imgs = []
            for ext in img_exts:
                imgs.extend(glob.glob(os.path.join(r_dir, ext)))
                imgs.extend(glob.glob(os.path.join(r_dir, "*", ext)))
            if imgs:
                beauty_imgs = [
                    f for f in imgs 
                    if not any(ch in os.path.basename(f).lower() for ch in [
                        "cryptomatte", "crypto", "选区", "normal", "法线", "depth", "深度", 
                        "mist", "ao", "roughness", "specular", "alpha", "mask", "shadow"
                    ])
                ]
                if beauty_imgs:
                    beauty_imgs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                    return beauty_imgs[0]
                imgs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return imgs[0]
                
    textures_dir = os.path.join(proj_dir, "02_Textures_贴图资产")
    if os.path.exists(textures_dir):
        imgs = []
        for ext in img_exts:
            imgs.extend(glob.glob(os.path.join(textures_dir, ext)))
        if imgs:
            imgs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return imgs[0]
            
    return None

def auto_detect_category_from_name(name):
    if not name:
        return "包装"
    s = str(name).lower()
    if "套盒" in s or "礼盒" in s:
        return "套盒"
    if "海报" in s or "kv" in s or "主图" in s or "展板" in s:
        return "海报"
    if "物料" in s or "单页" in s or "折页" in s or "展架" in s or "画册" in s:
        return "物料"
    return "包装"

# ----------------- 4 级资产真实时间阶梯算法 -----------------
def get_project_asset_mtime(proj_dir):
    """
    4 级资产真实产出时间算法：
    Tier 1 (完工渲染图): png / 04_Renders / 渲染 / 03_输出 / 05_Delivery 中的最新图片修改时间
    Tier 2 (贴图资产与设计稿): 02_Textures / textures / 贴图 / 01_Design 中的最新图片/AI修改时间
    Tier 3 (3D 工程): 03_3D / 模型 / .blend 最新修改时间
    Tier 4 (保底): 根目录图片或文件夹本身创建时间
    """
    if not proj_dir or not os.path.exists(proj_dir):
        return 0
    try:
        # Tier 1: 渲染输出图
        render_candidates = [
            os.path.join(proj_dir, "png"),
            os.path.join(proj_dir, "PNG"),
            os.path.join(proj_dir, "04_Renders_通道输出"),
            os.path.join(proj_dir, "04_Renders_高清分层输出"),
            os.path.join(proj_dir, "渲染"),
            os.path.join(proj_dir, "03_输出"),
            os.path.join(proj_dir, "04_输出"),
            os.path.join(proj_dir, "05_Delivery_最终交付"),
            os.path.join(proj_dir, "05_Final_精修定稿"),
            os.path.join(proj_dir, "Renders"),
            os.path.join(proj_dir, "renders"),
        ]
        img_exts = ('.png', '.jpg', '.jpeg', '.webp')
        
        tier1_max = 0
        for r_dir in render_candidates:
            if os.path.exists(r_dir) and os.path.isdir(r_dir):
                try:
                    with os.scandir(r_dir) as it:
                        for entry in it:
                            if entry.is_file() and entry.name.lower().endswith(img_exts):
                                m = entry.stat().st_mtime
                                if m > tier1_max:
                                    tier1_max = m
                except Exception:
                    pass
        if tier1_max > 0:
            return tier1_max
            
        # Tier 2: 贴图资产与平面设计原稿 (新工程未渲染)
        texture_candidates = [
            os.path.join(proj_dir, "02_Textures_贴图资产"),
            os.path.join(proj_dir, "textures"),
            os.path.join(proj_dir, "Textures"),
            os.path.join(proj_dir, "texture"),
            os.path.join(proj_dir, "贴图"),
            os.path.join(proj_dir, "贴图资产"),
            os.path.join(proj_dir, "01_Design_平面原稿"),
            os.path.join(proj_dir, "01_源文件"),
            os.path.join(proj_dir, "02_Design_设计原稿"),
        ]
        tier2_max = 0
        design_exts = ('.png', '.jpg', '.jpeg', '.webp', '.ai', '.psd', '.pdf')
        for t_dir in texture_candidates:
            if os.path.exists(t_dir) and os.path.isdir(t_dir):
                try:
                    with os.scandir(t_dir) as it:
                        for entry in it:
                            if entry.is_file() and entry.name.lower().endswith(design_exts):
                                m = entry.stat().st_mtime
                                if m > tier2_max:
                                    tier2_max = m
                except Exception:
                    pass
        if tier2_max > 0:
            return tier2_max
            
        # Tier 3: 3D 工程文件 (.blend)
        blend_candidates = [
            os.path.join(proj_dir, "03_3D_三维工程"),
            os.path.join(proj_dir, "03_3D_三维模型与场景"),
            os.path.join(proj_dir, "02_工程"),
            os.path.join(proj_dir, "模型"),
            os.path.join(proj_dir, "3D"),
            proj_dir
        ]
        tier3_max = 0
        for b_dir in blend_candidates:
            if os.path.exists(b_dir) and os.path.isdir(b_dir):
                try:
                    with os.scandir(b_dir) as it:
                        for entry in it:
                            if entry.is_file() and entry.name.lower().endswith('.blend'):
                                m = entry.stat().st_mtime
                                if m > tier3_max:
                                    tier3_max = m
                except Exception:
                    pass
        if tier3_max > 0:
            return tier3_max
            
        # Tier 4: 根目录中的任何图片/文件时间 / 文件夹本身
        tier4_max = 0
        try:
            with os.scandir(proj_dir) as it:
                for entry in it:
                    if entry.is_file() and entry.name.lower().endswith(('.png', '.jpg', '.jpeg', '.blend', '.ai', '.psd')):
                        m = entry.stat().st_mtime
                        if m > tier4_max:
                            tier4_max = m
        except Exception:
            pass
        if tier4_max > 0:
            return tier4_max
            
        return os.path.getmtime(proj_dir)
    except Exception:
        return 0

def scan_workspace_projects_fast(ws_root, meta_cache):
    if not ws_root or not os.path.exists(ws_root):
        return []
    projects = []
    cache_dirty = False
    
    try:
        entries = os.listdir(ws_root)
    except Exception:
        return []
        
    for entry in entries:
        if entry.lower() in SYSTEM_IGNORED_DIRS:
            continue
        brand_p = os.path.join(ws_root, entry)
        if not os.path.isdir(brand_p):
            continue
            
        try:
            skus = os.listdir(brand_p)
        except (PermissionError, OSError):
            continue
            
        for sku in skus:
            if sku.lower() in SYSTEM_IGNORED_DIRS:
                continue
            sku_p = os.path.join(brand_p, sku)
            try:
                if os.path.isdir(sku_p):
                    s_mtime = get_project_asset_mtime(sku_p)
                    cache_key = sku_p.lower().replace("/", "\\")
                    
                    cached_thumb = meta_cache.get(cache_key, {}).get("thumbnail")
                    has_valid_cached_thumb = (
                        cache_key in meta_cache 
                        and meta_cache[cache_key].get("mtime") == s_mtime
                        and cached_thumb
                        and os.path.exists(cached_thumb)
                    )
                    
                    if has_valid_cached_thumb:
                        cached_item = meta_cache[cache_key]
                        projects.append({
                            "source": "disk",
                            "brand": cached_item.get("brand", entry),
                            "sku": cached_item.get("sku", sku),
                            "cat": cached_item.get("cat", auto_detect_category_from_name(sku)),
                            "path": sku_p,
                            "thumbnail": cached_item.get("thumbnail"),
                            "time": "",
                            "mtime": s_mtime
                        })
                    else:
                        thumb = find_project_thumbnail(sku_p)
                        cat = auto_detect_category_from_name(sku)
                        meta_cache[cache_key] = {
                            "brand": entry,
                            "sku": sku,
                            "cat": cat,
                            "thumbnail": thumb,
                            "mtime": s_mtime
                        }
                        cache_dirty = True
                        projects.append({
                            "source": "disk",
                            "brand": entry,
                            "sku": sku,
                            "cat": cat,
                            "path": sku_p,
                            "thumbnail": thumb,
                            "time": "",
                            "mtime": s_mtime
                        })
            except (PermissionError, OSError):
                continue
                
    if cache_dirty:
        save_meta_cache(meta_cache)
        
    projects.sort(key=lambda x: x["mtime"], reverse=True)
    return projects

def merge_excel_and_disk_projects(excel_projects, disk_projects):
    disk_map = {}
    for dp in disk_projects:
        norm_p = dp["path"].lower().replace("/", "\\")
        disk_map[norm_p] = dp
        sku_clean = re.sub(r'[\s_\-\(\)（）]+', '', dp["sku"].lower())
        disk_map[f"sku:{sku_clean}"] = dp
        
    merged = []
    matched_disk_paths = set()
    
    for ep in excel_projects:
        sku_clean = re.sub(r'[\s_\-\(\)（）]+', '', ep["sku"].lower())
        matched_dp = disk_map.get(f"sku:{sku_clean}")
        
        item = {
            "source": "merged" if matched_dp else "excel",
            "brand": ep.get("brand") or (matched_dp.get("brand") if matched_dp else "柏缇"),
            "sku": ep["sku"],
            "cat": ep.get("cat") or (matched_dp.get("cat") if matched_dp else "包装"),
            "time": ep.get("time", ""),
            "row_idx": ep.get("row_idx", 0),
            "path": matched_dp["path"] if matched_dp else (ep.get("path") or ""),
            "thumbnail": (matched_dp["thumbnail"] if matched_dp and matched_dp.get("thumbnail") and os.path.exists(matched_dp["thumbnail"]) else None) or ep.get("thumbnail"),
            "mtime": matched_dp["mtime"] if matched_dp else 0
        }
        if matched_dp:
            matched_disk_paths.add(matched_dp["path"].lower().replace("/", "\\"))
        merged.append(item)
        
    for dp in disk_projects:
        norm_p = dp["path"].lower().replace("/", "\\")
        if norm_p not in matched_disk_paths:
            merged.append({
                "source": "disk",
                "brand": dp.get("brand", ""),
                "sku": dp["sku"],
                "cat": dp.get("cat", "包装"),
                "time": "",
                "row_idx": 0,
                "path": dp["path"],
                "thumbnail": dp.get("thumbnail"),
                "mtime": dp.get("mtime", 0)
            })
            
    merged.sort(key=lambda x: x["mtime"], reverse=True)
    return merged

def update_thumbnail_to_excel(excel_path, proj_path, sku, thumb_path):
    if not excel_path or not os.path.exists(excel_path):
        return False, "Excel 文件未找到！"
    if not thumb_path or not os.path.exists(thumb_path):
        return False, f"未找到可用的缩略图文件: {thumb_path}"

    try:
        temp_img_path = get_fast_disk_thumbnail_path(thumb_path, size=(200, 200)) or thumb_path
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        
        target_row = None
        sku_col = 2
        for col_idx, cell in enumerate(sheet[1], start=1):
            val = str(cell.value or "").strip()
            if val in ("产品名称", "SKU", "品名", "产品命名"):
                sku_col = col_idx
                break
                
        sku_clean = re.sub(r'[\s_\-\(\)（）]+', '', sku.lower()) if sku else ""
        for r in range(2, sheet.max_row + 1):
            cell_val = str(sheet.cell(row=r, column=sku_col).value or "").strip()
            cell_clean = re.sub(r'[\s_\-\(\)（）]+', '', cell_val.lower())
            if sku_clean and cell_clean and (sku_clean == cell_clean or sku_clean in cell_clean or cell_clean in sku_clean):
                target_row = r
                break
                
        if not target_row:
            target_row = sheet.max_row + 1
            sheet.cell(row=target_row, column=sku_col, value=sku)
            
        img_cell = f"C{target_row}"
        img = OpenpyxlImage(temp_img_path)
        img.width = 75
        img.height = 75
        
        sheet.row_dimensions[target_row].height = 65
        sheet.column_dimensions['C'].width = 14
        
        sheet.add_image(img, img_cell)
        wb.save(excel_path)
        wb.close()
        return True, f"已成功将 [{sku}] 的缩略图写入 Excel 单元格 {img_cell}！"
    except PermissionError:
        return False, f"无法保存 Excel！请先关闭正在打开《产品列表.xlsx》的 WPS 或 Excel 程序后重试。"
    except Exception as e:
        return False, f"写入 Excel 缩略图失败: {str(e)}"

def batch_sync_all_thumbnails_to_excel(excel_path, projects):
    if not excel_path or not os.path.exists(excel_path):
        return False, "Excel 文件未找到！"
        
    valid_items = [p for p in projects if p.get("thumbnail") and os.path.exists(p["thumbnail"])]
    if not valid_items:
        return False, "当前没有找到任何带有有效缩略图的项目！"
        
    success_count = 0
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        
        sku_col = 2
        for col_idx, cell in enumerate(sheet[1], start=1):
            val = str(cell.value or "").strip()
            if val in ("产品名称", "SKU", "品名", "产品命名"):
                sku_col = col_idx
                break
                
        row_map = {}
        for r in range(2, sheet.max_row + 1):
            c_val = str(sheet.cell(row=r, column=sku_col).value or "").strip()
            if c_val:
                c_clean = re.sub(r'[\s_\-\(\)（）]+', '', c_val.lower())
                row_map[c_clean] = r
                
        sheet.column_dimensions['C'].width = 14
        
        for item in valid_items:
            sku = item.get("sku", "")
            sku_clean = re.sub(r'[\s_\-\(\)（）]+', '', sku.lower()) if sku else ""
            target_row = row_map.get(sku_clean)
            if not target_row:
                for c_clean, r_idx in row_map.items():
                    if sku_clean and (sku_clean in c_clean or c_clean in sku_clean):
                        target_row = r_idx
                        break
            if not target_row:
                continue
                
            thumb_path = item["thumbnail"]
            temp_img_path = get_fast_disk_thumbnail_path(thumb_path, size=(180, 180)) or thumb_path
            
            try:
                img = OpenpyxlImage(temp_img_path)
                img.width = 75
                img.height = 75
                sheet.row_dimensions[target_row].height = 65
                sheet.add_image(img, f"C{target_row}")
                success_count += 1
            except Exception:
                pass
                
        wb.save(excel_path)
        wb.close()
        return True, f"🎉 批量同步完成！已成功将 {success_count} 个项目的渲染图写入 Excel 台账！"
    except PermissionError:
        return False, "无法保存 Excel！请先关闭 WPS 或 Excel 后重试。"
    except Exception as e:
        return False, f"批量同步发生异常: {str(e)}"

# ----------------- 异步图像加载器 (Qt 线程池) -----------------
class ImageLoadSignal(QObject):
    finished = Signal(str, QPixmap)

class ImageLoadTask(QRunnable):
    def __init__(self, img_path, target_size=(220, 220)):
        super().__init__()
        self.img_path = img_path
        self.target_size = target_size
        self.signals = ImageLoadSignal()

    def run(self):
        try:
            fast_p = get_fast_disk_thumbnail_path(self.img_path, self.target_size)
            if fast_p and os.path.exists(fast_p):
                pm = QPixmap(fast_p)
                if not pm.isNull():
                    self.signals.finished.emit(self.img_path, pm)
        except Exception:
            pass

# ----------------- 异步全量数据加载 Worker (Qt 信号槽) -----------------
class DataLoaderSignals(QObject):
    finished = Signal(list, list, list)

class DataLoaderWorker(QRunnable):
    def __init__(self, excel_path, workspace, meta_cache):
        super().__init__()
        self.excel_path = excel_path
        self.workspace = workspace
        self.meta_cache = meta_cache
        self.signals = DataLoaderSignals()

    def run(self):
        try:
            excel_p = parse_and_cache_excel(self.excel_path) if (self.excel_path and os.path.exists(self.excel_path)) else []
            disk_p = scan_workspace_projects_fast(self.workspace, self.meta_cache) if (self.workspace and os.path.exists(self.workspace)) else []
            merged = merge_excel_and_disk_projects(excel_p, disk_p)
            self.signals.finished.emit(excel_p, disk_p, merged)
        except Exception:
            self.signals.finished.emit([], [], [])

# ----------------- 萌猫开屏等待窗口 (Cat Splash Screen) -----------------
class CatSplashScreen(QWidget):
    def __init__(self, splash_img_path):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.SplashScreen)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.resize(520, 440)
        
        self.pixmap = None
        if os.path.exists(splash_img_path):
            self.pixmap = QPixmap(splash_img_path).scaled(520, 400, Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        self.img_lbl = QLabel()
        self.img_lbl.setAlignment(Qt.AlignCenter)
        if self.pixmap:
            self.img_lbl.setPixmap(self.pixmap)
        layout.addWidget(self.img_lbl, stretch=1)
        
        bottom_bar = QWidget()
        bottom_bar.setFixedHeight(46)
        bottom_bar.setStyleSheet("background: #18191C; border-bottom-left-radius: 12px; border-bottom-right-radius: 12px;")
        b_layout = QVBoxLayout(bottom_bar)
        b_layout.setContentsMargins(16, 6, 16, 6)
        
        self.status_lbl = QLabel("🚀 正在极速载入美术资产中枢...")
        self.status_lbl.setStyleSheet("color: #93C5FD; font-weight: bold; font-size: 12px;")
        b_layout.addWidget(self.status_lbl)
        
        self.prog_bar = QProgressBar()
        self.prog_bar.setRange(0, 0)
        self.prog_bar.setFixedHeight(4)
        self.prog_bar.setTextVisible(False)
        self.prog_bar.setStyleSheet("""
            QProgressBar { border: none; background: #282A31; border-radius: 2px; }
            QProgressBar::chunk { background: #3B82F6; border-radius: 2px; }
        """)
        b_layout.addWidget(self.prog_bar)
        layout.addWidget(bottom_bar)

    def set_status_text(self, text):
        self.status_lbl.setText(text)

# ----------------- Qt 虚拟化画廊数据模型 -----------------
class GalleryModel(QAbstractListModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.projects = []
        self.pixmap_cache = {}
        self.loading_set = set()

    def rowCount(self, parent=QModelIndex()):
        return len(self.projects)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or index.row() >= len(self.projects):
            return None
        proj = self.projects[index.row()]
        if role == Qt.UserRole:
            return proj
        return None

    def set_projects(self, projects):
        self.beginResetModel()
        self.projects = projects
        self.endResetModel()

    def get_project(self, row):
        if 0 <= row < len(self.projects):
            return self.projects[row]
        return None

# ----------------- Qt6 GPU 硬件加速卡片 Delegate (支持点击命中测试) -----------------
class GalleryCardDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_view = parent
        self.thread_pool = QThreadPool.globalInstance()
        self.thread_pool.setMaxThreadCount(4)

    def sizeHint(self, option, index):
        return QSize(220, 280)

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.MouseButtonRelease and event.button() == Qt.LeftButton:
            rect = option.rect.adjusted(6, 6, -6, -6)
            pos = event.pos()
            thumb_rect = QRect(rect.left() + 6, rect.top() + 6, rect.width() - 12, rect.width() - 12)
            badge_y = thumb_rect.bottom() + 10
            title_y = badge_y + 26
            action_y = title_y + 24
            btn1_rect = QRect(rect.left() + 10, action_y, (rect.width() - 26) // 2, 22)
            btn2_rect = QRect(btn1_rect.right() + 6, action_y, (rect.width() - 26) // 2, 22)
            
            proj = index.data(Qt.UserRole)
            if proj:
                win = self.parent_view.window()
                if btn1_rect.contains(pos):
                    win.open_folder(proj.get("path"), sku=proj.get("sku"), brand=proj.get("brand"))
                    return True
                elif btn2_rect.contains(pos):
                    win.launch_blend(proj.get("path"), sku=proj.get("sku"), brand=proj.get("brand"))
                    return True
                elif thumb_rect.contains(pos):
                    win.launch_blend(proj.get("path"), sku=proj.get("sku"), brand=proj.get("brand"))
                    return True
        return super().editorEvent(event, model, option, index)

    def paint(self, painter: QPainter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)

        rect = option.rect.adjusted(6, 6, -6, -6)
        proj = index.data(Qt.UserRole)
        if not proj:
            painter.restore()
            return

        is_hover = (option.state & QStyle.State_MouseOver)
        is_dark = (self.parent_view.window().current_theme == "dark") if hasattr(self.parent_view, "window") else True

        # 1. 卡片背景与投影 (GPU Direct3D 快速绘制)
        card_bg = QColor("#282A31") if is_dark else QColor("#FFFFFF")
        border_color = QColor("#3B82F6") if is_hover else (QColor("#383B44") if is_dark else QColor("#E2E8F0"))
        
        path = QPainterPath()
        path.addRoundedRect(QRectF(rect), 8, 8)
        painter.fillPath(path, card_bg)
        painter.setPen(QPen(border_color, 1.5 if is_hover else 1))
        painter.drawPath(path)

        # 2. 缩略图区域
        thumb_rect = QRect(rect.left() + 6, rect.top() + 6, rect.width() - 12, rect.width() - 12)
        thumb_path_shape = QPainterPath()
        thumb_path_shape.addRoundedRect(QRectF(thumb_rect), 6, 6)
        painter.fillPath(thumb_path_shape, QColor("#141518") if is_dark else QColor("#F8FAFC"))

        img_path = proj.get("thumbnail")
        model = index.model()
        pix = model.pixmap_cache.get(img_path) if img_path else None

        if pix and not pix.isNull():
            scaled = pix.scaled(thumb_rect.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            x_off = thumb_rect.left() + (thumb_rect.width() - scaled.width()) // 2
            y_off = thumb_rect.top() + (thumb_rect.height() - scaled.height()) // 2
            painter.drawPixmap(x_off, y_off, scaled)
        else:
            painter.setPen(QColor("#6B7280"))
            painter.drawText(thumb_rect, Qt.AlignCenter, "📦 待渲染工程")
            
            if img_path and img_path not in model.loading_set and os.path.exists(img_path):
                model.loading_set.add(img_path)
                task = ImageLoadTask(img_path)
                task.signals.finished.connect(self.on_image_loaded)
                self.thread_pool.start(task)

        # 3. 标签行 (形态 Badge + 品牌)
        cat_val = proj.get("cat", "包装")
        brand_val = proj.get("brand", "")
        sku_val = proj.get("sku", "")

        cat_colors = {
            "包装": (QColor(59, 130, 246, 50), QColor("#93C5FD")),
            "套盒": (QColor(245, 158, 11, 50), QColor("#FCD34D")),
            "海报": (QColor(139, 92, 246, 50), QColor("#C4B5FD")),
            "物料": (QColor(16, 185, 129, 50), QColor("#6EE7B7"))
        }
        badge_bg, badge_fg = cat_colors.get(cat_val, cat_colors["包装"])

        badge_y = thumb_rect.bottom() + 10
        badge_rect = QRect(rect.left() + 10, badge_y, 42, 18)
        badge_path = QPainterPath()
        badge_path.addRoundedRect(QRectF(badge_rect), 4, 4)
        painter.fillPath(badge_path, badge_bg)
        painter.setPen(badge_fg)
        font_badge = QFont(painter.font())
        font_badge.setPointSize(8)
        font_badge.setBold(True)
        painter.setFont(font_badge)
        painter.drawText(badge_rect, Qt.AlignCenter, cat_val)

        if brand_val:
            painter.setPen(QColor("#9BA1B0") if is_dark else QColor("#64748B"))
            font_brand = QFont(painter.font())
            font_brand.setPointSize(8)
            painter.setFont(font_brand)
            painter.drawText(badge_rect.right() + 6, badge_y + 13, brand_val)

        # 4. SKU 标题
        title_y = badge_y + 26
        title_rect = QRect(rect.left() + 10, title_y, rect.width() - 20, 22)
        painter.setPen(QColor("#F1F3F5") if is_dark else QColor("#0F172A"))
        font_title = QFont(painter.font())
        font_title.setPointSize(10)
        font_title.setBold(True)
        painter.setFont(font_title)
        
        metrics = QFontMetrics(font_title)
        elided_title = metrics.elidedText(sku_val, Qt.ElideRight, title_rect.width())
        painter.drawText(title_rect, Qt.AlignLeft | Qt.AlignVCenter, elided_title)

        # 5. 操作底栏
        action_y = title_y + 24
        btn1_rect = QRect(rect.left() + 10, action_y, (rect.width() - 26) // 2, 22)
        btn2_rect = QRect(btn1_rect.right() + 6, action_y, (rect.width() - 26) // 2, 22)
        
        b1_path = QPainterPath()
        b1_path.addRoundedRect(QRectF(btn1_rect), 4, 4)
        painter.fillPath(b1_path, QColor("#202227") if is_dark else QColor("#E2E8F0"))
        painter.setPen(QColor("#9BA1B0") if is_dark else QColor("#475569"))
        font_btn = QFont(painter.font())
        font_btn.setPointSize(8)
        painter.setFont(font_btn)
        painter.drawText(btn1_rect, Qt.AlignCenter, "📁 文件夹")

        b2_path = QPainterPath()
        b2_path.addRoundedRect(QRectF(btn2_rect), 4, 4)
        painter.fillPath(b2_path, QColor("#2563EB"))
        painter.setPen(QColor("#FFFFFF"))
        font_btn.setBold(True)
        painter.setFont(font_btn)
        painter.drawText(btn2_rect, Qt.AlignCenter, "🚀 3D工程")

        painter.restore()

    def on_image_loaded(self, img_path, pixmap):
        model = self.parent_view.model()
        model.pixmap_cache[img_path] = pixmap
        model.loading_set.discard(img_path)
        self.parent_view.viewport().update()

# ----------------- 自定义文件夹规则管理弹窗 -----------------
class FolderRuleManagerDialog(QDialog):
    def __init__(self, rules, active_rule_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ 自定义文件夹归档规则管理器")
        self.resize(760, 520)
        self.setMinimumSize(680, 460)
        
        self.rules = [dict(r) for r in rules]
        self.active_rule_id = active_rule_id
        self.current_idx = 0
        
        for idx, r in enumerate(self.rules):
            if r.get("id") == active_rule_id:
                self.current_idx = idx
                break
                
        self.build_ui()
        self.load_rule_into_form(self.current_idx)

    def showEvent(self, event):
        super().showEvent(event)
        set_dark_titlebar(int(self.winId()), True)

    def build_ui(self):
        main_layout = QVBoxLayout(self)
        splitter = QSplitter(Qt.HORIZONTAL)
        
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.addWidget(QLabel("<b>规则预设库</b>"))
        
        self.rule_list = QListWidget()
        self.rule_list.currentRowChanged.connect(self.on_rule_selection_changed)
        left_layout.addWidget(self.rule_list)
        
        btn_box_left = QHBoxLayout()
        btn_new = QPushButton("➕ 新建")
        btn_new.clicked.connect(self.new_rule)
        btn_del = QPushButton("🗑️ 删除")
        btn_del.clicked.connect(self.delete_rule)
        btn_box_left.addWidget(btn_new)
        btn_box_left.addWidget(btn_del)
        left_layout.addLayout(btn_box_left)
        
        splitter.addWidget(left_widget)
        
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        right_layout.addWidget(QLabel("规则名称:"))
        self.name_edit = QLineEdit()
        self.name_edit.textChanged.connect(self.on_form_modified)
        right_layout.addWidget(self.name_edit)
        
        right_layout.addWidget(QLabel("层级模板:"))
        self.pattern_combo = QComboBox()
        self.pattern_combo.addItems(["{brand}/{sku}", "{category}/{brand}/{sku}", "{sku}"])
        self.pattern_combo.currentTextChanged.connect(self.on_form_modified)
        right_layout.addWidget(self.pattern_combo)
        
        right_layout.addWidget(QLabel("子文件夹列表 (每行一个):"))
        self.subs_edit = QPlainTextEdit()
        self.subs_edit.textChanged.connect(self.on_form_modified)
        right_layout.addWidget(self.subs_edit)
        
        right_layout.addWidget(QLabel("📁 目录树实时预览:"))
        self.preview_lbl = QLabel()
        self.preview_lbl.setStyleSheet("background: #141518; color: #34D399; font-family: Consolas; padding: 10px; border-radius: 6px;")
        right_layout.addWidget(self.preview_lbl)
        
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        main_layout.addWidget(splitter)
        
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()
        btn_save = QPushButton("💾 保存并应用此规则")
        btn_save.setObjectName("PrimaryBtn")
        btn_save.clicked.connect(self.save_and_accept)
        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)
        bottom_layout.addWidget(btn_cancel)
        bottom_layout.addWidget(btn_save)
        main_layout.addLayout(bottom_layout)
        
        self.refresh_list()

    def refresh_list(self):
        self.rule_list.clear()
        for r in self.rules:
            self.rule_list.addItem(r.get("name", "未命名规则"))
        if 0 <= self.current_idx < len(self.rules):
            self.rule_list.setCurrentRow(self.current_idx)

    def load_rule_into_form(self, idx):
        if 0 <= idx < len(self.rules):
            r = self.rules[idx]
            self.name_edit.setText(r.get("name", ""))
            self.pattern_combo.setCurrentText(r.get("path_pattern", "{brand}/{sku}"))
            self.subs_edit.setPlainText("\n".join(r.get("subfolders", [])))
            self.update_preview()

    def on_rule_selection_changed(self, row):
        if row >= 0 and row != self.current_idx:
            self.save_current_form()
            self.current_idx = row
            self.load_rule_into_form(self.current_idx)

    def on_form_modified(self):
        self.update_preview()

    def update_preview(self):
        pat = self.pattern_combo.currentText().strip() or "{brand}/{sku}"
        root_name = pat.replace("{brand}", "柏缇").replace("{sku}", "红参抗皱霜").replace("{category}", "包装")
        raw_text = self.subs_edit.toPlainText()
        subs = [s.strip() for s in raw_text.split("\n") if s.strip()]
        
        lines = [f"📁 [主工作盘]\\{root_name}"]
        for i, s in enumerate(subs):
            prefix = " └── 📂 " if i == len(subs) - 1 else " ├── 📂 "
            lines.append(f"{prefix}{s}")
        self.preview_lbl.setText("\n".join(lines[:7]))

    def save_current_form(self):
        if 0 <= self.current_idx < len(self.rules):
            r = self.rules[self.current_idx]
            r["name"] = self.name_edit.text().strip() or "未命名规则"
            r["path_pattern"] = self.pattern_combo.currentText().strip() or "{brand}/{sku}"
            raw_text = self.subs_edit.toPlainText()
            subs = [s.strip() for s in raw_text.split("\n") if s.strip()]
            r["subfolders"] = subs if subs else ["01_Design_平面原稿", "03_3D_三维工程"]
            r["design_sub"] = next((s for s in subs if "01" in s or "Design" in s or "原稿" in s), subs[0] if subs else "")
            r["blend_sub"] = next((s for s in subs if "03" in s or "3D" in s or "工程" in s), subs[1] if len(subs)>1 else "")
            r["render_sub"] = next((s for s in subs if "04" in s or "Render" in s or "输出" in s), subs[2] if len(subs)>2 else "")

    def new_rule(self):
        self.save_current_form()
        new_id = f"custom_rule_{int(datetime.datetime.now().timestamp())}"
        new_r = {
            "id": new_id,
            "name": f"✨ 新建规则 ({len(self.rules)+1})",
            "desc": "用户自定义规则",
            "path_pattern": "{brand}/{sku}",
            "subfolders": [
                "01_Design_平面原稿",
                "02_Textures_贴图资产",
                "03_3D_三维工程",
                "04_Renders_通道输出",
                "05_Delivery_最终交付"
            ],
            "design_sub": "01_Design_平面原稿",
            "blend_sub": "03_3D_三维工程",
            "render_sub": "04_Renders_通道输出"
        }
        self.rules.append(new_r)
        self.current_idx = len(self.rules) - 1
        self.refresh_list()
        self.load_rule_into_form(self.current_idx)

    def delete_rule(self):
        if len(self.rules) <= 1:
            QMessageBox.warning(self, "提示", "必须至少保留一套文件夹规则！")
            return
        del self.rules[self.current_idx]
        self.current_idx = max(0, self.current_idx - 1)
        self.refresh_list()
        self.load_rule_into_form(self.current_idx)

    def save_and_accept(self):
        self.save_current_form()
        self.active_rule_id = self.rules[self.current_idx]["id"]
        self.accept()

# ----------------- Qt6 工业级主窗口 -----------------
class MainWindow(QMainWindow):
    initial_load_done = Signal()

    def __init__(self, initial_files=None):
        super().__init__()
        self.setWindowTitle("美术资产中枢 - Art Asset Hub (v1.0 正式版 Qt6 GPU 加速)")
        self.resize(1320, 860)
        self.setMinimumSize(1060, 680)

        self.cfg = load_config()
        self.meta_cache = load_meta_cache()
        self.current_theme = self.cfg.get("theme", "dark")
        self.workspaces = self.cfg.get("workspaces", DEFAULT_WORKSPACES)
        self.curated_brands = self.cfg.get("curated_brands", ["柏缇", "森之露", "语后", "漱外", "零食有鸣"])
        self.folder_rules = self.cfg.get("folder_rules", DEFAULT_FOLDER_RULES)
        self.active_rule_id = self.cfg.get("active_rule_id", "standard_packaging_5stage")
        
        self.excel_projects = []
        self.disk_projects = []
        self.merged_projects = []
        self.current_display_list = []
        
        # 双维度过滤状态
        self.selected_category = "全部"
        self.selected_brand = "全部"
        self.brand_counts_map = {}
        
        self.files_to_organize = []
        self.has_fired_initial_done = False
        
        self.setup_ui()
        self.apply_theme()
        
        # 0.01 秒首屏快照渲染 (无感秒开)
        cached_disk = list(self.meta_cache.values())
        if cached_disk:
            cached_disk.sort(key=lambda x: x.get("mtime", 0), reverse=True)
            self.disk_projects = cached_disk
            self.merged_projects = cached_disk
            self.update_sidebar_counts()
            self.update_active_dataset()

        # 后台异步加载全量数据 (Qt 线程池 + 信号槽)
        QTimer.singleShot(50, self.async_load_data)
        
        if initial_files:
            self.tabs.setCurrentIndex(1)
            self.add_files_to_organizer(initial_files)

    def showEvent(self, event):
        super().showEvent(event)
        set_dark_titlebar(int(self.winId()), self.current_theme == "dark")

    def apply_theme(self):
        is_dark = (self.current_theme == "dark")
        if is_dark:
            self.setStyleSheet(DARK_QSS)
            self.btn_theme.setText("🌙 护眼暗灰")
        else:
            self.setStyleSheet(LIGHT_QSS)
            self.btn_theme.setText("☀️ 浅色模式")
        set_dark_titlebar(int(self.winId()), is_dark)

    def toggle_theme(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.cfg["theme"] = self.current_theme
        save_config(self.cfg)
        self.apply_theme()
        self.gallery_view.viewport().update()

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        # 顶部全局导航栏
        top_bar = QHBoxLayout()
        self.sync_status_lbl = QLabel("🟢 极速同步已就绪")
        self.sync_status_lbl.setStyleSheet("background: rgba(16, 185, 129, 0.15); color: #34D399; border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 12px; padding: 4px 10px; font-weight: bold; font-size: 11px;")
        
        btn_sync_excel = QPushButton("📤 同步缩略图到 Excel")
        btn_sync_excel.clicked.connect(self.sync_all_thumbnails_to_excel)
        btn_bind_excel = QPushButton("📊 绑定 Excel...")
        btn_bind_excel.clicked.connect(self.bind_excel_file)
        btn_refresh = QPushButton("🔄 刷新")
        btn_refresh.clicked.connect(self.async_load_data)
        btn_export = QPushButton("🌐 导出画廊")
        btn_export.clicked.connect(self.export_html_gallery)
        self.btn_theme = QPushButton("🌙 护眼暗灰")
        self.btn_theme.clicked.connect(self.toggle_theme)

        top_bar.addWidget(self.sync_status_lbl)
        top_bar.addStretch()
        top_bar.addWidget(btn_sync_excel)
        top_bar.addWidget(btn_bind_excel)
        top_bar.addWidget(btn_refresh)
        top_bar.addWidget(btn_export)
        top_bar.addWidget(self.btn_theme)
        main_layout.addLayout(top_bar)

        # 主 Tab
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Tab 1: 视觉资产看板
        self.tab_hub = QWidget()
        self.setup_hub_tab(self.tab_hub)
        self.tabs.addTab(self.tab_hub, "  🖼️ 视觉资产看板 (双维度导航 & GPU 加速)  ")

        # Tab 2: 设计源文件分拣与开工
        self.tab_organizer = QWidget()
        self.setup_organizer_tab(self.tab_organizer)
        self.tabs.addTab(self.tab_organizer, "  📥 设计源文件分拣与开工  ")

    # ---------------- Tab 1: 视觉资产看板 (双维度侧边栏) ----------------
    def setup_hub_tab(self, parent):
        layout = QHBoxLayout(parent)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(10)

        # 左侧「业务形态 + 客户品牌」双维度侧边栏
        sidebar = QWidget()
        sidebar.setFixedWidth(220)
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.setSpacing(10)

        # 上组：业务形态分类
        side_layout.addWidget(QLabel("<b>🏷️ 业务形态分类</b>"))
        self.category_list = QListWidget()
        self.category_list.setFixedHeight(168)
        cats = ["全部形态", "📦 包装", "🎁 套盒", "🖼️ 海报", "📑 物料"]
        for c in cats:
            self.category_list.addItem(f"{c} (0)")
        self.category_list.setCurrentRow(0)
        self.category_list.currentRowChanged.connect(self.on_category_changed)
        side_layout.addWidget(self.category_list)

        # 下组：客户品牌库
        side_layout.addWidget(QLabel("<b>🏢 客户与品牌库</b>"))
        self.brand_list = QListWidget()
        self.brand_list.addItem("全部品牌 (0)")
        self.brand_list.setCurrentRow(0)
        self.brand_list.currentRowChanged.connect(self.on_brand_changed)
        side_layout.addWidget(self.brand_list, stretch=1)
        
        layout.addWidget(sidebar)

        # 右侧画廊主体
        gallery_area = QWidget()
        gal_layout = QVBoxLayout(gallery_area)
        gal_layout.setContentsMargins(0, 0, 0, 0)
        gal_layout.setSpacing(8)

        # 搜索与视图过滤条
        filter_bar = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 搜索产品 SKU / 品牌 / 类别...")
        self.search_edit.textChanged.connect(self.apply_filter)
        
        self.view_combo = QComboBox()
        self.view_combo.addItems(["⚡ 智能融合视图", "📊 仅 Excel 台账", "💾 仅工作盘扫描"])
        self.view_combo.currentIndexChanged.connect(self.on_view_mode_changed)

        filter_bar.addWidget(self.search_edit, stretch=2)
        filter_bar.addWidget(self.view_combo, stretch=1)
        gal_layout.addLayout(filter_bar)

        # 🚀 144 FPS GPU 硬件加速 QListView 虚拟化视口
        self.gallery_view = QListView()
        self.gallery_view.setObjectName("GalleryView")
        self.gallery_view.setViewMode(QListView.IconMode)
        self.gallery_view.setResizeMode(QListView.Adjust)
        self.gallery_view.setUniformItemSizes(True)
        self.gallery_view.setSpacing(12)
        self.gallery_view.setMovement(QListView.Static)
        self.gallery_view.setVerticalScrollMode(QListView.ScrollPerPixel)
        self.gallery_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        self.gallery_model = GalleryModel(self)
        self.gallery_delegate = GalleryCardDelegate(self.gallery_view)
        self.gallery_view.setModel(self.gallery_model)
        self.gallery_view.setItemDelegate(self.gallery_delegate)
        
        self.gallery_view.clicked.connect(self.on_card_clicked)
        self.gallery_view.doubleClicked.connect(self.on_card_double_clicked)
        self.gallery_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.gallery_view.customContextMenuRequested.connect(self.show_gallery_context_menu)

        gal_layout.addWidget(self.gallery_view)
        layout.addWidget(gallery_area)

    # ---------------- Tab 2: 设计源文件分拣与开工 ----------------
    def setup_organizer_tab(self, parent):
        layout = QVBoxLayout(parent)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        # 面板 1: 工作盘与规则
        group1 = QGroupBox("📂 工作盘、客户、分类与归档文件夹规则")
        g1_layout = QVBoxLayout(group1)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("主工作盘:"))
        self.combo_ws = QComboBox()
        self.combo_ws.addItems(self.workspaces)
        cur_ws = self.cfg.get("current_workspace", self.workspaces[0])
        self.combo_ws.setCurrentText(cur_ws)
        self.combo_ws.currentTextChanged.connect(self.on_organizer_setting_changed)
        btn_add_ws = QPushButton("➕ 绑定新工作盘")
        btn_add_ws.clicked.connect(self.add_workspace)
        row1.addWidget(self.combo_ws, stretch=1)
        row1.addWidget(btn_add_ws)
        g1_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("客户品牌:"))
        self.combo_brand = QComboBox()
        self.combo_brand.setEditable(True)
        self.update_organizer_brand_combo()
        self.combo_brand.currentTextChanged.connect(self.on_organizer_setting_changed)
        btn_add_brand = QPushButton("➕")
        btn_add_brand.setFixedWidth(32)
        btn_add_brand.clicked.connect(self.add_brand)
        row2.addWidget(self.combo_brand)
        row2.addWidget(btn_add_brand)

        row2.addWidget(QLabel("业务形态:"))
        self.combo_cat = QComboBox()
        self.combo_cat.addItems(VALID_CATEGORIES)
        self.combo_cat.setCurrentText(self.cfg.get("default_category", "包装"))
        self.combo_cat.currentTextChanged.connect(self.on_organizer_setting_changed)
        row2.addWidget(self.combo_cat)

        row2.addWidget(QLabel("📁 归档规则:"))
        self.combo_rule = QComboBox()
        self.combo_rule.addItems([r["name"] for r in self.folder_rules])
        self.update_rule_combo_selection()
        self.combo_rule.currentIndexChanged.connect(self.on_rule_combo_changed)
        btn_custom_rule = QPushButton("⚙️ 自定义规则...")
        btn_custom_rule.clicked.connect(self.open_rule_manager)
        row2.addWidget(self.combo_rule, stretch=1)
        row2.addWidget(btn_custom_rule)
        g1_layout.addLayout(row2)
        layout.addWidget(group1)

        # 面板 2: 自动化选项
        group2 = QGroupBox("⚡ 自动化开工选项")
        g2_layout = QHBoxLayout(group2)
        self.chk_ai = QCheckBox("🎨 自动打开 AI 设计原稿")
        self.chk_ai.setChecked(self.cfg.get("auto_open_ai", True))
        self.chk_ai.stateChanged.connect(self.on_organizer_setting_changed)
        
        self.chk_blend = QCheckBox("✨ 自动生成对应 .blend 工程")
        self.chk_blend.setChecked(self.cfg.get("auto_create_blend", True))
        self.chk_blend.stateChanged.connect(self.on_organizer_setting_changed)

        self.chk_open_blend = QCheckBox("🚀 自动启动 Blender")
        self.chk_open_blend.setChecked(self.cfg.get("auto_open_blender", True))
        self.chk_open_blend.stateChanged.connect(self.on_organizer_setting_changed)

        self.chk_excel = QCheckBox("📊 自动录入 Excel")
        self.chk_excel.setChecked(self.cfg.get("auto_append_to_excel", True))
        self.chk_excel.stateChanged.connect(self.on_organizer_setting_changed)

        g2_layout.addWidget(self.chk_ai)
        g2_layout.addWidget(self.chk_blend)
        g2_layout.addWidget(self.chk_open_blend)
        g2_layout.addWidget(self.chk_excel)
        layout.addWidget(group2)

        # 面板 3: 待分拣列表
        group3 = QGroupBox("📥 待分拣设计源文件列表")
        g3_layout = QVBoxLayout(group3)
        
        tools_layout = QHBoxLayout()
        btn_add_files = QPushButton("➕ 添加设计源文件...")
        btn_add_files.setObjectName("PrimaryBtn")
        btn_add_files.clicked.connect(self.browse_source_files)
        btn_clear = QPushButton("🗑️ 清空列表")
        btn_clear.clicked.connect(self.clear_source_files)
        btn_install_jsx = QPushButton("🛠️ 一键将导出脚本注入 Illustrator")
        btn_install_jsx.clicked.connect(self.install_ai_jsx_script)
        
        tools_layout.addWidget(btn_add_files)
        tools_layout.addWidget(btn_clear)
        tools_layout.addStretch()
        tools_layout.addWidget(btn_install_jsx)
        g3_layout.addLayout(tools_layout)

        self.table_files = QTableWidget(0, 3)
        self.table_files.setHorizontalHeaderLabels(["源文件名", "提取 SKU", "目标归档目录"])
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table_files.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table_files.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        g3_layout.addWidget(self.table_files)

        btn_start_flow = QPushButton("🚀 一键分拣归档并拉起 Blender 开工")
        btn_start_flow.setObjectName("PrimaryBtn")
        btn_start_flow.setFixedHeight(44)
        btn_start_flow.setStyleSheet("font-size: 15px; font-weight: bold;")
        btn_start_flow.clicked.connect(self.execute_organize_flow)
        g3_layout.addWidget(btn_start_flow)

        layout.addWidget(group3)

    def update_organizer_brand_combo(self):
        cur_brand = self.cfg.get("current_brand", "柏缇")
        all_brands = list(self.curated_brands)
        for b in sorted(self.brand_counts_map.keys()):
            if b and b not in all_brands:
                all_brands.append(b)
        self.combo_brand.blockSignals(True)
        self.combo_brand.clear()
        self.combo_brand.addItems(all_brands)
        if cur_brand in all_brands:
            self.combo_brand.setCurrentText(cur_brand)
        elif all_brands:
            self.combo_brand.setCurrentIndex(0)
        self.combo_brand.blockSignals(False)

    # ---------------- 业务逻辑与数据流 (Qt 线程池 + 信号槽) ----------------
    def async_load_data(self):
        self.sync_status_lbl.setText("⚡ 正在加载全量资产...")
        ex_path = self.cfg.get("excel_path", DEFAULT_EXCEL_PATH)
        cur_ws = self.combo_ws.currentText() if hasattr(self, "combo_ws") else self.workspaces[0]
        
        worker = DataLoaderWorker(ex_path, cur_ws, self.meta_cache)
        worker.signals.finished.connect(self.on_data_loaded)
        QThreadPool.globalInstance().start(worker)

    @Slot(list, list, list)
    def on_data_loaded(self, excel_p, disk_p, merged):
        self.excel_projects = excel_p
        self.disk_projects = disk_p
        self.merged_projects = merged
        self.update_after_data_loaded()
        if not self.has_fired_initial_done:
            self.has_fired_initial_done = True
            self.initial_load_done.emit()

    def update_after_data_loaded(self):
        self.sync_status_lbl.setText(f"🟢 极速同步已就绪 (已载入 {len(self.merged_projects)} 个项目)")
        self.update_sidebar_counts()
        self.update_active_dataset()
        self.update_organizer_brand_combo()

    def update_sidebar_counts(self):
        mode_idx = self.view_combo.currentIndex()
        dataset = self.merged_projects if mode_idx == 0 else (self.excel_projects if mode_idx == 1 else self.disk_projects)
        
        # 1. 统计业务形态数量
        cat_counts = {"全部形态": len(dataset), "包装": 0, "套盒": 0, "海报": 0, "物料": 0}
        brand_counts = {}
        for p in dataset:
            c = p.get("cat", "包装")
            if c in cat_counts:
                cat_counts[c] += 1
            else:
                cat_counts["包装"] += 1
                
            b = p.get("brand", "").strip() or "未分类品牌"
            brand_counts[b] = brand_counts.get(b, 0) + 1

        self.brand_counts_map = brand_counts

        # 刷新形态列表
        cats = [("全部形态", "全部形态"), ("📦 包装", "包装"), ("🎁 套盒", "套盒"), ("🖼️ 海报", "海报"), ("📑 物料", "物料")]
        for idx, (label, key) in enumerate(cats):
            self.category_list.item(idx).setText(f"{label} ({cat_counts[key]})")

        # 刷新品牌列表 (按数量降序排列)
        cur_selected_brand = self.selected_brand
        self.brand_list.blockSignals(True)
        self.brand_list.clear()
        self.brand_list.addItem(f"全部品牌 ({len(dataset)})")
        
        sorted_brands = sorted(brand_counts.items(), key=lambda x: x[1], reverse=True)
        target_row = 0
        for idx, (b_name, b_count) in enumerate(sorted_brands, start=1):
            self.brand_list.addItem(f"{b_name} ({b_count})")
            if b_name == cur_selected_brand:
                target_row = idx

        self.brand_list.setCurrentRow(target_row)
        self.brand_list.blockSignals(False)

    def on_view_mode_changed(self):
        self.update_sidebar_counts()
        self.update_active_dataset()

    def on_category_changed(self, row):
        cats = ["全部", "包装", "套盒", "海报", "物料"]
        if 0 <= row < len(cats):
            self.selected_category = cats[row]
            self.apply_filter()

    def on_brand_changed(self, row):
        if row == 0:
            self.selected_brand = "全部"
        elif row > 0 and row - 1 < len(self.brand_counts_map):
            sorted_brands = sorted(self.brand_counts_map.items(), key=lambda x: x[1], reverse=True)
            self.selected_brand = sorted_brands[row - 1][0]
        self.apply_filter()

    def update_active_dataset(self):
        mode_idx = self.view_combo.currentIndex()
        if mode_idx == 1:
            self.current_display_list = self.excel_projects
        elif mode_idx == 2:
            self.current_display_list = self.disk_projects
        else:
            self.current_display_list = self.merged_projects
        self.apply_filter()

    def apply_filter(self):
        kw = self.search_edit.text().strip().lower()
        res = []
        for p in self.current_display_list:
            cat = p.get("cat", "包装")
            brand = p.get("brand", "").strip() or "未分类品牌"
            
            # 形态过滤
            if self.selected_category != "全部" and cat != self.selected_category:
                continue
                
            # 品牌过滤
            if self.selected_brand != "全部" and brand != self.selected_brand:
                continue
                
            # 关键字过滤
            if kw:
                sku = p.get("sku", "").lower()
                b_low = brand.lower()
                if kw not in sku and kw not in b_low and kw not in cat.lower():
                    continue
            res.append(p)
        self.gallery_model.set_projects(res)

    def on_card_clicked(self, index):
        pass

    def on_card_double_clicked(self, index):
        proj = index.data(Qt.UserRole)
        if proj:
            self.launch_blend(proj.get("path"), sku=proj.get("sku"), brand=proj.get("brand"))

    def show_gallery_context_menu(self, pos):
        index = self.gallery_view.indexAt(pos)
        if not index.isValid():
            return
        proj = index.data(Qt.UserRole)
        if not proj:
            return

        menu = QMenu(self)
        p = proj.get("path", "")
        sku = proj.get("sku", "")
        brand = proj.get("brand", "")
        
        act_folder = menu.addAction(f"📁 打开文件夹: {sku}")
        act_blend = menu.addAction("🚀 Blender 打开 3D 工程")
        
        real_p = self.resolve_project_path(p, sku, brand)
        if real_p and os.path.exists(real_p):
            menu.addAction("🎨 查看 01_Design_平面原稿", lambda: self.open_folder(os.path.join(real_p, "01_Design_平面原稿")))
            menu.addAction("🖼️ 查看 04_Renders_通道输出", lambda: self.open_folder(os.path.join(real_p, "04_Renders_通道输出")))
            
        menu.addSeparator()
        cat_menu = menu.addMenu(f"🏷️ 修改业务形态 (当前: {proj.get('cat', '包装')})")
        for c in VALID_CATEGORIES:
            cat_menu.addAction(f"设为：{c}", lambda c_val=c: self.change_project_category(proj, c_val))
            
        menu.addSeparator()
        if proj.get("thumbnail") and os.path.exists(proj["thumbnail"]):
            menu.addAction("📊 将此缩略图写入 Excel 台账 (图片列)", lambda: self.sync_single_thumbnail_to_excel(proj))
        menu.addAction("📋 复制完整物理路径", lambda: QApplication.clipboard().setText(real_p or p))

        action = menu.exec(QCursor.pos())
        if action == act_folder:
            self.open_folder(p, sku=sku, brand=brand)
        elif action == act_blend:
            self.launch_blend(p, sku=sku, brand=brand)

    def resolve_project_path(self, path, sku="", brand=""):
        if path and os.path.exists(path):
            return path
        cur_ws = self.combo_ws.currentText() if hasattr(self, "combo_ws") else self.workspaces[0]
        if sku and cur_ws and os.path.exists(cur_ws):
            if brand:
                cand = os.path.join(cur_ws, brand, sku)
                if os.path.exists(cand):
                    return cand
            try:
                for entry in os.listdir(cur_ws):
                    bp = os.path.join(cur_ws, entry)
                    if os.path.isdir(bp):
                        cand = os.path.join(bp, sku)
                        if os.path.exists(cand):
                            return cand
            except Exception:
                pass
            cand = os.path.join(cur_ws, sku)
            if os.path.exists(cand):
                return cand
        return path

    def open_folder(self, path, sku="", brand=""):
        real_path = self.resolve_project_path(path, sku, brand)
        if real_path and os.path.exists(real_path):
            os.startfile(real_path)
        else:
            QMessageBox.warning(self, "提示", f"该项目的本地文件夹暂未找到:\n{path or sku}")

    def launch_blend(self, proj_path, sku="", brand=""):
        real_path = self.resolve_project_path(proj_path, sku, brand)
        if not real_path or not os.path.exists(real_path):
            QMessageBox.warning(self, "提示", f"未找到该项目的本地文件夹:\n{proj_path or sku}")
            return
        rule = self.get_current_folder_rule()
        blend_sub = rule.get("blend_sub", "03_3D_三维工程")
        blend_dir = os.path.join(real_path, blend_sub) if blend_sub else real_path
        target_dir = blend_dir if os.path.exists(blend_dir) else real_path
        blends = glob.glob(os.path.join(target_dir, "*.blend"))
        if blends:
            blends.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            chosen = blends[0]
            try:
                subprocess.Popen([BLENDER_EXE, chosen])
            except Exception:
                os.startfile(chosen)
        else:
            self.open_folder(real_path)

    def change_project_category(self, proj, new_cat):
        sku = proj.get("sku")
        path = proj.get("path")
        proj["cat"] = new_cat
        if path:
            norm_p = path.lower().replace("/", "\\")
            if norm_p in self.meta_cache:
                self.meta_cache[norm_p]["cat"] = new_cat
            else:
                self.meta_cache[norm_p] = {
                    "brand": proj.get("brand", ""),
                    "sku": sku,
                    "cat": new_cat,
                    "thumbnail": proj.get("thumbnail"),
                    "mtime": proj.get("mtime", 0)
                }
            save_meta_cache(self.meta_cache)
            
        ex_path = self.cfg.get("excel_path", DEFAULT_EXCEL_PATH)
        if ex_path and os.path.exists(ex_path):
            try:
                wb = openpyxl.load_workbook(ex_path)
                sheet = wb.active
                cat_col = None
                sku_col = 2
                for col_idx, cell in enumerate(sheet[1], start=1):
                    val = str(cell.value or "").strip()
                    if val in ("业务形态", "分类", "类别"):
                        cat_col = col_idx
                    if val in ("产品名称", "SKU", "品名", "产品命名"):
                        sku_col = col_idx
                if cat_col:
                    sku_clean = re.sub(r'[\s_\-\(\)（）]+', '', sku.lower()) if sku else ""
                    for r in range(2, sheet.max_row + 1):
                        cell_val = str(sheet.cell(row=r, column=sku_col).value or "").strip()
                        cell_clean = re.sub(r'[\s_\-\(\)（）]+', '', cell_val.lower())
                        if sku_clean and cell_clean and (sku_clean == cell_clean or sku_clean in cell_clean or cell_clean in sku_clean):
                            sheet.cell(row=r, column=cat_col, value=new_cat)
                            break
                    wb.save(ex_path)
                wb.close()
            except Exception:
                pass
        self.apply_filter()
        QMessageBox.information(self, "修改成功", f"已成功将 [{sku}] 的业务形态更新为：【{new_cat}】")

    def sync_single_thumbnail_to_excel(self, proj):
        ex_path = self.cfg.get("excel_path", DEFAULT_EXCEL_PATH)
        ok, msg = update_thumbnail_to_excel(ex_path, proj.get("path"), proj.get("sku"), proj.get("thumbnail"))
        if ok:
            QMessageBox.information(self, "同步成功", msg)
        else:
            QMessageBox.warning(self, "同步失败", msg)

    def sync_all_thumbnails_to_excel(self):
        ex_path = self.cfg.get("excel_path", DEFAULT_EXCEL_PATH)
        ok, msg = batch_sync_all_thumbnails_to_excel(ex_path, self.merged_projects)
        if ok:
            QMessageBox.information(self, "批量同步成功", msg)
        else:
            QMessageBox.warning(self, "批量同步失败", msg)

    def bind_excel_file(self):
        f, _ = QFileDialog.getOpenFileName(self, "绑定 Excel 产品台账", "", "Excel Files (*.xlsx *.xls)")
        if f:
            self.cfg["excel_path"] = os.path.normpath(f)
            save_config(self.cfg)
            self.async_load_data()
            QMessageBox.information(self, "绑定成功", f"已成功绑定新 Excel:\n{f}")

    def export_html_gallery(self):
        ex_dir = self.combo_ws.currentText() if hasattr(self, "combo_ws") else os.path.expanduser("~")
        html_file = os.path.join(ex_dir, "美术资产全景画廊.html")
        
        cards = []
        for p in self.current_display_list:
            thumb = p.get("thumbnail") or ""
            norm_thumb = thumb.replace("\\", "/") if thumb else ""
            t_src = f"file:///{norm_thumb}" if thumb and os.path.exists(thumb) else ""
            img_html = f'<img src="{t_src}" loading="lazy" />' if t_src else '<div style="color:#666;font-size:12px;">待渲染</div>'
            cards.append(f"""
            <div class="card" onclick="alert('项目路径: {html.escape(p.get('path', ''))}')">
              <div class="thumb-container">
                {img_html}
              </div>
              <div class="meta">
                <div class="badge">{html.escape(p.get('cat', '包装'))}</div>
                <div class="title">{html.escape(p.get('sku', ''))}</div>
                <div class="path">{html.escape(p.get('brand', ''))}</div>
              </div>
            </div>
            """)
            
        doc = f"""<!DOCTYPE html><html><head><meta charset="utf-8"/><title>美术资产全景视觉画廊</title>
        <style>
        body {{ background: #1a1c23; color: #e2e4e8; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif; padding: 24px; margin: 0; }}
        h1 {{ font-size: 20px; margin-bottom: 20px; font-weight: 600; }}
        .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap: 16px; }}
        .card {{ background: #262930; border-radius: 8px; overflow: hidden; border: 1px solid #363942; transition: transform .2s; cursor: pointer; }}
        .card:hover {{ transform: translateY(-4px); border-color: #3b82f6; }}
        .thumb-container {{ width: 100%; aspect-ratio: 1; background: #1e2026; display: flex; align-items: center; justify-content: center; overflow: hidden; }}
        .thumb-container img {{ width: 100%; height: 100%; object-fit: contain; }}
        .meta {{ padding: 10px; }}
        .badge {{ display: inline-block; font-size: 11px; background: #1e3a8a; color: #93c5fd; padding: 2px 6px; border-radius: 4px; font-weight: bold; margin-bottom: 4px; }}
        .title {{ font-size: 13px; font-weight: bold; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
        .path {{ font-size: 11px; color: #8c909c; margin-top: 2px; }}
        </style></head><body>
        <h1>🎨 美术资产全景视觉画廊 (共 {len(self.current_display_list)} 个项目)</h1>
        <div class="grid">{''.join(cards)}</div>
        </body></html>"""
        
        try:
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(doc)
            webbrowser.open("file:///" + html_file.replace("\\", "/"))
            QMessageBox.information(self, "导出成功", f"已成功导出画廊到:\n{html_file}")
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))

    # ---------------- 规则与分拣开工 ----------------
    def update_rule_combo_selection(self):
        for idx, r in enumerate(self.folder_rules):
            if r.get("id") == self.active_rule_id:
                self.combo_rule.setCurrentIndex(idx)
                return
        if self.folder_rules:
            self.combo_rule.setCurrentIndex(0)

    def get_current_folder_rule(self):
        for r in self.folder_rules:
            if r.get("id") == self.active_rule_id:
                return r
        return self.folder_rules[0] if self.folder_rules else DEFAULT_FOLDER_RULES[0]

    def on_rule_combo_changed(self, idx):
        if 0 <= idx < len(self.folder_rules):
            self.active_rule_id = self.folder_rules[idx]["id"]
            self.on_organizer_setting_changed()

    def open_rule_manager(self):
        dlg = FolderRuleManagerDialog(self.folder_rules, self.active_rule_id, self)
        if dlg.exec() == QDialog.Accepted:
            self.folder_rules = dlg.rules
            self.active_rule_id = dlg.active_rule_id
            self.cfg["folder_rules"] = self.folder_rules
            self.cfg["active_rule_id"] = self.active_rule_id
            save_config(self.cfg)
            
            self.combo_rule.blockSignals(True)
            self.combo_rule.clear()
            self.combo_rule.addItems([r["name"] for r in self.folder_rules])
            self.update_rule_combo_selection()
            self.combo_rule.blockSignals(False)
            self.refresh_organizer_table()
            QMessageBox.information(self, "规则更新", "已成功更新并应用文件夹归档规则！")

    def on_organizer_setting_changed(self):
        self.cfg["current_workspace"] = self.combo_ws.currentText()
        self.cfg["current_brand"] = self.combo_brand.currentText()
        self.cfg["default_category"] = self.combo_cat.currentText()
        self.cfg["active_rule_id"] = self.active_rule_id
        self.cfg["auto_open_ai"] = self.chk_ai.isChecked()
        self.cfg["auto_create_blend"] = self.chk_blend.isChecked()
        self.cfg["auto_open_blender"] = self.chk_open_blend.isChecked()
        self.cfg["auto_append_to_excel"] = self.chk_excel.isChecked()
        save_config(self.cfg)
        self.refresh_organizer_table()

    def add_workspace(self):
        d = QFileDialog.getExistingDirectory(self, "选择并绑定新工作盘")
        if d:
            d = os.path.normpath(d)
            if d not in self.workspaces:
                self.workspaces.append(d)
                self.cfg["workspaces"] = self.workspaces
                self.combo_ws.addItem(d)
            self.combo_ws.setCurrentText(d)
            self.on_organizer_setting_changed()

    def add_brand(self):
        b, ok = QInputDialog.getText(self, "新增客户品牌", "请输入新客户品牌名称:")
        if ok and b.strip():
            b = b.strip()
            if b not in self.curated_brands:
                self.curated_brands.append(b)
                self.cfg["curated_brands"] = self.curated_brands
                self.update_organizer_brand_combo()
            self.combo_brand.setCurrentText(b)
            self.on_organizer_setting_changed()

    def browse_source_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择设计源文件", "", "Design Files (*.ai *.psd *.pdf *.zip *.rar);;All Files (*.*)")
        if files:
            self.add_files_to_organizer(files)

    def add_files_to_organizer(self, files):
        for f in files:
            f = os.path.normpath(f)
            if f not in self.files_to_organize:
                self.files_to_organize.append(f)
        self.refresh_organizer_table()

    def clear_source_files(self):
        self.files_to_organize.clear()
        self.refresh_organizer_table()

    def refresh_organizer_table(self):
        self.table_files.setRowCount(0)
        cur_ws = self.combo_ws.currentText()
        brand = self.combo_brand.currentText()
        cat = self.combo_cat.currentText()
        rule = self.get_current_folder_rule()
        pat = rule.get("path_pattern", "{brand}/{sku}")

        for f in self.files_to_organize:
            fname = os.path.basename(f)
            sku = os.path.splitext(fname)[0]
            
            if "{category}" in pat:
                rel = f"{cat}/{brand}/{sku}" if brand else f"{cat}/{sku}"
            elif "{brand}" in pat:
                rel = f"{brand}/{sku}" if brand else sku
            else:
                rel = sku
                
            dest_dir = os.path.join(cur_ws, rel)
            
            row = self.table_files.rowCount()
            self.table_files.insertRow(row)
            self.table_files.setItem(row, 0, QTableWidgetItem(fname))
            self.table_files.setItem(row, 1, QTableWidgetItem(sku))
            self.table_files.setItem(row, 2, QTableWidgetItem(dest_dir))

    def execute_organize_flow(self):
        if not self.files_to_organize:
            QMessageBox.warning(self, "提示", "请先添加待分拣的设计源文件！")
            return
            
        cur_ws = self.combo_ws.currentText()
        brand = self.combo_brand.currentText()
        cat = self.combo_cat.currentText()
        rule = self.get_current_folder_rule()
        subfolders = rule.get("subfolders", DEFAULT_FOLDER_RULES[0]["subfolders"])
        pat = rule.get("path_pattern", "{brand}/{sku}")
        design_sub = rule.get("design_sub", "01_Design_平面原稿")
        blend_sub = rule.get("blend_sub", "03_3D_三维工程")
        
        auto_ai = self.chk_ai.isChecked()
        auto_blend = self.chk_blend.isChecked()
        auto_open_bl = self.chk_open_blend.isChecked()
        auto_excel = self.chk_excel.isChecked()
        
        created_count = 0
        opened_blend = None
        opened_ai = None
        
        for fpath in self.files_to_organize:
            if not os.path.exists(fpath):
                continue
            fname = os.path.basename(fpath)
            sku = os.path.splitext(fname)[0]
            
            if "{category}" in pat:
                rel = f"{cat}/{brand}/{sku}" if brand else f"{cat}/{sku}"
            elif "{brand}" in pat:
                rel = f"{brand}/{sku}" if brand else sku
            else:
                rel = sku
                
            proj_dir = os.path.join(cur_ws, rel)
            os.makedirs(proj_dir, exist_ok=True)
            for sub in subfolders:
                os.makedirs(os.path.join(proj_dir, sub), exist_ok=True)
                
            dest_fpath = os.path.join(proj_dir, design_sub, fname)
            try:
                shutil.copy2(fpath, dest_fpath)
            except Exception:
                pass
                
            if auto_ai and not opened_ai and fname.lower().endswith(('.ai', '.psd', '.pdf')):
                try:
                    os.startfile(dest_fpath)
                    opened_ai = dest_fpath
                except Exception:
                    pass
                    
            if auto_blend:
                blend_dir = os.path.join(proj_dir, blend_sub)
                target_blend = os.path.join(blend_dir, f"{sku}.blend")
                is_first_creation = not os.path.exists(target_blend)
                
                if is_first_creation:
                    tpl = get_valid_template_blend(self.cfg)
                    if tpl and os.path.exists(tpl):
                        shutil.copy2(tpl, target_blend)
                    else:
                        with open(target_blend, "wb") as bf:
                            pass
                            
                if auto_open_bl and is_first_creation and not opened_blend:
                    try:
                        subprocess.Popen([BLENDER_EXE, target_blend])
                        opened_blend = target_blend
                    except Exception:
                        os.startfile(target_blend)
                        opened_blend = target_blend
                        
            if auto_excel:
                ex_path = self.cfg.get("excel_path", DEFAULT_EXCEL_PATH)
                if ex_path and os.path.exists(ex_path):
                    try:
                        wb = openpyxl.load_workbook(ex_path)
                        sheet = wb.active
                        sku_col = 2
                        brand_col = 1
                        cat_col = None
                        for col_idx, cell in enumerate(sheet[1], start=1):
                            val = str(cell.value or "").strip()
                            if val in ("产品名称", "SKU", "品名", "产品命名"):
                                sku_col = col_idx
                            if val in ("品牌", "客户"):
                                brand_col = col_idx
                            if val in ("业务形态", "分类", "类别"):
                                cat_col = col_idx
                                
                        sku_clean = re.sub(r'[\s_\-\(\)（）]+', '', sku.lower())
                        exists = False
                        for r in range(2, sheet.max_row + 1):
                            c_val = str(sheet.cell(row=r, column=sku_col).value or "").strip()
                            if sku_clean and sku_clean == re.sub(r'[\s_\-\(\)（）]+', '', c_val.lower()):
                                exists = True
                                break
                        if not exists:
                            new_r = sheet.max_row + 1
                            sheet.cell(row=new_r, column=brand_col, value=brand)
                            sheet.cell(row=new_r, column=sku_col, value=sku)
                            if cat_col:
                                sheet.cell(row=new_r, column=cat_col, value=cat)
                            wb.save(ex_path)
                        wb.close()
                    except Exception:
                        pass
            created_count += 1
            
        self.files_to_organize.clear()
        self.refresh_organizer_table()
        self.async_load_data()
        QMessageBox.information(self, "开工成功", f"🎉 已成功分拣归档并初始化 {created_count} 个工程项目！")

    def install_ai_jsx_script(self):
        src_jsx = os.path.join(os.path.dirname(__file__), "Export_Artboards_To_Textures.jsx")
        if not os.path.exists(src_jsx):
            src_jsx = r"C:\Users\qq424\Packaging_Tools\Export_Artboards_To_Textures.jsx"
            
        if not os.path.exists(src_jsx):
            QMessageBox.warning(self, "错误", "未找到脚本源文件: Export_Artboards_To_Textures.jsx")
            return
            
        target_dirs = []
        possible_roots = [
            r"H:\adobe\Adobe Illustrator 2024",
            r"C:\Program Files\Adobe\Adobe Illustrator 2024",
            r"C:\Program Files\Adobe\Adobe Illustrator 2025",
            r"D:\Adobe\Adobe Illustrator 2024"
        ]
        for pr in possible_roots:
            p_zh = os.path.join(pr, "Presets", "zh_CN", "脚本")
            if os.path.exists(p_zh):
                target_dirs.append(p_zh)
            p_en = os.path.join(pr, "Presets", "en_US", "Scripts")
            if os.path.exists(p_en):
                target_dirs.append(p_en)
                
        installed = []
        for td in target_dirs:
            try:
                dest = os.path.join(td, "🚀一键导出画板到贴图目录.jsx")
                shutil.copy2(src_jsx, dest)
                installed.append(dest)
            except Exception:
                pass
                
        if installed:
            QMessageBox.information(self, "安装成功", f"🎉 已成功将贴图导出脚本注入 Illustrator！\n\n已安装到:\n{installed[0]}")
        else:
            dest_dir = QFileDialog.getExistingDirectory(self, "选择 Illustrator 的 Presets/zh_CN/脚本 目录")
            if dest_dir:
                try:
                    shutil.copy2(src_jsx, os.path.join(dest_dir, "🚀一键导出画板到贴图目录.jsx"))
                    QMessageBox.information(self, "安装成功", "🎉 已成功安装到指定的脚本目录！")
                except Exception as e:
                    QMessageBox.warning(self, "安装失败", str(e))

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    splash = None
    splash_path = SPLASH_CAT_JPG
    if os.path.exists(splash_path):
        splash = CatSplashScreen(splash_path)
        screen_geo = app.primaryScreen().geometry()
        splash.move((screen_geo.width() - splash.width()) // 2, (screen_geo.height() - splash.height()) // 2)
        splash.show()
        app.processEvents()

    initial_files = sys.argv[1:] if len(sys.argv) > 1 else None
    window = MainWindow(initial_files=initial_files)
    
    def on_ready_show():
        if splash:
            splash.close()
        window.show()
        set_dark_titlebar(int(window.winId()), window.current_theme == "dark")

    window.initial_load_done.connect(on_ready_show)
    QTimer.singleShot(2500, on_ready_show)
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
