# -*- coding: utf-8 -*-
"""
美术资产中枢 (Art Asset Hub - v1.0 正式版)
===================================================================
【核心功能体系】：
1. 🖼️ 视觉资产看板 (Visual Asset Hub)：
   - 物理磁盘 15KB JPEG 缩略图缓存池 + 30 槽位复用池 (Widget Pooling)，极速瞬切 < 5ms；
   - 自动融合 Excel 产品台账与本地硬盘工程；
   - 4 大业务形态精准分类：📦 包装 (默认) / 🎁 套盒 / 🖼️ 海报 / 📑 物料；
   - 自动抓取最新 Beauty 渲染图为封面，支持导出全景 HTML 交互式画廊。

2. 📥 设计源文件分拣与开工 (Source Organizer & Pipeline Launcher)：
   - 右键“发送到”一键唤起；
   - ⚙️ 自定义文件夹归档规则管理器：自由新建/编辑子目录结构，内置目录树实时预览；
   - 自动生成对应 Blender 母版工程并拉起 Blender 5.2 LTS 开工；
   - 自动将新项目录入《产品列表.xlsx》。

3. 🌿 温润工业石墨灰护眼设计 (Studio Graphite Eye-Care Dark Mode)：
   - 告别高反差死黑，对标 Blender/Lightroom/Eagle 工业级调色，柔和防眩光。
===================================================================
"""

import os
import sys
import re
import json
import glob
import shutil
import hashlib
import zipfile
import datetime
import threading
import webbrowser
import subprocess
import concurrent.futures
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from PIL import Image, ImageTk, ImageDraw
import openpyxl

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_suite_v7.json")
CACHE_DIR = os.path.join(os.path.expanduser("~"), ".packaging_asset_thumbnails")
EXCEL_CACHE_DIR = os.path.join(CACHE_DIR, "excel_images")
THUMB_CACHE_DIR = os.path.join(CACHE_DIR, "fast_thumbs")
META_CACHE_FILE = os.path.join(CACHE_DIR, "disk_meta_cache.json")
os.makedirs(EXCEL_CACHE_DIR, exist_ok=True)
os.makedirs(THUMB_CACHE_DIR, exist_ok=True)

APP_ICON_ICO = os.path.join(os.path.dirname(__file__), "app_icon.ico")
if not os.path.exists(APP_ICON_ICO):
    APP_ICON_ICO = r"C:\Users\qq424\Packaging_Tools\app_icon.ico"
APP_ICON_PNG = os.path.join(os.path.dirname(__file__), "app_icon.png")
if not os.path.exists(APP_ICON_PNG):
    APP_ICON_PNG = r"C:\Users\qq424\Packaging_Tools\app_icon.png"

DEFAULT_EXCEL_PATH = r"C:\Users\qq424\WorkBuddy\2026-08-26-15-33-05\产品列表.xlsx"

BLENDER_EXE = r"C:\Program Files\Blender Foundation\Blender 5.2\blender.exe"
if not os.path.exists(BLENDER_EXE):
    BLENDER_EXE = "blender"

DEFAULT_TEMPLATE = os.path.join(
    os.path.expanduser("~"), "Desktop", "AI_Blender包装渲染辅助工具", "templates", "Packaging_Master_Template.blend"
)

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

# 默认预设的文件夹规则库
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

THEMES = {
    "dark": {
        "bg": "#222429",
        "header_bg": "#2A2C32",
        "sidebar_bg": "#2A2C32",
        "panel_bg": "#2A2C32",
        "card_bg": "#30333A",
        "card_border": "#3D414A",
        "card_hover": "#4C84C4",
        "canvas_bg": "#222429",
        "thumb_bg": "#27292F",
        "input_bg": "#1A1B1F",
        "input_fg": "#E2E4E8",
        "fg": "#E2E4E8",
        "fg_muted": "#9699A2",
        "fg_dim": "#656872",
        "primary": "#3B78B8",
        "primary_hover": "#4989CE",
        "primary_fg": "#FFFFFF",
        "badge_brand_bg": "#3A3D46",
        "badge_brand_fg": "#B2B6BE",
        "cat_colors": {
            "包装": ("#26384C", "#96C2EC"),
            "套盒": ("#443522", "#E8C88A"),
            "海报": ("#382845", "#D2A2E8"),
            "物料": ("#223D32", "#8EE2BE")
        },
        "btn_secondary_bg": "#3A3D46",
        "btn_secondary_fg": "#D2D5DA",
        "status_bg": "#1D3328",
        "status_fg": "#6EE7B7"
    },
    "light": {
        "bg": "#F4F6F9",
        "header_bg": "#FFFFFF",
        "sidebar_bg": "#FFFFFF",
        "panel_bg": "#FFFFFF",
        "card_bg": "#FFFFFF",
        "card_border": "#E2E6EC",
        "card_hover": "#0284C7",
        "canvas_bg": "#F4F6F9",
        "thumb_bg": "#F8FAFC",
        "input_bg": "#FFFFFF",
        "input_fg": "#0F172A",
        "fg": "#1E293B",
        "fg_muted": "#64748B",
        "fg_dim": "#94A3B8",
        "primary": "#0078D7",
        "primary_hover": "#005A9E",
        "primary_fg": "#FFFFFF",
        "badge_brand_bg": "#F1F5F9",
        "badge_brand_fg": "#475569",
        "cat_colors": {
            "包装": ("#E1EFFF", "#005A9E"),
            "套盒": ("#FEF3C7", "#B45309"),
            "海报": ("#EDE9FE", "#6D28D9"),
            "物料": ("#ECFDF5", "#047857")
        },
        "btn_secondary_bg": "#F1F5F9",
        "btn_secondary_fg": "#1E293B",
        "status_bg": "#ECFDF5",
        "status_fg": "#059669"
    }
}

DEFAULT_CONFIG = {
    "workspaces": DEFAULT_WORKSPACES,
    "current_workspace": DEFAULT_WORKSPACES[0],
    "excel_path": DEFAULT_EXCEL_PATH if os.path.exists(DEFAULT_EXCEL_PATH) else "",
    "curated_brands": ["柏缇", "零食有鸣"],
    "current_brand": "柏缇",
    "default_category": "包装",
    "theme": "dark",
    "auto_create_blend": True,
    "auto_open_blender": True,
    "template_blend_path": DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else "",
    "auto_append_to_excel": True,
    "folder_rules": DEFAULT_FOLDER_RULES,
    "active_rule_id": "standard_packaging_5stage"
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in DEFAULT_CONFIG.items():
                    if k not in data:
                        data[k] = v
                # 检查 rules 是否完整
                if not data.get("folder_rules"):
                    data["folder_rules"] = DEFAULT_FOLDER_RULES
                return data
        except Exception:
            pass
    return DEFAULT_CONFIG

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_meta_cache():
    if os.path.exists(META_CACHE_FILE):
        try:
            with open(META_CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_meta_cache(cache):
    try:
        with open(META_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def get_fast_disk_thumbnail_path(orig_img_path, size=(190, 190)):
    if not orig_img_path or not os.path.exists(orig_img_path):
        return None
    try:
        mtime = os.path.getmtime(orig_img_path)
        h = hashlib.md5(f"{orig_img_path}_{mtime}_{size}".encode('utf-8')).hexdigest()
        cached_thumb_file = os.path.join(THUMB_CACHE_DIR, f"{h}.jpg")
        
        if os.path.exists(cached_thumb_file):
            return cached_thumb_file
            
        im = Image.open(orig_img_path)
        if im.mode != "RGB":
            im = im.convert("RGB")
        im.thumbnail(size, Image.Resampling.BILINEAR)
        im.save(cached_thumb_file, format="JPEG", quality=85)
        return cached_thumb_file
    except Exception:
        return orig_img_path


def normalize_category(raw_cat):
    if not raw_cat:
        return "包装"
    raw = str(raw_cat).strip()
    if any(k in raw for k in ["套盒", "礼盒", "套装", "组合装", "礼品装", "kit", "giftbox"]):
        return "套盒"
    elif any(k in raw.lower() for k in ["海报", "kv", "主视觉", "展板", "poster"]):
        return "海报"
    elif any(k in raw.lower() for k in ["物料", "易拉宝", "台卡", "展架", "堆头", "折页", "画册", "dm", "posm"]):
        return "物料"
    else:
        return "包装"


def auto_detect_category_from_name(filename):
    fn = filename.lower()
    if any(k in fn for k in ["套盒", "礼盒", "套装", "组合装", "礼品装", "kit", "giftbox"]):
        return "套盒"
    elif any(k in fn for k in ["海报", "kv", "主视觉", "展板", "poster"]):
        return "海报"
    elif any(k in fn for k in ["物料", "易拉宝", "台卡", "展架", "堆头", "折页", "画册", "dm", "posm"]):
        return "物料"
    else:
        return "包装"


def append_project_to_excel(excel_path, brand, sku, cat, proj_path):
    if not excel_path or not os.path.exists(excel_path):
        return False
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb['全部'] if '全部' in wb.sheetnames else wb.active
        norm_cat = normalize_category(cat)
        norm_proj_p = proj_path.replace('\\', '/').lower().strip('/')
        
        for r in range(2, ws.max_row + 1):
            ex_p = str(ws.cell(row=r, column=5).value or "").replace('\\', '/').lower().strip('/')
            ex_name = str(ws.cell(row=r, column=2).value or "").strip()
            if ex_name == sku or (norm_proj_p and ex_p == norm_proj_p):
                ws.cell(row=r, column=4, value=norm_cat)
                wb.save(excel_path)
                return False
                
        next_idx = ws.max_row
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        new_row = [next_idx, sku, None, norm_cat, proj_path.replace('\\', '/'), now_str]
        ws.append(new_row)
        try:
            ws.row_dimensions[ws.max_row].height = 40
        except Exception:
            pass
        wb.save(excel_path)
        return True
    except Exception as e:
        print(f"Error appending project to Excel: {e}")
        return False


def update_project_category_in_excel(excel_path, proj_path, sku, new_cat):
    if not excel_path or not os.path.exists(excel_path):
        return False
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb['全部'] if '全部' in wb.sheetnames else wb.active
        norm_cat = normalize_category(new_cat)
        norm_proj_p = proj_path.replace('\\', '/').lower().strip('/') if proj_path else ""
        
        updated = False
        for r in range(2, ws.max_row + 1):
            ex_p = str(ws.cell(row=r, column=5).value or "").replace('\\', '/').lower().strip('/')
            ex_name = str(ws.cell(row=r, column=2).value or "").strip()
            if (norm_proj_p and ex_p == norm_proj_p) or (sku and ex_name == sku):
                ws.cell(row=r, column=4, value=norm_cat)
                updated = True
                break
        if updated:
            wb.save(excel_path)
        return updated
    except Exception as e:
        print(f"Error updating category in Excel: {e}")
        return False


def get_file_md5(filepath):
    try:
        h = hashlib.md5()
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


JUNK_NAME_KEYWORDS = {
    '改', '改1', '改2', '改3', '改4', '改5', '修改', '修改版', '最新', '最终', '最终版', '定稿', '正稿',
    '未命名', '未命名-1', '新建', '新建画板', '新建画板1', '新建画板2', '1', '2', '3', 'a', 'b', 'c',
    '刀模', '包装', '瓶贴', '贴纸', '展开图', '画板', '副本', '111', '222', 'aaa'
}

def clean_and_parse_filename(filepath, fallback_brand="", valid_brands=None):
    raw_name = os.path.splitext(os.path.basename(filepath))[0]
    cleaned_base = re.sub(r'[\(\（]\s*\d+\s*[\)\）]', '', raw_name)
    cleaned_base = re.sub(r'[-_ ]*副本\s*\d*', '', cleaned_base)
    
    noise_patterns = [
        r'[-_ ]?(包装|刀模|展开图|正稿|定稿|完稿|原稿|印刷稿|平面|效果图)',
        r'[-_ ]?(修改版|修改|最新版|最终版|最终|定案|终版|初稿|打样|打样稿)',
        r'[-_ ]?(副本|\d{6,}|\d{4}年|\d{1,2}月\d{1,2}日)',
        r'[-_ ]?([vV]\d+(\.\d+)?|改\d*|版\d*)',
    ]
    for p in noise_patterns:
        cleaned_base = re.sub(p, '', cleaned_base, flags=re.IGNORECASE)
        
    cleaned_base = cleaned_base.strip(" -_")
    is_junk = cleaned_base.lower() in JUNK_NAME_KEYWORDS or len(cleaned_base) == 0
    
    brand = ""
    sku = ""
    if valid_brands:
        for vb in valid_brands:
            if cleaned_base.startswith(vb):
                brand = vb
                sku = cleaned_base[len(vb):].strip(" -_")
                break
                
    if not brand:
        parts = re.split(r'[-_—\s+]+', cleaned_base)
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) >= 2:
            brand = parts[0]
            sku = "_".join(parts[1:])
        elif len(parts) == 1:
            if is_junk:
                brand = fallback_brand
                sku = f"未命名_{raw_name}"
            else:
                brand = fallback_brand
                sku = parts[0]
        else:
            brand = fallback_brand
            sku = f"未命名_{raw_name}"
            
    return brand, sku, is_junk


def find_project_thumbnail(proj_path):
    if not proj_path or not os.path.exists(proj_path):
        return None

    render_dirs = [
        os.path.join(proj_path, "04_Renders_通道输出"),
        os.path.join(proj_path, "04_Renders_高清分层输出"),
        os.path.join(proj_path, "03_输出"),
        os.path.join(proj_path, "05_Delivery_最终交付"),
        os.path.join(proj_path, "05_Final_精修定稿"),
        os.path.join(proj_path, "渲染"),
        os.path.join(proj_path, "Renders"),
        proj_path
    ]
    candidates = []
    for rdir in render_dirs:
        if os.path.exists(rdir):
            try:
                for ext in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
                    candidates.extend(glob.glob(os.path.join(rdir, ext)))
            except Exception:
                pass
                
    if candidates:
        beauty_imgs = [
            c for c in candidates 
            if any(k in os.path.basename(c).lower() for k in ["beauty", "成品", "主图", "camera", "正面", "01_", "main", "render"])
            and not any(bad in os.path.basename(c).lower() for bad in ["mask", "alpha", "crypto", "选区", "蒙版", "normal", "depth", "roughness", "ao_"])
        ]
        if beauty_imgs:
            try:
                beauty_imgs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return beauty_imgs[0]
            except Exception:
                return beauty_imgs[0]
                
        filtered = [
            c for c in candidates 
            if not any(bad in os.path.basename(c).lower() for bad in ["mask", "alpha", "crypto", "选区", "蒙版", "normal", "depth", "roughness", "ao_", "diffuse"])
        ]
        if filtered:
            try:
                filtered.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return filtered[0]
            except Exception:
                return filtered[0]
                
        try:
            candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return candidates[0]
        except Exception:
            return candidates[0]
            
    return None


def parse_and_cache_excel(excel_path):
    projects = []
    if not excel_path or not os.path.exists(excel_path):
        return projects

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = wb['全部'] if '全部' in wb.sheetnames else wb.active
        
        row_image_map = {}
        with zipfile.ZipFile(excel_path, 'r') as z:
            if 'xl/drawings/drawing1.xml' in z.namelist() and 'xl/drawings/_rels/drawing1.xml.rels' in z.namelist():
                drawing_xml = z.read('xl/drawings/drawing1.xml')
                rels_xml = z.read('xl/drawings/_rels/drawing1.xml.rels')

                root_d = ET.fromstring(drawing_xml)
                root_r = ET.fromstring(rels_xml)

                rel_map = {}
                for rel in root_r:
                    r_id = rel.attrib.get('Id')
                    target = rel.attrib.get('Target')
                    t_clean = target.lstrip('/').replace('../', '')
                    if not t_clean.startswith('xl/'):
                        t_clean = 'xl/' + t_clean
                    rel_map[r_id] = t_clean

                ns = {
                    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                }

                for anchor in root_d.findall('.//xdr:twoCellAnchor', ns) + root_d.findall('.//xdr:oneCellAnchor', ns):
                    from_elem = anchor.find('xdr:from', ns)
                    if from_elem is not None:
                        row_idx = int(from_elem.find('xdr:row', ns).text) + 1
                        blip = anchor.find('.//a:blip', ns)
                        if blip is not None:
                            embed_id = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                            if embed_id and embed_id in rel_map:
                                media_path = rel_map[embed_id]
                                ext = os.path.splitext(media_path)[1]
                                out_filename = f"excel_img_r{row_idx}{ext}"
                                out_full = os.path.join(EXCEL_CACHE_DIR, out_filename)
                                with open(out_full, 'wb') as f_out:
                                    f_out.write(z.read(media_path))
                                row_image_map[row_idx] = out_full

        for r in range(2, sheet.max_row + 1):
            name = sheet.cell(row=r, column=2).value
            raw_cat = sheet.cell(row=r, column=4).value or "包装"
            path = sheet.cell(row=r, column=5).value or ""
            time_str = sheet.cell(row=r, column=6).value or ""

            if not name:
                continue

            thumb = row_image_map.get(r, None)
            norm_path = path.replace("/", "\\") if path else ""

            brand = "柏缇"
            if norm_path:
                parts = norm_path.strip("\\").split("\\")
                if len(parts) >= 2 and parts[-2] not in {"zjc", "Projects", "E:", "D:", "包装", "套盒", "海报", "物料"}:
                    brand = parts[-2]

            projects.append({
                "source": "excel",
                "brand": brand,
                "sku": str(name).strip(),
                "cat": normalize_category(raw_cat),
                "path": norm_path,
                "thumbnail": thumb,
                "time": str(time_str),
                "mtime": os.path.getmtime(norm_path) if (norm_path and os.path.exists(norm_path)) else 0
            })

    except Exception as e:
        print(f"Error parsing Excel: {e}")

    return projects


def scan_workspace_projects_fast(root_dir, meta_cache):
    projects = []
    if not os.path.exists(root_dir):
        return projects
    try:
        entries = os.listdir(root_dir)
    except (PermissionError, OSError):
        return projects
        
    cache_dirty = False
    
    for entry in entries:
        if entry.lower() in SYSTEM_IGNORED_DIRS or entry.startswith('.') or entry.startswith('$') or entry.startswith('_'):
            continue
        brand_p = os.path.join(root_dir, entry)
        try:
            if not os.path.isdir(brand_p):
                continue
            sub_entries = os.listdir(brand_p)
        except (PermissionError, OSError):
            continue
            
        for sku in sub_entries:
            if sku.lower() in SYSTEM_IGNORED_DIRS or sku.startswith('.') or sku.startswith('$') or sku.startswith('_'):
                continue
            sku_p = os.path.join(brand_p, sku)
            try:
                if os.path.isdir(sku_p):
                    s_mtime = os.path.getmtime(sku_p)
                    cache_key = sku_p.lower().replace("/", "\\")
                    
                    if cache_key in meta_cache and meta_cache[cache_key].get("mtime") == s_mtime:
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
        norm_p = dp["path"].lower().replace("/", "\\").strip("\\")
        disk_map[norm_p] = dp
        sku_key = f"{dp['brand']}@@{dp['sku']}".lower()
        disk_map[sku_key] = dp

    merged = []
    handled_disk_keys = set()

    for ep in excel_projects:
        norm_ep = ep["path"].lower().replace("/", "\\").strip("\\") if ep.get("path") else ""
        sku_key = f"{ep.get('brand', '')}@@{ep.get('sku', '')}".lower()
        
        matched_disk_proj = None
        if norm_ep and norm_ep in disk_map:
            matched_disk_proj = disk_map[norm_ep]
        elif sku_key in disk_map:
            matched_disk_proj = disk_map[sku_key]

        if matched_disk_proj:
            dp = matched_disk_proj
            handled_disk_keys.add(dp["path"].lower().replace("/", "\\").strip("\\"))
            thumb = dp["thumbnail"] if dp["thumbnail"] else ep["thumbnail"]
            mtime = dp["mtime"] if dp["mtime"] > 0 else ep["mtime"]
            merged.append({
                "source": "disk_prioritized",
                "brand": dp["brand"],
                "sku": dp["sku"],
                "cat": normalize_category(ep.get("cat", dp.get("cat", "包装"))),
                "path": dp["path"],
                "thumbnail": thumb,
                "time": ep.get("time", ""),
                "mtime": mtime
            })
        else:
            merged.append(ep)

    for dp in disk_projects:
        norm_p = dp["path"].lower().replace("/", "\\").strip("\\")
        if norm_p not in handled_disk_keys:
            merged.append(dp)

    merged.sort(key=lambda x: x.get("mtime", 0), reverse=True)
    return merged


# ==============================================================================
# 文件夹规则管理器弹窗 (Folder Rules Manager Modal Dialog)
# ==============================================================================
class FolderRuleManagerDialog(tk.Toplevel):
    def __init__(self, parent, rules, active_rule_id, on_save_callback, colors):
        super().__init__(parent)
        self.title("⚙️ 自定义文件夹归档规则管理器")
        self.geometry("820x620")
        self.minsize(720, 520)
        self.colors = colors
        self.configure(bg=colors["bg"])
        self.transient(parent)
        self.grab_set()
        
        self.rules = [dict(r) for r in rules]  # deep copy
        self.active_rule_id = active_rule_id
        self.on_save_callback = on_save_callback
        self.current_rule_idx = 0
        
        # 寻找当前激活的规则索引
        for i, r in enumerate(self.rules):
            if r.get("id") == self.active_rule_id:
                self.current_rule_idx = i
                break
                
        self.build_ui()
        self.load_rule_to_editor(self.current_rule_idx)

    def build_ui(self):
        c = self.colors
        
        # 主布局左右分栏
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)
        
        # 左侧：规则列表
        left_frame = ttk.LabelFrame(paned, text=" 📂 规则库列表 ", padding=8, width=240)
        paned.add(left_frame, weight=1)
        
        self.rule_listbox = tk.Listbox(
            left_frame,
            font=("Microsoft YaHei", 9),
            bg=c["panel_bg"],
            fg=c["fg"],
            selectbackground=c["primary"],
            selectforeground="#FFFFFF",
            relief=tk.FLAT,
            highlightthickness=0,
            activestyle="none"
        )
        self.rule_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.rule_listbox.bind("<<ListboxSelect>>", self.on_select_rule_list)
        
        left_btn_row = ttk.Frame(left_frame)
        left_btn_row.pack(fill=tk.X)
        ttk.Button(left_btn_row, text="➕ 新建", width=6, command=self.add_new_rule).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(left_btn_row, text="📋 复制", width=6, command=self.duplicate_rule).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(left_btn_row, text="🗑️ 删除", width=6, command=self.delete_rule).pack(side=tk.LEFT)
        
        # 右侧：规则编辑器
        right_frame = ttk.LabelFrame(paned, text=" 🛠️ 规则详细配置与目录树预览 ", padding=12)
        paned.add(right_frame, weight=3)
        
        # 1. 规则名称 & 描述
        f_info = ttk.Frame(right_frame)
        f_info.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(f_info, text="规则名称:").grid(row=0, column=0, sticky="w", pady=3)
        self.name_var = tk.StringVar()
        ttk.Entry(f_info, textvariable=self.name_var, width=32, font=("Microsoft YaHei", 9, "bold")).grid(row=0, column=1, sticky="ew", padx=(8, 0), pady=3)
        
        ttk.Label(f_info, text="规则说明:").grid(row=1, column=0, sticky="w", pady=3)
        self.desc_var = tk.StringVar()
        ttk.Entry(f_info, textvariable=self.desc_var, width=32).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=3)
        f_info.columnconfigure(1, weight=1)
        
        # 2. 根路径层级组织
        f_pattern = ttk.Frame(right_frame)
        f_pattern.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(f_pattern, text="目录层级架构:").pack(side=tk.LEFT, padx=(0, 8))
        self.pattern_var = tk.StringVar(value="{brand}/{sku}")
        self.combo_pattern = ttk.Combobox(
            f_pattern,
            textvariable=self.pattern_var,
            values=[
                "{brand}/{sku} (推荐：工作盘/客户/SKU)",
                "{category}/{brand}/{sku} (工作盘/业务形态/客户/SKU)",
                "{sku} (工作盘/仅SKU目录)"
            ],
            state="readonly",
            width=36
        )
        self.combo_pattern.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.combo_pattern.bind("<<ComboboxSelected>>", lambda e: self.update_tree_preview())

        # 3. 子文件夹列表编辑
        f_subs = ttk.LabelFrame(right_frame, text=" 📁 子文件夹清单 (每行一个文件夹，将自动按顺序创建) ", padding=8)
        f_subs.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        
        self.txt_subs = tk.Text(
            f_subs,
            height=6,
            font=("Consolas", 10),
            bg=c["input_bg"],
            fg=c["input_fg"],
            insertbackground=c["fg"],
            relief=tk.FLAT,
            bd=1
        )
        self.txt_subs.pack(fill=tk.BOTH, expand=True)
        self.txt_subs.bind("<KeyRelease>", lambda e: self.on_subs_text_changed())

        # 4. 自动分流关联设置
        f_mapping = ttk.Frame(right_frame)
        f_mapping.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(f_mapping, text="源文件存入:").grid(row=0, column=0, sticky="w", pady=3)
        self.design_sub_var = tk.StringVar()
        self.combo_design = ttk.Combobox(f_mapping, textvariable=self.design_sub_var, width=22)
        self.combo_design.grid(row=0, column=1, sticky="w", padx=(6, 16), pady=3)
        
        ttk.Label(f_mapping, text="3D工程生成于:").grid(row=0, column=2, sticky="w", pady=3)
        self.blend_sub_var = tk.StringVar()
        self.combo_blend = ttk.Combobox(f_mapping, textvariable=self.blend_sub_var, width=22)
        self.combo_blend.grid(row=0, column=3, sticky="w", padx=(6, 0), pady=3)

        ttk.Label(f_mapping, text="渲染图输出于:").grid(row=1, column=0, sticky="w", pady=3)
        self.render_sub_var = tk.StringVar()
        self.combo_render = ttk.Combobox(f_mapping, textvariable=self.render_sub_var, width=22)
        self.combo_render.grid(row=1, column=1, sticky="w", padx=(6, 16), pady=3)
        
        # 5. 实时效果预览
        f_prev = ttk.LabelFrame(right_frame, text=" 🌲 生成目录效果实时预览 ", padding=6)
        f_prev.pack(fill=tk.X, pady=(0, 8))
        
        self.lbl_preview = tk.Label(
            f_prev,
            text="",
            font=("Consolas", 9),
            bg=c["panel_bg"],
            fg=c["fg_muted"],
            justify="left",
            anchor="w",
            padx=6,
            pady=4
        )
        self.lbl_preview.pack(fill=tk.X)

        # 底部保存按钮行
        b_bar = ttk.Frame(self, padding=(16, 8))
        b_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        ttk.Button(b_bar, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=(8, 0))
        btn_apply = tk.Button(
            b_bar,
            text="💾 保存并设为当前活动规则",
            font=("Microsoft YaHei", 9, "bold"),
            bg=c["primary"],
            fg="#FFFFFF",
            relief=tk.FLAT,
            padx=16,
            pady=6,
            command=self.save_and_apply
        )
        btn_apply.pack(side=tk.RIGHT)
        
        self.refresh_rule_listbox()

    def refresh_rule_listbox(self):
        self.rule_listbox.delete(0, tk.END)
        for i, r in enumerate(self.rules):
            active_mark = " (当前)" if r.get("id") == self.active_rule_id else ""
            self.rule_listbox.insert(tk.END, f"{r.get('name', '未命名规则')}{active_mark}")
            
        if 0 <= self.current_rule_idx < len(self.rules):
            self.rule_listbox.select_set(self.current_rule_idx)

    def load_rule_to_editor(self, idx):
        if not (0 <= idx < len(self.rules)):
            return
        self.current_rule_idx = idx
        r = self.rules[idx]
        
        self.name_var.set(r.get("name", ""))
        self.desc_var.set(r.get("desc", ""))
        
        pat = r.get("path_pattern", "{brand}/{sku}")
        if "{category}" in pat:
            self.combo_pattern.current(1)
        elif "{brand}" in pat:
            self.combo_pattern.current(0)
        else:
            self.combo_pattern.current(2)
            
        subs = r.get("subfolders", [])
        self.txt_subs.delete("1.0", tk.END)
        self.txt_subs.insert(tk.END, "\n".join(subs))
        
        self.on_subs_text_changed(initial=True)
        self.design_sub_var.set(r.get("design_sub", subs[0] if subs else ""))
        self.blend_sub_var.set(r.get("blend_sub", subs[2] if len(subs) > 2 else (subs[0] if subs else "")))
        self.render_sub_var.set(r.get("render_sub", subs[3] if len(subs) > 3 else (subs[0] if subs else "")))
        
        self.update_tree_preview()

    def save_current_editor_to_memory(self):
        if not (0 <= self.current_rule_idx < len(self.rules)):
            return
        subs = [line.strip() for line in self.txt_subs.get("1.0", tk.END).split("\n") if line.strip()]
        pat_raw = self.combo_pattern.get()
        if "{category}" in pat_raw:
            pat = "{category}/{brand}/{sku}"
        elif "{brand}" in pat_raw:
            pat = "{brand}/{sku}"
        else:
            pat = "{sku}"
            
        r = self.rules[self.current_rule_idx]
        r["name"] = self.name_var.get().strip() or "自定义规则"
        r["desc"] = self.desc_var.get().strip()
        r["path_pattern"] = pat
        r["subfolders"] = subs
        r["design_sub"] = self.design_sub_var.get().strip() or (subs[0] if subs else "")
        r["blend_sub"] = self.blend_sub_var.get().strip() or (subs[2] if len(subs) > 2 else (subs[0] if subs else ""))
        r["render_sub"] = self.render_sub_var.get().strip() or (subs[3] if len(subs) > 3 else (subs[0] if subs else ""))

    def on_select_rule_list(self, event=None):
        sel = self.rule_listbox.curselection()
        if not sel:
            return
        self.save_current_editor_to_memory()
        self.load_rule_to_editor(sel[0])

    def on_subs_text_changed(self, initial=False):
        subs = [line.strip() for line in self.txt_subs.get("1.0", tk.END).split("\n") if line.strip()]
        self.combo_design["values"] = subs
        self.combo_blend["values"] = subs
        self.combo_render["values"] = subs
        
        if not initial:
            if subs:
                if self.design_sub_var.get() not in subs:
                    self.design_sub_var.set(subs[0])
                if self.blend_sub_var.get() not in subs:
                    self.blend_sub_var.set(subs[2] if len(subs) > 2 else subs[0])
                if self.render_sub_var.get() not in subs:
                    self.render_sub_var.set(subs[3] if len(subs) > 3 else (subs[-1] if subs else ""))
            self.update_tree_preview()

    def update_tree_preview(self):
        pat_raw = self.combo_pattern.get()
        if "{category}" in pat_raw:
            base_dir = "E:\\zjc\\包装\\柏缇\\玫瑰水乳"
        elif "{brand}" in pat_raw:
            base_dir = "E:\\zjc\\柏缇\\玫瑰水乳"
        else:
            base_dir = "E:\\zjc\\玫瑰水乳"
            
        subs = [line.strip() for line in self.txt_subs.get("1.0", tk.END).split("\n") if line.strip()]
        d_sub = self.design_sub_var.get()
        b_sub = self.blend_sub_var.get()
        r_sub = self.render_sub_var.get()
        
        lines = [f"📁 {base_dir}"]
        for i, s in enumerate(subs):
            prefix = "└── " if i == len(subs) - 1 else "├── "
            tags = []
            if s == d_sub:
                tags.append("📥 源文件归档处")
            if s == b_sub:
                tags.append("✨ 自动生成.blend")
            if s == r_sub:
                tags.append("🖼️ 渲染图封面来源")
            tag_str = f"  <-- [{' | '.join(tags)}]" if tags else ""
            lines.append(f"{prefix}📁 {s}{tag_str}")
            
        if not subs:
            lines.append("└── ⚠️ (请在上方输入至少一个子目录名称)")
            
        self.lbl_preview.config(text="\n".join(lines))

    def add_new_rule(self):
        self.save_current_editor_to_memory()
        new_id = f"custom_rule_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        new_rule = {
            "id": new_id,
            "name": f"✨ 自定义规则 {len(self.rules) + 1}",
            "desc": "用户自定义的个性化文件夹结构规则",
            "path_pattern": "{brand}/{sku}",
            "subfolders": ["01_Design_平面原稿", "02_Textures_贴图资产", "03_3D_三维工程", "04_Renders_通道输出", "05_Delivery_最终交付"],
            "design_sub": "01_Design_平面原稿",
            "blend_sub": "03_3D_三维工程",
            "render_sub": "04_Renders_通道输出"
        }
        self.rules.append(new_rule)
        self.current_rule_idx = len(self.rules) - 1
        self.refresh_rule_listbox()
        self.load_rule_to_editor(self.current_rule_idx)

    def duplicate_rule(self):
        if not (0 <= self.current_rule_idx < len(self.rules)):
            return
        self.save_current_editor_to_memory()
        cur = self.rules[self.current_rule_idx]
        new_id = f"custom_rule_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        new_rule = dict(cur)
        new_rule["id"] = new_id
        new_rule["name"] = f"{cur.get('name', '规则')} (副本)"
        new_rule["subfolders"] = list(cur.get("subfolders", []))
        self.rules.append(new_rule)
        self.current_rule_idx = len(self.rules) - 1
        self.refresh_rule_listbox()
        self.load_rule_to_editor(self.current_rule_idx)

    def delete_rule(self):
        if len(self.rules) <= 1:
            messagebox.showwarning("提示", "必须至少保留一个文件夹归档规则！")
            return
        r = self.rules[self.current_rule_idx]
        if messagebox.askyesno("确认删除", f"确定要删除规则【{r.get('name')}】吗？"):
            del self.rules[self.current_rule_idx]
            if self.active_rule_id == r.get("id"):
                self.active_rule_id = self.rules[0]["id"]
            self.current_rule_idx = max(0, self.current_rule_idx - 1)
            self.refresh_rule_listbox()
            self.load_rule_to_editor(self.current_rule_idx)

    def save_and_apply(self):
        self.save_current_editor_to_memory()
        cur = self.rules[self.current_rule_idx]
        self.active_rule_id = cur.get("id")
        self.on_save_callback(self.rules, self.active_rule_id)
        self.destroy()


# ==============================================================================
# 主应用程序 (Main Packaging Studio Suite)
# ==============================================================================
class PackagingStudioSuite:
    def __init__(self, root, initial_files=None):
        self.root = root
        self.root.title("美术资产中枢 - Art Asset Hub (v1.0 正式版)")
        self.root.geometry("1260x830")
        self.root.minsize(1020, 660)
        
        self.cfg = load_config()
        self.meta_cache = load_meta_cache()
        self.current_theme = self.cfg.get("theme", "dark")
        self.colors = THEMES.get(self.current_theme, THEMES["dark"])
        
        self.workspaces = self.cfg.get("workspaces", DEFAULT_WORKSPACES)
        cur_ws = self.cfg.get("current_workspace", self.workspaces[0])
        if not os.path.exists(cur_ws) and self.workspaces:
            cur_ws = self.workspaces[0]
        self.current_workspace_var = tk.StringVar(value=cur_ws)
        self.excel_path_var = tk.StringVar(value=self.cfg.get("excel_path", DEFAULT_EXCEL_PATH))
        
        # 文件夹规则库
        self.folder_rules = self.cfg.get("folder_rules", DEFAULT_FOLDER_RULES)
        self.active_rule_id = self.cfg.get("active_rule_id", "standard_packaging_5stage")
        self.active_rule_name_var = tk.StringVar()
        self.update_active_rule_name_var()
        
        # 归档页变量
        self.curated_brands = self.cfg.get("curated_brands", ["柏缇", "零食有鸣"])
        self.current_brand_var = tk.StringVar(value=self.cfg.get("current_brand", self.curated_brands[0]))
        self.current_cat_var = tk.StringVar(value=self.cfg.get("default_category", "包装"))
        self.auto_create_blend_var = tk.BooleanVar(value=self.cfg.get("auto_create_blend", True))
        self.auto_open_blender_var = tk.BooleanVar(value=self.cfg.get("auto_open_blender", True))
        self.auto_append_excel_var = tk.BooleanVar(value=self.cfg.get("auto_append_to_excel", True))
        self.files_to_organize = []
        
        # 资产看板页变量
        self.view_mode_var = tk.StringVar(value="merged")
        self.search_var = tk.StringVar()
        self.selected_category_var = tk.StringVar(value="全部")
        self.page_size = 30
        self.current_page = 0
        self.last_excel_mtime = 0
        self.search_debounce_job = None
        
        self.excel_projects = []
        self.disk_projects = []
        self.merged_projects = []
        self.current_display_list = []
        self.filtered_projects = []
        
        self.thumb_tk_cache = {}
        self.card_slots = []
        
        self.load_app_icon()
        self.setup_styles()
        self.build_ui()
        self.init_card_slots(30)
        self.load_all_asset_data()
        self.start_excel_auto_sync_watcher()
        
        if initial_files:
            self.notebook.select(1)
            self.add_files_to_organizer(initial_files)
        else:
            self.notebook.select(0)

    def update_active_rule_name_var(self):
        for r in self.folder_rules:
            if r.get("id") == self.active_rule_id:
                self.active_rule_name_var.set(r.get("name", "标准规则"))
                return
        if self.folder_rules:
            self.active_rule_id = self.folder_rules[0]["id"]
            self.active_rule_name_var.set(self.folder_rules[0]["name"])

    def get_current_folder_rule(self):
        for r in self.folder_rules:
            if r.get("id") == self.active_rule_id:
                return r
        return self.folder_rules[0] if self.folder_rules else DEFAULT_FOLDER_RULES[0]

    def load_app_icon(self):
        if os.path.exists(APP_ICON_ICO):
            try:
                self.root.iconbitmap(APP_ICON_ICO)
            except Exception:
                pass
        if os.path.exists(APP_ICON_PNG):
            try:
                img = Image.open(APP_ICON_PNG)
                photo = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, photo)
                self._app_icon_ref = photo
            except Exception:
                pass

    def setup_styles(self):
        c = self.colors
        self.root.configure(bg=c["bg"])
        
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
            
        style.configure(".", background=c["bg"], foreground=c["fg"], font=("Microsoft YaHei", 9))
        style.configure("TNotebook", background=c["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=c["header_bg"], foreground=c["fg_muted"], font=("Microsoft YaHei", 10), padding=[20, 7], borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", c["primary"])],
                  foreground=[("selected", c["primary_fg"])])
                  
        style.configure("TFrame", background=c["bg"])
        style.configure("TLabelframe", background=c["bg"], foreground=c["fg"], bordercolor=c["card_border"])
        style.configure("TLabelframe.Label", background=c["bg"], foreground=c["primary_hover"], font=("Microsoft YaHei", 9, "bold"))
        
        style.configure("TLabel", background=c["bg"], foreground=c["fg"])
        
        style.configure("TButton", background=c["btn_secondary_bg"], foreground=c["btn_secondary_fg"], font=("Microsoft YaHei", 9), padding=[8, 4], borderwidth=0)
        style.map("TButton",
                  background=[("active", c["primary"]), ("pressed", c["primary"])],
                  foreground=[("active", "#FFFFFF"), ("pressed", "#FFFFFF")])
                  
        style.configure("Primary.TButton", background=c["primary"], foreground="#FFFFFF", font=("Microsoft YaHei", 9, "bold"), padding=[10, 5], borderwidth=0)
        style.map("Primary.TButton", background=[("active", c["primary_hover"])])
        
        style.configure("TEntry", fieldbackground=c["input_bg"], foreground=c["input_fg"], insertcolor=c["fg"], bordercolor=c["card_border"])
        style.configure("TCombobox", fieldbackground=c["input_bg"], background=c["btn_secondary_bg"], foreground=c["input_fg"], bordercolor=c["card_border"], arrowcolor=c["fg_muted"])
        style.map("TCombobox", fieldbackground=[("readonly", c["input_bg"])], foreground=[("readonly", c["input_fg"])])

        style.configure("Treeview", background=c["panel_bg"], foreground=c["fg"], fieldbackground=c["panel_bg"], rowheight=26, font=("Microsoft YaHei", 9), borderwidth=0)
        style.configure("Treeview.Heading", background=c["header_bg"], foreground=c["fg_muted"], font=("Microsoft YaHei", 9, "bold"), borderwidth=1, relief=tk.FLAT)
        style.map("Treeview", background=[("selected", c["primary"])], foreground=[("selected", "#FFFFFF")])

    def toggle_theme(self):
        new_theme = "light" if self.current_theme == "dark" else "dark"
        self.current_theme = new_theme
        self.cfg["theme"] = new_theme
        self.colors = THEMES[new_theme]
        save_config(self.cfg)
        
        self.setup_styles()
        self.theme_btn.config(text="☀️ 浅色模式" if new_theme == "dark" else "🌙 护眼暗灰")
        self.restyle_all_ui()
        self.thumb_tk_cache.clear()
        self.render_cards()

    def restyle_all_ui(self):
        c = self.colors
        self.root.configure(bg=c["bg"])
        self.canvas.configure(bg=c["canvas_bg"])
        self.grid_container.configure(bg=c["canvas_bg"])
        self.category_listbox.configure(bg=c["panel_bg"], fg=c["fg"], selectbackground=c["primary"])
        self.sync_status_lbl.configure(bg=c["status_bg"], fg=c["status_fg"])
        
        for slot in self.card_slots:
            slot["card"].configure(bg=c["card_bg"], highlightbackground=c["card_border"])
            slot["img_lbl"].configure(bg=c["thumb_bg"])
            slot["meta_frame"].configure(bg=c["card_bg"])
            slot["badge_row"].configure(bg=c["card_bg"])
            slot["title_lbl"].configure(bg=c["card_bg"], fg=c["fg"])
            slot["action_frame"].configure(bg=c["card_bg"])
            slot["btn_open"].configure(bg=c["btn_secondary_bg"], fg=c["btn_secondary_fg"])
            slot["btn_blend"].configure(bg=c["primary"])

    def build_ui(self):
        c = self.colors
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: 视觉资产看板
        self.tab_assets = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_assets, text="  🖼️ 视觉资产看板  ")
        self.build_asset_hub_ui(self.tab_assets)
        
        # Tab 2: 设计源文件分拣与开工
        self.tab_organizer = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_organizer, text="  📥 设计源文件分拣与开工  ")
        self.build_organizer_ui(self.tab_organizer)

    # ---------------- 页面 1: 视觉资产看板 ----------------
    def build_asset_hub_ui(self, parent):
        c = self.colors
        
        top_bar = ttk.Frame(parent, padding=(16, 10))
        top_bar.pack(fill=tk.X)
        
        ttk.Label(top_bar, text="📊 视图:").pack(side=tk.LEFT, padx=(0, 4))
        self.combo_source = ttk.Combobox(
            top_bar,
            textvariable=self.view_mode_var,
            values=["merged", "excel", "disk"],
            state="readonly",
            width=22,
            font=("Microsoft YaHei", 9)
        )
        self.combo_source.pack(side=tk.LEFT, padx=(0, 14))
        self.combo_source.bind("<<ComboboxSelected>>", self.on_source_change)
        
        ttk.Label(top_bar, text="🔍 搜索:").pack(side=tk.LEFT, padx=(0, 4))
        search_entry = ttk.Entry(top_bar, textvariable=self.search_var, font=("Microsoft YaHei", 9), width=18)
        search_entry.pack(side=tk.LEFT, padx=(0, 12))
        self.search_var.trace_add("write", lambda *args: self.on_search_change_debounced())
        
        ttk.Button(top_bar, text="📊 绑定 Excel...", command=self.import_new_excel).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top_bar, text="🔄 刷新", command=self.load_all_asset_data).pack(side=tk.LEFT, padx=(0, 6))
        
        self.sync_status_lbl = tk.Label(
            top_bar,
            text="🟢 极速同步已就绪",
            font=("Microsoft YaHei", 8, "bold"),
            fg=c["status_fg"],
            bg=c["status_bg"],
            padx=8,
            pady=3
        )
        self.sync_status_lbl.pack(side=tk.LEFT, padx=(6, 0))
        
        self.theme_btn = ttk.Button(top_bar, text="☀️ 浅色模式" if self.current_theme == "dark" else "🌙 护眼暗灰", command=self.toggle_theme)
        self.theme_btn.pack(side=tk.RIGHT, padx=(8, 0))
        
        ttk.Button(top_bar, text="🌐 导出全景画廊 (HTML)", command=self.export_html_gallery).pack(side=tk.RIGHT)

        # 底栏分页
        self.bottom_bar = ttk.Frame(parent, padding=(16, 8))
        self.bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.page_info_lbl = ttk.Label(self.bottom_bar, text="", font=("Microsoft YaHei", 9))
        self.page_info_lbl.pack(side=tk.LEFT)
        
        self.btn_next = ttk.Button(self.bottom_bar, text="下一页 ➡️", command=self.next_page)
        self.btn_next.pack(side=tk.RIGHT, padx=(6, 0))
        
        self.btn_prev = ttk.Button(self.bottom_bar, text="⬅️ 上一页", command=self.prev_page)
        self.btn_prev.pack(side=tk.RIGHT)

        # 主视口
        main_pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 6))
        
        left_frame = ttk.LabelFrame(main_pane, text=" 🏷️ 形态分类 & 筛选 ", padding=8, width=200)
        main_pane.add(left_frame, weight=1)
        
        self.category_listbox = tk.Listbox(
            left_frame,
            font=("Microsoft YaHei", 10),
            selectmode=tk.SINGLE,
            relief=tk.FLAT,
            bg=c["panel_bg"],
            fg=c["fg"],
            selectbackground=c["primary"],
            selectforeground="white",
            highlightthickness=0,
            activestyle="none"
        )
        self.category_listbox.pack(fill=tk.BOTH, expand=True)
        self.category_listbox.bind("<<ListboxSelect>>", self.on_category_select)
        
        right_frame = ttk.Frame(main_pane)
        main_pane.add(right_frame, weight=5)
        
        self.canvas = tk.Canvas(right_frame, bg=c["canvas_bg"], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        
        self.grid_container = tk.Frame(self.canvas, bg=c["canvas_bg"])
        self.grid_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_container, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

    # ---------------- 卡片槽位复用池 (Widget Pool) ----------------
    def init_card_slots(self, count=30):
        c = self.colors
        for i in range(count):
            card = tk.Frame(
                self.grid_container,
                bg=c["card_bg"],
                bd=0,
                padx=8,
                pady=8,
                highlightthickness=1,
                highlightbackground=c["card_border"]
            )
            
            img_lbl = tk.Label(card, bg=c["thumb_bg"], cursor="hand2")
            img_lbl.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
            
            meta_frame = tk.Frame(card, bg=c["card_bg"], pady=4)
            meta_frame.pack(fill=tk.X)
            
            badge_row = tk.Frame(meta_frame, bg=c["card_bg"])
            badge_row.pack(fill=tk.X, pady=(0, 2))
            
            cat_tag = tk.Label(badge_row, text="包装", font=("Microsoft YaHei", 8, "bold"), padx=5, pady=1)
            cat_tag.pack(side=tk.LEFT, padx=(0, 4))
            
            brand_tag = tk.Label(badge_row, text="", font=("Microsoft YaHei", 8), padx=4, pady=1)
            brand_tag.pack(side=tk.LEFT)
            
            title_lbl = tk.Label(meta_frame, text="", font=("Microsoft YaHei", 9, "bold"), bg=c["card_bg"], fg=c["fg"], wraplength=180, justify="left")
            title_lbl.pack(anchor="w")
            
            action_frame = tk.Frame(card, bg=c["card_bg"], pady=4)
            action_frame.pack(fill=tk.X)
            
            btn_open = tk.Button(action_frame, text="📁 文件夹", font=("Microsoft YaHei", 8), bg=c["btn_secondary_bg"], fg=c["btn_secondary_fg"], relief=tk.FLAT, bd=0)
            btn_open.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
            
            btn_blend = tk.Button(action_frame, text="🚀 3D工程", font=("Microsoft YaHei", 8, "bold"), bg=c["primary"], fg="#FFFFFF", relief=tk.FLAT, bd=0)
            btn_blend.pack(side=tk.RIGHT)
            
            slot = {
                "card": card,
                "img_lbl": img_lbl,
                "meta_frame": meta_frame,
                "badge_row": badge_row,
                "cat_tag": cat_tag,
                "brand_tag": brand_tag,
                "title_lbl": title_lbl,
                "action_frame": action_frame,
                "btn_open": btn_open,
                "btn_blend": btn_blend,
                "active_proj": None
            }
            self.card_slots.append(slot)

    # ---------------- 页面 2: 设计源文件分拣与开工 ----------------
    def build_organizer_ui(self, parent):
        c = self.colors
        
        top_frame = ttk.LabelFrame(parent, text=" 📂 工作盘、客户、分类与归档文件夹规则 ", padding=10)
        top_frame.pack(fill=tk.X, padx=16, pady=8)
        
        # 行 1: 主工作盘
        row_dir = ttk.Frame(top_frame)
        row_dir.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(row_dir, text="主工作盘:").pack(side=tk.LEFT, padx=(0, 6))
        self.ws_combo_org = ttk.Combobox(
            row_dir,
            textvariable=self.current_workspace_var,
            values=self.workspaces,
            state="readonly",
            width=36,
            font=("Microsoft YaHei", 9)
        )
        self.ws_combo_org.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(row_dir, text="➕ 绑定新工作盘...", command=self.add_workspace).pack(side=tk.LEFT)
        
        # 行 2: 客户、形态与【自定义文件夹规则】
        row_brand = ttk.Frame(top_frame)
        row_brand.pack(fill=tk.X)
        ttk.Label(row_brand, text="指定客户:").pack(side=tk.LEFT, padx=(0, 6))
        self.brand_combo_org = ttk.Combobox(
            row_brand,
            textvariable=self.current_brand_var,
            values=self.curated_brands,
            state="readonly",
            width=14,
            font=("Microsoft YaHei", 9)
        )
        self.brand_combo_org.pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(row_brand, text="➕", width=3, command=self.add_brand).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(row_brand, text="🏷️ 业务形态:").pack(side=tk.LEFT, padx=(4, 4))
        self.cat_combo_org = ttk.Combobox(
            row_brand,
            textvariable=self.current_cat_var,
            values=VALID_CATEGORIES,
            state="readonly",
            width=8,
            font=("Microsoft YaHei", 9)
        )
        self.cat_combo_org.pack(side=tk.LEFT, padx=(0, 10))
        self.cat_combo_org.bind("<<ComboboxSelected>>", lambda e: self.save_cfg_all())

        # 🌟 自定义文件夹规则选择器与管理入口
        ttk.Label(row_brand, text="📁 归档规则:").pack(side=tk.LEFT, padx=(4, 4))
        self.combo_rule = ttk.Combobox(
            row_brand,
            textvariable=self.active_rule_name_var,
            values=[r["name"] for r in self.folder_rules],
            state="readonly",
            width=24,
            font=("Microsoft YaHei", 9)
        )
        self.combo_rule.pack(side=tk.LEFT, padx=(0, 6))
        self.combo_rule.bind("<<ComboboxSelected>>", self.on_rule_combo_selected)
        
        ttk.Button(row_brand, text="⚙️ 自定义规则...", command=self.open_rule_manager).pack(side=tk.LEFT)

        # 2. 自动化存盘与同步选项
        b_frame = ttk.LabelFrame(parent, text=" ⚡ 自动化与 Excel 双向同步设置 ", padding=8)
        b_frame.pack(fill=tk.X, padx=16, pady=(0, 6))
        row_b = ttk.Frame(b_frame)
        row_b.pack(fill=tk.X)
        ttk.Checkbutton(row_b, text="✨ 自动生成对应 .blend 工程", variable=self.auto_create_blend_var, command=self.save_cfg_all).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(row_b, text="🚀 自动启动 Blender 打开工程", variable=self.auto_open_blender_var, command=self.save_cfg_all).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(row_b, text="📊 归档时自动录入《产品列表.xlsx》", variable=self.auto_append_excel_var, command=self.save_cfg_all).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(row_b, text="📁 设置母版 .blend...", command=self.set_custom_template).pack(side=tk.RIGHT)

        # 3. 待处理列表
        list_frame = ttk.LabelFrame(parent, text=" 📋 待处理的设计源文件 (按当前规则自动映射路径，双击可自由微调) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 6))
        
        cols = ("file", "brand", "sku", "cat", "target_dir", "status")
        self.tree_org = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        self.tree_org.heading("file", text="待归档文件名")
        self.tree_org.heading("brand", text="归属客户")
        self.tree_org.heading("sku", text="核心SKU名")
        self.tree_org.heading("cat", text="业务分类")
        self.tree_org.heading("target_dir", text="目标归档目录")
        self.tree_org.heading("status", text="状态")
        
        self.tree_org.column("file", width=210, anchor="w")
        self.tree_org.column("brand", width=100, anchor="center")
        self.tree_org.column("sku", width=160, anchor="w")
        self.tree_org.column("cat", width=90, anchor="center")
        self.tree_org.column("target_dir", width=200, anchor="w")
        self.tree_org.column("status", width=80, anchor="center")
        
        scroll_org = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_org.yview)
        self.tree_org.configure(yscrollcommand=scroll_org.set)
        self.tree_org.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_org.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_org.bind("<Double-1>", self.on_org_double_click)

        # 4. 操作按钮行
        btn_frame = ttk.Frame(parent, padding=2)
        btn_frame.pack(fill=tk.X, padx=16, pady=(0, 6))
        ttk.Button(btn_frame, text="➕ 添加 AI / 设计文件...", command=self.browse_files_for_org).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="✏️ 批量应用当前客户与分类", command=self.apply_current_brand_and_cat_to_all_org).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="❌ 移除选中", command=self.remove_selected_org).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="🧹 清空", command=self.clear_org).pack(side=tk.LEFT)

        # 5. 底部执行大按钮
        exec_frame = ttk.Frame(parent, padding=6)
        exec_frame.pack(fill=tk.X, padx=16, pady=(0, 10))
        btn_exec = tk.Button(
            exec_frame,
            text="🚀 【 按当前文件夹规则一键创建项目、同步写入 Excel 并自动开工 】",
            font=("Microsoft YaHei", 11, "bold"),
            bg=c["primary"],
            fg="#FFFFFF",
            activebackground=c["primary_hover"],
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            height=2,
            bd=0,
            command=self.execute_organize_flow
        )
        btn_exec.pack(fill=tk.X)

    # ---------------- 规则联动与管理 ----------------
    def on_rule_combo_selected(self, event=None):
        name = self.active_rule_name_var.get()
        for r in self.folder_rules:
            if r.get("name") == name:
                self.active_rule_id = r.get("id")
                break
        self.save_cfg_all()
        self.refresh_organizer_table()

    def open_rule_manager(self):
        FolderRuleManagerDialog(
            parent=self.root,
            rules=self.folder_rules,
            active_rule_id=self.active_rule_id,
            on_save_callback=self.on_rules_updated,
            colors=self.colors
        )

    def on_rules_updated(self, new_rules, new_active_id):
        self.folder_rules = new_rules
        self.active_rule_id = new_active_id
        self.cfg["folder_rules"] = self.folder_rules
        self.cfg["active_rule_id"] = self.active_rule_id
        save_config(self.cfg)
        
        self.combo_rule["values"] = [r["name"] for r in self.folder_rules]
        self.update_active_rule_name_var()
        self.refresh_organizer_table()
        messagebox.showinfo("规则已更新", f"已成功切换并应用规则:\n【{self.active_rule_name_var.get()}】")

    def compute_target_relative_dir(self, brand, sku, cat):
        rule = self.get_current_folder_rule()
        pat = rule.get("path_pattern", "{brand}/{sku}")
        if "{category}" in pat:
            rel = f"{cat}/{brand}/{sku}" if brand else f"{cat}/{sku}"
        elif "{brand}" in pat:
            rel = f"{brand}/{sku}" if brand else sku
        else:
            rel = sku
        return rel

    # ---------------- 逻辑与事件处理 ----------------
    def save_cfg_all(self):
        self.cfg["current_workspace"] = self.current_workspace_var.get()
        self.cfg["current_brand"] = self.current_brand_var.get()
        self.cfg["default_category"] = self.current_cat_var.get()
        self.cfg["auto_create_blend"] = self.auto_create_blend_var.get()
        self.cfg["auto_open_blender"] = self.auto_open_blender_var.get()
        self.cfg["auto_append_to_excel"] = self.auto_append_excel_var.get()
        self.cfg["folder_rules"] = self.folder_rules
        self.cfg["active_rule_id"] = self.active_rule_id
        save_config(self.cfg)

    def add_workspace(self):
        d = filedialog.askdirectory(title="选择需要绑定的新工作盘/根目录")
        if d:
            norm_d = os.path.normpath(d)
            if norm_d not in self.workspaces:
                self.workspaces.insert(0, norm_d)
                self.cfg["workspaces"] = self.workspaces
            self.cfg["current_workspace"] = norm_d
            self.current_workspace_var.set(norm_d)
            save_config(self.cfg)
            self.ws_combo_org["values"] = self.workspaces
            self.load_all_asset_data()

    def add_brand(self):
        name = simpledialog.askstring("添加新客户", "请输入客户/品牌名称 (如：统一、农夫山泉):", parent=self.root)
        if name and name.strip():
            name = name.strip()
            if name not in self.curated_brands:
                self.curated_brands.insert(0, name)
                self.cfg["curated_brands"] = self.curated_brands
                self.cfg["current_brand"] = name
                save_config(self.cfg)
                self.current_brand_var.set(name)
                self.brand_combo_org["values"] = self.curated_brands
                cur_ws = self.current_workspace_var.get().strip()
                if cur_ws and os.path.exists(cur_ws):
                    os.makedirs(os.path.join(cur_ws, name), exist_ok=True)

    def set_custom_template(self):
        f = filedialog.askopenfilename(title="选择你的默认包装 Blender 母版工程 (.blend)", filetypes=[("Blender 工程", "*.blend"), ("所有文件", "*.*")])
        if f:
            self.cfg["template_blend_path"] = f
            save_config(self.cfg)
            messagebox.showinfo("设置成功", f"已设为默认母版:\n{os.path.basename(f)}")

    def add_files_to_organizer(self, filepaths):
        cur_brand = self.current_brand_var.get().strip()
        for fp in filepaths:
            fp = os.path.abspath(fp)
            if not os.path.exists(fp) or not os.path.isfile(fp):
                continue
            if any(item['filepath'] == fp for item in self.files_to_organize):
                continue
            brand, sku, is_junk = clean_and_parse_filename(fp, fallback_brand=cur_brand, valid_brands=self.curated_brands)
            detected_cat = auto_detect_category_from_name(os.path.basename(fp))
            item = {
                "filepath": fp,
                "filename": os.path.basename(fp),
                "brand": brand,
                "sku": sku,
                "cat": detected_cat,
                "is_junk": is_junk,
                "md5": get_file_md5(fp)
            }
            self.files_to_organize.append(item)
        self.refresh_organizer_table()

    def refresh_organizer_table(self):
        self.tree_org.delete(*self.tree_org.get_children())
        for item in self.files_to_organize:
            brand = item["brand"]
            sku = item["sku"]
            cat = item.get("cat", "包装")
            target_rel = self.compute_target_relative_dir(brand, sku, cat)
            status = "⚠️需确认" if item["is_junk"] else "✅就绪"
            self.tree_org.insert("", tk.END, values=(item["filename"], brand, sku, cat, target_rel, status))

    def browse_files_for_org(self):
        files = filedialog.askopenfilenames(title="选择设计源文件", filetypes=[("包装设计文件", "*.ai;*.pdf;*.psd;*.zip;*.rar;*.eps"), ("所有文件", "*.*")])
        if files:
            self.add_files_to_organizer(files)

    def apply_current_brand_and_cat_to_all_org(self):
        cur_brand = self.current_brand_var.get().strip()
        cur_cat = self.current_cat_var.get().strip()
        for item in self.files_to_organize:
            if cur_brand:
                item["brand"] = cur_brand
            if cur_cat:
                item["cat"] = cur_cat
        self.refresh_organizer_table()

    def remove_selected_org(self):
        selected = self.tree_org.selection()
        for s in reversed(selected):
            idx = self.tree_org.index(s)
            self.tree_org.delete(s)
            del self.files_to_organize[idx]

    def clear_org(self):
        self.tree_org.delete(*self.tree_org.get_children())
        self.files_to_organize.clear()

    def on_org_double_click(self, event):
        item_id = self.tree_org.focus()
        if not item_id:
            return
        idx = self.tree_org.index(item_id)
        cur_item = self.files_to_organize[idx]
        
        edit_win = tk.Toplevel(self.root)
        edit_win.title("✏️ 快速修改客户、分类与项目名")
        edit_win.geometry("430x290")
        edit_win.configure(bg=self.colors["bg"])
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text=f"原始文件: {cur_item['filename']}", wraplength=390).pack(padx=15, pady=10, anchor="w")
        b_var = tk.StringVar(value=cur_item["brand"])
        s_var = tk.StringVar(value=cur_item["sku"])
        c_var = tk.StringVar(value=cur_item.get("cat", "包装"))
        
        f_in = ttk.Frame(edit_win)
        f_in.pack(fill=tk.X, padx=15, pady=5)
        ttk.Label(f_in, text="归属客户:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(f_in, textvariable=b_var, values=self.curated_brands, width=24).grid(row=0, column=1, sticky="w", pady=4)
        
        ttk.Label(f_in, text="业务分类:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(f_in, textvariable=c_var, values=VALID_CATEGORIES, width=24).grid(row=1, column=1, sticky="w", pady=4)
        
        ttk.Label(f_in, text="核心SKU名:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(f_in, textvariable=s_var, width=26).grid(row=2, column=1, sticky="w", pady=4)
        
        def save_edit():
            cur_item["brand"] = b_var.get().strip()
            cur_item["sku"] = s_var.get().strip()
            cur_item["cat"] = normalize_category(c_var.get().strip())
            cur_item["is_junk"] = False
            self.refresh_organizer_table()
            edit_win.destroy()
            
        ttk.Button(edit_win, text="保存修改 (Enter)", style="Primary.TButton", command=save_edit).pack(pady=12)
        edit_win.bind("<Return>", lambda e: save_edit())

    def execute_organize_flow(self):
        if not self.files_to_organize:
            messagebox.showwarning("提示", "请先添加需要归档的设计源文件！")
            return
            
        root_dir = self.current_workspace_var.get().strip()
        if not root_dir or not os.path.exists(root_dir):
            try:
                os.makedirs(root_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建主工作盘: {root_dir}\n{e}")
                return
                
        self.save_cfg_all()
        rule = self.get_current_folder_rule()
        subfolders = rule.get("subfolders", ["01_Design_平面原稿", "02_Textures_贴图资产", "03_3D_三维工程", "04_Renders_通道输出", "05_Delivery_最终交付"])
        design_sub = rule.get("design_sub", subfolders[0] if subfolders else "")
        blend_sub = rule.get("blend_sub", subfolders[2] if len(subfolders) > 2 else (subfolders[0] if subfolders else ""))
        
        template_blend = self.cfg.get("template_blend_path", "")
        if not template_blend or not os.path.exists(template_blend):
            template_blend = DEFAULT_TEMPLATE
            
        auto_create_blend = self.auto_create_blend_var.get()
        auto_open_blender = self.auto_open_blender_var.get()
        auto_append_excel = self.auto_append_excel_var.get()
        excel_path = self.excel_path_var.get().strip()
        
        success_count = 0
        duplicate_count = 0
        excel_appended_count = 0
        last_blend = ""
        last_proj = ""
        
        for item in self.files_to_organize:
            brand = item["brand"].strip()
            sku = item["sku"].strip() if item["sku"].strip() else os.path.splitext(item["filename"])[0]
            cat = normalize_category(item.get("cat", "包装"))
            
            target_rel = self.compute_target_relative_dir(brand, sku, cat)
            proj_dir = os.path.join(root_dir, target_rel.replace("/", "\\"))
            
            # 按照用户规则创建所有子文件夹
            for sub in subfolders:
                os.makedirs(os.path.join(proj_dir, sub), exist_ok=True)
                
            design_dir = os.path.join(proj_dir, design_sub) if design_sub else proj_dir
            os.makedirs(design_dir, exist_ok=True)
            
            ext = os.path.splitext(item["filename"])[1]
            existing_files = os.listdir(design_dir) if os.path.exists(design_dir) else []
            is_dup = False
            if item.get("md5"):
                for ef in existing_files:
                    if get_file_md5(os.path.join(design_dir, ef)) == item["md5"]:
                        is_dup = True
                        break
            if is_dup:
                duplicate_count += 1
                last_proj = proj_dir
                continue
                
            dest_name = f"{sku}_v{len(existing_files)+1:02d}{ext}"
            try:
                shutil.copy2(item["filepath"], os.path.join(design_dir, dest_name))
                success_count += 1
                last_proj = proj_dir
            except Exception as e:
                print(f"Error copying: {e}")
                
            if auto_create_blend and blend_sub:
                target_blend_dir = os.path.join(proj_dir, blend_sub)
                os.makedirs(target_blend_dir, exist_ok=True)
                target_blend = os.path.join(target_blend_dir, f"{sku}.blend")
                if not os.path.exists(target_blend) and template_blend and os.path.exists(template_blend):
                    try:
                        shutil.copy2(template_blend, target_blend)
                        last_blend = target_blend
                    except Exception:
                        pass
                elif os.path.exists(target_blend):
                    last_blend = target_blend

            if auto_append_excel and excel_path and os.path.exists(excel_path):
                if append_project_to_excel(excel_path, brand, sku, cat, proj_dir):
                    excel_appended_count += 1

        msg = []
        if success_count > 0:
            msg.append(f"✅ 成功按【{rule.get('name')}】创建并归档 {success_count} 个标准项目！")
            if excel_appended_count > 0:
                msg.append(f"📊 自动同步将 {excel_appended_count} 个新项目录入《产品列表.xlsx》！")
        if duplicate_count > 0:
            msg.append(f"ℹ️ 自动跳过 {duplicate_count} 个重复接收文件。")
            
        if auto_open_blender and last_blend and os.path.exists(last_blend):
            try:
                subprocess.Popen([BLENDER_EXE, last_blend])
                msg.append(f"🚀 已自动启动 Blender 打开工程: [{os.path.basename(last_blend)}]")
            except Exception:
                os.startfile(last_blend)
        elif last_proj and os.path.exists(last_proj):
            try:
                os.startfile(last_proj)
            except Exception:
                pass
                
        messagebox.showinfo("🎉 处理完成", "\n".join(msg))
        self.clear_org()
        self.load_all_asset_data()

    # ---------------- 资产看板数据与极速同步 ----------------
    def start_excel_auto_sync_watcher(self):
        ex_path = self.excel_path_var.get().strip()
        if ex_path and os.path.exists(ex_path):
            try:
                current_mtime = os.path.getmtime(ex_path)
                if self.last_excel_mtime > 0 and current_mtime > self.last_excel_mtime:
                    self.last_excel_mtime = current_mtime
                    self.thumb_tk_cache.clear()
                    self.load_all_asset_data()
                    self.sync_status_lbl.config(text="⚡ Excel 已更新，已自动同步！", bg="#FEF3C7", fg="#B45309")
                    self.root.after(3500, lambda: self.sync_status_lbl.config(text="🟢 极速同步已就绪", bg=self.colors["status_bg"], fg=self.colors["status_fg"]))
                elif self.last_excel_mtime == 0:
                    self.last_excel_mtime = current_mtime
            except Exception:
                pass
        self.root.after(2000, self.start_excel_auto_sync_watcher)

    def load_all_asset_data(self):
        ex_path = self.excel_path_var.get().strip()
        if os.path.exists(ex_path):
            self.last_excel_mtime = os.path.getmtime(ex_path)
            
        self.excel_projects = parse_and_cache_excel(ex_path)
        cur_ws = self.current_workspace_var.get().strip()
        
        self.disk_projects = scan_workspace_projects_fast(cur_ws, self.meta_cache)
        self.merged_projects = merge_excel_and_disk_projects(self.excel_projects, self.disk_projects)
        
        self.combo_source["values"] = [
            f"⚡ 智能融合视图 ({len(self.merged_projects)})",
            f"仅 Excel 产品台账 ({len(self.excel_projects)})",
            f"仅本地工作盘扫描 ({len(self.disk_projects)})"
        ]
        
        mode = self.view_mode_var.get()
        if mode == "excel":
            self.combo_source.current(1)
        elif mode == "disk":
            self.combo_source.current(2)
        else:
            self.combo_source.current(0)
            
        self.update_active_dataset()

    def on_source_change(self, event=None):
        sel_idx = self.combo_source.current()
        if sel_idx == 1:
            self.view_mode_var.set("excel")
        elif sel_idx == 2:
            self.view_mode_var.set("disk")
        else:
            self.view_mode_var.set("merged")
        self.update_active_dataset()

    def update_active_dataset(self):
        mode = self.view_mode_var.get()
        if mode == "excel":
            self.current_display_list = self.excel_projects
        elif mode == "disk":
            self.current_display_list = self.disk_projects
        else:
            self.current_display_list = self.merged_projects
            
        cat_counts = {c: 0 for c in VALID_CATEGORIES}
        for p in self.current_display_list:
            c = normalize_category(p.get("cat", "包装"))
            p["cat"] = c
            cat_counts[c] = cat_counts.get(c, 0) + 1
            
        self.category_listbox.delete(0, tk.END)
        self.category_listbox.insert(tk.END, f"✨ 全部形态 ({len(self.current_display_list)})")
        
        cat_icons = {"包装": "📦", "套盒": "🎁", "海报": "🖼️", "物料": "📑"}
        for c in VALID_CATEGORIES:
            icon = cat_icons.get(c, "🏷️")
            self.category_listbox.insert(tk.END, f"{icon} {c} ({cat_counts.get(c, 0)})")
            
        self.category_listbox.select_set(0)
        self.selected_category_var.set("全部")
        self.current_page = 0
        self.apply_filter()

    def import_new_excel(self):
        f = filedialog.askopenfilename(title="选择包装列表 Excel 表格 (.xlsx)", filetypes=[("Excel 表格", "*.xlsx"), ("所有文件", "*.*")])
        if f:
            self.excel_path_var.set(f)
            self.cfg["excel_path"] = f
            save_config(self.cfg)
            self.load_all_asset_data()

    def on_search_change_debounced(self):
        if self.search_debounce_job:
            self.root.after_cancel(self.search_debounce_job)
        self.search_debounce_job = self.root.after(200, self.do_search_apply)

    def do_search_apply(self):
        self.current_page = 0
        self.apply_filter()

    def on_category_select(self, event=None):
        sel = self.category_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        text = self.category_listbox.get(idx)
        if idx == 0:
            self.selected_category_var.set("全部")
        else:
            m = re.search(r"[\u4e00-\u9fa5]+", text)
            if m:
                clean_name = m.group(0)
                if clean_name in VALID_CATEGORIES:
                    self.selected_category_var.set(clean_name)
        self.current_page = 0
        self.apply_filter()

    def apply_filter(self):
        cat = self.selected_category_var.get()
        kw = self.search_var.get().strip().lower()
        
        self.filtered_projects = []
        for p in self.current_display_list:
            p_cat = normalize_category(p.get("cat", "包装"))
            if cat != "全部" and p_cat != cat:
                continue
            if kw and (kw not in p["sku"].lower() and kw not in p.get("brand", "").lower() and kw not in p_cat.lower()):
                continue
            self.filtered_projects.append(p)
            
        total_items = len(self.filtered_projects)
        total_pages = max(1, (total_items + self.page_size - 1) // self.page_size)
        if self.current_page >= total_pages:
            self.current_page = total_pages - 1
            
        start_idx = self.current_page * self.page_size + 1 if total_items > 0 else 0
        end_idx = min((self.current_page + 1) * self.page_size, total_items)
        
        self.page_info_lbl.config(text=f"共 {total_items} 个项目 | 正在显示第 {start_idx} - {end_idx} 项 (第 {self.current_page + 1}/{total_pages} 页)")
        self.btn_prev.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
        self.btn_next.config(state=tk.NORMAL if self.current_page < total_pages - 1 else tk.DISABLED)
        
        self.render_cards()

    def next_page(self):
        self.current_page += 1
        self.apply_filter()
        self.canvas.yview_moveto(0)

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.apply_filter()
            self.canvas.yview_moveto(0)

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.render_cards()

    def on_mouse_wheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def get_scaled_thumbnail(self, img_path, size=(190, 190)):
        if not img_path or not os.path.exists(img_path):
            return self.get_placeholder_thumbnail(size)
            
        cache_key = f"{img_path}_{self.current_theme}"
        if cache_key in self.thumb_tk_cache:
            return self.thumb_tk_cache[cache_key]
            
        try:
            fast_thumb_p = get_fast_disk_thumbnail_path(img_path, size)
            if not fast_thumb_p or not os.path.exists(fast_thumb_p):
                return self.get_placeholder_thumbnail(size)
                
            im = Image.open(fast_thumb_p)
            tk_img = ImageTk.PhotoImage(im)
            self.thumb_tk_cache[cache_key] = tk_img
            return tk_img
        except Exception:
            return self.get_placeholder_thumbnail(size)

    def get_placeholder_thumbnail(self, size=(190, 190)):
        cache_key = f"placeholder_{self.current_theme}"
        if cache_key in self.thumb_tk_cache:
            return self.thumb_tk_cache[cache_key]
        c = self.colors
        bg_c = (39, 41, 47, 255) if self.current_theme == "dark" else (244, 246, 249, 255)
        fg_c = (110, 114, 122, 255) if self.current_theme == "dark" else (148, 163, 184, 255)
        im = Image.new("RGBA", size, bg_c)
        draw = ImageDraw.Draw(im)
        draw.text((size[0]//2 - 40, size[1]//2 - 10), "📦 待渲染工程", fill=fg_c)
        tk_img = ImageTk.PhotoImage(im)
        self.thumb_tk_cache[cache_key] = tk_img
        return tk_img

    def render_cards(self):
        c = self.colors
        container_width = self.canvas.winfo_width()
        if container_width < 100:
            container_width = 800
        card_w = 210
        cols = max(1, container_width // (card_w + 16))

        start_idx = self.current_page * self.page_size
        end_idx = start_idx + self.page_size
        page_items = self.filtered_projects[start_idx:end_idx]
        total_page_items = len(page_items)

        for idx in range(len(self.card_slots)):
            slot = self.card_slots[idx]
            if idx < total_page_items:
                proj = page_items[idx]
                slot["active_proj"] = proj
                row = idx // cols
                col = idx % cols
                
                slot["card"].grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
                
                tk_thumb = self.get_scaled_thumbnail(proj["thumbnail"])
                slot["img_lbl"].config(image=tk_thumb)
                slot["img_lbl"].image = tk_thumb
                
                cat_val = normalize_category(proj.get("cat", "包装"))
                bg_c, fg_c = c["cat_colors"].get(cat_val, ("#26384C", "#96C2EC"))
                slot["cat_tag"].config(text=cat_val, bg=bg_c, fg=fg_c)
                
                b_name = proj.get("brand", "")
                if b_name:
                    slot["brand_tag"].config(text=b_name, bg=c["badge_brand_bg"], fg=c["badge_brand_fg"])
                    slot["brand_tag"].pack(side=tk.LEFT)
                else:
                    slot["brand_tag"].pack_forget()
                    
                slot["title_lbl"].config(text=proj["sku"])
                
                has_path = bool(proj.get("path") and os.path.exists(proj["path"]))
                p = proj.get("path")
                slot["btn_open"].config(
                    text="📁 文件夹" if has_path else "📁 未就绪",
                    fg=c["btn_secondary_fg"] if has_path else c["fg_dim"],
                    command=lambda p_path=p: self.open_folder(p_path)
                )
                slot["btn_blend"].config(command=lambda p_path=p: self.launch_blend(p_path))
                
                for w in (slot["card"], slot["img_lbl"], slot["title_lbl"], slot["meta_frame"]):
                    w.bind("<Button-1>", lambda e, p_path=p: self.open_folder(p_path))
                    w.bind("<Double-1>", lambda e, p_path=p: self.launch_blend(p_path))
                    w.bind("<Button-3>", lambda e, pr=proj: self.show_context_menu(e, pr))
            else:
                slot["active_proj"] = None
                slot["card"].grid_remove()

    def open_folder(self, path):
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("打开错误", str(e))
        else:
            messagebox.showwarning("提示", f"该项目的本地文件夹暂未找到或路径不存在:\n{path}")

    def launch_blend(self, proj_path):
        if not proj_path or not os.path.exists(proj_path):
            messagebox.showwarning("提示", f"未找到该项目的本地文件夹:\n{proj_path}")
            return
        rule = self.get_current_folder_rule()
        blend_sub = rule.get("blend_sub", "03_3D_三维工程")
        
        blend_dir = os.path.join(proj_path, blend_sub) if blend_sub else proj_path
        target_dir = blend_dir if os.path.exists(blend_dir) else proj_path
        blends = glob.glob(os.path.join(target_dir, "*.blend"))
        if blends:
            blends.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            chosen = blends[0]
            try:
                subprocess.Popen([BLENDER_EXE, chosen])
            except Exception:
                os.startfile(chosen)
        else:
            self.open_folder(proj_path)

    def show_context_menu(self, event, proj):
        menu = tk.Menu(self.root, tearoff=0, bg=self.colors["panel_bg"], fg=self.colors["fg"])
        p = proj.get("path", "")
        sku = proj.get("sku", "")
        rule = self.get_current_folder_rule()
        d_sub = rule.get("design_sub", "01_Design_平面原稿")
        r_sub = rule.get("render_sub", "04_Renders_通道输出")
        
        menu.add_command(label=f"📁 打开文件夹: {sku}", command=lambda: self.open_folder(p))
        menu.add_command(label="🚀 Blender 打开 3D 工程", command=lambda: self.launch_blend(p))
        if p and os.path.exists(p):
            if d_sub:
                menu.add_command(label=f"🎨 查看 {d_sub}", command=lambda: self.open_folder(os.path.join(p, d_sub)))
            if r_sub:
                menu.add_command(label=f"🖼️ 查看 {r_sub}", command=lambda: self.open_folder(os.path.join(p, r_sub)))
            
        menu.add_separator()
        
        cat_submenu = tk.Menu(menu, tearoff=0, bg=self.colors["panel_bg"], fg=self.colors["fg"])
        for cat_item in VALID_CATEGORIES:
            cat_submenu.add_command(
                label=f"设为：{cat_item}",
                command=lambda c=cat_item, pr=proj: self.change_project_category(pr, c)
            )
        menu.add_cascade(label=f"🏷️ 修改业务形态 (当前: {proj.get('cat', '包装')})", menu=cat_submenu)
        
        menu.add_separator()
        menu.add_command(label="📋 复制完整物理路径", command=lambda: self.copy_path_to_clipboard(p))
        menu.tk_popup(event.x_root, event.y_root)

    def change_project_category(self, proj, new_cat):
        new_cat = normalize_category(new_cat)
        proj["cat"] = new_cat
        ex_path = self.excel_path_var.get().strip()
        
        update_project_category_in_excel(ex_path, proj.get("path"), proj.get("sku"), new_cat)
        
        if proj.get("path"):
            cache_key = proj["path"].lower().replace("/", "\\")
            if cache_key in self.meta_cache:
                self.meta_cache[cache_key]["cat"] = new_cat
                save_meta_cache(self.meta_cache)
                
        self.update_active_dataset()
        self.sync_status_lbl.config(text=f"✅ [{proj['sku']}] 已更新为 【{new_cat}】！", bg="#1D3328", fg="#6EE7B7")
        self.root.after(3500, lambda: self.sync_status_lbl.config(text="🟢 极速同步已就绪", bg=self.colors["status_bg"], fg=self.colors["status_fg"]))

    def copy_path_to_clipboard(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制路径到剪贴板:\n{text}")

    def export_html_gallery(self):
        cur_ws = self.current_workspace_var.get().strip()
        html_file = os.path.join(cur_ws if os.path.exists(cur_ws) else os.path.expanduser("~"), "🎨_美术资产全景视觉画廊.html")
        cards_html = []
        for p in self.current_display_list:
            thumb_rel = p.get("thumbnail") or ""
            thumb_src = "file:///" + thumb_rel.replace("\\", "/") if (thumb_rel and os.path.exists(thumb_rel)) else "https://via.placeholder.com/300x300?text=No+Render"
            folder_uri = "file:///" + p["path"].replace("\\", "/") if p.get("path") else "#"
            cat_name = normalize_category(p.get("cat", "包装"))
            
            cards_html.append(f"""
            <div class="card" onclick="window.open('{folder_uri}')">
                <div class="thumb-container"><img src="{thumb_src}" alt="{p['sku']}" loading="lazy"></div>
                <div class="meta">
                    <div style="display:flex; gap:6px; margin-bottom:6px;">
                        <span class="badge" style="background:#26384C; color:#96C2EC;">{cat_name}</span>
                        <span class="badge" style="background:#3A3D46; color:#B2B6BE;">{p.get('brand', '')}</span>
                    </div>
                    <h3 class="title">{p['sku']}</h3>
                    <p class="path">{p.get('path', '')}</p>
                </div>
            </div>
            """)
        full_html = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>🎨 美术资产全景视觉画廊</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: #222429; color: #e2e4e8; margin: 0; padding: 24px; }}
h1 {{ font-size: 24px; font-weight: 700; margin-bottom: 24px; color: #e2e4e8; }}
.grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 20px; }}
.card {{ background: #30333A; border-radius: 10px; overflow: hidden; border: 1px solid #3d414a; cursor: pointer; transition: transform 0.2s, border-color 0.2s; }}
.card:hover {{ transform: translateY(-4px); border-color: #4c84c4; box-shadow: 0 12px 24px -10px rgba(0,0,0,0.5); }}
.thumb-container {{ width: 100%; aspect-ratio: 1; background: #27292f; display: flex; align-items: center; justify-content: center; }}
.thumb-container img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
.meta {{ padding: 12px; }}
.badge {{ display: inline-block; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 9999px; }}
.title {{ font-size: 14px; font-weight: 600; margin: 0 0 4px 0; color: #e2e4e8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.path {{ font-size: 11px; color: #9699a2; margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
</style></head><body><h1>🎨 美术资产全景视觉画廊 (共 {len(self.current_display_list)} 个项目)</h1><div class="grid">{"".join(cards_html)}</div></body></html>"""
        try:
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(full_html)
            webbrowser.open("file:///" + html_file.replace("\\", "/"))
        except Exception as e:
            messagebox.showerror("导出错误", str(e))


if __name__ == "__main__":
    args_files = sys.argv[1:] if len(sys.argv) > 1 else None
    root_win = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = PackagingStudioSuite(root_win, initial_files=args_files)
    root_win.mainloop()
