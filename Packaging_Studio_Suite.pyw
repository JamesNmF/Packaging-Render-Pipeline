# -*- coding: utf-8 -*-
"""
包装设计与资产综合中枢 (Packaging Studio Suite - v4.1 自动双向追加同步版)
新增核心特性：
1. 【新项目全自动追加写入 Excel】：
   - 微信收到文件一键归档时，自动将新产品名、分类(瓶装/袋装/盒装/套盒)、路径与时间追加写入《产品列表.xlsx》！
2. 【资产看板实时自动上架】：
   - Excel 追加写入后，资产管理器 0.1 秒自动感应，新卡片瞬间上架展示！
3. 【智能防重复检测】：
   - 自动检测已有产品，绝不在 Excel 中产生重复行。
"""

import os
import sys
import re
import json
import glob
import zipfile
import datetime
import webbrowser
import subprocess
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from PIL import Image, ImageTk, ImageDraw
import openpyxl

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_suite_v4.json")
CACHE_DIR = os.path.join(os.path.expanduser("~"), ".packaging_asset_thumbnails")
EXCEL_CACHE_DIR = os.path.join(CACHE_DIR, "excel_images")
os.makedirs(EXCEL_CACHE_DIR, exist_ok=True)

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

DEFAULT_CONFIG = {
    "workspaces": DEFAULT_WORKSPACES,
    "current_workspace": DEFAULT_WORKSPACES[0],
    "excel_path": DEFAULT_EXCEL_PATH if os.path.exists(DEFAULT_EXCEL_PATH) else "",
    "curated_brands": ["柏缇", "零食有鸣"],
    "current_brand": "柏缇",
    "auto_create_blend": True,
    "auto_open_blender": True,
    "template_blend_path": DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else "",
    "auto_append_to_excel": True
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in DEFAULT_CONFIG.items():
                    if k not in data:
                        data[k] = v
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


def append_project_to_excel(excel_path, brand, sku, cat, proj_path):
    """将新创建的项目自动追加写入到 Excel 产品台账中"""
    if not excel_path or not os.path.exists(excel_path):
        return False
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb['全部'] if '全部' in wb.sheetnames else wb.active
        
        # 查重检测
        norm_proj_p = proj_path.replace('\\', '/').lower().strip('/')
        for r in range(2, ws.max_row + 1):
            ex_p = str(ws.cell(row=r, column=5).value or "").replace('\\', '/').lower().strip('/')
            ex_name = str(ws.cell(row=r, column=2).value or "").strip()
            if ex_name == sku or (norm_proj_p and ex_p == norm_proj_p):
                return False # 已存在
                
        next_idx = ws.max_row
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        
        new_row = [next_idx, sku, None, cat, proj_path.replace('\\', '/'), now_str]
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


def get_file_md5(filepath):
    try:
        import hashlib
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
                sku = f"未命名产品_{raw_name}"
            else:
                brand = fallback_brand
                sku = parts[0]
        else:
            brand = fallback_brand
            sku = f"未命名产品_{raw_name}"
            
    return brand, sku, is_junk


def find_project_thumbnail(proj_path):
    render_dirs = [
        os.path.join(proj_path, "04_Renders_通道输出"),
        os.path.join(proj_path, "05_Delivery_最终交付"),
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
        beauty_imgs = [c for c in candidates if any(k in os.path.basename(c).lower() for k in ["beauty", "成品", "主图", "camera", "正面", "01_"])]
        if beauty_imgs:
            try:
                beauty_imgs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
                return beauty_imgs[0]
            except Exception:
                return beauty_imgs[0]
        filtered = [c for c in candidates if not any(k in os.path.basename(c).lower() for k in ["mask", "alpha", "crypto", "选区", "蒙版", "normal", "depth", "roughness"])]
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
            cat = sheet.cell(row=r, column=4).value or "未分类"
            path = sheet.cell(row=r, column=5).value or ""
            time_str = sheet.cell(row=r, column=6).value or ""

            if not name:
                continue

            thumb = row_image_map.get(r, None)
            norm_path = path.replace("/", "\\") if path else ""

            brand = "柏缇"
            if norm_path:
                parts = norm_path.strip("\\").split("\\")
                if len(parts) >= 2 and parts[-2] not in {"zjc", "Projects", "E:", "D:"}:
                    brand = parts[-2]

            projects.append({
                "source": "excel",
                "brand": brand,
                "sku": str(name).strip(),
                "cat": str(cat).strip(),
                "path": norm_path,
                "thumbnail": thumb,
                "time": str(time_str),
                "mtime": os.path.getmtime(norm_path) if (norm_path and os.path.exists(norm_path)) else 0
            })

    except Exception as e:
        print(f"Error parsing Excel: {e}")

    return projects


def scan_workspace_projects(root_dir):
    projects = []
    if not os.path.exists(root_dir):
        return projects
    try:
        entries = os.listdir(root_dir)
    except (PermissionError, OSError):
        return projects
        
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
                    thumb = find_project_thumbnail(sku_p)
                    try:
                        s_mtime = os.path.getmtime(sku_p)
                    except Exception:
                        s_mtime = 0
                    projects.append({
                        "source": "disk",
                        "brand": entry,
                        "sku": sku,
                        "cat": "三维项目",
                        "path": sku_p,
                        "thumbnail": thumb,
                        "time": "",
                        "mtime": s_mtime
                    })
            except (PermissionError, OSError):
                continue
                
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
                "cat": ep.get("cat", "三维项目"),
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


class PackagingStudioSuite:
    def __init__(self, root, initial_files=None):
        self.root = root
        self.root.title("📦 包装设计与视觉资产中枢 (Packaging Studio Suite v4.1)")
        self.root.geometry("1200x800")
        self.root.minsize(980, 640)
        
        self.cfg = load_config()
        self.workspaces = self.cfg.get("workspaces", DEFAULT_WORKSPACES)
        cur_ws = self.cfg.get("current_workspace", self.workspaces[0])
        if not os.path.exists(cur_ws) and self.workspaces:
            cur_ws = self.workspaces[0]
        self.current_workspace_var = tk.StringVar(value=cur_ws)
        self.excel_path_var = tk.StringVar(value=self.cfg.get("excel_path", DEFAULT_EXCEL_PATH))
        
        # 归档页变量
        self.curated_brands = self.cfg.get("curated_brands", ["柏缇", "零食有鸣"])
        self.current_brand_var = tk.StringVar(value=self.cfg.get("current_brand", self.curated_brands[0]))
        self.current_cat_var = tk.StringVar(value="瓶装") # 默认归档分类
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
        
        self.excel_projects = []
        self.disk_projects = []
        self.merged_projects = []
        self.current_display_list = []
        self.filtered_projects = []
        self.thumb_cache = {}
        
        self.build_ui()
        self.load_all_asset_data()
        self.start_excel_auto_sync_watcher()
        
        if initial_files:
            self.notebook.select(1)
            self.add_files_to_organizer(initial_files)
        else:
            self.notebook.select(0)

    def build_ui(self):
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Microsoft YaHei", 10, "bold"), padding=[16, 6])
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: 视觉资产看板
        self.tab_assets = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_assets, text="  🖼️ 视觉资产看板  ")
        self.build_asset_hub_ui(self.tab_assets)
        
        # Tab 2: 微信文件分拣归档
        self.tab_organizer = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_organizer, text="  📥 微信AI文件归档与开工  ")
        self.build_organizer_ui(self.tab_organizer)

    # ---------------- 页面 1: 视觉资产看板 ----------------
    def build_asset_hub_ui(self, parent):
        top_bar = ttk.Frame(parent, padding=(15, 10))
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
        self.combo_source.pack(side=tk.LEFT, padx=(0, 15))
        self.combo_source.bind("<<ComboboxSelected>>", self.on_source_change)
        
        ttk.Label(top_bar, text="🔍 搜索:").pack(side=tk.LEFT, padx=(0, 4))
        search_entry = ttk.Entry(top_bar, textvariable=self.search_var, font=("Microsoft YaHei", 9), width=18)
        search_entry.pack(side=tk.LEFT, padx=(0, 12))
        self.search_var.trace_add("write", lambda *args: self.on_search_change())
        
        ttk.Button(top_bar, text="📊 指定 Excel...", command=self.import_new_excel).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(top_bar, text="🔄 刷新", command=self.load_all_asset_data).pack(side=tk.LEFT, padx=(0, 6))
        
        self.sync_status_lbl = tk.Label(top_bar, text="🟢 扫描优先实时同步已就绪", font=("Microsoft YaHei", 8), fg="#059669", bg="#ECFDF5", padx=6, pady=2)
        self.sync_status_lbl.pack(side=tk.LEFT, padx=(6, 0))
        
        ttk.Button(top_bar, text="🌐 导出网页画廊 (HTML)...", command=self.export_html_gallery).pack(side=tk.RIGHT)

        # 底栏分页
        self.bottom_bar = ttk.Frame(parent, padding=(15, 8))
        self.bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.page_info_lbl = ttk.Label(self.bottom_bar, text="", font=("Microsoft YaHei", 9))
        self.page_info_lbl.pack(side=tk.LEFT)
        
        self.btn_next = ttk.Button(self.bottom_bar, text="下一页 ➡️", command=self.next_page)
        self.btn_next.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.btn_prev = ttk.Button(self.bottom_bar, text="⬅️ 上一页", command=self.prev_page)
        self.btn_prev.pack(side=tk.RIGHT)

        # 主视口
        main_pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 5))
        
        left_frame = ttk.LabelFrame(main_pane, text=" 🏷️ 形态分类 & 筛选 ", padding=8, width=200)
        main_pane.add(left_frame, weight=1)
        
        self.category_listbox = tk.Listbox(
            left_frame,
            font=("Microsoft YaHei", 10),
            selectmode=tk.SINGLE,
            relief=tk.FLAT,
            bg="#F8F9FA",
            selectbackground="#0078D7",
            selectforeground="white",
            highlightthickness=0,
            activestyle="none"
        )
        self.category_listbox.pack(fill=tk.BOTH, expand=True)
        self.category_listbox.bind("<<ListboxSelect>>", self.on_category_select)
        
        right_frame = ttk.Frame(main_pane)
        main_pane.add(right_frame, weight=5)
        
        self.canvas = tk.Canvas(right_frame, bg="#F0F2F5", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        
        self.grid_container = tk.Frame(self.canvas, bg="#F0F2F5")
        self.grid_container.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.grid_container, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)

    # ---------------- 页面 2: 微信文件分拣归档 ----------------
    def build_organizer_ui(self, parent):
        top_frame = ttk.LabelFrame(parent, text=" 📂 工作盘、客户与形态分类 ", padding=10)
        top_frame.pack(fill=tk.X, padx=15, pady=8)
        
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
        
        row_brand = ttk.Frame(top_frame)
        row_brand.pack(fill=tk.X)
        ttk.Label(row_brand, text="指定客户:").pack(side=tk.LEFT, padx=(0, 6))
        self.brand_combo_org = ttk.Combobox(
            row_brand,
            textvariable=self.current_brand_var,
            values=self.curated_brands,
            state="readonly",
            width=16,
            font=("Microsoft YaHei", 9)
        )
        self.brand_combo_org.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(row_brand, text="➕ 新增客户...", command=self.add_brand).pack(side=tk.LEFT, padx=(0, 6))
        
        ttk.Label(row_brand, text="包装形态:").pack(side=tk.LEFT, padx=(8, 4))
        self.cat_combo_org = ttk.Combobox(
            row_brand,
            textvariable=self.current_cat_var,
            values=["瓶装", "袋装", "盒装", "套盒", "软管", "罐装", "通用"],
            state="readonly",
            width=10,
            font=("Microsoft YaHei", 9)
        )
        self.cat_combo_org.pack(side=tk.LEFT, padx=(0, 8))

        # 2. 自动化存盘与同步选项
        b_frame = ttk.LabelFrame(parent, text=" ⚡ 自动化与 Excel 双向同步设置 ", padding=8)
        b_frame.pack(fill=tk.X, padx=15, pady=(0, 6))
        row_b = ttk.Frame(b_frame)
        row_b.pack(fill=tk.X)
        ttk.Checkbutton(row_b, text="✨ 自动生成对应 .blend 工程", variable=self.auto_create_blend_var, command=self.save_cfg_all).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(row_b, text="🚀 自动启动 Blender 打开工程", variable=self.auto_open_blender_var, command=self.save_cfg_all).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(row_b, text="📊 归档时自动将新产品追加录入《产品列表.xlsx》", variable=self.auto_append_excel_var, command=self.save_cfg_all).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Button(row_b, text="📁 设置母版 .blend...", command=self.set_custom_template).pack(side=tk.RIGHT)

        # 3. 待处理列表
        list_frame = ttk.LabelFrame(parent, text=" 📋 待处理的微信源文件 (自动剔除 (1)(2) 并智能查重，双击可编辑) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 6))
        
        cols = ("file", "brand", "sku", "cat", "target_dir", "status")
        self.tree_org = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        self.tree_org.heading("file", text="微信接收的文件名")
        self.tree_org.heading("brand", text="归属客户")
        self.tree_org.heading("sku", text="核心SKU名")
        self.tree_org.heading("cat", text="形态分类")
        self.tree_org.heading("target_dir", text="目标归档目录")
        self.tree_org.heading("status", text="状态")
        
        self.tree_org.column("file", width=200, anchor="w")
        self.tree_org.column("brand", width=100, anchor="center")
        self.tree_org.column("sku", width=160, anchor="w")
        self.tree_org.column("cat", width=80, anchor="center")
        self.tree_org.column("target_dir", width=180, anchor="w")
        self.tree_org.column("status", width=80, anchor="center")
        
        scroll_org = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree_org.yview)
        self.tree_org.configure(yscrollcommand=scroll_org.set)
        self.tree_org.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_org.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_org.bind("<Double-1>", self.on_org_double_click)

        # 4. 操作按钮行
        btn_frame = ttk.Frame(parent, padding=2)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 6))
        ttk.Button(btn_frame, text="➕ 添加 AI / 源文件...", command=self.browse_files_for_org).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="✏️ 批量应用当前客户与分类", command=self.apply_current_brand_to_all_org).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="❌ 移除选中", command=self.remove_selected_org).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="🧹 清空", command=self.clear_org).pack(side=tk.LEFT)

        # 5. 底部执行大按钮
        exec_frame = ttk.Frame(parent, padding=6)
        exec_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        btn_exec = tk.Button(
            exec_frame,
            text="🚀 【 一键创建工业级标准项目、同步写入 Excel 并自动开工 】",
            font=("Microsoft YaHei", 11, "bold"),
            bg="#0078D7",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            relief=tk.FLAT,
            height=2,
            command=self.execute_organize_flow
        )
        btn_exec.pack(fill=tk.X)

    # ---------------- 逻辑与事件处理 ----------------
    def save_cfg_all(self):
        self.cfg["current_workspace"] = self.current_workspace_var.get()
        self.cfg["current_brand"] = self.current_brand_var.get()
        self.cfg["auto_create_blend"] = self.auto_create_blend_var.get()
        self.cfg["auto_open_blender"] = self.auto_open_blender_var.get()
        self.cfg["auto_append_to_excel"] = self.auto_append_excel_var.get()
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
        cur_cat = self.current_cat_var.get().strip()
        for fp in filepaths:
            fp = os.path.abspath(fp)
            if not os.path.exists(fp) or not os.path.isfile(fp):
                continue
            if any(item['filepath'] == fp for item in self.files_to_organize):
                continue
            brand, sku, is_junk = clean_and_parse_filename(fp, fallback_brand=cur_brand, valid_brands=self.curated_brands)
            item = {
                "filepath": fp,
                "filename": os.path.basename(fp),
                "brand": brand,
                "sku": sku,
                "cat": cur_cat,
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
            cat = item.get("cat", "瓶装")
            target_proj = f"{brand}/{sku}" if brand else sku
            status = "⚠️需确认" if item["is_junk"] else "✅就绪"
            self.tree_org.insert("", tk.END, values=(item["filename"], brand, sku, cat, target_proj, status))

    def browse_files_for_org(self):
        files = filedialog.askopenfilenames(title="选择微信接收的 AI / 包装文件", filetypes=[("包装设计文件", "*.ai;*.pdf;*.psd;*.zip;*.rar;*.eps"), ("所有文件", "*.*")])
        if files:
            self.add_files_to_organizer(files)

    def apply_current_brand_to_all_org(self):
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
        edit_win.geometry("420x280")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text=f"原始文件: {cur_item['filename']}", wraplength=380).pack(padx=15, pady=10, anchor="w")
        b_var = tk.StringVar(value=cur_item["brand"])
        s_var = tk.StringVar(value=cur_item["sku"])
        c_var = tk.StringVar(value=cur_item.get("cat", "瓶装"))
        
        f_in = ttk.Frame(edit_win)
        f_in.pack(fill=tk.X, padx=15, pady=5)
        ttk.Label(f_in, text="归属客户:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(f_in, textvariable=b_var, values=self.curated_brands, width=24).grid(row=0, column=1, sticky="w", pady=4)
        
        ttk.Label(f_in, text="包装形态:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(f_in, textvariable=c_var, values=["瓶装", "袋装", "盒装", "套盒", "软管", "罐装", "通用"], width=24).grid(row=1, column=1, sticky="w", pady=4)
        
        ttk.Label(f_in, text="核心SKU名:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(f_in, textvariable=s_var, width=26).grid(row=2, column=1, sticky="w", pady=4)
        
        def save_edit():
            cur_item["brand"] = b_var.get().strip()
            cur_item["sku"] = s_var.get().strip()
            cur_item["cat"] = c_var.get().strip()
            cur_item["is_junk"] = False
            self.refresh_organizer_table()
            edit_win.destroy()
            
        ttk.Button(edit_win, text="保存修改 (Enter)", command=save_edit).pack(pady=12)
        edit_win.bind("<Return>", lambda e: save_edit())

    def execute_organize_flow(self):
        if not self.files_to_organize:
            messagebox.showwarning("提示", "请先添加需要归档的 AI 文件！")
            return
            
        root_dir = self.current_workspace_var.get().strip()
        if not root_dir or not os.path.exists(root_dir):
            try:
                os.makedirs(root_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建主工作盘: {root_dir}\n{e}")
                return
                
        self.save_cfg_all()
        template_blend = self.cfg.get("template_blend_path", "")
        if not template_blend or not os.path.exists(template_blend):
            template_blend = DEFAULT_TEMPLATE
            
        auto_create_blend = self.auto_create_blend_var.get()
        auto_open_blender = self.auto_open_blender_var.get()
        auto_append_excel = self.auto_append_excel_var.get()
        excel_path = self.excel_path_var.get().strip()
        
        subfolders = ["01_Design_平面原稿", "02_Textures_贴图资产", "03_3D_三维工程", "04_Renders_通道输出", "05_Delivery_最终交付"]
        success_count = 0
        duplicate_count = 0
        excel_appended_count = 0
        last_blend = ""
        last_proj = ""
        
        for item in self.files_to_organize:
            brand = item["brand"].strip()
            sku = item["sku"].strip() if item["sku"].strip() else os.path.splitext(item["filename"])[0]
            cat = item.get("cat", "瓶装")
            proj_dir = os.path.join(root_dir, brand, sku) if brand else os.path.join(root_dir, sku)
            
            for sub in subfolders:
                os.makedirs(os.path.join(proj_dir, sub), exist_ok=True)
                
            design_dir = os.path.join(proj_dir, "01_Design_平面原稿")
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
                
            if auto_create_blend:
                target_blend = os.path.join(proj_dir, "03_3D_三维工程", f"{sku}.blend")
                if not os.path.exists(target_blend) and template_blend and os.path.exists(template_blend):
                    try:
                        shutil.copy2(template_blend, target_blend)
                        last_blend = target_blend
                    except Exception:
                        pass
                elif os.path.exists(target_blend):
                    last_blend = target_blend

            # 自动追加写入 Excel 产品台账
            if auto_append_excel and excel_path and os.path.exists(excel_path):
                if append_project_to_excel(excel_path, brand, sku, cat, proj_dir):
                    excel_appended_count += 1

        msg = []
        if success_count > 0:
            msg.append(f"✅ 成功归档并创建 {success_count} 个标准项目！")
            if excel_appended_count > 0:
                msg.append(f"📊 自动同步将 {excel_appended_count} 个新项目录入《产品列表.xlsx》！")
        if duplicate_count > 0:
            msg.append(f"ℹ️ 自动跳过 {duplicate_count} 个微信重复接收文件。")
            
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

    # ---------------- 资产看板数据与同步 ----------------
    def start_excel_auto_sync_watcher(self):
        ex_path = self.excel_path_var.get().strip()
        if ex_path and os.path.exists(ex_path):
            try:
                current_mtime = os.path.getmtime(ex_path)
                if self.last_excel_mtime > 0 and current_mtime > self.last_excel_mtime:
                    self.last_excel_mtime = current_mtime
                    self.thumb_cache.clear()
                    self.load_all_asset_data()
                    self.sync_status_lbl.config(text="⚡ Excel 已更新，已自动同步！", bg="#FEF3C7", fg="#B45309")
                    self.root.after(3500, lambda: self.sync_status_lbl.config(text="🟢 扫描优先实时同步已就绪", bg="#ECFDF5", fg="#059669"))
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
        self.disk_projects = scan_workspace_projects(cur_ws)
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
            
        cat_counts = {}
        for p in self.current_display_list:
            c = p.get("cat", "未分类")
            cat_counts[c] = cat_counts.get(c, 0) + 1
            
        self.category_listbox.delete(0, tk.END)
        self.category_listbox.insert(tk.END, f"✨ 全部形态 ({len(self.current_display_list)})")
        for c in sorted(cat_counts.keys()):
            self.category_listbox.insert(tk.END, f"📦 {c} ({cat_counts[c]})")
            
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

    def on_search_change(self):
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
            m = re.search(r"📦\s*(.*?)\s*\(\d+\)", text)
            if m:
                self.selected_category_var.set(m.group(1))
        self.current_page = 0
        self.apply_filter()

    def apply_filter(self):
        cat = self.selected_category_var.get()
        kw = self.search_var.get().strip().lower()
        
        self.filtered_projects = []
        for p in self.current_display_list:
            if cat != "全部" and p.get("cat") != cat:
                continue
            if kw and (kw not in p["sku"].lower() and kw not in p.get("brand", "").lower() and kw not in p.get("cat", "").lower()):
                continue
            self.filtered_projects.append(p)
            
        total_items = len(self.filtered_projects)
        total_pages = max(1, (total_items + self.page_size - 1) // self.page_size)
        if self.current_page >= total_pages:
            self.current_page = total_pages - 1
            
        start_idx = self.current_page * self.page_size + 1 if total_items > 0 else 0
        end_idx = min((self.current_page + 1) * self.page_size, total_items)
        
        self.page_info_lbl.config(text=f"共 {total_items} 个产品 | 正在显示第 {start_idx} - {end_idx} 项 (第 {self.current_page + 1}/{total_pages} 页)")
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
        if img_path in self.thumb_cache:
            return self.thumb_cache[img_path]
        try:
            im = Image.open(img_path)
            im.thumbnail(size, Image.Resampling.LANCZOS)
            thumb = Image.new("RGBA", size, (255, 255, 255, 255))
            offset_x = (size[0] - im.width) // 2
            offset_y = (size[1] - im.height) // 2
            if im.mode == "RGBA":
                thumb.paste(im, (offset_x, offset_y), im)
            else:
                thumb.paste(im, (offset_x, offset_y))
            tk_img = ImageTk.PhotoImage(thumb)
            self.thumb_cache[img_path] = tk_img
            return tk_img
        except Exception:
            return self.get_placeholder_thumbnail(size)

    def get_placeholder_thumbnail(self, size=(190, 190)):
        cache_key = "placeholder"
        if cache_key in self.thumb_cache:
            return self.thumb_cache[cache_key]
        im = Image.new("RGBA", size, (240, 243, 246, 255))
        draw = ImageDraw.Draw(im)
        draw.text((size[0]//2 - 35, size[1]//2 - 10), "📦 暂无渲染图", fill=(160, 170, 185, 255))
        tk_img = ImageTk.PhotoImage(im)
        self.thumb_cache[cache_key] = tk_img
        return tk_img

    def render_cards(self):
        for widget in self.grid_container.winfo_children():
            widget.destroy()
            
        if not self.filtered_projects:
            no_lbl = tk.Label(self.grid_container, text="📭 没有找到匹配的包装产品", font=("Microsoft YaHei", 12), bg="#F0F2F5", fg="#888", pady=60)
            no_lbl.pack()
            return

        container_width = self.canvas.winfo_width()
        if container_width < 100:
            container_width = 800
        card_w = 210
        cols = max(1, container_width // (card_w + 16))

        start_idx = self.current_page * self.page_size
        end_idx = start_idx + self.page_size
        page_items = self.filtered_projects[start_idx:end_idx]

        for idx, proj in enumerate(page_items):
            row = idx // cols
            col = idx % cols
            
            card = tk.Frame(self.grid_container, bg="white", bd=1, relief=tk.SOLID, padx=8, pady=8, highlightthickness=0)
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            
            tk_thumb = self.get_scaled_thumbnail(proj["thumbnail"])
            img_lbl = tk.Label(card, image=tk_thumb, bg="white", cursor="hand2")
            img_lbl.image = tk_thumb
            img_lbl.pack(fill=tk.BOTH, expand=True)
            
            meta_frame = tk.Frame(card, bg="white", pady=4)
            meta_frame.pack(fill=tk.X)
            
            badge_row = tk.Frame(meta_frame, bg="white")
            badge_row.pack(fill=tk.X, pady=(0, 2))
            
            cat_tag = tk.Label(badge_row, text=proj.get("cat", "包装"), font=("Microsoft YaHei", 8, "bold"), bg="#E1EFFF", fg="#005A9E", padx=4, pady=1)
            cat_tag.pack(side=tk.LEFT, padx=(0, 4))
            
            if proj.get("brand"):
                b_tag = tk.Label(badge_row, text=proj["brand"], font=("Microsoft YaHei", 8), bg="#F0F0F0", fg="#555", padx=4, pady=1)
                b_tag.pack(side=tk.LEFT)
            
            title_lbl = tk.Label(meta_frame, text=proj["sku"], font=("Microsoft YaHei", 9, "bold"), bg="white", fg="#222", wraplength=180, justify="left")
            title_lbl.pack(anchor="w")
            
            action_frame = tk.Frame(card, bg="white", pady=4)
            action_frame.pack(fill=tk.X)
            
            has_path = bool(proj.get("path") and os.path.exists(proj["path"]))
            btn_open = tk.Button(action_frame, text="📁 打开文件夹" if has_path else "📁 路径未就绪", font=("Microsoft YaHei", 8), bg="#F0F0F0" if has_path else "#FAFAFA", fg="#222" if has_path else "#999", relief=tk.FLAT, command=lambda p=proj.get("path"): self.open_folder(p))
            btn_open.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
            
            btn_blend = tk.Button(action_frame, text="🚀 3D工程", font=("Microsoft YaHei", 8), bg="#EBF5FB", fg="#005A9E", relief=tk.FLAT, command=lambda p=proj.get("path"): self.launch_blend(p))
            btn_blend.pack(side=tk.RIGHT)
            
            for w in (card, img_lbl, title_lbl, meta_frame):
                w.bind("<Button-1>", lambda e, p=proj.get("path"): self.open_folder(p))
                w.bind("<Double-1>", lambda e, p=proj.get("path"): self.launch_blend(p))
                w.bind("<Button-3>", lambda e, pr=proj: self.show_context_menu(e, pr))

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
        blend_dir = os.path.join(proj_path, "03_3D_三维工程")
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
        menu = tk.Menu(self.root, tearoff=0)
        p = proj.get("path", "")
        menu.add_command(label=f"📁 打开项目文件夹: {proj['sku']}", command=lambda: self.open_folder(p))
        menu.add_command(label="🚀 启动 Blender 打开 3D 工程", command=lambda: self.launch_blend(p))
        if p and os.path.exists(p):
            menu.add_command(label="🎨 打开 01_Design 平面原稿", command=lambda: self.open_folder(os.path.join(p, "01_Design_平面原稿")))
            menu.add_command(label="🖼️ 查看 04_Renders 渲染大图", command=lambda: self.open_folder(os.path.join(p, "04_Renders_通道输出")))
        menu.add_separator()
        menu.add_command(label="📋 复制项目完整路径", command=lambda: self.copy_path_to_clipboard(p))
        menu.tk_popup(event.x_root, event.y_root)

    def copy_path_to_clipboard(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制路径到剪贴板:\n{text}")

    def export_html_gallery(self):
        cur_ws = self.current_workspace_var.get().strip()
        html_file = os.path.join(cur_ws if os.path.exists(cur_ws) else os.path.expanduser("~"), "📦_包装项目全景视觉画廊.html")
        cards_html = []
        for p in self.current_display_list:
            thumb_rel = p.get("thumbnail") or ""
            thumb_src = "file:///" + thumb_rel.replace("\\", "/") if (thumb_rel and os.path.exists(thumb_rel)) else "https://via.placeholder.com/300x300?text=No+Render"
            folder_uri = "file:///" + p["path"].replace("\\", "/") if p.get("path") else "#"
            cards_html.append(f"""
            <div class="card" onclick="window.open('{folder_uri}')">
                <div class="thumb-container"><img src="{thumb_src}" alt="{p['sku']}" loading="lazy"></div>
                <div class="meta">
                    <div style="display:flex; gap:6px; margin-bottom:6px;">
                        <span class="badge" style="background:#0369a1;">{p.get('cat', '包装')}</span>
                        <span class="badge" style="background:#475569;">{p.get('brand', '')}</span>
                    </div>
                    <h3 class="title">{p['sku']}</h3>
                    <p class="path">{p.get('path', '')}</p>
                </div>
            </div>
            """)
        full_html = f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>📦 包装项目全景视觉画廊</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 24px; }}
h1 {{ font-size: 24px; font-weight: 700; margin-bottom: 24px; }}
.grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 20px; }}
.card {{ background: #1e293b; border-radius: 12px; overflow: hidden; border: 1px solid #334155; cursor: pointer; transition: transform 0.2s; }}
.card:hover {{ transform: translateY(-4px); border-color: #38bdf8; }}
.thumb-container {{ width: 100%; aspect-ratio: 1; background: #0f172a; display: flex; align-items: center; justify-content: center; }}
.thumb-container img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
.meta {{ padding: 12px; }}
.badge {{ display: inline-block; color: #e0f2fe; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 9999px; }}
.title {{ font-size: 14px; font-weight: 600; margin: 0 0 4px 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.path {{ font-size: 11px; color: #94a3b8; margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
</style></head><body><h1>📦 包装项目全景视觉画廊 (共 {len(self.current_display_list)} 个产品)</h1><div class="grid">{"".join(cards_html)}</div></body></html>"""
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
