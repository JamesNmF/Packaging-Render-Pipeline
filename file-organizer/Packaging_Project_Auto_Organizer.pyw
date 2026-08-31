# -*- coding: utf-8 -*-
"""
微信 AI 文件一键智能归档与项目脚手架创建器 (v2.6 Blender 首次创建自动化版)
核心功能：
1. 【首次创建 0 步骤开工】：归档 AI 文件时，自动克隆母版模板并命名为 "[产品名].blend" 放入 03_3D_三维工程/，并直接自动启动 Blender 打开！
2. 【增量自动归拢】：自动识别 "(1)", "（2）", "- 副本" 等增量后缀，归拢至同一项目主干。
3. 【MD5 智能查重】：秒级识别微信重复文件，杜绝假更新。
4. 【多工作盘秒切 & 客户白名单】：多盘无缝切换，列表 100% 纯净精准。
"""

import os
import sys
import re
import json
import shutil
import hashlib
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_organizer_v6.json")

# 默认 Blender 5.2 路径与母版模板路径
BLENDER_EXE = r"C:\Program Files\Blender Foundation\Blender 5.2\blender.exe"
if not os.path.exists(BLENDER_EXE):
    BLENDER_EXE = "blender"

DEFAULT_TEMPLATE = os.path.join(
    os.path.expanduser("~"), "Desktop", "AI_Blender包装渲染辅助工具", "templates", "Packaging_Master_Template.blend"
)

DEFAULT_WORKSPACES = []
for p in ["E:\\zjc", "D:\\Projects", "E:\\Projects", "D:\\", "E:\\"]:
    if os.path.exists(p) and p not in DEFAULT_WORKSPACES:
        DEFAULT_WORKSPACES.append(p)
if not DEFAULT_WORKSPACES:
    DEFAULT_WORKSPACES = ["D:\\Projects"]

DEFAULT_CONFIG = {
    "workspaces": DEFAULT_WORKSPACES,
    "current_workspace": DEFAULT_WORKSPACES[0],
    "curated_brands": ["柏缇", "零食有鸣"],
    "current_brand": "柏缇",
    "auto_create_blend": True,
    "auto_open_blender": True,
    "template_blend_path": DEFAULT_TEMPLATE if os.path.exists(DEFAULT_TEMPLATE) else "",
    "blender_exe_path": BLENDER_EXE
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not data.get("workspaces"):
                    data["workspaces"] = DEFAULT_WORKSPACES
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
                sku = f"未命名产品_{raw_name}"
            else:
                brand = fallback_brand
                sku = parts[0]
        else:
            brand = fallback_brand
            sku = f"未命名产品_{raw_name}"
            
    return brand, sku, is_junk


def get_next_design_version_filename(design_dir, brand, sku, ext, source_md5=None):
    if not os.path.exists(design_dir):
        return f"{sku}_v01{ext}", False, ""
        
    existing_files = os.listdir(design_dir)
    if not existing_files:
        return f"{sku}_v01{ext}", False, ""
        
    if source_md5:
        for ef in existing_files:
            ef_path = os.path.join(design_dir, ef)
            if os.path.isfile(ef_path):
                if get_file_md5(ef_path) == source_md5:
                    return ef, True, ef
                    
    pattern = re.compile(rf"^{re.escape(sku)}_v(\d+)", re.IGNORECASE)
    max_v = 0
    for ef in existing_files:
        m = pattern.match(ef)
        if m:
            try:
                v_num = int(m.group(1))
                if v_num > max_v:
                    max_v = v_num
            except Exception:
                pass
                
    if max_v > 0:
        return f"{sku}_v{max_v + 1:02d}{ext}", False, ""
    else:
        return f"{sku}_v{len(existing_files) + 1:02d}{ext}", False, ""


class OrganizerApp:
    def __init__(self, root, initial_files=None):
        self.root = root
        self.root.title("📦 包装 AI 文件智能归档器 (v2.6 首次创建自动化版)")
        self.root.geometry("880x660")
        self.root.minsize(780, 560)
        
        self.cfg = load_config()
        self.workspaces = self.cfg.get("workspaces", DEFAULT_WORKSPACES)
        self.current_workspace_var = tk.StringVar(value=self.cfg.get("current_workspace", self.workspaces[0]))
        
        self.curated_brands = self.cfg.get("curated_brands", ["柏缇", "零食有鸣"])
        self.current_brand_var = tk.StringVar(value=self.cfg.get("current_brand", self.curated_brands[0]))
        
        self.auto_create_blend_var = tk.BooleanVar(value=self.cfg.get("auto_create_blend", True))
        self.auto_open_blender_var = tk.BooleanVar(value=self.cfg.get("auto_open_blender", True))
        
        self.files_data = []
        
        self.build_ui()
        self.update_workspace_combo()
        self.update_brand_combo()
        
        if initial_files:
            self.add_files(initial_files)

    def build_ui(self):
        # 1. 顶部控制栏：多工作盘 + 客户白名单
        top_frame = ttk.LabelFrame(self.root, text=" 📂 多工作盘快速切换 & 客户品牌管理 ", padding=10)
        top_frame.pack(fill=tk.X, padx=15, pady=6)
        
        row_dir = ttk.Frame(top_frame)
        row_dir.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(row_dir, text="当前主工作盘:").pack(side=tk.LEFT, padx=(0, 6))
        
        self.workspace_combo = ttk.Combobox(
            row_dir,
            textvariable=self.current_workspace_var,
            font=("Microsoft YaHei", 9),
            width=36,
            state="readonly"
        )
        self.workspace_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.workspace_combo.bind("<<ComboboxSelected>>", self.on_workspace_select)
        
        ttk.Button(row_dir, text="➕ 绑定新工作盘...", command=self.add_workspace).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(row_dir, text="❌ 移除此工作盘", command=self.remove_workspace).pack(side=tk.LEFT)
        
        row_brand = ttk.Frame(top_frame)
        row_brand.pack(fill=tk.X)
        ttk.Label(row_brand, text="指定客户/品牌:").pack(side=tk.LEFT, padx=(0, 6))
        
        self.brand_combo = ttk.Combobox(
            row_brand,
            textvariable=self.current_brand_var,
            font=("Microsoft YaHei", 9),
            width=20,
            state="readonly"
        )
        self.brand_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.brand_combo.bind("<<ComboboxSelected>>", self.on_brand_select)
        
        ttk.Button(row_brand, text="➕ 添加新客户...", command=self.add_brand).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(row_brand, text="📁 点选已有客户文件夹...", command=self.pick_brand_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(row_brand, text="❌ 移除此客户", command=self.remove_brand).pack(side=tk.LEFT)

        # 2. Blender 自动化首次创建配置行
        b_frame = ttk.LabelFrame(self.root, text=" ⚡ Blender 首次创建自动化 (0 步骤自动建工程) ", padding=8)
        b_frame.pack(fill=tk.X, padx=15, pady=(0, 6))
        
        row_b = ttk.Frame(b_frame)
        row_b.pack(fill=tk.X)
        
        ttk.Checkbutton(
            row_b,
            text="✨ 归档时自动生成对应 .blend 工程",
            variable=self.auto_create_blend_var,
            command=self.save_blend_options
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Checkbutton(
            row_b,
            text="🚀 创建后自动启动 Blender 打开工程",
            variable=self.auto_open_blender_var,
            command=self.save_blend_options
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Button(row_b, text="📁 设置我的默认包装母版 .blend...", command=self.set_custom_template).pack(side=tk.RIGHT)

        # 3. 中间文件列表与智能解析表格
        list_frame = ttk.LabelFrame(self.root, text=" 📋 待处理源文件 (自动剔除 (1) (2) 归拢至同项目，支持双击编辑) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 6))
        
        cols = ("file", "brand", "sku", "target_dir", "status")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("file", text="微信接收的文件名")
        self.tree.heading("brand", text="归属客户/品牌")
        self.tree.heading("sku", text="核心项目/SKU主干")
        self.tree.heading("target_dir", text="目标归档目录")
        self.tree.heading("status", text="分析状态")
        
        self.tree.column("file", width=220, anchor="w")
        self.tree.column("brand", width=110, anchor="center")
        self.tree.column("sku", width=180, anchor="w")
        self.tree.column("target_dir", width=180, anchor="w")
        self.tree.column("status", width=90, anchor="center")
        
        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # 4. 快捷操作栏
        btn_frame = ttk.Frame(self.root, padding=2)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 6))
        
        ttk.Button(btn_frame, text="➕ 添加 AI / 源文件...", command=self.browse_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="✏️ 批量应用当前客户至全部", command=self.apply_current_brand_to_all).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="❌ 移除选中文件", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="🧹 清空列表", command=self.clear_all).pack(side=tk.LEFT)
        
        # 5. 底部执行大按钮
        exec_frame = ttk.Frame(self.root, padding=6)
        exec_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        self.btn_exec = tk.Button(
            exec_frame,
            text="🚀 【 一键创建工业级标准项目、生成 .blend 并自动开工 】",
            font=("Microsoft YaHei", 11, "bold"),
            bg="#0078D7",
            fg="white",
            activebackground="#005A9E",
            activeforeground="white",
            relief=tk.FLAT,
            height=2,
            command=self.execute_organize
        )
        self.btn_exec.pack(fill=tk.X)

    def save_blend_options(self):
        self.cfg["auto_create_blend"] = self.auto_create_blend_var.get()
        self.cfg["auto_open_blender"] = self.auto_open_blender_var.get()
        save_config(self.cfg)

    def set_custom_template(self):
        f = filedialog.askopenfilename(
            title="选择你的默认包装 Blender 母版工程 (.blend)",
            filetypes=[("Blender 工程文件", "*.blend"), ("所有文件", "*.*")]
        )
        if f:
            self.cfg["template_blend_path"] = f
            save_config(self.cfg)
            messagebox.showinfo("设置成功", f"已成功将默认母版设为:\n{os.path.basename(f)}")

    def update_workspace_combo(self):
        self.workspace_combo["values"] = self.workspaces
        cur = self.current_workspace_var.get().strip()
        if (not cur or cur not in self.workspaces) and self.workspaces:
            self.current_workspace_var.set(self.workspaces[0])
            self.cfg["current_workspace"] = self.workspaces[0]
            save_config(self.cfg)

    def on_workspace_select(self, event=None):
        cur = self.current_workspace_var.get().strip()
        self.cfg["current_workspace"] = cur
        save_config(self.cfg)
        self.refresh_table_views()

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
            self.update_workspace_combo()
            self.refresh_table_views()
            messagebox.showinfo("成功", f"已成功绑定新工作盘: [{norm_d}]！")

    def remove_workspace(self):
        cur = self.current_workspace_var.get().strip()
        if len(self.workspaces) <= 1:
            messagebox.showwarning("提示", "至少需保留一个工作盘路径！")
            return
        if messagebox.askyesno("确认", f"确定从列表中移除工作盘 [{cur}] 吗？\n(不会删除硬盘文件)"):
            self.workspaces.remove(cur)
            self.cfg["workspaces"] = self.workspaces
            self.cfg["current_workspace"] = self.workspaces[0]
            self.current_workspace_var.set(self.workspaces[0])
            save_config(self.cfg)
            self.update_workspace_combo()
            self.refresh_table_views()

    def update_brand_combo(self):
        self.brand_combo["values"] = self.curated_brands
        cur = self.current_brand_var.get().strip()
        if (not cur or cur not in self.curated_brands) and self.curated_brands:
            self.current_brand_var.set(self.curated_brands[0])
            self.cfg["current_brand"] = self.curated_brands[0]
            save_config(self.cfg)

    def on_brand_select(self, event=None):
        cur = self.current_brand_var.get().strip()
        self.cfg["current_brand"] = cur
        save_config(self.cfg)

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
                self.update_brand_combo()
                
                cur_ws = self.current_workspace_var.get().strip()
                if cur_ws and os.path.exists(cur_ws):
                    os.makedirs(os.path.join(cur_ws, name), exist_ok=True)
                    
                messagebox.showinfo("成功", f"已成功添加客户: [{name}]！")

    def pick_brand_folder(self):
        cur_ws = self.current_workspace_var.get().strip()
        initial = cur_ws if os.path.exists(cur_ws) else None
        d = filedialog.askdirectory(title="点选指定的客户/品牌文件夹", initialdir=initial)
        if d:
            brand_name = os.path.basename(os.path.normpath(d))
            if brand_name:
                if brand_name not in self.curated_brands:
                    self.curated_brands.insert(0, brand_name)
                    self.cfg["curated_brands"] = self.curated_brands
                self.cfg["current_brand"] = brand_name
                save_config(self.cfg)
                self.current_brand_var.set(brand_name)
                self.update_brand_combo()

    def remove_brand(self):
        cur = self.current_brand_var.get().strip()
        if len(self.curated_brands) <= 1:
            messagebox.showwarning("提示", "列表中至少需保留一个客户名称！")
            return
        if messagebox.askyesno("确认", f"确定从快速选择列表中移除客户 [{cur}] 吗？\n(不会删除硬盘文件)"):
            self.curated_brands.remove(cur)
            self.cfg["curated_brands"] = self.curated_brands
            self.cfg["current_brand"] = self.curated_brands[0]
            self.current_brand_var.set(self.curated_brands[0])
            save_config(self.cfg)
            self.update_brand_combo()

    def apply_current_brand_to_all(self):
        cur_brand = self.current_brand_var.get().strip()
        if not cur_brand:
            return
        for item in self.files_data:
            item["brand"] = cur_brand
        self.refresh_table_views()

    def refresh_table_views(self):
        self.tree.delete(*self.tree.get_children())
        cur_ws = self.current_workspace_var.get().strip()
        
        for item in self.files_data:
            brand = item["brand"]
            sku = item["sku"]
            target_proj = f"{brand}/{sku}" if brand else sku
            status = "⚠️需确认" if item["is_junk"] else "✅就绪"
            self.tree.insert("", tk.END, values=(item["filename"], brand, sku, target_proj, status))

    def browse_files(self):
        files = filedialog.askopenfilenames(
            title="选择微信接收的 AI / 包装设计文件",
            filetypes=[("包装设计源文件", "*.ai;*.pdf;*.psd;*.zip;*.rar;*.eps;*.cdr"), ("所有文件", "*.*")]
        )
        if files:
            self.add_files(files)

    def add_files(self, filepaths):
        cur_brand = self.current_brand_var.get().strip()
        for fp in filepaths:
            fp = os.path.abspath(fp)
            if not os.path.exists(fp) or not os.path.isfile(fp):
                continue
            if any(item['filepath'] == fp for item in self.files_data):
                continue
                
            brand, sku, is_junk = clean_and_parse_filename(fp, fallback_brand=cur_brand, valid_brands=self.curated_brands)
            
            item = {
                "filepath": fp,
                "filename": os.path.basename(fp),
                "brand": brand,
                "sku": sku,
                "is_junk": is_junk,
                "md5": get_file_md5(fp)
            }
            self.files_data.append(item)
            
        self.refresh_table_views()

    def remove_selected(self):
        selected = self.tree.selection()
        for s in reversed(selected):
            idx = self.tree.index(s)
            self.tree.delete(s)
            del self.files_data[idx]

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.files_data.clear()

    def on_double_click(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        idx = self.tree.index(item_id)
        cur_item = self.files_data[idx]
        
        edit_win = tk.Toplevel(self.root)
        edit_win.title("✏️ 快速修改客户与项目名")
        edit_win.geometry("420x240")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text=f"原始文件: {cur_item['filename']}", wraplength=380).pack(padx=15, pady=10, anchor="w")
        
        b_var = tk.StringVar(value=cur_item["brand"])
        s_var = tk.StringVar(value=cur_item["sku"])
        
        f_in = ttk.Frame(edit_win)
        f_in.pack(fill=tk.X, padx=15, pady=5)
        
        ttk.Label(f_in, text="归属客户/品牌:").grid(row=0, column=0, sticky="w", pady=5)
        b_entry = ttk.Combobox(f_in, textvariable=b_var, values=self.curated_brands, width=26)
        b_entry.grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(f_in, text="核心项目/SKU名:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(f_in, textvariable=s_var, width=28).grid(row=1, column=1, sticky="w", pady=5)
        
        def save_edit():
            cur_item["brand"] = b_var.get().strip()
            cur_item["sku"] = s_var.get().strip()
            cur_item["is_junk"] = False
            self.refresh_table_views()
            edit_win.destroy()
            
        btn_s = ttk.Button(edit_win, text="保存修改 (Enter)", command=save_edit)
        btn_s.pack(pady=12)
        edit_win.bind("<Return>", lambda e: save_edit())

    def execute_organize(self):
        if not self.files_data:
            messagebox.showwarning("提示", "请先添加需要归档的 AI 文件！")
            return
            
        root_dir = self.current_workspace_var.get().strip()
        if not root_dir or not os.path.exists(root_dir):
            try:
                os.makedirs(root_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建主工作盘路径: {root_dir}\n{e}")
                return
                
        self.cfg["current_workspace"] = root_dir
        save_config(self.cfg)
        
        success_count = 0
        duplicate_count = 0
        last_created_blend_path = ""
        last_created_proj_dir = ""
        
        subfolders = [
            "01_Design_平面原稿",
            "02_Textures_贴图资产",
            "03_3D_三维工程",
            "04_Renders_通道输出",
            "05_Delivery_最终交付"
        ]
        
        template_blend = self.cfg.get("template_blend_path", "")
        if not template_blend or not os.path.exists(template_blend):
            template_blend = DEFAULT_TEMPLATE
            
        auto_create_blend = self.auto_create_blend_var.get()
        auto_open_blender = self.auto_open_blender_var.get()
        
        for item in self.files_data:
            brand = item["brand"].strip()
            sku = item["sku"].strip() if item["sku"].strip() else os.path.splitext(item["filename"])[0]
            
            if brand:
                proj_dir = os.path.join(root_dir, brand, sku)
            else:
                proj_dir = os.path.join(root_dir, sku)
                
            for sub in subfolders:
                os.makedirs(os.path.join(proj_dir, sub), exist_ok=True)
                
            design_dir = os.path.join(proj_dir, "01_Design_平面原稿")
            ext = os.path.splitext(item["filename"])[1]
            
            dest_file_name, is_duplicate, dup_name = get_next_design_version_filename(
                design_dir, brand, sku, ext, source_md5=item.get("md5")
            )
            
            if is_duplicate:
                duplicate_count += 1
                last_created_proj_dir = proj_dir
                continue
                
            dest_file_path = os.path.join(design_dir, dest_file_name)
            try:
                shutil.copy2(item["filepath"], dest_file_path)
                success_count += 1
                last_created_proj_dir = proj_dir
            except Exception as e:
                print(f"Error copying {item['filepath']}: {e}")
                
            # 自动生成 03_3D_三维工程\[SKU].blend
            if auto_create_blend:
                target_blend = os.path.join(proj_dir, "03_3D_三维工程", f"{sku}.blend")
                if not os.path.exists(target_blend) and template_blend and os.path.exists(template_blend):
                    try:
                        shutil.copy2(template_blend, target_blend)
                        last_created_blend_path = target_blend
                    except Exception as e:
                        print(f"Error cloning blend template: {e}")
                elif os.path.exists(target_blend):
                    last_created_blend_path = target_blend
                    
        # 结果反馈与自动启动 Blender
        msg_parts = []
        if success_count > 0:
            msg_parts.append(f"✅ 成功归档/增量录入 {success_count} 个版本！")
            if auto_create_blend:
                msg_parts.append("✨ 自动初始化了对应的 3D Blender 工程！")
        if duplicate_count > 0:
            msg_parts.append(f"ℹ️ 自动识别并跳过 {duplicate_count} 个微信完全重复接收的文件。")
            
        if auto_open_blender and last_created_blend_path and os.path.exists(last_created_blend_path):
            blender_bin = self.cfg.get("blender_exe_path", BLENDER_EXE)
            try:
                subprocess.Popen([blender_bin, last_created_blend_path])
                msg_parts.append(f"🚀 已自动启动 Blender 5.2 打开工程: [{os.path.basename(last_created_blend_path)}]")
            except Exception as e:
                os.startfile(last_created_blend_path)
        elif last_created_proj_dir and os.path.exists(last_created_proj_dir):
            try:
                os.startfile(last_created_proj_dir)
            except Exception:
                pass
                
        messagebox.showinfo("🎉 处理完成", "\n".join(msg_parts))
        self.clear_all()


if __name__ == "__main__":
    args_files = sys.argv[1:] if len(sys.argv) > 1 else None
    root_win = tk.Tk()
    
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
        
    app = OrganizerApp(root_win, initial_files=args_files)
    root_win.mainloop()