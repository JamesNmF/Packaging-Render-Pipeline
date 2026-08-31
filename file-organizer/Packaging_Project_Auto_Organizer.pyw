# -*- coding: utf-8 -*-
"""
微信 AI 文件一键智能归档与项目脚手架创建器 (v2.3 精准白名单管理版)
核心改动：
1. 彻底禁用全盘盲目扫描，杜绝把 "素材/临时/字体/下载" 等杂乱文件夹混入品牌列表！
2. 纯净用户白名单：只有你手动添加或点选的客户才会进入列表。
3. 提供【➕ 新增】、【❌ 移除】、【📁 选择单个客户文件夹】，100% 精准干净。
"""

import os
import sys
import re
import json
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_organizer_v3.json")

DEFAULT_ROOT_DIR = "D:\\Projects"
if os.path.exists("E:\\zjc"):
    DEFAULT_ROOT_DIR = "E:\\zjc"
elif os.path.exists("E:\\Projects"):
    DEFAULT_ROOT_DIR = "E:\\Projects"
elif os.path.exists("E:\\"):
    DEFAULT_ROOT_DIR = "E:\\"

DEFAULT_CONFIG = {
    "root_dir": DEFAULT_ROOT_DIR,
    "curated_brands": ["柏缇", "零食有鸣"],
    "current_brand": "柏缇"
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_CONFIG

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


JUNK_NAME_KEYWORDS = {
    '改', '改1', '改2', '改3', '改4', '改5', '修改', '修改版', '最新', '最终', '最终版', '定稿', '正稿',
    '未命名', '未命名-1', '新建', '新建画板', '新建画板1', '新建画板2', '1', '2', '3', 'a', 'b', 'c',
    '刀模', '包装', '瓶贴', '贴纸', '展开图', '画板', '副本', '111', '222', 'aaa'
}

def clean_and_parse_filename(filepath, fallback_brand="", valid_brands=None):
    raw_name = os.path.splitext(os.path.basename(filepath))[0]
    
    noise_patterns = [
        r'[-_ ]?(包装|刀模|展开图|正稿|定稿|完稿|原稿|印刷稿|平面|效果图)',
        r'[-_ ]?(修改版|修改|最新版|最终版|最终|定案|终版|初稿|打样|打样稿)',
        r'[-_ ]?(副本|\d{6,}|\d{4}年|\d{1,2}月\d{1,2}日)',
        r'[-_ ]?([vV]\d+(\.\d+)?|改\d*|版\d*)',
    ]
    
    cleaned = raw_name
    for p in noise_patterns:
        cleaned = re.sub(p, '', cleaned, flags=re.IGNORECASE)
        
    cleaned = cleaned.strip(" -_")
    is_junk = cleaned.lower() in JUNK_NAME_KEYWORDS or len(cleaned) == 0
    
    parts = re.split(r'[-_—\s+]+', cleaned)
    parts = [p.strip() for p in parts if p.strip()]
    
    # 优先匹配已知白名单品牌
    brand = ""
    sku = ""
    
    if valid_brands:
        for vb in valid_brands:
            if cleaned.startswith(vb):
                brand = vb
                sku = cleaned[len(vb):].strip(" -_")
                break
                
    if not brand:
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


class OrganizerApp:
    def __init__(self, root, initial_files=None):
        self.root = root
        self.root.title("📦 包装 AI 文件智能归档与项目创建器 (精准白名单版)")
        self.root.geometry("800x600")
        self.root.minsize(720, 500)
        
        self.cfg = load_config()
        self.root_dir_var = tk.StringVar(value=self.cfg.get("root_dir", DEFAULT_ROOT_DIR))
        self.current_brand_var = tk.StringVar(value=self.cfg.get("current_brand", "柏缇"))
        self.curated_brands = self.cfg.get("curated_brands", ["柏缇", "零食有鸣"])
        
        self.files_data = []
        
        self.build_ui()
        self.update_combo_values()
        
        if initial_files:
            self.add_files(initial_files)

    def build_ui(self):
        # 1. 顶部控制栏：纯净客户白名单 + 工作盘
        top_frame = ttk.LabelFrame(self.root, text=" 📂 项目主工作盘与指定客户/品牌 ", padding=10)
        top_frame.pack(fill=tk.X, padx=15, pady=8)
        
        # 根目录行
        row_dir = ttk.Frame(top_frame)
        row_dir.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(row_dir, text="主工作盘:").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(row_dir, textvariable=self.root_dir_var, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(row_dir, text="选择工作盘...", command=self.browse_root_dir).pack(side=tk.RIGHT)
        
        # 品牌选择与白名单管理行
        row_brand = ttk.Frame(top_frame)
        row_brand.pack(fill=tk.X)
        ttk.Label(row_brand, text="当前指定客户/品牌:").pack(side=tk.LEFT, padx=(0, 6))
        
        self.brand_combo = ttk.Combobox(
            row_brand,
            textvariable=self.current_brand_var,
            font=("Microsoft YaHei", 9),
            width=18,
            state="readonly"
        )
        self.brand_combo.pack(side=tk.LEFT, padx=(0, 8))
        self.brand_combo.bind("<<ComboboxSelected>>", self.on_brand_select)
        
        # 精准管理按钮（不盲目全盘扫描）
        ttk.Button(row_brand, text="➕ 添加新客户...", command=self.add_brand).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(row_brand, text="📁 点选已有客户文件夹...", command=self.pick_brand_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(row_brand, text="❌ 移除选中客户", command=self.remove_brand).pack(side=tk.LEFT)

        # 2. 中间文件列表与智能解析表格
        list_frame = ttk.LabelFrame(self.root, text=" 📋 待处理的微信源文件 (双击任意一行可随时修改品牌或SKU) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 8))
        
        cols = ("file", "brand", "sku", "status")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("file", text="微信接收的文件名")
        self.tree.heading("brand", text="归属【客户/品牌】")
        self.tree.heading("sku", text="创建的【产品/SKU工程名】")
        self.tree.heading("status", text="状态")
        
        self.tree.column("file", width=240, anchor="w")
        self.tree.column("brand", width=140, anchor="center")
        self.tree.column("sku", width=240, anchor="w")
        self.tree.column("status", width=80, anchor="center")
        
        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # 3. 快捷操作栏
        btn_frame = ttk.Frame(self.root, padding=2)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 8))
        
        ttk.Button(btn_frame, text="➕ 添加 AI / 源文件...", command=self.browse_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="✏️ 批量应用当前客户至全部", command=self.apply_current_brand_to_all).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="❌ 移除选中文件", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="🧹 清空列表", command=self.clear_all).pack(side=tk.LEFT)
        
        # 4. 底部执行大按钮
        exec_frame = ttk.Frame(self.root, padding=8)
        exec_frame.pack(fill=tk.X, padx=15, pady=(0, 12))
        
        self.btn_exec = tk.Button(
            exec_frame,
            text="🚀 【 一键创建工业级标准项目并归档 】",
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

    def update_combo_values(self):
        """仅更新纯净白名单列表"""
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
        """精准添加新客户"""
        name = simpledialog.askstring("添加新客户", "请输入客户/品牌名称 (如：统一、农夫山泉):", parent=self.root)
        if name and name.strip():
            name = name.strip()
            if name not in self.curated_brands:
                self.curated_brands.insert(0, name)
                self.cfg["curated_brands"] = self.curated_brands
                self.cfg["current_brand"] = name
                save_config(self.cfg)
                self.current_brand_var.set(name)
                self.update_combo_values()
                
                # 自动在工作盘创建该客户文件夹
                root_dir = self.root_dir_var.get().strip()
                if root_dir and os.path.exists(root_dir):
                    os.makedirs(os.path.join(root_dir, name), exist_ok=True)
                    
                messagebox.showinfo("成功", f"已精准添加客户: [{name}]！")

    def pick_brand_folder(self):
        """只添加用户明确点选的那个文件夹，绝不盲目读取整盘"""
        root_dir = self.root_dir_var.get().strip()
        initial = root_dir if os.path.exists(root_dir) else None
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
                self.update_combo_values()

    def remove_brand(self):
        """从白名单移除选中的客户（不删硬盘文件）"""
        cur = self.current_brand_var.get().strip()
        if not cur:
            return
        if len(self.curated_brands) <= 1:
            messagebox.showwarning("提示", "列表中至少需保留一个客户名称！")
            return
        if messagebox.askyesno("确认", f"确定从快速选择列表中移除客户 [{cur}] 吗？\n(注意：不会删除硬盘中的实际文件)"):
            self.curated_brands.remove(cur)
            self.cfg["curated_brands"] = self.curated_brands
            self.cfg["current_brand"] = self.curated_brands[0]
            save_config(self.cfg)
            self.current_brand_var.set(self.curated_brands[0])
            self.update_combo_values()

    def apply_current_brand_to_all(self):
        cur_brand = self.current_brand_var.get().strip()
        if not cur_brand:
            return
        for idx, item in enumerate(self.files_data):
            item["brand"] = cur_brand
            self.tree.item(self.tree.get_children()[idx], values=(item["filename"], item["brand"], item["sku"], "已就绪"))

    def browse_root_dir(self):
        d = filedialog.askdirectory(initialdir=self.root_dir_var.get())
        if d:
            self.root_dir_var.set(d)
            self.cfg["root_dir"] = d
            save_config(self.cfg)

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
            status = "⚠️需确认" if is_junk else "已就绪"
            
            item = {
                "filepath": fp,
                "filename": os.path.basename(fp),
                "brand": brand,
                "sku": sku,
                "is_junk": is_junk
            }
            self.files_data.append(item)
            self.tree.insert("", tk.END, values=(item["filename"], item["brand"], item["sku"], status))

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
        edit_win.title("✏️ 快速修改客户与产品名")
        edit_win.geometry("400x220")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text=f"原始文件: {cur_item['filename']}", wraplength=360).pack(padx=15, pady=10, anchor="w")
        
        b_var = tk.StringVar(value=cur_item["brand"])
        s_var = tk.StringVar(value=cur_item["sku"])
        
        f_in = ttk.Frame(edit_win)
        f_in.pack(fill=tk.X, padx=15, pady=5)
        
        ttk.Label(f_in, text="归属客户/品牌:").grid(row=0, column=0, sticky="w", pady=5)
        b_entry = ttk.Combobox(f_in, textvariable=b_var, values=self.curated_brands, width=26)
        b_entry.grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(f_in, text="产品/SKU名称:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(f_in, textvariable=s_var, width=28).grid(row=1, column=1, sticky="w", pady=5)
        
        def save_edit():
            cur_item["brand"] = b_var.get().strip()
            cur_item["sku"] = s_var.get().strip()
            self.tree.item(item_id, values=(cur_item["filename"], cur_item["brand"], cur_item["sku"], "已确认"))
            edit_win.destroy()
            
        btn_s = ttk.Button(edit_win, text="保存修改 (Enter)", command=save_edit)
        btn_s.pack(pady=12)
        edit_win.bind("<Return>", lambda e: save_edit())

    def execute_organize(self):
        if not self.files_data:
            messagebox.showwarning("提示", "请先添加需要归档的 AI 文件！")
            return
            
        root_dir = self.root_dir_var.get().strip()
        if not root_dir or not os.path.exists(root_dir):
            try:
                os.makedirs(root_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建主工作盘路径: {root_dir}\n{e}")
                return
                
        self.cfg["root_dir"] = root_dir
        save_config(self.cfg)
        
        success_count = 0
        last_created_dir = ""
        
        subfolders = [
            "01_Design_平面原稿",
            "02_Textures_贴图资产",
            "03_3D_三维工程",
            "04_Renders_通道输出",
            "05_Delivery_最终交付"
        ]
        
        for item in self.files_data:
            brand = item["brand"].strip()
            sku = item["sku"].strip() if item["sku"].strip() else os.path.splitext(item["filename"])[0]
            
            if brand:
                proj_dir = os.path.join(root_dir, brand, sku)
            else:
                proj_dir = os.path.join(root_dir, sku)
                
            for sub in subfolders:
                os.makedirs(os.path.join(proj_dir, sub), exist_ok=True)
                
            dest_file_path = os.path.join(proj_dir, "01_Design_平面原稿", item["filename"])
            try:
                shutil.copy2(item["filepath"], dest_file_path)
                success_count += 1
                last_created_dir = proj_dir
            except Exception as e:
                print(f"Error copying {item['filepath']}: {e}")
                
        if success_count > 0:
            if last_created_dir and os.path.exists(last_created_dir):
                try:
                    os.startfile(last_created_dir)
                except Exception:
                    pass
                    
            messagebox.showinfo(
                "🎉 归档成功",
                f"成功创建并归档了 {success_count} 个标准包装项目！\n\n源文件已安全移入各自的 [01_Design_平面原稿] 中。"
            )
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