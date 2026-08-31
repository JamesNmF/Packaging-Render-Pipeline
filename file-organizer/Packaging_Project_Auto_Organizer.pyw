# -*- coding: utf-8 -*-
"""
微信 AI 文件一键智能归档与项目脚手架创建工具 (Packaging Project Auto-Organizer)
支持：
1. 拖拽一个或多个 .ai / .pdf / .psd / .zip 包装源文件
2. 智能解析 [品牌/客户] 与 [产品/SKU名称]，自动剔除 "改/最终版/刀模" 等杂乱后缀
3. 自动在工作盘创建标准五级目录
4. 自动将源文件转移到 01_Design_平面原稿/ 中
5. 支持右键发送到 (SendTo) 快捷调用
"""

import os
import sys
import re
import shutil
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# 默认工作根目录配置
DEFAULT_ROOT_DIR = "D:\\Projects"
if os.path.exists("E:\\zjc"):
    DEFAULT_ROOT_DIR = "E:\\zjc"
elif os.path.exists("E:\\Projects"):
    DEFAULT_ROOT_DIR = "E:\\Projects"
elif os.path.exists("E:\\"):
    DEFAULT_ROOT_DIR = "E:\\"

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_organizer_config.txt")

def load_saved_root_dir():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                d = f.read().strip()
                if os.path.exists(d):
                    return d
        except Exception:
            pass
    return DEFAULT_ROOT_DIR

def save_root_dir(path):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(path.strip())
    except Exception:
        pass


def clean_and_parse_filename(filepath):
    """
    智能解析文件名：
    例1：柏缇_蜜桃气泡水_包装刀模_改3.ai -> 品牌: 柏缇, SKU: 蜜桃气泡水
    例2：元气森林-白桃苏打水-最终版.ai   -> 品牌: 元气森林, SKU: 白桃苏打水
    例3：零食有鸣_香辣牛肉干_正稿.pdf    -> 品牌: 零食有鸣, SKU: 香辣牛肉干
    """
    filename = os.path.splitext(os.path.basename(filepath))[0]
    
    # 过滤杂乱关键词
    noise_patterns = [
        r'[-_ ]?(包装|刀模|展开图|正稿|定稿|完稿|原稿|印刷稿|平面|效果图)',
        r'[-_ ]?(修改版|修改|最新版|最终版|最终|定案|终版|初稿|打样|打样稿)',
        r'[-_ ]?(副本|\d{6,}|\d{4}年|\d{1,2}月\d{1,2}日)',
        r'[-_ ]?([vV]\d+(\.\d+)?|改\d*|版\d*)',
    ]
    
    cleaned = filename
    for p in noise_patterns:
        cleaned = re.sub(p, '', cleaned, flags=re.IGNORECASE)
        
    cleaned = cleaned.strip(" -_")
    
    # 按常见分隔符切分
    parts = re.split(r'[-_—\s+]+', cleaned)
    parts = [p.strip() for p in parts if p.strip()]
    
    if len(parts) >= 2:
        client_brand = parts[0]
        sku_name = "_".join(parts[1:])
    elif len(parts) == 1:
        client_brand = ""
        sku_name = parts[0]
    else:
        client_brand = ""
        sku_name = filename
        
    return client_brand, sku_name


class OrganizerApp:
    def __init__(self, root, initial_files=None):
        self.root = root
        self.root.title("📦 包装 AI 文件智能归档与项目脚手架创建器")
        self.root.geometry("720x520")
        self.root.minsize(640, 420)
        
        self.root_dir_var = tk.StringVar(value=load_saved_root_dir())
        self.files_data = [] # list of dicts: {filepath, brand, sku}
        
        self.build_ui()
        
        if initial_files:
            self.add_files(initial_files)

    def build_ui(self):
        # 1. 顶部工作盘设置
        top_frame = ttk.LabelFrame(self.root, text=" 📂 项目主工作盘 / 根目录 ", padding=10)
        top_frame.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Entry(top_frame, textvariable=self.root_dir_var, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(top_frame, text="浏览选择...", command=self.browse_root_dir).pack(side=tk.RIGHT)
        
        # 2. 中间文件列表与智能解析表格
        list_frame = ttk.LabelFrame(self.root, text=" 📋 待处理的微信源文件列表 (支持直接拖拽/点击添加) ", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))
        
        cols = ("file", "brand", "sku")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("file", text="微信接收的文件名")
        self.tree.heading("brand", text="识别出的【客户/品牌】")
        self.tree.heading("sku", text="识别出的【产品/SKU名】")
        
        self.tree.column("file", width=280, anchor="w")
        self.tree.column("brand", width=140, anchor="center")
        self.tree.column("sku", width=200, anchor="w")
        
        scroll_y = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 双击可微调品牌或SKU
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # 3. 操作按钮栏
        btn_frame = ttk.Frame(self.root, padding=5)
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        ttk.Button(btn_frame, text="➕ 添加 AI / 源文件...", command=self.browse_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="❌ 移除选中项", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="🧹 清空列表", command=self.clear_all).pack(side=tk.LEFT)
        
        # 4. 底部执行大按钮
        exec_frame = ttk.Frame(self.root, padding=10)
        exec_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        self.btn_exec = tk.Button(
            exec_frame,
            text="🚀 【 一键创建工业级目录并归档源文件 】",
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

    def browse_root_dir(self):
        d = filedialog.askdirectory(initialdir=self.root_dir_var.get())
        if d:
            self.root_dir_var.set(d)
            save_root_dir(d)

    def browse_files(self):
        files = filedialog.askopenfilenames(
            title="选择微信接收的 AI / 包装设计文件",
            filetypes=[("包装设计文件", "*.ai;*.pdf;*.psd;*.zip;*.rar;*.eps;*.cdr"), ("所有文件", "*.*")]
        )
        if files:
            self.add_files(files)

    def add_files(self, filepaths):
        for fp in filepaths:
            fp = os.path.abspath(fp)
            if not os.path.exists(fp) or not os.path.isfile(fp):
                continue
            if any(item['filepath'] == fp for item in self.files_data):
                continue
                
            brand, sku = clean_and_parse_filename(fp)
            item = {
                "filepath": fp,
                "filename": os.path.basename(fp),
                "brand": brand,
                "sku": sku
            }
            self.files_data.append(item)
            self.tree.insert("", tk.END, values=(item["filename"], item["brand"], item["sku"]))

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
        
        # 弹出微调小窗口
        edit_win = tk.Toplevel(self.root)
        edit_win.title("✏️ 微调产品与品牌名称")
        edit_win.geometry("380x200")
        edit_win.transient(self.root)
        edit_win.grab_set()
        
        ttk.Label(edit_win, text=f"文件名: {cur_item['filename']}", wraplength=340).pack(padx=15, pady=10, anchor="w")
        
        b_var = tk.StringVar(value=cur_item["brand"])
        s_var = tk.StringVar(value=cur_item["sku"])
        
        f_in = ttk.Frame(edit_win)
        f_in.pack(fill=tk.X, padx=15, pady=5)
        
        ttk.Label(f_in, text="客户/品牌:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(f_in, textvariable=b_var, width=28).grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(f_in, text="产品/SKU:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(f_in, textvariable=s_var, width=28).grid(row=1, column=1, sticky="w", pady=5)
        
        def save_edit():
            cur_item["brand"] = b_var.get().strip()
            cur_item["sku"] = s_var.get().strip()
            self.tree.item(item_id, values=(cur_item["filename"], cur_item["brand"], cur_item["sku"]))
            edit_win.destroy()
            
        ttk.Button(edit_win, text="保存修改", command=save_edit).pack(pady=12)

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
                
        save_root_dir(root_dir)
        
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
            brand = item["brand"]
            sku = item["sku"] if item["sku"] else os.path.splitext(item["filename"])[0]
            
            if brand:
                proj_dir = os.path.join(root_dir, brand, sku)
            else:
                proj_dir = os.path.join(root_dir, sku)
                
            # 1. 创建五级子目录
            for sub in subfolders:
                os.makedirs(os.path.join(proj_dir, sub), exist_ok=True)
                
            # 2. 拷贝/转移源文件至 01_Design_平面原稿
            dest_file_path = os.path.join(proj_dir, "01_Design_平面原稿", item["filename"])
            try:
                shutil.copy2(item["filepath"], dest_file_path)
                success_count += 1
                last_created_dir = proj_dir
            except Exception as e:
                print(f"Error copying {item['filepath']}: {e}")
                
        # 3. 结果反馈与自动弹出
        if success_count > 0:
            if last_created_dir and os.path.exists(last_created_dir):
                try:
                    os.startfile(last_created_dir)
                except Exception:
                    pass
                    
            messagebox.showinfo(
                "🎉 归档成功",
                f"成功创建并归档了 {success_count} 个包装项目！\n\n所有源文件已自动安全分类至各自的 [01_Design_平面原稿] 中。"
            )
            self.clear_all()


if __name__ == "__main__":
    args_files = sys.argv[1:] if len(sys.argv) > 1 else None
    root_win = tk.Tk()
    
    # 设置 Windows 原生高 DPI 适配
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
        
    app = OrganizerApp(root_win, initial_files=args_files)
    root_win.mainloop()