# -*- coding: utf-8 -*-
"""
包装视觉资产管理器 (Packaging Visual Asset Hub)
专为设计师打造的 Eagle / Pinterest 风格三维渲染与项目资产看板：
1. 视觉化缩略图卡片流：自动抓取各项目 04_Renders / 05_Delivery 中的高清渲染图作为封面！
2. 品牌分类与即时搜索：左侧品牌一键筛选，顶部输入即时过滤。
3. 一键直达交互：
   - 单击卡片/按钮 -> 0.1秒弹出 Windows 对应项目文件夹！
   - 双击卡片 -> 自动启动 Blender 5.2 打开该产品的 3D 工程！
   - 右键菜单 -> 打开平面原稿 .ai / 查看高清渲染大图 / 复制路径。
4. 支持导出轻量独立 HTML 全景视觉看板。
"""

import os
import sys
import json
import glob
import webbrowser
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".packaging_organizer_v6.json")
CACHE_DIR = os.path.join(os.path.expanduser("~"), ".packaging_asset_thumbnails")
os.makedirs(CACHE_DIR, exist_ok=True)

BLENDER_EXE = r"C:\Program Files\Blender Foundation\Blender 5.2\blender.exe"
if not os.path.exists(BLENDER_EXE):
    BLENDER_EXE = "blender"

DEFAULT_WORKSPACES = []
for p in ["E:\\zjc", "D:\\Projects", "E:\\Projects", "D:\\", "E:\\"]:
    if os.path.exists(p) and p not in DEFAULT_WORKSPACES:
        DEFAULT_WORKSPACES.append(p)
if not DEFAULT_WORKSPACES:
    DEFAULT_WORKSPACES = ["D:\\Projects"]

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
    return {
        "workspaces": DEFAULT_WORKSPACES,
        "current_workspace": DEFAULT_WORKSPACES[0]
    }


def find_project_thumbnail(proj_path):
    """查找项目中最合适的高清渲染图或预览图作为缩略图"""
    # 1. 优先在 04_Renders_通道输出 / 05_Delivery_最终交付 中查找 Beauty 成品渲染图
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
            for ext in ("*.png", "*.jpg", "*.jpeg", "*.webp"):
                candidates.extend(glob.glob(os.path.join(rdir, ext)))
                
    if candidates:
        # 优先选择包含 "Beauty", "成品", "主图", "Camera" 的图
        beauty_imgs = [c for c in candidates if any(k in os.path.basename(c).lower() for k in ["beauty", "成品", "主图", "camera", "正面"])]
        if beauty_imgs:
            # 取最新修改的图
            beauty_imgs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return beauty_imgs[0]
        # 排除包含 mask, alpha, crypto 的黑白遮罩图
        filtered = [c for c in candidates if not any(k in os.path.basename(c).lower() for k in ["mask", "alpha", "crypto", "选区", "蒙版"])]
        if filtered:
            filtered.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return filtered[0]
        candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        return candidates[0]
        
    return None


def scan_workspace_projects(root_dir):
    """扫描工作盘下的所有包装三维与设计项目"""
    projects = []
    if not os.path.exists(root_dir):
        return projects
        
    try:
        entries = os.listdir(root_dir)
    except Exception:
        return projects
        
    for entry in entries:
        full_p = os.path.join(root_dir, entry)
        if not os.path.isdir(full_p) or entry.startswith('.') or entry.startswith('$') or entry.startswith('_'):
            continue
            
        # 检查这是否直接是一个项目 (含 01_Design, 03_3D 等)
        sub_items = [d.lower() for d in os.listdir(full_p) if os.path.isdir(os.path.join(full_p, d))]
        is_direct_proj = any("03_3d" in s or "01_design" in s or "04_renders" in s for s in sub_items)
        
        if is_direct_proj:
            thumb = find_project_thumbnail(full_p)
            projects.append({
                "brand": "通用/未分类",
                "sku": entry,
                "path": full_p,
                "thumbnail": thumb,
                "mtime": os.path.getmtime(full_p)
            })
        else:
            # 这是一个客户/品牌文件夹 (如 E:\zjc\柏缇\)，扫描其二级目录
            brand_name = entry
            try:
                sku_entries = os.listdir(full_p)
                for sku in sku_entries:
                    sku_p = os.path.join(full_p, sku)
                    if os.path.isdir(sku_p) and not sku.startswith('.') and not sku.startswith('_'):
                        thumb = find_project_thumbnail(sku_p)
                        projects.append({
                            "brand": brand_name,
                            "sku": sku,
                            "path": sku_p,
                            "thumbnail": thumb,
                            "mtime": os.path.getmtime(sku_p)
                        })
            except Exception:
                pass
                
    projects.sort(key=lambda x: x["mtime"], reverse=True)
    return projects


class AssetHubApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📦 包装视觉资产管理器 (Packaging Visual Asset Hub)")
        self.root.geometry("1120x740")
        self.root.minsize(920, 600)
        
        self.cfg = load_config()
        self.workspaces = self.cfg.get("workspaces", DEFAULT_WORKSPACES)
        self.current_workspace_var = tk.StringVar(value=self.cfg.get("current_workspace", self.workspaces[0]))
        self.search_var = tk.StringVar()
        self.selected_brand_var = tk.StringVar(value="全部品牌")
        
        self.all_projects = []
        self.filtered_projects = []
        self.thumb_cache = {} # path -> ImageTk.PhotoImage
        
        self.build_ui()
        self.refresh_projects()

    def build_ui(self):
        # 1. 顶部操作栏
        top_bar = ttk.Frame(self.root, padding=(15, 10))
        top_bar.pack(fill=tk.X)
        
        ttk.Label(top_bar, text="📂 工作盘:").pack(side=tk.LEFT, padx=(0, 6))
        self.ws_combo = ttk.Combobox(
            top_bar,
            textvariable=self.current_workspace_var,
            values=self.workspaces,
            state="readonly",
            width=28,
            font=("Microsoft YaHei", 9)
        )
        self.ws_combo.pack(side=tk.LEFT, padx=(0, 15))
        self.ws_combo.bind("<<ComboboxSelected>>", self.on_workspace_change)
        
        ttk.Label(top_bar, text="🔍 快速搜索:").pack(side=tk.LEFT, padx=(0, 6))
        search_entry = ttk.Entry(top_bar, textvariable=self.search_var, font=("Microsoft YaHei", 9), width=24)
        search_entry.pack(side=tk.LEFT, padx=(0, 15))
        self.search_var.trace_add("write", lambda *args: self.apply_filter())
        
        ttk.Button(top_bar, text="🔄 刷新", command=self.refresh_projects).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top_bar, text="🌐 导出全景网页看板 (HTML)...", command=self.export_html_gallery).pack(side=tk.RIGHT)

        # 2. 主体区域 (左侧品牌树 + 右侧卡片瀑布流)
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 12))
        
        # 左侧品牌筛选栏
        left_frame = ttk.LabelFrame(main_pane, text=" 🏷️ 客户/品牌分类 ", padding=8, width=200)
        main_pane.add(left_frame, weight=1)
        
        self.brand_listbox = tk.Listbox(
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
        self.brand_listbox.pack(fill=tk.BOTH, expand=True)
        self.brand_listbox.bind("<<ListboxSelect>>", self.on_brand_select)
        
        # 右侧卡片滚动视口
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

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.render_cards()

    def on_mouse_wheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def on_workspace_change(self, event=None):
        cur_ws = self.current_workspace_var.get().strip()
        self.cfg["current_workspace"] = cur_ws
        self.refresh_projects()

    def refresh_projects(self):
        cur_ws = self.current_workspace_var.get().strip()
        self.all_projects = scan_workspace_projects(cur_ws)
        
        # 更新左侧品牌列表与计数
        brand_counts = {}
        for p in self.all_projects:
            b = p["brand"]
            brand_counts[b] = brand_counts.get(b, 0) + 1
            
        self.brand_listbox.delete(0, tk.END)
        self.brand_listbox.insert(tk.END, f"✨ 全部项目 ({len(self.all_projects)})")
        
        brands_sorted = sorted(brand_counts.keys())
        for b in brands_sorted:
            self.brand_listbox.insert(tk.END, f"🏷️ {b} ({brand_counts[b]})")
            
        self.brand_listbox.select_set(0)
        self.selected_brand_var.set("全部品牌")
        self.apply_filter()

    def on_brand_select(self, event=None):
        sel = self.brand_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        text = self.brand_listbox.get(idx)
        if idx == 0:
            self.selected_brand_var.set("全部品牌")
        else:
            # 提取品牌名 "🏷️ 柏缇 (8)" -> "柏缇"
            m = re.search(r"🏷️\s*(.*?)\s*\(\d+\)", text)
            if m:
                self.selected_brand_var.set(m.group(1))
        self.apply_filter()

    def apply_filter(self):
        brand = self.selected_brand_var.get()
        kw = self.search_var.get().strip().lower()
        
        self.filtered_projects = []
        for p in self.all_projects:
            if brand != "全部品牌" and p["brand"] != brand:
                continue
            if kw and (kw not in p["sku"].lower() and kw not in p["brand"].lower()):
                continue
            self.filtered_projects.append(p)
            
        self.render_cards()

    def get_scaled_thumbnail(self, img_path, size=(190, 190)):
        if not img_path or not os.path.exists(img_path):
            return self.get_placeholder_thumbnail(size)
            
        if img_path in self.thumb_cache:
            return self.thumb_cache[img_path]
            
        try:
            im = Image.open(img_path)
            im.thumbnail(size, Image.Resampling.LANCZOS)
            
            # 创建居中正方形画板
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
        draw.text((size[0]//2 - 40, size[1]//2 - 10), "📦 暂无渲染图", fill=(160, 170, 185, 255))
        tk_img = ImageTk.PhotoImage(im)
        self.thumb_cache[cache_key] = tk_img
        return tk_img

    def render_cards(self):
        for widget in self.grid_container.winfo_children():
            widget.destroy()
            
        if not self.filtered_projects:
            no_lbl = tk.Label(
                self.grid_container,
                text="📭 没有找到匹配的包装项目\n(提示：可通过顶部切换工作盘或点击刷新)",
                font=("Microsoft YaHei", 12),
                bg="#F0F2F5",
                fg="#888",
                pady=60
            )
            no_lbl.pack()
            return

        # 计算每行卡片数
        container_width = self.canvas.winfo_width()
        if container_width < 100:
            container_width = 800
        card_w = 210
        cols = max(1, container_width // (card_w + 16))

        for idx, proj in enumerate(self.filtered_projects):
            row = idx // cols
            col = idx % cols
            
            card = tk.Frame(
                self.grid_container,
                bg="white",
                bd=1,
                relief=tk.SOLID,
                padx=8,
                pady=8,
                highlightthickness=0
            )
            card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            
            # 缩略图图片
            tk_thumb = self.get_scaled_thumbnail(proj["thumbnail"])
            img_lbl = tk.Label(card, image=tk_thumb, bg="white", cursor="hand2")
            img_lbl.image = tk_thumb
            img_lbl.pack(fill=tk.BOTH, expand=True)
            
            # 品牌小胶囊标签 + 产品名
            meta_frame = tk.Frame(card, bg="white", pady=4)
            meta_frame.pack(fill=tk.X)
            
            b_tag = tk.Label(
                meta_frame,
                text=proj["brand"],
                font=("Microsoft YaHei", 8),
                bg="#E1EFFF",
                fg="#005A9E",
                padx=4,
                pady=1
            )
            b_tag.pack(anchor="w", pady=(0, 2))
            
            title_lbl = tk.Label(
                meta_frame,
                text=proj["sku"],
                font=("Microsoft YaHei", 9, "bold"),
                bg="white",
                fg="#222",
                wraplength=180,
                justify="left"
            )
            title_lbl.pack(anchor="w")
            
            # 底部动作按钮行
            action_frame = tk.Frame(card, bg="white", pady=4)
            action_frame.pack(fill=tk.X)
            
            btn_open = tk.Button(
                action_frame,
                text="📁 打开文件夹",
                font=("Microsoft YaHei", 8),
                bg="#F0F0F0",
                relief=tk.FLAT,
                command=lambda p=proj["path"]: self.open_folder(p)
            )
            btn_open.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
            
            btn_blend = tk.Button(
                action_frame,
                text="🚀 3D工程",
                font=("Microsoft YaHei", 8),
                bg="#EBF5FB",
                fg="#005A9E",
                relief=tk.FLAT,
                command=lambda p=proj["path"]: self.launch_blend(p)
            )
            btn_blend.pack(side=tk.RIGHT)
            
            # 绑定卡片交互：单击打开文件夹，双击打开 Blender，右键弹出菜单
            for w in (card, img_lbl, title_lbl, meta_frame):
                w.bind("<Button-1>", lambda e, p=proj["path"]: self.open_folder(p))
                w.bind("<Double-1>", lambda e, p=proj["path"]: self.launch_blend(p))
                w.bind("<Button-3>", lambda e, pr=proj: self.show_context_menu(e, pr))

    def open_folder(self, path):
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("打开错误", str(e))

    def launch_blend(self, proj_path):
        """寻找该项目下的 .blend 文件并用 Blender 启动"""
        blend_dir = os.path.join(proj_path, "03_3D_三维工程")
        target_dir = blend_dir if os.path.exists(blend_dir) else proj_path
        
        blends = glob.glob(os.path.join(target_dir, "*.blend"))
        if blends:
            # 优先选择主工程
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
        menu.add_command(label=f"📁 打开项目文件夹: {proj['sku']}", command=lambda: self.open_folder(proj["path"]))
        menu.add_command(label="🚀 启动 Blender 打开 3D 工程", command=lambda: self.launch_blend(proj["path"]))
        menu.add_command(label="🎨 打开 01_Design 平面原稿文件夹", command=lambda: self.open_folder(os.path.join(proj["path"], "01_Design_平面原稿")))
        menu.add_command(label="🖼️ 查看 04_Renders 渲染大图文件夹", command=lambda: self.open_folder(os.path.join(proj["path"], "04_Renders_通道输出")))
        menu.add_separator()
        menu.add_command(label="📋 复制项目完整路径", command=lambda: self.copy_path_to_clipboard(proj["path"]))
        menu.tk_popup(event.x_root, event.y_root)

    def copy_path_to_clipboard(self, text):
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制路径到剪贴板:\n{text}")

    def export_html_gallery(self):
        """导出精美独立的本地 HTML 视觉全景画廊"""
        cur_ws = self.current_workspace_var.get().strip()
        html_file = os.path.join(cur_ws, "📦_包装项目全景视觉画廊.html")
        
        cards_html = []
        for p in self.all_projects:
            thumb_rel = p["thumbnail"] if p["thumbnail"] else ""
            thumb_src = "file:///" + thumb_rel.replace("\\", "/") if thumb_rel else "https://via.placeholder.com/300x300?text=No+Render"
            folder_uri = "file:///" + p["path"].replace("\\", "/")
            
            cards_html.append(f"""
            <div class="card" onclick="window.open('{folder_uri}')">
                <div class="thumb-container">
                    <img src="{thumb_src}" alt="{p['sku']}" loading="lazy">
                </div>
                <div class="meta">
                    <span class="badge">{p['brand']}</span>
                    <h3 class="title">{p['sku']}</h3>
                    <p class="path">{p['path']}</p>
                </div>
            </div>
            """)
            
        full_html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>📦 包装项目全景视觉画廊</title>
<style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background: #0f172a; color: #f8fafc; margin: 0; padding: 24px; }}
    h1 {{ font-size: 24px; font-weight: 700; margin-bottom: 24px; display: flex; align-items: center; gap: 10px; }}
    .grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 20px; }}
    .card {{ background: #1e293b; border-radius: 12px; overflow: hidden; border: 1px solid #334155; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; }}
    .card:hover {{ transform: translateY(-4px); box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5); border-color: #38bdf8; }}
    .thumb-container {{ width: 100%; aspect-ratio: 1; background: #0f172a; display: flex; align-items: center; justify-content: center; }}
    .thumb-container img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
    .meta {{ padding: 12px; }}
    .badge {{ display: inline-block; background: #0369a1; color: #e0f2fe; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 9999px; margin-bottom: 6px; }}
    .title {{ font-size: 14px; font-weight: 600; margin: 0 0 4px 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
    .path {{ font-size: 11px; color: #94a3b8; margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
</style>
</head>
<body>
    <h1>📦 包装项目全景视觉资产画廊 (共 {len(self.all_projects)} 个项目)</h1>
    <div class="grid">
        {"".join(cards_html)}
    </div>
</body>
</html>"""

        try:
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(full_html)
            webbrowser.open("file:///" + html_file.replace("\\", "/"))
        except Exception as e:
            messagebox.showerror("导出错误", str(e))


if __name__ == "__main__":
    root_win = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    app = AssetHubApp(root_win)
    root_win.mainloop()