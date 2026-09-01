# -*- coding: utf-8 -*-
"""
🎨 美术资产中枢 · Blender 包装贴图直达助手 (Quick Texture Jumper for Blender 5.2+)
===================================================================================
【核心功能】：
1. 顶部 Header 栏常驻按钮：Shader 编辑器顶部直接显示 [📂 项目贴图 (Alt+O)]
2. 右键菜单集成：在 Shader 视口右键任意节点 ➔ [📂 打开项目贴图目录]
3. N 键侧边栏专用面板：按 N 呼出侧边栏 ➔ [🎨 贴图直达] 选项卡
4. 双快捷键支持：支持 Alt + O (字母) 与 Alt + 0 (数字) 一键唤起
5. 老工程 & 新工程自适应探测 (02_Textures_贴图资产 / png / PNG / 贴图)
===================================================================================
"""

bl_info = {
    "name": "美术资产中枢 - 贴图直达助手",
    "author": "Art Asset Hub",
    "version": (1, 0, 1),
    "blender": (5, 0, 0),
    "location": "Shader Editor > Top Header, Right Click & N-Panel",
    "description": "打开贴图时自动定位至当前工程同级的 02_Textures 或 png 贴图目录",
    "category": "Node",
}

import os
import bpy
from bpy.app.handlers import persistent
from bpy_extras.io_utils import ImportHelper

TEXTURE_CANDIDATES = [
    "02_Textures_贴图资产",
    "02_Textures",
    "png",
    "PNG",
    "02_贴图资产",
    "Textures",
    "贴图",
    "02_贴图"
]

def get_current_project_texture_dir():
    blend_path = bpy.data.filepath
    if not blend_path:
        return ""
    
    current_dir = os.path.dirname(blend_path)
    parent_dir = os.path.dirname(current_dir)
    
    # 1. 优先在父级目录寻找 (新工程: 03_3D_三维工程/.. ➔ 02_Textures; 老工程: 模型/.. ➔ png)
    for name in TEXTURE_CANDIDATES:
        cand = os.path.join(parent_dir, name)
        if os.path.exists(cand):
            return os.path.normpath(cand)
            
    # 2. 尝试在当前同级目录寻找
    for name in TEXTURE_CANDIDATES:
        cand = os.path.join(current_dir, name)
        if os.path.exists(cand):
            return os.path.normpath(cand)
            
    # 3. 兜底返回父目录下的标准 02_Textures 目录
    default_dir = os.path.join(parent_dir, "02_Textures_贴图资产")
    return os.path.normpath(default_dir)


def update_blender_bookmarks(tex_dir):
    if not tex_dir or not os.path.exists(tex_dir):
        return
    try:
        config_dir = bpy.utils.user_resource('CONFIG')
        bm_file = os.path.join(config_dir, "bookmarks.txt")
        
        lines = []
        if os.path.exists(bm_file):
            with open(bm_file, "r", encoding="utf-8", errors="ignore") as f:
                lines = [l.strip() for l in f.readlines() if l.strip()]
                
        norm_target = os.path.normpath(tex_dir).rstrip("\\/") + "\\"
        
        bookmarks_sec = []
        recent_sec = []
        cur_sec = None
        
        for l in lines:
            if l == "[Bookmarks]":
                cur_sec = "bm"
                continue
            elif l == "[Recent]":
                cur_sec = "rc"
                continue
            if cur_sec == "bm":
                if l != norm_target:
                    bookmarks_sec.append(l)
            elif cur_sec == "rc":
                if l != norm_target:
                    recent_sec.append(l)
                    
        bookmarks_sec.insert(0, norm_target)
        recent_sec.insert(0, norm_target)
        
        out_content = "[Bookmarks]\n" + "\n".join(bookmarks_sec[:20]) + "\n\n[Recent]\n" + "\n".join(recent_sec[:20]) + "\n"
        with open(bm_file, "w", encoding="utf-8") as f:
            f.write(out_content)
    except Exception as e:
        print(f"[ArtAssetHub] Error updating bookmarks: {e}")


@persistent
def on_load_blend_post(dummy):
    tex_dir = get_current_project_texture_dir()
    if tex_dir:
        update_blender_bookmarks(tex_dir)


# ---------------- 核心操作符 ----------------
class NODE_OT_open_project_texture(bpy.types.Operator, ImportHelper):
    """直接打开当前项目的贴图目录 (02_Textures / png) 选择图片并赋给选中节点"""
    bl_idname = "node.open_project_texture"
    bl_label = "📂 打开项目贴图目录"
    bl_options = {'REGISTER', 'UNDO'}

    filename_ext = ".png;.jpg;.jpeg;.psd;.tif;.tiff;.exr;.hdr"
    filter_glob: bpy.props.StringProperty(
        default="*.png;*.jpg;*.jpeg;*.psd;*.tif;*.tiff;*.exr;*.hdr",
        options={'HIDDEN'}
    )

    def invoke(self, context, event):
        tex_dir = get_current_project_texture_dir()
        if tex_dir and os.path.exists(tex_dir):
            self.directory = os.path.normpath(tex_dir) + os.sep
            update_blender_bookmarks(tex_dir)
        return ImportHelper.invoke(self, context, event)

    def execute(self, context):
        if not self.filepath:
            return {'CANCELLED'}
            
        node = getattr(context, 'active_node', None)
        if node and node.type == 'TEX_IMAGE':
            try:
                img = bpy.data.images.load(self.filepath, check_existing=True)
                node.image = img
                self.report({'INFO'}, f"✅ 成功载入贴图: {os.path.basename(self.filepath)}")
                return {'FINISHED'}
            except Exception as e:
                self.report({'ERROR'}, f"载入贴图失败: {e}")
                return {'CANCELLED'}
        else:
            # 如果当前没有选中 Image Texture 节点，仍然打开图片并载入到 Blender 数据中
            try:
                img = bpy.data.images.load(self.filepath, check_existing=True)
                self.report({'INFO'}, f"✅ 贴图已载入 Blender: {os.path.basename(self.filepath)}")
                return {'FINISHED'}
            except Exception as e:
                self.report({'ERROR'}, str(e))
                return {'CANCELLED'}


class NODE_OT_explore_texture_folder(bpy.types.Operator):
    """在 Windows 文件管理器中直接打开当前项目的贴图文件夹"""
    bl_idname = "node.explore_texture_folder"
    bl_label = "📁 在资源管理器中打开贴图文件夹"
    bl_options = {'REGISTER'}

    def execute(self, context):
        tex_dir = get_current_project_texture_dir()
        if tex_dir and os.path.exists(tex_dir):
            os.startfile(tex_dir)
            return {'FINISHED'}
        else:
            self.report({'WARNING'}, f"未找到贴图目录: {tex_dir}")
            return {'CANCELLED'}


# ---------------- 界面挂载点 ----------------

# 1. Shader 编辑器顶部 Header 栏直接显示常驻按钮
def draw_shader_header(self, context):
    if context.space_data.tree_type == 'ShaderNodeTree':
        layout = self.layout
        row = layout.row(align=True)
        tex_dir = get_current_project_texture_dir()
        folder_name = os.path.basename(tex_dir) if tex_dir else "贴图"
        row.operator("node.open_project_texture", text=f"📂 贴图目录: {folder_name} (Alt+O)", icon='IMAGE_DATA')


# 2. Shader 编辑器右键上下文菜单
def draw_shader_context_menu(self, context):
    layout = self.layout
    layout.separator()
    layout.operator("node.open_project_texture", text="📂 从项目贴图目录选择 (Alt+O)", icon='FILE_FOLDER')
    layout.operator("node.explore_texture_folder", text="📁 打开贴图文件夹 (Explorer)", icon='FOLDER_REDIRECT')


# 3. Shader 编辑器 N 键侧边栏专用面板
class NODE_PT_art_asset_texture_panel(bpy.types.Panel):
    bl_label = "🎨 美术资产中枢"
    bl_idname = "NODE_PT_art_asset_texture_panel"
    bl_space_type = 'NODE_EDITOR'
    bl_region_type = 'UI'
    bl_category = "🎨 贴图直达"

    def draw(self, context):
        layout = self.layout
        tex_dir = get_current_project_texture_dir()
        
        box = layout.box()
        box.label(text="当前项目贴图目录:", icon='FILE_FOLDER')
        if tex_dir and os.path.exists(tex_dir):
            box.label(text=f"📁 {os.path.basename(tex_dir)}")
            box.label(text=tex_dir)
            col = layout.column(align=True)
            col.operator("node.open_project_texture", text="🖼️ 选择贴图赋给选中节点 (Alt+O)", icon='IMAGE_DATA')
            col.operator("node.explore_texture_folder", text="📂 打开贴图文件夹", icon='FOLDER_REDIRECT')
        else:
            box.label(text="⚠️ 未找到关联贴图目录", icon='ERROR')
            if bpy.data.filepath:
                box.label(text="请确保同级有 02_Textures 或 png 文件夹")


# ---------------- 注册与热键绑定 ----------------
classes = (
    NODE_OT_open_project_texture,
    NODE_OT_explore_texture_folder,
    NODE_PT_art_asset_texture_panel,
)

addon_keymaps = []

def register():
    for cls in classes:
        bpy.utils.register_class(cls)
        
    bpy.app.handlers.load_post.append(on_load_blend_post)
    bpy.types.NODE_HT_header.append(draw_shader_header)
    bpy.types.NODE_MT_context_menu.append(draw_shader_context_menu)

    # 绑定快捷键 Alt + O (字母 O) 与 Alt + 0 (数字 0)
    wm = bpy.context.window_manager
    kc = wm.keyconfigs.addon
    if kc:
        km = kc.keymaps.new(name='Node Editor', space_type='NODE_EDITOR')
        # 字母 O
        kmi1 = km.keymap_items.new("node.open_project_texture", 'O', 'PRESS', alt=True)
        addon_keymaps.append((km, kmi1))
        # 数字 0 (防止用户按错)
        kmi2 = km.keymap_items.new("node.open_project_texture", 'ZERO', 'PRESS', alt=True)
        addon_keymaps.append((km, kmi2))

def unregister():
    for km, kmi in addon_keymaps:
        km.keymap_items.remove(kmi)
    addon_keymaps.clear()
    
    if on_load_blend_post in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.remove(on_load_blend_post)
        
    bpy.types.NODE_HT_header.remove(draw_shader_header)
    bpy.types.NODE_MT_context_menu.remove(draw_shader_context_menu)
    
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)

if __name__ == "__main__":
    register()
