# -*- coding: utf-8 -*-
"""
🎨 美术资产中枢 · Blender 包装贴图直达助手 (Quick Texture Jumper for Blender 5.2+)
===================================================================================
【功能说明】：
1. 自动寻址：打开 .blend 时自动向上查找并锁定项目的「02_Textures_贴图资产」目录；
2. 书签置顶：自动将贴图目录写入 Blender 文件浏览器的置顶书签；
3. 节点直达：在 Shader 编辑器的「图像纹理」节点增加 [📂 打开项目贴图] 按钮与快捷键 (Alt+O)；
4. 零侵入性：完全不修改任何材质连线和节点结构，纯粹作为路径辅助。
===================================================================================
"""

bl_info = {
    "name": "美术资产中枢 - 贴图直达助手",
    "author": "Art Asset Hub",
    "version": (1, 0, 0),
    "blender": (5, 0, 0),
    "location": "Shader Editor > Image Texture Node & Alt+O",
    "description": "打开贴图时自动定位至当前工程同级的 02_Textures 贴图目录",
    "category": "Material",
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
    
    # 1. 尝试在父级目录寻找 (标准工程: 03_3D_三维工程/.. ➔ 02_Textures_贴图资产)
    for name in TEXTURE_CANDIDATES:
        cand = os.path.join(parent_dir, name)
        if os.path.exists(cand):
            return os.path.normpath(cand)
            
    # 2. 尝试在同级目录寻找
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
        
        # 整理书签
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
                    
        # 将当前贴图目录置顶
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


# ---------------- 节点直达选择操作符 ----------------
class NODE_OT_open_project_texture(bpy.types.Operator, ImportHelper):
    """直接打开当前项目的 02_Textures 贴图目录选择图片"""
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
            
        node = context.active_node
        if node and node.type == 'TEX_IMAGE':
            try:
                # 载入或获取图片数据块
                img = bpy.data.images.load(self.filepath, check_existing=True)
                node.image = img
                self.report({'INFO'}, f"✅ 成功载入贴图: {os.path.basename(self.filepath)}")
                return {'FINISHED'}
            except Exception as e:
                self.report({'ERROR'}, f"载入贴图失败: {e}")
                return {'CANCELLED'}
        else:
            self.report({'WARNING'}, "请先在着色器编辑器中选中一个「图像纹理 (Image Texture)」节点！")
            return {'CANCELLED'}


# 在 Shader 节点的 UI 上绘制直达按钮
def draw_image_texture_node_button(self, context):
    if self.type == 'TEX_IMAGE':
        layout = self.layout
        row = layout.row(align=True)
        row.operator("node.open_project_texture", text="📂 贴图目录 (Alt+O)", icon='FILE_FOLDER')


# ---------------- 注册与热键绑定 ----------------
classes = (
    NODE_OT_open_project_texture,
)

addon_keymaps = []

def register():
    for cls in classes:
        bpy.utils.register_class(cls)
        
    bpy.app.handlers.load_post.append(on_load_blend_post)
    bpy.types.NODE_HT_header.append(lambda s, c: None)
    bpy.types.ShaderNodeTexImage.draw_buttons_ext = draw_image_texture_node_button

    # 绑定快捷键 Alt + O (在 Shader Node Editor 中生效)
    wm = bpy.context.window_manager
    kc = wm.keyconfigs.addon
    if kc:
        km = kc.keymaps.new(name='Node Editor', space_type='NODE_EDITOR')
        kmi = km.keymap_items.new("node.open_project_texture", 'O', 'PRESS', alt=True)
        addon_keymaps.append((km, kmi))

def unregister():
    for km, kmi in addon_keymaps:
        km.keymap_items.remove(kmi)
    addon_keymaps.clear()
    
    if on_load_blend_post in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.remove(on_load_blend_post)
        
    for cls in reversed(classes):
        bpy.utils.unregister_class(cls)

if __name__ == "__main__":
    register()
