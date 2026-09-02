bl_info = {
    "name": "包装渲染与多通道自动输出助手 (Packaging Render Pipeline)",
    "author": "Antigravity Pipeline",
    "version": (2, 1, 0),
    "blender": (5, 2, 0),
    "location": "3D 视口 > 侧边栏 (N) > 包装渲染",
    "description": "专为包装渲染定制：内置标准项目脚手架创建、智能产品目录识别、非阻塞多机位批量连拍、电商画幅/采样率快切、一键阴影捕捉与多通道输出",
    "category": "Render",
}

import bpy
import os
import re
import subprocess
from bpy.props import StringProperty, BoolProperty, EnumProperty, PointerProperty
from bpy.types import Panel, Operator, PropertyGroup
from bpy.app.handlers import persistent

class PackagingPipelineProperties(PropertyGroup):
    # 0. 标准项目脚手架创建
    scaffold_root_dir: StringProperty(
        name="工作根目录",
        description="选择或输入存放所有项目的主工作盘/根目录 (如 E:/Projects 或 D:/Projects)",
        subtype='DIR_PATH',
        default="D:\\Projects\\"
    )
    scaffold_client_name: StringProperty(
        name="品牌/客户名",
        description="品牌或客户名称 (如：Brand_A 或 柏缇，可留空)",
        default=""
    )
    scaffold_project_name: StringProperty(
        name="产品/SKU名",
        description="产品或SKU名称 (如：Peach_Soda 或 蜜桃气泡水)",
        default=""
    )
    
    # 1. 机位模式
    camera_mode: EnumProperty(
        name="机位模式",
        description="选择仅渲染当前视角或勾选批量连拍机位",
        items=[
            ('ALL', "📸 勾选批量连拍机位", "自由勾选需要连拍的摄像机列表"),
            ('ACTIVE', "仅当前活动机位", "渲染当前选定的摄像机视角"),
        ],
        default='ALL'
    )
    
    # 2. 输出配置
    image_format: EnumProperty(
        name="输出格式",
        description="选择渲染图片输出格式",
        items=[
            ('JPEG', "JPG (100% 极高清)", "保存为 100% 最高画质 JPG 格式 (Quality: 100)，体积缩减 85%"),
            ('PNG', "PNG (无损透明)", "保存为无损 RGBA PNG 格式"),
        ],
        default='JPEG'
    )
    output_directory: StringProperty(
        name="保存目录",
        description="渲染图与通道图保存目录 (// 代表当前工程同级目录)",
        subtype='DIR_PATH',
        default="//"
    )
    sku_name: StringProperty(
        name="自定义前缀 (选填)",
        description="留空则自动从文件夹目录结构中提取产品名称 (如：Orange_Soda_Can)",
        default=""
    )
    auto_increment: BoolProperty(
        name="自动版本增量 (防覆盖 v01/v02)",
        description="开启后每次渲染自动递增版本号 (如 Product_v01, Product_v02)，绝不覆盖旧图",
        default=True
    )
    export_beauty: BoolProperty(name="成品图 (Beauty)", default=True)
    export_alpha: BoolProperty(name="纯黑剪切蒙版 (Alpha)", default=False)
    export_cryptomatte: BoolProperty(name="Cryptomatte 智能选区", default=True)
    auto_open_folder: BoolProperty(name="渲染完成后自动弹出文件夹", default=True)


RENDER_SUBFOLDER_CANDIDATES = [
    "04_Renders_通道输出",
    "04_Renders_高清分层输出",
    "04_Renders",
    "03_输出",
    "05_Delivery_最终交付",
    "渲染",
    "Renders",
    "Output"
]

def get_resolved_output_directory(scene):
    """
    智能解析保存目录：
    1. 若用户在面板手动指定了绝对路径，优先使用；
    2. 若为默认 // 或空，自动向上探测父级是否存在「04_Renders_通道输出」或「渲染」；
    3. 若工程位于「03_3D_三维工程」等子目录，自动回退到父级并归入「04_Renders_通道输出」；
    4. 绝不会错误保存在 03 建模工程目录中！
    """
    raw_dir = scene.packaging_props.output_directory.strip()
    
    if bpy.data.filepath:
        if raw_dir and raw_dir not in ("//", ".", ".\\", "./"):
            abs_dir = bpy.path.abspath(raw_dir)
            if abs_dir:
                return abs_dir
                
        blend_dir = os.path.dirname(os.path.abspath(bpy.data.filepath))
        parent_dir = os.path.dirname(blend_dir)
        
        # 1. 优先扫描父级目录是否已有标准渲染输出文件夹
        for c in RENDER_SUBFOLDER_CANDIDATES:
            cand_p = os.path.join(parent_dir, c)
            if os.path.exists(cand_p) and os.path.isdir(cand_p):
                return cand_p
                
        # 2. 扫描同级目录是否已有渲染输出文件夹
        for c in RENDER_SUBFOLDER_CANDIDATES:
            cand_p = os.path.join(blend_dir, c)
            if os.path.exists(cand_p) and os.path.isdir(cand_p):
                return cand_p

        # 3. 检查当前目录是否是工程子目录 (如 03_3D_三维工程, 03_3D, 02_工程, 模型, 3D)
        blend_dir_name = os.path.basename(blend_dir).lower()
        if any(k in blend_dir_name for k in ["03_3d", "3d", "三维", "工程", "模型", "02_工程", "blend"]):
            target_dir = os.path.join(parent_dir, "04_Renders_通道输出")
            try:
                os.makedirs(target_dir, exist_ok=True)
            except Exception:
                pass
            return target_dir
            
        return blend_dir
    else:
        if raw_dir and raw_dir not in ("//", ".", ".\\", "./"):
            abs_dir = bpy.path.abspath(raw_dir)
            if abs_dir:
                return abs_dir
        desktop = os.path.join(os.path.expanduser("~"), "Desktop", "Renders")
        return desktop


def auto_detect_project_name(scene):
    """
    智能穿透并识别产品名称：
    自动跳过 '03_3D_三维工程', '04_Renders_通道输出', '渲染', 'Renders', 'Output', '3D' 等通用归档层，
    精确提取上级目录作为产品名
    """
    custom = scene.packaging_props.sku_name.strip()
    if custom:
        return custom
        
    fp = bpy.data.filepath
    if fp:
        dir_path = os.path.dirname(fp)
    else:
        dir_path = get_resolved_output_directory(scene)
        
    if not dir_path or dir_path == "//" or dir_path == ".":
        return "Render_Product"
        
    parts = os.path.normpath(dir_path).split(os.sep)
    generic_names = {
        '渲染', 'renders', 'render', 'output', '3d', '工程', 'blend', 'temp', 'textures', '',
        '01_design_平面原稿', '02_textures_贴图资产', '03_3d_三维工程', '04_renders_通道输出', '05_delivery_最终交付',
        '01_design', '02_textures', '03_3d', '04_renders', '05_delivery'
    }
    
    for part in reversed(parts):
        clean_part = part.strip().lower()
        if clean_part and clean_part not in generic_names and not part.endswith(':'):
            return part
            
    return "Render_Product"


def get_next_incremental_sku(output_dir, base_sku):
    """自动扫描输出文件夹，计算下一个自增版本号 (如 SKU_v01 -> SKU_v02)"""
    if not os.path.exists(output_dir):
        return f"{base_sku}_v01"
        
    try:
        existing_files = os.listdir(output_dir)
    except Exception:
        return f"{base_sku}_v01"
        
    pattern = re.compile(rf"^{re.escape(base_sku)}_v(\d+)", re.IGNORECASE)
    
    max_v = 0
    has_base = False
    for f in existing_files:
        match = pattern.match(f)
        if match:
            try:
                v_num = int(match.group(1))
                if v_num > max_v:
                    max_v = v_num
            except Exception:
                pass
        elif f.startswith(base_sku):
            has_base = True
            
    if max_v > 0:
        return f"{base_sku}_v{max_v + 1:02d}"
    elif has_base:
        return f"{base_sku}_v02"
    else:
        return f"{base_sku}_v01"


def get_compositor_tree(scene):
    """兼容 Blender 5.2 compositing_node_group 与旧版 node_tree"""
    if hasattr(scene, "compositing_node_group"):
        if scene.compositing_node_group is None:
            ng = bpy.data.node_groups.new("Compositing Nodetree", "CompositorNodeTree")
            scene.compositing_node_group = ng
        return scene.compositing_node_group
    elif hasattr(scene, "node_tree"):
        scene.use_nodes = True
        return scene.node_tree
    return None


def get_active_view_layer(context, scene):
    """安全获取当前活动 ViewLayer"""
    if hasattr(context, "view_layer") and context.view_layer:
        return context.view_layer
    if hasattr(scene, "view_layers") and len(scene.view_layers) > 0:
        return scene.view_layers[0]
    return None


def find_final_beauty_socket(tree, rl_node):
    """智能寻找经 Render Raw 等调色后的最终 Beauty Socket"""
    for n in tree.nodes:
        if n.type in ('COMPOSITE', 'GROUP_OUTPUT'):
            for inp in n.inputs:
                if inp.is_linked and len(inp.links) > 0:
                    for link in inp.links:
                        if link.from_node.type != 'OUTPUT_FILE':
                            return link.from_socket

    for n in tree.nodes:
        if ('render raw' in n.name.lower()) or (n.type == 'GROUP' and n.node_tree and 'render raw' in n.node_tree.name.lower()):
            for out in n.outputs:
                if 'image' in out.name.lower() or out.type == 'RGBA':
                    return out

    if rl_node and 'Image' in rl_node.outputs:
        return rl_node.outputs['Image']
    return None


def setup_compositor_and_passes(context, props, effective_prefix):
    """配置合成器：通过物理 Socket 索引连接，彻底免疫超长中文名称导致的 KeyError"""
    scene = context.scene
    tree = get_compositor_tree(scene)
    if not tree:
        return
        
    scene.render.film_transparent = True
    
    vl = get_active_view_layer(context, scene)
    if vl and props.export_cryptomatte:
        vl.use_pass_cryptomatte_object = True
        
    rl_node = None
    for n in tree.nodes:
        if n.type == 'R_LAYERS':
            rl_node = n
            break
    if not rl_node:
        rl_node = tree.nodes.new('CompositorNodeRLayers')
        rl_node.location = (0, 0)
        
    fo_node = tree.nodes.get("Packaging_FileOutput")
    if not fo_node:
        fo_node = tree.nodes.new('CompositorNodeOutputFile')
        fo_node.name = "Packaging_FileOutput"
        fo_node.label = "【包装多通道自动导出】"
    fo_node.location = (800, -200)
    
    abs_out_dir = get_resolved_output_directory(scene)
    if abs_out_dir and not os.path.exists(abs_out_dir):
        os.makedirs(abs_out_dir, exist_ok=True)
        
    if hasattr(fo_node, "directory"):
        fo_node.directory = abs_out_dir
    if hasattr(fo_node, "base_path"):
        fo_node.base_path = abs_out_dir
        
    if hasattr(fo_node, "file_name"):
        fo_node.file_name = ""

    if hasattr(fo_node.format, "media_type"):
        fo_node.format.media_type = 'IMAGE'
        
    if props.image_format == 'JPEG':
        fo_node.format.file_format = 'JPEG'
        fo_node.format.color_mode = 'RGB'
        fo_node.format.quality = 100
    else:
        fo_node.format.file_format = 'PNG'
        fo_node.format.color_mode = 'RGBA'
        fo_node.format.color_depth = '8'

    final_beauty_socket = find_final_beauty_socket(tree, rl_node)

    # 智能查找或自动创建 Cryptomatte 节点并连线 (Render Layers [图像] -> Cryptomatte [图像], Cryptomatte [选取/Pick] -> File Output)
    crypto_socket = None
    if props.export_cryptomatte:
        crypto_node = None
        for n in tree.nodes:
            if n.type in ('CRYPTOMATTE', 'CRYPTOMATTE_V2') or ('crypto' in n.name.lower() and n.type != 'OUTPUT_FILE' and 'packaging' not in n.name.lower()):
                crypto_node = n
                break
                
        if not crypto_node:
            try:
                crypto_node = tree.nodes.new(type='CompositorNodeCryptomatteV2')
            except Exception:
                try:
                    crypto_node = tree.nodes.new(type='CompositorNodeCryptomatte')
                except Exception:
                    crypto_node = None
                    
        if crypto_node:
            crypto_node.location = (rl_node.location.x + 350, rl_node.location.y - 250)
            if hasattr(crypto_node, "source"):
                crypto_node.source = 'RENDER'
            if hasattr(crypto_node, "layer_name"):
                crypto_node.layer_name = 'ViewLayer.CryptoObject'
                
            # 严格连接：Render Layers [图像/Image] -> Cryptomatte [图像/Image]
            if 'Image' in rl_node.outputs and 'Image' in crypto_node.inputs:
                tree.links.new(rl_node.outputs['Image'], crypto_node.inputs['Image'])
                
            # 取出 Cryptomatte 的【选取 / Pick】彩色彩图输出（供 PS 魔棒一键抠图选区）
            if 'Pick' in crypto_node.outputs:
                crypto_socket = crypto_node.outputs['Pick']
            elif '选取' in crypto_node.outputs:
                crypto_socket = crypto_node.outputs['选取']
            elif 'Matte' in crypto_node.outputs:
                crypto_socket = crypto_node.outputs['Matte']
            elif 'Image' in crypto_node.outputs:
                crypto_socket = crypto_node.outputs['Image']

    # 兼容 Blender 5.2 file_output_items (使用物理索引进行安全连线)
    if hasattr(fo_node, "file_output_items"):
        fo_node.file_output_items.clear()
        
        # 1. 成品图 (Beauty)
        if props.export_beauty and final_beauty_socket:
            slot_name = f"{effective_prefix}_01_Beauty_成品"
            it = fo_node.file_output_items.new('RGBA', slot_name)
            if hasattr(it, "save_as_render"):
                it.save_as_render = True
            target_in_socket = fo_node.inputs[len(fo_node.file_output_items) - 1]
            tree.links.new(final_beauty_socket, target_in_socket)
            
        # 2. 黑白剪切蒙版 (Alpha)
        if props.export_alpha and 'Alpha' in rl_node.outputs:
            slot_name = f"{effective_prefix}_02_Mask_Alpha蒙版"
            it = fo_node.file_output_items.new('FLOAT', slot_name)
            if hasattr(it, "save_as_render"):
                it.save_as_render = False
            target_in_socket = fo_node.inputs[len(fo_node.file_output_items) - 1]
            tree.links.new(rl_node.outputs['Alpha'], target_in_socket)

        # 3. Cryptomatte 智能选区
        if props.export_cryptomatte and crypto_socket:
            slot_name = f"{effective_prefix}_03_Crypto_选区"
            it = fo_node.file_output_items.new('RGBA', slot_name)
            if hasattr(it, "save_as_render"):
                it.save_as_render = False
            target_in_socket = fo_node.inputs[len(fo_node.file_output_items) - 1]
            tree.links.new(crypto_socket, target_in_socket)

    elif hasattr(fo_node, "file_slots"):
        fo_node.file_slots.clear()
        
        if props.export_beauty and final_beauty_socket:
            slot_name = f"{effective_prefix}_01_Beauty_成品"
            fo_node.file_slots.new(slot_name)
            target_in_socket = fo_node.inputs[len(fo_node.file_slots) - 1]
            tree.links.new(final_beauty_socket, target_in_socket)
            
        if props.export_alpha and 'Alpha' in rl_node.outputs:
            slot_name = f"{effective_prefix}_02_Mask_Alpha蒙版"
            fo_node.file_slots.new(slot_name)
            target_in_socket = fo_node.inputs[len(fo_node.file_slots) - 1]
            tree.links.new(rl_node.outputs['Alpha'], target_in_socket)

        if props.export_cryptomatte and crypto_socket:
            slot_name = f"{effective_prefix}_03_Crypto_选区"
            fo_node.file_slots.new(slot_name)
            target_in_socket = fo_node.inputs[len(fo_node.file_slots) - 1]
            tree.links.new(crypto_socket, target_in_socket)


# ==============================================================================
# 非阻塞多机位异步接力队列管理器
# ==============================================================================

camera_render_queue = []
current_versioned_sku = ""
original_scene_camera = None
target_open_dir = ""
is_in_batch = False

@persistent
def on_render_complete_batch_dispatcher(scene):
    global camera_render_queue, current_versioned_sku, original_scene_camera, target_open_dir, is_in_batch
    
    if not is_in_batch:
        if target_open_dir and os.path.exists(target_open_dir):
            try:
                os.startfile(target_open_dir)
            except Exception:
                pass
            target_open_dir = ""
        return
        
    if camera_render_queue:
        def trigger_next_camera():
            global camera_render_queue, current_versioned_sku
            if not camera_render_queue:
                return None
            next_cam = camera_render_queue.pop(0)
            scene.camera = next_cam
            prefix = f"{current_versioned_sku}_{next_cam.name}"
            setup_compositor_and_passes(bpy.context, scene.packaging_props, prefix)
            print(f"📸 正在连拍机位: {next_cam.name} (剩余 {len(camera_render_queue)} 个)...")
            bpy.ops.render.render('INVOKE_DEFAULT', write_still=True)
            return None

        bpy.app.timers.register(trigger_next_camera, first_interval=0.15)
    else:
        is_in_batch = False
        if original_scene_camera:
            scene.camera = original_scene_camera
            original_scene_camera = None
            
        print("🎉 全部所选机位连拍完成！")
        if target_open_dir and os.path.exists(target_open_dir):
            try:
                os.startfile(target_open_dir)
            except Exception:
                pass
            target_open_dir = ""

@persistent
def on_render_cancel_handler(scene):
    global is_in_batch, camera_render_queue, original_scene_camera, target_open_dir
    is_in_batch = False
    camera_render_queue.clear()
    if original_scene_camera and scene:
        scene.camera = original_scene_camera
        original_scene_camera = None
    target_open_dir = ""
    print("⚠️ 批量连拍已由用户手动中断 (Cancel)，已成功复原场景机位。")


# ==============================================================================
# 操作符集合：项目脚手架创建、机位管理、画幅/采样率快选、阴影捕捉、批量连拍
# ==============================================================================

class PROJECT_OT_create_scaffold(Operator):
    """一键创建标准五级工业项目目录并自动另存工程"""
    bl_idname = "project.create_scaffold"
    bl_label = "一键创建标准项目目录并另存工程"
    bl_options = {'REGISTER'}

    def execute(self, context):
        props = context.scene.packaging_props
        
        proj_name = props.scaffold_project_name.strip()
        if not proj_name:
            self.report({'WARNING'}, "请输入【产品/SKU名称】！")
            return {'CANCELLED'}
            
        root_dir = bpy.path.abspath(props.scaffold_root_dir.strip())
        if not root_dir or root_dir == "//":
            root_dir = os.path.join(os.path.expanduser("~"), "Desktop", "Projects")
            
        client_name = props.scaffold_client_name.strip()
        if client_name:
            project_base_dir = os.path.join(root_dir, client_name, proj_name)
        else:
            project_base_dir = os.path.join(root_dir, proj_name)
            
        # 1. 创建标准五级目录
        subfolders = [
            "01_Design_平面原稿",
            "02_Textures_贴图资产",
            "03_3D_三维工程",
            "04_Renders_通道输出",
            "05_Delivery_最终交付"
        ]
        
        for sub in subfolders:
            os.makedirs(os.path.join(project_base_dir, sub), exist_ok=True)
            
        # 2. 另存当前工程到 03_3D_三维工程
        blend_save_path = os.path.join(project_base_dir, "03_3D_三维工程", f"{proj_name}.blend")
        bpy.ops.wm.save_as_mainfile(filepath=blend_save_path)
        
        # 3. 锁定相对路径与渲染输出目录
        try:
            bpy.ops.file.make_paths_relative()
        except Exception:
            pass
            
        props.output_directory = "//../04_Renders_通道输出/"
        props.sku_name = ""  # 留空以自动识别
        
        # 4. 弹出项目根目录
        try:
            os.startfile(project_base_dir)
        except Exception:
            pass
            
        self.report({'INFO'}, f"🎉 成功创建标准项目 [{proj_name}] 并另存工程！")
        return {'FINISHED'}


class RENDER_OT_select_all_cameras(Operator):
    """一键全选 / 全不选 / 反选场景中的摄像机"""
    bl_idname = "render.select_all_cameras"
    bl_label = "选择所有摄像机"
    bl_options = {'REGISTER', 'UNDO'}

    action: EnumProperty(
        items=[
            ('ALL', "全选", "勾选所有摄像机"),
            ('NONE', "全不选", "取消勾选所有摄像机"),
            ('INVERT', "反选", "反转勾选状态"),
        ]
    )

    def execute(self, context):
        cams = [obj for obj in context.scene.objects if obj.type == 'CAMERA']
        for cam in cams:
            if self.action == 'ALL':
                cam.use_for_batch_render = True
            elif self.action == 'NONE':
                cam.use_for_batch_render = False
            elif self.action == 'INVERT':
                cam.use_for_batch_render = not cam.use_for_batch_render
        return {'FINISHED'}


class RENDER_OT_set_active_camera(Operator):
    """一键将选中的机位设为当前活动视角"""
    bl_idname = "render.set_active_camera"
    bl_label = "设为当前机位"
    bl_options = {'REGISTER', 'UNDO'}

    cam_name: StringProperty()

    def execute(self, context):
        cam = context.scene.objects.get(self.cam_name)
        if cam and cam.type == 'CAMERA':
            context.scene.camera = cam
            self.report({'INFO'}, f"当前活动摄像机已切换为: {cam.name}")
        return {'FINISHED'}


class RENDER_OT_set_aspect_ratio(Operator):
    """一键切换电商标准画幅与分辨率"""
    bl_idname = "render.set_aspect_ratio"
    bl_label = "设置画幅比例"
    bl_options = {'REGISTER', 'UNDO'}

    ratio_type: EnumProperty(
        items=[
            ('1_1', "1:1 (2000x2000)", "电商白底/正方形主图"),
            ('3_4', "3:4 (1500x2000)", "小红书/移动端竖版"),
            ('16_9', "16:9 (1920x1080)", "横版海报/电商Banner"),
            ('4_3', "4:3 (1600x1200)", "标准详情图"),
        ]
    )

    def execute(self, context):
        render = context.scene.render
        render.resolution_percentage = 100
        
        if self.ratio_type == '1_1':
            render.resolution_x = 2000
            render.resolution_y = 2000
            desc = "1:1 正方形主图 (2000 × 2000)"
        elif self.ratio_type == '3_4':
            render.resolution_x = 1500
            render.resolution_y = 2000
            desc = "3:4 移动端竖版 (1500 × 2000)"
        elif self.ratio_type == '16_9':
            render.resolution_x = 1920
            render.resolution_y = 1080
            desc = "16:9 横版海报 (1920 × 1080)"
        elif self.ratio_type == '4_3':
            render.resolution_x = 1600
            render.resolution_y = 1200
            desc = "4:3 详情页图 (1600 × 1200)"
            
        self.report({'INFO'}, f"已切换画幅为: {desc}")
        return {'FINISHED'}


class RENDER_OT_set_render_quality(Operator):
    """一键切换渲染品质与采样率档位"""
    bl_idname = "render.set_render_quality"
    bl_label = "设置渲染采样品质"
    bl_options = {'REGISTER', 'UNDO'}

    quality: EnumProperty(
        items=[
            ('PREVIEW', "⚡ 极速预览档 (64采样)", "快速预览光影构图"),
            ('FINAL', "💎 高清交付档 (1024采样+降噪)", "商业交付最终高精输出"),
        ]
    )

    def execute(self, context):
        scene = context.scene
        if scene.render.engine == 'CYCLES':
            if self.quality == 'PREVIEW':
                scene.cycles.samples = 64
                scene.cycles.use_denoising = True
                desc = "⚡ 极速预览档 (64 采样 / 开启降噪)"
            else:
                scene.cycles.samples = 1024
                scene.cycles.use_denoising = True
                desc = "💎 高清交付档 (1024 采样 / 高精降噪)"
        else:
            if hasattr(scene, "eevee"):
                if self.quality == 'PREVIEW':
                    scene.eevee.taa_render_samples = 32
                    desc = "⚡ EEVEE 预览档 (32 采样)"
                else:
                    scene.eevee.taa_render_samples = 128
                    desc = "💎 EEVEE 高清档 (128 采样)"
            else:
                desc = "已应用采样设置"
                
        self.report({'INFO'}, f"已设置渲染品质: {desc}")
        return {'FINISHED'}


class OBJECT_OT_toggle_shadow_catcher(Operator):
    """一键将选中的物体设为/切换阴影捕捉 (Shadow Catcher)"""
    bl_idname = "object.toggle_shadow_catcher"
    bl_label = "设为阴影捕捉"
    bl_options = {'REGISTER', 'UNDO'}

    def execute(self, context):
        sel_objs = [obj for obj in context.selected_objects if obj.type == 'MESH']
        if not sel_objs:
            self.report({'WARNING'}, "请先在 3D 视口中选中需要作为地面的物体！")
            return {'CANCELLED'}
            
        context.scene.render.film_transparent = True
        
        count = 0
        all_already = all(obj.is_shadow_catcher for obj in sel_objs)
        target_state = not all_already
        
        for obj in sel_objs:
            obj.is_shadow_catcher = target_state
            count += 1
            
        state_str = "阴影捕捉 (Shadow Catcher)" if target_state else "普通可见物体"
        self.report({'INFO'}, f"已将选中的 {count} 个物体设为: {state_str}")
        return {'FINISHED'}


class RENDER_OT_packaging_export_all(Operator):
    """一键渲染（非阻塞实时显示进度窗口，支持单机位或勾选多机位连拍）"""
    bl_idname = "render.packaging_export_all"
    bl_label = "渲染并导出全套通道"
    bl_options = {'REGISTER'}

    def execute(self, context):
        global camera_render_queue, current_versioned_sku, original_scene_camera, target_open_dir, is_in_batch
        scene = context.scene
        props = scene.packaging_props
        
        # 1. 智能提取产品项目名
        project_name = auto_detect_project_name(scene)
        abs_out_dir = get_resolved_output_directory(scene)
        if abs_out_dir and not os.path.exists(abs_out_dir):
            os.makedirs(abs_out_dir, exist_ok=True)
            
        # 2. 版本增量计算
        if props.auto_increment:
            versioned_sku = get_next_incremental_sku(abs_out_dir, project_name)
        else:
            versioned_sku = project_name
            
        current_versioned_sku = versioned_sku
        target_open_dir = abs_out_dir if props.auto_open_folder else ""

        # 挂载异步接力监听
        if on_render_complete_batch_dispatcher not in bpy.app.handlers.render_complete:
            bpy.app.handlers.render_complete.append(on_render_complete_batch_dispatcher)

        # 3. 执行多机位连拍 vs 单机位渲染
        if props.camera_mode == 'ALL':
            all_cams = [obj for obj in scene.objects if obj.type == 'CAMERA']
            if not all_cams:
                self.report({'WARNING'}, "场景中未找到任何摄像机！")
                return {'CANCELLED'}
                
            selected_cams = [cam for cam in all_cams if getattr(cam, "use_for_batch_render", True)]
            if not selected_cams:
                self.report({'WARNING'}, "请在面板列表中至少勾选一个需要连拍的摄像机！")
                return {'CANCELLED'}
                
            original_scene_camera = scene.camera
            is_in_batch = True
            
            # 建立连拍队列
            camera_render_queue = list(selected_cams)
            
            # 启动第一个机位的实时渲染窗口
            first_cam = camera_render_queue.pop(0)
            scene.camera = first_cam
            effective_prefix = f"{versioned_sku}_{first_cam.name}"
            setup_compositor_and_passes(context, props, effective_prefix)
            
            self.report({'INFO'}, f"📸 开始连拍 [{project_name}] 的 {len(selected_cams)} 个机位...")
            bpy.ops.render.render('INVOKE_DEFAULT', write_still=True)
            return {'FINISHED'}
            
        else:
            is_in_batch = False
            camera_render_queue = []
            active_cam = scene.camera
            cam_suffix = f"_{active_cam.name}" if active_cam else ""
            effective_prefix = f"{versioned_sku}{cam_suffix}"
            
            setup_compositor_and_passes(context, props, effective_prefix)
            self.report({'INFO'}, f"正在渲染 [{effective_prefix}]...")
            bpy.ops.render.render('INVOKE_DEFAULT', write_still=True)
            return {'FINISHED'}


class RENDER_OT_open_output_folder(Operator):
    bl_idname = "render.open_packaging_folder"
    bl_label = "打开输出文件夹"

    def execute(self, context):
        scene = context.scene
        abs_out_dir = get_resolved_output_directory(scene)
        if abs_out_dir and not os.path.exists(abs_out_dir):
            os.makedirs(abs_out_dir, exist_ok=True)
        if abs_out_dir:
            os.startfile(abs_out_dir)
        return {'FINISHED'}


# ==============================================================================
# UI 侧边栏面板 (N 键)
# ==============================================================================

class VIEW3D_PT_packaging_pipeline(Panel):
    bl_label = "包装渲染与多通道输出"
    bl_idname = "VIEW3D_PT_packaging_pipeline"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = "包装渲染"

    def draw(self, context):
        layout = self.layout
        scene = context.scene
        props = scene.packaging_props
        
        # 模块 1：画幅与渲染品质快选
        box_quick = layout.box()
        box_quick.label(text="📐 画幅比例与品质快切", icon='IMAGE_PLANE')
        
        row_res = box_quick.row(align=True)
        op1 = row_res.operator("render.set_aspect_ratio", text="1:1 (2K主图)")
        op1.ratio_type = '1_1'
        op2 = row_res.operator("render.set_aspect_ratio", text="3:4 (竖版)")
        op2.ratio_type = '3_4'
        op3 = row_res.operator("render.set_aspect_ratio", text="16:9 (横版)")
        op3.ratio_type = '16_9'
        
        row_qual = box_quick.row(align=True)
        q1 = row_qual.operator("render.set_render_quality", text="⚡ 极速预览 (64采样)")
        q1.quality = 'PREVIEW'
        q2 = row_qual.operator("render.set_render_quality", text="💎 高清交付 (1024采样)")
        q2.quality = 'FINAL'
        
        # 模块 2：地面阴影工具
        box_util = layout.box()
        box_util.label(text="地面阴影工具", icon='MOD_OPACITY')
        row_sc = box_util.row(align=True)
        row_sc.scale_y = 1.2
        sel_active = context.active_object
        is_sc = sel_active.is_shadow_catcher if (sel_active and sel_active.type == 'MESH') else False
        btn_text = "🔘 切换阴影捕捉 (当前: 开)" if is_sc else "🔘 选中所选一键设为阴影捕捉"
        btn_icon = 'CHECKBOX_HLT' if is_sc else 'CHECKBOX_DEHLT'
        row_sc.operator("object.toggle_shadow_catcher", text=btn_text, icon=btn_icon)

        # 模块 3：机位选择与动态多机位勾选列表
        box_cam = layout.box()
        box_cam.label(text="机位选择与批量勾选", icon='CAMERA_DATA')
        box_cam.prop(props, "camera_mode", text="机位模式")
        
        all_cams = [obj for obj in scene.objects if obj.type == 'CAMERA']
        
        if props.camera_mode == 'ALL':
            row_sel_btn = box_cam.row(align=True)
            op_all = row_sel_btn.operator("render.select_all_cameras", text="全选")
            op_all.action = 'ALL'
            op_none = row_sel_btn.operator("render.select_all_cameras", text="全不选")
            op_none.action = 'NONE'
            op_inv = row_sel_btn.operator("render.select_all_cameras", text="反选")
            op_inv.action = 'INVERT'
            
            if all_cams:
                col_cams = box_cam.column(align=True)
                for cam in all_cams:
                    row_c = col_cams.row(align=True)
                    is_active = (cam == scene.camera)
                    row_c.prop(cam, "use_for_batch_render", text=cam.name)
                    op_set = row_c.operator("render.set_active_camera", text="", icon='VIEW_CAMERA' if is_active else 'RESTRICT_VIEW_OFF')
                    op_set.cam_name = cam.name
            else:
                box_cam.label(text="场景中未检测到任何摄像机！", icon='ERROR')
        else:
            cur_cam = scene.camera.name if scene.camera else "未指定"
            box_cam.label(text=f"当前活动机位: {cur_cam}", icon='VIEW_CAMERA')
                
        # 模块 4：输出配置与通道
        box_out = layout.box()
        box_out.label(text="输出配置与通道", icon='OUTPUT')
        
        row_fmt = box_out.row()
        row_fmt.prop(props, "image_format", expand=True)
        
        box_out.prop(props, "output_directory", text="保存目录")
        resolved_out = get_resolved_output_directory(scene)
        if resolved_out:
            box_out.label(text=f"📂 目标目录: {os.path.basename(resolved_out)}", icon='FILE_FOLDER')
        
        detected_name = auto_detect_project_name(scene)
        box_out.prop(props, "sku_name", text="自定义前缀", placeholder=f"自动识别: {detected_name}")
        box_out.prop(props, "auto_increment", text="自动版本增量 (v01, v02 防覆盖)")
        
        row_pass = box_out.row()
        row_pass.prop(props, "export_beauty")
        row_pass.prop(props, "export_cryptomatte")
        box_out.prop(props, "auto_open_folder")
        
        # 模块 5：核心大按钮
        layout.separator()
        col_btn = layout.column(align=True)
        col_btn.scale_y = 2.0
        
        if props.camera_mode == 'ALL':
            checked_count = len([c for c in all_cams if getattr(c, "use_for_batch_render", True)])
            btn_main_text = f"📸 一键连拍勾选机位 ({checked_count}/{len(all_cams)}) 并导出通道"
            btn_main_icon = 'OUTLINER_OB_CAMERA'
        else:
            btn_main_text = "渲染并导出全套通道 (当前机位)"
            btn_main_icon = 'RENDER_STILL'
            
        col_btn.operator("render.packaging_export_all", text=btn_main_text, icon=btn_main_icon)
        
        row_folder = layout.row()
        row_folder.scale_y = 1.2
        row_folder.operator("render.open_packaging_folder", text="📁 打开输出文件夹", icon='FILE_FOLDER')


classes = (
    PackagingPipelineProperties,
    PROJECT_OT_create_scaffold,
    RENDER_OT_select_all_cameras,
    RENDER_OT_set_active_camera,
    RENDER_OT_set_aspect_ratio,
    RENDER_OT_set_render_quality,
    OBJECT_OT_toggle_shadow_catcher,
    RENDER_OT_packaging_export_all,
    RENDER_OT_open_output_folder,
    VIEW3D_PT_packaging_pipeline,
)

def register():
    bpy.types.Object.use_for_batch_render = BoolProperty(
        name="批量渲染此机位",
        description="勾选后在批量连拍时将渲染此机位",
        default=True
    )
    for cls in classes:
        bpy.utils.register_class(cls)
    bpy.types.Scene.packaging_props = PointerProperty(type=PackagingPipelineProperties)
    if on_render_complete_batch_dispatcher not in bpy.app.handlers.render_complete:
        bpy.app.handlers.render_complete.append(on_render_complete_batch_dispatcher)
    if on_render_cancel_handler not in bpy.app.handlers.render_cancel:
        bpy.app.handlers.render_cancel.append(on_render_cancel_handler)

def unregister():
    if hasattr(bpy.types.Object, "use_for_batch_render"):
        del bpy.types.Object.use_for_batch_render
    if on_render_complete_batch_dispatcher in bpy.app.handlers.render_complete:
        bpy.app.handlers.render_complete.remove(on_render_complete_batch_dispatcher)
    if on_render_cancel_handler in bpy.app.handlers.render_cancel:
        bpy.app.handlers.render_cancel.remove(on_render_cancel_handler)
    if hasattr(bpy.types.Scene, "packaging_props"):
        del bpy.types.Scene.packaging_props
    for cls in reversed(classes):
        try:
            bpy.utils.unregister_class(cls)
        except Exception:
            pass

if __name__ == "__main__":
    register()