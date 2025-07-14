bl_info = {
    "name": "Furniture Tools",
    "description": "Инструмент для работы с шаблонами мебели",
    "author": "Ваше имя или псевдоним",
    "version": (1, 0, 0),
    "blender": (3, 0, 0),
    "location": "3D View > N Panel > Furniture",
    "category": "Object",
}

import bpy
import os
from mathutils import Vector

TEMPLATE_FOLDER = "C:\\Users\\sagin\\OneDrive\\Desktop\\samples"

import bpy
from mathutils import Vector
import os

def vert_coord(obj):
    """
    Возвращает глобальные координаты первой вершины объекта.
    :param obj: объект типа MESH
    :return: глобальные координаты (Vector) или None
    """
    if obj.type != 'MESH':
        print("Объект не является мешем!")
        return None

    if len(obj.data.vertices) == 0:
        print("У объекта нет вершин!")
        return None

    first_vertex = obj.data.vertices[0]
    
    local_coords = first_vertex.co

    global_coords = obj.matrix_world @ local_coords
    
    print(global_coords)

    return global_coords


def load_and_replace_template(cube_object, template_name):
    import_object_path = os.path.join(TEMPLATE_FOLDER, template_name)
    
    if not os.path.exists(import_object_path):
        print(f"Файл {import_object_path} не найден!")
        return

    cube_vert = vert_coord(cube_object)
    cube_location = cube_object.location.copy()
    cube_dimensions = cube_object.dimensions.copy()

    with bpy.data.libraries.load(import_object_path, link=False) as (data_from, data_to):
        data_to.objects = data_from.objects

    imported_objects = []
    for obj in data_to.objects:
        if obj is not None:
            bpy.context.collection.objects.link(obj)
            imported_objects.append(obj)

    if not imported_objects:
        print(f"Не удалось загрузить объекты из {template_name}")
        return
    
    first_obj = vert_coord(imported_objects[0])
    x_distance = cube_vert[0] - first_obj[0]
    y_distance = cube_vert[1] - first_obj[1]
    z_distance = cube_vert[2] - first_obj[2]
    
    print(x_distance, y_distance, z_distance)
    
    for obj in imported_objects:
        original_location = obj.location.copy()
        
        obj.location = (
            obj.location.x + x_distance,
            obj.location.y + y_distance,
            obj.location.z + z_distance
            )
        
        print(obj.location.x, obj.location.y, obj.location.z)
        
        obj.select_set(True)
        print(f"Обработан объект: {obj.name} с новыми координатами {obj.location}")

    bpy.context.view_layer.objects.active = imported_objects[-1]

def get_cube_object():
    obj = bpy.context.active_object
    if obj is None or obj.type != 'MESH' or obj.name.lower()[:4] != "cube":
        print("Выберите объект-куб!")
        return None
    return obj

class FurnitureToolsPanel(bpy.types.Panel):
    bl_label = "Furniture Tools"
    bl_idname = "VIEW3D_PT_furniture_tools"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = 'Furniture'

    def draw(self, context):
        layout = self.layout
        col = layout.column()

        col.label(text="Шаги:")
        col.label(text="1. Создайте куб и задайте размеры.")
        col.label(text="2. Выберите куб.")
        col.operator("object.show_templates", text="Выбрать шаблон")

class ShowTemplatesOperator(bpy.types.Operator):
    bl_idname = "object.show_templates"
    bl_label = "Выбрать шаблон"
    bl_description = "Выберите шаблон для применения к кубу"

    def execute(self, context):
        cube_object = get_cube_object()
        if cube_object is None:
            self.report({'WARNING'}, "Выберите куб перед использованием!")
            return {'CANCELLED'}

        templates = [f for f in os.listdir(TEMPLATE_FOLDER) if f.endswith(".blend")]
        if not templates:
            self.report({'WARNING'}, "Шаблоны не найдены!")
            return {'CANCELLED'}
        
        def draw_func(self, context):
            layout = self.layout
            for template in templates:
                layout.operator("object.apply_template", text=template).template_name = template

        context.window_manager.popup_menu(draw_func, title="Шаблоны мебели", icon='FILE_FOLDER')
        return {'FINISHED'}


class ApplyTemplateOperator(bpy.types.Operator):
    bl_idname = "object.apply_template"
    bl_label = "Применить шаблон"
    bl_description = "Применить выбранный шаблон к кубу"
    
    template_name: bpy.props.StringProperty() # type: ignore

    def execute(self, context):
        cube_object = get_cube_object()
        if cube_object is None:
            self.report({'WARNING'}, "Выберите куб перед использованием!")
            return {'CANCELLED'}

        load_and_replace_template(cube_object, self.template_name)
        return {'FINISHED'}

# Регистрация классов
def register():
    bpy.utils.register_class(FurnitureToolsPanel)
    bpy.utils.register_class(ShowTemplatesOperator)
    bpy.utils.register_class(ApplyTemplateOperator)

def unregister():
    bpy.utils.unregister_class(FurnitureToolsPanel)
    bpy.utils.unregister_class(ShowTemplatesOperator)
    bpy.utils.unregister_class(ApplyTemplateOperator)

if __name__ == "__main__":
    register()
