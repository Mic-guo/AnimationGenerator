import bpy
import math

############################################
# Name        : createPlane
# Called by   : main(filePathName, sheetName, ledRowStart, ledColStart, ledSections) - line 105
# Parameters  : N/A
# Returns     : none
# Description : Creates the black background for the LED animation 
############################################
def createPlane():
    bpy.ops.mesh.primitive_plane_add(size=1000,location=(0,0,0),rotation=(0,0,math.pi/2))
    plane = bpy.context.active_object
    planeMaterial = bpy.data.materials.new(name='planeBlack')
    planeMaterial.use_nodes = True
    plane.active_material = planeMaterial
    nodes = planeMaterial.node_tree.nodes
    output_node = nodes.get('Material Output')
    emission_node = nodes.new(type='ShaderNodeEmission')
    emission_node.inputs[0].default_value = (0.625, 0.818, 1, 1) #White glow
    emission_node.inputs[1].default_value = 0.0  # Black
    links = planeMaterial.node_tree.links
    links.new(emission_node.outputs[0], output_node.inputs[0])
    emission_node.inputs[1].keyframe_insert(data_path='default_value', frame=0)

class CreatePlaneOperator(bpy.types.Operator):
    bl_idname = "object.add_plane_operator"
    bl_label = "Create background"
    ############################################
    # Name        : execute
    # Called by   : Panel UI 'Create background' button
    # Parameters  : self, context
    # Returns     : 'FINISHED'
    # Description : Run createPlane function when button is pressed
    ############################################
    def execute(self, context):
        createPlane()
        return {'FINISHED'}

############################################
# Name        : remove_cubes
# Called by   : RemoveCubesOperator & RemovePlaneOperator execute - line 86, 100
# Parameters  : N/A
# Returns     : N/A
# Description : Remove all cubes or planes in workspace
############################################
def remove_object(type):
    bpy.ops.object.select_all(action='DESELECT')
   # Iterate over all objects in the scene collection
    for obj in bpy.context.scene.collection.objects:
        # Check if the object is a mesh object with the name starting with "Cube"
        if obj.type == 'MESH' and (obj.name.startswith(type)):
            # Select the object
            obj.select_set(True)
    
    # Delete the selected objects
    bpy.ops.object.delete()
    
    # Remove unlinked materials
    for material in bpy.data.materials:
        if not material.users:
            bpy.data.materials.remove(material)

    # Remove unlinked shader node trees
    for node_tree in bpy.data.node_groups:
        if not node_tree.users:
            bpy.data.node_groups.remove(node_tree)
            
    # Remove unused materials and shader node trees
    bpy.context.view_layer.objects.active = None

############################################
# Name        : purgeData
# Called by   : PurgeOperator execute - line 121
# Parameters  : N/A
# Returns     : N/A
# Description : Purge all orphaned data
############################################
def purgeData():
    for i in range(3):
        bpy.ops.outliner.orphans_purge()


class RemoveCubesOperator(bpy.types.Operator):
    bl_idname = "object.remove_cubes_operator"
    bl_label = "Remove Cubes"
    ############################################
    # Name        : execute
    # Called by   : Panel UI 'Remove Cubes' button
    # Parameters  : self, context
    # Returns     : 'FINISHED'
    # Description : Run remove_cubes function when button is pressed
    ############################################
    def execute(self, context):
        remove_object("Cube")
        return {'FINISHED'}
    
class RemovePlaneOperator(bpy.types.Operator):
    bl_idname = "object.remove_plane_operator"
    bl_label = "Remove Plane"
    ############################################
    # Name        : execute
    # Called by   : Panel UI 'Remove Panel' button
    # Parameters  : self, context
    # Returns     : 'FINISHED'
    # Description : Run remove_panel function when button is pressed
    ############################################
    def execute(self, context):
        remove_object("Plane")
        return {'FINISHED'}

class PurgeOperator(bpy.types.Operator):
    bl_idname = "object.purge_operator"
    bl_label = "Purge data"
    ############################################
    # Name        : execute
    # Called by   : Panel UI 'Purge Data' button
    # Parameters  : self, context
    # Returns     : 'FINISHED'
    # Description : Run purgeData function when button is pressed
    ############################################
    def execute(self, context):
        purgeData()
        return {'FINISHED'}

class RemoveCubesPanel(bpy.types.Panel):
    bl_idname = "OBJECT_PT_remove_cubes_panel"
    bl_label = "Viewing Panel"
    bl_space_type = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category = 'LED Animation'
    ############################################
    # Name        : draw
    # Called by   : register() - line 145
    # Parameters  : self, context
    # Returns     : N/A
    # Description : Creates UI for 'Remove Cubes Panel'
    ############################################
    def draw(self, context):
        layout = self.layout
        layout.operator("object.add_plane_operator")
        layout.operator("object.remove_cubes_operator")
        layout.operator("object.remove_plane_operator")
        layout.operator("object.purge_operator")

############################################
# Name        : register
# Called by   : Panel Creation main - line 369
# Parameters  : N/A
# Returns     : N/A
# Description : Register custom panel and operator in workspace
############################################
def register():
    bpy.utils.register_class(CreatePlaneOperator)
    bpy.utils.register_class(RemoveCubesOperator)
    bpy.utils.register_class(RemoveCubesPanel)
    bpy.utils.register_class(RemovePlaneOperator)
    bpy.utils.register_class(PurgeOperator)

############################################
# Name        : unregister
# Called by   : called when blender closes
# Parameters  : N/A
# Returns     : N/A
# Description : Used to clean up space and delete the remove cubes panel add-on
############################################
def unregister():
    bpy.utils.unregister_class(CreatePlaneOperator)
    bpy.utils.unregister_class(RemoveCubesOperator)
    bpy.utils.unregister_class(RemoveCubesPanel)
    bpy.utils.unregister_class(RemovePlaneOperator)
    bpy.utils.unregister_class(PurgeOperator)
