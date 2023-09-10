import bpy 
import pandas as pds
from bpy.props import StringProperty, IntProperty, BoolProperty
from bpy.types import Scene, WindowManager

class confirmRowCol(bpy.types.PropertyGroup):
    startRow: IntProperty(name="", min=0, default=0)
    startCol: IntProperty(name="", min=0, default=0)
    
class popUpTest(bpy.types.Operator):
    bl_idname = "wm.pop_up_test"
    bl_label = "Test sheet name, start row, and start column"
    
    visualStartRow: IntProperty(name="", min=0, default=0)
    testRowNum: IntProperty(name="", min=0, default=0)
    
    ############################################
    # Name        : execute
    # Called by   : Popup UI 'OK' button
    # Parameters  : self, context
    # Returns     : 'FINISHED'
    # Description : Allows for the rest of the input fields to show if valid path and sheet name
    ############################################
    def execute(self, context):
        scene = context.scene
        try:
            xls_data = pds.read_excel(scene.filePathName, sheet_name=scene.sheetName)
            scene.showInput = True
        except:
            scene.showInput = False
        return {'FINISHED'}
    
    ############################################
    # Name        : invoke
    # Called by   : Panel UI 'Test sheet name, start row, start column'
    # Parameters  : self, context, event
    # Returns     : drawing of popup dialog 
    # Description : Calls the draw function of the popup, 
    #               showing the popup in the panel window
    ############################################
    def invoke(self, context, event):
        return context.window_manager.invoke_props_dialog(self)
    
    ############################################
    # Name        : draw
    # Called by   : invoke(self, context, event) - line 34
    # Parameters  : self, context
    # Returns     : none
    # Description : Creates the UI for the popup dialog and 
    #               allows for user trial and error to find start row and start column
    ############################################
    def draw(self, context):
        layout = self.layout
        layout.scale_y = 1.2
        layout.ui_units_x = 30
        
        scene = context.scene
        confirmInfo = scene.confirmRowCol
        
        try:
            xls_data = pds.read_excel(scene.filePathName, sheet_name=scene.sheetName)
        except:
            layout.label(text="There is something wrong with excel sheet name or file path.")
            return None

        layout.label(text="Choose a starting row number to test")

        layout.prop(confirmInfo, "startRow")
        
        num_test = 10 #Arbitrarily chosen
        
        layout.label(text="Verify this is the first 10 inputs in the time/frames column")
        box = layout.box()
        row = box.row()
        
        frames_list = xls_data.iloc[confirmInfo.startRow:confirmInfo.startRow + num_test,0]
        for frame in frames_list:
            split = row.split()
            split.label(text=f"{frame}")
        
        split = layout.split()
        split.label(text="Starting row number seen in excel")
        split.label(text="Choose a row to test first 10 columns")
        
        split = layout.split()
        split.prop(self, "visualStartRow")
        split.prop(self, "testRowNum")
        
        layout.label(text="Choose a starting column number to test")

        layout.prop(confirmInfo, "startCol")
        
        box = layout.box()
        row = box.row()
        
        rowDifference = abs(self.visualStartRow - confirmInfo.startRow)
        test_row = 0
        if confirmInfo.startRow < self.visualStartRow:
            test_row = self.testRowNum - rowDifference
        elif confirmInfo.startRow > self.visualStartRow:
            test_row = self.testRowNum + rowDifference
        else:
            pass
        
        if (test_row < len(xls_data.index)) and test_row >= 0:
            row.label(text="Verify this is the first 10 inputs in the selected row starting from the start column")
            row = box.row()
            for i in range(confirmInfo.startCol, confirmInfo.startCol + num_test):
                split = row.split()
                split.label(text=f"{checkDataInput(xls_data.iloc[test_row][i])}")
        else:
            row = box.row()
            row.label(text="No valid data for the chosen row")

############################################
# Name        : checkDataInput
# Called by   : draw(self, context) line 45
# Parameters  : data from excel sheet
# Returns     : int or any
# Description : Change all number inputs to int and if its not 
#               float or int return original value
############################################        
def checkDataInput(input):
    try:
        val = int(input)
        return val
    except ValueError:
        return input

############################################
# Name        : register
# Called by   : Panel Creation main - line 359
# Parameters  : N/A
# Returns     : N/A
# Description : Loads scene variables and classes for created add-on
############################################
def register():
    bpy.utils.register_class(popUpTest)
    bpy.utils.register_class(confirmRowCol)
    bpy.types.Scene.confirmRowCol = bpy.props.PointerProperty(type=confirmRowCol)
    
############################################
# Name        : unregister
# Called by   : called when blender closes
# Parameters  : N/A
# Returns     : N/A
# Description : Used to clean up space and delete the add-on
############################################
def unregister():
    bpy.utils.unregister_class(popUpTest)
    bpy.utils.unregister_class(confirmRowCol)