import bpy
import math
import pandas as pds
import sys
from dataclasses import dataclass
from enum import Enum 

class Color(Enum):
    White = 1
    Yellow = 2
    Red = 3
    Black = 4

class ExitError(Exception):
    pass

@dataclass
class LedSection():
    row: int
    col: int
    start: int
    end: int
    vert: bool
    reverse: bool
    color: Color
    difSize: bool
    optionalSize: int

############################################
# Name        : test_main
# Called by   : Panel Creation PopupMenu execute - line 298
# Parameters  : file path, sheet number, led start row, led column start, list of LED sections
# Returns     : none
# Description : Here for debugging purposes to make sure the UI is working with the backend
############################################
def test_main(filePathName, sheetNumber, ledRowStart, ledColStart, ledSections):
    print("\n\n\nINSIDE main.py\n\n\n")
    print(filePathName)
    print(sheetNumber)
    print(ledRowStart)
    print(ledColStart)
    for led in ledSections:
        print("Start: " + str(led.start), end=' ')
        print("End: " + str(led.end), end=' ')
        print("Row: " + str(led.start), end=' ')
        print("Col: " + str(led.start))

############################################
# Name        : main
# Called by   : Panel Creation PopupMenu execute - line 300
# Parameters  : file path, sheet number, led start row, led column start, list of LED sections
# Returns     : none
# Description : main function for backend -> create LED's and create animation sequence
############################################
def main(filePathName, sheetName, ledRowStart, ledColStart, ledSections):
    print("\nSTART OF NEW TEST IN main.py\n")
    context = bpy.context
    scene = context.scene

    # Turn on bloom effect
    scene.render.engine = 'BLENDER_EEVEE'
    scene.eevee.use_bloom = True

    excel_file_path = filePathName  
    
    # Read excel file and store the data in an array
    xls_data = pds.read_excel(excel_file_path, sheet_name=sheetName)

    frames_list = xls_data.iloc[ledRowStart:,0]

    #The 2 is since the first frame is always 0, no matter milliseconds or seconds
    #If the value is in milliseconds (less than 1), multiply milliseconds to frames
    if frames_list.iloc[2] < 1:
        for index, frame in enumerate(frames_list):
            frames_list.iloc[index] = frame * 1000

    size_of_led  = 2

    materials_list = []
    extra_materials_list = []
    led_row_dict = {}
    led_col_dict = {}
    
    #Find starting and ending positions for all LED Sections
    findAllStart(ledSections, led_col_dict, led_row_dict)
    
    #Create each individual LED cube and add a material to each cube
    materials_list, extra_materials_list = createBars(size_of_led, ledSections, led_row_dict, led_col_dict)
    #Apply emission node for glow and apply for each keyframe specified in excel
    createKeyFrames(ledSections, materials_list, frames_list, xls_data, ledRowStart, ledColStart, extra_materials_list)

    #FINISHED
    context.active_object.select_set(False)

############################################
# Name        : addToDict
# Called by   : findAllStartHelper - line 192, 198
# Parameters  : led strip index, dictionary of LEDS, led section, size of led, axis
# Returns     : none
# Description : Add a LED section to the dictionary that maps row, column number to a start and end position
############################################
def addToDict(led_strip_num, led_dict, led, size_of_led, axis):
    led_start = 0
     # ADD TO ANY EXISTING ROW INDEXING
    for prev_index in range(led_strip_num - 1, -1, -1):
        if prev_index in led_dict:
            led_start = led_dict[prev_index][1]
            break
    #ONLY FIRST INDEX SHOULD INPUT WITH LED_START_COL as 0
    if axis == "Row":
        if led.vert:
            if led.difSize:
                led_dict[led_strip_num] = (led_start, led_start - led.optionalSize * size_of_led)
            else:
                led_dict[led_strip_num] = (led_start, led_start - (abs(led.end - led.start) * size_of_led))
        else:
            led_dict[led_strip_num] = (led_start, led_start - size_of_led)
    if axis == "Column":
        if not led.vert:
            if led.difSize:
                led_dict[led_strip_num] = (led_start, led_start + led.optionalSize * size_of_led)
            else:
                led_dict[led_strip_num] = (led_start, led_start + (abs(led.end - led.start) * size_of_led))
        else:
            led_dict[led_strip_num] = (led_start, led_start + size_of_led)
    
############################################
# Name        : updateDict
# Called by   : findAllStartHelper - line 193, 199
# Parameters  : led strip index, number of dimensions in the user input matrix, dictionary of LEDS
# Returns     : none
# Description : Update current dictionary if an LED strip index is in between existing indices
############################################
def updateDict(led_strip_num, total_num_dimension, led_dict):
    #PLACE NEW INDEX IN MIDDLE OF TWO INDICES
    for next_index in range(led_strip_num + 1, total_num_dimension):
        if next_index in led_dict:
            cur_index = next_index
            
            for cur in range(next_index, -1, -1):
                cur -= 1
                if cur in led_dict:
                    cur_index = cur
                    break
            #ERROR CHECK
            if cur_index == next_index:
                sys.exit('NO PREVIOUS INDEX')
            
            #FIND RANGE OF PREVIOUS INDEX
            increment = abs(led_dict[cur_index][1] - led_dict[cur_index][0])
            new_start = led_dict[next_index][0] + increment
            new_end = led_dict[next_index][1] + increment
            
            led_dict[next_index] = (new_start, new_end)
            
############################################
# Name        : findAllStartHelper
# Called by   : findAllStart - line 226, 229
# Parameters  : led strip index, dictionary of LEDS, led section, size of led, number of dimensions in matrix, axis
# Returns     : none
# Description : Helper function to add and update each led strip into dictionary
############################################
def findAllStartHelper(led_strip_num, led_dict, led, size_of_led, total_num_dimension, axis):
    #IF FIRST TIME SEEING THIS COLUMN NUMBER ADD START, END TO DICT
        if led_strip_num not in led_dict: 
            
            addToDict(led_strip_num, led_dict, led, size_of_led, axis)
            updateDict(led_strip_num, total_num_dimension, led_dict)
        
        #IF THE PREVIOUS ADDITION WAS NOT THE MAX FOR THAT ROW
        elif abs(led_dict[led_strip_num][1] - led_dict[led_strip_num][0]) == 2:
            
            addToDict(led_strip_num, led_dict, led, size_of_led, axis)
            updateDict(led_strip_num, total_num_dimension, led_dict)
            
        #ALREADY UPDATED THE RANGE
        else:
            pass

############################################
# Name        : findAllStart
# Called by   : main - line 107
# Parameters  : led strip list, led column index dictionary, led row index dictionary
# Returns     : none
# Description : Function to create two dictionaries of row, col indices mapped to start 
#               and end (x,y) coordinates for LED cube creation
############################################
def findAllStart(ledSections, led_col_dict, led_row_dict):
    led_start_col = 0
    led_start_row = 0
    
    total_num_dimension = 5
    
    size_of_led = 2
    
    for led in ledSections: 
        led_strip_col_num = led.col
        led_strip_row_num = led.row
        
        #CONFIGURE COLUMN DICTIONARY
        findAllStartHelper(led_strip_col_num, led_col_dict, led, size_of_led, total_num_dimension, "Column")
        
        #CONFIGURE ROW DICTIONARY
        findAllStartHelper(led_strip_row_num, led_row_dict, led, size_of_led, total_num_dimension, "Row")

############################################
# Name        : createBars
# Called by   : main - line 110
# Parameters  : size of led, list of led strips, led row dictionary, led column dictionary
# Returns     : list of materials attached to cubes, list of materials attached to cubes not orignally in excel
# Description : Create each individual LED cube and add a material to each cube
############################################
def createBars(size_of_led, ledSections, led_row_dict, led_col_dict):
    materials_list = []
    extra_materials_list = []
    led_count = 0
    
    for led in ledSections:
        led_strip_col_num = led.col
        led_strip_row_num = led.row
        
        #SET THE START TO THE START STORED IN DICT
        if not led.reverse: #IF NOT REVERSE 
            led_start_x = led_col_dict[led_strip_col_num][0]
            led_start_y = led_row_dict[led_strip_row_num][0]
        
        elif not led.vert:  #IF REVERSE AND NOT VERT
            led_start_x = led_col_dict[led_strip_col_num][1]
            led_start_y = led_row_dict[led_strip_row_num][0]
        else:               #IF REVERSE AND VERT
            led_start_x = led_col_dict[led_strip_col_num][0]
            led_start_y = led_row_dict[led_strip_row_num][1]

        size = (abs(led.end - led.start) + 1)
        if led.difSize:
            size = led.optionalSize + 1
        
        for i in range(size):
            bpy.ops.mesh.primitive_cube_add()
            ob = bpy.context.active_object
            ob.dimensions = [size_of_led,size_of_led,size_of_led]
            ob.location = [led_start_x, led_start_y, 0]
            
            material = bpy.data.materials.new(name='LedMaterial')
            material.use_nodes = True
            ob.active_material = material
            
            if led.difSize:
                extra_materials_list.append(material)
            else:
                materials_list.append(material)
            
            if led.reverse:
                if led.vert:
                    led_start_y += size_of_led
                else:
                    led_start_x -= size_of_led
            else:
                if led.vert:
                    led_start_y -= size_of_led
                else:
                    led_start_x += size_of_led

    return materials_list, extra_materials_list

############################################
# Name        : createBars
# Called by   : main - line 112
# Parameters  : list of led strips, list of materials tied to cubes, time/frames column from excel, 
#               data from excel, row to start reading from in excel, column to start reading from in excel, 
#               extra list of materials for cubes not in excel
# Returns     : N/A
# Description : Apply emission node for glow and apply for each keyframe specified in excel
############################################
def createKeyFrames(ledSections, materials_list, frames_list, xls_data, ledRowStart, ledColStart, extra_materials_list):
    material_index = 0
    extra_from_difSize = 0
    material: bpy.data.materials
    print(f"Length of extraList: {len(extra_materials_list)}")
    for led in ledSections:
        led_start = led.start
        led_end = led.end
        color = led.color.name
        duplicateSize = 1
        
        if led.reverse:
            led_start = led.end
            led_end = led.start
        print(f"Led Start: {led_start}, Led End: {led_end}")
        for ledIndex in range(led_start, (led_end + 1)):
            if led.difSize:
                ledRange = abs(led.end - led.start) + 1
                duplicateSize = (led.optionalSize + 1) // ledRange
                if ledIndex == led_end:
                    duplicateSize += (led.optionalSize + 1) - (duplicateSize * ledRange)
#            print(f"DuplicateSize: {duplicateSize}")
            for duplicateLED in range(duplicateSize):
                if led.difSize:
#                    print(f"extra_from_difSize: {extra_from_difSize}")
                    material = extra_materials_list[extra_from_difSize]
                else:
                    material = materials_list[material_index]
                nodes = material.node_tree.nodes
                output_node = nodes.get('Material Output')
                emission_node = nodes.new(type='ShaderNodeEmission')
                
                if color == "White":
                    emission_node.inputs[0].default_value = (0.625, 0.818, 1, 1)  # White glow
                elif color == "Yellow":
                    emission_node.inputs[0].default_value = (0.98, 0.85, 0.22, 1) # Yellow glow
                else:
                    emission_node.inputs[0].default_value = (1, 0, 0, 1)          # Red glow
                
                for frame_index,(frame) in enumerate(frames_list):
                    #Error Check
                    checkUserInput(xls_data.iloc[frame_index+ledRowStart,ledIndex+ledColStart], (frame_index + ledRowStart), (ledIndex+ledColStart))
                    if xls_data.iloc[frame_index+ledRowStart,ledIndex+ledColStart] == 0.0:
                        emission_node.inputs[1].default_value = 0 # no emission glow
                    else:
                        emission_node.inputs[1].default_value = (0.371327) * math.exp(4.20955 * (xls_data.iloc[frame_index+ledRowStart,ledIndex+ledColStart] / 100.0))  # Exponential Glow function
                    emission_node.inputs[1].keyframe_insert(data_path='default_value', frame=frame)
                    
                    if frame_index == 0:
                        links = material.node_tree.links
                        links.new(emission_node.outputs[0], output_node.inputs[0])
                if led.difSize:
                    extra_from_difSize += 1
                else:
                    material_index += 1
                    
############################################
# Name        : createBars
# Called by   : createKeyFrames - line 300
# Parameters  : input data from excel, rough estimate of row number in excel, rough estimate of column number in excel
# Returns     : N/A
# Description : Check data to make sure the data is an integer and can be converted into a emission value
############################################
def checkUserInput(input, rowNum, colNum):
    try:
        # Convert it into integer
        val = int(input)
    except ValueError:
        try:
            # Convert it into float
            val = float(input)
        except ValueError:
            raise ExitError("Something went wrong in the EXCEL sheet. Each value must be either an int or a float. Row: %i, Col: %i. Got: %s" %(rowNum, colNum, input))
