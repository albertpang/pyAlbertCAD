import win32com.client
from pywintypes import com_error

import math
import numpy as np
import pandas as pd

from datatypes import APoint
import wait
import re

# Importing Class Objects
from entity import Fitting, Line, PolyLine, Text, Viewport, LeaderLine
from dataframes import * 

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

class Sheet:
    """
    Initialize the Sheet object with layout, lines data frame, 
    fittings data frame, texts data frame, and bill of materials data frame.
    """
    def __init__(self, layout):
        self.layout = layout

    def find_blocks(self):
        """ Finds all the blocks in the current drawing and classifies them. 
        The information about these objects is then stored in respective data
        frames.
        """
        entities = wait.wait_for_attribute(self.layout, "Block")
        entitiesCount = wait.wait_for_attribute(entities, "Count") 
        i, errorCount = 0, 1
        while i < entitiesCount and errorCount <= 3:
            if i == (entitiesCount // 2):
                print ("--- 50% done ---")
            # try:
            entity = wait.wait_for_method_return(entities, "Item", i)
            entityObjectName = wait.wait_for_attribute(entity, "ObjectName")
            # Line Object
            if entityObjectName == 'AcDbLine': # and entity.Layer == 'C-PR-WATER':
                if entity.Length > 1:
                    l = Line(entity, self.layout.name, False)
                    LinesDF.loc[len(LinesDF.index)] = [l.ID, l.sheet, l.layer, 
                                                        l.startX, l.startY, l.endX, 
                                                        l.endY, l.length, l.slope]

            # Polyline
            elif entityObjectName == 'AcDbPolyline': # and entity.Layer == 'C-PR-WATER':
                if entity.Length > 1:
                    # If there are only two coordinates, convert Polyline into a line
                    if len(wait.wait_for_attribute(entity, "Coordinates")) == 4:
                        l = Line(entity, self.layout.name, True)
                        LinesDF.loc[len(LinesDF.index)] = [l.ID, l.sheet, l.layer, 
                                                            l.startX, l.startY, l.endX, 
                                                            l.endY, l.length, l.slope]
                    # Otherwise, it is an authentic Polyline
                    else:
                        pl = PolyLine(entity, self.layout.name, LinesDF)

            elif entityObjectName == 'AcDbBlockReference': # and entity.Name.startswith("WATER"):
                f = Fitting(entity, self.layout.name)
                FittingsDF.loc[len(FittingsDF.index)] = [f.ID, f.sheet, f.BlockName, 
                                                            f.locationX, f.locationY, "N/A", "N/A"]
                
            # Text Items
            elif (entityObjectName == 'AcDbMText' or entity.ObjectName == 'AcDbText'):
                t = Text(entity, self.layout.name)
                TextsDF.loc[len(TextsDF.index)] = [t.ID, t.sheet, t.text, 
                                                    t.locationX, t.locationY, 
                                                    "N/A", "N/A"]
            # MLeader Object
            elif entityObjectName == 'AcDbMLeader':
                le = LeaderLine(entity, self.layout.name)
                MLeadersDF.loc[len(MLeadersDF.index)] = [le.ID, le.sheet, le.text, le.type, le.startVertexCoordinateX, 
                                                            le.startVertexCoordinateY, le.endVertexCoordinateX, 
                                                            le.endVertexCoordinateY, "N/A", "N/A", "N/A"
                                                            ]
            elif entityObjectName == "AcDbMText":
                print(wait.wait_for_attribute(entity, "TextString"))

            # elif entityObjectName == 'AcDbViewport':
            #     pass
            errorCount = 0 
            i += 1
            # except com_error as e:
            #     time.sleep(0.5)
            #     print("autocad Exception")
            # except Exception as e:
            #     errorCount += 1
            #     print(f"\tAttempt: {errorCount}", i, entityObjectName, e)
    
    
    def is_collinear(self, x1, y1, x2, y2, x3, y3) -> bool:
        """ Checks if three points are collinear."""
        collinearity = x1*(y3-y2)+x3*(y2-y1)+x2*(y1-y3)
        return (abs(collinearity) < 0.005)
    
    def calc_midpoint(self, x1, y1, x2, y2) -> float:
        """Calculates the midpoint between two points in a 2D space."""
        return ((x1 + x2) / 2, (y1 + y2) / 2)
        
    def calc_distance(self, x1, y1, x2, y2) -> float:
        """Calculates the Euclidean distance between two points in a 2D space."""
        return ((x1 - x2) **2 + (y1 - y2) **2) ** 0.5

    def find_appropriate_callout(self):
        """Associates all MLeader Callouts with Appropriate fitting based on distance."""
        def find_fitting(searchString) -> tuple:
            """Calculate closest appropriate fitting based on Callout type"""
            minDistance = 999999
            fittingIndex = 0
            for fittingIndex in FittingsDF.index:
                if FittingsDF['Block Description'][fittingIndex].contains(searchString):
                    x2, y2 = FittingsDF['Block X'][FittingsDF], FittingsDF['Block Y'][FittingsDF]
                    distance = self.calc_distance(x1, y1, x2, y2)
                    if distance < minDistance:
                        minDistance = distance
                        minFitting = FittingsDF['Block Description'][fittingIndex]
            return minFitting, minDistance
        

        print("Associating Closest Callout to Fitting")
        mleaderIndex = 0
        MLeadersDF['Fitting Description'] = 'N/A'
        MLeadersDF['Distance to Appropriate Fitting'] = 'N/A'

        for mleaderIndex in MLeadersDF.index:
            if MLeadersDF['Is MLeader'][mleaderIndex] == False:
                x1, y1 = self.calc_midpoint(MLeadersDF['Starting Vertex Coordinate X'][mleaderIndex], 
                                            MLeadersDF['Starting Vertex Coordinate Y'][mleaderIndex],
                                            MLeadersDF['Ending Vertex Coordinate Y'][mleaderIndex],
                                            MLeadersDF['Ending Vertex Coordinate Y'][mleaderIndex])
                print(MLeadersDF['Text'][mleaderIndex])
                if MLeadersDF['Text'][mleaderIndex] == '1':
                    fitting = find_fitting('valve')
                    MLeadersDF['Fitting Description'] = fitting[0]
                    MLeadersDF['Distance to Appropriate Fitting'] = fitting[1]
                elif MLeadersDF['Text'][mleaderIndex] == '2':
                    fitting = find_fitting('hydr')
                    MLeadersDF['Fitting Description'] = fitting[0]
                    MLeadersDF['Distance to Appropriate Fitting'] = fitting[1]
                elif MLeadersDF['Text'][mleaderIndex] == '3':
                    fitting = find_fitting('valve')
                    MLeadersDF['Fitting Description'] = fitting[0]
                    MLeadersDF['Distance to Appropriate Fitting'] = fitting[1]
                elif MLeadersDF['Text'][mleaderIndex] == '4':
                    fitting = find_fitting('bend')
                    MLeadersDF['Fitting Description'] = fitting[0]
                    MLeadersDF['Distance to Appropriate Fitting'] = fitting[1]
        

    def find_fitting_size(self):
        """Associates all fittings with a Line ID based on collinearity."""
        print("Associating ModelSpace Fittings to Lines")
        fittingIndex, lineIndex = 0, 0
        for fittingIndex in FittingsDF.index:
            x2 = FittingsDF['Block X'][fittingIndex]
            y2 = FittingsDF['Block Y'][fittingIndex]
            for lineIndex in LinesDF.index:
                x1 = LinesDF['Start X'][lineIndex]
                y1 = LinesDF['Start Y'][lineIndex]
                x3 = LinesDF['End X'][lineIndex]
                y3 = LinesDF['End Y'][lineIndex]
                if self.is_collinear(x1, y1, x2, y2, x3, y3):
                    FittingsDF.loc[fittingIndex, 'Matching Line ID'] = \
                        LinesDF['ID'][lineIndex]
                    FittingsDF.loc[fittingIndex, 'Matching Line Length'] = \
                        LinesDF['Length'][lineIndex]
                    
    def find_line_size(self):
        """Associates all fittings with a Line ID based on collinearity."""
        print("Associating ModelLeader to Lines")
        mleaderIndex, lineIndex = 0, 0
        for mleaderIndex in MLeadersDF.index:
            x2 = MLeadersDF['Starting Vertex Coordinate X'][mleaderIndex]
            y2 = MLeadersDF['Starting Vertex Coordinate Y'][mleaderIndex]
            for lineIndex in LinesDF.index:
                x1 = LinesDF['Start X'][lineIndex]
                y1 = LinesDF['Start Y'][lineIndex]
                x3 = LinesDF['End X'][lineIndex]
                y3 = LinesDF['End Y'][lineIndex]
                if self.is_collinear(x1, y1, x2, y2, x3, y3):
                    MLeadersDF.loc[mleaderIndex, 'Matching Line ID'] = \
                        LinesDF['ID'][lineIndex]
                    MLeadersDF.loc[mleaderIndex, 'Matching Line Layer'] = \
                        LinesDF['Layer'][lineIndex]
                    MLeadersDF.loc[mleaderIndex, 'Matching Line Length'] = \
                        LinesDF['Length'][lineIndex]
                    
    def find_assoc_text(self):
        # Currently O(n^2)
        """Associates text blocks in a sheet with the nearest text block.

        Parameters:
            self: The Sheet object.

        Returns:
            None
        """
        '''TODO: Divide and Conquer Algorithm'''

        print("Associating Texts by Distance")        
        for sheet_name, group in TextsDF.groupby('Sheet'):
            # Iterate over rows within the current sheet group
            for textIndex in group.index:
                minDistance = 999
                minIndex = 0
                x1 = group.loc[textIndex, 'Block X']
                y1 = group.loc[textIndex, 'Block Y']
                for relativeTextIndex in group.index:
                    if relativeTextIndex != textIndex:
                        x2 = group.loc[relativeTextIndex, 'Block X']
                        y2 =group.loc[relativeTextIndex, 'Block Y']
                        distance = self.calc_distance(x1, y1, x2, y2)
                        if distance < minDistance:
                            minDistance = distance
                            minIndex = relativeTextIndex
                TextsDF.loc[textIndex, 'Associated Text ID'] = \
                    TextsDF.loc[minIndex, 'ID']
                TextsDF.loc[textIndex, 'Associated Text String'] = \
                    TextsDF.loc[minIndex, 'Text']

        for textIndex in TextsDF.index:
            if ("DUCTILE" in TextsDF.loc[textIndex, 'Associated Text String'] and 
                "'" in TextsDF.loc[textIndex, 'Text']):
                print(f"\t{TextsDF['Sheet'][textIndex]} : "
                      f"{TextsDF['Text'][textIndex]} of "
                      f"{TextsDF['Associated Text String'][textIndex]}")


    def find_bill_of_materials(self):  
        """Identify the bill of materials in the TextsDF DataFrame, extract the
        relevant information, and store it in the BillOfMaterialsDF DataFrame.

        The function applies a regular expression to identify the rows in the 
        TextsDF DataFrame that contain the 'BILL OF MATERIALS' string. The relevant
        columns are extracted and stored in the BillOfMaterialsDF DataFrame.
        The sheet number is also extracted from the 'Sheet' column and used to
        sort the rows in ascending order. The 'Associated Text String' column is
        then cleaned up and split into a list of comma-separated values.

        Returns:
        None
        """
        def format():
            '''Removes text style symbols attached to strings'''
            BillOfMaterialsDF['Associated Text String'] =            \
                BillOfMaterialsDF['Associated Text String'].apply    \
                (lambda x: re.sub(r'\\A1;', '', x))
            BillOfMaterialsDF['Associated Text String'] =            \
                BillOfMaterialsDF['Associated Text String'].apply    \
                (lambda x: re.sub(r'\\P', ',', x)).str.replace("\t", "-")
            BillOfMaterialsDF.drop(['Sheet #'], axis=1, inplace= True)
            BillOfMaterialsDF['Associated Text String'].apply(lambda x: x[1:].split(","))

        print("\nFinding Bill of Materials")
        BillOfMaterialsFilter = \
            TextsDF['Text'].str.contains('BILL OF MATERIALS')
        BillOfMaterialsDF = \
            TextsDF.loc[BillOfMaterialsFilter, ['Sheet', 'Associated Text String']]
        BillOfMaterialsDF['Sheet #'] = \
            BillOfMaterialsDF['Sheet'].apply(lambda x: int(re.sub(r'\D', '', str(x))))
        BillOfMaterialsDF.sort_values("Sheet #", inplace=True)
        format()

        fittingColumn = 'Associated Text String'
        BillOfMaterialsDF['Fitting List'] = BillOfMaterialsDF[fittingColumn].str.split(',')
        BillOfMaterialsDF = BillOfMaterialsDF.explode('Fitting List')
        BillOfMaterialsDF['Quantity'] = BillOfMaterialsDF['Fitting List'].str.split('-').str[0]
        BillOfMaterialsDF['Fitting'] = BillOfMaterialsDF['Fitting List'].str.split('-').str[1]
        BillOfMaterialsDF.drop(['Fitting List'], axis=1, inplace= True)


    def assign_block_to_sheet(self):
        def is_inside_quadrilateral(a, b, c, d, p):
            # Compute winding numbers
            a, b, c, d = arrange_points_clockwise(a, b, c, d)
            wn_abp = compute_winding_number(a, b, p)
            wn_bcp = compute_winding_number(b, c, p)
            wn_cdp = compute_winding_number(c, d, p)
            wn_dap = compute_winding_number(d, a, p)

            # Check if point is inside quadrilateral
            if (wn_abp == wn_bcp == wn_cdp == wn_dap) and (wn_abp != 0):
                return True
            else:
                return False

        def compute_winding_number(start, end, point):
            # Compute vector from start to point
            v = (point[0] - start[0], point[1] - start[1])
            # Compute vector from start to end
            w = (end[0] - start[0], end[1] - start[1])
            # Compute the 2D cross product of v and w
            cross_product = (v[0] * w[1]) - (v[1] * w[0])
            # Determine winding number based on sign of cross product
            if cross_product > 0:
                return 1
            elif cross_product < 0:
                return -1
            else:
                return 0
            
        def arrange_points_clockwise(a, b, c, d):
            '''Sort four points of a quadrilateral in a CW manner'''
            # Find centroid
            centroid_x = (a[0] + b[0] + c[0] + d[0]) / 4
            centroid_y = (a[1] + b[1] + c[1] + d[1]) / 4
            centroid = (centroid_x, centroid_y)

            # Compute angles between centroid and points
            angles = []
            for point in [a, b, c, d]:
                x_diff = point[0] - centroid[0]
                y_diff = point[1] - centroid[1]
                angle = math.atan2(y_diff, x_diff)
                angles.append(angle)

            # Sort points by angle
            sorted_points = [x for _, x in sorted(zip(angles, [a, b, c, d]))]
            return sorted_points
           
        # Creating new ViewportsDF Column based on associted Viewport to Fitting
        FittingsDF['Matching Viewport ID'] = 'N/A'
        FittingsDF['Matching Viewport Sheet'] = 'N/A'

        # Iteratively gather all coordinates by going through Viewports that are 
        # main BasePlan ModelSpaces
        viewportIndex, fittingIndex = 0, 0
        for viewportIndex in ViewportsDF.index:
            if ViewportsDF['Is BasePlan ModelSpace'][viewportIndex] == True:
                corner1 = (ViewportsDF['ModelSpace Coordinate Corner1 X'][viewportIndex],
                                ViewportsDF['ModelSpace Coordinate Corner1 Y'][viewportIndex])
                corner2 = (ViewportsDF['ModelSpace Coordinate Corner2 X'][viewportIndex],
                                ViewportsDF['ModelSpace Coordinate Corner2 Y'][viewportIndex])
                corner3 = (ViewportsDF['ModelSpace Coordinate Corner3 X'][viewportIndex],
                                ViewportsDF['ModelSpace Coordinate Corner3 Y'][viewportIndex])
                corner4 = (ViewportsDF['ModelSpace Coordinate Corner4 X'][viewportIndex],
                                ViewportsDF['ModelSpace Coordinate Corner4 Y'][viewportIndex])
                
                for fittingIndex in FittingsDF.index:
                    fittingPoint = (FittingsDF['Block X'][fittingIndex],
                                    FittingsDF['Block Y'][fittingIndex])
                    # Only check for Fittings inside ModelSpace
                    if FittingsDF['Sheet'][fittingIndex] == 'Model':
                        if is_inside_quadrilateral(corner1, corner2, corner3, corner4, fittingPoint):
                            FittingsDF.loc[fittingIndex, 'Matching Viewport ID'] = \
                                ViewportsDF['ID'][viewportIndex]
                            FittingsDF.loc[fittingIndex, 'Matching Viewport Sheet'] = \
                                ViewportsDF['Sheet'][viewportIndex]

class PyHelp():
    def __init__(self) -> None:
        self.create_albert_layer()
        self.create_layout_list()
        # self.find_viewports()
        self.find_papersheets()
        self.remove_albert_layer()

    def create_layout_list(self):
        '''Identifies all Layouts in Document and Sorts by Numerical Value'''
        self.layoutList = []
        layouts = wait.wait_for_attribute(doc, "Layouts")
        for layout in layouts:
            self.layoutList.append(layout.Name)
        self.layoutList.sort(key=lambda x:x[2:])

    def create_albert_layer(self):
        '''Creates the layer property to use for all required PyHelp methods'''
        coordinateLayer = doc.layers.Add("AlbertToolLayer")
        coordinateLayer.LayerOn
        coordinateLayer.color = 40


    def remove_albert_layer(self):
        '''Removes all items that use AlbertLayer in the document'''
        print("Removing Albert's Calculation Linework")
        for index, row in LinesDF.iterrows():
        # Check if the layer of the current row is 'AlbertToolLayer'
            try:
                if row['Layer'] == 'AlbertToolLayer':
                    line = acad.ActiveDocument.ObjectIDtoObject(int(row['ID']))
                    line.Delete()
            except:
                print("Failed")


    def validate_viewport(self, entity):
        '''Returns Boolean of Viewport Entity fitting within Sheet and lying within Sheet'''
        # Fits within the bounds of the page
        def is_viewport_size(entity):
            '''Check if Viewport is smaller than the Sheet'''
            PIXELtoINCH = 25.4
            layoutPixelWidth = self.layoutWidth / PIXELtoINCH
            layoutPixelHeight = self.layoutHeight / PIXELtoINCH
            return ((wait.wait_for_attribute(entity, "Height") < layoutPixelHeight) and 
                    (wait.wait_for_attribute(entity, "Width") < layoutPixelWidth))
        
        # Starts and ends within the bounds of the page
        def is_within_page(entity):
            '''Check if Viewport is fully within the Page'''
            corner1 = (wait.wait_for_attribute(entity, "Center")[0] - (abs(wait.wait_for_attribute(entity, "Width")) / 2), 
                        wait.wait_for_attribute(entity, "Center")[1] + (abs(wait.wait_for_attribute(entity, "Height") / 2)))
            corner2 = (wait.wait_for_attribute(entity, "Center")[0] + (abs(wait.wait_for_attribute(entity, "Width")) / 2), 
                        wait.wait_for_attribute(entity, "Center")[1] - (abs(wait.wait_for_attribute(entity, "Height") / 2)))
            if (corner1[0] > 0 and corner1[1]  > 0 and 
                corner2[0] > 0 and corner2[1] > 0):
                return True
            else:
                return False
        return (is_viewport_size(entity) and is_within_page(entity))


    def find_viewports(self):
        # If this is the first sheet, AutoCAD needs to go slow
        doc.ActiveLayer = doc.Layers("AlbertToolLayer")
        # Loop over all layouts and print their names
        print("Finding Viewports")

        for layout in self.layoutList:
            if layout != "Model":
                wait.set_attribute(doc, "ActiveLayout", wait.wait_for_method_return(doc, "Layouts", layout))
                currentLayout = wait.wait_for_attribute(doc, "ActiveLayout")
                validPaperSize = False
                while not validPaperSize:
                    try:
                        self.layoutHeight, self.layoutWidth  = wait.wait_for_method_return(currentLayout, "GetPaperSize")
                        validPaperSize = True
                    except:
                        continue
                
                wait.wait_for_method_return(doc, "SendCommand", "pspace z a  ")

                entities = wait.wait_for_attribute(currentLayout, "Block")
                entitiesCount = wait.wait_for_attribute(entities,"Count")

                i, errorCount = 0, 0
                while i < entitiesCount and errorCount < 3:
                    entity = wait.wait_for_method_return(entities, "Item", i)
                    entityName = wait.wait_for_attribute(entity, "EntityName")
                    if entityName == "AcDbViewport" and self.validate_viewport(entity):
                        vp = Viewport(entity, currentLayout)
                        ViewportsDF.loc[len(ViewportsDF.index)] = [vp.ID, vp.sheet, vp.width, 
                                                                    vp.height, vp.type, vp.isCenter,
                                                                    vp.numFrozenLayers, False,
                                                                    vp.psCorner1[0], vp.psCorner1[1],
                                                                    vp.psCorner2[0], vp.psCorner2[1],
                                                                    vp.msCorner1[0], vp.msCorner1[1], 
                                                                    vp.msCorner2[0], vp.msCorner2[1],
                                                                    vp.psCorner3[0], vp.psCorner3[1],
                                                                    vp.psCorner4[0], vp.psCorner4[1],
                                                                    vp.msCorner3[0], vp.msCorner3[1], 
                                                                    vp.msCorner4[0], vp.msCorner4[1],
                                                                    ]
                        # Group by Sheet and find the Viewport with the fewest frozen layer
                    errorCount = 0
                    i += 1
        self.sort_viewport_dataframe()


    def sort_viewport_dataframe(self):
        _msViewportDF = ViewportsDF[ViewportsDF['Type'] == "Model View"].reset_index(drop=True)
        _msOverlapDF = _msViewportDF[_msViewportDF['Overlaps Center'] == True].reset_index(drop=True)
        indexMinList = _msOverlapDF.groupby('Sheet')['Num of Frozen Layers'].idxmin().to_list()
        for index in indexMinList:
            id = _msOverlapDF['ID'].iloc[index]
            vpIndex = ViewportsDF.index[ViewportsDF['ID'] == id][0]
            ViewportsDF.loc[vpIndex, 'Is BasePlan ModelSpace'] = True


    def find_papersheets(self):
        # , linesDF, FittingsDF, TextsDF
        skipModelSpace = False
        inModelSpace = True
        print("Finding Blocks")
        for layout in self.layoutList:
            print(layout)
            wait.set_attribute(doc, "ActiveLayout", wait.wait_for_method_return(doc, "Layouts", layout))
            currentLayout = wait.wait_for_attribute(doc, "ActiveLayout")
            # Skip Model Space
            if skipModelSpace:
                # On Model Space Sheet
                if layout == "Model":
                    continue
                # Any Paper Sheet
                else:
                    s = Sheet(currentLayout)
                    s.find_blocks()
            # Don't Skip Model Space
            else:
                if layout == "Model":
                    s = Sheet(currentLayout)
                    s.find_blocks()
                    s.find_fitting_size()
                    s.find_appropriate_callout()
                    s.find_line_size()
                    inModelSpace = False
                    continue
                else:
                    s = Sheet(currentLayout)
                    s.find_blocks()

        s.find_assoc_text()
        s.find_bill_of_materials()
        s.assign_block_to_sheet()

# LAUNCH THIS PROGRAM
p = PyHelp()
save_dataframe()