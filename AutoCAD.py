import win32com.client
from pywintypes import com_error
import wait
import numpy as np
from ACAD_DataTypes import APoint
import pandas as pd
import time
import re

# Importing Class Objects
from entity import Fitting, Line, PolyLine, Text, Viewport

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

LinesDF = pd.DataFrame(columns=['ID', 'Sheet', 'Layer', 'Start X', 
                                    'Start Y', 'End X', 'End Y', 'Length', 'Slope'])
FittingsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Block Description', 
                                'Block X', 'Block Y', 'Matching Line ID', 
                                'Matching Line Length'])
TextsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Text', 'Block X', 'Block Y', 
                                'Associated Text ID', 'Associated Text String'])
        
ViewportsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Width', 'Height', 'Type', 'Overlaps Center', 
                                    'Num of Frozen Layers', 'Is BasePlan ModelSpace',
                                    'PaperSpace Coordinate Corner1 X', 'PaperSpace Coordinate Corner1 Y',
                                    'PaperSpace Coordinate Corner2 X', 'PaperSpace Coordinate Corner2 Y',
                                    'ModelSpace Coordinate Corner1 X', 'ModelSpace Coordinate Corner1 Y',
                                    'ModelSpace Coordinate Corner2 X', 'ModelSpace Coordinate Corner2 Y',
                                    'PaperSpace Coordinate Corner3 X', 'PaperSpace Coordinate Corner3 Y',
                                    'PaperSpace Coordinate Corner4 X', 'PaperSpace Coordinate Corner4 Y',
                                    'ModelSpace Coordinate Corner3 X', 'ModelSpace Coordinate Corner3 Y',
                                    'ModelSpace Coordinate Corner4 X', 'ModelSpace Coordinate Corner4 Y',
                                    ])

BillOfMaterialsDF = pd.DataFrame(columns=['Sheet', 'Associated Text String'])

class Sheet:
    """
    Initialize the Sheet object with layout, lines data frame, 
    fittings data frame, texts data frame, and bill of materials data frame.

    Parameters:
    -----------
    layout : Layout
        An AutoCAD Layout object representing the current drawing layout.
    """

    def __init__(self, layout):
        self.__layout = layout

    def findBlocks(self):
        """ Finds all the blocks in the current drawing and classifies them. 
        The information about these objects is then stored in respective data
        frames.

        Raises:
            Exception: If there is an error while iterating over the entities 
            in the layout.
        """
        entities = wait.wait_for_attribute(self.__layout, "Block")
        entitiesCount = wait.wait_for_attribute(entities, "Count") 
        i, errorCount = 0, 1
        while i < entitiesCount and errorCount <= 3:
            # if i == (entitiesCount // 2):
            #     print ("--- 50% done ---")
            try:
                entity = wait.wait_for_method_return(entities, "Item", i)
                entityObjectName = wait.wait_for_attribute(entity, "ObjectName")

                # Line Object
                if entityObjectName == 'AcDbLine': # and entity.Layer == 'C-PR-WATER':
                    if entity.Length > 1:
                        l = Line(entity, self.__layout.name, False)
                        LinesDF.loc[len(LinesDF.index)] = [l.ID, l.sheet, l.layer, 
                                                           l.startX, l.startY, l.endX, 
                                                           l.endY, l.length, l.slope]

                # Polyline
                elif entityObjectName == 'AcDbPolyline': # and entity.Layer == 'C-PR-WATER':
                    if entity.Length > 1:
                        # If there are only two coordinates, convert Polyline into a line
                        if len(entity.Coordinates) == 4:
                            l = Line(entity, self.__layout.name, True)
                            LinesDF.loc[len(LinesDF.index)] = [l.ID, l.sheet, l.layer, 
                                                               l.startX, l.startY, l.endX, 
                                                               l.endY, l.length, l.slope]
                        # Otherwise, it is an authentic Polyline
                        else:
                            pl = PolyLine(entity, self.__layout.name, LinesDF)
                elif entityObjectName == 'AcDbBlockReference': # and entity.Name.startswith("WATER"):
                    f = Fitting(entity, self.__layout.name)
                    FittingsDF.loc[len(FittingsDF.index)] = [f.ID, f.sheet, f.BlockName, 
                                                             f.locationX, f.locationY, "N/A", "N/A"]
                # Text Items
                elif (entityObjectName == 'AcDbMText' or entity.ObjectName == 'AcDbText'):
                    t = Text(entity, self.__layout.name)
                    TextsDF.loc[len(TextsDF.index)] = [t.ID, t.sheet, t.text, 
                                                       t.locationX, t.locationY, 
                                                       "N/A", "N/A"]
                # MLeader Object
                elif entityObjectName == 'AcDbMLeader': # and "DUCTILE" in entity.textString:
                    # t = Text(entity, self.__layout.name)
                    # t.appendToDF(self.TextsDF)
                    pass
                elif entityObjectName == 'AcDbViewport':
                    pass
                errorCount = 0 
                i += 1
            except com_error as e:
                time.sleep(0.5)
                print("here")
            except Exception as e:
                errorCount += 1
                print(f"\tAttempt: {errorCount}", i, entityObjectName, e)
    
    
    def isCollinear(self, x1, y1, x2, y2, x3, y3) -> bool:
        """ Checks if three points are collinear.

        Parameters:
            x1 (float): The x coordinate of the first point.
            y1 (float): The y coordinate of the first point.
            x2 (float): The x coordinate of the second point.
            y2 (float): The y coordinate of the second point.
            x3 (float): The x coordinate of the third point.
            y3 (float): The y coordinate of the third point.

        Returns:
            bool: True if the points are collinear, False otherwise.
        """
        collinearity = x1*(y3-y2)+x3*(y2-y1)+x2*(y1-y3)
        return (abs(collinearity) < 0.005)
    
        
    def calculateDistance(self, x1, y1, x2, y2) -> float:
        """Calculates the Euclidean distance between two points in a 2D space.

        Parameters:
            self: The Sheet object.
            x1 (int): The X coordinate of the first point.
            y1 (int): The Y coordinate of the first point.
            x2 (int): The X coordinate of the second point.
            y2 (int): The Y coordinate of the second point.

        Returns:
            int: The Euclidean distance between the two points.
        """
        return ((x1 - x2) **2 + (y1 - y2) **2) ** 0.5


    def findFittingSize(self):
        """Associates all fittings with a Line ID based on collinearity.

        Parameters:
            self: The Sheet object.

        Returns:
            None
        """
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
                if self.isCollinear(x1, y1, x2, y2, x3, y3):
                    FittingsDF.loc[fittingIndex, 'Matching Line ID'] = \
                        LinesDF['ID'][lineIndex]
                    FittingsDF.loc[fittingIndex, 'Matching Line Length'] = \
                        LinesDF['Length'][lineIndex]
    
    
    def findAssociatedText(self):
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
                        distance = self.calculateDistance(x1, y1, x2, y2)
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


    def findBillOfMaterials(self):  
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


    def assignBlockToSheet(self):
        def liesWithin(c1, c2, c3, c4, fittingPoint):
            # I have absolutely no clue why my x
            x, y = fittingPoint
            minX = min(abs(c1[0]), abs(c2[0]), abs(c3[0]), abs(c4[0]))
            maxX = max(abs(c1[0]), abs(c2[0]), abs(c3[0]), abs(c4[0]))
            minY = min(abs(c1[1]), abs(c2[1]), abs(c3[1]), abs(c4[1]))
            maxY = max(abs(c1[1]), abs(c2[1]), abs(c3[1]), abs(c4[1]))
            insideX = (minX <= x) and (x <= maxX)
            insideY = (minY <= y) and (y <= maxY)
            return (insideX and insideY)
        
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
                        if liesWithin(corner1, corner2, corner3, corner4, fittingPoint):
                            FittingsDF.loc[fittingIndex, 'Matching Viewport ID'] = \
                                ViewportsDF['ID'][viewportIndex]
                            FittingsDF.loc[fittingIndex, 'Matching Viewport Sheet'] = \
                                ViewportsDF['Sheet'][viewportIndex]

class PyHelp():
    def __init__(self) -> None:
        self.createAlbertLayer()
        self.findViewports()
        self.findPaperSheets()
        self.removeAlbertTool()


    def createAlbertLayer(self):
        coordinateLayer = doc.layers.Add("AlbertToolLayer")
        coordinateLayer.LayerOn
        coordinateLayer.color = 40


    def removeAlbertTool(self):
        print("Removing Albert's Calculation Linework")
        for index, row in LinesDF.iterrows():
        # Check if the layer of the current row is 'AlbertToolLayer'
            try:
                if row['Layer'] == 'AlbertToolLayer':
                    line = acad.ActiveDocument.ObjectIDtoObject(int(row['ID']))
                    line.Delete()
            except:
                print("Failed")


    def validateViewport(self, entity, layout):
        # Fits within the bounds of the page
        def isViewPortSize(layout, entity):
            PIXELtoINCH = 25.4
            height, width = wait.wait_for_method_return(layout, "GetPaperSize")
            width /= PIXELtoINCH
            height /= PIXELtoINCH
            return (wait.wait_for_attribute(entity, "Height") < height and 
                    wait.wait_for_attribute(entity, "Width") < width)
        # Starts and ends within the bounds of the page
        def isWithinPage(entity):
            corner1 = (entity.Center[0] - (abs(entity.Width) / 2), 
                                entity.Center[1] + (abs(entity.Height) / 2))
            corner2 = (entity.Center[0] + (abs(entity.Width) / 2), 
                                entity.Center[1] - (abs(entity.Height) / 2))
            if (corner1[0] > 0 and corner1[1]  > 0 and 
                corner2[0] > 0 and corner2[1] > 0):
                return True
            else:
                return False
            
        return (isViewPortSize(layout, entity) and isWithinPage(entity))


    def findViewports(self):
        # If this is the first sheet, AutoCAD needs to go slow
        layouts = wait.wait_for_attribute(doc, "Layouts")
        numLayouts = wait.wait_for_attribute(layouts, "Count")
        doc.ActiveLayer = doc.Layers("AlbertToolLayer")
        # Loop over all layouts and print their names
        print("Finding Viewports")
        for layout in layouts:
            print(wait.wait_for_attribute(layout, "Name"))
            if layout.Name != "Model":
                doc.ActiveLayout = doc.Layouts(wait.wait_for_attribute(layout, "Name"))
                paperFlag = False
                while paperFlag == False:
                    try: 
                        doc.SendCommand("pspace z a  ")
                        paperFlag = True
                    except:
                        break
                entities = wait.wait_for_attribute(layout, "Block")
                entitiesCount = wait.wait_for_attribute(entities,"Count")
                i, errorCount = 0, 0
                while i < entitiesCount and errorCount < 3:
                    try:
                        entity = wait.wait_for_method_return(entities, "Item", i)
                        entityName = wait.wait_for_attribute(entity, "EntityName")
                        if entityName == "AcDbViewport" and self.validateViewport(entity, layout):
                            vp = Viewport(entity, layout)
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
                    except Exception as e:
                        errorCount += 1
                        print(f"\tAttempt: {errorCount}", e)

        self.sortViewportDF()


    def sortViewportDF(self):
        _msViewportDF = ViewportsDF[ViewportsDF['Type'] == "Model View"].reset_index(drop=True)
        _msOverlapDF = _msViewportDF[_msViewportDF['Overlaps Center'] == True].reset_index(drop=True)
        indexMinList = _msOverlapDF.groupby('Sheet')['Num of Frozen Layers'].idxmin().to_list()
        for index in indexMinList:
            id = _msOverlapDF['ID'].iloc[index]
            vpIndex = ViewportsDF.index[ViewportsDF['ID'] == id][0]
            ViewportsDF.loc[vpIndex, 'Is BasePlan ModelSpace'] = True


    def findPaperSheets(self):
        # , linesDF, FittingsDF, TextsDF
        skipModelSpace = False
        inModelSpace = True
        print("Finding Blocks")
        for layout in acad.activeDocument.layouts:
            # Skip Model Space
            if skipModelSpace:
                # On Model Space Sheet
                if inModelSpace:
                    inModelSpace = False
                    continue
                # Any Paper Sheet
                else:
                    s = Sheet(layout)
                    s.findBlocks()
            # Don't Skip Model Space
            else:
                if inModelSpace:
                    s = Sheet(layout)
                    s.findBlocks()
                    s.findFittingSize()
                    inModelSpace = False
                    continue
                else:
                    s = Sheet(layout)
                    s.findBlocks()

        s.findAssociatedText()
        s.findBillOfMaterials()
        s.assignBlockToSheet()


def saveDF():
    """Save the DataFrame objects to CSV files.

    The function saves each of the four DataFrame objects to a separate CSV file.
    The file names are hardcoded as 'LinesCSV', 'FittingsCSV', 'TextsCSV', and
    'BillOfMaterialsCSV', respectively. The function logs a message to the console
    after each file is saved to indicate which DataFrame object was saved.

    Returns:
    None
    """
    print("Saving CSVs")
    LinesDF.to_csv('LinesCSV')
    print("-Logged to Lines")
    FittingsDF.to_csv('FittingsCSV')
    print("--Logged to Fittings")
    TextsDF.to_csv('TextsCSV')
    print("---Logged to Texts")
    BillOfMaterialsDF.to_csv('BillOfMaterialsCSV')
    print("----Logged to BOM")
    ViewportsDF.to_csv('ViewportsCSV')
    print("-----Logged to ViewportsCSV")

p = PyHelp()
saveDF()