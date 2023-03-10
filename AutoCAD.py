import win32com.client
import numpy as np
from ACAD_DataTypes import APoint
import pandas as pd
import time
import re

# Importing Class Objects
from line import Line, PolyLine
from fitting import Fitting
from text import Text

acad = win32com.client.Dispatch("AutoCAD.Application")

class Sheet:
    """
    Initialize the Sheet object with layout, lines data frame, 
    fittings data frame, texts data frame, and bill of materials data frame.

    Parameters:
    -----------
    layout : Layout
        An AutoCAD Layout object representing the current drawing layout.

    linesDF : pandas.DataFrame
        A data frame containing information about lines in the current drawing.

    FittingsDF : pandas.DataFrame
        A data frame containing information about fittings in the current drawing.

    TextsDF : pandas.DataFrame
        A data frame containing information about texts in the current drawing.

    BoMDF : pandas.DataFrame
        A data frame containing information about bill of materials 
        in the current drawing.
    """
    def __init__(self, layout, linesDF, FittingsDF, TextsDF, BoMDF):
        self.__layout = layout
        self.LinesDF = linesDF
        self.FittingsDF = FittingsDF
        self.TextsDF = TextsDF
        self.BillOfMaterialsDF = BoMDF


    def findBlocks(self):
        """
        Finds all the blocks in the current drawing and classifies them as lines, 
        polylines, dynamic blocks, texts or mleader objects. The information 
        about these objects is then stored in respective data frames.

        Raises:
            Exception: If there is an error while iterating over the entities 
            in the layout.
        """
        acUtils = acad.Activedocument.Utility
        entities = self.__layout.Block
        entitiesCount = entities.Count 
        i, errorCount = 0, 1
        while i < entitiesCount and errorCount <= 3:
            if i == (entitiesCount // 2):
                print ("--- 50% done ---")
                
            try:
                entity = entities.Item(i)
                entityObjectName = entity.ObjectName

                # Line Object
                if entityObjectName == 'AcDbLine': # and entity.Layer == 'C-PR-WATER':
                    if entity.Length > 1:
                        l = Line(entity, self.__layout.name, False)
                        l.appendToDF(self.LinesDF)

                # Polyline
                elif entityObjectName == 'AcDbPolyline': # and entity.Layer == 'C-PR-WATER':
                    if entity.Length > 1:
                        # If there are only two coordinates, convert Polyline into a line
                        if len(entity.Coordinates) == 4:
                            l = Line(entity, self.__layout.name, True)
                            l.appendToDF(self.LinesDF)
                        # Otherwise, it is an authentic Polyline
                        else:
                            pl = PolyLine(entity, self.__layout.name, self.LinesDF)

                # Dynamic Blocks
                elif entityObjectName == 'AcDbBlockReference': # and entity.Name.startswith("WATER"):
                    f = Fitting(entity, self.__layout.name)
                    f.appendToDF(self.FittingsDF)
        
                # Text Items
                elif (entityObjectName == 'AcDbMText' or entity.ObjectName == 'AcDbText'):
                    t = Text(entity, self.__layout.name)
                    t.appendToDF(self.TextsDF)          

                # MLeader Object
                elif entityObjectName == 'AcDbMLeader': # and "DUCTILE" in entity.textString:
                    # t = Text(entity, self.__layout.name)
                    # t.appendToDF(self.TextsDF)
                    pass
                elif entityObjectName == 'AcDbViewport':
                    pass


                    
                else:
                    pass

                i += 1

            except Exception as e:
                errorCount += 1
                print(f"Attempt Count: {errorCount}", i, entityObjectName, e)
    
    



    # def translateCoordinates(self):
        # tp = APoint(18.50201834, 17.10606439)
        # print(acad.ActiveDocument.Utility.TranslateCoordinates(tp, 2, 0, False))
        # "2706723.25739601, 270394.52176793"

        # psCenter = (entity.Center[0], entity.Center[1])
        # psleftTopCorner = (entity.Center[0] - (abs(entity.Width) / 2), entity.Center[1] + (abs(entity.Height) / 2))
        # psrightBotCorner = (entity.Center[0] + (abs(entity.Width) / 2), entity.Center[1] - (abs(entity.Height) / 2))
        
        # psleftTopCornerPoint = APoint(psleftTopCorner[0], psleftTopCorner[1])
        # psrightBotCornerPoint = APoint(psrightBotCorner[0], psrightBotCorner[1])
        # psCenterPoint = APoint(entity.Center[0], entity.Center[1])

        # # Translate Coordinates only works when in Model Space
        # acad.ActiveDocument.SendCommand("_Model ")
        # wcsCenter = acad.ActiveDocument.Utility.TranslateCoordinates(psCenterPoint, 1, 0, False)
        # wcsleftTop = acad.ActiveDocument.Utility.TranslateCoordinates(psleftTopCornerPoint, 1,0, False)
        # wcsrightBot = acad.ActiveDocument.Utility.TranslateCoordinates(psrightBotCornerPoint, 1, 0, False)
                            # ViewportsDF.loc[len(ViewportsDF.index)] = [entity.ObjectID, self.__layout.name,
                    #                                            entity.Width, entity.Height,
                    #                                            psCenter, wcsCenter, 
                    #                                            psleftTopCorner, psrightBotCorner,
            #                                            wcsleftTop, wcsrightBot]

    def isCollinear(self, x1, y1, x2, y2, x3, y3) -> bool:
        """
        Checks if three points are collinear.

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
        for fittingIndex in self.FittingsDF.index:
            x2 = self.FittingsDF['Block X'][fittingIndex]
            y2 = self.FittingsDF['Block Y'][fittingIndex]
            for lineIndex in self.LinesDF.index:
                x1 = self.LinesDF['Start X'][lineIndex]
                y1 = self.LinesDF['Start Y'][lineIndex]
                x3 = self.LinesDF['End X'][lineIndex]
                y3 = self.LinesDF['End Y'][lineIndex]
                if self.isCollinear(x1, y1, x2, y2, x3, y3):
                    self.FittingsDF.loc[fittingIndex, 'Matching Line ID'] = \
                        self.LinesDF['ID'][lineIndex]
                    self.FittingsDF.loc[fittingIndex, 'Matching Line Length'] = \
                        self.LinesDF['Length'][lineIndex]
    
    
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
        for sheet_name, group in self.TextsDF.groupby('Sheet'):
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
                self.TextsDF.loc[textIndex, 'Associated Text ID'] = \
                    self.TextsDF.loc[minIndex, 'ID']
                self.TextsDF.loc[textIndex, 'Associated Text String'] = \
                    self.TextsDF.loc[minIndex, 'Text']

        for textIndex in self.TextsDF.index:
            if ("DUCTILE" in self.TextsDF.loc[textIndex, 'Associated Text String'] and 
                "'" in self.TextsDF.loc[textIndex, 'Text']):
                print(f"{self.TextsDF['Sheet'][textIndex]} : "
                      f"{self.TextsDF['Text'][textIndex]} of "
                      f"{self.TextsDF['Associated Text String'][textIndex]}")


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
            self.BillOfMaterialsDF['Associated Text String'] =            \
                self.BillOfMaterialsDF['Associated Text String'].apply    \
                (lambda x: re.sub(r'\\A1;', '', x))
            self.BillOfMaterialsDF['Associated Text String'] =            \
                self.BillOfMaterialsDF['Associated Text String'].apply    \
                (lambda x: re.sub(r'\\P', ',', x)).str.replace("\t", "-")
            self.BillOfMaterialsDF.drop(['Sheet #'], axis=1, inplace= True)
            
            
            
            self.BillOfMaterialsDF['Associated Text String'].apply(lambda x: x[1:].split(","))

        print("\nFinding Bill of Materials")
        BillOfMaterialsFilter = \
            self.TextsDF['Text'].str.contains('BILL OF MATERIALS')
        self.BillOfMaterialsDF = \
            self.TextsDF.loc[BillOfMaterialsFilter, ['Sheet', 'Associated Text String']]
        self.BillOfMaterialsDF['Sheet #'] = \
            self.BillOfMaterialsDF['Sheet'].apply(lambda x: int(re.sub(r'\D', '', str(x))))
        self.BillOfMaterialsDF.sort_values("Sheet #", inplace=True)
        format()

        fittingColumn = 'Associated Text String'
        self.BillOfMaterialsDF['Fitting List'] = self.BillOfMaterialsDF[fittingColumn].str.split(',')
        self.BillOfMaterialsDF = self.BillOfMaterialsDF.explode('Fitting List')
        self.BillOfMaterialsDF['Quantity'] = self.BillOfMaterialsDF['Fitting List'].str.split('-').str[0]
        self.BillOfMaterialsDF['Fitting'] = self.BillOfMaterialsDF['Fitting List'].str.split('-').str[1]
        self.BillOfMaterialsDF.drop(['Fitting List'], axis=1, inplace= True)

        

        # 
        # self.BillOfMaterialsDF.assign(fittingColumn = 

        
    def saveDF(self):
        """Save the DataFrame objects to CSV files.

        The function saves each of the four DataFrame objects to a separate CSV file.
        The file names are hardcoded as 'LinesCSV', 'FittingsCSV', 'TextsCSV', and
        'BillOfMaterialsCSV', respectively. The function logs a message to the console
        after each file is saved to indicate which DataFrame object was saved.

        Returns:
        None
        """
        print("Saving CSVs")
        self.LinesDF.to_csv('LinesCSV')
        print("-Logged to Lines")
        self.FittingsDF.to_csv('FittingsCSV')
        print("--Logged to Fittings")
        self.TextsDF.to_csv('TextsCSV')
        print("---Logged to Texts")
        self.BillOfMaterialsDF.to_csv('BillOfMaterialsCSV')
        print("----Logged to BOM")
        ViewportsDF.to_csv('ViewportsCSV')

    
def purgeZombieEntity():
    """ -- PROXY Errors due to WINCOM -- Obsolete Code -- """
    """Delete any zombie AutoCAD entities from the Modelspace.

    The function iterates over all the entities in the Modelspace of the active
    AutoCAD document and prints their object names to the console. This is useful
    for identifying any entities that may have been left behind by a program that
    terminated unexpectedly or otherwise failed to clean up after itself. The 
    function does not delete any entities.

    Returns:
    None
    """
    i = 0
    db = acad.ActiveDocument.Modelspace
   
    for i in range(db.count):
        print(db.Item(i).ObjectName)      

def createViewportLayer():
    coordinateLayer = acad.ActiveDocument.layers.Add("AlbertToolLayer")
    coordinateLayer.LayerOn
    coordinateLayer.color = 40
        # Get the active document
    doc = acad.ActiveDocument

    # Get the Layouts collection for the current document
    layouts = doc.Layouts

    # Loop over all layouts and print their names
    for layout in layouts:
        if layout.Name != "Model":
            print(layout.Name)
            doc.ActiveLayout = doc.Layouts(layout.Name)
            
            for entity in acad.ActiveDocument.ActiveLayout.Block:
                if entity.EntityName == "AcDbViewport" and fitPage(layout, entity):
                    psleftTopCorner = (entity.Center[0] - (abs(entity.Width) / 2), 
                                        entity.Center[1] + (abs(entity.Height) / 2))
                    psrightBotCorner = (entity.Center[0] + (abs(entity.Width) / 2), 
                                        entity.Center[1] - (abs(entity.Height) / 2))
                    psleftTopCornerPoint = APoint(psleftTopCorner[0], psleftTopCorner[1])
                    psrightBotCornerPoint = APoint(psrightBotCorner[0], psrightBotCorner[1])
                    acad.ActiveDocument.PaperSpace.AddLine(psleftTopCornerPoint, psrightBotCornerPoint)
                    acad.SendCommand("_select\n")
                    acad.SendCommand("_single\n")
                    acad.SendCommand("_pickfirst\n")
                    acad.SendCommand("_circle\n")
                    # doc.SendCommand("CHSPACE ")
                    continue
            for entity in acad.ActiveDocument.ActiveLayout.Block:
                if entity.EntityName == "AcDbLine" and entity.Layer == "AlbertToolLayer":
                    entity.Highlight()
                    entity.Update()
                    x, y = (entity.StartPoint[0], entity.StartPoint[1])
                    doc.SendCommand("_line " + str(x) + ',' + str(y) + " ")

def fitPage(layout, entity):
    height, width = layout.GetPaperSize()
    width /= 25.4
    height /= 25.4

    return (entity.Height < height and entity.Width < width)

def findPaperSheets(linesDF, FittingsDF, TextsDF):
    skipModelSpace = True
    inModelSpace = True

    for layout in acad.activeDocument.layouts:
        print(layout.Name)
        # Skip Model Space
        if skipModelSpace:
            # On Model Space Sheet
            if inModelSpace:
                inModelSpace = False
                continue
            # Any Paper Sheet
            else:
                s = Sheet(layout, linesDF, FittingsDF, TextsDF, BillOfMaterialsDF)
                
                s.findBlocks()
        # Don't Skip Model Space
        else:
            if inModelSpace:
                s = Sheet(layout, linesDF, FittingsDF, TextsDF, BillOfMaterialsDF)
                s.findBlocks()
                s.findFittingSize()
                inModelSpace = False
                continue
            else:
                s = Sheet(layout, linesDF, FittingsDF, TextsDF, BillOfMaterialsDF)
                s.findBlocks()

    s.findAssociatedText()
    s.findBillOfMaterials()
    s.saveDF()


LinesDF = pd.DataFrame(columns=['ID', 'Sheet', 'Block Description', 'Start X', 
                                    'Start Y', 'End X', 'End Y', 'Length', 'Slope'])
FittingsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Block Description', 
                                'Block X', 'Block Y', 'Matching Line ID', 
                                'Matching Line Length'])
TextsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Text', 'Block X', 'Block Y', 
                                'Associated Text ID', 'Associated Text String'])        
BillOfMaterialsDF = pd.DataFrame(columns=['Sheet', 'Associated Text String'])
ViewportsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Width', 'Height', 'Center Point',
                                    'ms Center Point', 'topLeftBoundary', 'botRightBoundary',
                                    'ms topLeftBoundary', 'ms botRightBoundary'])

createViewportLayer()
findPaperSheets(LinesDF, FittingsDF, TextsDF)
