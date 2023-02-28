import win32com.client
import numpy as np
import pandas as pd
import time
import re

# Importing Class Objects
from line import Line, PolyLine
from fitting import Fitting
from text import Text

acad = win32com.client.Dispatch("AutoCAD.Application")

# iterate through all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes.
class Sheet:
    def __init__(self, layout, linesDF, FittingsDF, TextsDF, BoMDF):
        self.__layout = layout
        self.LinesDF = linesDF
        self.FittingsDF = FittingsDF
        self.TextsDF = TextsDF
        self.BillOfMaterialsDF = BoMDF


    def findBlocks(self):
        entities = self.__layout.Block
        entitiesCount = entities.Count 
        i = 0
        while i < entitiesCount:
            if i == (entitiesCount // 2):
                print ("--- 50% done ---")
            try:
                entity = entities.Item(i)
                entityObjectName = entity.ObjectName
                # print(i, entityObjectName)

                # Line Object
                if entityObjectName == 'AcDbLine': # and entity.Layer == 'C-PR-WATER':
                    if entity.Length > 1:
                        l = Line(entity, self.__layout.name)
                        l.appendToDF(self.LinesDF)

                # Polyline
                elif entityObjectName == 'AcDbPolyline': # and entity.Layer == 'C-PR-WATER':
                    pl = PolyLine(entity, self.__layout.name, self.LinesDF)

                # Dynamic Blocks
                elif entityObjectName == 'AcDbBlockReference': # and entity.Name.startswith("WATER"):
                    f = Fitting(entity, self.__layout.name)
                    f.appendToDF(self.FittingsDF)
        
                # Text Items
                elif entityObjectName == 'AcDbMText' or entity.ObjectName == 'AcDbText':
                    t = Text(entity, self.__layout.name)
                    t.appendToDF(self.TextsDF)          

                # MLeader Object
                elif entityObjectName == 'AcDbMLeader': # and "DUCTILE" in entity.textString:
                    # t = Text(entity, self.__layout.name)
                    # t.appendToDF(self.TextsDF)
                    pass

                else:
                    pass
                i += 1

            except Exception as e:
                print(i, entityObjectName, e)


    def isCollinear(self, x1, y1, x2, y2, x3, y3) -> bool:
        collinearity = x1*(y3-y2)+x3*(y2-y1)+x2*(y1-y3)
        return (abs(collinearity) < 0.005)
    
        
    def calculateDistance(self, x1, y1, x2, y2) -> int:
        return ((x1 - x2) **2 + (y1 - y2) **2) ** 0.5


    def findFittingSize(self):
        '''Associate all fittings with a Line ID based on collinearity'''
        print("Associating ModelSpace Fittings to Lines")
        fittingIndex, lineIndex = 0, 0
        for fittingIndex in self.FittingsDF.index:
            x2, y2 = self.FittingsDF['Block X'][fittingIndex], self.FittingsDF['Block Y'][fittingIndex]
            for lineIndex in self.LinesDF.index:
                x1, y1  = self.LinesDF['Start X'][lineIndex], self.LinesDF['Start Y'][lineIndex]
                x3, y3 = self.LinesDF['End X'][lineIndex], self.LinesDF['End Y'][lineIndex]
                if self.isCollinear(x1, y1, x2, y2, x3, y3):
                    self.FittingsDF.loc[fittingIndex, 'Matching Line ID'] = self.LinesDF['ID'][lineIndex]
                    self.FittingsDF.loc[fittingIndex, 'Matching Line Length'] = self.LinesDF['Length'][lineIndex]
    
    
    def findAssociatedText(self):
        # Currently O(n^2)
        '''TODO: Divide and Conquer Algorithm'''
        print("Associating Texts by Distance")        
        for sheet_name, group in self.TextsDF.groupby('Sheet'):
            # Iterate over rows within the current sheet group
            for textIndex in group.index:
                minDistance = 999
                minIndex = 0
                x1, y1 = group.loc[textIndex, 'Block X'], group.loc[textIndex, 'Block Y']
                for relativeTextIndex in group.index:
                    if relativeTextIndex != textIndex:
                        x2, y2 = group.loc[relativeTextIndex, 'Block X'], group.loc[relativeTextIndex, 'Block Y']
                        distance = self.calculateDistance(x1, y1, x2, y2)
                        if distance < minDistance:
                            minDistance = distance
                            minIndex = relativeTextIndex
                self.TextsDF.loc[textIndex, 'Associated Text ID'] = self.TextsDF.loc[minIndex, 'ID']
                self.TextsDF.loc[textIndex, 'Associated Text String'] = self.TextsDF.loc[minIndex, 'Text']

        for textIndex in self.TextsDF.index:
            if "DUCTILE" in self.TextsDF.loc[textIndex, 'Associated Text String'] and "'" in self.TextsDF.loc[textIndex, 'Text'] :
                print(f"{self.TextsDF['Sheet'][textIndex]} : {self.TextsDF['Text'][textIndex]} of {self.TextsDF['Associated Text String'][textIndex]}")

    def findBillOfMaterials(self):
        def format():
            self.BillOfMaterialsDF['Associated Text String'] =            \
                self.BillOfMaterialsDF['Associated Text String'].apply    \
                (lambda x: re.sub(r'\\A1;|\\P', ',', x)).str.replace("\t", "-")
            
            self.BillOfMaterialsDF['Associated Text String'].apply(lambda x: x[1:].split(","))

        print("\nFinding Bill of Materials")
        BillOfMaterialsFilter = self.TextsDF['Text'].str.contains('BILL OF MATERIALS')
        self.BillOfMaterialsDF = self.TextsDF.loc[BillOfMaterialsFilter, ['Sheet', 'Associated Text String']]
        self.BillOfMaterialsDF['Sheet_no_letters'] = self.BillOfMaterialsDF['Sheet'].apply(lambda x: re.sub(r'\D', '', str(x)))
        self.BillOfMaterialsDF.sort_values("Sheet_no_letters", inplace=True)
        format()

        
    def saveDF(self):
        print("Saving CSVs")
        self.LinesDF.to_csv('LinesCSV')
        print("-Logged to Lines")
        self.FittingsDF.to_csv('FittingsCSV')
        print("--Logged to Fittings")
        self.TextsDF.to_csv('TextsCSV')
        print("---Logged to Texts")
        self.BillOfMaterialsDF.to_csv('BillOfMaterialsCSV')
        print("----Logged to BOM")

    
def purgeZombieEntity():
    i = 0
    db = acad.ActiveDocument.Modelspace
    for i in range(db.count):
        print(db.Item(i).ObjectName)      


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

findPaperSheets(LinesDF, FittingsDF, TextsDF)



        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

