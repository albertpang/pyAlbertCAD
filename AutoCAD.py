import win32com.client
import numpy as np
import pandas as pd

# Importing Class Objects
from line import Line, PolyLine
from fitting import Fitting
from text import Text

acad = win32com.client.Dispatch("AutoCAD.Application")

# iterate through all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes.
class Sheet:
    def __init__(self, layout, linesDF, FittingsDF, TextsDF):
        self.__layout = layout
        self.LinesDF = linesDF
        self.FittingsDF = FittingsDF
        self.TextsDF = TextsDF
    
    def purgeZombieEntity(self):
        for entity in acad.ActiveDocument.ModelSpace:
            if entity.objectName == 'AcDbZombieEntity':
                entity.Erase()        
                print("erased")
                

    def findBlocks(self):
        entities = self.__layout.Block  
        for i in range(entities.Count):
            if i == (entities.Count // 2):
                print ("--- 50% done ---")
            entity = entities.Item(i)

            if entity.ObjectName == 'AcDbLine' and entity.Layer == 'C-PR-WATER':
                if entity.Length > 1:
                    l = Line(entity, self.__layout.name)
                    l.appendToDF(self.LinesDF)
            # Polyline
            elif entity.ObjectName == 'AcDbPolyline' and entity.Layer == 'C-PR-WATER': 
                pl = PolyLine(entity, self.__layout.name, self.LinesDF)
            # Dynamic Blocks                                               
            elif entity.ObjectName == 'AcDbBlockReference' and entity.EffectiveName.startswith("WATER"):
                f = Fitting(entity, self.__layout.name)
                f.appendToDF(self.FittingsDF)
            # Text Items
            elif entity.ObjectName == 'AcDbMText'or entity.ObjectName == 'AcDbText':
                t = Text(entity, self.__layout.name)
                t.appendToDF(self.TextsDF)          
            elif entity.ObjectName == 'AcDbMLeader'and "DUCTILE" in entity.textString:
                pass
                # print(entity.textString)
                # print(entity.DoglegLength)
                # print(entity.GetLeaderLineVertices(0))
                # m = Leader(entity)
                # m.appendToDF(self.LeadersDF)
            else:
                pass
                # print (entity.ObjectName)
    def isCollinear(self, x1, y1, x2, y2, x3, y3) -> bool:
        collinearity = x1*(y3-y2)+x3*(y2-y1)+x2*(y1-y3)
        return (abs(collinearity) < 0.005)
    
        
    def calculateDistance(self, x1, y1, x2, y2) -> int:
        return ((x1 - x2) **2 + (y1 - y2) **2) ** 0.5


    def findFittingSize(self):
        '''Associate all fittings with a Line ID based on collinearity'''
        print("Associating Fittings to Lines")
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
        for textIndex in self.TextsDF.index:
            minDistance = 999
            minIndex = 0
            x1, y1 = self.TextsDF['Block X'][textIndex], self.TextsDF['Block Y'][textIndex]
            for relativeTextIndex in self.TextsDF.index:
                if relativeTextIndex != textIndex:
                    x2, y2  = self.TextsDF['Block X'][relativeTextIndex], self.TextsDF['Block Y'][relativeTextIndex]
                    distance = self.calculateDistance(x1, y1, x2, y2)
                    if distance < minDistance:
                        minDistance = distance
                        minIndex = relativeTextIndex
            self.TextsDF.loc[textIndex, 'Associated Text ID'] = self.TextsDF['ID'][minIndex]
            self.TextsDF.loc[textIndex, 'Associated Text String'] = self.TextsDF['Text'][minIndex]
            
        for textIndex in self.TextsDF.index:
            if "DUCTILE" in self.TextsDF.loc[textIndex, 'Associated Text String'] and "'" in self.TextsDF.loc[textIndex, 'Text'] :
                print(f"{self.TextsDF['Sheet'][textIndex]} : {self.TextsDF['Text'][textIndex]} of {self.TextsDF['Associated Text String'][textIndex]}")
            
    
    def saveDF(self):
        print("SAVING CSVs")
        self.LinesDF.to_csv('LinesCSV')
        print("-Logged to Lines")
        self.FittingsDF.to_csv('FittingsCSV')
        print("--Logged to Fittings")
        self.TextsDF.to_csv('TextsCSV')
        print("---Logged to Texts")



def findPaperSheets(linesDF, FittingsDF, TextsDF):
    inModelSpace = True
    for layout in acad.activeDocument.layouts:
        print(layout.Name)
        if inModelSpace:
            s = Sheet(layout, linesDF, FittingsDF, TextsDF)
            s.findBlocks()
            s.findFittingSize()
            inModelSpace = False
        else:
            s = Sheet(layout, linesDF, FittingsDF, TextsDF)
            s.findBlocks()
    s.findAssociatedText()
    s.saveDF()


LinesDF = pd.DataFrame(columns=['ID', 'Sheet', 'Block Description', 'Start X', 
                                    'Start Y', 'End X', 'End Y', 'Length', 'Slope'])
FittingsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Block Description', 
                                'Block X', 'Block Y', 'Matching Line ID', 
                                'Matching Line Length'])
TextsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Text', 'Block X', 'Block Y', 
                                'Associated Text ID', 'Associated Text String'])        

findPaperSheets(LinesDF, FittingsDF, TextsDF)
# sheet.purgeZombieEntity()
# sheet.findFittingSize()
# sheet.saveDF()


        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

