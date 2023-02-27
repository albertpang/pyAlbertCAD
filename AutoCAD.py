import win32com.client
import numpy as np
import pandas as pd
from line import Line
from fitting import Fitting
from text import Text

acad = win32com.client.Dispatch("AutoCAD.Application")

# iterate through all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes.
class Sheet:
    def __init__(self):
        self.LinesDF = pd.DataFrame(columns=['ID', 'Block Description', 'Start X', 
                                            'Start Y', 'End X', 'End Y', 'Slope'])
        self.FittingsDF = pd.DataFrame(columns=['ID', 'Block Description', 
                                        'Block X', 'Block Y', 'Matching Line ID'])
        self.TextsDF = pd.DataFrame(columns=['ID', 'Text', 
                                        'Block X', 'Block Y', 'Matching Line ID'])
        

    def findBlocks(self):
        for i, entity in enumerate(acad.ActiveDocument.Modelspace):
            # print (i, entity.ObjectName)
            if entity.ObjectName == 'AcDbLine' and entity.Layer == 'C-PR-WATER':
                l = Line(entity)
                l.appendToDF(self.LinesDF)
                                                            
            if entity.ObjectName == 'AcDbBlockReference' and entity.EffectiveName.startswith("WATER"):
                f = Fitting(entity)
                f.appendToDF(self.FittingsDF)
            
            if entity.ObjectName == 'AcDbMText':
                t = Text(entity)
                t.appendToDF(self.TextsDF)
            
            if entity.ObjectName == 'AcDbMLeader'and "DUCTILE" in entity.textString:
                print(entity.textString)
                # m = Leader(entity)
                # m.appendToDF(self.LeadersDF)

    def isCollinear(self, x1, y1, x2, y2, x3, y3):
        return (x1*(y3-y2)+x3*(y2-y1)+x2*(y1-y3) == 0)
        
    
    def findFittingSize(self):
        fittingIndex, lineIndex = 0, 0
        for fittingIndex in self.FittingsDF.index:
            x2, y2 = self.FittingsDF['Block X'][fittingIndex], self.FittingsDF['Block Y'][fittingIndex]
            for lineIndex in self.LinesDF.index:
                x1, y1  = self.LinesDF['Start X'][lineIndex], self.LinesDF['Start Y'][lineIndex]
                x3, y3 = self.LinesDF['End X'][lineIndex], self.LinesDF['End Y'][lineIndex]
                if self.isCollinear(x1, y1, x2, y2, x3, y3):
                    self.FittingsDF.loc[fittingIndex, 'Matching Line ID'] = self.LinesDF['ID'][lineIndex]
                    print("found", self.LinesDF['ID'][lineIndex])


    def saveDF(self):
        self.LinesDF.to_csv('LinesCSV')
        print("logged to Lines")
        self.FittingsDF.to_csv('FittingsCSV')
        print("logged to Fittings")
        self.TextsDF.to_csv('TextsCSV')
        print("logged to Texts")


def findText(sheet):
    entities = sheet.Block     
    diSet = set() 
    for i in range(entities.Count):
        entity = entities.Item(i)
        if entity.ObjectName == 'AcDbMText' and "IRON" in entity.TextString:
            if sheet.Name == "W1":
                if entity.TextString[10:] in diSet:
                    print(entities.Item(i+1).textString, entity.TextString[10:])
                else:
                    diSet.add(entity.TextString[10:])
            else:
                print(entities.Item(i+1).textString, entity.TextString[10:])


def findPaperSheets():
    skipModelSpace = True
    for sheet in acad.activeDocument.Layouts:
        if skipModelSpace:
            skipModelSpace = False
            continue
        print(sheet.Name)
        findText(sheet)
        

sheet = Sheet()
# findPaperSheets()
sheet.findBlocks()
# sheet.findFittingSize()
# sheet.saveDF()

# sheet.findFittings()
# sheet.findLines()


        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

