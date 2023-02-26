import win32com.client
import numpy as np
import pandas as pd
from line import Line

from fitting import Fitting

acad = win32com.client.Dispatch("AutoCAD.Application")

# iterate through all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes.
class Sheet:
    def __init__(self):
        self.LinesDF = pd.DataFrame(columns=['ID', 'Block Description', 'Start X', 
                                            'Start Y', 'End X', 'End Y', 'Slope'])
        self.FittingsDF = pd.DataFrame(columns=['ID', 'Block Description', 
                                        'Block X', 'Block Y', 'Matching Line ID'])
        
    def findBlocks(self):
        for i, entity in enumerate(acad.ActiveDocument.Modelspace):
            if entity.ObjectName == 'AcDbLine' and entity.Layer == 'C-PR-WATER':
                l = Line(entity)
                l.appendToDF(self.LinesDF)
                                                            
            if entity.ObjectName == 'AcDbBlockReference' and entity.EffectiveName.startswith("WATER"):
                f = Fitting(entity)
                f.appendToDF(self.FittingsDF)
            
            if entity.ObjectName == 'AcDbMText':
                print(entity.TextString)




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
        print("logged to fittings")

sheet = Sheet()
sheet.findBlocks()
sheet.findFittingSize()
sheet.saveDF()

# sheet.findFittings()
# sheet.findLines()


        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

