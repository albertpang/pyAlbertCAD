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
                                            'Start Y', 'End X', 'End Y'])
        self.FittingsDF = pd.DataFrame(columns=['ID', 'Block Description', 
                                        'Block X', 'Block Y'])
        
    def findBlocks(self):
        for i, entity in enumerate(acad.ActiveDocument.Modelspace):
            print(i, entity.ObjectName)
            if entity.ObjectName == 'AcDbLine' and entity.Layer == 'C-PR-WATER':
                l = Line(entity)
                l.appendToDF(self.LinesDF)
                                                            
            if entity.ObjectName == 'AcDbBlockReference' and entity.EffectiveName.startswith("WATER"):
                f = Fitting(entity)
                f.appendToDF(self.FittingsDF)

        l.saveDF(self.LinesDF)
        f.saveDF(self.FittingsDF)
    
sheet = Sheet()
sheet.findBlocks()

# sheet.findFittings()
# sheet.findLines()


        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

