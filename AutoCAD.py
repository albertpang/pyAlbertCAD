import win32com.client
import numpy as np
import pandas as pd
from line import Line
acad = win32com.client.Dispatch("AutoCAD.Application")

# iterate through all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes.
class Sheet:
    def __init__(self):
        self.FittingsDF = pd.DataFrame(columns=['ID', 'Block Description', 'Block X', 'Block Y'])

    def findLines(self):
        for entity in acad.ActiveDocument.Modelspace:
            if entity.ObjectName == 'AcDbLine' and entity.Layer == 'C-PR-WATER':
                l = Line(entity)
                print (l.getStart(), l.getEnd())


    def findFittings(self):
        for entity in acad.ActiveDocument.Modelspace:
            name = entity.EntityName
            if name == 'AcDbBlockReference' and entity.EffectiveName.startswith("WATER"):
                self.getFittingsProps(entity)
        print(self.FittingsDF)
        

    def getFittingsProps(self, block):
        FittingID = block.ObjectID
        FittingsX = block.insertionPoint[0]
        FittingsY = block.insertionPoint[1]
        blockDynamicProperties = block.GetDynamicBlockProperties()
        for prop in blockDynamicProperties:
            if prop.PropertyName == 'Visibility1':
                dynamicBlockName = prop.Value
        self.FittingsDF.loc[len(self.FittingsDF.index)] = [FittingID, dynamicBlockName, FittingsX, FittingsY]
    
    def saveDF(self):
        self.FittingsDF.to_csv('FittingsCSV')

sheet = Sheet()
sheet.findFittings()
sheet.saveDF()
# sheet.findLines()


        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

