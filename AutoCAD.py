import win32com.client
import numpy as np
import pandas as pdt
acad = win32com.client.Dispatch("AutoCAD.Application")

# iterate through all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes.
class Sheet:
    def __init__(self):
        self.blockDF = pd.DataFrame(columns=['Block Description', 'Block X', 'Block Y'])

    def findLines(self):
        for entity in acad.ActiveDocument.Modelspace:
            if entity.ObjectName == 'AcDbLine' and entity.Layer == 'C-PR-WATER':
                l = Line(entity)
                print (l.getStart(), l.getEnd())


    def findBlocks(self):
        for entity in acad.ActiveDocument.Modelspace:
            name = entity.EntityName
            if name == 'AcDbBlockReference' and entity.EffectiveName.startswith("WATER"):
                self.getBlockProps(entity)
        print(self.blockDF)
        

    def getBlockProps(self, block):
        blockX = block.insertionPoint[0]
        blockY = block.insertionPoint[1]
        blockDynamicProperties = block.GetDynamicBlockProperties()
        for prop in blockDynamicProperties:
            if prop.PropertyName == 'Visibility1':
                dynamicBlockName = prop.Value
        self.blockDF.loc[len(self.blockDF.index)] = [ dynamicBlockName, blockX, blockY]
    
    def saveDF(self):
        self.blockDF.to_csv('blockCSV')

    
    

class Line:
    def __init__(self, block):
        self._startX = block.StartPoint[0]
        self._startY = block.StartPoint[1]
        self._endX = block.EndPoint[0]
        self._endY = block.EndPoint[1]
    
    def getStart(self):
        return (self._startX, self._startY)


    def getEnd(self):
        return (self._endX, self._endY)
    

    def isCollinear(self):
        pass

sheet = Sheet()
sheet.findBlocks()
sheet.saveDF()
# sheet.findLines()


        # Depending on Tool Palette used, there may be a unique method for creating blocks
        # HasAttributes = entity.HasAttributes
        # if HasAttributes:
        #     for attrib in entity.GetAttributes():
        #         print(attrib.TextString, attrib.TagString)

