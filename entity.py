import time
import pandas as pd
import math
import win32com.client
from ACAD_DataTypes import APoint
acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

class Entity:
    def __init__(self, block, layout):
        self.ID = block.ObjectID
        self.sheet = layout
        self.locationX = block.insertionPoint[0]
        self.locationY = block.insertionPoint[1]

class Fitting(Entity):
    def __init__(self, block, layout):
        super().__init__(block, layout)
        if block.IsDynamicBlock:
            blockDynamicProperties = block.GetDynamicBlockProperties()

            for prop in blockDynamicProperties:
                # InHouse Palette use DBlock -- Vis1 To store Name
                if prop.PropertyName == 'Visibility1':
                    self.BlockName = prop.Value
                    return
            self.BlockName = block.EffectiveName
        else:
            self.BlockName = block.Name


        
class Text(Entity):
    def __init__(self, block, layout):
        super().__init__(block, layout)
        self.text = block.textString

class Viewport(Entity):
    def __init__(self, block, layout):
        time.sleep(0.5)
        self.ID = block.ObjectID
        self.center = block.Center
        self.height = block.Height
        self.width = block.Width
        self.XData = block.GetXData("")
        self.sheet = layout.Name
        self.numFrozenLayers = self.count_frozen_layers()
        self.isMain = None
        self.sheetHeight, self.sheetWidth = layout.GetPaperSize()
        self.scale = round(block.CustomScale, 3)
        self.type = self.classify_viewport()
        self.crossLine = None

        self.psCorner1 = (self.center[0] - (abs(self.width) / 2), 
                            self.center[1] + (abs(self.height) / 2))
        self.psCorner2 = (self.center[0] + (abs(self.width) / 2), 
                            self.center[1] - (abs(self.height) / 2))
        
        # Create AutoCAD Point object
        psCorner1Point = APoint(self.psCorner1[0], self.psCorner1[1])
        psCorner2Point = APoint(self.psCorner2[0], self.psCorner2[1])
        self.convertLinePaperSpace(psCorner1Point, psCorner2Point)
        self.classify_viewport()

    def count_frozen_layers(self):
        """ Uses block.getXData() to count all "Freeze VP" Layers """
        self.numFrozenLayers = 0 
        for data in self.XData[0]:
            self.numFrozenLayers += 1
        return self.numFrozenLayers
            

    def classify_viewport(self):
        if self.scale == .25:
            self.type = "Section View"
            self.isMain = 'Not Applicable'

        elif self.scale == .05:
            self.type = "Model View"
            # if
            self.isMain = 'TRUE'
            # else:
            self.isMain = 'FALSE'
            
        else:
            return (f"Incorrect Scale: {self.scale}")
    
    def convertLinePaperSpace(self, p1, p2):
        self.crossLine = doc.PaperSpace.AddLine(p1, p2)
        doc.SendCommand("select last  chspace  ")
        self.msCorner1 = (self.crossLine.StartPoint[0], self.crossLine.StartPoint[1])
        self.msCorner2 = (self.crossLine.EndPoint[0], self.crossLine.EndPoint[1])
        doc.SendCommand("pspace ") 

class Line(Entity):
    def __init__(self, block, layout, isPolyline):
        self.ID = block.ObjectID
        self.sheet = layout
        self.layer = block.Layer
        self.length = block.length

        if isPolyline:
            self.startX = block.Coordinates[0]
            self.startY = block.Coordinates[1]
            self.endX = block.Coordinates[2]
            self.endY = block.Coordinates[3]
        else:
            self.startX = block.StartPoint[0]
            self.startY = block.StartPoint[1]
            self.endX = block.EndPoint[0]
            self.endY = block.EndPoint[1]

        self.slope = self.calculateSlope(self.startX, self.startY, 
                                         self.endX, self.endY)


    def calculateLength(self, x1, y1, x2, y2):
        length = math.dist([x1, y1], [x2, y2])
        return round(length, 2)


    def calculateSlope(self, x1, y1, x2, y2):
        if (x2 - x1) == 0:
            return "undef"
        if (y2 - y1) == 0:
            return "inf"
        slope = (y2 - y1) / (x2 - x1)
        return round(slope, 3)

class PolyLine(Line):

    def __init__(self, block, layout, DF):
        self.block = block
        self.sheet = layout
        self.layer = block.Layer
        self.DF = DF
        
    

    def breakPolyline(self):
        i = 0
        while i < len(self.block.Coordinates) - 2:
            self.startX = self.block.Coordinates[i]
            self.startY = self.block.Coordinates[i + 1]
            self.endX = self.block.Coordinates[i + 2]
            self.endY = self.block.Coordinates[i + 3]
            self.length = self.calculateLength(self.startX, self.startY, 
                                                self.endX, self.endY)
            self.slope = self.calculateSlope(self.startX, self.startY, 
                                                self.endX, self.endY)
            self.appendToDF()
            i += 2

    def appendToDF(self):
        self.DF.loc[len(self.DF.index)] = [f"Polyline - {self.ID}", 
                                            self.sheet, self.layer, 
                                            self.startX, self.startY, 
                                            self.endX, self.endY, self.length,
                                            self.slope]

                