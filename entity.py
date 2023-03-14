import pandas as pd
import math
import win32com.client
from entity import Entity
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

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.BlockName, 
                                self.locationX, self.locationY, "N/A", "N/A"]
        
class Text(Entity):
    def __init__(self, block, layout):
        super().__init__(block, layout)
        self.text = block.textString

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.text, 
                                 self.locationX, self.locationY, "N/A", "N/A"]

class Viewport(Entity):
    def __init__(self, entity, layout):
        
        objectID = entity.ObjectID
        sheet = layout.Name
        self.crossLine = None
        psLeftTopCorner = (entity.Center[0] - (abs(entity.Width) / 2), 
                            entity.Center[1] + (abs(entity.Height) / 2))
        psRightBotCorner = (entity.Center[0] + (abs(entity.Width) / 2), 
                            entity.Center[1] - (abs(entity.Height) / 2))
        
        # Create AutoCAD Point object
        psleftTopCornerPoint = APoint(psLeftTopCorner[0], psLeftTopCorner[1])
        psrightBotCornerPoint = APoint(psRightBotCorner[0], psRightBotCorner[1])
        self.convertLinePaperSpace()
        self.appendToDF()
    
    
    def convertLinePaperSpace(self, p1, p2):
        self.crossLine = doc.PaperSpace.AddLine(p1, p2)
        doc.SendCommand("select last  chspace  ")
        doc.SendCommand("pspace ")


    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [Viewport.objectID, Viewport.sheet,
                                 Viewport.Width, Viewport.Height,
                                Viewport.psLeftTopCorner, Viewport.psRightBotCorner,
                                (self.crossLine.StartPoint[0], self.crossLine.StartPoint[1]), 
                                (self.crossLine.EndPoint[0], self.crossLine.EndPoint[1])]

    # TranslateCoordinates Method -- Does Not Work
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

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.layer, self.startX,
                                self.startY, self.endX, self.endY, self.length,
                                self.slope]

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
        self.lineDF = DF
        self.breakPolyline()
    

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
        self.lineDF.loc[len(self.lineDF.index)] = [f"Polyline - {self.block.ObjectId}", 
                                                    self.sheet, self.layer, 
                                                    self.startX, self.startY, 
                                                    self.endX, self.endY, self.length,
                                                    self.slope]   
        