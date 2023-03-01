from entity import Entity
import pandas as pd
import math

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
        slope = (x1 - y1) / (x2 - y2)
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