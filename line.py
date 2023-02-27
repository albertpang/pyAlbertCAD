from entity import Entity
import pandas as pd

class Line(Entity):
    def __init__(self, block, layout):
        self.ID = block.ObjectID
        self.sheet = layout
        self.layer = block.Layer
        self.length = block.length
        self.startX = block.StartPoint[0]
        self.startY = block.StartPoint[1]
        self.endX = block.EndPoint[0]
        self.endY = block.EndPoint[1]
        self.slope = self.calculateSlope()

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.layer, self.startX,
                                self.startY, self.endX, self.endY, self.length,
                                self.slope]

    def calculateLength(self):
        length = ((self.startX - self.endX) **2 + (self.startY - self.endY) **2) ** 0.5
        return round(length, 2)


    def calculateSlope(self):
        slope = (self.endY - self.startY) / (self.endX - self.endY)
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
            self.appendToDF()
            i += 2


    def appendToDF(self):
        self.lineDF.loc[len(self.lineDF.index)] = [f"Polyline - {self.block.ObjectId}", 
                                                    self.sheet, self.layer, 
                                                    self.startX, self.startY, 
                                                    self.endX, self.endY, self.calculateLength(),
                                                    self.calculateSlope()]    
        
