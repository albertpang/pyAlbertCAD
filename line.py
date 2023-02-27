from entity import Entity
import pandas as pd

class Line(Entity):
    def __init__(self, block, layout):
        self.ID = block.ObjectID
        self.sheet = layout
        self.layer = block.Layer
        self.startX = round(block.StartPoint[0], 2)
        self.startY = round(block.StartPoint[1], 2)
        self.endX = round(block.EndPoint[0], 2)
        self.endY = round(block.EndPoint[1], 2)
        self.slope = self.calculateSlope()

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.layer, self.startX,
                                self.startY, self.endX, self.endY, 
                                self.slope]

    def calculateSlope(self):
        slope = (self.endY - self.startY) / (self.endX - self.endY)
        return slope   