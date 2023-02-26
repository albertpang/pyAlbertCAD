import pandas as pd

class Line:
    def __init__(self, block):
        self.ID = block.ObjectID
        self.layer = block.Layer
        self.startX = round(block.StartPoint[0], 2)
        self.startY = round(block.StartPoint[1], 2)
        self.endX = round(block.EndPoint[0], 2)
        self.endY = round(block.EndPoint[1], 2)
        self.slope = self.calculateSlope()

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.layer, self.startX,
                                self.startY, self.endX, self.endY, self.slope]

    def calculateSlope(self):
        slope = (self.endY - self.startY) / (self.endX - self.endY)
        return slope   

    def saveDF(self, DF):
        DF.to_csv('LinesCSV')


    def isCollinear(self):
        pass
