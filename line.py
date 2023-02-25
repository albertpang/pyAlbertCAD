import pandas as pd

class Line:
    def __init__(self, block):
        self.ID = block.ObjectID
        self.layer = block.Layer
        self.startX = block.StartPoint[0]
        self.startY = block.StartPoint[1]
        self.endX = block.EndPoint[0]
        self.endY = block.EndPoint[1]

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.layer, self.startX,
                                self.startY, self.endX, self.endY]
        
    def saveDF(self, DF):
        DF.to_csv('LinesCSV')


    # def getStart(self):
    #     return (self._startX, self._startY)


    # def getEnd(self):
    #     return (self._endX, self._endY)
    

    def isCollinear(self):
        pass
