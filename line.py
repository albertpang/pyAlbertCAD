import pandas as pd

class Line:
    def __init__(self, block, DF):
        self.ID = block.ObjectID
        self.layer = block.Layer
        self.startX = block.StartPoint[0]
        self.startY = block.StartPoint[1]
        self.endX = block.EndPoint[0]
        self.endY = block.EndPoint[1]
        self.DF = DF
        self.DF.loc[len(self.DF.index)] = [self.ID, self.layer, self.startX,
                                            self.startY, self.endX, self.endY]

    def appendDF(self):
        self.LinesDF.loc[len(self.LinesDF.index)] = [self.ID, self.layer, self.startX,
                                                    self.startY, self.endX, self.endY]

    def saveDF(self):
        
        self.DF.to_csv('LinesCSV')


    # def getStart(self):
    #     return (self._startX, self._startY)


    # def getEnd(self):
    #     return (self._endX, self._endY)
    

    def isCollinear(self):
        pass
