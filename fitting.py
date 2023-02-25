class Fitting:
    def __init__(self, block):
        self.startX = block.StartPoint[0]
        self.startY = block.StartPoint[1]
        self.endX = block.EndPoint[0]
        self.endY = block.EndPoint[1]

    
    # def getStart(self):
    #     return (self._startX, self._startY)


    # def getEnd(self):
    #     return (self._endX, self._endY)
    

    def isCollinear(self):
        pass
