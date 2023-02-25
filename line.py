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
