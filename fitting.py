import pandas as pd

class Fitting:
    def __init__(self, block):
        self.FittingID = block.ObjectID
        self.FittingsX = block.insertionPoint[0]
        self.FittingsY = block.insertionPoint[1]
        blockDynamicProperties = block.GetDynamicBlockProperties()
        for prop in blockDynamicProperties:
            if prop.PropertyName == 'Visibility1':
                self.dynamicBlockName = prop.Value

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.FittingID, self.dynamicBlockName, 
                                           self.FittingsX, self.FittingsY]

    def saveDF(self, DF):
        DF.to_csv('FittingsCSV')
    
    def isCollinear(self):
        pass
