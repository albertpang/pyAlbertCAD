import pandas as pd

class Fitting:
    def __init__(self, block):
        self.FittingID = block.ObjectID
        self.FittingsX = round(block.insertionPoint[0], 2)
        self.FittingsY = round(block.insertionPoint[1], 2)
        blockDynamicProperties = block.GetDynamicBlockProperties()
        for prop in blockDynamicProperties:
            if prop.PropertyName == 'Visibility1':
                self.dynamicBlockName = prop.Value

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.FittingID, self.dynamicBlockName, 
                                self.FittingsX, self.FittingsY, "N/A"]

    def saveDF(self, DF):
        DF.to_csv('FittingsCSV')
    
    def isCollinear(self):
        pass
