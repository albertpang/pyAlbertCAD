import pandas as pd

class Fitting:
    def __init__(self, block, DF):
        FittingID = block.ObjectID
        FittingsX = block.insertionPoint[0]
        FittingsY = block.insertionPoint[1]
        blockDynamicProperties = block.GetDynamicBlockProperties()
        for prop in blockDynamicProperties:
            if prop.PropertyName == 'Visibility1':
                dynamicBlockName = prop.Value

        self.DF = DF
        self.DF.loc[len(self.DF.index)] = [FittingID, dynamicBlockName, FittingsX, FittingsY]

    def saveDF(self):
        self.DF.to_csv('FittingsCSV')
    
    def isCollinear(self):
        pass
