from entity import Entity
import pandas as pd

class Fitting(Entity):
    def __init__(self, block):
        super().__init__(block)
        blockDynamicProperties = block.GetDynamicBlockProperties()
        for prop in blockDynamicProperties:
            if prop.PropertyName == 'Visibility1':
                self.dynamicBlockName = prop.Value

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.dynamicBlockName, 
                                self.locationX, self.locationY, "N/A"]