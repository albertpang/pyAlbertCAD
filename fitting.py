from entity import Entity
import pandas as pd

class Fitting(Entity):
    def __init__(self, block, layout):
        super().__init__(block, layout)
        if block.IsDynamicBlock:
            blockDynamicProperties = block.GetDynamicBlockProperties()

            for prop in blockDynamicProperties:
                # InHouse Palette use DBlock -- Vis1 To store Name
                if prop.PropertyName == 'Visibility1':
                    self.BlockName = prop.Value
                    return
            self.BlockName = block.EffectiveName
        else:
            self.BlockName = block.Name

        

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.BlockName, 
                                self.locationX, self.locationY, "N/A", "N/A"]