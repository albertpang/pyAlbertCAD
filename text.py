from entity import Entity
import pandas as pd

class Text(Entity):
    def __init__(self, block, layout):
        super().__init__(block, layout)
        self.text = block.textString

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.sheet, self.text, 
                                 self.locationX, self.locationY, "N/A", "N/A"]