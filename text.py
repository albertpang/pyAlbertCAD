from entity import Entity
import pandas as pd

class Text(Entity):
    def __init__(self, block):
        super().__init__(block)
        self.text = block.textString

    def appendToDF(self, DF):
        DF.loc[len(DF.index)] = [self.ID, self.text, self.locationX, self.locationY, "N/A"]