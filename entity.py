import pandas as pd

class Entity:
    def __init__(self, block, layout):
        self.ID = block.ObjectID
        self.sheet = layout
        self.locationX = block.insertionPoint[0]
        self.locationY = block.insertionPoint[1]