import pandas as pd

class Entity:
    def __init__(self, block):
        self.ID = block.ObjectID
        self.locationX = round(block.insertionPoint[0], 2)
        self.locationY = round(block.insertionPoint[1], 2)