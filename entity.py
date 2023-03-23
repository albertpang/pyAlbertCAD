import pandas as pd
import math
import wait
import win32com.client
from ACAD_DataTypes import APoint
import time

acad = win32com.client.Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument

'''WRAP ALL ATTRIBUTES AND METHODS USING WAIT CLASS FUNCTIONS OTHERWISE
YOU WILL GET A COM_ERROR.'''

class Entity:
    def __init__(self, block, layout):
        self.ID = wait.wait_for_attribute(block, "ObjectID")
        self.sheet = layout
        self.locationX = wait.wait_for_attribute(block, "insertionPoint")[0]
        self.locationY = wait.wait_for_attribute(block, "insertionPoint")[1]


class Fitting(Entity):
    def __init__(self, block, layout):
        '''Initailize a Fitting Object based on Block Type'''
        super().__init__(block, layout)
        # Check if the blocks are Dynamic. This changes how properties are held
        if block.IsDynamicBlock:
            blockDynamicProperties = wait.wait_for_method_return(block, "GetDynamicBlockProperties")
            # Isolate the block Fitting Type
            for prop in blockDynamicProperties:
                # InHouse Palette use DBlock -- Vis1 To store Name
                if wait.wait_for_attribute(prop, "PropertyName") == 'Visibility1':
                    self.BlockName = wait.wait_for_attribute(prop, "Value")
                    return
            # Sometimes the Name of a Dynamic Block is held in EffectiveName tag
            self.BlockName = wait.wait_for_attribute(block, "EffectiveName")
        # If not a Dynamic Block, we can just check Name field for the Fitting
        else:
            self.BlockName = wait.wait_for_attribute(block, "Name")


class Text(Entity):
    def __init__(self, block, layout):
        super().__init__(block, layout)
        self.text = wait.wait_for_attribute(block, "textString")


class Viewport(Entity):
    def __init__(self, block, layout):
        '''Initailize a Viewport Object based on Block Type'''
        while True:
            try:
                wait.set_attribute(block, "ViewportOn", False)
                wait.set_attribute(block, "ViewportOn", True)
                break
            except:
                print("Viewport ERROR")
                time.sleep(0.3)
                continue

        self.ID = wait.wait_for_attribute(block, "ObjectID")
        self.center = wait.wait_for_attribute(block, "Center")
        self.height = wait.wait_for_attribute(block, "Height")
        self.width = wait.wait_for_attribute(block, "Width")
        self.XData = wait.wait_for_method_return(block, "GetXData" ,"")
        self.sheet = wait.wait_for_attribute(layout, "Name")
        self.numFrozenLayers = self.count_frozen_layers()
        self.scale = round(wait.wait_for_attribute(block, "CustomScale"), 3)
        self.msCenter = self.XData[1][4]

        # PaperSpace Coordinates for Viewport
        self.psCorner1 = (self.center[0] - (abs(self.width) / 2), 
                          self.center[1] + (abs(self.height) / 2))
        self.psCorner2 = (self.center[0] + (abs(self.width) / 2), 
                            self.center[1] - (abs(self.height) / 2))
        self.psCorner3 = (self.center[0] + (abs(self.width) / 2), 
                          self.center[1] + (abs(self.height) / 2))
        self.psCorner4 = (self.center[0] - (abs(self.width) / 2), 
                            self.center[1] - (abs(self.height) / 2))
                             
        self.isCenter = self.is_center_viewport()
        self.type = self.classify_viewport()       

        # Create AutoCAD Point object
        psCorner1Point = APoint(self.psCorner1[0], self.psCorner1[1])
        psCorner2Point = APoint(self.psCorner2[0], self.psCorner2[1])
        psCorner3Point = APoint(self.psCorner3[0], self.psCorner3[1])
        psCorner4Point = APoint(self.psCorner4[0], self.psCorner4[1])
        
        self.convertLinePaperSpace(psCorner1Point, psCorner2Point, 
                                   psCorner3Point, psCorner4Point)

    def count_frozen_layers(self):
        """ Uses block.GetXData() to count all "Freeze VP" Layers. Used to find
         most likely main ModelSpace viewport."""
        self.numFrozenLayers = 0 
        for data in self.XData[0]:
            if data == 1003:
                self.numFrozenLayers += 1
        return self.numFrozenLayers
    
    def is_center_viewport(self):
        '''Checks if Viewport contains the midpoint of a papersheet. Creates 
        priority for finding the ModelSpace viewport of a sheet.'''
        sheetHeight, sheetWidth  = wait.wait_for_method_return(wait.wait_for_attribute(doc, "ActiveLayout"), "GetPaperSize")

        def liesWithin(cp, c1, c2) -> bool:
            '''Assuming a viewport is relatively rectangular, does it contain
            the midpoint of the papersheet'''
            x, y = cp
            minX = min(c1[0], c2[0])
            maxX = max(c1[0], c2[0])
            minY = min(c1[1], c2[1])
            maxY = max(c1[1], c2[1])
            insideX = (minX <= x <= maxX)
            insideY = (minY <= y <= maxY)
            return (insideX and insideY)
        
        # Constant unit conversion
        PIXELtoINCH = 25.4
        sheetHeight /= PIXELtoINCH
        sheetWidth /= PIXELtoINCH
        layoutCenter = (sheetHeight / 2), (sheetWidth / 2)
        return liesWithin(layoutCenter, self.psCorner1, self.psCorner2)

    def classify_viewport(self):
        if self.scale == .25:
            return "Section View"
        elif self.scale == .05:
            return "Model View"            
        else:
            return (f"Incorrect Scale: {self.scale}")
    
    def convertLinePaperSpace(self, p1, p2, p3, p4):
        '''This method will take all 4 corners of the viewport, and draw two 
        intersecting lines that we will use to get coordinates of the viewport.
        We need two lines, because you cannot get the askewed viewports otherwise'''

        # Draw the Lines and declare as crossLines
        paperSpace = wait.wait_for_attribute(doc, "PaperSpace")

        self.crossLine1 = wait.wait_for_method_return(paperSpace, "AddLine", p1, p2)
        wait.wait_for_method_return(doc, "SendCommand", "chspace last   ")
        # Get the Coordinates for the diagonal inside the Viewport
        msCorner1 = wait.wait_for_attribute(self.crossLine1, "StartPoint")
        msCorner2 = wait.wait_for_attribute(self.crossLine1, "EndPoint")
        wait.wait_for_method_return(doc, "SendCommand", "pspace ")

        self.crossLine2 = wait.wait_for_method_return(paperSpace, "AddLine", p3, p4)
        wait.wait_for_method_return(doc, "SendCommand", "chspace last   ")
        # Get the Coordinates for the diagonal inside the Viewport
        msCorner3 = wait.wait_for_attribute(self.crossLine2, "StartPoint")
        msCorner4 = wait.wait_for_attribute(self.crossLine2, "EndPoint") 
        wait.wait_for_method_return(doc, "SendCommand", "pspace ")

        # Convert all 4 corner points into tuples of coordinates for the viewport
        self.msCorner1 = (msCorner1[0], msCorner1[1])
        self.msCorner2 = (msCorner2[0], msCorner2[1])
        self.msCorner3 = (msCorner3[0], msCorner3[1])
        self.msCorner4 = (msCorner4[0], msCorner4[1])
        

class Line(Entity):
    def __init__(self, block, layout, isPolyline):
        self.ID = block.ObjectID
        self.sheet = layout
        self.layer = block.Layer
        self.length = block.length

        # Get PolyLine Coordinates of a Single segment
        if isPolyline:
            self.startX = block.Coordinates[0]
            self.startY = block.Coordinates[1]
            self.endX = block.Coordinates[2]
            self.endY = block.Coordinates[3]
        # Get Line Coordinates
        else:
            self.startX = block.StartPoint[0]
            self.startY = block.StartPoint[1]
            self.endX = block.EndPoint[0]
            self.endY = block.EndPoint[1]

        self.slope = self.calculateSlope(self.startX, self.startY, 
                                         self.endX, self.endY)


    def calculateLength(self, x1, y1, x2, y2):
        '''Calculates Distance of a line'''
        length = math.dist([x1, y1], [x2, y2])
        return round(length, 2)


    def calculateSlope(self, x1, y1, x2, y2):
        '''Calculates Slope of a line'''
        if (x2 - x1) == 0:
            return "undef"
        if (y2 - y1) == 0:
            return "inf"
        slope = (y2 - y1) / (x2 - x1)
        return round(slope, 3)


class PolyLine(Line):
    def __init__(self, block, layout, DF):

        self.block = block
        self.sheet = layout
        self.layer = block.Layer
        self.DF = DF
        

    def breakPolyline(self):
        '''This breaks down polylines into individual lines for LinesDF'''
        i = 0
        while i < len(self.block.Coordinates) - 2:
            self.startX = self.block.Coordinates[i]
            self.startY = self.block.Coordinates[i + 1]
            self.endX = self.block.Coordinates[i + 2]
            self.endY = self.block.Coordinates[i + 3]
            self.length = self.calculateLength(self.startX, self.startY, 
                                                self.endX, self.endY)
            self.slope = self.calculateSlope(self.startX, self.startY, 
                                                self.endX, self.endY)
            self.appendToDF()
            i += 2

    def appendToDF(self):
        self.DF.loc[len(self.DF.index)] = [f"Polyline - {self.ID}", 
                                            self.sheet, self.layer, 
                                            self.startX, self.startY, 
                                            self.endX, self.endY, self.length,
                                            self.slope]
        

# class LeaderLine(Entity):
#     def __init__(self, block, layout):
#         self.ID = wait.wait_for_attribute(block, "ObjectID")
#         self.sheet = layout
#         self.leaderlineVertices = wait.wait_for_method_return(block, "GetLeaderLineVertices", 0)
#         self.leaderlineIndexes = wait.wait_for_method_return(block, "GetLeaderLineIndexes")
#         self.leaderlineIndex = wait.wait_for_method_return(block, "GetLeaderIndex")

#     def break_leader_vertices(self):
#         numberofVertices = len(self.leaderlineVertices)
#         vertexList = 
         
#         for index, vertex in enum(self.leaderlineVertices):
#             if index +

#         for 
#         calloutPointX, calloutPointY = self.leaderlineVertices


        