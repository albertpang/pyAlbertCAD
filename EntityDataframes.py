import pandas as pd

LinesDF = pd.DataFrame(columns=['ID', 'Sheet', 'Layer', 'Start X', 
                                    'Start Y', 'End X', 'End Y', 'Length', 'Slope'])
FittingsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Block Description', 
                                'Block X', 'Block Y', 'Matching Line ID', 
                                'Matching Line Length'])
TextsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Text', 'Block X', 'Block Y', 
                                'Associated Text ID', 'Associated Text String'])
        
ViewportsDF = pd.DataFrame(columns=['ID', 'Sheet', 'Width', 'Height', 'Type', 'Overlaps Center', 
                                    'Num of Frozen Layers', 'Is BasePlan ModelSpace',
                                    'PaperSpace Coordinate Corner1 X', 'PaperSpace Coordinate Corner1 Y',
                                    'PaperSpace Coordinate Corner2 X', 'PaperSpace Coordinate Corner2 Y',
                                    'ModelSpace Coordinate Corner1 X', 'ModelSpace Coordinate Corner1 Y',
                                    'ModelSpace Coordinate Corner2 X', 'ModelSpace Coordinate Corner2 Y',
                                    'PaperSpace Coordinate Corner3 X', 'PaperSpace Coordinate Corner3 Y',
                                    'PaperSpace Coordinate Corner4 X', 'PaperSpace Coordinate Corner4 Y',
                                    'ModelSpace Coordinate Corner3 X', 'ModelSpace Coordinate Corner3 Y',
                                    'ModelSpace Coordinate Corner4 X', 'ModelSpace Coordinate Corner4 Y',
                                    ])

MLeadersDF = pd.DataFrame(columns=['ID', 'Sheet', ''])

BillOfMaterialsDF = pd.DataFrame(columns=['Sheet', 'Associated Text String'])