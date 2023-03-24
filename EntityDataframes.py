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

MLeadersDF = pd.DataFrame(columns=['ID', 'Sheet', 'Text', 'Starting Vertex Coordinate X', 'Starting Vertex Coordinate Y',
                                   'Ending Vertex Coordinate X', 'Ending Vertex Coordinate Y', 'Matching Line ID',
                                   'Matching Line Layer', 'Matching Line Length'
                                   ])

BillOfMaterialsDF = pd.DataFrame(columns=['Sheet', 'Associated Text String'])

def save_dataframe():
    """Save the DataFrame objects to CSV files.

    The function saves each of the four DataFrame objects to a separate CSV file.
    The file names are hardcoded as 'LinesCSV', 'FittingsCSV', 'TextsCSV', and
    'BillOfMaterialsCSV', respectively. The function logs a message to the console
    after each file is saved to indicate which DataFrame object was saved.

    Returns:
    None
    """
    print("Saving CSVs")
    LinesDF.to_csv('LinesCSV')
    print("-Logged to Lines")
    FittingsDF.to_csv('FittingsCSV')
    FittingsDF.to_excel('FittingExcel.xlsx')
    print("--Logged to Fittings")
    TextsDF.to_csv('TextsCSV')
    print("---Logged to Texts")
    BillOfMaterialsDF.to_csv('BillOfMaterialsCSV')
    print("----Logged to BOM")
    ViewportsDF.to_csv('ViewportsCSV')
    print("-----Logged to ViewportsCSV")
    MLeadersDF.to_csv('MLeadersCSV')
    print("------Logged to MLeadersCSV")
