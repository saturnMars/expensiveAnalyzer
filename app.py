from os import path, makedirs
from utils import dataLoader, stats
from pathlib import Path
from locale import getlocale

cutOffYear = 2023

if __name__ == '__main__':

    projectFolder = path.dirname(__file__)
    language = getlocale()[0].split('_')[0]

    # Import transaction file
    importPath = path.join(Path.home(), "Download" if language == 'it' else "Downloads")
    dataLoader.importTransactionFile(projectFolder, importPath)

    # Read movements
    df = dataLoader.loadTransactions(projectFolder)

    # Create the output folder
    outputFolder = path.join(projectFolder, 'outputs')
    graphFolder = path.join(outputFolder, 'graphs')
    if not path.exists(graphFolder):
        makedirs(graphFolder)

    # Compute income stats
    stats.computeIncomes(df.copy(), outputFolder = outputFolder)

    # Compute monthly stats
    stats.monthlyStats(df.copy(), outputFolder = outputFolder)
    
    # Compute expensive by ABI code
    stats.groupExpensives(df.copy(), outputFolder = outputFolder, feature = "CAUSALE ABI", include_incomes = False, cutoff_year = cutOffYear)
    stats.groupExpensives(df.copy(), outputFolder = outputFolder, feature = "CATEGORIA", include_incomes = False, cutoff_year = cutOffYear)

    # Create the area graphs
    stats.visualizeExpensives(df, outputFolder = graphFolder, cutoff_year = cutOffYear, groupby = "MESE" )
    stats.visualizeExpensives(df, outputFolder = graphFolder, cutoff_year = cutOffYear, groupby = "TRIMESTRE")

    stats.visualizeExpensives(df, outputFolder = graphFolder, cutoff_year = cutOffYear, feature = "CAUSALE ABI", groupby = "TRIMESTRE" )