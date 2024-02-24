from os import path, getcwd
from pathlib import Path
from locale import getlocale
from datetime import datetime
from win11toast import toast

# LOCAL IMPORTS
from utils import dataLoader, stats

cutOffYear = 2024

if __name__ == '__main__':

    # Init folders
    projectFolder = getcwd()
    dataFolder, outputFolder, graphFolder = dataLoader.initFolders(projectFolder)

    # Import transaction file
    importPath = path.join(Path.home(), "Download" if getlocale()[0].split('_')[0] == 'it' else "Downloads") 
    importedFileName = dataLoader.importTransactionFile(projectFolder, importPath)

    # Read movements
    df = dataLoader.loadTransactions(projectFolder)
    
    # Select the first and last date
    last_day, first_day =  df['VALUTA'].iloc[0], df['VALUTA'].iloc[-1]
    days_ago, hours_ago, *_ = (datetime.now() - last_day).components
    last_transaction = df[['DESC', 'CATEGORIA', 'IMPORTO']].iloc[0].astype('str').to_dict()
    print("\n--> LAST: ", last_day.strftime('%d-%m-%Y'),f'({days_ago} days and {hours_ago} hours ago)',
          "\n--> FIRST:", first_day.strftime('%d-%m-%Y'), "\n")

    if importedFileName:
        transactions_date = importedFileName.split('_')[:3]
        transactions_date[0] = transactions_date[0][-2:]
        
        tost_message = [f"Imported transaction up to {'-'.join(transactions_date)}"]
    else:
    
        tost_message = ['The expensives have been analyzsed']
    tost_message.append(f"Last Transaction was {days_ago} days and {hours_ago} hours ago "\
                            f"({last_transaction['DESC']} - {last_transaction['CATEGORIA']}, {last_transaction['IMPORTO']} €).")

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

    # Window Message
    toast(*tost_message, 
          icon = path.join(projectFolder, 'images', 'inbank.ico'),
          audio = {'silent': 'true'}, 
          duration='long',
          buttons = [
              {'activationType': 'protocol', 'arguments': path.join(projectFolder, 'outputs', 'expensives.xlsx'), 'content': 'Expensives'},
              {'activationType': 'protocol', 'arguments': path.join(projectFolder, 'outputs', 'graphs', 'expensivesByMonth.png'), 'content': 'Month Graph'},
              {'activationType': 'protocol', 'arguments': path.join(projectFolder, 'outputs', 'graphs', 'expensivesByQuarters.png'), 'content': 'Quarter Graph'}] 
    )