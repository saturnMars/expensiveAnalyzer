from os import path, getcwd
from pathlib import Path
from locale import getlocale
from datetime import datetime
from win11toast import toast
from configparser import ConfigParser
from numpy import datetime64, timedelta64
from threading import Thread
from multiprocessing import Process

# LOCAL IMPORTS
from utils import dataLoader, stats

if __name__ == '__main__':

    config = ConfigParser()
    config.read(path.join('src','config.ini'))
    reporting_period = int(config.get('PERIOD', 'reporting_months'))

    # Init folders
    projectFolder = getcwd()
    dataFolder, outputFolder, graphFolder = dataLoader.initFolders(projectFolder)

    # Import transaction file
    importPath = path.join(Path.home(), "Download" if getlocale()[0].split('_')[0] == 'it' else "Downloads") 
    importedFileName = dataLoader.importCSVTransactions(projectFolder, importPath)

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

    # Compute monthly stats
    Thread(target = stats.monthly_stats, args=(df, outputFolder)).start()

    # Consider only the selected period
    if reporting_period > 0:
        cutOff = datetime64('today', 'M')  - timedelta64(reporting_period, 'M')
        df = df[df['VALUTA'].dt.to_period('M') > str(cutOff)]

        if df.empty:
            print(f"NO TRANSACTION IN THE LAST {reporting_period} MONTHS")
            exit()
        print(f"REPORTING PERIOD: {reporting_period} months\nCUTOFF: {cutOff} ({df['VALUTA'].iloc[-1].date()} <--> {df['VALUTA'].iloc[0].date()})\n")

    # Compute income stats
    Thread(target= stats.compute_incomes, args=(df, outputFolder)).start()
    
    # Compute expensive by ABI code
    threads = [Thread(target = stats.group_expensive, args=(df, outputFolder, feature, False)) for feature in ["CAUSALE ABI", "CATEGORIA"]]
    [thread.start() for thread in threads]

    # Create the area graphs
    processes = [Process(target = stats.expensive_graph, args=(df, graphFolder, period, "CATEGORIA"), daemon=True) 
                 for period in ["MESE", "TRIMESTRE"]]
    processes.append(Process(target =  stats.expensive_graph, args=(df, graphFolder, 'TRIMESTRE', "CAUSALE ABI"), daemon=True))
    [process.start() for process in processes]

    # Window Message
    items = {'Expensives':  path.join(projectFolder, 'outputs', 'expensives.xlsx'),
             'Month Graph':  path.join(projectFolder, 'outputs', 'graphs', 'expensivesByMonth.png'),
             'Quarter Graph':  path.join(projectFolder, 'outputs', 'graphs', 'expensivesByQuarters.png')}
    toast(*tost_message, icon = path.join(projectFolder, 'images', 'inbank.ico'), audio = {'silent': 'true'}, duration='long',
          buttons = [{'activationType': 'protocol', 'arguments':path, 'content': title} for title, path in items.items()])