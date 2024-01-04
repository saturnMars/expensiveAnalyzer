import pandas as pd
from os import path, listdir, remove
from json import load
import shutil

def importTransactionFile(projectFolder, importPath):
    print(importPath)
    
    if not path.exists(importPath):
        print(f"\n[WARNING] The import path does not exist! {importPath}\n--> (A) Change che import path in the app.py file.\n--> (B) Put the file into the ./data folder.\n")
        return False

    for fileName in listdir(importPath):
        if  fileName.endswith('.csv') and 'listamovimenti' in fileName.lower():

            # Move the file
            shutil.move(src = path.join(importPath, fileName), dst = path.join(projectFolder, 'data'))

            print("--> Imported file:", fileName, "\n")

            # Rename the file
            #rename(src = path.join(folderPath, fileName), dst = path.join(folderPath, 'movements.csv'))

            return True
    return False

def loadTransactions(projectFolder, outputfileName = 'transactions.xlsx'):

    # Attach new transactions
    df = importTransactions(projectFolder)

    # Import the main dataframe
    dataFolder = path.join(projectFolder, 'data')
    folderFiles = listdir(dataFolder)
    if outputfileName in folderFiles:

        # Load the excel file
        main_df = pd.read_excel(path.join(dataFolder, outputfileName))

        # Merge the dataframes
        df = pd.concat([main_df, df]).drop_duplicates(subset=['VALUTA', 'IMPORTO']).reset_index(drop = True)

    if len(df) == 0:
        raise Exception('No data! Neither in the DATA folder nor in the DOWNLOAD folder.\n')

    # Sort the new dataframe
    df = df.sort_values(by = ['VALUTA'], ascending= False)

    # Create the month column 
    df['MESE'] = df['VALUTA'].dt.to_period('M') #.strftime('%B %Y')
    df['TRIMESTRE'] = df['VALUTA'].dt.to_period('Q')
    df['ANNO'] = df['VALUTA'].dt.to_period('Y').dt.year

    # Select the first and last date
    last_day, first_day =  df['VALUTA'].iloc[0], df['VALUTA'].iloc[-1]
    print("--> LAST: ", last_day.strftime('%d-%m-%Y'), "\n--> FIRST:", first_day.strftime('%d-%m-%Y'), "\n")

    # Erase the folder 
    for fileName in folderFiles:
       file_path = path.join(dataFolder, fileName)
       if path.isfile(file_path):
           remove(file_path)

    # Save the dataframe
    df.to_excel(path.join(dataFolder, f"transactions.xlsx"), index = False, sheet_name = 'Transactions', freeze_panes = (1,1))

    return df


def importTransactions(projectFolder):

    dataFolder = path.join(projectFolder, 'data')
     
    # Scan the folder
    folderFiles = listdir(dataFolder)

    dfs = []
    for fileName in folderFiles:
        if fileName.endswith('.csv'):
          dfs.append(pd.read_csv(path.join(dataFolder, fileName), sep = ';', skipfooter = 3, parse_dates=['DATA', 'VALUTA'],  engine='python', dayfirst=True))
    
    if len(dfs) == 0: 
        return pd.DataFrame()
    df = pd.concat(dfs).drop_duplicates().reset_index(drop = True)

    # Drop the invalid entries
    df = df.dropna(subset = 'CAUSALE ABI')

    # Parse the data
    df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True)

    # Modify the numerical representation
    for col in ['AVERE', 'DARE']:
        df[col] = df[col].str.replace('.', '').str.replace(',', '.').astype('float')
    df['CAUSALE ABI'] = df['CAUSALE ABI'].astype(pd.Int32Dtype())

    # Replace 
    df['IMPORTO'] = df.apply(lambda df_row: df_row['AVERE'] if pd.isna(df_row['DARE']) else -df_row['DARE'], axis = 1)

    # Clean description
    df['DESC'] = df['DESCRIZIONE OPERAZIONE'].map(cleanDescription)
    df['DESC'] = df['DESC'].map(cleanIncomeDescription)
    
    # Remove the last empty column
    df = df.drop(columns = ['Unnamed: 4', 'Unnamed: 7', 'DARE', 'AVERE'])

    # Map the ABI codes
    df = mapAbiCodes(df, folderData = path.join(projectFolder, 'taxonomies'))

    # Map the categories
    df = addExpensiveCategories(df, folderData = path.join(projectFolder, 'taxonomies'))

    return df

def mapAbiCodes(df, folderData):

    # Load the taxonomies
    with open(path.join(folderData, 'causaliABI.json')) as jsonFile:
        causaliAbi = load(jsonFile)
    try:
        df['CAUSALE ABI'] = df['CAUSALE ABI'].map(lambda code: causaliAbi[str(code)] if not pd.isna(code) else "")
    except KeyError as e:
        print('Missing taxonomy for ABI code:', e)
        raise Exception()

    return df


def mapExpensives(taxonomy, desc):
    for expensiveName, expensiveCategory in taxonomy.items():
        if expensiveName.lower() in desc.lower():
            if 'paypal' in desc.lower() and ('IMP. E 5,99' in desc):
                return "Education & Culture"
            elif 'paypal' in desc.lower() and ('IMP. E 3,10' in desc or 'IMP. E 6,20' in desc):
                return 'Transportation'
            else:
                return expensiveCategory
    return 'Other'

def addExpensiveCategories(df, folderData):

    # Load taxonomy
    with open(path.join(folderData, 'expensiveCategories.json')) as jsonFile:
        expensiveCategories = load(jsonFile)
    expensiveMapping = {expensive: cat for cat, expList in expensiveCategories.items() for expensive in expList}

    # Map the expensive
    df.loc[df['IMPORTO'] < 0, 'CATEGORIA'] = df.loc[df['IMPORTO'] < 0,'DESCRIZIONE OPERAZIONE'].map(lambda desc: mapExpensives(expensiveMapping, desc))
    return df

def cleanDescription(desc):
    if 'vostra disposizione a favore' in desc.lower():
        start = desc.lower().find('a fav: ')
        end = desc.lower().find('id.msg')
        return desc[start: end if end > start else len(desc)].split('-')[0].strip(' -.')
    elif 'pagamento tramite pos' in desc.lower():

        start_pos1 = desc.lower().find('presso:') 
        start_pos2 = desc.lower().find('.00 ')
        start_pos3 = desc.lower().find(':')

        if start_pos1 > 0:
            start_pos = start_pos1 + 7
        elif start_pos2 > 0:
            start_pos = start_pos2 + 4
        elif start_pos3 > 0:
            start_pos = start_pos3 + 12
        else:
            start_pos = 0       
        return desc[start_pos:].strip(' -.')
    
    elif "pagamenti diversi cred. " in desc.lower():
        minimal = desc[desc.lower().find('cred. ') + 6: desc.lower().find('id.mandato')].strip(' -.')
        minimal = 'PayPal' if 'paypal' in minimal.lower() else minimal
        return minimal
    
    elif "imposte e tasse polizza" in desc.lower():
        return desc[desc.lower().find('periodo bollo'): ].strip(' -.')
    return desc

def cleanIncomeDescription(desc):
    if 'ordinante' and 'causale' in desc.lower():
        return desc[desc.lower().find('ordinante: ') + 11: desc.lower().find('causale')].strip(' -.')
    elif 'emolumenti' in desc.lower():
        return desc[desc.lower().find('per emolumenti') + 15: desc.lower().find('accredito competenze')].strip(' -.')
    elif 'cedole' in desc.lower():
        #if 'btp' in desc.lower():
        #    return desc[desc.lower().find('btp'): desc.lower().find('quantit')].strip(' -.')
        return desc[desc.lower().find('cedole ') + 7: desc.lower().find('quantit')].strip(' -.')
    return desc

def loadBudget():
    folderData = path.join(path.dirname(__file__), '..', 'taxonomies')
    budget = pd.read_excel(path.join(folderData, 'budget.xlsx'), sheet_name='Budget', index_col = 0, usecols = [0,1])
    budget = budget['BUDGET'].astype('int').to_dict()
    return budget

def saveDataframe(df, outputFolder = 'data'):

    # Reorder columns
    df = df[df.columns[:2] + df.columns[3:] + 'DESCRIZIONE OPERAZIONE']

    print(df.iloc[0])

    