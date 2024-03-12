from os import path, listdir, remove, makedirs
from json import load
from hashlib import md5
from pandas.io.formats.excel import ExcelFormatter
import pandas as pd
import shutil

def initFolders(projectFolder):
    dataFolder = path.join(projectFolder, 'data')
    if not path.exists(dataFolder):
        makedirs(dataFolder)

    outputFolder = path.join(projectFolder, 'outputs')
    graphFolder = path.join(outputFolder, 'graphs')
    if not path.exists(graphFolder):
        makedirs(graphFolder)

    return dataFolder, outputFolder, graphFolder

def importCSVTransactions(projectFolder, importPath):
    if not path.exists(importPath):
        print(f"\n[WARNING] The import path does not exist! {importPath} --> (A) Change che import path in the app.py file.--> (B) Put the file into the ./data folder.\n")
        return False

    for fileName in listdir(importPath):
        if  fileName.endswith('.csv') and 'listamovimenti' in fileName.lower():
            try:
                shutil.move(src = path.join(importPath, fileName), dst = path.join(projectFolder, 'data'))
                print("--> Imported file:", fileName, "\n")
                return fileName
            except shutil.Error:
                return None
    return None

def loadTransactions(projectFolder):

    outputfileName = 'Transactions.xlsx'

    # Attach new transactions
    df = importTransactions(projectFolder)
    
    # Import the main dataframe
    dataFolder = path.join(projectFolder, 'data')
    
    folderFiles = listdir(dataFolder)
    if outputfileName in folderFiles:

        # Load the excel file
        main_df = pd.read_excel(path.join(dataFolder, outputfileName), sheet_name='Transactions') 
        main_df['ID'] = main_df['DESCRIZIONE OPERAZIONE'].map(lambda desc: md5(desc.encode('UTF-8')).hexdigest())

        # Upload the categories
        main_df = addExpensiveCategories(main_df, folderData = path.join(projectFolder, 'taxonomies'))

        # Merge the dataframes
        df = pd.concat([main_df, df]).drop_duplicates(subset=['VALUTA', 'IMPORTO']).reset_index(drop = True)

    if len(df) == 0:
        raise Exception('No data! Neither in the DATA folder nor in the DOWNLOAD folder.\n')

    # Generate the ID (M5 Hash) for each transaction
    df['ID'] = df['DESCRIZIONE OPERAZIONE'].map(lambda desc: md5(desc.encode('UTF-8')).hexdigest())
    
    # Sort the new dataframe
    df = df.sort_values(by = ['VALUTA', 'DATA'], ascending = False)

    # Create the month column 
    df['MESE'] = df['VALUTA'].dt.to_period('M') #.strftime('%B %Y')
    df['TRIMESTRE'] = df['VALUTA'].dt.to_period('Q')
    df['ANNO'] = df['VALUTA'].dt.to_period('Y').dt.year

    # Erase the folder 
    for fileName in folderFiles:
       file_path = path.join(dataFolder, fileName)
       if path.isfile(file_path):
           remove(file_path)

    # Save the dataframe
    ExcelFormatter.header_style = None
    with pd.ExcelWriter(path.join(dataFolder, outputfileName),  engine = 'xlsxwriter', datetime_format="d mmm yyyy") as excelFile:
        
        # Main
        df.drop(columns = ['ID']).to_excel(excelFile, index = False, sheet_name = 'Transactions', freeze_panes = (1,0))

        # IDs
        columns = ['ID', 'VALUTA', 'CATEGORIA','IMPORTO' ,'DESCRIZIONE OPERAZIONE']
        df[columns].to_excel(excelFile, index = False, sheet_name = 'IDs', freeze_panes = (1, 0))
        
        # Graphical settings
        grey_format = excelFile.book.add_format({'bg_color': '#EEEEEE'})
        white_format = excelFile.book.add_format({'bg_color': '#FFFFFF'})
        header_format = excelFile.book.add_format({'bg_color': '#9DBC98', 'font_color': 'white', 'bold': False, 'valign': 'center'})
   
        for sheet in excelFile.sheets.values():
            for idk_row in range(1, len(df)):
                sheet.set_row(idk_row, cell_format = white_format if idk_row % 2 ==0 else grey_format)

            sheet.set_row(0, None, header_format)
            sheet.autofit()

        # Colorbar
        excelFile.sheets['Transactions'].conditional_format('E2:E999', {
            'type': '3_color_scale', 'min_type': 'percentile', 'min_value': 5, 'min_color': "#FF8080", 
            'mid_color': "white", 'mid_type': 'num', 'mid_value': 0, 'max_color': "#99BC85"})
        
        # Hidden columns
        for i in [0, 2]:
            excelFile.sheets['Transactions'].set_column(first_col = i, last_col = i, options = {"hidden": True})

        # Set custom width
        #excelFile.sheets['Transactions'].set_column(first_col = 1, last_col = 1, width = 11)
        excelFile.sheets['Transactions'].set_column(first_col = 5, last_col = 5, width = 50)
        #excelFile.sheets['Transactions'].set_column(first_col = 10, last_col = 10, width = 10)

        # Autofilter
        excelFile.sheets['Transactions'].autofilter('A1:J9999')

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
    df['VALUTA'] = pd.to_datetime(df['VALUTA'], dayfirst=True)

    # Modify the numerical representation
    for col in ['AVERE', 'DARE']:
        cond = ~df[col].isna()
        if cond.any():
            df.loc[cond, col] = df.loc[cond, col].str.replace('.', '').str.replace(',', '.').astype('float')
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
        raise Exception(f"Missing taxonomy for ABI code: {e}\n")
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

    # Load expensive taxonomy
    with open(path.join(folderData, 'expensiveCategories.json')) as jsonFile:
        expensiveCategories = load(jsonFile)
    expensiveMapping = {expensive: cat for cat, expList in expensiveCategories.items() for expensive in expList}

    # Map the expensive
    expensiveFilter_cond = df['IMPORTO'] < 0
    df.loc[expensiveFilter_cond, 'CATEGORIA'] = df.loc[expensiveFilter_cond,'DESCRIZIONE OPERAZIONE'].map(
        lambda desc: mapExpensives(expensiveMapping, desc))
        
    # Load one-off transaction
    with open(path.join(folderData, 'oneOffTransactions.json')) as jsonFile:
        oneOffTransactions = load(jsonFile)

    # Map the one-off expensive
    for category, transactions in oneOffTransactions.items():
        for transaction_id in transactions:
            df.loc[df['ID'] == transaction_id, "CATEGORIA"] = category
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

def loadBudget(folderData = path.join('taxonomies')): #path.join(path.dirname(__file__), '..', 'taxonomies')
    budget = pd.read_excel(path.join(folderData, 'budget.xlsx'), sheet_name='Budget', index_col = 0, usecols = [0,1])
    budget = budget['BUDGET'].astype('int').to_dict()
    return budget    