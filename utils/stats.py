import pandas as pd
from os import path
from collections import Counter

from utils import graphs
from utils.dataLoader import loadBudget

import numpy as np

def groupExpensives(df, outputFolder, feature = "CAUSALE ABI", include_incomes = False, cutoff_year = None, verbose = False): 

    df['DESC'] += "!"
    df['#'] = 1
    
    # Data filtering
    if not include_incomes:
        df = df[df['IMPORTO'] < 0]
    if cutoff_year != None:
        df = df[df['ANNO'] >= cutoff_year]

    # (1) Group expensives by month
    expensivesByMonth = df[['MESE', 'IMPORTO', 'DESC', '#', feature]].groupby(by = ['MESE', feature], as_index = True).sum()


    # (1.a) Add expensives by description
    groupedByDesc = df[['DESC', 'IMPORTO', 'MESE']].groupby(by = ['DESC', 'MESE']).sum()

    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['DESC'].str.split('!').map(
        lambda items: [item for item in items if item != ""] if isinstance(items, list) else items)
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['OPERAZIONI'].map(Counter).map(Counter.items)
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth.apply(
        lambda df_row: [expensive + tuple([groupedByDesc.loc[(expensive[0] + "!", df_row.name[0]), 'IMPORTO'].round(2) * -1])
                        for expensive in df_row['OPERAZIONI']], axis = 1)
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['OPERAZIONI'].map(lambda items: sorted(items, key = lambda item: (item[2], item[1]), reverse = True))
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['OPERAZIONI'].map(
        lambda items: '\n '.join([item[0] + (f" (x{item[1]}, " if item[1] > 1 else ' (') + str(int(item[2]) if item[2] >= 1 else item[2]) + " €)"
                                   for item in items]))
    expensivesByMonth = expensivesByMonth.drop(columns = ['DESC']).sort_index(ascending = False)

    # (2) Group expensives by code
    df['TRIMESTRE'] = df['TRIMESTRE'].dt.strftime('Q%q')
    groupedByCategory = df[['TRIMESTRE', 'ANNO',  'IMPORTO', feature]].groupby(by = [feature, 'ANNO', 'TRIMESTRE']).sum() 

    # (2.a) Sort index by code importance
    topAbiCodesByExpensive = df[['IMPORTO', feature]].groupby(by = [feature]).sum().sort_values(by ='IMPORTO', ascending = True).index.to_list()
    groupedByCategory['RANK'] = [topAbiCodesByExpensive.index(entry[0]) for entry in groupedByCategory.index] 
    groupedByCategory = groupedByCategory.sort_values(by = ['RANK', 'TRIMESTRE',  'IMPORTO'], ascending = [True, False, False]).drop(columns = 'RANK')

    # (2.b) Round the imports
    groupedByCategory['IMPORTO'] = groupedByCategory['IMPORTO'].round(0)

    # (2.c) Add budget
    if feature == 'CATEGORIA':
        budget = loadBudget()
        groupedByCategory['Δ Budget'] = groupedByCategory.apply(lambda df_row: round(- df_row['IMPORTO'] - (budget[df_row.name[0]] * 3), 0) , axis = 1)
        groupedByCategory['Δ Budget (%)'] = groupedByCategory.apply(
            lambda df_row: df_row['Δ Budget'] / (budget[df_row.name[0]] * 3) if budget[df_row.name[0]] > 0 else 1, axis = 1)

    # Save the excel file
    fileName = 'expensiveBy' + ('AbiCode' if feature == "CAUSALE ABI" else 'Category') + '.xlsx'
    with pd.ExcelWriter(path.join(outputFolder, fileName), engine = 'xlsxwriter') as excelFile:
        groupedByCategory.to_excel(excelFile, sheet_name = 'Overview', freeze_panes = (1,1))

        euro_fmt = excelFile.book.add_format({'num_format': '#,##0 €'})
        perc_fmt = excelFile.book.add_format({"num_format": "0%"})
        excelFile.sheets['Overview'].set_column('D:E', width = 8, cell_format = euro_fmt)
        excelFile.sheets['Overview'].set_column('F:F', width = 12, cell_format = perc_fmt)
        excelFile.sheets['Overview'].set_column('A:A', width = 20)
        excelFile.sheets['Overview'].conditional_format('E1:E999', {
            'type': '3_color_scale', 'min_color': "#4CAF50",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#EF5350"})
        excelFile.sheets['Overview'].conditional_format('F1:F999', {
            'type': '3_color_scale', 'min_color': "#4CAF50",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#EF5350"})

        # Monthy stats
        for month in expensivesByMonth.index.get_level_values(0).unique():
            partial_df = expensivesByMonth.loc[month, :].sort_values(by = 'IMPORTO', ascending = True) 

            # Add new columns 
            # Percentages
            partial_df.insert(loc = 1, column = '%', value = (partial_df['IMPORTO'] / partial_df['IMPORTO'].sum()).round(2))

            # Delta from budget
            if feature == 'CATEGORIA':
                partial_df.insert(loc = 2, column = 'Δ Budget', value = partial_df.apply(
                    lambda df_row: round(-df_row['IMPORTO'] - budget[df_row.name], 0), axis = 1))
                partial_df.insert(loc = 3, column = 'Δ Budget (%)', value = partial_df.apply(
                    lambda df_row: df_row['Δ Budget'] / budget[df_row.name] if budget[df_row.name] > 0 else 1, axis = 1))

                partial_df.loc['_TOTAL', 'IMPORTO'] = partial_df['IMPORTO'].sum()
                partial_df.loc['_TOTAL', 'Δ Budget'] = partial_df['Δ Budget'].sum()
                
                totalBudget = np.sum(list(budget.values()))
                if totalBudget > 0:
                    partial_df.loc['_TOTAL', 'Δ Budget (%)'] = partial_df['Δ Budget'].sum() / totalBudget 
            
            # Save the sheet
            sheetName = month.strftime('%B %Y')
            partial_df.to_excel(excelFile, sheet_name = sheetName, index = True)

            # Graphical settings
            excelFile.sheets[sheetName].set_column('B:B', width = 8, cell_format = euro_fmt)
            excelFile.sheets[sheetName].set_column('C:C', width = 5, cell_format = perc_fmt)
            excelFile.sheets[sheetName].set_column('A:A', width = 20)
            excelFile.sheets[sheetName].set_column(first_col = len(partial_df.columns), last_col =len(partial_df.columns), width = 150)

            if feature == 'CATEGORIA':
                excelFile.sheets[sheetName].set_column('D:D', width = 8, cell_format = euro_fmt)
                excelFile.sheets[sheetName].set_column('E:E', width = 12, cell_format = perc_fmt)
                excelFile.sheets[sheetName].conditional_format(f'D1:D{len(partial_df) -1}', {
                    'type': '3_color_scale', 'min_color': "#4CAF50",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#EF5350"})
                excelFile.sheets[sheetName].conditional_format(f'E1:E{len(partial_df) -1}', {
                    'type': '3_color_scale', 'min_color': "#4CAF50",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#EF5350"})

    if verbose:
        print(groupedByCategory)

    print("[DONE] Grouped expensive by:", feature, "\n")


def visualizeExpensives(df, outputFolder, cutoff_year = None, feature = "CATEGORIA", groupby = 'TRIMESTRE'):
    if cutoff_year != None:
        df = df[df['ANNO'] >= cutoff_year]
    df = df[df['IMPORTO'] < 0]

    if len(df) == 0:
        print(f"\n[ERROR] No expensives found for the year {cutoff_year}.\n")
        return

    groupedByCategory = df[[groupby, feature, 'IMPORTO']].groupby(by = [groupby, feature], as_index=False).sum() 
    graphs.creteAreaPlots(groupedByCategory, outputFolder, feature = feature, groupby = groupby)

    print(f"[DONE] Graph ({groupby})\n")

def computeIncomes(df, outputFolder):

    col_to_group = ['ANNO', 'TRIMESTRE','MESE']
    
    # Filter only incomes
    df = df[df['IMPORTO'] > 0]

    # Filter only relevant columns
    df = df[['CAUSALE ABI', 'DESC', 'IMPORTO'] + col_to_group]
    df['DESC'] += "!"
    df['#'] = 1

    stats = dict()

    # Compute the overview
    stats['Overview'] = df[['CAUSALE ABI', 'IMPORTO']].groupby(by = ['CAUSALE ABI']).sum().sort_values(by = 'IMPORTO', ascending = False)
    orderedFeatures = stats['Overview'].index.to_list()

    # Group incomes
    for col in col_to_group:
        grouped_df = df[['CAUSALE ABI', 'IMPORTO', '#','DESC', col]].groupby(by = ['CAUSALE ABI', col]).sum()
        
        groupedByDesc = df[['DESC', 'IMPORTO', col]].groupby(by = ['DESC', col]).sum()

        # Add the descriptions
        grouped_df['DESC'] = grouped_df['DESC'].str.split('!').map(lambda items: [item for item in items if item != ""]).apply(Counter)
        grouped_df['DESC'] = grouped_df['DESC'].map(lambda count: sorted(count.items(), key = lambda item: item[1], reverse=True))

        grouped_df['DESC'] = grouped_df.apply(
            lambda df_row: [expensive + tuple([groupedByDesc.loc[[(expensive[0] + '!', df_row.name[1])], 'IMPORTO'].round(2).iloc[0]])
                            for expensive in df_row['DESC']], axis = 1)
        grouped_df['DESC'] = grouped_df['DESC'].map(lambda items: '\n '.join([item[0] + 
                                                                             (f" (x{item[1]}, " if item[1] > 1 else ' (') + 
                                                                             str(int(item[2]) if item[2] >= 1 else item[2]) + " €)"
                                                                             for item in items]))

        # Sort dataframe
        grouped_df['rank'] = grouped_df.apply(lambda df_row: orderedFeatures.index(df_row.name[0]), axis = 1)
        grouped_df = grouped_df.sort_values(by = ['rank', col], ascending = [True, False]).drop(columns = ['rank'])

        stats[col] = grouped_df

    # Save the stats
    with pd.ExcelWriter(path.join(outputFolder, 'incomes.xlsx')) as excelFile:
        for featureName, grouped_df in stats.items():
            grouped_df.to_excel(excelFile, sheet_name = featureName)


def monthlyStats(df, outputFolder):

    # Save the period
    period = pd.Series({'Last Transaction': df['VALUTA'].iloc[0], 'First Transaction': df['VALUTA'].iloc[-1]}, name  = 'Date')
    period = period.dt.strftime('%d/%m/%Y')

    # Create the macro-category
    df['MACRO-CATEGORIA'] = df['IMPORTO'].map(lambda value: 'ENTRATE' if value > 0 else 'USCITE')
    df.loc[df['CATEGORIA'] == 'Investments', 'MACRO-CATEGORIA'] = 'INVESTIMENTI'

    # Group the months
    groupedDf = df[['MESE', 'IMPORTO', 'MACRO-CATEGORIA']].groupby(by = ['MESE', 'MACRO-CATEGORIA']).sum() #, as_index = False

    # Create the new data representation
    monthlyStats = []
    for month in groupedDf.index.get_level_values(0).unique():
        stats = groupedDf.loc[month, 'IMPORTO']
        stats['MESE'] = str(month)
        monthlyStats.append(stats)
    monthlyStats = pd.DataFrame(monthlyStats)
    monthlyStats = monthlyStats[['MESE', 'ENTRATE', 'USCITE', 'INVESTIMENTI']]
    monthlyStats = monthlyStats.fillna(value = 0).set_index('MESE')

    # Save the findings
    with pd.ExcelWriter(path.join(outputFolder, 'monthlyStats.xlsx'),  engine = 'xlsxwriter') as excelFile:

        # Save the main sheet
        monthlyStats.to_excel(excelFile, sheet_name = 'Months', index = True)

        # Graphical settings
        colors = {'ENTRATE': {'MIN': '#C8E6C9', 'MAX': '#388E3C'}, 
                  'USCITE': {'MAX': '#EF9A9A', 'MIN': '#E53935'}, 
                  'INVESTIMENTI': {'MAX': '#81D4FA', 'MIN': '#0288D1'}}
        euro_fmt = excelFile.book.add_format({'num_format': '#,##0 €'})
        for col_idk, colName in enumerate(monthlyStats.columns):
            excelFile.sheets['Months'].conditional_format(0, col_idk + 1, 999, col_idk + 1, {
                    'type': '2_color_scale', 'min_color': colors[colName]['MIN'], 'max_color': colors[colName]['MAX']})
            excelFile.sheets['Months'].set_column(first_col = col_idk + 1, last_col = col_idk + 1, width = len(colName) + 2, cell_format = euro_fmt)

        period.to_excel(excelFile, sheet_name = 'Period', index = True)