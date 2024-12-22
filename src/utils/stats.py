from configparser import ConfigParser
from os import path
from collections import Counter
import pandas as pd
import numpy as np

# LOCAL IMPORTS
from utils import graphs
from utils.dataLoader import loadBudget

def group_expensive(df, outputFolder, feature = "CAUSALE ABI", include_incomes = False): 
    
    df = df.copy()
    df['DESC'] += "!"
    df['#'] = 1
    
    # Data filtering
    if not include_incomes:
        df = df[df['IMPORTO'] < 0]

    # (1.a) Add expensives by description
    groupedByDesc = df[['DESC', 'IMPORTO', 'MESE']].groupby(by = ['DESC', 'MESE']).sum()

    # (1) Group expensives by month
    expensivesByMonth = df[['MESE', 'IMPORTO', 'DESC', '#', feature]].groupby(by = ['MESE', feature], as_index = True).sum()    # 
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['DESC'].str.split('!').map(
        lambda items: [item for item in items if item != ""] if isinstance(items, list) else items)
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['OPERAZIONI'].map(Counter).map(Counter.items)
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth.apply(
        lambda df_row: [expensive + tuple([groupedByDesc.loc[(expensive[0] + "!", df_row.name[0]), 'IMPORTO'].round(2) * -1])
                        for expensive in df_row['OPERAZIONI']], axis = 1)
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['OPERAZIONI'].map(lambda items: sorted(items, key = lambda item: (item[2], item[1]), reverse = True))
    expensivesByMonth['OPERAZIONI'] = expensivesByMonth['OPERAZIONI'].map(
        lambda items: ' | '.join([item[0] + (f" (x{item[1]}, " if item[1] > 1 else ' (') + str(int(item[2]) if item[2] >= 1 else item[2]) + " €)"
                                   for item in items]))
    
    expensivesByMonth = expensivesByMonth.drop(columns = ['DESC']).sort_index(ascending = False)

    # (2) Group expensives by code
    df['TRIMESTRE'] = df['TRIMESTRE'].dt.strftime('Q%q')
    groupedByCategory = df[['TRIMESTRE', 'ANNO',  'IMPORTO', feature]].groupby(by = [feature, 'ANNO', 'TRIMESTRE']).sum() 

    # (2.a) Sort index by code importance
    topAbiCodesByExpensive = df[['IMPORTO', feature]].groupby(by = [feature]).sum().sort_values(by ='IMPORTO', ascending = True).index.to_list()
    groupedByCategory['RANK'] = [topAbiCodesByExpensive.index(entry[0]) for entry in groupedByCategory.index] 

    groupedByCategory = groupedByCategory.sort_values(by = ['RANK', 'ANNO', 'TRIMESTRE',  'IMPORTO'], ascending = [True, False, False, False]).drop(columns = 'RANK')

    # (2.b) Round the imports
    groupedByCategory['IMPORTO'] = groupedByCategory['IMPORTO'].round(0)

    # (2.c) Add budget
    if feature == 'CATEGORIA':
        budget = loadBudget()

        try:
            groupedByCategory['Δ BUDGET'] = groupedByCategory.apply(
                lambda df_row: round(- df_row['IMPORTO'] - (budget[df_row.name[0]] * 3), 0) , axis = 1)
            
            groupedByCategory['Δ BUDGET (%)'] = groupedByCategory.apply(
                lambda df_row: df_row['Δ BUDGET'] / (budget[df_row.name[0]] * 3) if budget[df_row.name[0]] > 0 else 9.99, axis = 1)
        
        except KeyError as missingBudgetCategory:
            raise Exception(f'\n{missingBudgetCategory} is not in the budget! Please include it in the budget file.\n')

    # Monthly stats
    warnings = []
    monthly_dfs = dict()
    for month in expensivesByMonth.index.get_level_values(0).unique():
        partial_df = expensivesByMonth.loc[month, :].sort_values(by = 'IMPORTO', ascending = True) 

        # Add new columns 
        partial_df.insert(loc = 1, column = '%', value = (partial_df['IMPORTO'] / partial_df['IMPORTO'].sum()).round(2))

        # Delta from budget
        if feature == 'CATEGORIA':
            partial_df.insert(loc = 2, column = 'Δ BUDGET', value = partial_df.apply(
                lambda df_row: round(-df_row['IMPORTO'] - budget[df_row.name], 0), axis = 1))
            partial_df.insert(loc = 3, column = 'Δ BUDGET (%)', value = partial_df.apply(
                lambda df_row: df_row['Δ BUDGET'] / budget[df_row.name] if budget[df_row.name] > 0 else 9.99, axis = 1))

            partial_df.loc[''] = None
            partial_df.loc['_TOTAL'] = {'IMPORTO': partial_df['IMPORTO'].sum(), 'Δ BUDGET' : partial_df['Δ BUDGET'].sum(),
                                        'Δ BUDGET (%)': partial_df['Δ BUDGET'].sum() / np.sum(list(budget.values()))}
            
            partial_df.insert(loc = 3, column = "!", value = partial_df['Δ BUDGET (%)'].map(lambda x: 1 if x >= 0.5 else 0 if x >=0 else -1))
            partial_df.loc[['_TOTAL', ''], '!'] = None

            monthlyWarnings = partial_df.loc[partial_df['!'] == 1, ['Δ BUDGET (%)']]
            monthlyWarnings.insert(loc = 0, column = 'MESE', value = month)

            if 'Investments' in monthlyWarnings.index:
                monthlyWarnings = monthlyWarnings.drop(index = 'Investments')
            monthlyWarnings = monthlyWarnings.reset_index()
            
            warnings.append(monthlyWarnings)
        monthly_dfs[month] = partial_df

    # Save the excel file
    fileName = 'expensives' + ('byAbiCode' if feature == "CAUSALE ABI" else '') + '.xlsx'
    try:    
        with pd.ExcelWriter(path.join(outputFolder, fileName), engine = 'xlsxwriter') as excelFile:
            groupedByCategory.to_excel(excelFile, sheet_name = 'Overview', freeze_panes = (1,1))

            euro_fmt = excelFile.book.add_format({'num_format': '#,##0 €', 'font_size': 16})
            perc_fmt = excelFile.book.add_format({"num_format": "0%", 'font_size': 16})
            index_fmt = excelFile.book.add_format({"align": "left", 'bold': True, 'font_size': 16, 'align': 'vcenter'})
            header_format = excelFile.book.add_format({'bg_color': '#3D3B40', 'font_color': 'white', 'bold': False, 'valign': 'center', 'font_size': 20})
            excelFile.book.formats[0].set_font_size(16)
            excelFile.book.formats[0].set_align('vcenter')
            excelFile.book.formats[0].set_align('center')
            excelFile.sheets['Overview'].set_row(0, None, header_format)
            excelFile.sheets['Overview'].set_column('D:E', cell_format = euro_fmt)
            excelFile.sheets['Overview'].set_column('F:F', cell_format = perc_fmt) #  width = 12, 
            excelFile.sheets['Overview'].set_column('A:A', cell_format = index_fmt)
            #excelFile.sheets['Overview'].set_column('A:A', width = 20)
            excelFile.sheets['Overview'].conditional_format('E1:E999', {
                'type': '3_color_scale', 'min_color': "#99BC85",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#C83E3E"})
            excelFile.sheets['Overview'].conditional_format('F1:F999', {
            'type': '3_color_scale', 'min_color': "#99BC85",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#C83E3E"})
            excelFile.sheets['Overview'].conditional_format('D1:D999', {
                        'type': '2_color_scale', 'min_color': '#C83E3E', 'max_color': '#E6A8A8'})
            excelFile.sheets['Overview'].autofit()

            if len(warnings) > 0:
                warnings = pd.pivot_table(pd.concat(warnings), index=['CATEGORIA', 'MESE'])
                
                ranks = warnings[['Δ BUDGET (%)']].groupby('CATEGORIA').sum().sort_values(by = 'Δ BUDGET (%)', ascending=False).index.to_list()
                warnings['ranks'] = [ranks.index(cat[0]) for cat in warnings.index]
                warnings = warnings.sort_values(by = ['ranks', 'Δ BUDGET (%)', 'MESE'], ascending = [True, False, False]).drop(columns = 'ranks')
    
                warnings.to_excel(excelFile, sheet_name = 'Warnings', freeze_panes = (1,1)) 
                excelFile.sheets['Warnings'].set_column('A:A', cell_format = index_fmt)
                excelFile.sheets['Warnings'].set_column('C:C', cell_format = perc_fmt)
                excelFile.sheets['Warnings'].set_row(0, None, header_format)
                excelFile.sheets['Warnings'].conditional_format(f'C2:C{len(warnings) + 1}', {
                        'type': '2_color_scale', 'min_color': '#E6A8A8', 'max_color': '#C83E3E'})
                excelFile.sheets['Warnings'].autofit()

            for month, monthly_df in monthly_dfs.items():

                # Save the sheet
                sheetName = month.strftime('%B %Y')
                monthly_df.to_excel(excelFile, sheet_name = sheetName, index = True,  freeze_panes = (1,1))

                # Graphical settings
                excelFile.sheets[sheetName].set_row(0, None, header_format)
                excelFile.sheets[sheetName].set_column('A:A', cell_format = index_fmt)
                excelFile.sheets[sheetName].set_column('B:B', cell_format = euro_fmt) #  width = 8, 
                excelFile.sheets[sheetName].set_column('C:C', cell_format = perc_fmt) #  width = 5,
                excelFile.sheets[sheetName].set_column(first_col = len(partial_df.columns), last_col =len(partial_df.columns), 
                                                    width = 60, cell_format = excelFile.book.add_format({"align": "left", 'font_size': 16}))
                
                excelFile.sheets[sheetName].conditional_format(f'B2:B{len(monthly_df) - 1}', {
                        'type': '2_color_scale', 'min_color': '#C83E3E', 'max_color': '#E6A8A8'})

                if feature == 'CATEGORIA':
                    excelFile.sheets[sheetName].set_column('D:D', width = 8, cell_format = euro_fmt)
                    excelFile.sheets[sheetName].set_column('F:F', width = 12, cell_format = perc_fmt)
                    excelFile.sheets[sheetName].set_column('E:E', width = 2)
                    excelFile.sheets[sheetName].set_column('G:G', width = 2)

                    excelFile.sheets[sheetName].conditional_format(f'C2:C{len(monthly_df) - 1}', {
                        "type": "data_bar", "min_type": "num", "max_type": "num", "min_value": 0, "max_value": 1, 
                        "bar_color": "#CFD8DC", "bar_solid": True, "bar_only": False, "bar_direction":'right'})
                    excelFile.sheets[sheetName].conditional_format(f'E2:E{len(monthly_df) - 1 }', {
                        'type': 'icon_set', 'icon_style': '3_symbols_circled', 'icons_only': True, 'reverse_icons': True,
                        'icons': [
                            {'criteria': '>=', 'type': 'number', 'value': 1},
                            {'criteria': '<=', 'type': 'number', 'value': 0},
                            {'criteria': '<',  'type': 'number', 'value': -1}]
                        })
                    
                    excelFile.sheets[sheetName].conditional_format(f'D2:D{len(monthly_df) -1}', {
                        'type': '3_color_scale', 'min_color': "#99BC85",'mid_color': "white", 'mid_type': 'num', 'mid_value': 0, 'max_color': "#C83E3E"})
                    excelFile.sheets[sheetName].conditional_format(f'F2:F{len(monthly_df) - 1}', {
                        'type': '3_color_scale', 'min_color': "#99BC85",'mid_color': "white", 'mid_type': 'num',  'mid_value': 0, 'max_color': "#C83E3E"})
                excelFile.sheets[sheetName].autofit()
            
            if feature == 'CATEGORIA':
                mapping = dict(zip(map(str, monthly_dfs.keys()), list(excelFile.sheets.keys())[2:]))
                for idk, item in enumerate(warnings.index.get_level_values(1).astype(str)):
                    excelFile.sheets['Warnings'].write_url(idk + 1, 1, f"internal:'{mapping[item]}'!A1", string = item, 
                                                        cell_format = excelFile.book.add_format({'font_size': 16, 'font_color': 'black'}))
    except PermissionError:
        raise PermissionError("Close the spreadsheet...")
    print("[DONE] Grouped expensive by:", feature, "\n")


def expensive_graph(df, outputFolder, groupby = 'TRIMESTRE', feature = "CATEGORIA"):
    df = df[df['IMPORTO'] < 0]

    groupedByCategory = df[[groupby, feature, 'IMPORTO']].groupby(by = [groupby, feature], as_index=False).sum() 
    graphs.creteAreaPlots(groupedByCategory, outputFolder, feature = feature, groupby = groupby)

    print(f"[DONE] GRAPH of {feature} by {groupby}\n")

def compute_incomes(df, outputFolder):

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
          
        # Formats
        excelFile.book.formats[0].set_font_size(16)
        excelFile.book.formats[0].set_align('vcenter')

        euro_fmt = excelFile.book.add_format({'num_format': '#,##0 €', 'font_size': 16})
        header_format = excelFile.book.add_format({'bg_color': '#3D3B40', 'font_color': 'white', 'bold': False, 'valign': 'center', 'font_size': 16})
        for featureName, grouped_df in stats.items():
            grouped_df.to_excel(excelFile, sheet_name = featureName)

            # Graph settings
            sheet = excelFile.sheets[featureName]
            sheet.set_row(0, None, header_format)
            value_col = 'B1:B999' if featureName == 'Overview' else 'C1:C999' 
            sheet.set_column(value_col, width = 8, cell_format = euro_fmt)
            sheet.conditional_format(value_col, {'type': '2_color_scale', 'min_color': '#E1F0DA', 'max_color': '#99BC85'})
            sheet.autofit()

def cutoff_period():

    # Read config file
    config = ConfigParser()
    config.read(path.join('src','config.ini'))

    # Read the cutoff date and cash amount
    cutoff_date = pd.to_datetime(config.get('CUTOFF', 'date'))
    cutoff_cashAmount = float(config.get('CUTOFF', 'cash_amount'))

    return cutoff_date, cutoff_cashAmount

def monthly_stats(df, outputFolder):

    # Save the period
    period = pd.Series({'Last Transaction': df['VALUTA'].iloc[0], 'First Transaction': df['VALUTA'].iloc[-1]}, name  = 'Date')
    period = period.dt.strftime('%d/%m/%Y')

    # Create the macro-category
    df['MACRO-CATEGORIA'] = df['IMPORTO'].map(lambda value: 'ENTRATE' if value > 0 else 'USCITE')
    df.loc[df['CATEGORIA'] == 'Investments', 'MACRO-CATEGORIA'] = 'INVESTIMENTI'

    # Sort the dataframe
    df = df.sort_values(by = 'DATA').reset_index(drop = True)

    # Get the cutoff date and amount
    cutoff_date, cutoff_amount = cutoff_period()
    actual_cutoff_date = df['DATA'].iloc[0]
    assert np.abs(actual_cutoff_date - cutoff_date) < np.timedelta64(7, 'D'), f"The cutoff date ({cutoff_date}) is not correct! First date found: {actual_cutoff_date}"
    
    # Add the cash amount
    cutoff_transaction = df[df['DATA'] <= actual_cutoff_date].iloc[-1].name
    df.loc[cutoff_transaction:, "LIQUIDITA"] = cutoff_amount + df.loc[cutoff_transaction:, 'IMPORTO'].cumsum()

    #df.to_excel(path.join(outputFolder, 'tmp.xlsx'))

    # Group the months
    groupedDf = df[['MESE', 'IMPORTO', 'MACRO-CATEGORIA']].groupby(by = ['MESE', 'MACRO-CATEGORIA']).sum() #, as_index = False

    # Create the new data representation
    monthlyStats = []
    for month in groupedDf.index.get_level_values(0).unique():

        # Stats
        stats = groupedDf.loc[month, 'IMPORTO']
        stats['MESE'] = month

        # Cash 
        cash = df.loc[df['VALUTA'].dt.to_period('M') == month, 'LIQUIDITA'].dropna()
        if len(cash) > 0:
            stats["LIQUIDITA'"] = cash.iloc[-1]
        monthlyStats.append(stats)
    monthlyStats = pd.DataFrame(monthlyStats)
    monthlyStats = monthlyStats[['MESE', 'ENTRATE', 'USCITE', 'INVESTIMENTI', "LIQUIDITA'"]]
    monthlyStats = monthlyStats.fillna(value = 0).set_index('MESE')
    monthlyStats = monthlyStats.sort_index(ascending = False)

    # Save the findings
    colors = {'ENTRATE': {'MIN': '#C8E6C9', 'MAX': '#388E3C'}, 
              'USCITE': {'MAX': '#EF9A9A', 'MIN': '#E53935'}, 
              'INVESTIMENTI': {'MAX': '#81D4FA', 'MIN': '#0288D1'}, 
              "LIQUIDITA'": {'MAX': '#F4511E', 'MIN': '#FFCCBC'}}
    with pd.ExcelWriter(path.join(outputFolder, 'monthlyStats.xlsx'),  engine = 'xlsxwriter', datetime_format="mmmm yyyy") as excelFile:

        # Save the main sheet
        monthlyStats.to_excel(excelFile, sheet_name = 'Months', index = True, freeze_panes=(1, 0))
        sheet = excelFile.sheets['Months']

        # Graphical settings
        excelFile.book.formats[0].set_font_size(18)
        
        header_format = excelFile.book.add_format({'bg_color': '#3D3B40', 'font_color': 'white', 'bold': False, 'valign': 'center', 'font_size': 18})
        grey_format = excelFile.book.add_format({'bg_color': '#EEEEEE','font_size': 18})
        euro_fmt = excelFile.book.add_format({'num_format': '#,##0 €', 'font_size': 18})
        sheet.set_column('B1:E999', cell_format = euro_fmt)

        for col_idk, colName in enumerate(monthlyStats.columns):
            sheet.conditional_format(0, col_idk + 1, 999, col_idk + 1, {
                    'type': '2_color_scale', 'min_color': colors[colName]['MIN'], 'max_color': colors[colName]['MAX']})
        sheet.set_row(0, None, header_format)

        sheet.conditional_format(f'A2:A{len(monthlyStats)}', {
            'type':'formula', 'criteria': "=MOD(ROW(),2)=0", 'format':  grey_format})
      
        sheet.autofit()
        period.to_excel(excelFile, sheet_name = 'Period', index = True)