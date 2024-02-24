from os import path
import numpy as np
import pandas as pd
import matplotlib.pylab as plt
import matplotlib.colors as mcolors
import matplotlib.ticker as mtick

from utils.dataLoader import loadBudget

def creteAreaPlots(df, outputFolder, feature = 'CATEGORIA', groupby = "TRIMESTRE"):

    if feature == 'CATEGORIA':
        budget = loadBudget()

    # Turn quarter names into string
    df[groupby] = df[groupby].astype('str')

    # Save the list of quarters
    quarters = df[groupby].unique()
    
    # Compute total expensive by category
    summedExpensiveByCategory = df[[feature, 'IMPORTO']].groupby(by = feature).sum().sort_values(by = 'IMPORTO', ascending = True)
    minValue, maxValue  = summedExpensiveByCategory['IMPORTO'].min(),  summedExpensiveByCategory['IMPORTO'].max()
    summedExpensiveByCategory /= summedExpensiveByCategory['IMPORTO'].sum() 
    summedExpensiveByCategory *= 100
    summedExpensiveByCategory = summedExpensiveByCategory.round(1)

    # Create the figure
    ncols = 3
    nrows = int(np.ceil(len(summedExpensiveByCategory) / ncols))
    fig, axes = plt.subplots(nrows = nrows, ncols = ncols, figsize = (7 * ncols, 4 * nrows), sharex = False)
    
    if axes.ndim == 1:
        axes = axes.reshape(1, -1)

    # Colors
    colors = plt.cm.get_cmap("Oranges_r")
    norm = mcolors.Normalize(vmin = minValue * 1.1, vmax = maxValue * 1.5)

    # Create the subplots
    idk_row, idk_col = 0, 0
    for category in summedExpensiveByCategory.index:
        
        # Isolate the subset of observations
        to_viz = df[df[feature] == category]

        # Create zero values for missing quarters
        toFill = set(quarters) - set(to_viz[groupby])
        if len(toFill) > 0:
            to_add = pd.DataFrame([{groupby:item, 'CATEGORIA':category,  'IMPORTO':0} for item in toFill])
            to_viz = pd.concat([to_viz, to_add]).sort_values(by = groupby, ascending = True)

        # Data
        x = to_viz[groupby]
        y = to_viz['IMPORTO']

        # Color normalizer
        summedQuarterExpensives = int(np.round(to_viz['IMPORTO'].sum()))
        normalized_color = norm(summedQuarterExpensives)

        # (1) Area plot
        axes[idk_row, idk_col].stackplot(x, y, color = colors(normalized_color), labels = [f"Total ({summedQuarterExpensives:,.0f} €)"])
        
        # (2) Scatter plot
        axes[idk_row, idk_col].scatter(x, y, s = 70, marker = 'o', color = colors(1 - normalized_color), edgecolors = "grey" )

        # (3) Median value
        axes[idk_row, idk_col].hlines(y = y.median(), xmin = x.iloc[0], xmax = x.iloc[-1], colors='grey', linestyles='--', lw = 2, alpha = 0.7,
                                      label=f"Median ({int(np.round(to_viz['IMPORTO'].median())):,.0f} €)")
        
        # (4) Add budget
        if feature == 'CATEGORIA':
            categoryBudget = budget[category] * -3 if groupby == 'TRIMESTRE' else budget[category] * -1
            
            axes[idk_row, idk_col].hlines(y = categoryBudget, xmin = x.iloc[0], xmax = x.iloc[-1],
                                          colors='firebrick', linestyles='-', lw = 2, alpha = 0.4,
                                          label=f"Budget ({categoryBudget}) €)")
        
        # Subplot settings
        categoryWeight = summedExpensiveByCategory.loc[category, 'IMPORTO']
        axes[idk_row, idk_col].set_title(r"$\bf{" + category.replace(' ', '~') + r'}$ ' + f"({int(categoryWeight) if categoryWeight.is_integer() else categoryWeight} %)", fontsize = 22)
        axes[idk_row, idk_col].yaxis.set_major_formatter(mtick.StrMethodFormatter('{x:,.0f} €')) 
        axes[idk_row, idk_col].xaxis.set_major_locator(mtick.MaxNLocator(nbins = x.unique().size)) 
        axes[idk_row, idk_col].legend()
    
        # Swicth subgraph
        if idk_col < ncols -1:
            idk_col += 1
        else:
            idk_col = 0
            idk_row += 1
    
    # Hide gempty graphs
    diff =  (ncols * nrows) - len(summedExpensiveByCategory)
    if diff > 0:
        for i in range(diff):
            axes[-1, -1 -i].axis('off')

    # Save the graphs
            #feature
    fig.tight_layout()
    fileName = "expensivesBy" + ('Quarters' if groupby == 'TRIMESTRE' else 'Month') + ('AbiCauses' if feature == 'CAUSALE ABI' else '') + '.png'
    plt.savefig(path.join(outputFolder, fileName))

def createBarPlots(df, outputFolder, feature = 'CATEGORIA', groupby = "TRIMESTRE"):
    
    # Turn quarter names into string
    df[groupby] = df[groupby].astype('str')

    print(df)