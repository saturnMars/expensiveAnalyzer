# Analisi delle spese - Inbank

## Requisiti 
- Python 3 ([Windows Store](https://apps.microsoft.com/detail/9NRWMJP3717K))

## Installazione
1. Eseguire *./init.bat*

## Setup
1. Aggiornare la tassonomia delle spese nel file *./taxonomies/expensiveCategories.json*. 
    - Personalizzare la tassonomia con le proprie spese (transazioni scononosciute vengono catalogate come "Other")
    - Utilizzare la descrizione del pagamento ed inserire la sottostringa che identifica la spesa, ad es.: "Pagamento tramite POS DATA/ORA ... COOP BOLOGNANO" --> "FOOD: ["COOP"]
2. Modificare il budget mensile nel file *./taxonomies/budget.xlsx*
3. Modificare eventualmente l'anno di partenza delle statistiche (cutOffYear, default = 2023) all'interno del file *./app.py*

## Utilizzo
1. Scaricare la lista movimenti in formato CSV dal portale inbank: "Ultimi Movimenti" > "Movimenti conto (.csv)"
   - Per statistiche più complete selezionare un periodo ampio (ad es.: transazioni degli ultimi tre mesi/ultimo anno)
2. Eseguire il file *./spese.bat* 
    - Il programma importerà la transazione contenute nel file "ListaMovimentiCsv..." e le salverà nel file incrementale *./data/transactions.xlsx*
    - Genererà statistiche tabulari (excel files) e visuali (grafici) nella cartella *./outputs*
