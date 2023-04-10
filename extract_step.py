## LIBRAIRIES 

import pandas as pd 
import os 
import xlrd
import openpyxl


## DATASETS 

colnames = ['department', 'declarations_number', 'total_vineyard_area', 'aoc_vineyard_area', 'vdqs_vineyard_area', 'cognac_arma_vineyard_area', 'other_vineyard_area', 
            'aoc_white_quantity', 'aoc_red_rose_quantity', 'vdqs_white_quantity', 'vdqs_red_rose_quantity', 'cognac_arma_quantity', 'country_white_quantity', 'country_red_rose_quantity', 'other_white_quantity', 
            'other_red_rose_quantity']

wine_harvest_2009 = pd.read_excel('data/harvest_by_year/2009-stats-recolte.xls', 
                                  skiprows=20, 
                                  usecols='A:P')

wine_harvest_2009.columns = colnames
wine_harvest_2009 = wine_harvest_2009.loc[(wine_harvest_2009['department'].isna() == False) &
                            (wine_harvest_2009['department'] != 'TOTAUX'), ]

wine_harvest_2009.to_excel("wine_harvest_2009.xlsx", index=False)

## All the files
import os 
# folder path
dir_path = r'data/harvest_by_year/'

# list to store files
filenames = []

# Iterate directory
for path in os.listdir(dir_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(dir_path, path)):
        filenames.append(path)


for filename in filenames: 
    wine_harvest = pd.read_excel('data/harvest_by_year/' + filename, 
                                 skiprows=20, 
                                 usecols='A:P')
    
    wine_harvest.columns = colnames 


    wine_harvest = wine_harvest.loc[(wine_harvest['department'].isna() == False) &
                            (wine_harvest['department'] != 'TOTAUX'), ]
    
    print(wine_harvest)

    wine_harvest.to_excel('data/harvest_by_year/harvest_by_year_clean/' + filename.split('.')[0] + '.xlsx', index=False)