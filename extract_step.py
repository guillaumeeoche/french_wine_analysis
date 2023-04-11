## LIBRAIRIES 

import pandas as pd 
import os 
import xlrd
import openpyxl
import re

## DATASETS 
# Every document doesn't have the same columns, so we have to cut all the datasets into multiple parts. 
# We don't take 2009 & 2010 because the appelation VDQS (vin délimité de qualité supérieure) is not supported after these years. 

# Our analysis is based on the years between 2011 and 2022.

# YEAR 2011

colnames = ['department', 'declarations_number', 'total_area', 'aop_area', 'cognac_area', 'other_area', 'aop_white_quantity', 
            'aop_red_rose_quantity', 'cognac_quantity', 'igp_white_quantity', 'igp_red_rose_quantity', 'other_white_quantity', 'other_red_rose_quantity', 'total_white_quantity', 'total_red_rose_quantity', 
            'total_quantity']

wine_harvest_2011 = pd.read_excel('data/harvest_by_year/2011-stats-recolte.xls', 
                                  skiprows=20)

wine_harvest_2011.columns = colnames
wine_harvest_2011 = wine_harvest_2011.loc[(wine_harvest_2011['department'].isna() == False) &
                            (wine_harvest_2011['department'] != 'TOTAUX'), ]

wine_harvest_2011.to_excel("wine_harvest_2011.xlsx", index=False)

## YEAR 2012 & 2013 
# Since 2014, there is the VSI. It's a reserve system used by the winemakers so as not to lose wine.

# folder path
dir_path = r'data/harvest_by_year/'

# list to store files
filenames = []

# Iterate directory
for path in os.listdir(dir_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(dir_path, path)):
        if '2012' in path or '2013' in path:
            filenames.append(path)

for filename in filenames: 
    wine_harvest = pd.read_excel('data/harvest_by_year/' + filename, 
                                 skiprows=20)
    
    wine_harvest.columns = ['department', 'declarations_number', 'total_area', 'aop_area', 'cognac_area', 'igp_area', 'vsig_area', 
            'aop_white_quantity', 'aop_red_quantity', 'aop_rose_quantity', 'igp_white_quantity', 'igp_red_quantity', 'igp_rose_quantity', 'vsig_white_quantity', 
            'vsig_red_quantity', 'vsig_rose_quantity', 'total_white_quantity', 'total_red_quantity', 'total_rose_quantity', 'total_cognac_quantity', 'total_non_marketable_quantity', 'total_quantity']

    wine_harvest = wine_harvest.loc[(wine_harvest['department'].isna() == False) &
                            (wine_harvest['department'] != 'TOTAUX') & ('*' in wine_harvest['department']), ]
    
    print(wine_harvest)

    wine_harvest.to_excel('data/harvest_by_year/harvest_by_year_clean/' + filename.split('.')[0] + '.xlsx', index=False)