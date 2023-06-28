import pandas as pd
import numpy as np
import tabulate
import xlrd
import openpyxl
from Tools.scripts.dutree import display
import argparse

ap = argparse.ArgumentParser()
ap.add_argument('--f_path', help="File Location for Upload")
ap.add_argument('--d_path', help="Destination File Location for Upload")
args = vars(ap.parse_args())

# Importing excel sheet
fileloc = args['f_path']
kansas = pd.read_excel(fileloc, skiprows = 4)
ks = pd.DataFrame(data=kansas)
ks1 = ks

#Creating column mapping
columnnames = ['county', 'licnum', 'CompanyName', 'State', 'id', 'Gal']
mapping = {ks1.columns[0]: 'county', ks1.columns[1]: 'licnum', ks1.columns[2]: 'CompanyName',
           ks1.columns[3]: 'City', ks1.columns[4]: 'State', ks1.columns[5]: 'id', ks1.columns[6]: 'Taxable glns'}
ks2 = ks1.rename(columns=mapping)

# Filling NA's
ks2[ks2['county']==""] = np.NaN
ks2['county'] =  ks2['county'].fillna(method='ffill')
# Dropping NA's
ks2.dropna(subset=['licnum','CompanyName', 'City', 'State'], how='all', inplace=True)


ks3 =ks2[['county', 'licnum', 'CompanyName', 'State', 'id', 'Taxable glns']]

# Converts dataframe to a styler and left aligns it
left_aligned_df = ks3.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})
left_aligned_df

# The path which will upload the data frame to the text file - ohio1.txt
path = r'C:\wip\som\rick\KS\kansas.txt'

# export DataFrame to text file
with open(path, 'a') as f:
    left_aligned_df_string = left_aligned_df.hide(axis="index").hide(axis=1).to_string(sparse_columns=True, sparse_index=True, delimiter='\t')
    f.write(left_aligned_df_string)
    print("Export Complete!")
