import pandas as pd
import tabulate
import xlrd
import openpyxl
import argparse
from Tools.scripts.dutree import display
from xlrd import open_workbook, XLRDError

def file_path(string):
    if os.path.isfile(string):
        return string
    else:
        raise FileNotFoundError(string)

def dest_path(string):
    if os.path.isfile(string):
        return string
    else:
        raise FileNotFoundError(string)

ap = argparse.ArgumentParser()
ap.add_argument('--f_path', type=file_path, help="File Location for Upload")
ap.add_argument('--d_path', type=dest_path, help="Destination File Location for Upload")
args = vars(ap.parse_args())

fileloc = args['f_path']
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# INSERT FIlE LOCATION HERE
file_loc = fileloc
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Column names
columnnames = ['Licnum', 'CompanyName', 'City', 'State', 'Year', 'Gal', 'Month']


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Importing excel - CHANGE SHEET NAMES IF NEEDED OPTIONAL
nebraska = pd.read_excel(file_loc, sheet_name=['Jan 22', 'Feb 22','Mar 22','Apr 22','May 22','Jun 22',
                                'Jul 22','Aug 22','Sep 22','Oct 22','Nov 22','Dec 22'], header=None, names=columnnames)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# Combine data from all worksheets
nebraska_all = pd.concat(nebraska.values(), ignore_index=True)
# print(df_all)

# Dropping the last two rows of each sheet that has 'miscellaneous'
nebraska_all1 = nebraska_all.dropna(subset=['Licnum', 'CompanyName'], how='all')
# print(nebraska_all1)

# Filling NA values with 0
nebraska_all2 = nebraska_all1.fillna(value=0, axis=1)

# Converting numeric columns to type int
nebraska_all2.Gal = nebraska_all2.Gal.astype('int')
nebraska_all2.Year = nebraska_all2.Year.astype('int')
nebraska_all2.Month = nebraska_all2.Month.astype('int')

# Rearranging the order of the columns to required format
result1 = nebraska_all2[['Licnum', 'CompanyName', 'State', 'Year', 'Gal', 'Month', 'City']]

# Converts dataframe to a styler and left aligns it
left_aligned_df = result1.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})
# print(left_aligned_df)


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# INSERT THE DESTINATION PATH HERE
#       The path which will upload the data frame to the text file - nebraska.txt
fileloc1 = args['d_path']
path = fileloc1
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# export DataFrame to text file
with open(path, 'a') as f:
    left_aligned_df_string = left_aligned_df.hide(axis="index").to_string(sparse_index=True, delimiter='\t')
    f.write(left_aligned_df_string)
    print("Export Complete!")

