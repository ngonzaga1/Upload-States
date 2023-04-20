import pandas as pd
import tabulate
import xlrd
import openpyxl
from Tools.scripts.dutree import display

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# INSERT FIlE LOCATION HERE
file_loc = r'C:\wip\som\rick\OH\Montly Fuel Report for Ohio Dec22-Jan22.xlsx'
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Column names
columnnames = ['ID', 'CompanyName', 'Address', 'City', 'State', 'Zip', 'Gal', 'Aviation', 'Year', 'Month']

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Importing excel
ohio = pd.read_excel(file_loc, sheet_name=['Dec 22', 'Jan 23'], header=0, names=columnnames)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Combine data from all worksheets
ohio_all = pd.concat(ohio.values(), ignore_index=True)

# Rearranging the order of the columns to required format
result1 = ohio_all[['ID', 'CompanyName', 'Year', 'Month', 'Gal']]
print(result1)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Rearranging the order of the columns to required format
result2 = result1[['ID', 'CompanyName', 'Year', 'Month', 'Gal']]

# Converts dataframe to a styler and left aligns it
left_aligned_df = result2.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})
print(left_aligned_df)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# The path which will upload the data frame to the text file - nebraska.txt
path = r'C:\wip\som\rick\OH\ohio1.txt'
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


# export DataFrame to text file
with open(path, 'a') as f:
  left_aligned_df_string = left_aligned_df.hide(axis="index").hide(axis=1).to_string(sparse_columns=True, sparse_index=True, delimiter='\t')
  f.write(left_aligned_df_string)
  print("Export Complete!")
