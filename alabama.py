import pandas as pd
import tabulate
import xlrd
import openpyxl
from Tools.scripts.dutree import display

# Column Names
columnnames = ['Company', 'City', '11B', 'Gallons', 'Gallons.1']

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# Importing Month
# Insert the excel file location in the read_excel function, in between the quotation marks
#   - must end with .xls or .xlsx
month = pd.read_excel(r'C:\wip\som\rick\AL\Gasoline January 1-31 2023.xls', header=2, names = columnnames, sheet_name='Sheet2')
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

month_df = pd.DataFrame(data=month)

# Slicing the data frame and rearranging the columns
month_df1 = month_df.iloc[1:66]
month1 = month_df1[['Company', 'City','Gallons', 'Gallons.1']]
month1

# New column names
columnnames = ['Licnum', 'Company', 'Cityy']
# Importing Company List
complist = pd.read_excel(r'C:\wip\som\rick\AL\Copy of z_AL Company List.xls', header=None, names = columnnames, sheet_name = 'Licnum 4')
complist

# Merging the month excel sheet and the company list by the company name
merged = pd.merge(month1, complist, how = 'left', on = 'Company')
merged

# Filling NA values with 0
merged = merged.fillna(value=0, axis=1)

# Converting numeric columns to type int
merged.Gallons = merged.Gallons.astype('int')
merged['Gallons.1'] = merged['Gallons.1'].astype('int')

merged1= merged[['Licnum','Company','City', 'Gallons', 'Gallons.1']]
merged['Licnum'] = merged['Licnum'].astype('int')
pd.DataFrame(data=merged1)

merged1['Gal'] = merged1.loc[:,['Gallons','Gallons.1']].sum(axis=1)

merged1["Year"] = 2023
merged1['Month'] = 1

merged2= merged1[['Licnum','Company','Year', 'Month', 'City', 'Gal']]

# Converts dataframe to a styler and left aligns it
left_aligned_df = merged2.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})
left_aligned_df

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# The path which will upload the data frame to the text file

# Create a text file name - for example: Alabama0423.txt
# Insert the path location of where you want the exported file to appear,
#       and then add your text file name at the end
path = r'C:\wip\som\rick\AL\Alabama0423.txt'
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# export DataFrame to text file
with open(path, 'a') as f:
  left_aligned_df_string = left_aligned_df.hide(axis="index").hide(axis=1).to_string(sparse_columns=True, sparse_index=True, delimiter='\t')
  f.write(left_aligned_df_string)
  print("Export Complete!")