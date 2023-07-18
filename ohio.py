import pandas as pd
import tabulate
import os
import argparse
from Tools.scripts.dutree import display

ap = argparse.ArgumentParser()
ap.add_argument('--f_path', help="File Location for Upload")
ap.add_argument('--d_path', help="Destination File Location for Upload")
ap.add_argument('--y', help = "Enter the Year")
ap.add_argument('--m', help = 'Enter the Month')
args = vars(ap.parse_args())

fileloc = args['f_path']

# Column names
columnnames = ['ID', 'CompanyName', 'Address', 'City', 'State', 'Zip', 'Gal', 'Aviation']

# Importing january
month = pd.read_excel(fileloc, names=columnnames, skiprows =5, skipfooter=3, header=None)

# Adding new column with a constant value: 2023 and 1
# Adding the year and month column
month["Year"]= args['y']
month['Month'] = args['m']
month

# Slicing rows
month1 = month
month1_df = pd.DataFrame(data=month1)
month1_df


# Rearranging the order of the columns to required format
result2 = month1_df[['ID', 'CompanyName', 'Year', 'Month', 'Gal']]
result2


# Converts dataframe to a styler and left aligns it
left_aligned_df = result2.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})
left_aligned_df

destloc = args['d_path']
# export DataFrame to text file
file_path = os.path.join(destloc, 'OH.txt')

# export DataFrame to text file
with open(file_path, 'w') as f:
    left_aligned_df_string = left_aligned_df.hide(axis="index").hide(axis=1).to_string(sparse_columns=True, sparse_index=True, delimiter='\t')
    f.write(left_aligned_df_string)
    print("Export Complete!")