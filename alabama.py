import pandas as pd
import tabulate
import argparse
import os
from Tools.scripts.dutree import display
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

ap = argparse.ArgumentParser()
ap.add_argument('--f_path', help="File Location for Upload")
ap.add_argument('--d_path', help="Destination File Location for Upload")
ap.add_argument('--y', help = "Enter the Year")
ap.add_argument('--m', help = 'Enter the Month')
args = vars(ap.parse_args())

fileloc = args['f_path']

# Column Names
columnnames = ['Company', 'City', '11B', 'Gallons', 'Gallons.1']


month = pd.read_excel(fileloc, header=15, skipfooter=7, index_col =None, sheet_name='Sheet1')
month_df = pd.DataFrame(data=month)

# Extracting the columns needed
month1 = month_df.iloc[:,[0, 5,8, 9,10]]

# Renaming the columns
mapping = {month1.columns[0]: 'Company', month1.columns[1]: 'City', month1.columns[2]: '11B',
           month1.columns[3]: 'Gallons', month1.columns[4]: 'Gallons.1'}
month2 = month1.rename(columns=mapping)

# Importing Company List
columnnames = ['Licnum', 'Company', 'Cityy', 'Cityy2']
comp_list = pd.read_excel(r'C:\wip\som\rick\AL\Copy of z_AL Company List.xls', header=None, names = columnnames, sheet_name = 'Licnum 4', skiprows=4)
complist = comp_list.sort_values('Company')

# Function to find the closest match for a given company and city based on a similarity threshold
def find_closest_match(row):
    company = row['Company']
    city = row['City']
    matches = complist[(complist['Company'] == company) & (complist['Cityy2'] == city)]
    if len(matches) > 10:
        return matches.iloc[0]['Company']
    else:
        match = process.extractOne(company, complist['Company'])
        if match[1] >= 84:  # Set a threshold for the similarity score (adjust as needed)
            return match[0]
        else:
            return None

# Apply the find_closest_match function to the "Company" and "City" columns
month2['Company'] = month2.apply(find_closest_match, axis=1)


# Adding the Gallons and Gallons.1 Columns
month2['Gal'] = month2[['Gallons', 'Gallons.1']].sum(axis=1)

# Adding the year and month column
month2["Year"]= args['y']
month2['Month'] = args['m']

# Merging the month excel sheet and the company list by the company name
merged = pd.merge(month2, complist, how='left', on='Company')
merged.sort_values('Company')

# Filling NA values with 0
merged = merged.fillna(value=0, axis=1)

# Converting numeric columns to type int
merged.Gallons = merged.Gallons.astype('int')
merged['Gallons.1'] = merged['Gallons.1'].astype('int')
merged['Gal'] = merged['Gal'].astype('int')
merged['Licnum'] = merged['Licnum'].astype('int')
merged

#Choosing the required columns needed for text file
merged2= merged[['Licnum','Company','Year', 'Month', 'City', 'Gal']]
merged2.sort_values('Company')

# Convert DataFrame to a string
merged2.replace(0, '')
merged3=merged2.replace(0, '')
merged3

# Converts dataframe to a styler and left aligns it
left_aligned_df = merged3.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})
left_aligned_df

destloc = args['d_path']

file_path = os.path.join(destloc, 'AL.txt')

# export DataFrame to text file
with open(file_path, 'w') as f:
    left_aligned_df_string = left_aligned_df.hide(axis="index").hide(axis=1).to_string(sparse_columns=True, sparse_index=True, delimiter='\t')
    f.write(left_aligned_df_string)
    print("Export Complete!")