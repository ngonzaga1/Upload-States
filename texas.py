import pandas as pd
import os
import tabulate
import xlrd
import openpyxl
import argparse
from Tools.scripts.dutree import display
from xlrd import open_workbook, XLRDError
from datetime import datetime
import math

from pandas.io.formats import excel

excel.ExcelFormatter.header_style = None

# Directory path
# dir_path = r"C:\wip\som\rick\TX\January Revised.TXT"

ap = argparse.ArgumentParser()
ap.add_argument('--f_path', help="File Location for Upload")
ap.add_argument('--d_path', help="Destination File Location for Upload")
# ap.add_argument('--e_path', help="Excel File Destional Location")
args = vars(ap.parse_args())

fileloc = args['f_path']


with open(fileloc, 'r') as f:
    lines = f.readlines()

# Defining the length of the columns0
col_lengths = (11, 50, 50, 30, 2, 5, 4, 10, 2, 2, 50, 10,
               2, 11, 10, 10, 5, 1, 1, 17, 17, 17, 17, 17,
               17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17,
               17, 17, 17, 17, 17, 17, 17, 17, 17, 17, 17,
               17, 17, 17, 17)

# Assigning the appropriate length to each column using a for loop
col_positions = []
start_pos = 0
for length in col_lengths:
    end_pos = start_pos + length
    col_positions.append((start_pos, end_pos))
    start_pos = end_pos

with open(fileloc, 'r') as f:
    data = []
    for line in f:
        row = {}
        for i, pos in enumerate(col_positions):
            start_pos, end_pos = pos
            row[f'col{i + 1}'] = line[start_pos:end_pos].strip()
        data.append(row)

# Converting the data to a dataframe
df = pd.DataFrame(data)

pd.set_option('display.max_columns', None)

# Converts dataframe to a styler and left aligns it
left_aligned_df = df.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})

# The path which will upload the data frame to the text file - ohio1.txt
# path = r'C:\wip\som\rick\TX\TX.txt'

destloc = args['d_path']

# export DataFrame to text file
file_path = os.path.join(destloc, 'TX.txt')

# export DataFrame to text file
with open(file_path, 'w') as f:
  left_aligned_df_string = left_aligned_df.hide(axis="index").hide(axis=1).to_string(sparse_columns=True, sparse_index=True, delimiter='\t')
  f.write(left_aligned_df_string)
  print("Export Complete!")


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# SENDING TO BECKY

# Cleaning data-only need data where col10 = 07 and has a value that's not 0
df1 = df[df['col10'] == '07']
df1.sort_values(by=['col1'])

df1 = df1[df1["col48"].str.contains("0000000000000.00") == False]


# Rearranging the gallonage column to have the negatives in front
def str_to_int(s):
    sign = '-' if s.endswith('-') else ''
    no_sign = s[:-1] if s.endswith('-') else s
    return int(sign + no_sign.replace('.', ''))


df1['col48'] = df['col48'].apply(str_to_int)

# Removing the two 2 digits of the gallanage numbers
df2 = df1.sort_values(by=['col2', 'col12'], ascending=[True, False])


def remove_last_two_digits(s):
    s = str(s)[:-2]
    return s


df2['col48'] = df1['col48'].apply(remove_last_two_digits)

# Reassigning column names and only extracting the columns that are needed
column_names = ['End Date', 'licnum', 'Company', 'Gallonage', 'Address', 'City', 'State', 'Zip']

df2 = df2[['col12', 'col1', 'col2', 'col48', 'col3', 'col4', 'col5', 'col6']]

df2 = df2.rename(columns={'col12': 'End Date', 'col1': 'licnum', 'col2': 'Company',
                          'col48': 'Gallonage', 'col3': 'Address', 'col4': 'City', 'col5': 'State', 'col6': 'Zip'})


# define function to format date string
def format_date(date_str):
    date_obj = datetime.strptime(date_str, '%m/%d/%Y')
    return date_obj.strftime('%d-%b-%y')


# apply function to 'date' column
df2['End Date'] = df2['End Date'].apply(format_date)

# Converts dataframe to a styler and left aligns it
left_aligned_df1 = df2.reset_index(drop=True).style.set_properties(**{'text-align': 'left'})

# The path which will upload the data frame to the text file -TX_D2_4SOM_python.txt
path1 = r'C:\wip\som\rick\TX\TX_D2_4SOM_python.xlsx'

# destloc1 = args['e_path']

# export DataFrame to text filefd
with open(path1, 'w') as f:
     left_aligned_df_string1 = left_aligned_df1.hide(axis="index").hide(axis=1).to_excel(excel_writer=path1, inf_rep=str,
                                                                         index=False)
     print("Export Complete!")
