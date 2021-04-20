import os
import pandas as pd
import xlrd
import dateparser

root = r'V:\APLA\jobs\AP\OGDCL\[024092 OGDCL 2020MY Pakistan\01 Data\01 Organized\03 Production\Non Operated JV' \
       r' Prod July -June 2017 to June 2020 Monthwise'

# Create a mother dataframe
df = pd.DataFrame()

# Iterate through each file in each folder
for dir, subdir, files in os.walk(root):

    # If files is not empty
    if files:
        for f in files:
            loc = dir + '\\' + f

            # Skip first 2 rows in excel, read into df
            x = pd.read_excel(loc, skiprows=2)
            x = x.drop(columns=['TOTAL.1'])
            # Rename columns
            x = x.rename(columns={'TOTAL': 'Total_Oil', 'OGDCL SHARE': 'OGDCL_Share_Oil', 'avg': 'avg_Oil',
                                  'TOTAL.2': 'Total_Gas', 'OGDCL SHARE.1': 'OGDCL_Share_Gas', 'avg.1': 'avg_Gas',
                                  'TOTAL.3': 'Total_LPG', 'OGDCL SHARE.2': 'OGDCL_Share_LPG', 'avg.2': 'avg_LPG'})

            # Detect date
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(0)
            date = sheet.cell_value(0, 8)
            date_ = dateparser.parse(date)

            try:
                date = date_.strftime("%m/%Y")

            except AttributeError:
                date = 'date not found'

            # Add date column to df
            num_rows = x.shape[0]
            date = [date] * num_rows
            x['Date'] = date

            file_name = [f] * num_rows
            x['File Name'] = file_name

            # Append df to mother df
            df = df.append(other=x, ignore_index=True)
            # print(x)

# print(df)

# Write out to an excel file
df.to_excel(root + '\\output.xlsx')
