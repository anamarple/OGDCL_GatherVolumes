import os
import pandas as pd
import xlrd
import numpy as np

root = r'V:\APLA\jobs\AP\OGDCL\024092 OGDCL 2020MY Pakistan\01 Data\00 Incoming Data\09-30-2020 FTP\2020 RES Study ' \
       r'OGDCL (SLB Data)\OGDCL-2018-Schlumberger-Study-Geological-Data\OGDCL_Excel_Sheets'

# Create a mother dataframe
mother_df = pd.DataFrame(columns=['Well_Name', 'Field_Name', 'Division', 'Form_Name', 'Company', 'Country', 'Rsvr_Lith',
                                  'Uncertainty', 'ContactDepth_ft', 'Operator', 'Concession/Block', 'Lease_Type',
                                  'LeaseExp_Date', 'Company_WI', 'GRV_m3', 'Area_m2', 'Gross_m', 'Net_m', 'NTG',
                                  'NRV_m3','po', '1-sw', 'HCPV', 'Bg', 'GIIP', 'Has_Sheet2'])

for dir, subdir, files in os.walk(root):

    # If files is not empty
    if files:
        for f in files:
            loc = dir + '\\' + f

            # Pull uncertainty from file name
            if 'P10' in f:
                unc = 'P10'
            elif 'P50' in f:
                unc = 'P50'
            elif 'P90' in f:

                unc = 'P90'
            else:
                unc = ''

            try:
                # Open file
                wb = xlrd.open_workbook(loc)
                sheets = wb.sheet_names()
                sheet = wb.sheet_by_index(0)

                # See if file has 1 or 2 sheets
                if len(sheets) > 1:
                    has_sheet2 = True
                else:
                    has_sheet2 = False

                # Read first block of info
                x = pd.read_excel(loc, skiprows=2, nrows=14, usecols='B')
                x_array = x.to_numpy()
                x_reshaped = np.reshape(x_array, (1, -1))
                x = pd.DataFrame(x_reshaped, columns=['Well_Name', 'Field_Name', 'Division', 'Form_Name', 'Company',
                                                      'Country', 'Rsvr_Lith', 'Uncertainty', 'ContactDepth_ft',
                                                      'Operator',
                                                      'Concession/Block', 'Lease_Type', 'LeaseExp_Date', 'Company_WI'])
                x['Uncertainty'] = unc

                # Read second block of info
                y = pd.read_excel(loc, skiprows=2, nrows=11, usecols='E')
                y_array = y.to_numpy()
                y_reshaped = np.reshape(y_array, (1, -1))
                y = pd.DataFrame(y_reshaped, columns=['GRV_m3', 'Area_m2', 'Gross_m', 'Net_m', 'NTG', 'NRV_m3', 'po',
                                                      '1-sw', 'HCPV', 'Bg', 'GIIP'])
                y['Has_Sheet2'] = has_sheet2

                # Append x and y dfs to each other
                xy = pd.concat([x, y], axis=1)

                # Append df to mother df
                mother_df = mother_df.append(other=xy, ignore_index=True)

            except AssertionError:
                print('Skipped: ' + f)

print(mother_df)

# Write out to an excel file
mother_df.to_excel(r'V:\APLA\jobs\AP\OGDCL\024092 OGDCL 2020MY Pakistan\01 Data\00 Incoming Data\09-30-2020 FTP\2020 '
                   r'RES Study OGDCL (SLB Data)\OGDCL-2018-Schlumberger-Study-Geological-Data\output.xlsx')
