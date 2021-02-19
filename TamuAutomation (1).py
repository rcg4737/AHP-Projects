import pandas as pd
import openpyxl
import os
import re
import sys
import datetime

downloadsFolder = os.path.join(os.environ['USERPROFILE'], 'Downloads') # current users downloads folder
os.chdir(downloadsFolder)

# Enter the name of the Excel file that needs to be processed
FileName  = input()


# Pandas reads in the file
if FileName.endswith('.xlsx'):
    df = pd.read_excel(FileName)
    FileName = FileName[:-5]
else:
    df = pd.read_excel(FileName + '.xlsx')



# Inserts AHP file load header if it is not present or Replace headers with AHP file load headers if headers are incorrect or Kick file back if columns are not aligned
new_colnames = ['Transaction Type',	'School',	'SSN',	'Student ID', 'Last Name','First Name',	'Middle Name',	'Address 1', 	'Address 2',	'City', 	'State', 	'Zip',	'Zip+4', 	'Email', 	'Phone', 	'Gender', 	'DOB', 	'Coverage Period', 	'Eff',	'Term', 	'Coverage Type', 	'Classification',	'Product Type']
old_colnames = df.columns
DateList = [df.iat[8, 16], df.iat[8, 18]]  # Selects the 8th row's DOB, Eff and Term date
NewDatelist = list(map(str, DateList))                    # Converts DateList into string
Str_columns = list(map(str, df.columns))
df1 = df.columns.str.contains('(?i)transaction|transaction code|transaction type') # checks if transaction is present in column header


##########   Checks if column headers are present and inserts AHP header if it is not present   ##########

if df1[0] == False:
    for dates in Str_columns:
        if re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", dates) or  re.match(r"^\d{6,8}$", dates):
            print('dates are in the header')
    new_df = pd.read_excel(FileName + '.xlsx', header= None) 
    df = new_df   
    df.rename(columns={i:j for i,j in zip(df.columns,new_colnames)}, inplace=True)


##########  If header is present: Check if DOB, Eff and Term are in the correct columns on the 8th row and kicks back file if colums are misaligned  ##########
if df1[0] == True:   
    for ValueCheck in NewDatelist:
        if not re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", ValueCheck) and not re.match(r"^\d{6,8}$", ValueCheck):
            df.to_csv('COLUMNS_ARE_NOT_ALIGNED_CORRECTLY.xlsx', index=False)
            sys.exit()    
    df.rename(columns={i:j for i,j in zip(old_colnames,new_colnames)}, inplace=True)



##########  Changes Transaction Type to U   ################# 
df.loc[df['Transaction Type'] != "U", 'Transaction Type'] = "U"



# Convert dates into the correct format in DOB, Eff and Term columns
if df["DOB"].dtype == 'int64': 
    df["DOB"] = df["DOB"].apply(str)
    df["DOB"] = pd.to_datetime(df["DOB"]).dt.strftime('%m/%d/%Y')
df["DOB"] = pd.to_datetime(df["DOB"]).dt.strftime('%m/%d/%Y')

if df["Eff"].dtype == 'int64': 
    df["Eff"] = df["Eff"].apply(str)
    df["Eff"] = pd.to_datetime(df["Eff"]).dt.strftime('%m/%d/%Y')
df["Eff"] = pd.to_datetime(df["Eff"]).dt.strftime('%m/%d/%Y')


###########     Changes Term date to 8/31/2020     ####################     - UPDATE ANNUALLY
df.loc[df['Term'] != "8/31/2021", 'Term'] = "8/31/2021"        



##########     changes the day of each date to 01 (xx/01/xxxx)  #####################
if df["Eff"].dtype == 'int64': 
    df["Eff"] = df["Eff"].apply(str)
    df["Eff"] = pd.to_datetime(df["Eff"]).dt.strftime('%m/01/%Y')
df["Eff"] = pd.to_datetime(df["Eff"]).dt.strftime('%m/01/%Y')


############    Crates dataframes based on Eff date     ###################
September = pd.DataFrame(df[df.Eff == '09/01/2020'])
October = pd.DataFrame(df[df.Eff == '10/01/2020'])
November = pd.DataFrame(df[df.Eff == '11/01/2020'])
December = pd.DataFrame(df[df.Eff == '12/01/2020'])
January = pd.DataFrame(df[df.Eff == '01/01/2021'])
Febuary = pd.DataFrame(df[df.Eff == '02/01/2021'])
March = pd.DataFrame(df[df.Eff == '03/01/2021'])
April = pd.DataFrame(df[df.Eff == '04/01/2021'])
May = pd.DataFrame(df[df.Eff == '05/01/2021'])
June = pd.DataFrame(df[df.Eff == '06/01/2021'])
July = pd.DataFrame(df[df.Eff == '07/01/2021'])
August = pd.DataFrame(df[df.Eff == '08/01/2021'])


############    Writes Eff date dataframe to a worksheets on the same document and saves document  ###############
writer = pd.ExcelWriter(FileName+"FINAL.xlsx", engine="xlsxwriter")

September.to_excel(writer, index=False, sheet_name='September')
October.to_excel(writer, index=False, sheet_name='October')
November.to_excel(writer, index=False, sheet_name='November')
December.to_excel(writer, index=False, sheet_name='December')
January.to_excel(writer, index=False, sheet_name='January')
Febuary.to_excel(writer, index=False, sheet_name='Febuary')
March.to_excel(writer, index=False, sheet_name='March')
April.to_excel(writer, index=False, sheet_name='April')
May.to_excel(writer, index=False, sheet_name='May')
June.to_excel(writer, index=False, sheet_name='June')
July.to_excel(writer, index=False, sheet_name='July')
August.to_excel(writer, index=False, sheet_name='August')

writer.save()

 


