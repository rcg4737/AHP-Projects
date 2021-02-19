import pandas as pd
import os
import openpyxl
import re
import sys
import json


os.chdir('C:\\Users\\robert graham\\Downloads')

# Enter the name of the CSV file that needs to be processed
FileName  = input()


# Pandas reads in the file
df = pd.read_csv(FileName + '.csv')

# deletes all rows that have no data
df.dropna(axis=0, how='all',inplace=True)

# Inserts AHP file load header if it is not present or Replace headers with AHP file load headers if headers are incorrect or Kick file back if columns are not aligned
new_colnames = ['Transaction Type',	'School',	'SSN',	'Student ID', 'Last Name','First Name',	'Middle Name',	'Address 1', 	'Address 2',	'City', 	'State', 	'Zip',	'Zip+4', 	'Email', 	'Phone', 	'Gender', 	'DOB', 	'Coverage Period', 	'Eff',	'Term', 	'Coverage Type', 	'Classification',	'Product Type']
old_colnames = df.columns
DateList = [df.iat[3, 16], df.iat[3, 18], df.iat[3, 19]]  # Selects the 8th row's DOB, Eff and Term date
NewDatelist = list(map(str, DateList))                    # Converts DateList into string
df1 = df.columns.str.contains('(?i)transaction|transaction code|transaction type') # checks if transaction is present in column header

##########   Checks if column headers are present and inserts AHP header if it is not present   ##########

if df1[0] == False:
    for dates in df.columns:
        if re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", dates) or  re.match(r"^\d{6,8}$", dates):
            print('dates are in the header')
    new_df = pd.read_csv(FileName + '.csv', header= None) 
    df = new_df   
    df.rename(columns={i:j for i,j in zip(df.columns,new_colnames)}, inplace=True)



##########  If header is present: Check if DOB, Eff and Term are in the correct columns on the 8th row and kicks back file if colums are misaligned  ##########

if df1[0] == True:   
    for ValueCheck in NewDatelist:
        if not re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", ValueCheck) and not re.match(r"^\d{6,8}$", ValueCheck):
            df.to_csv('COLUMNS_ARE_NOT_ALIGNED_CORRECTLY.csv', index=False)
            sys.exit()    
    df.rename(columns={i:j for i,j in zip(old_colnames,new_colnames)}, inplace=True)
    


# Creates dataframe without rows with DOB as a DOB (Deletes duplicate column headers placed in rows)
df = df[df.DOB != 'DOB' ]


#delete all rows with more than 3 characters in Transaction Type column 
df = df[df['Transaction Type'].str.len() <= 3]




# Convert dates into the correct format in DOB, Eff and Term columns

if df["DOB"].dtype == 'int64': 
    df["DOB"] = df["DOB"].apply(str)
    df["DOB"] = pd.to_datetime(df["DOB"]).dt.strftime('%m/%d/%Y')
df["DOB"] = pd.to_datetime(df["DOB"]).dt.strftime('%m/%d/%Y')

if df["Eff"].dtype == 'int64': 
    df["Eff"] = df["Eff"].apply(str)
    df["Eff"] = pd.to_datetime(df["Eff"]).dt.strftime('%m/%d/%Y')
df["Eff"] = pd.to_datetime(df["Eff"]).dt.strftime('%m/%d/%Y')

if df["Term"].dtype == 'int64': 
    df["Term"] = df["Term"].apply(str)
    df["Term"] = pd.to_datetime(df["Term"]).dt.strftime('%m/%d/%Y')
df["Term"] = pd.to_datetime(df["Term"]).dt.strftime('%m/%d/%Y')



# Modify student ID based on school code (add more schools in this field as needed)
f = open('C:\\Users\\robert graham\\OneDrive - ACADEMIC HEALTHPLANS\\Desktop\\python\\schoolcodes.json')
data = json.load(f)

for i in data['schools']:
    for col in df['School']:
        for scode in i['SchoolCode']:
            if scode in col:
                df['Student ID'] = df['Student ID'].apply(lambda x: ('{0:0>'+str(i['SchoolIdLength'])+'}').format(x))


# converts  csv into xlsx
#os.chdir('C:\\Users\\robert graham\\Downloads')      # change file path to where you want xlsx file. 
df.to_excel(FileName + '.xlsx', index=False)
os.remove(FileName + '.csv')