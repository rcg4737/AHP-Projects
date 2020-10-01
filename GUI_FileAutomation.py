import pandas as pd
import os
import openpyxl
import re
import sys
import json
from tkinter import *
import tkinter.messagebox
from tkinter import ttk  # Normal Tkinter.* widgets are not themed!
from ttkthemes import ThemedTk
from threading import Thread

os.chdir('C:\\Users\\robert graham\\Downloads')
root = ThemedTk(theme="breeze")
root.title('File Formatter')
root.geometry("400x150")

def automation_threaded():
    Thread(target=automation).start()


def automation():
    try:
        FileName = myentry.get()
        dataframe = pd.read_csv(FileName + '.csv')
        dataframe.dropna(axis=0, how='all',inplace=True)  # deletes all empty rows
        # Inserts AHP file load header if it is not present or Replace headers with AHP file load headers if headers are incorrect or Kick file back if columns are not aligned
        new_colnames = ['Transaction Type',	'School',	'SSN',	'Student ID', 'Last Name','First Name',	'Middle Name',	'Address 1', 	'Address 2',	'City', 	'State', 	'Zip',	'Zip+4', 	'Email', 	'Phone', 	'Gender', 	'DOB', 	'Coverage Period', 	'Eff',	'Term', 	'Coverage Type', 	'Classification',	'Product Type']
        old_colnames = dataframe.columns
        DateList = [dataframe.iat[3, 16], dataframe.iat[3, 18], dataframe.iat[3, 19]]  # Selects the 8th row's DOB, Eff and Term date
        NewDatelist = list(map(str, DateList))                    # Converts DateList into string
        dataframe1 = dataframe.columns.str.contains('(?i)transaction|transaction code|transaction type') # checks if transaction is present in column header

        ##########   Checks if column headers are present and inserts AHP header if it is not present   ##########
        ##########  If header is present: Check if DOB, Eff and Term are in the correct columns on the 8th row and kicks back file if colums are misaligned  ##########

        if dataframe1[0] == False:
            for dates in dataframe.columns:
                if re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", dates) or  re.match(r"^\d{6,8}$", dates):
                    print("success")
            new_dataframe = pd.read_csv(FileName + '.csv', header= None) 
            dataframe = new_dataframe   
            dataframe.rename(columns={i:j for i,j in zip(dataframe.columns,new_colnames)}, inplace=True)
        elif dataframe1[0] == True:   
            for ValueCheck in NewDatelist:
                if not re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", ValueCheck) and not re.match(r"^\d{6,8}$", ValueCheck):
                    tkinter.messagebox.showerror("Column Header Error", "The columns in this file are not aligned!")
                    sys.exit()    
            dataframe.rename(columns={i:j for i,j in zip(old_colnames,new_colnames)}, inplace=True)

        # Creates dataframe without rows with DOB as a DOB (Deletes duplicate column headers placed in rows)
        dataframe = dataframe[dataframe.DOB != 'DOB' ]


        #delete all rows with more than 3 characters in Transaction Type column 
        dataframe = dataframe[dataframe['Transaction Type'].str.len() <= 3]


        # Checks file dates for the correct format in DOB, Eff and Term columns

        if dataframe["DOB"].dtype == 'int64': 
            dataframe["DOB"] = dataframe["DOB"].apply(str)
            dataframe["DOB"] = pd.to_datetime(dataframe["DOB"]).dt.strftime('%m/%d/%Y')
        dataframe["DOB"] = pd.to_datetime(dataframe["DOB"]).dt.strftime('%m/%d/%Y')

        if dataframe["Eff"].dtype == 'int64': 
            dataframe["Eff"] = dataframe["Eff"].apply(str)
            dataframe["Eff"] = pd.to_datetime(dataframe["Eff"]).dt.strftime('%m/%d/%Y')
        dataframe["Eff"] = pd.to_datetime(dataframe["Eff"]).dt.strftime('%m/%d/%Y')

        if dataframe["Term"].dtype == 'int64': 
            dataframe["Term"] = dataframe["Term"].apply(str)
            dataframe["Term"] = pd.to_datetime(dataframe["Term"]).dt.strftime('%m/%d/%Y')
        dataframe["Term"] = pd.to_datetime(dataframe["Term"]).dt.strftime('%m/%d/%Y')


        # Modify student ID based on school code in json file (add more schools in this field as needed)
        studentIDbySchoolCode = open('schoolcodes.json')
        data = json.load(studentIDbySchoolCode)

        for i in data['schools']: 
            for col in dataframe['School']:
                for scode in i['SchoolCode']:
                    if scode in col:
                        dataframe['Student ID'] = dataframe['Student ID'].apply(lambda x: ('{0:0>'+str(i['SchoolIdLength'])+'}').format(x))

        #dataframe.School = [x.upper() for x in dataframe.School] 
        #dataframe.School = dataframe.School.str.upper()

        # converts csv into xlsx in downloads folder 
        dataframe.to_excel(FileName + '.xlsx', index=False)
        # os.remove(FileName + '.csv')
    except:
        if not myentry.get( ) == "":
            tkinter.messagebox.showerror("File Error", "The file was not processed due to unknown error!  Please manually process the file.")
        else:
            tkinter.messagebox.showerror("File Name Error", "Please enter a file name.")


mylabel = ttk.Label(root, text="Enter the name of the file you want to format")
mylabel.pack(pady=10)

myentry = ttk.Entry(root, width=50 )
myentry.pack()


mybutton = ttk.Button(root, text='Format', command= automation_threaded)
mybutton.pack(pady=10)

root.mainloop()