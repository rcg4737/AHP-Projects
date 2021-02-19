import pandas as pd     
import os
import openpyxl
import re
import sys
import json
import tkinter
from tkinter import messagebox
from tkinter import ttk  
from ttkthemes import ThemedTk
from threading import Thread



downloadsFolder = os.path.join(os.environ['USERPROFILE'], 'Downloads') # current users downloads folder
os.chdir(downloadsFolder)

root = ThemedTk(theme="breeze")
root.title('File Formatter')
root.geometry("400x150")

# function to find any file in users machine
def find(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)


def automation_threaded():
    Thread(target=automation).start()


def automation():
    
    #  Load file into dataframe based on filename (Pulls csv file from current users downloads folder)
    if not myentry.get( ) == "":
        try:
            FileName = myentry.get()
            if FileName.endswith('.csv'):
                dataframe = pd.read_csv(FileName)
                FileName = FileName[:-4]
            else:
                dataframe = pd.read_csv(FileName + '.csv')
        except:
            tkinter.messagebox.showerror("File Format Error", "Please use the CSV filetype.")
            myentry.delete(0, 'end')
            sys.exit()
    else:
        tkinter.messagebox.showerror("File Name Error", "Please enter a file name.")
        sys.exit()

    # deletes all empty rows
    dataframe.dropna(axis=0, how='all',inplace=True)  


    # Inserts AHP file load header if it is not present | Replace headers with AHP file load headers if headers are incorrect | The file is kicked back if columns are not aligned.
    new_colnames = ['Transaction Type',	'School',	'SSN',	'Student ID', 'Last Name','First Name',	'Middle Name',	'Address 1', 	'Address 2',	'City', 	'State', 	'Zip',	'Zip+4', 	'Email', 	'Phone', 	'Gender', 	'DOB', 	'Coverage Period', 	'Eff',	'Term', 	'Coverage Type', 	'Classification',	'Product Type']
    old_colnames = dataframe.columns
    DateList = [dataframe.iat[3, 16], dataframe.iat[3, 18], dataframe.iat[3, 19]]  # Selects the 3rd row of columns Q, S and T (DOB, Eff and Term date)
    NewDatelist = list(map(str, DateList))                    # Converts DateList into string
    dataframe1 = dataframe.columns.str.contains('(?i)transaction|transaction code|transaction type') # checks if "transaction" is present in column A header



    ##########   Checks if column headers are present and inserts AHP header if it is not present   ##########
    ##########  If header is present: Check if DOB, Eff and Term are in the correct columns on the 8th row and kicks back file if columns are misaligned  ##########
    if dataframe1[0] == False:
        for dates in dataframe.columns:
            if re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", dates) or  re.match(r"^\d{6,8}$", dates):  # Looks for dates in column headers
                continue
        new_dataframe = pd.read_csv(FileName + '.csv', header= None) 
        dataframe = new_dataframe   
        dataframe.rename(columns={i:j for i,j in zip(dataframe.columns,new_colnames)}, inplace=True)
    elif dataframe1[0] == True:   
        for ValueCheck in NewDatelist:
            if not re.match(r"^\d{1,2}\/\d{1,2}\/\d{4}$", ValueCheck) and not re.match(r"^\d{6,8}$", ValueCheck):
                tkinter.messagebox.showerror("Column Data Error", "The columns in " + FileName + " are not aligned or data is missing!")
                myentry.delete(0, 'end')
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



    #Check for schoolcode.json file on the users machine | error message if schoolcode file not found
    schoolCodesFilePath = find('schoolcodes.json', os.environ['USERPROFILE'])
    if not schoolCodesFilePath == "": 
        studentIDbySchoolCode = open(schoolCodesFilePath)
        data = json.load(studentIDbySchoolCode)
    else: 
        tkinter.messagebox.showerror("Missing School Code File", "The schoolcodes.json file was not found on your machine.")
        myentry.delete(0,'end')
        sys.exit()


    # Modify student ID based on school code in json file (add more schools in this field as needed)
    # Does nothing if school is not found
    for i in data['schools']: 
        for col in dataframe['School']:
            for scode in i['SchoolCode']:
                if scode in col:
                    dataframe['Student ID'] = dataframe['Student ID'].apply(lambda x: ('{0:0>'+str(i['SchoolIdLength'])+'}').format(x))


    # converts csv into xlsx in downloads folder 
    dataframe.to_excel(FileName + '.xlsx', index=False)
    
    # Optional: replaces the csv in downloads folder 
    #dataframe.to_csv(FileName + '.csv', index=False)

    # Optional: Removes the original CSV file
    # os.remove(FileName + '.csv') 

    # Clear the entry field in the GUI
    myentry.delete(0,'end') 
    


label1 = ttk.Label(root, text="Enter the name of the CSV file you want to format.")
label1.pack(pady=10)

myentry = ttk.Entry(root, width=50 )
myentry.pack()


mybutton = ttk.Button(root, text='Format', command= automation_threaded)
mybutton.pack(pady=10)

root.mainloop()