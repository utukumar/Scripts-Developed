import pandas as pd
import datetime
import calendar
import os
import openpyxl
from pathlib import Path
from tkinter import filedialog
from tkinter.filedialog import askopenfile

import tkinter as tk

#path = r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\sorting only required issues\documentSearch_utukumar (8).xls'
#creating pop-up open window for selecting the downloaded excel sheet
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfile(mode = 'r')
#assigning the absolute path to the path variable
path = os.path.abspath(file_path.name)
df = pd.read_excel(path, parse_dates=['CreateDate', 'ResolvedDate'])

#required_datetime = datetime.datetime(2022, 9, 8, 11, 30, 0)
# #getting current year, month, day
year = int(datetime.datetime.now().strftime("%Y"))
month = int(datetime.datetime.now().strftime("%#m"))
day = int(datetime.datetime.now().strftime("%#d"))
dayNum = int(datetime.datetime.now().strftime("%w"))#0 is Sunday and 6 is Saturday
leap_year_bool = calendar.isleap(year)
#required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
#print(required_datetime)
if dayNum == 1 and day == 1:
    if month == 1:
        month = 12
        day = int(input("Enter the previous month's last Friday's date : "))
        required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    else:
        month = month - 1
        day = int(input("Enter the previous month's last Friday's date : "))
        required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
elif dayNum == 1:
    day = int(input("Enter the previous Friday's date : "))
    required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
elif day == 1:
    if month == 2:
        month = 1
        date = 31
        required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    elif month in [4,6,9,11]:
        month = month - 1
        day = 31
        required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    elif month in [5,7,8,10,12]:
        month = month - 1
        day = 30
        required_datetime = datetime.datetime(year, month, day, 11, 30 , 0)
    elif month == 3:
        month = 2
        if leap_year_bool:
            day = 29
        else:
            day = 28
        required_datetime = datetime.datetime(year, month, day, 11, 30, 00)
else:
    required_datetime = datetime.datetime(year, month, day-1, 11, 30, 00)
required_timestamp = required_datetime.replace(tzinfo=datetime.timezone.utc).timestamp()
print(required_timestamp)
#list for storing  the issues that are required
required_issues = []
short_id_dict = {}
open_count = 0
#Getting the issues that were created after 5PM the previous day
for i in range(len(df)):
    if required_timestamp -  df.iloc[i,6].timestamp() < 0:
        open_count +=1
        short_id_dict[df.iloc[i,0]] = 1
        result = {'ShortId':df.iloc[i,0],
                  'Title':df.iloc[i,1],
                  'Priority':df.iloc[i,2],
                  'Status':df.iloc[i,3],
                  'RequesterIdentity':df.iloc[i,4],
                  'AssigneeIdentity':df.iloc[i,5],
                  'CreateDate':df.iloc[i,6],
                  'ResolvedDate':df.iloc[i,7],
                  'IssueUrl':df.iloc[i,8],
                  'Tags':df.iloc[i,9],
                  'Labels':df.iloc[i,10]
            
                    }
        required_issues.append(result)
r_count = 0
for i in range(len(df)):
    if df.iloc[i,0] not in short_id_dict:
        if pd.isna(df.iloc[i,7]) == False:
            if required_timestamp -  df.iloc[i,7].timestamp() < 0:
                r_count += 1
                result = {'ShortId':df.iloc[i,0],
                'Title':df.iloc[i,1],
                'Priority':df.iloc[i,2],
                'Status':df.iloc[i,3],
                'RequesterIdentity':df.iloc[i,4],
                'AssigneeIdentity':df.iloc[i,5],
                'CreateDate':df.iloc[i,6],
                'ResolvedDate':df.iloc[i,7],
                'IssueUrl':df.iloc[i,8],
                'Tags':df.iloc[i,9],
                'Labels':df.iloc[i,10]
            
                    }
                required_issues.append(result)
#creating a data frame with the issues that are required
print(open_count)
print()
print(short_id_dict)
print(r_count)
required_issues_df = pd.DataFrame(required_issues)
required_issues_df.to_excel(r"C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\sorting only required issues\output.xlsx", index=False, engine='openpyxl' )