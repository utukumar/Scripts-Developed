import pandas as pd
import datetime
import calendar
import os
import openpyxl
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfile


#creating pop-up for selecting the file
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfile(mode = 'r')

#assigning assigning path to path vatiable
path = os.path.abspath(file_path.name)
df = pd.read_excel(path, parse_dates=["Created", "Resolved"])

#getting current day, month, year
year = int(datetime.datetime.now().strftime("%Y"))
month = int(datetime.datetime.now().strftime("%#m"))
day = int(datetime.datetime.now().strftime("%#d"))

#getting the dayNumber 0 for Sunday, 1 for Monday and so on
dayNum = int(datetime.datetime.now().strftime("%w"))

#gettign the leapyear boolean
leapYearBool = calendar.isleap(year)

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
        if leapYearBool:
            day = 29
        else:
            day = 28
        required_datetime = datetime.datetime(year, month, day, 11, 30, 00)
else:
    required_datetime = datetime.datetime(year, month, day-1, 11, 30, 00)
    

#converting datetime object to timestamp
required_timestamp = required_datetime.replace(tzinfo=datetime.timezone.utc).timestamp()

#list for storing issues that are required
required_issues = []
issue_key_dict = {}
open_issues = 0

#getting the issues that were created after 5PM yesterday
for i in range(len(df)):
    if required_timestamp - df.iloc[i,7].timestamp() < 0:
        open_issues += 1
        issue_key_dict[df.iloc[i,1]] = 1
        result_dict = {
                    "Issue Type":df.iloc[i,3],
                    "Issue key":df.iloc[i,1],
                    "Summary":df.iloc[i,0],
                    "Status":df.iloc[i,4],
                    "Priority":df.iloc[i,11],
                    "Assignee":df.iloc[i,13],
                    "Reporter":df.iloc[i,14],
                    "Created":df.iloc[i,16],
                    "Resolved":df.iloc[i,19],
                    "Resolution":df.iloc[i,12],
                    "Labels":f"{df.iloc[i,25]},{df.iloc[i,26]},{df.iloc[i,27]},{df.iloc[i,28]},{df.iloc[i,29]},{df.iloc[i,30]},{df.iloc[i,31]},{df.iloc[i,32]}".rstrip(','),
                    "Bug Found in Origin":df.iloc[i,75]
                   
    }
        required_issues.append(result_dict)

resolved_count = 0

for i in range(len(df)):
    if df.iloc[i,1] not in issue_key_dict:
        if pd.isna(df.iloc[i,8] == False):
            if required_timestamp - df.iloc[i,8].timestamp() < 0:
                resolved_count += 1
                result_dict = {
                    "Issue Type":df.iloc[i,3],
                    "Issue key":df.iloc[i,1],
                    "Summary":df.iloc[i,0],
                    "Status":df.iloc[i,4],
                    "Priority":df.iloc[i,11],
                    "Assignee":df.iloc[i,13],
                    "Reporter":df.iloc[i,14],
                    "Created":df.iloc[i,16],
                    "Resolved":df.iloc[i,19],
                    "Resolution":df.iloc[i,12],
                    "Labels":f"{df.iloc[i,25]},{df.iloc[i,26]},{df.iloc[i,27]},{df.iloc[i,28]},{df.iloc[i,29]},{df.iloc[i,30]},{df.iloc[i,31]},{df.iloc[i,32]}".rstrip(','),
                    "Bug Found in Origin":df.iloc[i,75]
                   
    }
                required_issues.append(result_dict)
                
print("Open Count : ",open_issues)
print()
print("Resolved Count : ",resolved_count)
print()

required_issues_df = pd.DataFrame(required_issues)
required_issues_df.to_excel(r"C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\DEM-Team\Sorting relevant issues based on date\relevantIssues.xlsx", index=False, engine='openpyxl')


        


