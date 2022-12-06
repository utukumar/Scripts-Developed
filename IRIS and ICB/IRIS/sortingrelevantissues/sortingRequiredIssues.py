import calendar

import pandas as pd
import datetime
import os
from tkinter import filedialog
from tkinter.filedialog import askopenfile

# opening the Excel file required
file_name = filedialog.askopenfile(mode='r')
file_path = os.path.abspath(file_name.name)

# creating dataframe from the excel sheet
df = pd.read_excel(file_path, parse_dates=['CreateDate', 'ResolvedDate'])

# getting user input for manual or automated date entry option
user_selection = int(input(" Enter 1 for manual date entry and 2 for automated date selection: "))

if user_selection == 1:
    # getting current year, month, day
    year = int(input("Enter Year : "))
    month = int(input("Enter Month : "))
    day = int(input("Enter Day : "))
    dayNumber = int(datetime.datetime(year, month, day).strftime('%w'))  # 0 for Sunday and 1 for monday and so on

    # checking if year is leap or not
    leap_year_bool = calendar.isleap(year)

    if dayNumber == 1 and day == 1:
        if month == 1:
            month = 12
            day = int(input("Enter the previous month's last Friday's date : "))
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        else:
            month = month - 1
            day = int(input("Enter the previous month's last Friday's date : "))
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    elif dayNumber == 1:
        day = int(input("Enter the previous Friday's date : "))
        required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    elif day == 1:
        if month == 2:
            month = 1
            date = 31
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        elif month in [4, 6, 9, 11]:
            month = month - 1
            day = 31
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        elif month in [5, 7, 8, 10, 12]:
            month = month - 1
            day = 30
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        elif month == 3:
            month = 2
            if leap_year_bool:
                day = 29
            else:
                day = 28
            required_datetime = datetime.datetime(year, month, day, 11, 30, 00)
    else:
        required_datetime = datetime.datetime(year, month, day - 1, 11, 30, 00)

    required_timestamp = required_datetime.replace(tzinfo=datetime.timezone.utc).timestamp()
    print(required_timestamp)


elif user_selection == 2:
    # getting current year, month, day
    year = int(datetime.datetime.now().strftime("%Y"))
    month = int(datetime.datetime.now().strftime("%#m"))
    day = int(datetime.datetime.now().strftime("%#d"))
    dayNumber = int(datetime.datetime.now().strftime("%w")) # 0 for Sunday and 1 for monday and so on

    # checking if year is leap or not
    leap_year_bool = calendar.isleap(year)

    if dayNumber == 1 and day == 1:
        if month == 1:
            month = 12
            day = int(input("Enter the previous month's last Friday's date : "))
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        else:
            month = month - 1
            day = int(input("Enter the previous month's last Friday's date : "))
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    elif dayNumber == 1:
        day = int(input("Enter the previous Friday's date : "))
        required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
    elif day == 1:
        if month == 2:
            month = 1
            date = 31
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        elif month in [4, 6, 9, 11]:
            month = month - 1
            day = 31
            required_datetime = datetime.datetime(year, month, day, 11, 30, 0)
        elif month in [5, 7, 8, 10, 12]:
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


# python_dictionary to track the seen issues
seen_issues = {}

# required issues list to store the required issues details
required_issues = []

# open and resolved issue count
open_count = 0
resolved_count = 0

# reading the dataframe for sorting open issues
for i in range(len(df)):
    if required_timestamp - df.iloc[i, 6].timestamp() < 0:
        seen_issues[df.iloc[i, 0]] = 1
        open_count += 1
        result = {
                  'ShortId': df.iloc[i, 0],
                  'Title': df.iloc[i, 1],
                  'Priority': df.iloc[i, 2],
                  'Status': df.iloc[i, 3],
                  'RequesterIdentity': df.iloc[i, 4],
                  'AssigneeIdentity': df.iloc[i, 5],
                  'CreateDate': df.iloc[i, 6],
                  'ResolvedDate': df.iloc[i, 7],
                  'IssueUrl': df.iloc[i, 8],
                  'Tags': df.iloc[i, 9],
                  'Labels': df.iloc[i, 10]
                  }
        required_issues.append(result)

# reading dataframe for sorting resolved issues
for i in range(len(df)):
    if df.iloc[i, 0] not in seen_issues:
        if not pd.isna(df.iloc[i, 7]):
            if required_timestamp - df.iloc[i, 7].timestamp() < 0:
                resolved_count += 1
                result = {
                    'ShortId': df.iloc[i, 0],
                    'Title': df.iloc[i, 1],
                    'Priority': df.iloc[i, 2],
                    'Status': df.iloc[i, 3],
                    'RequesterIdentity': df.iloc[i, 4],
                    'AssigneeIdentity': df.iloc[i, 5],
                    'CreateDate': df.iloc[i, 6],
                    'ResolvedDate': df.iloc[i, 7],
                    'IssueUrl': df.iloc[i, 8],
                    'Tags': df.iloc[i, 9],
                    'Labels': df.iloc[i, 10]
                }
                required_issues.append(result)

print(open_count)
print(resolved_count)

# converting required_issued list to pandas dataframe
required_issues_df = pd.DataFrame(required_issues)

# data frame to excel
required_issues_df.to_excel(r"C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\indianConsumerBusinnes\iris\sortingrelevantissues\output.xlsx", index=False)

#output prompt
print("Output.xlsx file generated")


