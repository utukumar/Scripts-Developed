from datetime import datetime
import os
import pandas as pd
import os
from pathlib import Path
from openpyxl import workbook, load_workbook
import tkinter as tk
import datetime
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import smtplib as smt
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#creating pop-up open window for selecting the downloaded excel sheet
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfile(mode = 'r')

#assigning the absolute path to the path variable
path = os.path.abspath(file_path.name)

#reading excel file as pandas data frame
df = pd.read_excel(path, na_filter=False)

#labels used for mercury team
gdq_label = 'GDQ_QS_Detected_AMP'
team_label = 'QA-Offshore-Team'
adhoc_label = 'QS_Adhoc'
testcase_label = 'QS_Testcase'
valid_label = 'QS_Detected_Valid'
invalid_label = 'QS_Detected_Invalid'
new_feature = 'QS_New_Feature'
by_design = 'By Design'
test_case_update_needed = 'Testcase update needed'
test_environment_issue = 'Test Environment Issue'

#variable required
total_issues = 0
open_issues = 0
resolved_issues = 0

#total issue
total_issues = len(df)
#calcualting total open and resolved issues
for i in range(len(df)):
    if 'Resolved' not in df.iloc[i,4]:
        open_issues += 1
    elif 'Resolved' in df.iloc[i,4]:
        resolved_issues += 1
print("Total Open Issues : ",open_issues)
print("Total Resolved Issues : ",resolved_issues)

#Missing information list
missing_info = []
missing_info_dict = {}

#proper issue list
no_missing_info = []
######################################
#having test case yes/no/empty -- seen for both QS_Adhoc and QS_Teastcase
#using haing testcase? Yes for QS_Testcase and No for QS_Adhoc
######################################
#qs valid and invalid should not be used for open issues
#for QS_Detected_Valid Invalid Bug - Category shoud be empty
#for QS_Detected_Invalid Invalid Bug - Category should have either By_Design, Test_Environment_Issue, Test_Case_UpdateNeeded
#status is either resolved or others
#TODO:figure out what "n/a" is priority is not applicable or not available
#Case -1-- team_label, gdq_label, adhoc_label and test_case label are missing
######################################
#checking if proper lables are present or not
for i in range(len(df)):
    if gdq_label not in df.iloc[i,-1]:
        if team_label not in df.iloc[i,-1]:
            if adhoc_label not in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                #gdq label, team label, adhoc label and testcase label are not present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label}, {testcase_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label}, {testcase_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label}, {team_label}, {testcase_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 116", df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label}, {adhoc_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label}, {adhoc_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label}, {team_label}, {adhoc_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 156",df.iloc[i,0])
                    missing_info.append(result_dict)
            elif adhoc_label in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                #gdq label, team label not present
                #adhoc label is present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing and {adhoc_label} present, {testcase_label} required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing and {adhoc_label} present, {testcase_label} required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label}, {team_label} are missing and {adhoc_label} present, {testcase_label} required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 198",df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label}, {team_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 237",df.iloc[i,0])
                    missing_info.append(result_dict)
            elif adhoc_label not in df.iloc[i,-1] and testcase_label in df.iloc[i,-1]:
                #gdq label, team label not present
                #test case label is present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label}, {team_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 279",df.iloc[i,0])
                    missing_info.append(result_dict)
                
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing and TC present Adhoc required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {team_label} are missing and TC present Adhoc required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label}, {team_label} are missing and TC present Adhoc required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 319",df.iloc[i,0])
                    missing_info.append(result_dict)
        
        elif team_label in df.iloc[i,-1]:
            #gdq label not present, team label is present
            if adhoc_label not in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                #adhoc label and testcase label are not present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {testcase_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {testcase_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label} and {testcase_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 363",df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {adhoc_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label}, {adhoc_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label} and {adhoc_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 402",df.iloc[i,0])
                    missing_info.append(result_dict)
            
            elif adhoc_label in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                #gdq label not present, team label present, adhoc label present
                #test case label not present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue, {gdq_label} label is missing and {adhoc_label} present {testcase_label} required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue, {gdq_label} label is missing and {adhoc_label} present {testcase_label} required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label} is missing and {adhoc_label} present {testcase_label} required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 445",df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label} label is missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label} label is missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label} label is missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 484",df.iloc[i,0])
                    missing_info.append(result_dict)
            
            elif adhoc_label not in df.iloc[i,-1] and testcase_label in df.iloc[i,-1]:
                #gdq label not present, team label present, adhoc label not present
                #test case label is present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {gdq_label} label is missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {gdq_label} label is missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label} label is missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 527",df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue, {gdq_label} label is missing and TC present Adhoc required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue, {gdq_label} label is missing and TC present Adhoc required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{gdq_label} label is missing and TC present Adhoc required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 566",df.iloc[i,0])
                    missing_info.append(result_dict)
    elif gdq_label in df.iloc[i,-1]:
        if team_label not in df.iloc[i,-1]:
            #gdd label is present and team label is not present
            if adhoc_label not in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                #adhoc and testcase label not present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {team_label}, {testcase_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {team_label}, {testcase_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{team_label}, {testcase_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 610",df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {team_label}, {adhoc_label} are missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {team_label}, {adhoc_label} are missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{team_label}, {adhoc_label} are missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 649",df.iloc[i,0])
                    missing_info.append(result_dict)
            elif adhoc_label in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                #gdq label is present, team label is not present and adhoc label is present
                #test case label is not present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue, {team_label} label is missing and {adhoc_label} present {testcase_label} required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue, {team_label} label is missing and {adhoc_label} present {testcase_label} required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{team_label} is missing and {adhoc_label} present, {testcase_label} required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 691",df.iloc[i,0])
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    #No -- adhoc label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {team_label} is missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {team_label} is missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{team_label} is missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 730",df.iloc[i,0])
                    missing_info.append(result_dict)
            elif adhoc_label not in df.iloc[i,-1] and testcase_label in df.iloc[i,-1]:
                #gdq label is present, team label is not present, test case label is present
                #adhoc label is not present
                if "Yes" in df.iloc[i,8]:
                    #Yes -- testcase label should be present
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        #if priority is missing 'Please add priority' will be used else the assigned priority will be used
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {team_label} label is missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {team_label} label is missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{team_label} label is missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 687")
                    missing_info.append(result_dict)
                
                elif "No" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue, {team_label} label is missing and TC present Adhoc required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue, {team_label} label is missing and TC present Adhoc required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{team_label} label is missing and TC present Adhoc required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 723")
                    missing_info.append(result_dict)
        
        elif team_label in df.iloc[i,-1]:
            if adhoc_label not in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                if "Yes" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue, {testcase_label} label is missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue, {testcase_label} label is missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{testcase_label} label is missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 761")
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue, {adhoc_label} label is missing'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue, {adhoc_label} label is missing'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{adhoc_label} label is missing'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 796")
                    missing_info.append(result_dict)
            
            elif adhoc_label in df.iloc[i,-1] and testcase_label not in df.iloc[i,-1]:
                if "Yes" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and {adhoc_label} present {testcase_label} required'
                    }
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and {adhoc_label} present {testcase_label} required'
                    }
                    else:                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{adhoc_label} present {testcase_label} required'
                        }
                    if df.iloc[i,0] not in missing_info_dict:
                        missing_info_dict[df.iloc[i,0]] = 1
                    print("Here 833")
                    missing_info.append(result_dict)
                elif "No" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue'
                    }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 848")
                        missing_info.append(result_dict)
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue'
                    }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 862")
                        missing_info.append(result_dict)
                    else:
                        #no issues here will go to no issues list                       
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':'NA'
                        }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("here 877")
                        no_missing_info.append(result_dict)
            
            elif adhoc_label not in df.iloc[i,-1] and testcase_label in df.iloc[i,-1]:
                if "Yes" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue'
                    }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 894")
                        missing_info.append(result_dict)
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue'
                    }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 908")
                        missing_info.append(result_dict)
                    else:
                        #no issues here will go to no issues list                         
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':'NA'
                        }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 923")
                        no_missing_info.append(result_dict)        
                elif "No" in df.iloc[i,8]:
                    if "Resolved" not in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{valid_label} is used for {df.iloc[i,3]} issue and TC present Adhoc required'
                    }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 851")
                        missing_info.append(result_dict)
                    elif "Resolved" not in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,0],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                        'Havign Testcase ?':df.iloc[i,-2],
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{invalid_label} is used for {df.iloc[i,3]} issue and TC present Adhoc required'
                    }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 864")
                        missing_info.append(result_dict)
                    else:
                        #no issues here will go to no issues list                        
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':'TC present Adhoc required'
                        }
                        if df.iloc[i,0] not in missing_info_dict:
                            missing_info_dict[df.iloc[i,0]] = 1
                        print("Here 967")
                        missing_info.append(result_dict)

#going to resolved issue
for i in range(len(df)):
    if "Resolved" in df.iloc[i,3]:
        if valid_label not in df.iloc[i,-1] or invalid_label not in df.iloc[i,-1]:
                result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{valid_label} or {invalid_label} is missing'
                        }
                if df.iloc[i,0] not in missing_info_dict:
                    missing_info_dict[df.iloc[i,0]] = 1
                print("Here 985")
                missing_info.append(result_dict)

for i in range(len(df)):
    if df.iloc[i,0] not in missing_info_dict:
        if df.iloc[i,2] not in ["P0","P1","P2","P3"]:
            result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':'NA'
                        }
            if df.iloc[i,0] not in missing_info_dict:
                missing_info_dict[df.iloc[i,0]] = 1
            print("Here 920")
            missing_info.append(result_dict)

for i in range(len(df)):
    if df.iloc[i,0] not in missing_info_dict:
        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,0],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please add priority',
                            'Havign Testcase ?':df.iloc[i,-2],
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':'NA'
                        }
        no_missing_info.append(result_dict)

#converting list's to pandas dataframe
missing_info_dataframe = pd.DataFrame(missing_info)
no_missing_info_dataframe = pd.DataFrame(no_missing_info)

#converting pandas dataframe to html
data_frame_to_html_missing_info = missing_info_dataframe.to_html(index=False)
data_frame_to_html_missing_info = "<h4>Rectification Required</h4>"+data_frame_to_html_missing_info
data_frame_to_html_no_info_missing = no_missing_info_dataframe.to_html(index=False)
data_frame_to_html_no_info_missing = "<h4>All Good</h4>"+data_frame_to_html_no_info_missing

#getting today's date for email subject
today_date = datetime.datetime.now().strftime('%x')

#email subject line
sub = f'Bug Audit Report for {today_date}.'

#generating email
try:
    smtpserver = "smtp.amazon.com"
    server = smt.SMTP(smtpserver)
    msg = EmailMessage()
    # from_ = input("Enter From Email Address : ")
    # to_ = input("Enter To Email Address : ")
    from_ = 'utukumar@amazon.com'
    to_ = 'utukumar@amazon.com'
    msg['To'] = to_
    msg['From'] = from_
    msg['Subject'] = sub
    with open(path, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=path)
    if missing_info and no_missing_info:
        email_msg = '''
                    <html>
                    <head>
                    <title>Weekly SIM Metrics</title>
                    <meta http–equiv="Content-Type" content="text/html; charset=UTF-8" />
                    <meta http–equiv=“X-UA-Compatible” content=“IE=edge” />
                    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                    <style>
                    table, th, td {border: 1.5px solid black; border-collapse: collapse; padding: 7px;}
                    </style>
                    </head>
                    <body>
                        <font face=Calibri>
                        <font size=2>
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for JIRA details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(total_issues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>'''+str(resolved_issues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>'''+str(open_issues)+'''</td></font></tr>
                        </table>
                        <br><br>
                        {missing_info}
                        <br><br>
                        {proper_info}
                    '''.format(missing_info=data_frame_to_html_missing_info, proper_info=data_frame_to_html_no_info_missing)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_,to_,msg.as_string())
        server.quit()
    elif missing_info:
        email_msg = '''
                    <html>
                    <head>
                    <title>Weekly SIM Metrics</title>
                    <meta http–equiv="Content-Type" content="text/html; charset=UTF-8" />
                    <meta http–equiv=“X-UA-Compatible” content=“IE=edge” />
                    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                    <style>
                    table, th, td {border: 1.5px solid black; border-collapse: collapse; padding: 7px;}
                    </style>
                    </head>
                    <body>
                        <font face=Calibri>
                        <font size=2>
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for JIRA details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(total_issues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>'''+str(resolved_issues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>'''+str(open_issues)+'''</td></font></tr>
                        </table>
                        <br><br>
                        {missing_info}
                        <br>
                    '''.format(missing_info=data_frame_to_html_missing_info)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_,to_,msg.as_string())
        server.quit()
    else:
        email_msg = '''
                    <html>
                    <head>
                    <title>Weekly SIM Metrics</title>
                    <meta http–equiv="Content-Type" content="text/html; charset=UTF-8" />
                    <meta http–equiv=“X-UA-Compatible” content=“IE=edge” />
                    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                    <style>
                    table, th, td {border: 1px solid black; border-collapse: collapse;}
                    </style>
                    </head>
                    <body>
                        <font face=Calibri>
                        <font size=2>
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for JIRA details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(total_issues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>'''+str(resolved_issues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>'''+str(open_issues)+'''</td></font></tr>
                        </table>
                        <br><br>
                        {}
                        <br><br>
                        
                    '''.format(data_frame_to_html_no_info_missing)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_,to_,msg.as_string())
        server.quit()
    print("Email Sent")
except Exception as e:
    print(e)
        

        
                       
            
            
                
                
                 
    
    
             
            

                           