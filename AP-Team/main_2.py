from datetime import datetime
from turtle import left
import pandas as pd
import os
from pathlib import Path
from openpyxl import workbook, load_workbook
import tkinter as tk
import datetime
from tkinter import CENTER, filedialog
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

#Labels used for AP team
ap_label = 'GDQ_QS_Detected_AP'
testcase_label = 'QS_Testcase'
adhoc_label = 'QS_Adhoc'
valid_label = 'QS_Detected_Valid'
invalid_label = 'QS_Detected_Invalid'

#tags used for AP team
adhoc_regression = 'Adhoc_Regression'
adhoc_progression = 'Adhoc_Progression'
tc_regression = 'TC_Regression'
tc_progression = 'TC_Progression'
fixed = 'QS-Fixed'
cannot_reproduce = 'QS-CannotReproduce'
valid_by_design = 'Valid-Bydesign'
duplicate = 'QS-Duplicate'
not_an_issue = 'QS-NotanIssue'
backlog_issue = "QS-Backlogs"
feature_change = "QS-FeatureChange"
deferred = "QS-Deferred"
feature_enhancement = "QS_Feature_Enhancement"



#information that needs to be calculated
total_issues_raised = 0
resolved_issues = 0
open_issues = 0
high = 0
medium = 0
low = 0
resolved_high = 0
resolved_medium = 0
resolved_low = 0
open_high = 0
open_medium = 0
open_low = 0

#calcualting total issues
total_issues_raised = len(df)
print("Total Issues Raised : ",total_issues_raised)

#calculating open and resolved issues:
for i in range(len(df)):
    if 'Open' in df.iloc[i,3]:
        open_issues += 1
    elif 'Resolved' in df.iloc[i,3]:
        resolved_issues += 1
print("Total Open Issues : ",open_issues)
print("Total Resolved Issues : ",resolved_issues)

#calculating high medium low issues:
for i in range(len(df)):
    if 'High' in df.iloc[i,2]:
        high += 1
    elif 'Medium' in df.iloc[i,2]:
        medium += 1
    elif 'Low' in df.iloc[i,2]:
        low += 1
print()
print("High Issues : ",high)
print("Medium Issues : ",medium)
print("Low Issues : ",low)

#high medium low open issues
for i in range(len(df)):
    if 'Open' in df.iloc[i,3] and 'High' in df.iloc[i,2]:
        open_high += 1
    elif 'Open' in df.iloc[i,3] and 'Medium' in df.iloc[i,2]:
        open_medium += 1
    elif 'Open' in df.iloc[i,3] and 'Low' in df.iloc[i,2]:
        open_low += 1
print("Open High Issues : ",open_high)
print("Open Medium Issues : ",open_medium)
print("Open Low Issues : ",open_low)

#high medium low resolved issues
for i in range(len(df)):
    if 'Resolved' in df.iloc[i,3] and 'High' in df.iloc[i,2]:
        resolved_high += 1
    elif 'Resolved' in df.iloc[i,3] and 'Medium' in df.iloc[i,2]:
        resolved_medium += 1
    elif 'Resolved' in df.iloc[i,3] and 'Low' in df.iloc[i,2]:
        resolved_low += 1
print("Resolved High Issues : ",resolved_high)
print("Resolved Medium Issues : ",resolved_medium)
print("Resolved Low Issues : ",resolved_low)

#list of storing issues which have some information missing
missing_information = []
#list to store the issues that don't have any problem
issues_no_missing_information = []

#short id dict for storing the issue id, so we don't have repeating issues
shortIdDict = {} 
       
#checking if user has added proper adhoc  labels or not
for i in range(len(df)):
    if (adhoc_progression not in df.iloc[i,-2] and adhoc_regression not in df.iloc[i,-2]):
        if (adhoc_label in df.iloc[i,-1]):
            if ap_label in df.iloc[i,-1]:
                if (tc_progression not in df.iloc[i,-2] and tc_regression not in df.iloc[i,-2]):
                    #if adhoc tag is not present, adhoc_label is present, ap_label is present and tc tag is not present
                    #Assuming that adhoc label is correct we need Adhoc Tag
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"Adhoc Required"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"Adhoc Required"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':"NA",
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"Adhoc Required"                       
                        }
                elif (tc_progression in df.iloc[i,-2] or tc_regression in df.iloc[i,-2]):
                    #if adhoc tag is not present, adhoc_label is present, ap_label is present and tc tag is present
                    #Assuming tc_tag is correct we need TC label insted of Adhoc label
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'"{invalid_label} used for Open Issue" and "Adhoc Present, TC Required"',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"NA"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'"{valid_label} used for Open Issue" and "Adhoc Present, TC Required"',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"NA"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':"Adhoc Present, TC Required",
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"NA"                       
                                    }
                #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here1',df.iloc[i,0])
                #appending the result_dictionary to missing_information list
                missing_information.append(result_dict)
            elif ap_label not in df.iloc[i,-1]:
                if (tc_progression not in df.iloc[i,-2] and tc_regression not in df.iloc[i,-2]):
                    #if adhoc_tag is not present, adhoc_label is present, ap_label is not present and tc tag is not present
                    #Assuming Adhoc label is correct, we will need ap_label and Adhoc tag
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue and {ap_label} required',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"Adhoc Required"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue and {ap_label} required',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"Adhoc Required"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':ap_label,
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"Adhoc Required"                       
                        }
                elif (tc_progression in df.iloc[i,-2] or tc_regression in df.iloc[i,-2]):
                    #if adhoc_tag is not present, adhoc_label is present, ap_label is not present and tc tag is present
                    #Assuming tc tag is correct, we will need ap_label and need to replace Adhoc_Label with TC_label
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue and Adhoc Present, TC Required',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"NA"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue and Adhoc Present, TC Required',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"NA"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{ap_label}, "Adhoc Present, TC Required"',
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"NA"                       
                        }
                #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here2',df.iloc[i,0])
                #appending the result_dictionary to missing_information list
                missing_information.append(result_dict)      
    elif (adhoc_progression in df.iloc[i,-2] or adhoc_regression in df.iloc[i,-2]):
        if adhoc_label not in df.iloc[i,-1]:
            if ap_label in df.iloc[i,-1]:
                if testcase_label not in df.iloc[i,-1]:
                    #if adhoc_tag is present, adhoc_label is not present, ap_label is present and tc_label is not present
                    #Assuming adhoc_tag is correct, we will need adhoc_label
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue and {adhoc_label} missing',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"NA"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue and {adhoc_label} missing',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"NA"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':adhoc_label,
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"NA"                       
                        }
                elif testcase_label in df.iloc[i,-1]:
                    #if adhoc_tag is present, adhoc_label is not present, ap_label is present and tc_label is present
                    #Assuming adhoc_tag is correct, also tc_label is present and need to be replaced with adhoc label
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue and "TC Present, Adhoc Required"',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"NA"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue and "TC Present, Adhoc Required"',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"NA"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':"TC Present, Adhoc Required",
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"NA"                       
                        }
                #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here5',df.iloc[i,0])
                #appending the result_dictionary to missing_information list
                missing_information.append(result_dict)
            elif ap_label not in df.iloc[i,-1]:
                if testcase_label not in df.iloc[i,-1]:
                    #if adhoc_tag is present, adhoc_label is not present, ap_label is not present and tc_label is not present
                    #Assuming adhoc_tag is correct, we need adhoc_label and ap_label
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue and {ap_label}, {adhoc_label} are missing',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"NA"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue and {ap_label}, {adhoc_label} are missing',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"NA"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{ap_label}, {adhoc_label}',
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"NA"                       
                        }
                else:
                    #if adhoc_tag is present, adhoc_label is not present, ap_label is not present and tc_label is present
                    #Assuming adhoc_tag is correct, we need ap_label and replace tc_label with adhoc_label
                    if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                        result_dict = {
                                        'User Alias':df.iloc[i,4],
                                        'Issue URL':df.iloc[i,8],
                                        'Status':df.iloc[i,3],
                                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                        'Label Present':df.iloc[i,-1],
                                        'Label Missing':f'{invalid_label} used for Open Issue , {ap_label} missing and "TC Present, Adhoc Required"',
                                        'Tags Present':df.iloc[i,-2],
                                        'Missing Tag':"NA"
                                        }
                    elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                        result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{valid_label} used for Open Issue, {ap_label} missing and "TC Present, Adhoc Required"',
                                'Tags Present':df.iloc[i,-2],
                                'Missing Tag':"NA"
                                    }
                    else:
                        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{ap_label} missing , "TC Present, Adhoc Required"',
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag':"NA"                       
                        }
                #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here6',df.iloc[i,0])
                #appending the result_dictionary to missing_information list
                missing_information.append(result_dict)
        elif adhoc_label in df.iloc[i,-1]:
            if ap_label in df.iloc[i,-1]:
                #adhoc_tag is present, adhoc_label is present and ap_label is present
                #we need to verify only if the open issues don't have valid_label or invalid_label
                if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                    result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{invalid_label} used for Open Issue',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"NA"
                                    }
                    if df.iloc[i,0] not in shortIdDict:
                        shortIdDict[df.iloc[i,0]] = 1
                    missing_information.append(result_dict)
                    print("here458",df.iloc[i,0])
                elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                    result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{valid_label} used for Open Issue',
                            'Tags Present':df.iloc[i,-2],
                            'Missing Tag':"NA"
                                }
                    if df.iloc[i,0] not in shortIdDict:
                        shortIdDict[df.iloc[i,0]] = 1
                    missing_information.append(result_dict)
                    print("here479",df.iloc[i,0])
                else:
                    if df.iloc[i,0] not in shortIdDict:
                        if df.iloc[i,2] not in ["High", "Medium", "Low"]:
                            result_dict = {
                                            'User Alias':df.iloc[i,4],
                                            'Issue URL':df.iloc[i,8],
                                            'Status':df.iloc[i,3],
                                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                            'Label Present':df.iloc[i,-1],
                                            'Label Missing':"NA",
                                            #'Mismatch Between Label and Tag':"",
                                            'Tags Present':df.iloc[i,-2],                        
                                            'Missing Tag': "NA"                       
                                        }
                            if df.iloc[i,0] not in shortIdDict:
                                shortIdDict[df.iloc[i,0]] = 1
                            print('Here492',df.iloc[i,0])
                            missing_information.append(result_dict)
                        else:
                            if ("Resolved" in df.iloc[i,3] and valid_label in df.iloc[i,-1] and (fixed in df.iloc[i,-2] or cannot_reproduce in df.iloc[i,-2] or valid_by_design in df.iloc[i,-2] or backlog_issue in df.iloc[i,-2] or feature_change in df.iloc[i,-2] or deferred in df.iloc[i,-2] or feature_enhancement in df.iloc[i,-2])):
                                result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':"NA",
                                    'Tags Present':df.iloc[i,-2],                        
                                    'Missing Tag':"NA"                       
                                }
                                if df.iloc[i,0] not in shortIdDict:
                                    shortIdDict[df.iloc[i,0]] = 1
                                issues_no_missing_information.append(result_dict)
                                print("here_no_issue513",df.iloc[i,0])
                            elif ("Resolved" in df.iloc[i,3] and invalid_label in df.iloc[i,-1] and (not_an_issue in df.iloc[i,-2] or duplicate in df.iloc[i,-2])):
                                result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':"NA",
                                    'Tags Present':df.iloc[i,-2],                        
                                    'Missing Tag':"NA"                       
                                }
                                if df.iloc[i,0] not in shortIdDict:
                                    shortIdDict[df.iloc[i,0]] = 1
                                issues_no_missing_information.append(result_dict)
                                print("here_no_issue528",df.iloc[i,0])
                
                #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                #if df.iloc[i,0] not in shortIdDict:
                #    shortIdDict[df.iloc[i,0]] = 1
                #print('Here7')
                #appending the result_dictionary to missing_information list
                
            elif ap_label not in df.iloc[i,-1]:
                #adhoc_tag is present, adhoc_label is present and ap_label is not present
                #we need ap_label and need to verify only if the open issues don't have valid_label or invalid_label
                if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                    result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{invalid_label} used for Open Issue and {ap_label} is missing',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"NA"
                                    }
                elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                    result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{valid_label} used for Open Issue and {ap_label} is missing',
                            'Tags Present':df.iloc[i,-2],
                            'Missing Tag':"NA"
                                }
                else:
                    result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,8],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f'{ap_label}',
                        'Tags Present':df.iloc[i,-2],                        
                        'Missing Tag':"NA"                       
                    }
                #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here8',df.iloc[i,0])
                #appending the result_dictionary to missing_information list
                missing_information.append(result_dict)
             

#checking ap lable with tctag:
for i in range(len(df)):
    if (adhoc_progression not in df.iloc[i,-2] and adhoc_regression not in df.iloc[i,-2]):
        if adhoc_label not in df.iloc[i,-1]:
            if testcase_label not in df.iloc[i,-1]:
                if ap_label in df.iloc[i,-1]:
                    if tc_progression in df.iloc[i,-2] or tc_regression in df.iloc[i,-2]:
                        #if adhoc_tag is not present, adhoc_label is not present, ap_label is present and tc_tag is present
                        #Assuming tc_tag is correct and we have ap_label, so we need tc_label
                        if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                            result_dict = {
                                            'User Alias':df.iloc[i,4],
                                            'Issue URL':df.iloc[i,8],
                                            'Status':df.iloc[i,3],
                                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                            'Label Present':df.iloc[i,-1],
                                            'Label Missing':f'{invalid_label} used for Open Issue and {testcase_label} missing',
                                            'Tags Present':df.iloc[i,-2],
                                            'Missing Tag':"NA"
                                            }
                        elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                            result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{valid_label} used for Open Issue and {testcase_label} missing',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"NA"
                                        }
                        else:
                            result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':testcase_label,
                                'Tags Present':df.iloc[i,-2],                        
                                'Missing Tag':"NA"                       
                            }
                    else:
                        #if adhoc_tag is not present, adhoc_label is not present, ap_label is present and tc_tag is not present
                        #As we don't have proper label and tags -- "No Adhoc or TC label present" and "No Adhoc or TC tag present", but we have ap_label
                        if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                            result_dict = {
                                            'User Alias':df.iloc[i,4],
                                            'Issue URL':df.iloc[i,8],
                                            'Status':df.iloc[i,3],
                                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                            'Label Present':df.iloc[i,-1],
                                            'Label Missing':f'{invalid_label} used for Open Issue and No Adhoc or TC Label present',
                                            'Tags Present':df.iloc[i,-2],
                                            'Missing Tag':"No Adhoc or TC tag present"
                                            }
                        elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                            result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{valid_label} used for Open Issue and No Adhoc or TC Label present',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"No Adhoc or TC tag present"
                                        }
                        else:
                            result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':'No Adhoc or TC Label present',
                                'Tags Present':df.iloc[i,-2],                        
                                'Missing Tag':"No Adhoc or TC tag present"                       
                            }
                    #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                    if df.iloc[i,0] not in shortIdDict:
                        shortIdDict[df.iloc[i,0]] = 1
                    print('Here9',df.iloc[i,0])
                    #appending the result_dictionary to missing_information list
                    missing_information.append(result_dict)
                else:
                    if tc_progression in df.iloc[i,-2] or tc_regression in df.iloc[i,-2]:
                        #if adhoc_tag is not present, adhoc_label is not present, ap_label and testcase_label are not present and tc_tag is present
                        #Assuming tc_tag is correct, we need ap_label and tc_label
                        if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                            result_dict = {
                                            'User Alias':df.iloc[i,4],
                                            'Issue URL':df.iloc[i,8],
                                            'Status':df.iloc[i,3],
                                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                            'Label Present':df.iloc[i,-1],
                                            'Label Missing':f'{invalid_label} used for Open Issue and {ap_label}, {testcase_label} are missing',
                                            'Tags Present':df.iloc[i,-2],
                                            'Missing Tag':"NA"
                                            }
                        elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                            result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{valid_label} used for Open Issue and {ap_label}, {testcase_label} are missing',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"NA"
                                        }
                        else:
                            result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{ap_label}, {testcase_label}',
                                'Tags Present':df.iloc[i,-2],                        
                                'Missing Tag':"NA"                       
                            }
                    else:
                        #if adhoc_tag is not present, adhoc_label is not present, ap_label is not present and tc_label, tc_tag is not present
                        #then we need ap_label, "No Adhoc or tc label present", "No Adhoc or TC tag present"
                        if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                            result_dict = {
                                            'User Alias':df.iloc[i,4],
                                            'Issue URL':df.iloc[i,8],
                                            'Status':df.iloc[i,3],
                                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                            'Label Present':df.iloc[i,-1],
                                            'Label Missing':f'{invalid_label} used for Open Issue, {ap_label} is missing and No Adhoc or TC label present',
                                            'Tags Present':df.iloc[i,-2],
                                            'Missing Tag':"No Adhoc or TC tag present"
                                            }
                        elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                            result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{valid_label} used for Open Issue, {ap_label} is missing and No Adhoc or TC label present',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"No Adhoc or TC tag present"
                                        }
                        else:
                            result_dict = {
                                'User Alias':df.iloc[i,4],
                                'Issue URL':df.iloc[i,8],
                                'Status':df.iloc[i,3],
                                'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                'Label Present':df.iloc[i,-1],
                                'Label Missing':f'{ap_label}, No Adhoc or TC label present',
                                'Tags Present':df.iloc[i,-2],                        
                                'Missing Tag':"No Adhoc or TC tag present"                       
                            }
                    #putting the issue id in shortid dict so the same issues are not replicated in missing_information list
                    if df.iloc[i,0] not in shortIdDict:
                        shortIdDict[df.iloc[i,0]] = 1
                    print('Here10',df.iloc[i,0])
                    #appending the result_dictionary to missing_information list
                    missing_information.append(result_dict)
            
#checking if valid_label and invalid_label are not used with open issues having tc_tag and tc_label                 
for i in range(len(df)):
    if df.iloc[i,0] not in shortIdDict:
        #print(df.iloc[i,0])
        if (tc_progression in df.iloc[i,-2] or tc_regression in df.iloc[i,-2]) and testcase_label in df.iloc[i,-1]:
            if ap_label not in df.iloc[i,-1]:
                if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                    result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{invalid_label} used for Open Issue and {ap_label} missing',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"NA"
                                    }
                elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                    result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{valid_label} used for Open Issue and {ap_label} missing',
                            'Tags Present':df.iloc[i,-2],
                            'Missing Tag':"NA"
                                }
                else:
                    result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{ap_label} is missing',
                            'Tags Present':df.iloc[i,-2],
                            'Missing Tag':"NA"
                                }

                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                    print('Herenext',df.iloc[i,0])
                missing_information.append(result_dict)
            #TODO:tc tag, tc label and ap label are present but valid/invalid label is used
            #if any error are seen in the mail check code block below
            elif ap_label in df.iloc[i,-1]:
                if 'Open' in df.iloc[i,3] and invalid_label in df.iloc[i,-1]:
                    result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':f'{invalid_label} used for Open Issue',
                                    'Tags Present':df.iloc[i,-2],
                                    'Missing Tag':"NA"
                                    }
                    if df.iloc[i,0] not in shortIdDict:
                        shortIdDict[df.iloc[i,0]] = 1
                    print('Herenext766',df.iloc[i,0])
                    missing_information.append(result_dict)
                elif 'Open' in df.iloc[i,3] and valid_label in df.iloc[i,-1]:
                    result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':f'{valid_label} used for Open Issue',
                            'Tags Present':df.iloc[i,-2],
                            'Missing Tag':"NA"
                                }
                    if df.iloc[i,0] not in shortIdDict:
                        shortIdDict[df.iloc[i,0]] = 1
                    print('Herenext781',df.iloc[i,0])
                    missing_information.append(result_dict)
                else:
                    if df.iloc[i,0] not in shortIdDict:
                        if df.iloc[i,2] not in ["High", "Medium", "Low"]:
                            result_dict = {
                                            'User Alias':df.iloc[i,4],
                                            'Issue URL':df.iloc[i,8],
                                            'Status':df.iloc[i,3],
                                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                            'Label Present':df.iloc[i,-1],
                                            'Label Missing':"NA",
                                            #'Mismatch Between Label and Tag':"",
                                            'Tags Present':df.iloc[i,-2],                        
                                            'Missing Tag': "NA"                       
                                        }
                            if df.iloc[i,0] not in shortIdDict:
                                shortIdDict[df.iloc[i,0]] = 1
                            print('Here817',df.iloc[i,0])
                            missing_information.append(result_dict)
                        else:
                            #checks have to be written to verify  for resolved issues same for adhoc loop as well
                            if ("Resolved" in df.iloc[i,3] and valid_label in df.iloc[i,-1] and (fixed in df.iloc[i,-2] or cannot_reproduce in df.iloc[i,-2] or valid_by_design in df.iloc[i,-2] or backlog_issue in df.iloc[i,-2] or feature_change in df.iloc[i,-2] or deferred in df.iloc[i,-2] or feature_enhancement in df.iloc[i,-2])):
                                result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':"NA",
                                    'Tags Present':df.iloc[i,-2],                        
                                    'Missing Tag':"NA"                       
                                }
                                if df.iloc[i,0] not in shortIdDict:
                                    shortIdDict[df.iloc[i,0]] = 1
                                issues_no_missing_information.append(result_dict)
                                print("here_no_issue839",df.iloc[i,0])
                            elif ("Resolved" in df.iloc[i,3] and invalid_label in df.iloc[i,-1] and (not_an_issue in df.iloc[i,-2] or duplicate in df.iloc[i,-2])):
                                result_dict = {
                                    'User Alias':df.iloc[i,4],
                                    'Issue URL':df.iloc[i,8],
                                    'Status':df.iloc[i,3],
                                    'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                                    'Label Present':df.iloc[i,-1],
                                    'Label Missing':"NA",
                                    'Tags Present':df.iloc[i,-2],                        
                                    'Missing Tag':"NA"                       
                                }
                                if df.iloc[i,0] not in shortIdDict:
                                    shortIdDict[df.iloc[i,0]] = 1
                                issues_no_missing_information.append(result_dict)
                                print("here_no_issue854",df.iloc[i,0])
            
     
#checking for resolved issues
for i in range(len(df)):
    if 'Resolved' in df.iloc[i,3]:
        if valid_label not in df.iloc[i,-1]:
            if fixed in df.iloc[i,-2] or cannot_reproduce in df.iloc[i,-2] or valid_by_design in df.iloc[i,-2] or backlog_issue in df.iloc[i,-2] or feature_change in df.iloc[i,-2] or deferred in df.iloc[i,-2] or feature_enhancement in df.iloc[i,-2]:
                #if the issue is resolved, and valid_label is not present and proper tags are present
                #the we need valid_label
                result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,8],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':valid_label,
                        'Tags Present':df.iloc[i,-2],                        
                        'Missing Tag':"NA"                       
                    }
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here13',df.iloc[i,0])
                missing_information.append(result_dict)
        elif valid_label in df.iloc[i,-1]:
            if fixed not in df.iloc[i,-2] and cannot_reproduce not in df.iloc[i,-2] and valid_by_design not in df.iloc[i,-2] and backlog_issue in df.iloc[i,-2] and feature_change in df.iloc[i,-2] and deferred in df.iloc[i,-2] and feature_enhancement in df.iloc[i,-2]:
                #if the issue is resolved, valid_label is present and proper tags are not present
                #the we need "Resolution tags are missing"
                result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,8],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':"NA",
                        #'Mismatch Between Label and Tag':"",
                        'Tags Present':df.iloc[i,-2],                        
                        'Missing Tag':"Resolution Tags are missing"                       
                    }
                #print('Inside valid resolution')
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here14',df.iloc[i,0])
                missing_information.append(result_dict)


#checking for resolved issues invalid label
for i in range(len(df)):
    if 'Resolved' in df.iloc[i,3]:
        if invalid_label not in df.iloc[i,-1]:
            if not_an_issue in df.iloc[i,-2] or duplicate in df.iloc[i,-2]:
                #have to use invalid by design tag in the above if condition
                result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,8],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':invalid_label,
                        #'Mismatch Between Label and Tag':"",
                        'Tags Present':df.iloc[i,-2],                        
                        'Missing Tag':"NA"                       
                    }
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here15',df.iloc[i,0])
                missing_information.append(result_dict)
        elif invalid_label in df.iloc[i,-1]:
            if not_an_issue not in df.iloc[i,-2] and duplicate not in df.iloc[i,-2]:
                #have to consider the case where invalid-label is present with fixed, cannotreproduce, validbydesign tags
                result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,8],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':"NA",
                        #'Mismatch Between Label and Tag':"",
                        'Tags Present':df.iloc[i,-2],                        
                        'Missing Tag': "Resolution Tags are missing"                       
                    }
                #print('Inside invalid resolution')
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here16',df.iloc[i,0])
                missing_information.append(result_dict)

#checking if resolved issue does not have valid/invlaid labels are missing and resolution tags are missing
for i in range(len(df)):
    if "Resolved" in df.iloc[i,3]:
        if valid_label not in df.iloc[i,-1] and invalid_label not in df.iloc[i,-1]:
            if fixed not in df.iloc[i,-2] and cannot_reproduce not in df.iloc[i,-2] and valid_by_design not in df.iloc[i,-2] and backlog_issue not in df.iloc[i,-2] and feature_change not in df.iloc[i,-2] and deferred not in df.iloc[i,-2] and feature_enhancement not in df.iloc[i,-2] and not_an_issue not in df.iloc[i,-2] and duplicate not in df.iloc[i,-2]:
                result_dict = {
                        'User Alias':df.iloc[i,4],
                        'Issue URL':df.iloc[i,8],
                        'Status':df.iloc[i,3],
                        'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                        'Label Present':df.iloc[i,-1],
                        'Label Missing':f"'{valid_label}' OR '{invalid_label}' Missing",
                        #'Mismatch Between Label and Tag':"",
                        'Tags Present':df.iloc[i,-2],                        
                        'Missing Tag': "Resolution Tags are missing"                       
                    }
                #print('Inside invalid resolution')
                if df.iloc[i,0] not in shortIdDict:
                    shortIdDict[df.iloc[i,0]] = 1
                print('Here17',df.iloc[i,0])
                missing_information.append(result_dict)
                
#Checking if the proper issues are having priority or not
for i in range(len(df)):
    if df.iloc[i,0] not in shortIdDict:
        if df.iloc[i,2] not in ["High", "Medium", "Low"]:
            result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':"NA",
                            #'Mismatch Between Label and Tag':"",
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag': "NA"                       
                        }
            if df.iloc[i,0] not in shortIdDict:
                shortIdDict[df.iloc[i,0]] = 1
            print('Here18',df.iloc[i,0])
            missing_information.append(result_dict)

#debub print
print(shortIdDict)
            
#appending other issue to issues_no_missing_information list
for i in range(len(df)):
    if df.iloc[i,0] not in shortIdDict:
        result_dict = {
                            'User Alias':df.iloc[i,4],
                            'Issue URL':df.iloc[i,8],
                            'Status':df.iloc[i,3],
                            'Priority':df.iloc[i,2] if df.iloc[i,2] else 'Please Add Priority',
                            'Label Present':df.iloc[i,-1],
                            'Label Missing':"NA",
                            #'Mismatch Between Label and Tag':"",
                            'Tags Present':df.iloc[i,-2],                        
                            'Missing Tag': "NA"                       
                        }
        issues_no_missing_information.append(result_dict)
#converting list's to pandas dataframe
missing_df = pd.DataFrame(missing_information)
no_missing_df = pd.DataFrame(issues_no_missing_information)


#missing_df.to_excel(r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\output.xlsx', index=False)
#no_missing_df.to_excel(r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\proper.xlsx',index=False)
#converting pandas dataframe to html
data_frame_to_html = missing_df.to_html(justify="center",index=False)
#data_frame_to_html.replace('<td>','<td style="text-align: right;">')
data_frame_to_html = "<h4>Rectification Required</h4>"+data_frame_to_html
data_frame_to_html_no_info_missing = no_missing_df.to_html(index=False)
data_frame_to_html_no_info_missing = "<h4>All Good</h4>"+data_frame_to_html_no_info_missing
#with open(r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\datafram.txt',"w") as f:
#    f.write(data_frame_to_html)
#debug print
print(missing_information)

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
    if missing_information and issues_no_missing_information:
        email_msg = '''
                    <html>
                    <head>
                    <title>Weekly SIM Metrics</title>
                    <meta http–equiv="Content-Type" content="text/html; charset=UTF-8" />
                    <meta http–equiv=“X-UA-Compatible” content=“IE=edge” />
                    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                    <style>
                    table, th, td,tr {border: 1.5px solid black; border-collapse: collapse; padding: 7px; text-align: center}
                    </style>
                    </head>
                    <body>
                        <font face=Calibri>
                        <font size=2>
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(total_issues_raised)+'''</td></font></tr>
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
                    '''.format(missing_info=data_frame_to_html, proper_info=data_frame_to_html_no_info_missing)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_,to_,msg.as_string())
        server.quit()
    elif missing_information:
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(total_issues_raised)+'''</td></font></tr>
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
                    '''.format(missing_info=data_frame_to_html)
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(total_issues_raised)+'''</td></font></tr>
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
        
