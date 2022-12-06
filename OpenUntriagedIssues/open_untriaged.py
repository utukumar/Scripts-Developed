#comment out the follwoing line, if report generation is not automated
#import report_download
import xlrd
import openpyxl
import datetime
import tkinter as tk
from tkinter import filedialog
from tabulate import tabulate
import os
from pathlib import Path
import pandas as pd
untriagedOpen = 0
untopenassigned = 0
untopenunassigned = 0
user_dict = dict()
# pmohanap+OR+khpob+OR+jegathra+OR+janakaar+OR+bpriyant+OR+utukumar+OR+gaukuman+pjjonask+OR+gurbalat+OR+bpriyant+OR+vaijayam+OR+depik+OR+tijuwils+OR+vsofiya+OR+vshrth+OR+santhiyv+OR+babupree+OR+shisaxen+OR+saipriyr+OR+nivethp+OR+balaratr+OR+gznare+OR+loganat+OR+kalaiarm+OR+priydh

#Selecting the file from windows dialog
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfilename()
path = file_path

#selecting the current directory for saving the file
current_directory = os.getcwd()
save_file = current_directory
#print(save_file)

#reading excel sheet
excel_worksheet = xlrd.open_workbook(path)
excel_sheet = excel_worksheet.sheet_by_index(0)
sheet_name = str(excel_worksheet.sheet_names()[0])
column = excel_sheet.ncols
row_name = str(excel_sheet.cell_value(0, column-1))

df = pd.read_excel(path, sheet_name, parse_dates=["CreateDate"])
#open untriaged issues:
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,3)
    Column_tags = excel_sheet.cell_value(i,9)
    if "Open" in Column_status and "Triaged_No" in Column_tags:
        untriagedOpen = untriagedOpen+1

#open untriaged unassigned and assigned
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,3)
    Column_tags = excel_sheet.cell_value(i,9)
    Column_assigned = excel_sheet.cell_value(i,5)
    if "Open" in Column_status and "Triaged_No" in Column_tags and len(Column_assigned):
        untopenassigned +=1
    elif "Open" in Column_status and "Triaged_No" in Column_tags and len(Column_assigned) == 0:
        untopenunassigned +=1       

#reading user name list and project name list
project_list = list()
user_list = list()
with open(Path(save_file+'\\project_list.txt') , 'r') as f:
    for word in f.readlines():
        project_list.append(word.strip())
f.close()
with open(Path(save_file+'\\RequesterIdentity.txt') , 'r') as f:
    for word in f.readlines():
        user_list.append(word.strip())
f.close()

#getting open untriaged count based on user and project involved
user_and_project_wise = dict()
for i in range(1,excel_sheet.nrows):
    for project in project_list:
        for user in user_list:
            if "Open" in excel_sheet.cell_value(i,3) and "Triaged_No" in excel_sheet.cell_value(i,9):
                if user in excel_sheet.cell_value(i,4) and project in excel_sheet.cell_value(i,9):
                    if (user,project) in user_and_project_wise:                    
                        user_and_project_wise[(user, project)] +=1
                    else:                    
                        user_and_project_wise[(user, project)] =1
#print(user_and_project_wise)
#generating headers for tabulating the data
print(user_and_project_wise)
headers_1 = ['User','Project' ,'Untriaged Open Issue']
headers = ['User','Project' ,'Untriaged Open Issue', 'Ageing of Defects[{"days":"count"}]']
date_dict=dict()
#date_list = []
for i in range(1,excel_sheet.nrows):
    for project in project_list:
        for user in user_list:
            if "Open" in excel_sheet.cell_value(i,3) and "Triaged_No" in excel_sheet.cell_value(i,9):            
                if user in excel_sheet.cell_value(i,4) and project in excel_sheet.cell_value(i,9):                
                    if (user, project) not in date_dict:
                        date_dict[(user,project)] = list()
                        date_dict[(user,project)].append((datetime.datetime.now()-df[df.columns[6]][i-1]).days)
                    else:
                        date_dict[(user,project)].append((datetime.datetime.now()-df[df.columns[6]][i-1]).days)
common_keys = user_and_project_wise.keys() & date_dict.keys()                    
# with smtplib.SMTP('smtp.amazon.com') as smtp:
#     smtp.ehlo()
#getting shortid by user
short_id_dict = {}
                                                                    
for i in range(1,excel_sheet.nrows):
    for project in project_list:
        for user in user_list:
            if "Open" in excel_sheet.cell_value(i,3) and "Triaged_No" in excel_sheet.cell_value(i,9):
                if user in excel_sheet.cell_value(i,4) and project in excel_sheet.cell_value(i,9):
                    if (user,project) not in short_id_dict:
                        short_id_dict[(user,project)] = list()
                        short_id_dict[(user,project)].append(excel_sheet.cell_value(i,0))
                    else:
                        short_id_dict[(user,project)].append(excel_sheet.cell_value(i,0))    
#getting issue days by user and project
date_dict=dict()
for i in range(1,excel_sheet.nrows):
    for project in project_list:
        for user in user_list:
            if "Open" in excel_sheet.cell_value(i,3) and "Triaged_No" in excel_sheet.cell_value(i,9):            
                if user in excel_sheet.cell_value(i,4) and project in excel_sheet.cell_value(i,9):                
                    if (user, project) not in date_dict:
                        date_dict[(user,project)] = list()
                        date_dict[(user,project)].append((datetime.datetime.now()-df[df.columns[6]][i-1]).days)
                    else:
                        date_dict[(user,project)].append((datetime.datetime.now()-df[df.columns[6]][i-1]).days)

#header list for writing to excel sheet                    
headers_issueid_wise = ['User','Project','Count','IssueId and Days']
#converting list comprehension of user_project_issueID_ageing_of_defects to pandas datafram
df_mine = pd.DataFrame(sorted([(k1[0],k1[1],v3,[(j,i) for i,j in zip(v1,v2)])for(k1,v1),(k2,v2),(k3,v3) in zip(date_dict.items(), short_id_dict.items(), user_and_project_wise.items())]), columns = headers_issueid_wise)
#writing above dataframe to excelsheet
df_mine.to_excel(Path(save_file+'\\sorted_by_issueid.xlsx'), columns = headers_issueid_wise, index=False)                   
with open(Path(save_file+"\\Untriaged Counter.txt"), "w", encoding='utf-8') as f:
    f.write("--------------------------------------------------------------------------")
    f.write("\n")
    f.write("Created on : "+str(datetime.datetime.now().strftime('%x')))
    f.write("\n")
    #comment out the following line, if report download is not automated
    #f.write("From[MM/DD/YYYY] : "+str(report_download.start_date)+" to "+str(report_download.end_date))
    f.write("\n")
    f.write("\n")
    f.write("Open Untriaged Issues = "+str(untriagedOpen))
    f.write("\n")
    f.write("Open Untraiged Assigned = "+str(untopenassigned))
    f.write("\n")
    f.write("Open Untriaged Unassigned = "+str(untopenunassigned))
    f.write("\n")
    f.write("\n")
    f.write("--------------------------------------------------------------------------")
    f.write("\n")
    f.write("\n")
    # for key,value in user_dict.items():
    #     f.write(str(key)+" -- "+"  "+str(value))
    f.write(tabulate(sorted([(k[0], k[1], v) for k,v in user_and_project_wise.items()]), headers=headers_1, tablefmt='fancy_grid',numalign = "center"))
    f.write("\n")
    f.write("\n")
f.close()
print("File Created Successfully")                  
          