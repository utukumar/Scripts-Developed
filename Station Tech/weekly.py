from poplib import CR
from time import strftime
import xlrd
import nltk
import os
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
import datetime
#import stationtech_report_download
import smtplib as smt
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfilename()
path = file_path
#user_name = str(input('Enter your username(amazon alias): '))

current_directory = os.getcwd()
save_file = current_directory
#print(save_file)

high = 0
medium = 0
low = 0
openIssues = 0
openHigh = 0
openMedium = 0
openLow = 0
resolvedIssues = 0
resolvedHigh = 0
resolvedMedium = 0
resolvedLow = 0
assigned = 0
assignedHigh = 0
assignedMedium = 0
assignedLow = 0
unassigned = 0
unassignedHigh = 0
unassignedMedium = 0
unassignedLow = 0
adhoc = 0
adhocHigh = 0
adhocMedium = 0
adhocLow = 0
tc = 0
tcHigh = 0
tcMedium = 0
tcLow = 0
fixed =0
fixedHigh = 0
fixedMedium = 0
fixedLow = 0
triaged = 0
triagedHigh = 0
triagedMedium = 0
triagedLow = 0
untriaged = 0
untriagedHigh = 0
untriagedMedium = 0
untriagedLow = 0
wontFix = 0
wontFixHigh = 0
wontFixMedium = 0
wontFixLow = 0
deferred = 0
deferredHigh = 0
deferredMedium = 0
deferredLow = 0
cantrepo = 0
cantrepoHigh = 0
cantrepoMedium = 0
cantrepoLow = 0
dup = 0
dupHigh = 0
dupMedium = 0
dupLow = 0
notbug = 0
notbugHigh = 0
notbugMedium = 0
notbugLow = 0
isbug = 0
isbugHigh = 0
isbugMedium = 0
isbugLow = 0
defectFixRatio = 0
validDefectRatio = 0
bydesign = 0
bydesignHigh = 0
bydesignMedium = 0
bydesignLow = 0
# valid = 0
# validHigh = 0
# validMedium = 0
# validLow = 0
valid_label = 0
valid_labelHigh = 0
valid_labelMedium = 0
valid_labelLow = 0
invalid_label = 0
invalid_labelHigh = 0
invalid_labelMedium = 0
invalid_labelLow = 0

# Critical = 'Priority_Critical'
# Blocker = 'Priority_Blocker'
# Medium = 'Priority_Medium'
# Minor = 'Priority_Minor'
adhocTag = 'QS_ST_Adhoc'
adhocTag_new = 'QS_Adhoc'
tcTag = 'QS_ST_TC'
tcTag_new = 'QS_Testcase'
tcTagNotRequired = 'QS_ST_TC_NotRequired'
Triaged = 'QS_ST_Triaged'
Fixed = 'QS_ST_Fixed'
Fixed_new = 'Fixed'
WontFix_new = "Won’t fix"
Deferred_new = 'Deferred'
CantReproduce_new = 'Cannot_reproduce'
Duplicate_new = 'Duplicate'
NotBug_new = 'Not_a_bug'
WontFix = 'QS_ST_WontFix'
Deferred = 'QS_ST_Deferred'
CantReproduce = 'QS_ST_NotRepo'
Duplicate = 'QS_ST_Duplicate'
NotBug = 'QS_ST_NotABug'
ByDesign = 'QS_ST_ByDesign'
ByDesign_new = 'By_Design'
ValidLabel = 'QS_Detected_Valid'
#IsBug = 'QS_ST_Bug'

labels_list = [adhocTag, adhocTag_new, tcTag, tcTag_new, Triaged, Fixed, WontFix, Deferred, CantReproduce, Duplicate, NotBug, ByDesign]

#path = 'C:\\Users\\utukumar\\Desktop\\UsersutukumarDownloadsdocumentSearch_utukumar.xls'

df = pd.read_excel(path, "Data")
text = df.iloc[: , -2]
row_name = 'Labels'
csv_name = '\\frame_weekly.csv'
text.to_csv(Path(save_file+csv_name), index=False, header=True)
data_word_count = pd.read_csv(Path(save_file+csv_name))
d = data_word_count[row_name].str.cat(sep=',')
words = nltk.tokenize.word_tokenize(d)
word_dist = nltk.FreqDist(words)
result = pd.DataFrame(word_dist.most_common(100), columns=['Word', 'Frequency'])
resolution_column = df.iloc[: , -1]
resolution_row_name = "Resolution (string)"
resolution_csv_name = '\\resolution_weekly.csv'
resolution_column.to_csv(Path(save_file+resolution_csv_name), index=False, header=True)
resolution_counter = pd.read_csv(Path(save_file+resolution_csv_name))
reso = resolution_counter[resolution_row_name].str.cat(sep=',')
reso_tags = nltk.tokenize.word_tokenize(reso)
reso_tag_dist = nltk.FreqDist(reso_tags)
reso_result = pd.DataFrame(reso_tag_dist.most_common(100), columns=['Word', 'Frequency'])
reso_dict = reso_result.to_dict('split') 
#print(reso_dict)
freq_dict = result.to_dict('split')
#print(freq_dict['data'])
label_counter_name = '\labels_counter_weekly_report.txt'
#result.to_csv('count.csv', index=False, columns=['Word', 'Frequency'])
excel_worksheet = xlrd.open_workbook(path)
excel_sheet = excel_worksheet.sheet_by_index(0)
#Counting Total Issues
totalDefects = excel_sheet.nrows-1
#("Total Issues = ",excel_sheet.nrows-1,'\n')

#Open and Resolved Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,3)
    if(Column_status == 'Open'):
        openIssues = openIssues+1
    elif(Column_status == 'Resolved'):
        resolvedIssues = resolvedIssues+1
#("Open Issues = ",openIssues)
#("Resolved Issues = ",resolvedIssues) 
#('\n')

#High Medium Low Issues
for i in range(1,excel_sheet.nrows):
    Column_Priority = excel_sheet.cell_value(i,2)
    if ("High" in Column_Priority):
        high = high+1
    if ("Medium" in Column_Priority):
        medium = medium+1
    if ("Low" in Column_Priority):
        low = low+1        
        
#('High Issues = ', high)
#('Medium Issues = ', medium)    
#('Low Issues = ', low)
#('\n')

#Assigned and Unassigned Issues
for i in range(1,excel_sheet.nrows):     
    Column_Assigned = excel_sheet.cell_value(i,5)
    if(len(Column_Assigned)):
        assigned = assigned + 1
    else:
        unassigned = unassigned+1
#("Assigned = ",assigned)
#("Unassigned = ",unassigned)
#('\n')



#Adhoc and TC Issues
for i in range(1,excel_sheet.nrows):
    Column_tags = excel_sheet.cell_value(i,9)
    if((tcTag in Column_tags and tcTagNotRequired not in Column_tags) or (tcTag_new in Column_tags and tcTagNotRequired not in Column_tags)):
        tc = tc+1
    elif(adhocTag in Column_tags or adhocTag_new in Column_tags):
        adhoc = adhoc+1
# print("Adhoc Issues = ",adhoc)
# print("TC Issues = ",tc)  
# print('\n')

#Fixed Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(Fixed in Column_Label or Fixed_new in Column_Resolution):
        fixed = fixed+1
# print("Fixed = ",fixed)
# print('\n')

#Triaged and Untriaged Issued
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    if(Triaged in Column_Label):
        triaged = triaged+1
    elif(Triaged not in Column_Label):
        untriaged = untriaged+1          
#("Triaged = ", triaged)
#("Untriaged = ", untriaged)

#High Medium Low Open and Resolved Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,3)
    Column_Priority = excel_sheet.cell_value(i,2)
    if((Column_status == 'Open' and "High" in Column_Priority)):
          openHigh = openHigh + 1
    elif(Column_status == 'Open' and "Medium" in Column_Priority):
        openMedium = openMedium + 1
    elif(Column_status == 'Open' and "Low" in Column_Priority):
        openLow = openLow + 1
    elif((Column_status == 'Resolved' and "High" in Column_Priority)):
        resolvedHigh = resolvedHigh + 1
    elif(Column_status == 'Resolved' and "Medium" in Column_Priority):
        resolvedMedium = resolvedMedium + 1
    elif(Column_status == 'Resolved' and "Low" in Column_Priority):
        resolvedLow = resolvedLow + 1
        
#("Open High = ", openHigh)
#("Open Medium = ", openMedium)
#("Open Low = ", openLow)
#("Resolved High = ", resolvedHigh)
#("Resolved Medium = ", resolvedMedium)
#("Resolved Low = ", resolvedLow)
#('\n')

#High Medium Low Assigned and Unassigned
for i in range(1,excel_sheet.nrows):
    Column_Assigned = excel_sheet.cell_value(i,5)    
    Column_Priority = excel_sheet.cell_value(i,2)
    if((len(Column_Assigned) and "High" in Column_Priority)):
        assignedHigh = assignedHigh + 1
    elif(len(Column_Assigned) and "Medium" in Column_Priority):
        assignedMedium = assignedMedium + 1
    elif(len(Column_Assigned) and "Low" in Column_Priority):
        assignedLow = assignedLow + 1
    elif((len(Column_Assigned) == 0 and "High" in Column_Priority)):
        unassignedHigh = unassignedHigh + 1
    elif(len(Column_Assigned) == 0 and "Medium" in Column_Priority):
        unassignedMedium = unassignedMedium + 1
    elif(len(Column_Assigned) == 0 and "Low" in Column_Priority):
        unassignedLow = unassignedLow + 1        
#("Assigned High = ",assignedHigh)
#("Assigned Medium = ",assignedMedium)
#("Assigned Low = ",assignedLow)
#("Unassigned High = ",unassignedHigh)
#("Unassigned Medium = ",unassignedMedium)
#("Unassigned Low = ",unassignedLow)
#('\n')                                   

#High Medium Low Triaged and Untriaged Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    if((Triaged in Column_Label and "High" in Column_Priority)):
        triagedHigh = triagedHigh + 1
    elif(Triaged in Column_Label and "Medium" in Column_Priority):
        triagedMedium = triagedMedium + 1
    elif(Triaged in Column_Label and "Low" in Column_Priority):
        triagedLow = triagedLow + 1
    elif((Triaged not in Column_Label and "High" in Column_Priority)):
        untriagedHigh = untriagedHigh + 1
    elif(Triaged not in Column_Label and "Medium" in Column_Priority):
        untriagedMedium = untriagedMedium + 1
    elif(Triaged not in Column_Label and "Low" in Column_Priority):
        untriagedLow = untriagedLow + 1
#("Triaged High = ",triagedHigh)
#("Triaged Medium = ",triagedMedium)
#("Triaged Low = ",triagedLow)
#("Untriaged High = ",untriagedHigh)
#("Untriaged Medium = ",untriagedMedium)
#("Untriaged Low = ",untriagedLow)
#('\n')                            

#High Medium Low Adhoc and TC
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    if((((tcTag in Column_Label and tcTagNotRequired not in Column_Label) or (tcTag_new in Column_Label and tcTagNotRequired not in Column_Label)) and "High" in Column_Priority)):
        tcHigh = tcHigh + 1
    elif((((tcTag in Column_Label and tcTagNotRequired not in Column_Label) or (tcTag_new in Column_Label and tcTagNotRequired not in Column_Label)) and "Medium" in Column_Priority)):
        tcMedium = tcMedium + 1
    elif((((tcTag in Column_Label and tcTagNotRequired not in Column_Label) or (tcTag_new in Column_Label and tcTagNotRequired not in Column_Label)) and "Low" in Column_Priority)):
        tcLow = tcLow + 1
    elif((adhocTag in Column_Label or adhocTag_new in Column_Label) and "High" in Column_Priority):
        adhocHigh = adhocHigh + 1
    elif((adhocTag in Column_Label or adhocTag_new in Column_Label) and "Medium" in Column_Priority):
        adhocMedium = adhocMedium + 1
    elif((adhocTag in Column_Label or adhocTag_new in Column_Label) and "Low" in Column_Priority):
        adhocLow = adhocLow + 1    
# print("Adhoc High = ",adhocHigh)
# print("Adhoc Medium = ",adhocMedium)
# print("Adhoc Low = ",adhocLow)
# print("TC High = ",tcHigh) 
# print("TC Medium = ",tcMedium)
# print("TC Low == ",tcLow)
# print('\n')                           

#High Medium Low Fixed Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(((Fixed in Column_Label or Fixed_new in Column_Resolution) and "High" in Column_Priority)):
        fixedHigh = fixedHigh + 1
    elif((Fixed in Column_Label or Fixed_new in Column_Resolution) and "Medium" in Column_Priority):
        fixedMedium = fixedMedium + 1
    elif((Fixed in Column_Label or Fixed_new in Column_Resolution) and "Low" in Column_Priority):
        fixedLow = fixedLow + 1
# print("Fixed High = ",fixedHigh)
# print("Fixed Medium = ",fixedMedium)
# print("Fixed Low = ",fixedLow)
# print('\n')

#WontFix Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(WontFix in Column_Label or WontFix_new in Column_Resolution):
        wontFix = wontFix + 1
# print("Won't Fix Issue = ",wontFix)
# print('\n')

#High Medium Low Won't Fix Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(((WontFix in Column_status or WontFix_new in Column_Resolution) and "High" in Column_Priority)):
          wontFixHigh = wontFixHigh + 1
    elif((WontFix in Column_status or WontFix_new in Column_Resolution) and "Medium" in Column_Priority):
        wontFixMedium = wontFixMedium + 1
    elif((WontFix in Column_status or WontFix_new in Column_Resolution) and "Low" in Column_Priority):
        wontFixLow = wontFixLow + 1                       
# print("Won't Fix High = ",wontFixHigh)
# print("Won't Fix Medium = ",wontFixMedium)
# print("Won't Fix Low = ",wontFixLow)
# print('\n')

#Deferred Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(Deferred in Column_Label or Deferred_new in Column_Resolution):
        deferred = deferred + 1
# print("Deferred Issue = ",deferred)

#High Medium Low Deferred Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(((Deferred in Column_status or Deferred_new in Column_Resolution) and "High" in Column_Priority)):
          deferredHigh = deferredHigh + 1
    elif((Deferred in Column_status or Deferred_new in Column_Resolution) and "Medium" in Column_Priority):
        deferredMedium = deferredMedium + 1
    elif((Deferred in Column_status or Deferred_new in Column_Resolution) and "Low" in Column_Priority):
        deferredLow = deferredLow + 1                       
# print("Deferred High = ",deferredHigh)
# print("Deferred Medium = ",deferredMedium)
# print("Deferred Low = ",deferredLow)
# print('\n')

#Cannot Reproduce Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(CantReproduce in Column_Label or CantReproduce_new in Column_Resolution):
        cantrepo = cantrepo + 1
# print("Cannot Reproduce Issue = ",cantrepo)

#High Medium Low Cannot Reproduce Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(((CantReproduce in Column_status or CantReproduce_new in Column_Resolution) and "High" in Column_Priority)):
          cantrepoHigh = cantrepoHigh + 1
    elif((CantReproduce in Column_status or CantReproduce_new in Column_Resolution) and "Medium" in Column_Priority):
        cantrepoMedium = cantrepoMedium + 1
    elif((CantReproduce in Column_status or CantReproduce_new in Column_Resolution) and "Low" in Column_Priority):
        cantrepoLow = cantrepoLow + 1                       
# print("Cannot Reproduce High = ",cantrepoHigh)
# print("Cannot Reproduce Medium = ",cantrepoMedium)
# print("Cannot Reproduce Low = ",cantrepoLow)
# print('\n')

#Duplicate Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(Duplicate in Column_Label or Duplicate_new in Column_Resolution):
        dup = dup + 1
# print("Duplicate  = ",dup)

#High Medium Low Duplicate Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if((Duplicate in Column_Label or Duplicate_new in Column_Resolution) and "High" in Column_Priority):
          dupHigh = dupHigh + 1
    elif((Duplicate in Column_Label or Duplicate_new in Column_Resolution) and "Medium" in Column_Priority):
        dupMedium = dupMedium + 1
    elif((Duplicate in Column_Label or Duplicate_new in Column_Resolution) and "Low" in Column_Priority):
        dupLow = dupLow + 1                       
# print("Duplicate  High = ",dupHigh)
# print("Duplicate  Medium = ",dupMedium)
# print("Duplicate  Low = ",dupLow)
# print('\n')

#Not a Bug Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(NotBug in Column_Label or NotBug_new in Column_Resolution):
        notbug = notbug + 1
# print("Not a Bug  = ",notbug)

#High Medium Low Not a Bug Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if((NotBug in Column_status or NotBug_new in Column_Resolution) and "High" in Column_Priority):
          notbugHigh = notbugHigh + 1
    elif((NotBug in Column_status or NotBug_new in Column_Resolution) and "Medium" in Column_Priority):
        notbugMedium = notbugMedium + 1
    elif((NotBug in Column_status or NotBug_new in Column_Resolution) and "Low" in Column_Priority):
        notbugLow = notbugLow + 1                       
# print("Not a Bug  High = ",notbugHigh)
# print("Not a Bug  Medium = ",notbugMedium)
# print("Not a Bug  Low = ",notbugLow)
# print('\n')

#bydesign issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if ByDesign in Column_Label or ByDesign_new in Column_Resolution:
        bydesign = bydesign + 1
# print("By Design Issues = ", bydesign)

#high medium low by design issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if (((ByDesign in Column_Label or ByDesign_new in Column_Resolution) and "High" in Column_Priority)):
        bydesignHigh = bydesignHigh+1
    elif ((ByDesign in Column_Label or ByDesign_new in Column_Resolution) and "Medium" in Column_Priority):
        bydesignMedium = bydesignMedium+1
    elif ((ByDesign in Column_Label or ByDesign_new in Column_Resolution) and "Low" in Column_Priority):
        bydesignLow = bydesignLow+1
# print("By Design High = ",bydesignHigh)
# print("By Design Medium = ",bydesignMedium)
# print("By Design Low = ",bydesignLow)
# print('\n')

# #valid issues
# for i in range(1,excel_sheet.nrows):
#     Column_Label = excel_sheet.cell_value(i,9)
#     Column_Resolution = excel_sheet.cell_value(i,10)
#     if ((ValidLabel in Column_Label)or(Fixed in Column_Label or Fixed_new in Column_Resolution)or(Deferred in Column_Label or Deferred_new in Column_Resolution)or(WontFix in Column_Label or WontFix_new in Column_Resolution)or(ByDesign in Column_Label or ByDesign_new in Column_Resolution)or(CantReproduce in Column_Label or CantReproduce_new in Column_Resolution)):
#         valid = valid+1
# print("Valid Issues = ", valid)

# #high medium low valid issues
# for i in range(1,excel_sheet.nrows):
#     Column_Label = excel_sheet.cell_value(i,9)
#     Column_Priority = excel_sheet.cell_value(i,9)
#     if (ValidLabel in Column_Label and Critical in Column_Priority) or (ValidLabel in Column_Label and Blocker in Column_Priority):
#         validHigh +=1
#     elif ValidLabel in Column_Label and Medium in Column_Priority:
#         validMedium +=1
#     elif ValidLabel in Column_Label and Minor in Column_Priority:
#         validLow +=1
# print("Valid High = ",validHigh)
# print("Valid Medium = ",validMedium)
# print("Valid Low = ",validLow)
# print('\n')
                
#total valid issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if ((Fixed in Column_Label or Fixed_new in Column_Resolution) or (Deferred in Column_Label or Deferred_new in Column_Resolution) or (CantReproduce in Column_Label or CantReproduce_new in Column_Priority) or (WontFix in Column_Label or WontFix_new in Column_Resolution) or (ByDesign in Column_Label or ByDesign_new in Column_Resolution)):
        valid_label += 1
# print("Valid Issues_n = ",valid_label)

#high medium low valid issues:
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    Column_Priority = excel_sheet.cell_value(i,2)
    if ((((Fixed in Column_Label or Fixed_new in Column_Resolution) or (Deferred in Column_Label or Deferred_new in Column_Resolution) or (CantReproduce in Column_Label or CantReproduce_new in Column_Priority) or (WontFix in Column_Label or WontFix_new in Column_Resolution) or (ByDesign in Column_Label or ByDesign_new in Column_Resolution)) and "High" in Column_Priority)):
        valid_labelHigh += 1
    elif ((((Fixed in Column_Label or Fixed_new in Column_Resolution) or (Deferred in Column_Label or Deferred_new in Column_Resolution) or (CantReproduce in Column_Label or CantReproduce_new in Column_Priority) or (WontFix in Column_Label or WontFix_new in Column_Resolution) or (ByDesign in Column_Label or ByDesign_new in Column_Resolution)) and "Medium" in Column_Priority)):
        valid_labelMedium +=1
    elif ((((Fixed in Column_Label or Fixed_new in Column_Resolution) or (Deferred in Column_Label or Deferred_new in Column_Resolution) or (CantReproduce in Column_Label or CantReproduce_new in Column_Priority) or (WontFix in Column_Label or WontFix_new in Column_Resolution) or (ByDesign in Column_Label or ByDesign_new in Column_Resolution)) and "Low" in Column_Priority)):
        valid_labelLow +=1
# print("Valid High_n = ",valid_labelHigh)
# print("Valid Medium_n ",valid_labelMedium)
# print("Valid Low_n = ",valid_labelLow)                      

#total invalid issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if (Duplicate in Column_Label or Duplicate_new in Column_Resolution) or (NotBug in Column_Label or NotBug_new in Column_Resolution):
        invalid_label +=1
#print("Invalid Issues : ",invalid_label)

#high medium low invalid issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    Column_Priority = excel_sheet.cell_value(i,2)
    if (((Duplicate in Column_Label or Duplicate_new in Column_Resolution) or (NotBug in Column_Label or NotBug_new in Column_Resolution)) and "High" in Column_Priority):
        invalid_labelHigh += 1
    elif (((Duplicate in Column_Label or Duplicate_new in Column_Resolution) or (NotBug in Column_Label or NotBug_new in Column_Resolution)) and "Medium" in Column_Priority):
        invalid_labelMedium += 1
    elif (((Duplicate in Column_Label or Duplicate_new in Column_Resolution) or (NotBug in Column_Label or NotBug_new in Column_Resolution)) and "Low" in Column_Priority):
        invalid_labelLow += 1
# print("Invalid High : ",invalid_labelHigh)
# print("Invalid Medium : ",invalid_labelMedium)
# print("Invalid Low : ",invalid_labelLow)                       
                                       
        

'''#IsBug Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    if(IsBug in Column_Label):
        isbug = isbug + 1
#("IsBug  = ",isbug)'''

'''#High Medium Low IsBug Issues
for i in range(1,excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,9)
    if((IsBug in Column_status and Critical in Column_Priority) or (IsBug in Column_status and Blocker in Column_Priority)):
          isbugHigh = isbugHigh + 1
    elif(IsBug in Column_status and Medium in Column_Priority):
        isbugMedium = isbugMedium + 1
    elif(IsBug in Column_status and Minor in Column_Priority):
        isbugLow = isbugLow + 1                       
#("IsBug  High = ",isbugHigh)
#("IsBug  Medium = ",isbugMedium)
#("IsBug  Low = ",isbugLow)
#('\n')'''

#Defect Fix Ratio
if(resolvedIssues == 0):
    defectFixRatio = 0
else:
    defectFixRatio = int((fixed*100)/resolvedIssues)
    
##(defectFixRatio)    
#valid defect ration
if(resolvedIssues == 0):
    validDefectRatio = 0
else:
    validDefectRatio = int((valid_label*100)/resolvedIssues)    
#Defect Priority
defectPriority = int((high*100)/ totalDefects)
##(defectPriority)

#Test Suit Efficiency
suitEfficiency = int((tc*100)/totalDefects) 
##(suitEfficiency)

#Triaged %
triagedRatio = int((triaged*100)/totalDefects)
##(triagedRatio)

get_week = datetime.datetime.now()
#("week"+" - "+get_week.strftime("%U"))
week_of_year = (int(get_week.strftime("%U")))-1
#(week_of_year)
sub = "Weekly SIM metrics - Week {weeknum}".format(weeknum = week_of_year)


try:
    smtpserver = "smtp.amazon.com"
    from_ = input("Enter From Email Address : ")
    to_ = input("Enter To Email Address : ")
    server = smt.SMTP(smtpserver)
    msg = EmailMessage()
    msg['Subject'] = sub
    with open(path, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=path)
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
                    Hello Everyone, <br><br>Please find the below table for week '''+str(week_of_year)+''' bug analysis report and attached excel for SIM details.
                    <br><br>
                    <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                    <tr bgcolor=#D8BFD8>
                    <font size=3 face=Calibri>
                    <th align=center>Bug Split up</th><th align=center>High</th><th align=center>Medium</th><th align=center>Low</th><th align=center>Total</th></font></tr>
                    <tr bgcolor=#D8BFD8>
                    <font size=3 face=Calibri>
                    <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(high)+'''</td><td align=center>'''+str(medium)+'''</td><td align=center>'''+str(low)+'''</td><td align=center>'''+str(totalDefects)+'''</td></font></tr>
                    <tr bgcolor=#98FB98>
                    <font size=3 face=Calibri>
                    <td align=center><b>Resolved</b></td><td align=center>'''+str(resolvedHigh)+'''</td><td align=center>'''+str(resolvedMedium)+'''</td><td align=center>'''+str(resolvedLow)+'''</td><td align=center>'''+str(resolvedIssues)+'''</td></font></tr>
                    <tr bgcolor=#98FB98>
                    <font size=3 face=Calibri>
                    <td align=center><b>Unresolved</b></td><td align=center>'''+str(openHigh)+'''</td><td align=center>'''+str(openMedium)+'''</td><td align=center>'''+str(openLow)+'''</td><td align=center>'''+str(openIssues)+'''</td></font></tr>
                    <tr bgcolor=#FFC0CB>
                    <font size=3 face=Calibri>
                    <td align=center><b>Triaged</b></td><td align=center>'''+str(triagedHigh)+'''</td><td align=center>'''+str(triagedMedium)+'''</td><td align=center>'''+str(triagedLow)+'''</td><td align=center>'''+str(triaged)+'''</td></font></tr>
                    <tr bgcolor=#FFC0CB>
                    <font size=3 face=Calibri>
                    <td align=center><b>Untriaged</b></td><td align=center>'''+str(untriagedHigh)+'''</td><td align=center>'''+str(untriagedMedium)+'''</td><td align=center>'''+str(untriagedLow)+'''</td><td align=center>'''+str(untriaged)+'''</td></font></tr>
                    <tr bgcolor=#B0E0E6>
                    <font size=3 face=Calibri>
                    <td align=center><b>Assigned</b></td><td align=center>'''+str(assignedHigh)+'''</td><td align=center>'''+str(assignedMedium)+'''</td><td align=center>'''+str(assignedLow)+'''</td><td align=center>'''+str(assigned)+'''</td></font></tr>
                    <tr bgcolor=#B0E0E6>
                    <font size=3 face=Calibri>
                    <td align=center><b>Unassigned</b></td><td align=center>'''+str(unassignedHigh)+'''</td><td align=center>'''+str(unassignedMedium)+'''</td><td align=center>'''+str(unassignedLow)+'''</td><td align=center>'''+str(unassigned)+'''</td></font></tr>
                    </table>
                    
                    <br>
                    <font size=3 face=Calibri><b>Resolution Split:</b></font>
                    <br>
                    <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                    <tr bgcolor=#ffffff>
                    <font size=3 face=Calibri>
                    <th align=center>Resolution</th><th align=center>High</th><th align=center>Medium</th><th align=center>Low</th><th align=center>Total</th></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Fixed</b></td><td align=center>'''+str(fixedHigh)+'''</td><td align=center>'''+str(fixedMedium)+'''</td><td align=center>'''+str(fixedLow)+'''</td><td align=center>'''+str(fixed)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Won't Fix</b></td><td align=center>'''+str(wontFixHigh)+'''</td><td align=center>'''+str(wontFixMedium)+'''</td><td align=center>'''+str(wontFixLow)+'''</td><td align=center>'''+str(wontFix)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Cannot Reproduce</b></td><td align=center>'''+str(cantrepoHigh)+'''</td><td align=center>'''+str(cantrepoMedium)+'''</td><td align=center>'''+str(cantrepoLow)+'''</td><td align=center>'''+str(cantrepo)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Duplicate</b></td><td align=center>'''+str(dupHigh)+'''</td><td align=center>'''+str(dupMedium)+'''</td><td align=center>'''+str(dupLow)+'''</td><td align=center>'''+str(dup)+'''</td></font></tr>
                    <!---
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Bug Exist</b></td><td align=center>'''+str(isbugHigh)+'''</td><td align=center>'''+str(isbugMedium)+'''</td><td align=center>'''+str(isbugLow)+'''</td><td align=center>'''+str(isbug)+'''</td></font></tr>
                    --->
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Not a Bug</b></td><td align=center>'''+str(notbugHigh)+'''</td><td align=center>'''+str(notbugMedium)+'''</td><td align=center>'''+str(notbugLow)+'''</td><td align=center>'''+str(notbug)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>By Design</b></td><td align=center>'''+str(bydesignHigh)+'''</td><td align=center>'''+str(bydesignMedium)+'''</td><td align=center>'''+str(bydesignLow)+'''</td><td align=center>'''+str(bydesign)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Valid Issues</b></td><td align=center>'''+str(valid_labelHigh)+'''</td><td align=center>'''+str(valid_labelMedium)+'''</td><td align=center>'''+str(valid_labelLow)+'''</td><td align=center>'''+str(valid_label)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Invalid Issues</b></td><td align=center>'''+str(invalid_labelHigh)+'''</td><td align=center>'''+str(invalid_labelMedium)+'''</td><td align=center>'''+str(invalid_labelLow)+'''</td><td align=center>'''+str(invalid_label)+'''</td></font></tr>
                    </table>
                    
                    <br>
                    <font size=3 face=Calibri><b>Test case Vs. Adhoc:</b></font>
                    <br>
                    <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                    <tr bgcolor=#ffffff>
                    <font size=3 face=Calibri>
                    <th align=center>Priority</th><th align=center>High</th><th align=center>Medium</th><th align=center>Low</th><th align=center>Total</th></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Adhoc</b></td><td align=center>'''+str(adhocHigh)+'''</td><td align=center>'''+str(adhocMedium)+'''</td><td align=center>'''+str(adhocLow)+'''</td><td align=center>'''+str(adhoc)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>TC</b></td><td align=center>'''+str(tcHigh)+'''</td><td align=center>'''+str(tcMedium)+'''</td><td align=center>'''+str(tcLow)+'''</td><td align=center>'''+str(tc)+'''</td></font></tr>
                    </table>
                    
                    <br>
                    <font size=3 face=Calibri><b>OverAll Metrics:</b></font>
                    <br>
                    <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                    <tr bgcolor=#ffffff>
                    <font size=3 face=Calibri>
                    <th align=center>KPI</th><th align=center>Percentage %</th></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Defect Valid Ratio %</b></td><td align=center>'''+str(validDefectRatio)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Defect Fix Ratio %</b></td><td align=center>'''+str(defectFixRatio)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Defect Priority %</b></td><td align=center>'''+str(defectPriority)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Triaged %</b></td><td align=center>'''+str(triagedRatio)+'''</td></font></tr>
                    <tr bgcolor = #ffffff>
                    <font size=3 face=Calibri>
                    <td align=center><b>Test Suite Efficiency %</b></td><td align=center>'''+str(suitEfficiency)+'''</td></font></tr>
                    </table>                    
                </body>
                </html>                
    '''
    
    text_part = MIMEText(email_msg,"html")
    msg.attach(text_part)
    server.sendmail(from_,to_,msg.as_string())
    server.quit()
    print("Email Sent")
except Exception as e:
    print(e)
    
    

with open(Path(save_file+label_counter_name), "w") as f:
    f.write('*****************************************************************************')
    f.write('\n')
    f.write("Created on : "+str(datetime.datetime.now().strftime('%x')))
    f.write('\n')
    f.write('Information'.center(22,"-"))
    f.write('\n')
    f.write('Total Issues are counted as the total number of rows, excluding the first row from the excel')
    f.write('\n')
    f.write('*****************************************************************************')
    f.write('\n')
    f.write('\n')
    for i in (freq_dict['data']):
        for j in labels_list:
            if (i[0] == j ):
                f.write(str(i))
                f.write('\n')
    f.write('-------------------------------------------------------------------------------')
    f.write('\n')
    f.write('\n')
    f.write('Total Issues : '+str(totalDefects)+'\n')
    f.write('Open Issues : '+str(openIssues)+'\n')
    f.write('Open High : '+str(openHigh)+'\n')
    f.write('Open Medium : '+str(openMedium)+'\n')
    f.write('Open Low : '+str(openLow))
    f.write('\n')
    f.write('\n')         
    f.write('Resolved Issues : '+str(resolvedIssues)+'\n')
    f.write('Resolved High : '+str(resolvedHigh)+'\n')
    f.write('Resolved Medium : '+str(resolvedMedium)+'\n')
    f.write('Resolved Low : '+str(resolvedLow))
    f.write('\n')
    f.write('\n')    
    f.write('High : '+str(high)+'\n')
    f.write('Medium : '+str(medium)+'\n')
    f.write('Low : '+str(low))
    f.write('\n')
    f.write('\n')
    f.write('Assigned : '+str(assigned)+'\n')
    f.write('Assigned High : '+str(assignedHigh)+'\n')
    f.write('Assigned Medium : '+str(assignedMedium)+'\n')
    f.write('Assigned Low : '+str(assignedLow))
    f.write('\n')
    f.write('\n')
    f.write('Unassigned : '+str(unassigned)+'\n')
    f.write('Unassigned High : '+str(unassignedHigh)+'\n')
    f.write('Unassigned Medium : '+str(unassignedMedium)+'\n')
    f.write('Unassigned Low : '+str(unassignedLow))
    f.write('\n')
    f.write('\n')
    f.write('Adhoc Issues : '+str(adhoc)+'\n')
    f.write('Adhoc High : '+str(adhocHigh)+'\n')
    f.write('Adhoc Medium : '+str(adhocMedium)+'\n')
    f.write('Adhoc Low : '+str(adhocLow))
    f.write('\n')
    f.write('\n')
    f.write('TC Issues : '+str(tc)+'\n')
    f.write('TC High : '+str(tcHigh)+'\n')
    f.write('TC Medium : '+str(tcMedium)+'\n')
    f.write('TC Low : '+str(tcLow))
    f.write('\n')
    f.write('\n')    
    f.write('Fixed Issues : '+str(fixed)+'\n')
    f.write('Fixed High : '+str(fixedHigh)+'\n')
    f.write('Fixed Medium : '+str(fixedMedium)+'\n')
    f.write('Fixed Low : '+str(fixedLow))
    f.write('\n')
    f.write('\n')
    f.write('Wont Fix : '+str(wontFix)+'\n')
    f.write('Wont Fix High : '+str(wontFixHigh)+'\n')
    f.write('Wont Fix Medium : '+str(wontFixMedium)+'\n')
    f.write('Wont Fix Low : '+str(wontFixLow))
    f.write('\n')
    f.write('\n')
    f.write('Cannot Reproduce Issues : '+str(cantrepo)+'\n')
    f.write('Cannot Reproduce High : '+str(cantrepoHigh)+'\n')
    f.write('Cannot Reproduce Medium : '+str(cantrepoMedium)+'\n')
    f.write('Cannot Reproduce Low : '+str(cantrepoLow))
    f.write('\n')
    f.write('\n')
    f.write('Duplicate Issues : '+str(dup)+'\n')
    f.write('Duplicate High : '+str(dupHigh)+'\n')
    f.write('Duplicate Medium : '+str(dupMedium)+'\n')
    f.write('Duplicate Low : '+str(dupLow))
    f.write('\n')
    f.write('\n')
    f.write('Not A Bug : '+str(notbug)+'\n')
    f.write('Not A Bug High : '+str(notbugHigh)+'\n')
    f.write('Not A Bug Medium : '+str(notbugMedium)+'\n')
    f.write('Not A Bug Low : '+str(notbugLow))
    f.write('\n') 
    f.write('\n')
    f.write('Bydesign Issue : '+str(bydesign)+'\n')
    f.write('Bydesign High : '+str(bydesignHigh)+'\n')
    f.write('Bydesign Medium : '+str(bydesignMedium)+'\n')
    f.write('Bydesign Low : '+str(bydesignLow))
    f.write('\n')
    f.write('\n')
    f.write('Valid Issues : '+str(valid_label)+'\n')
    f.write('Valid High : '+str(valid_labelHigh)+'\n')
    f.write('Valid Medium : '+str(valid_labelMedium)+'\n')
    f.write('Valid Low : '+str(valid_labelLow))
    f.write('\n')
    f.write('\n')
    f.write('Invalid Issues : '+str(invalid_label)+'\n')
    f.write('Invalid High : '+str(invalid_labelHigh)+'\n')
    f.write('Invalid Medium : '+str(invalid_labelMedium)+'\n')
    f.write('Invalid Low : '+str(invalid_labelLow))
    f.write('\n')
    f.write('\n')                
    f.write('-------------------------------------------------------------------------------')
    f.write('\n')
    f.write('\n')
    f.write('Frequency of All the Labels : ')
    f.write('\n')            
    for i in (freq_dict['data']):
        f.write(str(i))
        f.write('\n')
    f.write('-------------------------------------------------------------------------------')
    f.write('\n')
    f.write('Resolution Split')
    f.write('\n')
    f.write('\n')
    for i in (reso_dict['data']):
        f.write(str(i))
        f.write('\n')    
f.close()
print('Labels Counter File generated.')

#deleting the excel sheet after sending the mail    
try:
    os.path.exists(file_path)
    os.remove(path)
    ("Excel File Deleted Successfully")
except Exception as e:
    print(e)
     

try:
    os.path.exists(Path(save_file+csv_name))
    os.remove(Path(save_file+csv_name))
    #("File Deleted Successfully")
except Exception as e:
    print(e)
    
try:
    os.path.exists(Path(save_file+resolution_csv_name))
    os.remove(Path(save_file+resolution_csv_name))
    #("File Deleted Successfully")
except Exception as e:
    print(e)              



                       
    
            
    



