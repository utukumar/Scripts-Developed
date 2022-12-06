from poplib import CR
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

openIssues = 0
resolvedIssues = 0
high = 0
medium = 0
low = 0
openHigh = 0
openMedium = 0
openLow = 0
resolvedHigh = 0
resolvedMedium = 0
resolvedLow = 0
triaged = 0
triagedHigh = 0
triagedMedium = 0
triagedLow = 0
untriaged = 0
untriagedHigh = 0
untriagedMedium = 0
untriagedLow = 0
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
bydesign = 0
bydesignHigh = 0
bydesignMedium = 0
bydesignLow = 0
valid_label = 0
valid_labelHigh = 0
valid_labelMedium = 0
valid_labelLow = 0
invalid_label = 0
invalid_labelHigh = 0
invalid_labelMedium = 0
invalid_labelLow = 0
fix_15 = 0
fix_15_plus = 0
fix_30_plus = 0
un_tri_15 = 0
un_tri_15_plus = 0
un_tri_30_plus = 0


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
WontFix = 'QS_ST_WontFix'
Deferred = 'QS_ST_Deferred'
CantReproduce = 'QS_ST_NotRepo'
CantReproduce_new = 'Cannot_reproduce'
Duplicate = 'QS_ST_Duplicate'
Duplicate_new = 'Duplicate'
ByDesign = 'QS_ST_ByDesign'
ByDesign_new = 'By_Design'
ValidLabel = 'QS_Detected_Valid'
NotBug = 'QS_ST_NotABug'
NotBug_new = 'Not_a_bug'

labels_list = [adhocTag, adhocTag_new, tcTag, tcTag_new, Triaged, Fixed, WontFix, Deferred, CantReproduce, Duplicate, NotBug]

#Selecting the file from windows dialog
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfilename()
path = file_path

#selecting the current directory for saving the file
current_directory = os.getcwd()
save_file = current_directory
print(save_file)

#reading excel sheet
excel_worksheet = xlrd.open_workbook(path)
excel_sheet = excel_worksheet.sheet_by_index(0)
sheet_name = str(excel_worksheet.sheet_names()[0])
column = excel_sheet.ncols
row_name = str(excel_sheet.cell_value(0, column-2))


#excel to pandas data frame
df = pd.read_excel(path, sheet_name, parse_dates=["CreateDate", "ResolvedDate"])
text = df.iloc[:,-2]
csv_name = '\\frame_monthly.csv'
text.to_csv(Path(save_file+csv_name))
data_word_count = pd.read_csv(Path(save_file+csv_name))
d = data_word_count[row_name].str.cat(sep=",")
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
freq_dict = result.to_dict('split')
label_counter_name = '\labels_counter_monthly_report.txt'

#getting CreateDate and ResolvedDate pandas data frames
start_date_column = df[df.columns[6]]
resolve_date_column = df[df.columns[7]]

#getting total defects
total_defects = excel_sheet.nrows-1


#high medium and low issues
for i in range(1,excel_sheet.nrows):
    Column_Priority = excel_sheet.cell_value(i,2)
    if("High" in Column_Priority):
        high = high+1
    elif("Medium" in Column_Priority):
        medium = medium+1
    elif("Low" in Column_Priority):
        low = low+1
        
#open and resolved issues
for i in range(1, excel_sheet.nrows):
    Column_status = excel_sheet.cell_value(i,3)
    if(Column_status == 'Open'):
        openIssues = openIssues+1
    elif(Column_status == 'Resolved'):
        resolvedIssues = resolvedIssues+1    
        
#open and resolved High Medium and Low issues
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

#triaged and untriaged issues
for i in range(1,excel_sheet.nrows):
    Column_label = excel_sheet.cell_value(i,9)
    if(Triaged in Column_label):
        triaged = triaged+1
    elif(Triaged not in Column_label):
        untriaged = untriaged+1
                           
#triaged and untriaged High Medium Low issues
for i in range(1, excel_sheet.nrows):
    Column_label = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)
    if((Triaged in Column_label and "High" in Column_Priority)):
        triagedHigh = triagedHigh+1
    elif(Triaged in Column_label and "Medium" in Column_Priority):
        triagedMedium = triagedMedium+1
    elif(Triaged in Column_label and "Low" in Column_Priority):
        triagedLow = triagedLow+1
    elif((Triaged not in Column_label and "High" in Column_Priority)):
        untriagedHigh = untriagedHigh + 1
    elif(Triaged not in Column_label and "Medium" in Column_Priority):
        untriagedMedium = untriagedMedium + 1
    elif(Triaged not in Column_label and "Low" in Column_Priority):
        untriagedLow = untriagedLow + 1            

#Assigned and Unassigned Issues
for i in range(1,excel_sheet.nrows):     
    Column_Assigned = excel_sheet.cell_value(i,5)
    if(len(Column_Assigned)):
        assigned = assigned + 1
    else:
        unassigned = unassigned+1

#High Medium Low Assigned and Unassigned
for i in range(1,excel_sheet.nrows):
    Column_Assigned = excel_sheet.cell_value(i,5)    
    Column_Priority = excel_sheet.cell_value(i,9)
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

#Adhoc and TC Issues
for i in range(1,excel_sheet.nrows):
    Column_tags = excel_sheet.cell_value(i,9)
    if((tcTag in Column_tags and tcTagNotRequired not in Column_tags) or (tcTag_new in Column_tags and tcTagNotRequired not in Column_tags)):
        tc = tc+1
    elif(adhocTag in Column_tags or adhocTag_new in Column_tags):
        adhoc = adhoc+1
        
#High Medium Low Adhoc and TC
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Priority = excel_sheet.cell_value(i,2)    
    if(((tcTag in Column_tags and tcTagNotRequired not in Column_tags) or (tcTag_new in Column_tags and tcTagNotRequired not in Column_tags)) and "High" in Column_Priority):
        tcHigh = tcHigh + 1
    elif((((tcTag in Column_Label and tcTagNotRequired not in Column_Label) or (tcTag_new in Column_Label and tcTagNotRequired not in Column_Label)) and "Medium" in Column_Priority)):
        tcMedium = tcMedium + 1
    elif((((tcTag in Column_Label and tcTagNotRequired not in Column_Label) or (tcTag_new in Column_Label and tcTagNotRequired not in Column_Label)) and "Low" in Column_Priority)):
        tcLow = tcLow + 1
    elif(((adhocTag in Column_tags or adhocTag_new in Column_tags) and "High" in Column_Priority)):
        adhocHigh = adhocHigh + 1
    elif((adhocTag in Column_tags or adhocTag_new in Column_tags) and "Medium" in Column_Priority):
        adhocMedium = adhocMedium + 1
    elif((adhocTag in Column_tags or adhocTag_new in Column_tags) and "Low" in Column_Priority):
        adhocLow = adhocLow + 1

#Fixed Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(Fixed in Column_Label or Fixed_new in Column_Resolution):
        fixed = fixed+1
        
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

#WontFix Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(WontFix in Column_Label or WontFix_new in Column_Resolution):
        wontFix = wontFix + 1

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

#Deferred Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(Deferred in Column_Label or Deferred_new in Column_Resolution):
        deferred = deferred + 1

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
        
#Cannot Reproduce Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(CantReproduce in Column_Label or CantReproduce_new in Column_Resolution):
        cantrepo = cantrepo + 1

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

#Duplicate Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(Duplicate in Column_Label or Duplicate_new in Column_Resolution):
        dup = dup + 1

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

#Not a Bug Issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if(NotBug in Column_Label or NotBug_new in Column_Resolution):
        notbug = notbug + 1

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
        
#bydesign issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if ByDesign in Column_Label or ByDesign_new in Column_Resolution:
        bydesign = bydesign + 1
#print("By Design Issues = ", bydesign)

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

#total valid issues
for i in range(1,excel_sheet.nrows):
    Column_Label = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if ((Fixed in Column_Label or Fixed_new in Column_Resolution) or (Deferred in Column_Label or Deferred_new in Column_Resolution) or (CantReproduce in Column_Label or CantReproduce_new in Column_Priority) or (WontFix in Column_Label or WontFix_new in Column_Resolution) or (ByDesign in Column_Label or ByDesign_new in Column_Resolution)):
        valid_label += 1
#print("Valid Issues_n = ",valid_label)

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
                
#Defect Fix Ratio
if(resolvedIssues == 0):
    defectFixRatio = 0
else:
    defectFixRatio = int((fixed*100)/resolvedIssues)

#valid defect ration
if(resolvedIssues == 0):
    validDefectRatio = 0
else:
    validDefectRatio = int((valid_label*100)/resolvedIssues)    

#Defect Priority
defectPriority = int((high*100)/ total_defects)


#Test Suit Efficiency
suitEfficiency = int((tc*100)/total_defects) 


#Triaged %
triagedRatio = int((triaged*100)/total_defects)

#time taken to fix defects
for i in range(1,excel_sheet.nrows):
    column_labels = excel_sheet.cell_value(i,9)
    Column_Resolution = excel_sheet.cell_value(i,10)
    if "QS_ST_Fixed" in column_labels or "Fixed" in Column_Resolution:
        diff = resolve_date_column[i-1] - start_date_column[i-1]        
        if diff.days <= 15:
            fix_15 = fix_15 + 1
        elif 16<=diff.days <=30:
            fix_15_plus = fix_15_plus + 1
        elif diff.days> 30:
            fix_30_plus = fix_30_plus + 1
            


#ageing of non-triaged defects
for i in range(1,excel_sheet.nrows):
    Column_Labels = excel_sheet.cell_value(i,9)
    current_date = pd.to_datetime(datetime.date.today())
    if(Triaged not in Column_Labels):
        diff = current_date - start_date_column[i-1]
        if diff.days <= 15:
            un_tri_15 = un_tri_15 + 1
        elif 16<=diff.days<=30:
            un_tri_15_plus = un_tri_15_plus + 1
        elif diff.days>30:
            un_tri_30_plus = un_tri_30_plus + 1



lastmonth = datetime.date.today().replace(day=1) - datetime.timedelta(days=1)
lastmonth_year = lastmonth.strftime("%B %Y")
sub = "Monthly SIM metrics - " + str(lastmonth_year)


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
                    Hello Everyone, <br><br>Please find the below table for week '''+str(lastmonth_year)+''' bug analysis report and attached excel for SIM details.
                    <br><br>
                    <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                    <tr bgcolor=#D8BFD8>
                    <font size=3 face=Calibri>
                    <th align=center>Bug Split up</th><th align=center>High</th><th align=center>Medium</th><th align=center>Low</th><th align=center>Total</th></font></tr>
                    <tr bgcolor=#D8BFD8>
                    <font size=3 face=Calibri>
                    <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(high)+'''</td><td align=center>'''+str(medium)+'''</td><td align=center>'''+str(low)+'''</td><td align=center>'''+str(total_defects)+'''</td></font></tr>
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
                    
                    <br>
    				<font size=3 face=Calibri><b>Time taken to fix defects:</b></font>
    				<br>
    				<table cellspacing=1 cellpadding=6 bgcolor=#000000>
    				<tr bgcolor=#ffffff>
    				<font size=3 face=Calibri>
    				<th align=center>1 - 15 days</th><th align=center>16 - 30 days</th><th align=center>> 30 days</th></font></tr>
    				<tr bgcolor=#ffffff>
    				<font size=3 face=Calibri>
    				<td align=center>'''+str(fix_15)+'''</td><td align=center>'''+str(fix_15_plus)+'''</td><td align=center>'''+str(fix_30_plus)+'''</td>
    				</font>
    				</tr>
    				</table>

                    <br>
    				<font size=3 face=Calibri><b>Ageing of not triaged defects:</b></font>
    				<br>
    				<table cellspacing=1 cellpadding=6 bgcolor=#000000>
    				<tr bgcolor=#ffffff>
    				<font size=3 face=Calibri>
    				<th align=center>1 - 15 days</th><th align=center>16 - 30 days</th><th align=center>> 30 days</th></font></tr>
    				<tr bgcolor=#ffffff>
    				<font size=3 face=Calibri>
    				<td align=center>'''+str(un_tri_15)+'''</td><td align=center>'''+str(un_tri_15_plus)+'''</td><td align=center>'''+str(un_tri_30_plus)+'''</td>
    				</font>
    				</tr>
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
    f.write('Total Issues are counted as the total number of rows, excluding the first row from the excel sheet')
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
    f.write('Total Issues : '+str(total_defects)+'\n')
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
    f.write('Unassigned Low : '+str(untriagedLow))
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
    f.write('By Design Issues : '+str(bydesign)+'\n')
    f.write('By Design High : '+str(bydesignHigh)+'\n')
    f.write('By Design Medium : '+str(bydesignMedium)+'\n')
    f.write('By Design Low : '+str(bydesignLow))
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
    f.write('Invalis Low : '+str(invalid_labelLow))
    f.write('\n')
    f.write('\n')
    f.write('Time taken to Fix Defects, 1-15 Days : '+str(fix_15)+'\n')
    f.write('Time taken to Fix Defects, 16-30 Days : '+str(fix_15_plus)+'\n')
    f.write('Time taken to Fix Defects, 30 Days : '+str(fix_30_plus)+'\n')
    f.write('\n') 
    f.write('\n')
    f.write('Ageing of not Triaged Defects, 1-15 Days : '+str(un_tri_15)+'\n')
    f.write('Ageing of not Triaged Defects, 16-30 Days : '+str(un_tri_15_plus)+'\n')
    f.write('Ageing of not Triaged Defects, 30 Days : '+str(un_tri_30_plus)+'\n')
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
                            