import pandas as pd
import os
from datetime import datetime
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import smtplib as smt
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#variables required
totalIssues = 0
openIssues = 0
resolvedIssues = 0

#label used
gdqTeamLabel = 'GDQ_QS_Detected_DEM'
validLabel = 'QS_Detected_Valid'
invalidLabel = 'QS_Detected_Invalid'
adhocLabel = 'QS_Adhoc'
testcaseLabel = 'QS_Testcase'
adhocTag = 'AD-Hoc'
testcaseTag = 'Test Case'


#read the required file
file_name = filedialog.askopenfile(mode='r')
path = os.path.abspath(file_name.name)

#reading excel as pandas dataframe
df = pd.read_excel(path, parse_dates=["Created", "Resolved"], na_filter=False)


#total number of issues
totalIssues = len(df)

#open and resolved issues
for i in range(len(df)):
    if 'Closed' in df.iloc[i,3] or 'Resolved' in df.iloc[i,3]:
        resolvedIssues += 1
    elif 'Screen' in df.iloc[i,3]:
        openIssues += 1

print('Resolved Issues : ',resolvedIssues)
print('Open Issues : ',openIssues)

# result_dict{
#     "User Alias":df.iloc[i,6],
#     "Issue key":df.iloc[i,1],
#     "Status":df.iloc[i,3],
#     "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
#     "Bug Found in Origin":df.iloc[i,11],
#     "Label Present":df.iloc[i,10],
#     "Label Missing":
    
# }

#issues with missing info
missing_info = []

#issues with no missing info
issues_no_missing_info = []

#dict to track the issue key
missing_dict = {}

#getting the required issues
for i in range(len(df)):
    if 'GDQ_QS_Detected_DEM' not in df.loc[i,'Labels']:
        if adhocTag not in df.iloc[i,11] and testcaseTag not in df.iloc[i,11]:
            if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                result_dict = {
                    "User Alias":df.iloc[i,6],
                    "Issue key":df.iloc[i,1],
                    "Status":df.iloc[i,3],
                    "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                    "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                    "Label Present":df.iloc[i,10],
                    "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {gdqTeamLabel} is missing'
                }
            else:
                result_dict = {
                    "User Alias":df.iloc[i,6],
                    "Issue key":df.iloc[i,1],
                    "Status":df.iloc[i,3],
                    "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                    "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                    "Label Present":df.iloc[i,10],
                    "Label Missing":f'{gdqTeamLabel} is missing'
                }
            if df.iloc[i,1] not in missing_dict:
                missing_dict[df.iloc[i,1]] = 1
            print('Here 94',df.iloc[i,1])
            missing_info.append(result_dict)
        elif adhocTag in df.iloc[i,11]:
            if adhocLabel in df.iloc[i,10]:
                if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {gdqTeamLabel} is missing'
                    }
                else:
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{gdqTeamLabel} is missing'
                    }
                if df.iloc[i,1] not in missing_dict:
                    missing_dict[df.iloc[i,1]] = 1
                print('Here 120',df.iloc[i,1])
                missing_info.append(result_dict)
            
            elif testcaseLabel in df.iloc[i,10]:
                if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue, {gdqTeamLabel} is missing and {testcaseLabel} used {adhocLabel} Required'
                    }
                else:
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{gdqTeamLabel} is missing and {testcaseLabel} used {adhocLabel} Required'
                    }
                if df.iloc[i,1] not in missing_dict:
                    missing_dict[df.iloc[i,1]] = 1
                print('Here 146',df.iloc[i,1])
                missing_info.append(result_dict)
        elif testcaseTag in df.iloc[i,11]:
            if adhocLabel in df.iloc[i,10]:
                if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {gdqTeamLabel} is missing and {adhocLabel} used {testcaseLabel} Required'
                    }
                else:
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{gdqTeamLabel} is missing and {adhocLabel} used {testcaseLabel} Required'
                    }
                if df.iloc[i,1] not in missing_dict:
                    missing_dict[df.iloc[i,1]] = 1
                print('Here 172',df.iloc[i,1])
                missing_info.append(result_dict)
            
            elif testcaseLabel in df.iloc[i,10]:
                if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {gdqTeamLabel} is missing'
                    }
                else:
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{gdqTeamLabel} is missing'
                    }
                if df.iloc[i,1] not in missing_dict:
                    missing_dict[df.iloc[i,1]] = 1
                print('Here 198',df.iloc[i,1])
                missing_info.append(result_dict)
    elif gdqTeamLabel in df.iloc[i,10]:
        if adhocTag not in df.iloc[i,11] and testcaseTag not in df.iloc[i,11]:
            if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                result_dict = {
                    "User Alias":df.iloc[i,6],
                    "Issue key":df.iloc[i,1],
                    "Status":df.iloc[i,3],
                    "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                    "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                    "Label Present":df.iloc[i,10],
                    "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue'
                }
            else:
                result_dict = {
                    "User Alias":df.iloc[i,6],
                    "Issue key":df.iloc[i,1],
                    "Status":df.iloc[i,3],
                    "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                    "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                    "Label Present":df.iloc[i,10],
                    "Label Missing":'NA'
                }
            if df.iloc[i,1] not in missing_dict:
                missing_dict[df.iloc[i,1]] = 1
            print('Here 224',df.iloc[i,1])
            missing_info.append(result_dict)
        elif adhocTag in df.iloc[i,11]:
        #     if adhocLabel in df.iloc[i,10]:
        #         #need to make changes so that missing priority issues are also considered
        #         if ('Closed' not in df.iloc[i,3] or 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
        #             result_dict = {
        #                 "User Alias":df.iloc[i,6],
        #                 "Issue key":df.iloc[i,1],
        #                 "Status":df.iloc[i,3],
        #                 "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
        #                 "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
        #                 "Label Present":df.iloc[i,10],
        #                 "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {gdqTeamLabel} is missing'
        #             }
        #         else:
        #             result_dict = {
        #                 "User Alias":df.iloc[i,6],
        #                 "Issue key":df.iloc[i,1],
        #                 "Status":df.iloc[i,3],
        #                 "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
        #                 "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
        #                 "Label Present":df.iloc[i,10],
        #                 "Label Missing":f'{gdqTeamLabel} is missing'
        #             }
        #         if df.iloc[i,1] not in missing_dict:
        #             missing_dict[df.iloc[i,1]] = 1
        #         #missing_info.append(result_dict)
            
            if testcaseLabel in df.iloc[i,10]:
                if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {testcaseLabel} used {adhocLabel} Required'
                    }
                else:
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{testcaseLabel} used {adhocLabel} Required'
                    }
                if df.iloc[i,1] not in missing_dict:
                    missing_dict[df.iloc[i,1]] = 1
                print('Here 276',df.iloc[i,1])
                missing_info.append(result_dict)
        elif testcaseTag in df.iloc[i,11]:
            if adhocLabel in df.iloc[i,10]:
                if ('Closed' not in df.iloc[i,3] and 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {adhocLabel} used {testcaseLabel} Required'
                    }
                else:
                    result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{adhocLabel} used {testcaseLabel} Required'
                    }
                if df.iloc[i,1] not in missing_dict:
                    missing_dict[df.iloc[i,1]] = 1
                print('Here 302',df.iloc[i,1])
                missing_info.append(result_dict)
            
            # elif testcaseLabel in df.iloc[i,10]:
            #     if ('Closed' not in df.iloc[i,3] or 'Resolved' not in df.iloc[i,3]) and (validLabel in df.iloc[i,10] or invalidLabel in df.iloc[i,10]):
            #         result_dict = {
            #             "User Alias":df.iloc[i,6],
            #             "Issue key":df.iloc[i,1],
            #             "Status":df.iloc[i,3],
            #             "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
            #             "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
            #             "Label Present":df.iloc[i,10],
            #             "Label Missing":f'{validLabel} or {invalidLabel} used for Open Issue and {gdqTeamLabel} is missing'
            #         }
            #     else:
            #         result_dict = {
            #             "User Alias":df.iloc[i,6],
            #             "Issue key":df.iloc[i,1],
            #             "Status":df.iloc[i,3],
            #             "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
            #             "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
            #             "Label Present":df.iloc[i,10],
            #             "Label Missing":f'{gdqTeamLabel} is missing'
            #         }
            #     if df.iloc[i,1] not in missing_dict:
            #         missing_dict[df.iloc[i,1]] = 1
            #     missing_info.append(result_dict)


#checking the resolved issues
for i in range(len(df)):
    #if df.iloc[i,1] not in missing_dict:
        if ('Closed' in df.iloc[i,3]) or ('Resolved' in df.iloc[i,3]):
            if validLabel not in df.iloc[i,10] and invalidLabel not in df.iloc[i,10]:
                result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} is missing'
                    }
            #adding the issues in the list
            elif (validLabel not in df.iloc[i,10] and invalidLabel not in df.iloc[i,10]) and (adhocTag in df.iloc[i,11] and testcaseLabel in df.iloc[i,10]):
                result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} is missing and {testcaseLabel} used {adhocLabel} required'
                    }
            
            elif (validLabel not in df.iloc[i,10] and invalidLabel not in df.iloc[i,10]) and (testcaseTag in df.iloc[i,11] and adhocLabel in df.iloc[i,11]):
                result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} is missing and {adhocLabel} used {testcaseLabel} required'
                    }
            if df.iloc[i,1] not in missing_dict:
                missing_dict[df.iloc[i,1]] = 1
            print('Here 348',df.iloc[i,1])
            missing_info.append(result_dict)

#getting the issues without priority and bug found in origin
for i in range(len(df)):
    if df.iloc[i,1] not in missing_dict:
        if df.iloc[i,4] == "" or df.iloc[i,11] == "":
            result_dict = {
                        "User Alias":df.iloc[i,6],
                        "Issue key":df.iloc[i,1],
                        "Status":df.iloc[i,3],
                        "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                        "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                        "Label Present":df.iloc[i,10],
                        "Label Missing":f'{validLabel} or {invalidLabel} is missing'
                    }
            #adding the issues in the list
            if df.iloc[i,1] not in missing_dict:
                missing_dict[df.iloc[i,1]] = 1
            print('Here 367',df.iloc[i,1])
            missing_info.append(result_dict)

#getting all the issue, in which nothing is missing
for i in range(len(df)):
    if df.iloc[i,1] not in missing_dict:
        result_dict = {
                            "User Alias":df.iloc[i,6],
                            "Issue key":df.iloc[i,1],
                            "Status":df.iloc[i,3],
                            "Priority":df.iloc[i,4] if df.iloc[i,4] else "Please add priority",
                            "Bug Found in Origin":df.iloc[i,11] if df.iloc[i,11] else "Please add 'Bug Found in Origin'.",
                            "Label Present":df.iloc[i,10],
                            "Label Missing":'NA'
                        }
    if df.iloc[i,1] not in missing_dict:
        missing_dict[df.iloc[i,1]] = 1
    print('Here 383',df.iloc[i,1])
    issues_no_missing_info.append(result_dict)


#missing info data frame
missing_info_dataframe = pd.DataFrame(missing_info)

#issue without missing info data frame
issues_no_missing_info_dataframe = pd.DataFrame(issues_no_missing_info)

#saving data frame to excel 
#missing_info_dataframe.to_excel(r"C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\DEM-Team\Daily Mail\missing_info.xlsx", index=False)

#saving no missing dataframe to excel
#issues_no_missing_info_dataframe.to_excel(r"C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\DEM-Team\Daily Mail\issue_no_missing_info.xlsx",index=False)

#converting dataframe to html
missing_info_dataframe_to_html = missing_info_dataframe.to_html(index=False)
missing_info_dataframe_to_html = "<h4>Rectification Required</h4>"+missing_info_dataframe_to_html
issue_no_missing_info_dataframe_html = issues_no_missing_info_dataframe.to_html(index=False)
issue_no_missing_info_dataframe_html = "<h4>All Good</h4>"+issue_no_missing_info_dataframe_html

#getting today's date for e-mail subject
today_date = datetime.now().strftime("%x")

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
    if missing_info and issues_no_missing_info:
        email_msg = '''
                    <html>
                    <head>
                    <title>Weekly SIM Metrics</title>
                    <meta http–equiv="Content-Type" content="text/html; charset=UTF-8" />
                    <meta http–equiv=“X-UA-Compatible” content=“IE=edge” />
                    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
                    <style>
                    table, th, td, tr {border: 1.5px solid black; border-collapse: collapse; padding: 7px; text-align: center}
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
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(totalIssues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>'''+str(resolvedIssues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>'''+str(openIssues)+'''</td></font></tr>
                        </table>
                        <br><br>
                        {missing_info}
                        <br><br>
                        {proper_info}
                    '''.format(missing_info=missing_info_dataframe_to_html, proper_info=issue_no_missing_info_dataframe_html)
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for '''+str(today_date)+''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(totalIssues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>'''+str(resolvedIssues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>'''+str(open)+'''</td></font></tr>
                        </table>
                        <br><br>
                        {missing_info}
                        <br>
                    '''.format(missing_info=missing_info_dataframe_to_html)
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
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>'''+str(totalIssues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>'''+str(resolvedIssues)+'''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>'''+str(openIssues)+'''</td></font></tr>
                        </table>
                        <br><br>
                        {}
                        <br><br>
                        
                    '''.format(issue_no_missing_info_dataframe_html)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_,to_,msg.as_string())
        server.quit()
    print("Email Sent")
except Exception as e:
    print(e)







        
        
        