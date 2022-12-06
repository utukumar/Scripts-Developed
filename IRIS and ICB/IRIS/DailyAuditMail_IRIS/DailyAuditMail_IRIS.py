import pandas as pd
from tkinter import filedialog
from tkinter.filedialog import askopenfile
import os
import datetime
import smtplib as smt
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication



# labels and tags used
gdq_label = 'GDQ_QS_Detected_Indian Consumer Business'
adhoc_label = 'QS_Adhoc'
testcase_label = 'QS_Testcase'
valid_label = 'QS_Detected_Valid'
invalid_label = 'QS_Detected_Invalid'

iris_qs_tag = 'Iris_QS'
iris_tc_tag = 'Iris_TC'
iris_adhoc_tag = 'Iris_Adhoc'
iris_triaged_tag = 'Iris-QS-Triaged'

qs_blocker_tag = 'QS_Blocker'
qs_critical_tag = 'QS_Critical'

# mandatory tags for raising issues
qs_regression_tag = 'QS_Regression'
qs_newfeature_tag = 'QS_Newfeature'
qs_nonfunctional_tag = 'QS_Nonfunctional'

# Custom field
iris_resolution_string = 'Fixed'
iris_root_cause = 'Bug - Fixed'

# opening the Excel file
file_name = filedialog.askopenfile(mode='r')
file_path = os.path.abspath(file_name.name)

# reading the required issues Excel sheet as pandas dataframe
df = pd.read_excel(file_path, na_filter=False)

# issue count
total_issues = 0
open_issues = 0
resolved_issues = 0

# calculating issue count
total_issues = len(df)
print("Total Issues : ", total_issues)

for i in range(len(df)):
    if 'Open' in df.iloc[i,3]:
        open_issues += 1
    elif 'Resolved' in df.iloc[i,3]:
        resolved_issues += 1

print("Open Issues : ", open_issues)
print("Resolved Issues : ", resolved_issues)

# list of issues with some missing information
iris_issues_missing_info = []

# list of issues with no missing information
iris_issues_no_missing_info = []

# dictionary of seen issues
iris_seen_dict = {}

# checking if proper adhoc labels are added or not
for i in range(len(df)):
    if iris_adhoc_tag not in df.iloc[i, 9]:
        if adhoc_label in df.iloc[i, 10]:
            if gdq_label in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if iris_tc_tag not in df.iloc[i, 9]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f'{invalid_label} used for Open Issue',
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "Adhoc Required {mis}".format(mis='and ' + "QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f'{valid_label} used for Open Issue',
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "Adhoc Required {mis}".format(mis='and ' + "QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': "NA",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "Adhoc Required {mis}".format(mis='and ' + "QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                elif iris_tc_tag in df.iloc[i, 9]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open Issue and Adhoc Label used, TC label required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open Issue and Adhoc Label used, TC label required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"Adhoc Label used, TC label required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                # updating the issues_id in icb_seen_dict so same issue is not repeated in icb_issues_missing_info
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 141', df.iloc[i, 0])
                # appending the result_dict to icb_issues_missing_info list
                iris_issues_missing_info.append(result_dict)

            elif gdq_label not in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if iris_tc_tag not in df.iloc[i, 9]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open issues. {gdq_label} is Missing",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "Adhoc tag is required {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} used for Open issues. {gdq_label} is Missing",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "Adhoc tag is required {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{gdq_label} is Missing",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "Adhoc tag is required {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }

                elif iris_tc_tag in df.iloc[i, 9]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open Issue. {gdq_label} is Missing. Adhoc used, TC required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} used for Open Issue. {gdq_label} is Missing. Adhoc used, TC required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{gdq_label} is Missing. Adhoc used, TC required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 218', df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)

            elif gdq_label in df.iloc[i, 10] and iris_qs_tag not in df.iloc[i, 9]:
                if iris_tc_tag not in df.iloc[i, 9]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} is used for Open issue.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} and {tag2} are missing {mis}".format(tag1=iris_adhoc_tag, tag2=iris_qs_tag, mis='and '+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} is used for Open issue.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} and {tag2} are missing {mis}".format(tag1=iris_adhoc_tag, tag2=iris_qs_tag, mis='and '+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': "NA",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} and {tag2} are missing {mis}".format(tag1=iris_adhoc_tag, tag2=iris_qs_tag, mis='and '+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                elif iris_tc_tag in df.iloc[i, 9]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} is used for Open issue. Adhoc label used, TC required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis='and '+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} is used for Open issue. Adhoc label used, TC required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis='and '+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"Adhoc label used, TC required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis='and '+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 292', df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)
    elif iris_adhoc_tag in df.iloc[i, 9]:
        if adhoc_label not in df.iloc[i, 10]:
            if gdq_label in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if testcase_label not in df.iloc[i, 10]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} is used for Open issue. Adhoc label is missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} is used for Open issue. Adhoc label is missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"Adhoc label is missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                elif testcase_label in df.iloc[i, 10]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open issue. TC present, Adhoc required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} used for Open issue. TC present, Adhoc required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"TC present, Adhoc required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 367', df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)
            elif gdq_label not in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if testcase_label not in df.iloc[i, 10]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} is used for Open issue. {adhoc_label} and {gdq_label} are missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} is used for Open issue. {adhoc_label} and {gdq_label} are missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{adhoc_label} and {gdq_label} are missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                elif testcase_label in df.iloc[i, 10]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open issue. {gdq_label} is missing. TC Used, Adhoc required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} used for Open issue. {gdq_label} is missing. TC Used, Adhoc required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{gdq_label} is missing. TC Used, Adhoc required.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 440', df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)

            elif gdq_label not in df.iloc[i, 10]  and iris_qs_tag not in df.iloc[i, 9]:
                if testcase_label not in df.iloc[i, 10]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} used for Open issue. {gdq_label} and {adhoc_label} are missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} used for Open issue. {gdq_label} and {adhoc_label} are missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{gdq_label} and {adhoc_label} are missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                elif testcase_label in df.iloc[i, 10]:
                    if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{invalid_label} is used for Open issue. TC label used, Adhoc Required and {gdq_label} is missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"{valid_label} is used for Open issue. TC label used, Adhoc Required and {gdq_label} is missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                    else:
                        result_dict = {
                            'User Alias': df.iloc[i, 4],
                            'Issue URL': df.iloc[i, 8],
                            'Status': df.iloc[i, 3],
                            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                            'Label Present': df.iloc[i, 10],
                            'Label Missing': f"TC label used, Adhoc Required and {gdq_label} is missing.",
                            'Tags Present': df.iloc[i, 9],
                            'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                        }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print("Here 514", df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)
        elif adhoc_label in df.iloc[i, 10]:
            if gdq_label in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{invalid_label} used for Open issue.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 531', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{valid_label} used for Open issue.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 545', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                else:
                    if df.iloc[i, 0] not in iris_seen_dict:
                        if df.iloc[i, 2] not in ['High', 'Medium', 'Low']:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': "NA",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                            if df.iloc[i, 0] not in iris_seen_dict:
                                iris_seen_dict[df.iloc[i, 0]] = 1
                            print('Here 565', df.iloc[i, 0])
                            iris_issues_missing_info.append(result_dict)
                    else:
                        if 'Resolved' in df.iloc[i, 3] and (valid_label in df.ilo[i, 10] or invalid_label in df.iloc[i, 10]) and iris_triaged_tag in df.iloc[i, 9]:
                            result_dict = {
                                        'User Alias': df.iloc[i, 4],
                                        'Issue URL': df.iloc[i, 8],
                                        'Status': df.iloc[i, 3],
                                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                        'Label Present': df.iloc[i, 10],
                                        'Label Missing': f"NA",
                                        'Tags Present': df.iloc[i, 9],
                                        'Missing Tag': "NA",
                                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                                    }
                            if df.iloc[i, 0] not in iris_seen_dict:
                                iris_seen_dict[df.iloc[i, 0]] = 1
                            print('Here 997', df.iloc[i, 0])
                            iris_issues_no_missing_info.append(result_dict)
            elif gdq_label not in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{invalid_label} used for Open issue. {gdq_label} is missing.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 580', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{valid_label} used for Open issue. {gdq_label} is missing.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 595', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                else:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{gdq_label} is missing.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 609', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)


# checking for tc issues
for i in range(len(df)):
    if iris_adhoc_tag not in df.iloc[i, 9]:
        if adhoc_label not in df.iloc[i, 10]:
            if testcase_label not in df.iloc[i, 10]:
                if gdq_label in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                    if iris_tc_tag in df.iloc[i, 9]:
                        if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} is used for Open Issue. {testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} is used for Open Issue. {testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        else:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                    elif iris_tc_tag not in df.iloc[i, 9]:
                        if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} used for Open issue. No Adhoc or TC label present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "No Adhoc or TC tag present {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} used for Open issue. No Adhoc or TC label present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "No Adhoc or TC tag present {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        else:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"No Adhoc or TC label present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "No Adhoc or TC tag present {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 689', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                elif gdq_label not in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                    if iris_tc_tag in df.iloc[i, 9]:
                        if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} used for Open issue. {gdq_label} & {testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} used for Open issue. {gdq_label} & {testcase_label} are missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        else:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{gdq_label} & {testcase_label} are missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }

                    elif iris_tc_tag not in df.iloc[i, 9]:
                        if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} used for Open Issue. {gdq_label}is missing & No Adhoc or TC label present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "No Adhoc or TC tag present {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} used for Open Issue. {gdq_label}is missing & No Adhoc or TC label present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "No Adhoc or TC tag present {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        else:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{gdq_label}is missing & No Adhoc or TC label present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "No Adhoc or TC tag present {mis}".format(mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 763', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                elif gdq_label in df.iloc[i, 10] and iris_qs_tag not in df.iloc[i, 9]:
                    if iris_tc_tag in df.iloc[i, 9]:
                        if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} used for Open issue. {testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} used for Open issue. {testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        else:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{testcase_label} is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                    elif iris_tc_tag not in df.iloc[i, 9]:
                        if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} used for Open issue. No Adhoc or TC label is present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing. No Adhoc or TC tag is present  {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} used for Open issue. No Adhoc or TC label is present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing. No Adhoc or TC tag is present  {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                        else:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"No Adhoc or TC label is present",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing. No Adhoc or TC tag is present  {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 836', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)

# checking if valid or invalid label is used for open issues with tc_tag and tc_label
for i in range(len(df)):
    if df.iloc[i, 0] not in iris_seen_dict:
        if testcase_label in df.iloc[i, 10] and iris_tc_tag in df.iloc[i, 9]:
            # checking for gdq_label is present and qs_tag is not present
            if gdq_label in df.iloc[i, 10] and iris_qs_tag not in df.iloc[i, 9]:
                if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                    result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{invalid_label} used for Open issue.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{valid_label} used for Open issue.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                else:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"NA",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{tag1} is missing {mis}".format(tag1=iris_qs_tag, mis="and "+"QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 880', df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)
            elif gdq_label not in df.iloc[i, 10] and df.iloc[i, 10]:
                if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{invalid_label} used for Open issue. {gdq_label} is missing.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{valid_label} used for Open issue. {gdq_label} is missing.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                else:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{gdq_label} is missing.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                if df.iloc[i, 0] not in iris_seen_dict:
                    iris_seen_dict[df.iloc[i, 0]] = 1
                print('Here 918', df.iloc[i, 0])
                iris_issues_missing_info.append(result_dict)
            elif gdq_label in df.iloc[i, 10] and iris_qs_tag in df.iloc[i, 9]:
                if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{invalid_label} used for Open issue.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 934', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
                    result_dict = {
                        'User Alias': df.iloc[i, 4],
                        'Issue URL': df.iloc[i, 8],
                        'Status': df.iloc[i, 3],
                        'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                        'Label Present': df.iloc[i, 10],
                        'Label Missing': f"{valid_label} used for Open issue.",
                        'Tags Present': df.iloc[i, 9],
                        'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "NA"),
                        'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                    }
                    if df.iloc[i, 0] not in iris_seen_dict:
                        iris_seen_dict[df.iloc[i, 0]] = 1
                    print('Here 949', df.iloc[i, 0])
                    iris_issues_missing_info.append(result_dict)
                else:
                    if df.iloc[i, 0] not in iris_seen_dict:
                        if df.iloc[i, 2] not in ["High", "Medium", "Low"]:
                            result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"NA",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "NA",
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
                            if df.iloc[i, 0] not in iris_seen_dict:
                                iris_seen_dict[df.iloc[i, 0]] = 1
                            print('Here 984', df.iloc[i, 0])
                            iris_issues_missing_info.append(result_dict)
                        else:
                            if ('Resolved' in df.iloc[i, 3]) and (valid_label in df.iloc[i, 10] or invalid_label in df.iloc[i, 10]) and iris_triaged_tag in df.iloc[i, 9]:
                                result_dict = {
                                    'User Alias': df.iloc[i, 4],
                                    'Issue URL': df.iloc[i, 8],
                                    'Status': df.iloc[i, 3],
                                    'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                    'Label Present': df.iloc[i, 10],
                                    'Label Missing': f"NA",
                                    'Tags Present': df.iloc[i, 9],
                                    'Missing Tag': "NA",
                                    'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                                }
                                if df.iloc[i, 0] not in iris_seen_dict:
                                    iris_seen_dict[df.iloc[i, 0]] = 1
                                print('Here 1000', df.iloc[i, 0])
                                iris_issues_no_missing_info.append(result_dict)
                # have to write code for resolved issues
            # if gdq_label in df.iloc[i, 10] and icb_qs_tag in df.iloc[i, 9]:
            #     if 'Open' in df.iloc[i, 3] and invalid_label in df.iloc[i, 10]:
            #         result_dict = {
            #             'User Alias': df.iloc[i, 4],
            #             'Issue URL': df.iloc[i, 8],
            #             'Status': df.iloc[i, 3],
            #             'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
            #             'Label Present': df.iloc[i, 10],
            #             'Label Missing': f"{invalid_label} used for Open issue.",
            #             'Tags Present': df.iloc[i, 9],
            #             'Missing Tag': f"NA"
            #         }
            #     elif 'Open' in df.iloc[i, 3] and valid_label in df.iloc[i, 10]:
            #         result_dict = {
            #             'User Alias': df.iloc[i, 4],
            #             'Issue URL': df.iloc[i, 8],
            #             'Status': df.iloc[i, 3],
            #             'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
            #             'Label Present': df.iloc[i, 10],
            #             'Label Missing': f"{valid_label} used for Open issue.",
            #             'Tags Present': df.iloc[i, 9],
            #             'Missing Tag': f"NA"
            #         }
            #    put in list
            # elif no gdq but icb qs tag

# checking for Resolved issues
for i in range(len(df)):
    if 'Resolved' in df.iloc[i, 3]:
        if valid_label not in df.iloc[i, 10] and invalid_label not in df.iloc[i, 10]:
            if iris_triaged_tag not in df.iloc[i, 9]:
                if iris_root_cause in df.iloc[i,12] and iris_resolution_string not in df.iloc[i,11]:
                    result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} or {invalid_label} label is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing {mis}".format(tag1 = iris_triaged_tag, mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':f'{iris_resolution_string} Resolution is missing.'
                            }
                elif iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,11]:
                    result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} or {invalid_label} label is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{tag1} is missing {mis}".format(tag1 = iris_triaged_tag, mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
            elif iris_triaged_tag in df.iloc[i, 9]:
                if iris_root_cause in df.iloc[i,12] and iris_resolution_string not in df.iloc[i,11]:
                    result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} or {invalid_label} label is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':f'{iris_resolution_string} Resolution is missing.'
                            }
                elif iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,11]:
                    result_dict = {
                                'User Alias': df.iloc[i, 4],
                                'Issue URL': df.iloc[i, 8],
                                'Status': df.iloc[i, 3],
                                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                                'Label Present': df.iloc[i, 10],
                                'Label Missing': f"{valid_label} or {invalid_label} label is missing.",
                                'Tags Present': df.iloc[i, 9],
                                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9]  and 'QS_Nonfunctional' not in df.iloc[i, 9]  else "").rstrip(),
                                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
                            }
            if df.iloc[i, 0] not in iris_seen_dict:
                iris_seen_dict[df.iloc[i, 0]] = 1
            print('Here 1057', df.iloc[i, 0])
            iris_issues_missing_info.append(result_dict)

# checking if proper issues have priority or not
for i in range(len(df)):
    if df.iloc[i, 0] not in iris_seen_dict:
        if df.iloc[i, 2] not in ["High", "Medium", "Low"]:
            result_dict = {
                'User Alias': df.iloc[i, 4],
                'Issue URL': df.iloc[i, 8],
                'Status': df.iloc[i, 3],
                'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
                'Label Present': df.iloc[i, 10],
                'Label Missing': f"NA",
                'Tags Present': df.iloc[i, 9],
                'Missing Tag': "{mis}".format(mis="QS_Regression or QS_Newfeature or QS_Nonfunctional is missing." if 'QS_Regression' not in df.iloc[i, 9] and 'QS_Newfeature' not in df.iloc[i, 9] and 'QS_Nonfunctional' not in df.iloc[i, 9] else "NA"),
                'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
            }
            if df.iloc[i, 0] not in iris_seen_dict:
                print('Here 1027', df.iloc[i, 0])
            iris_issues_missing_info.append(result_dict)

# appending all the proper issues in the no missing info list
for i in range(len(df)):
    if df.iloc[i, 0] not in iris_seen_dict:
        result_dict = {
            'User Alias': df.iloc[i, 4],
            'Issue URL': df.iloc[i, 8],
            'Status': df.iloc[i, 3],
            'Priority': df.iloc[i, 2] if df.iloc[i, 2] else 'Please Add Priority',
            'Label Present': df.iloc[i, 10],
            'Label Missing': f"NA",
            'Tags Present': df.iloc[i, 9],
            'Missing Tag': "NA",
            'Custom Resolution':'{scenario}'.format(scenario = 'Fixed Resolution is present with Bug-Fixed Root Cause' if (iris_root_cause in df.iloc[i,12] and iris_resolution_string in df.iloc[i,12]) else "NA")
        }

        iris_issues_no_missing_info.append(result_dict)

missing_df = pd.DataFrame(iris_issues_missing_info)
no_missing_info_df = pd.DataFrame(iris_issues_no_missing_info)

#missing_df.to_excel(r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\indianConsumerBusinnes\icb\sortingRequiredIssues\missing.xlsx', index=False)
#no_missing_info_df.to_excel(r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\indianConsumerBusinnes\icb\sortingRequiredIssues\no_missing.xlsx', index=False)

missing_info_html = missing_df.to_html(justify="center", index=False)
no_missing_info_html = no_missing_info_df.to_html(justify="center", index=False)

missing_info_html = "<h4>Rectification Required</h4>"+missing_info_html
no_missing_info_html = "<h4>All Good</h4>"+no_missing_info_html

# getting today's date for email subject
today_date = datetime.datetime.now().strftime('%x')

# email subject line
sub = f'Bug Audit Report for {today_date}.'

# generating email
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
    with open(file_path, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=file_path)
    if iris_issues_missing_info and iris_issues_no_missing_info:
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for ''' + str(today_date) + ''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>''' + str(
            total_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>''' + str(resolved_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>''' + str(open_issues) + '''</td></font></tr>
                        </table>
                        <br><br>
                        {missing_info}
                        <br><br>
                        {proper_info}
                    '''.format(missing_info=missing_info_html, proper_info=no_missing_info_html)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_, to_, msg.as_string())
        server.quit()
    elif iris_issues_missing_info:
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for ''' + str(today_date) + ''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>''' + str(
            total_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>''' + str(resolved_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>''' + str(open_issues) + '''</td></font></tr>
                        </table>
                        <br><br>
                        {missing_info}
                        <br>
                    '''.format(missing_info=missing_info_html)
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_, to_, msg.as_string())
        server.quit()
    elif iris_issues_no_missing_info:
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for ''' + str(today_date) + ''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>''' + str(
            total_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>''' + str(resolved_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>''' + str(open_issues) + '''</td></font></tr>
                        </table>
                        <br><br>
                        {}
                        <br><br>

                    '''.format(no_missing_info_html)
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
                        Hello Everyone, <br><br>Please find the below Bug Audit Report for ''' + str(today_date) + ''' and attached excel for SIM details.
                        <br><br>
                        <table cellspacing=1 cellpadding=6 bgcolor=#000000>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <th align=center>Bug Split up</th><th align=center>Total</th></font></tr>
                        <tr bgcolor=#D8BFD8>
                        <font size=3 face=Calibri>
                        <td align=center><b>Total Incoming Defects</b></td><td align=center>''' + str(total_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Resolved</b></td><td align=center>''' + str(resolved_issues) + '''</td></font></tr>
                        <tr bgcolor=#98FB98>
                        <font size=3 face=Calibri>
                        <td align=center><b>Unresolved</b></td><td align=center>''' + str(open_issues) + '''</td></font></tr>
                        </table>
                        <br><br>
                        There are no issues Raised or Resolved.
                    '''
        part2 = MIMEText(email_msg, "html")
        msg.attach(part2)
        server.sendmail(from_, to_, msg.as_string())
        server.quit()
    print("Email Sent")
except Exception as e:
    print(e)





















