from uuid import NAMESPACE_X500
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import askopenfile

#path = os.getcwd()
#creating pop-up open window for selecting the downloaded excel sheet
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)
file_path = filedialog.askopenfile(mode = 'r')

#assigning the absolute path to the path variable
path = os.path.abspath(file_path.name)
df = pd.read_csv(path, parse_dates=['Created','Resolved'], na_filter=False)

prefix = "https://jira.music.amazon.dev/browse/"

required_result = []

for i in range(len(df)):
    result = {
                'Issue key':prefix+df.loc[i,'Issue key'],
                'Summary':f'{df.iloc[i,0]}',
                'Priority':f'{df.iloc[i,11]}',
                'Status':f'{df.iloc[i,4]}',
                'Reporter':f'{df.iloc[i,14]}',
                'Assignee':f'{df.iloc[i,13]}',
                'CreateDate':df.iloc[i,16],
                'ResolvedDate':df.iloc[i,19],
                'Having Testcase?':df.iloc[i,102],
                'Invalid Bug - Category':df.iloc[i,104],
                'Labels':f'{df.iloc[i,25]},{df.iloc[i,26]},{df.iloc[i,27]},{df.iloc[i,28]},{df.iloc[i,29]},{df.iloc[i,30]},{df.iloc[i,31]}'.rstrip(',')             
                
               }
    required_result.append(result)

required_result_df = pd.DataFrame(required_result)
required_result_df.to_excel(r'C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\Mercury\required.xlsx', index=False)


