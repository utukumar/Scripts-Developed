import pandas as pd
import os
import datetime
from tkinter import filedialog
from tkinter.filedialog import askopenfile

#selecting the required file
file_name = filedialog.askopenfile(mode='r')
path = os.path.abspath(file_name.name)

#creating a pandas dataframe
df = pd.read_csv(path, parse_dates=["Created", "Updated", "Resolved"],na_filter=False)
#print(f"{df.iloc[1,25]},{df.iloc[1,26]},{df.iloc[1,27]},{df.iloc[1,28]},{df.iloc[1,29]},{df.iloc[1,30]},{df.iloc[1,31]},{df.iloc[1,32]}".rstrip(','))

#list for creating required dataframe
required_results = []
#extracting the required issues
for i in range(len(df)):
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
    
    required_results.append(result_dict)
    
#creating dataframe with list
required_dataframe = pd.DataFrame(required_results)
required_dataframe.to_excel(r"C:\Users\utukumar\Documents\Scripts\dist\Daily mail AP-QS\DEM-Team\Getting Required Data From CSV\output.xlsx", index=False)