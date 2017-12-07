"""
This code checks if there have been two reports submitted on any given audio
recording.
"""

# Import all required modules for this code
import pandas as pd

# Set the path
path = r'C:\Users\cje4\Desktop\Integrity Study Coding Sheets\Compiled Data'
full_data = '\Full_Data.xlsx'

df_full_data = pd.read_excel(path + full_data)


filename = []
for files in df_full_data['FileName']:
    filenames = files[:-3]
    filename.append(filenames)

single = []
multiple = []
for x in filename:
    if filename.count(x) != 1:
        multiple.append(x)
    else:
        single.append(x)

singles = pd.DataFrame({'Single Coding Sheet Completed': single})
multiples = pd.DataFrame({'Multiple Coding Sheets Completed': multiple})

results = pd.concat([singles, multiples], axis=1)

results.to_excel(path + '\Coding Sheets Count.xlsx')
