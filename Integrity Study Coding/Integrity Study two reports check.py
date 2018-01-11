"""
This code checks if there have been two reports submitted on any given audio
recording.
"""

# Import all required modules for this code
import pandas as pd
import glob
import datetime as dt

current_date = str(dt.date.today())
# Set the path
path = r'C:\Users\cje4\Desktop\Integrity Study Coding Sheets\Compiled Data'

data_compiled_on = path + r'\Data Compiled on*'

folder = glob.glob(data_compiled_on)
full_data_files = []

for filename in folder:
    full_data_files.append(filename)

current_folder = full_data_files[-1]

folder = glob.glob(current_folder + '\Full_Data*')

current_filename = []
for file in folder:
    current_filename.append(file)

name = current_filename[0]

df_full_data = pd.read_excel(name)


filename = []
for files in df_full_data['FileName']:
    filenames = files[:-3]
    filename.append(filenames)

single = []
two = []
three = []
for x in filename:
    if filename.count(x) == 1:
        single.append(x)
    elif filename.count(x) == 2:
        two.append(x)
    else:
        three.append(x)

singles = pd.DataFrame({'Single Coding Sheet Completed': single})
doubles = pd.DataFrame({'Two Coding Sheets Completed': two})
triples = pd.DataFrame({'Three or more Coding Sheets Completed': three})

dfs = [singles, doubles, triples]
results = pd.concat(dfs, axis=1)

filename = current_folder + '\Coding Sheets Count ' + current_date + '.xlsx'

results.to_excel(filename)
