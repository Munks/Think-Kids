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


def completed_count(data_name):
    current_folder = full_data_files[-1]
    folder = glob.glob(current_folder + '\\' + data_name + '*')
    current_filename = []
    for file in folder:
        current_filename.append(file)

    name = current_filename[0]

    df_full_data = pd.read_excel(name)

    filename = []
    raters = []
    for files in df_full_data['FileName']:
        filenames = files[:-3]
        filename.append(filenames)
        rater = int(files[-2:])
        raters.append(rater)

    single = []
    single_rater = []
    two = []
    two_rater = []
    three = []
    three_rater = []
    for x, y in zip(filename, raters):
        if filename.count(x) == 1:
            single.append(x)
            single_rater.append(y)
        elif filename.count(x) == 2:
            two.append(x)
            two_rater.append(y)
        else:
            three.append(x)
            three_rater.append(y)

    singles = pd.DataFrame({'Single Coding Sheet Completed': single})
    singles['Rater'] = single_rater
    doubles = pd.DataFrame({'Two Coding Sheets Completed': two})
    doubles['Rater'] = two_rater
    triples = pd.DataFrame({'Three or more Coding Sheets Completed': three})
    triples['Rater'] = three_rater

    dfs = [singles, doubles, triples]
    results = pd.concat(dfs, axis=1)
    current_folder = full_data_files[-1] + '\Coding Sheets Count '

    filename = current_folder + current_date + ' ' + data_name

    results.to_excel(filename + '.xlsx')


completed_count('Full_Data')
completed_count('TIRF_Data')
completed_count('TPOCSA_Data')
