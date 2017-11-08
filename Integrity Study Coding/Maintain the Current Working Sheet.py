# This code compiles the 'Validating the CPS Integrity Coding System' data

# Import all required modules for this code
import pandas as pd
import numpy as np
import glob
import os

# Set the path
path = r'C:\Users\cje4\Desktop\Integrity Study Coding Sheets'
TIRF_Data = path + '\TIRF Data'
TPOCSA_Data = path + '\TPOCSA Data'
Compiled_Data = path + '\Compiled Data'


# Function to make folders
def make_folders():
    if not os.path.exists(TIRF_Data):
        os.makedirs(TIRF_Data)

    if not os.path.exists(TPOCSA_Data):
        os.makedirs(TPOCSA_Data)

    if not os.path.exists(Compiled_Data):
        os.makedirs(Compiled_Data)


# Making folders
make_folders()

# Compile the TIRF Data
folder = glob.glob(TIRF_Data + '\*.xls')

# Create an empty frame to store data in
frames = []

for file in folder:
    filename = file[-24:-9]
    coder = filename[-2:]
    date = filename[-9:-7] + '/' + filename[-7:-5] + '/20' + filename[-5:-3]
    family = filename[-12:-10]
    specialist = filename[-15:-13]
    df_TIRF = pd.read_excel(file, header=None)

# Replace the blanks with neutral values (np.NaN)
    df_TIRF.replace({
        'Thoroughness Rating NA  -  1  -  4': np.NaN, '____': np.NaN,
        'NA  -  1  -  2  -  3  -  4': np.NaN, 'NA': np.NaN, '□ _____': np.NaN,
        'N/A': np.NaN, 'na': np.NaN
    }, inplace=True)

# Drop out universally blank rows
    df_TIRF.drop(df_TIRF.index[[2, 3, 16, 19]], inplace=True)

# Drop out universally blank columns
    df_TIRF.drop([1], axis=1, inplace=True)

# Clean data sheets that have problems with columns
# 73_89_071817_02 has an extra column at the end of the sheet
    if filename == '73_89_071817_02':
        df_TIRF.drop([18], axis=1, inplace=True)

# reorient the sheet horizontally
    df_TIRF = df_TIRF.transpose()

# Clean data sheets that have problems with rows
# Both 34_18_020817_02 and 31_30_060517_02 have a question after the notes
    if filename == '34_18_020817_02':
        for i in range(11):
            df_TIRF.drop([21 + i], axis=1, inplace=True)
    elif filename == '31_30_060517_02':
        for i in range(14):
            df_TIRF.drop([21 + i], axis=1, inplace=True)

# Grab each cell of the data sheet and string them after eachother
# This puts the data in a single column
    cell_value = []
    for i in df_TIRF:
        for x in df_TIRF[i]:
            cell_value.append(x)

# Create a new data sheet
    data = pd.DataFrame(data=cell_value)

# Make the single column into a single row
    data = data.transpose()

# Create the first few columns
    data[0] = filename
    data[1] = date
    data[2] = int(coder)
    data[3] = int(family)
    data[4] = int(specialist)
    data[5] = 'Session Length'

# Add the new row under the previous row in the final version
    frames.append(data)

# Make the number array into a pandas DataFrame
frame = pd.concat(frames)

# Drop the extra columns in the dataframe
for i in range(50):
    frame.drop([i + 290], axis=1, inplace=True)
for i in range(27):
    frame.drop([i + 7], axis=1, inplace=True)
for i in range(10):
    frame.drop([i + 273], axis=1, inplace=True)
for i in range(3):
    frame.drop([i + 285], axis=1, inplace=True)

# Creating the names for columns
columns = [
    'FileName', 'Date', 'Coder', 'Family', 'Specialist', 'Session Length',
    'Attending'
]

for i in range(14):
    for x in range(17):
        if x == 0:
            header = ('Area_' + str(i + 1))
            columns.append(header)
        elif x == 1:
            header = ('Question_' + str(i + 1))
            columns.append(header)
        elif x != 16:
            header = ('Component_' + str(1 + i) + '_Time_' + str(1 + x))
            columns.append(header)
        else:
            skillfullness = ('Skillfullness_Component_' + str(1 + i))
            columns.append(skillfullness)
end_column_names = [
    'Capture_Integrity_Q', 'Capture_Integrity_A',
    'Global_Treatment_Integrity_Q', 'Global_Treatment_Integrity_A', 'Notes'
]

for i in end_column_names:
    columns.append(i)
# Renaming the columns
frame.columns = columns

# Save the dataframe as an excel file
frame.to_excel(Compiled_Data + '\Data.xlsx')
