# This code compiles the 'Validating the CPS Integrity Coding System' data

# Import all required modules for this code
import pandas as pd
import numpy as np
import glob
import os

# Set the path
path = r'C:\Users\cje4\Desktop\Integrity Study Coding Sheets'

# Make some folders
TIRF_Data = path + '\TIRF Data'
if not os.path.exists(TIRF_Data):
    os.makedirs(TIRF_Data)

TPOCSA_Data = path + '\TPOCSA Data'
if not os.path.exists(TPOCSA_Data):
    os.makedirs(TPOCSA_Data)

Compiled_Data = path + '\Compiled Data'
if not os.path.exists(Compiled_Data):
    os.makedirs(Compiled_Data)

# Compile the Data
folder = glob.glob(TIRF_Data + '\*.xls')

frames = []

for file in folder:
    filename = file[-23:-9]
    df_TIRF = pd.read_excel(file, header=None)
    df_TIRF.replace({
        'Thoroughness Rating NA  -  1  -  4': np.NaN, '____': np.NaN,
        'NA  -  1  -  2  -  3  -  4': np.NaN, 'NA': np.NaN, '□ _____': np.NaN,
        'N/A': np.NaN, 'na': np.NaN
    }, inplace=True)
    df_TIRF.drop(df_TIRF.index[[2, 3, 16, 19]], inplace=True)
    df_TIRF.drop([1], axis=1, inplace=True)
    df_TIRF = df_TIRF.transpose()
    cell_value = []
    for i in df_TIRF:
        for x in df_TIRF[i]:
            cell_value.append(x)

    data = pd.DataFrame(data=cell_value)
    data = data.transpose()
    data.insert(0, 'FileName', filename)
    to_drop = [
        273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 285, 286, 287
    ]
    for i in to_drop:
        data.drop(i, axis=1, inplace=True)
    for i in range(15):
        data.drop([2 + i], axis=1, inplace=True)
    for i in range(16):
        data.drop([18 + i], axis=1, inplace=True)
    frames.append(data)


frame = pd.concat(frames)

columns = ['FileName', 'Rating_Date', 'Coder', 'Session Length']
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
    'Capture_Integrity', 'Global_Treatment_Integrity_Q',
    'Global_Treatment_Integrity_A', 'Notes'
]
columns.append(end_column_names)
for i in range(240):
    columns.append(str(i))
frame.columns = columns
frame.to_excel(Compiled_Data + '\Data.xlsx')
