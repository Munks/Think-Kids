"""
This code compiles the 'Validating the CPS Integrity Coding System' data
"""

# Import all required modules for this code
import pandas as pd
import numpy as np
import glob
import os
import datetime as dt


# Set the path
path = r'\\Cifs2\thinkkid$\Research\Chris\Youth Villages\Integrity Study Coding Sheets'
tirf_data = path + r'\TIRF Data'
tpocsa_data = path + r'\TPOCSA Data'
compiled_data = path + r'\Compiled Data'
current_date = str(dt.date.today())

folder_date = compiled_data + '\\Data Compiled on ' + current_date

# Make a new folder for today
if not os.path.exists(folder_date):
    os.makedirs(folder_date)

# Part 1
# Compile the TIRF Data
folder = glob.glob(tirf_data + r'\*.xls')

# Create an empty frame to store data in
frames = []

for file in folder:
    # Create file name containers
    filename = file[-25:-9]
    coder = filename[-2:]
    date = filename[-9:-7] + '/' + filename[-7:-5] + '/20' + filename[-5:-3]
    family = filename[-13:-10]
    specialist = filename[-16:-14]
    df_TIRF = pd.read_excel(file, header=None)

    # Drop out universally blank rows
    df_TIRF.drop(df_TIRF.index[[2, 3, 16, 19]], inplace=True)

    # Drop out universally blank columns
    df_TIRF.drop([1], axis=1, inplace=True)

    # Replace the blanks with neutral values (np.NaN)
    df_TIRF.replace({
        'Thoroughness Rating NA  -  1  -  4': np.NaN,
        '____': np.NaN,
        'NA  -  1  -  2  -  3  -  4': np.NaN,
        '□ _____': np.NaN,
        'NA ': np.NaN,
        'NA  ': np.NaN,
        'N/A  ': np.NaN,
        'N/a': np.NaN,
        'n/a': np.NaN,
        'n/A': np.NaN,
        'N/A': np.NaN,
        ' 1  -  2  -  3  -  4': np.NaN,
        'Y / N': np.NaN
    }, inplace=True)

    # reorient the sheet horizontally
    df_TIRF = df_TIRF.transpose()
    count = -1
    for i in df_TIRF[17]:
        if i in [1, 2, 3, 4, 99]:
            count += 1
    session_lengths = {
        -1: np.NaN,
        0: np.NaN,
        1: 5,
        2: 10,
        3: 15,
        4: 20,
        5: 25,
        6: 30,
        7: 35,
        8: 40,
        9: 45,
        10: 50,
        11: 55,
        12: 60,
        13: 65,
        14: 70
    }
    time = session_lengths[count]
    # Grab each cell of the data sheet and string them after eachother
    # Put the data in a single column
    cell_value = []
    for i in df_TIRF:
        for x in df_TIRF[i]:
            cell_value.append(x)

    # Create a new data sheet
    data = pd.DataFrame(data=cell_value)

    # Make the single column into a single row
    data = data.transpose()

    capture_integrity = []

    if data[277].any() and data[278].any() and data[281].any():
        capture_integrity.append('')
    elif 'yes' in str(data[277]).lower():
        capture_integrity.append('Yes')
    elif 'no' in str(data[278]).lower():
        capture_integrity.append('No')
    elif 'maybe' in str(data[281]).lower():
        capture_integrity.append('Maybe')
    else:
        capture_integrity.append('')

    rightmost_column_count = 0

    for i in range(50, 221, 17):
        if data[i].any():
            rightmost_column_count += 1

    number_of_columns_with_cps = 0
    column_1 = 0
    for i in range(36, 207, 17):
        if data[i].any():
            column_1 += 1
    if column_1 > 0:
        number_of_columns_with_cps += 1
    column_2 = 0
    for i in range(37, 208, 17):
        if data[i].any():
            column_2 += 1
    if column_2 > 0:
        number_of_columns_with_cps += 1
    column_3 = 0
    for i in range(38, 209, 17):
        if data[i].any():
            column_3 += 1
    if column_3 > 0:
        number_of_columns_with_cps += 1
    column_4 = 0
    for i in range(39, 210, 17):
        if data[i].any():
            column_4 += 1
    if column_4 > 0:
        number_of_columns_with_cps += 1
    column_5 = 0
    for i in range(40, 211, 17):
        if data[i].any():
            column_5 += 1
    if column_5 > 0:
        number_of_columns_with_cps += 1
    column_6 = 0
    for i in range(41, 212, 17):
        if data[i].any():
            column_6 += 1
    if column_6 > 0:
        number_of_columns_with_cps += 1
    column_7 = 0
    for i in range(42, 213, 17):
        if data[i].any():
            column_7 += 1
    if column_7 > 0:
        number_of_columns_with_cps += 1
    column_8 = 0
    for i in range(43, 214, 17):
        if data[i].any():
            column_8 += 1
    if column_8 > 0:
        number_of_columns_with_cps += 1
    column_9 = 0
    for i in range(44, 215, 17):
        if data[i].any():
            column_9 += 1
    if column_9 > 0:
        number_of_columns_with_cps += 1
    column_10 = 0
    for i in range(45, 216, 17):
        if data[i].any():
            column_10 += 1
    if column_10 > 0:
        number_of_columns_with_cps += 1
    column_11 = 0
    for i in range(46, 217, 17):
        if data[i].any():
            column_11 += 1
    if column_11 > 0:
        number_of_columns_with_cps += 1
    column_12 = 0
    for i in range(47, 218, 17):
        if data[i].any():
            column_12 += 1
    if column_12 > 0:
        number_of_columns_with_cps += 1
    column_13 = 0
    for i in range(48, 219, 17):
        if data[i].any():
            column_13 += 1
    if column_13 > 0:
        number_of_columns_with_cps += 1
    column_14 = 0
    for i in range(49, 220, 17):
        if data[i].any():
            column_14 += 1
    if column_14 > 0:
        number_of_columns_with_cps += 1

    # Create the first few columns
    data[0] = filename
    data[2] = data[1]
    data[1] = date
    data[3] = int(coder)
    data[5] = int(family)
    data[6] = int(specialist)
    data[7] = time
    data[8] = data[21]
    data[9] = data[26]
    data[10] = data[29]
    data[11] = data[32]

    data[12] = data[277]
    data[13] = data[278]
    data[14] = data[281]
    data[15] = capture_integrity
    data['rightmost_column_count'] = rightmost_column_count
    data['number_of_columns_with_cps'] = number_of_columns_with_cps
    time_columns = time / 5
    data['percent_of_cps_time'] = round(
        number_of_columns_with_cps / time_columns, 2)

    # Add the new row under the previous row in the final version
    frames.append(data)

# Make the number array into a pandas DataFrame
frame = pd.concat(frames)

# Drop the extra columns in the dataframe
for i in range(50):
    frame.drop([i + 290], axis=1, inplace=True)
for i in range(18):
    frame.drop([i + 16], axis=1, inplace=True)
for i in range(12):
    frame.drop([i + 272], axis=1, inplace=True)
for i in range(3):
    frame.drop([i + 285], axis=1, inplace=True)

# Creating the names for columns
columns = [
    'FileName', 'Date', 'Rating Date', 'Coder', 'Raters Name', 'Family',
    'Specialist', 'Session Length Calculated', 'Session Length Reported',
    'Caregiver Attending', 'Youth Attending', 'Other Attending',
    'Capture Integrity Yes', 'Capture Integrity No', 'Capture Integrity Maybe',
    'Captured Integrity'
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
            header = ('Component_' + str(i + 1) + '_Time_' + str(x - 1))
            columns.append(header)
        else:
            skillfullness = ('Skillfullness_Component_' + str(1 + i))
            columns.append(skillfullness)
end_column_names = [
    'Global_Treatment_Integrity_Q', 'Global_Treatment_Integrity_A', 'Notes',
    'rightmost_column_count', 'number_of_columns_with_cps',
    'percent_of_cps_time'
]
count = []
c = -1
for i in frame.columns:
    c += 1
    count.append(c)
frame.columns = count

frame = frame[count[:260]]

for i in end_column_names:
    columns.append(i)

# Renaming the columns
frame.columns = columns

end_file_name = ' ' + current_date + '.xlsx'
# Save the result to an excel
frame.to_excel(folder_date + r'\TIRF_Data' + end_file_name)

# Part 2
# Compile the TPOCSA Data
folder = glob.glob(tpocsa_data + r'\*.xlsx')

# Create an empty frame to store data in
frames = []

for file in folder:
    # Create file name containers
    filename = file[-28:-12]
    coder = filename[-2:]
    date = filename[-9:-7] + '/' + filename[-7:-5] + '/20' + filename[-5:-3]
    family = filename[-13:-10]
    specialist = filename[-16:-14]

    # Open the TPOCSA Files
    df_TPOCSA = pd.read_excel(file, header=None)

    # Import the data into a dataframe
    data = pd.DataFrame(data=df_TPOCSA[7])

    # Make the single column into a single row
    data = data.transpose()

    # Create the first few columns
    data[0] = filename
    data[1] = date
    data[2] = int(coder)
    data[3] = int(family)
    data[4] = int(specialist)
    frames.append(data)
frame = pd.concat(frames)

# Drop empty columns
frame.drop([5, 6, 11, 12, 13, 14, 15, 21], axis=1, inplace=True)

# Creating the names for columns
columns = ['FileName', 'Date', 'Coder', 'Family', 'Specialist']

for i in range(10):
    if i < 4:
        header = ('TPOC_' + str(1 + i))
        columns.append(header)
    elif 3 < i < 9:
        header = ('TPOCAR_' + str(i - 3))
        columns.append(header)
    elif i > 8:
        header = ('TPOCBR_6')
        columns.append(header)

# Rename columns
frame.columns = columns


# Save the result to an excel
frame.to_excel(folder_date + r'\TPOCSA_Data' + end_file_name)

# Part 3
# Import the TIRF and TPOCSA data
df1 = pd.read_excel(folder_date + r'\TIRF_Data' + end_file_name)
df2 = pd.read_excel(folder_date + r'\TPOCSA_Data' + end_file_name)

# Merge the TIRF and TPOCSA data
result = pd.merge(df1, df2, how='outer', on=[
    'FileName', 'Date', 'Coder', 'Family', 'Specialist'
])

# Save the result to an excel
result.to_excel(folder_date + r'\Full_Data' + end_file_name)


"""
This section of the code checks if there have been two reports submitted on any
given audio recording.
"""


def completed_count(data_name, df):
    current_folder = folder_date
    filename = []
    raters = []
    for files in df['FileName']:
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
    current_folder = folder_date + r'\Coding Sheets Count '

    filename = current_folder + current_date + ' ' + data_name

    results.to_excel(filename + '.xlsx')


completed_count('Full_Data', result)
completed_count('TIRF_Data', df1)
completed_count('TPOCSA_Data', df2)
