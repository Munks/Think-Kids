"""
Cleaning the YOQ Data
"""

# Import all required modules for this code
import pandas as pd

# Create the path and locate the required files
path = r'C:\Users\cje4\Desktop\Data Repository'
yoq_data_loc = r'\2017Dec12_OQACustomReport.csv'

yoq_data = pd.read_csv(path + yoq_data_loc)

yoq_data = yoq_data.dropna(axis=1, how='all')
save = ['PersonID', 'AdministrationID', 'MedRecordNum', 'FirstName',
        'MiddleName', 'LastName']

df_save = yoq_data[save]
df_save = df_save.drop_duplicates(['PersonID'])
mrn = []
for i in df_save['MedRecordNum']:
    if len(i) == 10:
        n = i[3:]
        mrn.append(n)
    else:
        mrn.append(i)
df_save['MedRecordNum'] = mrn
remove = save[1:]
to_delete = [
    'Outpatient', 'Instrument', 'Clinic', 'SettingOfCare', 'InstrumentID'
]
for i in to_delete:
    remove.append(i)

df = []

for i, g in yoq_data.groupby(yoq_data.groupby('PersonID').cumcount()):
    df.append(
        g.drop(remove, 1).add_suffix('_' + str(i)).reset_index(drop=1))

data = pd.concat(df, 1)

new = pd.merge(
    df_save, data, right_on=['PersonID_0'], left_on=['PersonID'],
)
new = new.drop(['PersonID_0'], axis=1)

new.to_csv(path + '\yoq_data.csv')
