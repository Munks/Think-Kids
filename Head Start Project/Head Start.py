import pandas as pd
import numpy as np
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import scoring as score


# Functions
def df_time_columns(label, columns_list):
    for column in df.columns:
        if label in column[:7]:
            name = column[:]
            columns_list.append(name)


path = r'C:\Users\cje4\Desktop\Head Start Project'
file = r'\Head Start Data 11_21_17.xlsx'

df = pd.read_excel(path + file)
df = df.replace({99: np.nan})


# Make the PCRI DataFrames
pcri_df_t1 = ['ID #']
pcri_df_t2 = ['ID #', 'T1_PCRI_REL']
pcri_df_t3 = ['ID #', 'T1_PCRI_REL']

df_time_columns('T1_PCRI', pcri_df_t1)
df_time_columns('T2_PCRI', pcri_df_t2)
df_time_columns('T3_PCRI', pcri_df_t3)

column_names = []
pcri_1 = df[pcri_df_t1]
for column in pcri_1.columns:
    name = column[3:7] + '_' + column[7:]
    column_names.append(name.lower())

pcri_1.columns = column_names

column_names = []
pcri_2 = df[pcri_df_t2]
for column in pcri_2.columns:
    name = column[3:7] + '_' + column[7:]
    column_names.append(name.lower())

pcri_2.columns = column_names

column_names = []
pcri_3 = df[pcri_df_t3]
for column in pcri_3.columns:
    name = column[3:7] + '_' + column[7:]
    column_names.append(name.lower())

pcri_3.columns = column_names

pcri_df_t1.remove('ID #')
pcri_df_t2.remove('ID #')
pcri_df_t2.remove('T1_PCRI_REL')
pcri_df_t3.remove('ID #')
pcri_df_t3.remove('T1_PCRI_REL')
all_pcri = []
all_pcri.append(pcri_df_t1)
all_pcri.append(pcri_df_t2)
all_pcri.append(pcri_df_t3)
for i in all_pcri:
    df.drop(i, axis=1, inplace=True)

# Add the PCRI Results to the main DataFrame
results = []
results = score.pcri(pcri_1, results, '#_', 'pcri__rel', 't1')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_2, results, '#_', 'pcri__rel', 't2')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_3, results, '#_', 'pcri__rel', 't3')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)


# Make the CPS-AIM DataFrames
cps_aim_t1 = ['ID #']
cps_aim_t2 = ['ID #']
cps_aim_t3 = ['ID #']

df_time_columns('T1_TKCO', cps_aim_t1)
df_time_columns('T2_TKCO', cps_aim_t2)
df_time_columns('T3_TKCO', cps_aim_t3)

cps_aim_df_t1 = df[cps_aim_t1]
column_names = []
for column in cps_aim_df_t1.columns:
    name = column[3:8] + '_' + column[8:]
    column_names.append(name.lower())

cps_aim_df_t1.columns = column_names

cps_aim_df_t2 = df[cps_aim_t2]
column_names = []
for column in cps_aim_df_t2.columns:
    name = column[3:8] + '_' + column[8:]
    column_names.append(name.lower())

cps_aim_df_t2.columns = column_names

cps_aim_df_t3 = df[cps_aim_t3]
column_names = []
for column in cps_aim_df_t3.columns:
    name = column[3:8] + '_' + column[8:]
    column_names.append(name.lower())

cps_aim_df_t3.columns = column_names
cps_aim_t1.remove('ID #')
cps_aim_t2.remove('ID #')
cps_aim_t3.remove('ID #')
all_cps_aim = []
all_cps_aim.append(cps_aim_t1)
all_cps_aim.append(cps_aim_t2)
all_cps_aim.append(cps_aim_t3)
for i in all_cps_aim:
    df.drop(i, axis=1, inplace=True)


results = []
results = score.cps_aim(cps_aim_df_t1, results, '#_', 't1')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.cps_aim(cps_aim_df_t2, results, '#_', 't2')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.cps_aim(cps_aim_df_t3, results, '#_', 't3')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)

# Adding the CPS-AIM

df2 = df.mean()

df2 = pd.DataFrame(df2)
df2 = df2.transpose()
df2['Stats'] = 'Mean'

df3 = df.min()

df3 = pd.DataFrame(df3)
df3 = df3.transpose()
df3['Stats'] = 'Min'

df4 = df.max()

df4 = pd.DataFrame(df4)
df4 = df4.transpose()
df4['Stats'] = 'Max'

df5 = df.std()

df5 = pd.DataFrame(df5)
df5 = df5.transpose()
df5['Stats'] = 'Standard Deviation'

frame = [df, df2, df3, df4, df5]

results = pd.concat(frame)

results.to_csv(path + r'\Head Start Data Complied.csv')
