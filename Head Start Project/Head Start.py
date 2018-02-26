import pandas as pd
import numpy as np
import datetime as dt
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import scoring as score


# Functions
def df_time_columns(label, columns_list):
    for column in df.columns:
        if label in column[:7]:
            name = column[:]
            columns_list.append(name)


def split_dfs(df, slice_front, slice_back):
    front = int(slice_front)
    back = int(slice_back)
    for column in df.columns:
        name = column[front:back] + '_' + column[back:]
        column_names.append(name.lower())


current_date = str(dt.date.today())

path = r'\\Cifs2\thinkkid$\Research\Chris\Head Start Project\Head Start Project Data'
file = r'\Head Start Data Raw Data.xlsx'

df = pd.read_excel(path + file)
df = df.replace({99: np.nan})

# Make the CPS-AIM DataFrames
cps_aim_t1 = ['ID #']
cps_aim_t2 = ['ID #']
cps_aim_t3 = ['ID #']

df_time_columns('T1_TKCO', cps_aim_t1)
df_time_columns('T2_TKCO', cps_aim_t2)
df_time_columns('T3_TKCO', cps_aim_t3)

cps_aim_df_t1 = df[cps_aim_t1]
column_names = []
split_dfs(cps_aim_df_t1, 3, 8)
cps_aim_df_t1.columns = column_names

cps_aim_df_t2 = df[cps_aim_t2]
column_names = []
split_dfs(cps_aim_df_t2, 3, 8)
cps_aim_df_t2.columns = column_names

cps_aim_df_t3 = df[cps_aim_t3]
column_names = []
split_dfs(cps_aim_df_t3, 3, 8)
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

# Adding the CPS-AIM to the main DataFrame

results = []
results = score.cps_aim_parent(cps_aim_df_t1, results, '#_', 't1')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.cps_aim_parent(cps_aim_df_t2, results, '#_', 't2')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.cps_aim_parent(cps_aim_df_t3, results, '#_', 't3')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)

# Make the DERS DataFrames

# DERS is missing items 24 and 36
ders_t1 = ['ID #']
ders_t2 = ['ID #']
ders_t3 = ['ID #']

df_time_columns('T1_DERS', ders_t1)
df_time_columns('T2_DERS', ders_t2)
df_time_columns('T3_DERS', ders_t3)

ders_df_t1 = df[ders_t1]
column_names = []
split_dfs(ders_df_t1, 3, 7)
ders_col_names = ['#_']
for i in range(len(column_names)):
    if i != 23:
        ders_col_names.append('ders_' + str(i + 1))
ders_df_t1.columns = ders_col_names

ders_df_t2 = df[ders_t2]
column_names = []
split_dfs(ders_df_t2, 3, 7)
ders_col_names = ['#_']
for i in range(len(column_names)):
    if i != 23:
        ders_col_names.append('ders_' + str(i + 1))
ders_df_t2.columns = ders_col_names

ders_df_t3 = df[ders_t3]
column_names = []
split_dfs(ders_df_t3, 3, 7)
ders_col_names = ['#_']
for i in range(len(column_names)):
    if i != 23:
        ders_col_names.append('ders_' + str(i + 1))
ders_df_t3.columns = ders_col_names


ders_t1.remove('ID #')
ders_t2.remove('ID #')
ders_t3.remove('ID #')
all_ders = []
all_ders.append(ders_t1)
all_ders.append(ders_t2)
all_ders.append(ders_t3)
for i in all_ders:
    df.drop(i, axis=1, inplace=True)

# Adding the DERS to the main DataFrame

results = []
results = score.ders(ders_df_t1, results, '#_', 't1')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.ders(ders_df_t2, results, '#_', 't2')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.ders(ders_df_t3, results, '#_', 't3')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)


# Make the PCRI DataFrames
pcri_df_t1 = ['ID #']
pcri_df_t2 = ['ID #', 'T1_PCRI_REL']
pcri_df_t3 = ['ID #', 'T1_PCRI_REL']

df_time_columns('T1_PCRI', pcri_df_t1)
df_time_columns('T2_PCRI', pcri_df_t2)
df_time_columns('T3_PCRI', pcri_df_t3)

column_names = []
pcri_1 = df[pcri_df_t1]
split_dfs(pcri_1, 3, 7)
pcri_1.columns = column_names

column_names = []
pcri_2 = df[pcri_df_t2]
split_dfs(pcri_2, 3, 7)
pcri_2.columns = column_names

column_names = []
pcri_3 = df[pcri_df_t3]
split_dfs(pcri_3, 3, 7)
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
results.drop('id', axis=1, inplace=True)
results.drop('pcri_valid_check_t1', axis=1, inplace=True)
results = []
results = score.pcri(pcri_2, results, '#_', 'pcri__rel', 't2')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_3, results, '#_', 'pcri__rel', 't3')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
df.drop('id', axis=1, inplace=True)


# Adding simple statistics to the DataFrame
df2 = round(df.mean(), 2)

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

df5 = round(df.std(), 2)

df5 = pd.DataFrame(df5)
df5 = df5.transpose()
df5['Stats'] = 'Standard Deviation'

frame = [df, df2, df3, df4, df5]

results = pd.concat(frame)

filename = path + r'\Head Start Data Complied ' + current_date + '.xlsx'
results.to_excel(filename)
print('Results Saved to ===> ' + filename)
