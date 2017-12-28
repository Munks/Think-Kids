import pandas as pd
import numpy as np
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import scoring as score

path = r'C:\Users\cje4\Desktop\Head Start Project'
file = r'/Head Start Data 11_21_17.xlsx'

df = pd.read_excel(path + file)
df = df.replace({99: np.nan})


# Make the PCRI DataFrames
pcri_df_t1 = ['ID #']
pcri_df_t2 = ['ID #', 'T1_PCRI_REL']
pcri_df_t3 = ['ID #', 'T1_PCRI_REL']

for column in df.columns:
    if 'T1_PCRI' in column:
        name = column[:]
        pcri_df_t1.append(name)
    if 'T2_PCRI' in column:
        name = column[:]
        pcri_df_t2.append(name)
    if 'T3_PCRI' in column:
        name = column[:]
        pcri_df_t3.append(name)

pcri_1 = df[pcri_df_t1]
column_names = []
for column in pcri_1.columns:
    name = column[3:7] + '_' + column[7:]
    column_names.append(name.lower())

pcri_1.columns = column_names

pcri_2 = df[pcri_df_t2]
column_names = []
for column in pcri_2.columns:
    name = column[3:7] + '_' + column[7:]
    column_names.append(name.lower())

pcri_2.columns = column_names

pcri_3 = df[pcri_df_t3]
column_names = []
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
results = []
results = score.pcri(pcri_2, results, '#_', 'pcri__rel', 't2')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
results = []
results = score.pcri(pcri_3, results, '#_', 'pcri__rel', 't3')
df = pd.merge(df, results, how='left', left_on='ID #', right_on='id')
# Score the PCRI DataFrames


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


# Adding the CPS-AIM

results.to_csv(path + '/Head Start Data Complied.csv')
