"""Create the Data Repository"""

# Import all required modules for this code
import pandas as pd
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import scoring as score

# Create the path and locate the required files
path = r'C:\Users\cje4\Desktop\Data Repository'


"""REDcap Log"""
redcaplog_data = '\REDCap Log.xlsx'
# Open the REDcap Log
redcaplog = pd.read_excel(path + redcaplog_data)
redcap_log_yoq = pd.read_excel(path + redcaplog_data, sheet_name='YOQ_Link')

"""Child History Form"""
child_history = '\CHF Export Wide 12_06_2017b cleaned.xlsx'
# Open the Child History Form
child_history = pd.read_excel(path + child_history)

# Create a list of columns to drop
to_drop = []

for i in range(160):
    if i < 63:
        x = 'briefp_' + str(i + 1) + '_w0'
        to_drop.append(x)
    if i < 86:
        x = 'brief_' + str(i + 1) + '_w0'
        to_drop.append(x)
    if i < 134:
        x = 'bascps_' + str(i + 1) + '_w0'
        to_drop.append(x)
    if i < 150:
        x = 'bascadol_' + str(i + 1) + '_w0'
        to_drop.append(x)
    x = 'basc_' + str(i + 1) + '_w0'
    to_drop.append(x)

child_history = child_history.drop(to_drop, axis=1)


# Functions
def df_time_columns(label, time, columns_list):
    for column in child_history.columns:
        if label in column[:4] and time in column[-2:]:
            name = column[:]
            columns_list.append(name)


# Make the PCRI DataFrames
pcri_df_t1 = ['redcap_id', 'rel_to_child_w0']
pcri_df_t2 = ['redcap_id', 'rel_to_child_w3']
pcri_df_t3 = ['redcap_id', 'rel_to_child_w6']
pcri_df_t4 = ['redcap_id', 'rel_to_child_w9']
pcri_df_t5 = ['redcap_id', 'rel_to_child_w12']

df_time_columns('pcri', 'w0', pcri_df_t1)
df_time_columns('pcri', 'w3', pcri_df_t2)
df_time_columns('pcri', 'w6', pcri_df_t3)
df_time_columns('pcri', 'w9', pcri_df_t4)
df_time_columns('pcri', '12', pcri_df_t5)

column_names = []
pcri_1 = child_history[pcri_df_t1]
for column in pcri_1.columns:
    name = column[:-3]
    column_names.append(name.lower())

pcri_1.columns = column_names

column_names = []
pcri_2 = child_history[pcri_df_t2]
for column in pcri_2.columns:
    name = column[:-3]
    column_names.append(name.lower())

pcri_2.columns = column_names

column_names = []
pcri_3 = child_history[pcri_df_t3]
for column in pcri_3.columns:
    name = column[:-3]
    column_names.append(name.lower())

pcri_3.columns = column_names

column_names = []
pcri_4 = child_history[pcri_df_t4]
for column in pcri_4.columns:
    name = column[:-3]
    column_names.append(name.lower())

pcri_4.columns = column_names

column_names = []
pcri_5 = child_history[pcri_df_t5]
for column in pcri_5.columns:
    name = column[:-4]
    column_names.append(name.lower())

pcri_5.columns = column_names

pcri_df_t1.remove('redcap_id')
pcri_df_t1.remove('rel_to_child_w0')
pcri_df_t2.remove('redcap_id')
pcri_df_t2.remove('rel_to_child_w3')
pcri_df_t3.remove('redcap_id')
pcri_df_t3.remove('rel_to_child_w6')
pcri_df_t4.remove('redcap_id')
pcri_df_t4.remove('rel_to_child_w9')
pcri_df_t5.remove('redcap_id')
pcri_df_t5.remove('rel_to_child_w12')

all_pcri = []
all_pcri.append(pcri_df_t1)
all_pcri.append(pcri_df_t2)
all_pcri.append(pcri_df_t3)
all_pcri.append(pcri_df_t4)
all_pcri.append(pcri_df_t5)
for i in all_pcri:
    child_history.drop(i, axis=1, inplace=True)

# Add the PCRI Results to the main DataFrame
results = []
results = score.pcri(pcri_1, results, 'redcap', 'rel_to_child', 't1')
child_history = pd.merge(
    child_history, results, how='left', left_on='redcap_id', right_on='id'
)
child_history.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_2, results, 'redcap', 'rel_to_child', 't2')
child_history = pd.merge(
    child_history, results, how='left', left_on='redcap_id', right_on='id'
)
child_history.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_3, results, 'redcap', 'rel_to_child', 't3')
child_history = pd.merge(
    child_history, results, how='left', left_on='redcap_id', right_on='id'
)
child_history.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_4, results, 'redcap', 'rel_to_child', 't4')
child_history = pd.merge(
    child_history, results, how='left', left_on='redcap_id', right_on='id'
)
child_history.drop('id', axis=1, inplace=True)
results = []
results = score.pcri(pcri_5, results, 'redca', 'rel_to_child', 't5')
child_history = pd.merge(
    child_history, results, how='left', left_on='redcap_id', right_on='id'
)
child_history.drop('id', axis=1, inplace=True)

# Create the initial Final Results DataFrame
final_results = pd.merge(redcaplog, child_history, how='outer',
                         left_on=['REDcap ID'], right_on=['redcap_id'])

final_results = final_results.append(redcap_log_yoq)

"""BASC Data"""
basc_data = '\BASC score export 12_11_2017 complete.xlsx'
# Open the BASC Data
basc_data = pd.read_excel(path + basc_data, sheet_name=None)
basc_data = pd.concat(basc_data)

# Create a list of columns to drop
to_drop = []

for i in range(160):
    x = 'ITEM' + str(i + 1)
    to_drop.append(x)

basc_data = basc_data.drop(to_drop, axis=1)

# Add BASC Data to the Final Results DataFrame
final_results = pd.merge(final_results, basc_data, how='outer',
                         left_on=['Child MRN'], right_on=['C_ID'])


"""YOQ Data"""
yoq_data = '\yoq_data.csv'
# Open the YOQ Data
yoq_data = pd.read_csv(path + yoq_data)

# Add YOQ Data to the Final Results DataFrame
final_results = pd.merge(final_results, yoq_data, how='outer',
                         left_on=['Child MRN'], right_on=['MedRecordNum'])


final_results.to_csv(path + r'\Repository Data.csv')
