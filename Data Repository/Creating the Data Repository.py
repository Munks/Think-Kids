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
