"""
Create the Data Repository Enrollment Log
"""

# Import all required modules for this code
import pandas as pd
import numpy as np
from redcap import Project
import sys
sys.path.append('C:\Python Programs')
import tokens as tk

# Create the path and locate the required files
path = r'C:\Users\cje4\Desktop\Data Repository'

# Grab my REDcap Token
tk.child_history_form()

project = Project(tk.api_url, tk.api_token)

df_childhistory = project.export_records(format='df')
columns = ['sex', 'ch_race___1', 'ch_race___2', 'ch_race___3', 'ch_race___4',
           'ch_race___5', 'ch_race___6']
df_childhistory = df_childhistory[columns]
columns = [
    'participant_id', 'Event Name', 'Gender', 'African American or Black',
    'Asian American or Indian American',
    'European American or Caucasian/White', 'Latino(a)/Hispanic',
    'Native American (including Alaskan/Hawaiian Native)',
    'Other (please specify)'
]
df_childhistory.reset_index(
    level=['participant_id', 'redcap_event_name'], inplace=True
)

df_childhistory.columns = columns
df = df_childhistory
df_childhistory = df[df['Event Name'] == 'intake_arm_1']
df_childhistory = df_childhistory.replace({0: np.nan})

# Import the REDCap Log
redcaplog = '\REDCap Log.xlsx'
df_redcaplog = pd.read_excel(path + redcaplog)
# Take only the rows where 'Repository Consentnt?' has some value in it.
df_redcaplog = df_redcaplog.dropna(subset=['Repository Consent?'])
# Take only the following three columns as they are all that is relivent to the
# enrollment log
df_redcaplog = df_redcaplog.filter([
    'REDcap ID', 'Child MRN', 'Repository Consent?', 'Consent Date'
], axis=1)

# Merge the two files into a new dataframe
results = pd.merge(df_childhistory, df_redcaplog, left_on=['participant_id'],
                   right_on=['REDcap ID']
                   )
# Take only the rows where 'Repository Consent?' is filled out 'YES'
results = results[results['Repository Consent?'] == 'YES']
# Identify genders
results['Gender'] = results['Gender'].replace({1: 'Male', 2: 'Female'})
"""
Create a report called 'Enrollment Log Demographics' in our 'Data Repository'
folder on the desktop
"""
results.to_excel(path + '\Enrollment Log Demographics.xlsx')
