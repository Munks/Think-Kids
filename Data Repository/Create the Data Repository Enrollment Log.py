# Create the Data Repository Enrollment Log

# Import all required modules for this code
import pandas as pd
import numpy as np

# Create the path and locate the required files
path = r'C:\Users\cje4\Desktop\Data Repository'
# The child history downloaded is going to have the date and stuff after,
# Delete that stuff
childhistory = '\ThinkKidsChildHistor.csv'
# Grab a current copy of the REDCap Log from the shared folder
# place it in here
redcaplog = '\REDCap Log.xlsx'

# Import the child history from REDCap into the program
df = pd.read_csv(path + childhistory)
# Take only the rows that say 'Intake' under event name
df_childhistory = df[df['Event Name'] == 'Intake']
# Clean up the race/ethnicity section
df_childhistory = df_childhistory.replace({'Unchecked': np.nan, 'Checked': 1})
# Rename the columns so we have a 'REDcap ID' column to match with our
# REDCap Log
columns = [
    'REDcap ID', 'Event Name', 'Gender', 'African American or Black',
    'Asian American or Indian American',
    'European American or Caucasian/White', 'Latino(a)/Hispanic',
    'Native American (including Alaskan/Hawaiian Native)',
    'Other (please specify)'
]
df_childhistory.columns = columns

# Import the REDCap Log
df_redcaplog = pd.read_excel(path + redcaplog)
# Take only the rows where 'Repository Consentnt?' has some value in it.
df_redcaplog = df_redcaplog.dropna(subset=['Repository Consent?'])
# Take only the following three columns as they are all that is relivent to the
# enrollment log
df_redcaplog = df_redcaplog.filter([
    'REDcap ID', 'Repository Consent?', 'Consent Date'
], axis=1)

# Merge the two files into a new dataframe
results = pd.merge(df_redcaplog, df_childhistory, on=['REDcap ID'])
# Take only the rows where 'Repository Consent?' is filled out 'YES'
results = results[results['Repository Consent?'] == 'YES']
# Create a report called 'Results' in our 'Data Repository' folder on
# the desktop
results.to_excel(path + '\Results.xlsx')
