"""
Up Academy Readiness Completion Report
"""
import pandas as pd
import datetime as dt
from redcap import Project
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import tokens as tk


# User Interface Introduction
print("\n\n\t\t* * * * * * *   *    *      *")
print("\t\t      *        * *   *    *")
print("\t\t      *         *    *  *")
print("\t\t      *              **")
print("\t\t      *         *    *  *")
print("\t\t      *        * *   *    *")
print("\t\t      *         *    *      *")
print("\n\t\t\tUp Completion Reports!!\n\n")

print('\n\nThis file will be saved to "\Cifs2\\thinkkid$\Research\Readiness"')
input('\n\nPlease make sure that the Up Completion Report is closed. \
This program will not run if it is open. \n\n\tPress Enter to continue')

# Functions


def meta_cell_content(df, row_name, column_name):
    """
    Retrieve the content of a given cell in the Data Dictionary downlaoded from
    REDcap
    """
    dict_name = df.loc[[row_name], [column_name]]
    dict_name = dict_name[column_name]
    item = []
    for i in dict_name:
        item.append(i)
    return item


def meta_dict(df, row_name):
    """
    Create a python dictionary from the Data Dictionary downloaded from REDcap
    df = Data Dictionary DataFrame Pandas Object
    row_name = Name of Row
    The column for this Dictionary is always 'select_choices_or_calculations'
    """
    column_name = 'select_choices_or_calculations'
    item = meta_cell_content(df, row_name, column_name)
    dic_key = []
    dic_val = []
    for i in item:
        item = i.split(' | ')
        for i in item:
            item = i.split(', ', 1)
            a = 0
            for i in item:
                if a == 0:
                    dic_key.append(int(i))
                    a += 1
                else:
                    dic_val.append(i)
    dict_name = dict(zip(dic_key, dic_val))
    return dict_name


# Variables

current_date = str(dt.date.today().strftime('%#m/%#d/%Y'))

# project = Project(tk.api_url, tk.api_token)

print('\n\nDownloading the file from REDcap')
tk.readiness_3_part()
project = Project(tk.api_url, tk.api_token)
three_part_readiness = project.export_records(format='df')
metadata_three_part_readiness = project.export_metadata(format='df')

up_dictionary = meta_dict(metadata_three_part_readiness, 'up_academy_spec')
print('\n\nCounting only the people from Up!')
up_readiness = three_part_readiness[
    three_part_readiness['organization'] == 204]

up_readiness = up_readiness.dropna(thresh=41)

count = {'Total': 0}

for i in up_readiness.up_academy_spec:
    if up_dictionary[i] not in count.keys():
        count[up_dictionary[i]] = 1
        count['Total'] += 1
    else:
        count[up_dictionary[i]] += 1
        count['Total'] += 1
print('\n\nOpening the old Up Completion Report')
path = r'\\Cifs2\thinkkid$\Research\Readiness\Up_Completion_Report.csv'
up_readiness_report = pd.read_csv(path, index_col=0)
df = pd.DataFrame(data=count, index=[current_date])
for i in up_readiness_report.index:
    if i == current_date:
        up_readiness_report.drop([i], inplace=True)
frames = []
for i in df.index:
    if i not in up_readiness_report.index:
        frames = [up_readiness_report, df]
print('\n\nUpdated it with the current numbers from ' + str(current_date) +
      '\nShould be done any second!')

result = pd.concat(frames)
result.to_csv(path)
input('Press Enter to Close')
