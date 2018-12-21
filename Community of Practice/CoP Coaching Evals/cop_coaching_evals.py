import pandas as pd
import numpy as np
from redcap import Project
import sys
sys.path.append(r'C:\Python Programs\Think-Kids_Private')
import tokens as tk


def cell_content(df, row, column):
    """
    This function saves the content of a single cell in a DataFrame as a list
    """
    content = df.loc[[row], [column]]
    content = content[column]
    content_list = []
    for i in content:
        content_list.append(i)
    return content_list


def meta_dict(df, row):
    """
    Create a python dictionary from the Data Dictionary downloaded from REDcap

    df = Data Dictionary DataFrame Pandas Object.
    row = name of row in the Data Dictionary where you are looking for
        dictionary keys and values.
    The column name in the Data Dictionary is with keys and values is a
        constant and always 'select_choices_or_calculations'.
    """
    column = 'select_choices_or_calculations'
    content = cell_content(df, row, column)
    dictionary_keys = []
    dictionary_values = []
    for i in content:
        content = i.split(' | ')
        # Divide the string into lines as displayed in REDcap.
        for i in content:
            content = i.split(', ', 1)
            # Divide the lines into Answer and Key as input in REDcap and save
            #   as a list.
            a = 0
            for i in content:
                # First item in the list is the key. Add to the Keys List.
                if a == 0:
                    dictionary_keys.append(int(i))
                    a += 1
                # Second item in the list is the value. Add to the Values List.
                else:
                    dictionary_values.append(i)
    dictionary = dict(zip(dictionary_keys, dictionary_values))
    return dictionary


tk.cop_coaching_evals()

project = Project(tk.api_url, tk.api_token)
# CPS-AIM MetaData
df = project.export_records(format='df')
meta_df = project.export_metadata(format='df')

trainers = meta_dict(meta_df, 'trainer1')
percent_attend = meta_dict(meta_df, 'percent_coaching_attended')

percent_attend = dict((keys,
                       int(values)) for keys, values in percent_attend.items())

df['percent_coaching_attended'] = (
    df['percent_coaching_attended'].map(percent_attend))
df['trainer1'] = df['trainer1'].map(trainers)
rating_col_names = ['Lecuture', ' Slides', 'Videos', 'Handouts', 'Polling',
                    'Case Study', 'Pair Share', 'Large Group Discussion',
                    'Guiding Questions', 'Role Plays']
trainer_col_names = ['Speech', 'Enthusiasm', 'Interaction', 'Organization',
                     'Challenges', 'Role Play', 'Explained', 'Mindset',
                     'Answered', 'Examples', 'Pace', 'Apprecaited']
overall_col_names = ['Content', 'Relevant', 'Organized', 'Experience Level',
                     'Interesting', 'Apply', 'Experience', 'Recommend']
learn_content = []
understand_philosophy = []
improve_cps_ability = []
trainer_one = []
overall = []
comments = ['mosthelp', 'leasthelp', 'comments']

for i in meta_df.index:
    if 'activity_rate' in str(meta_df.loc[i]['matrix_group_name']):
        learn_content.append(i)
    elif 'activity3_rate' in str(meta_df.loc[i]['matrix_group_name']):
        understand_philosophy.append(i)
    elif 'activity4_rate' in str(meta_df.loc[i]['matrix_group_name']):
        improve_cps_ability.append(i)
    elif 'trainer1_rate' in str(meta_df.loc[i]['matrix_group_name']):
        trainer_one.append(i)
    elif 'overall_workshop' in str(meta_df.loc[i]['matrix_group_name']):
        overall.append(i)

content_philo_ability = [learn_content,
                         understand_philosophy,
                         improve_cps_ability]

for content_list in content_philo_ability:
    for column in content_list:
        df[column] = df[column].replace(9, np.nan)
for column in trainer_one:
    df[column] = df[column].replace(6, np.nan)
for column in overall:
    df[column] = df[column].replace(7, np.nan)

learn_content = dict(zip(learn_content, rating_col_names))
understand_philosophy = dict(zip(understand_philosophy, rating_col_names))
improve_cps_ability = dict(zip(improve_cps_ability, rating_col_names))
trainer_one = dict(zip(trainer_one, trainer_col_names))
overall = dict(zip(overall, overall_col_names))


content_philo_ability = [learn_content,
                         understand_philosophy,
                         improve_cps_ability,
                         trainer_one,
                         overall]

for content_list in content_philo_ability:
    df = df.rename(columns=content_list)

df.to_excel('cop_coaching_evals.xlsx')
