import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


df = pd.read_excel('coach_rated_TIRFs.xlsx')

skill_columns = {'Assess Skill 1': 'assess_skills_1',
                 'Assess Skill 2': 'assess_skills_2',
                 'Prioritize Skill 3': 'prioritize_skill_3',
                 'Prioritize Skill 4': 'prioritize_skill_4',
                 'Plan B Skill 5': 'plan_b_skill_5',
                 'Plan B Skill 6': 'plan_b_skill_6',
                 'Plan B Skill 7': 'plan_b_skill_7',
                 'Plan B Skill 8': 'plan_b_skill_8',
                 'Plan B Skill 9': 'plan_b_skill_9',
                 'Regulate Skill 10-11': 'regulate_skill_10_11',
                 'Differentiate Skill 12': 'differentiate_skill_12'}

for key, value in skill_columns.items():
    hist = df.hist(column=value,
                   bins=np.arange(6) - .5,
                   figsize=(4, 3))
    plt.title(key, fontsize=12)
    plt.xticks((0, 1, 2, 3, 4))  # labels along the bottom edge are off
    plt.ylim(0, 70)
    plt.axes().xaxis.grid(False)
    plt.savefig(key + '.png')

columns = {'Collaborative Stance': 'collab_stance_13',
           'Philosophy': 'philos_14',
           'Global Integrity': 'global_integrity_skill_15'}

for key, value in columns.items():
    hist = df.hist(column=value,
                   bins=np.arange(5) - .5,
                   figsize=(4, 3))
    plt.title(key, fontsize=12)
    plt.ylim(0, 70)
    plt.xlim(0.5, 4.5)
    plt.axes().xaxis.grid(False)
    plt.savefig(key + '.png')
