# This program builds the original spreadsheet with what we want the columns to be named.

import pandas as pd
import os

if not os.path.exists(r'C:\Users\cje4\Desktop\Integrity Study Coding Tests\Headers.xlsx'):
    headers = ['FileName', 'CodingDate', 'Specialist',
               'Family', 'Site', 'Coder', 'Length of Session']

    for i in range(12):
        for x in range(15):
            if x != 14:
                header = ('Component_' + str(1 + i) + '_Time_' + str(1 + x))
                headers.append(header)
                skillfullness = ('Skillfullness_Component_' + str(1 + i))
            else:
                headers.append(skillfullness)

    tpocsa_headers = ['TPOCS-A_1', 'TPOCS-A_2', 'TPOCS-A_3', 'TPOCS-A_4', 'TPOCS-A_R1',
                      'TPOCS-A_R2', 'TPOCS-A_R3', 'TPOCS-A_R4', 'TPOCS-A_R5', 'TPOCS-A_R6']
    for header in tpocsa_headers:
        headers.append(header)

    df = pd.DataFrame(columns=headers)

    df.to_excel(
        r'C:\Users\cje4\Desktop\Integrity Study Coding Tests\Headers.xlsx')
