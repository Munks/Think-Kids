"""
This code compiles the 'Validating the CPS Integrity Coding System' data
"""

# Import all required modules for this code
import pandas as pd
import datetime as dt

# Set the path
path = r'\\Cifs2\thinkkid$\Research\Chris\Youth Villages\Integrity Study Coding Sheets'
compiled_data = path + r'\Compiled Data'
current_date = str(dt.date.today())

folder_date = compiled_data + r'\Data Compiled on ' + current_date

data = pd.read_excel(folder_date + r'\Full_Data ' + current_date + '.xlsx')

new_columns = []


for i in data.columns:
    if 'Time_' not in i:
        if 'Area' not in i:
            if 'Q' not in i:
                if 'Notes' not in i:
                    new_columns.append(i)

df = pd.DataFrame(data, columns=new_columns)

df1 = df.groupby(['Date', 'Family', 'Specialist'])

new_dfs = []

for i in df1:
    i = list(i)
    i = i[1:2]
    for x in i:
        df = pd.DataFrame(x)
        if df.shape[0] == 2:
            # TIRFS Ratings
            df_skillfullness = df.loc[:, 'Skillfullness_Component_1':'Global_Treatment_Integrity_A']
            df_skillfullness.fillna('NA', inplace=True)
            # TPOCSA Ratings
            tpocsa = df.loc[:, 'TPOC_1':'TPOCBR_6']
            tpocsa.fillna('NA', inplace=True)
            # Start new DF
            output_df = df.iloc[0, [1, 5, 6]]
            # 1 = date
            # 5 = Family ID
            # 6 = Specialist ID
            coder1 = df.iloc[0, [3, 31, 32, 33, 7, 8, 9, 10, 11, 15]]
            coder1_tirf = df_skillfullness.iloc[0]
            coder1 = coder1.append(coder1_tirf)
            coder1_tpocsa = tpocsa.iloc[0]
            coder1 = coder1.append(coder1_tpocsa)
            coder1 = coder1.add_suffix('_Coder_1')
            output_df = output_df.append(coder1)
            coder2 = df.iloc[1, [3, 31, 32, 33, 7, 8, 9, 10, 11, 15]]
            coder2_tirf = df_skillfullness.iloc[1]
            coder2 = coder2.append(coder2_tirf)
            coder2_tpocsa = tpocsa.iloc[1]
            coder2 = coder2.append(coder2_tpocsa)
            coder2 = coder2.add_suffix('_Coder_2')
            output_df = output_df.append(coder2)
            output_df['Number_of_Raters'] = 2
            output_df['average_percent_cps_time'] = df.percent_of_cps_time.mean()
            output_df['average_rightmost_column_count'] = df.rightmost_column_count.mean()

            new_dfs.append(output_df)

data = pd.DataFrame(new_dfs)

path = r'\\Cifs2\thinkkid$\Research\Chris\Youth Villages\Integrity Study Coding Sheets\Data from YV\Chris Manipulated Files\Code Files'

data.to_excel(path + r'\Coder_Sheet.xlsx')
