import pandas as pd

path = r'\\Cifs2\thinkkid$\Research\Chris\Youth Villages\Integrity Study Coding Sheets\Data from YV\Chris Manipulated Files\Code Files'

coder_averages = path + r'\Coder_Sheet.xlsx'

# Coder_Averages.xlsx
coder_averages = pd.read_excel(coder_averages)

specialists_tirfs = path + r'\Specialist TIRFS.xlsx'

# Specialsts_TIRFS
specialists_tirfs = pd.read_excel(specialists_tirfs)

specialist_column_names = ['Date', 'Youth ID', 'Staff ID']
for column in specialists_tirfs.columns[3:]:
    specialist_column_names.append(column + '_Specialist')

specialists_tirfs.columns = specialist_column_names

merge_columns = ['Date', 'Youth ID', 'Staff ID']

coders_specialists = pd.merge(left=coder_averages,
                              right=specialists_tirfs,
                              on=merge_columns,
                              how='outer')

coders_specialists.to_excel(path + r'\Coders_Specialists.xlsx')
