import pandas as pd

path = r'\\Cifs2\thinkkid$\Research\Chris\Youth Villages\Integrity Study Coding Sheets\Data from YV\Chris Manipulated Files\Code Files'

consultation_key = path + r'\Consultation with Key.xlsx'

consultation_key = pd.read_excel(consultation_key)

coders_specialists = path + r'\Coders_Specialists.xlsx'

coders_specialists = pd. read_excel(coders_specialists)

merge_columns = ['Date', 'Youth ID', 'Staff ID']

coders_specialist_consultant = pd.merge(left=coders_specialists,
                                        right=consultation_key,
                                        on=merge_columns,
                                        how='outer')

coders_specialist_consultant.to_excel(
    path + r'\Coders_Specialists_Consultant.xlsx')
