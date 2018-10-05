import pandas as pd

path = r'\\Cifs2\thinkkid$\Research\Chris\Youth Villages\Integrity Study Coding Sheets\Data from YV\Chris Manipulated Files\Code Files'

consultation_key = path + r'\Consultation to Session Date key.xlsx'

consultation_key = pd.read_excel(consultation_key)

consultation_key.Consultation_Date = consultation_key.Consultation_Date.apply(
    str)

consultation_data = path + r'\Consultant TIRFS.xlsx'

consultation_data = pd.read_excel(consultation_data)

consultation_data.Consultation_Date = consultation_data.Consultation_Date.apply(
    str)

merge_columns = ['Consultant', 'Consultation_Date', 'Specialist']

consultation_with_key = pd.merge(left=consultation_key,
                                 right=consultation_data,
                                 on=merge_columns,
                                 how='outer')

consultation_with_key.to_excel(path + r'\Consultation with Key.xlsx')
