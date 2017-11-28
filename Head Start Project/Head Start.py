import pandas as pd
import numpy as np

path = r'C:\Users\cje4\Desktop\Head Start Data 11_21_17.xlsx'

df = pd.read_excel(path)
text_columns = (
    'ID #', 'DEM_CHILDAGE', 'DEM_CHILDAGE', 'DEM_RACE', 'DEM_EDU', 'DEM_PAR1',
    'DEM_PAR2', 'DEM_CUR', 'DEM_COM')
df = df.replace({99: np.nan})

df2 = df.mean()
df2 = pd.DataFrame(df2)
df2 = df2.transpose()

frame = [df, df2]

results = pd.concat(frame)

print(results)

results.to_csv(r'C:\Users\cje4\Desktop\Head Start Data.csv')