"""
Cleaning the YOQ Data
"""

# Import all required modules for this code
import pandas as pd

# Create the path and locate the required files
path = r'C:\Users\cje4\Desktop\Data Repository'
yoq_data_loc = r'\2017Dec12_OQACustomReport.csv'

yoq_data = pd.read_csv(path + yoq_data_loc)

yoq_data.set_index(['PersonID'])
