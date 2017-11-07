import os
import glob
import pandas as pd
import numpy as np

#	Set the home path for the program

path = os.path.expanduser("~/Desktop/")
os.chdir(path)

#	Create the function to import the excel files


def coding_import():
    files = glob.glob(path)
    data_frames = []
    coder_number = filename[-25:-22]
    for filename in files:
        print(filename)
        df = pd.read_excel(filename)
    data_frame = pd.concat(data_frames)
    data_frame.to_excel(path + '/' + 'Coding Data.xlsx')

#	Need a function to slice up the excel into rows. Each row is its own data frame. Reconnect the rows horizontally


def slice_rows():
    for row in df:
        print(row)


#	Build a dictionary to change column names

column_remane = {
}
coding_import()
