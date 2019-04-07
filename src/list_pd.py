import pandas as pd
import numpy as np

data = pd.read_excel('../ToParse_Python.xlsx')
print data
print data.keys()

start_row, start_col = 7, 2

new_data = data.iloc[start_row:, :]
print new_data
print new_data.keys()

new_header = new_data.iloc[0]   # grab the first row for the header
new_data = new_data[1:]         # take the data less the header row
new_data.columns = new_header   # set the header row as the df header
print 'keys are: ', new_data.keys()



