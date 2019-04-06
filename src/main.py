from __future__ import print_function
# for using print as a command, not as a command
from collections import OrderedDict

from models.listparser import ListParser

import xlrd
import datetime
import json


# the ToParse_python.xlsx is located one directory
# above the current file in this system
workbook = xlrd.open_workbook('../ToParse_Python.xlsx')
sheet = workbook.sheet_by_index(0)  # getting the first sheet




def get_cell_data(r, c):
    '''


    :param r: row
    :param c: column
    :return: the data, appropriately modified, for each cell
    '''


    cell_value = sheet.cell(r, c)
    data = cell_value.value
    type = cell_value.ctype




    if type == 3:       # implies that the data is a date
       data = xlrd.xldate.xldate_as_datetime(data, workbook.datemode)
       data = data.strftime('%Y-%m-%d')    # change it to the required format


    return data




if __name__ == '__main__':
    mydict = OrderedDict()


    # get the quote number
    row = 1
    col = 1
    qn = get_cell_data(row, col)
    col += 1
    qn_val = get_cell_data(row, col)
    mydict[qn] = int(qn_val)


    # get the date key-value pair
    col += 2
    dt = get_cell_data(row, col)
    col += 1
    dt_val = get_cell_data(row, col)
    mydict[dt] = dt_val


    # Get the shipping address
    row, col = 3, 1
    st = get_cell_data(row, col)
    col += 1
    st_val = get_cell_data(row, col)
    mydict[st] = st_val


    # Get the name
    row, col = 5, 1
    name_text = get_cell_data(row, col)
    key, name = str(name_text).strip().split(':')  # separate the name from the key
    mydict[key] = name


    # initiate the column parsing procedure
    row, col = 9, 2
    MAX_COL = 6
    MAX_ROW = 2


    items_list = []  # list to add items


    # to get the headers for each column. This can be done dynamically also
    # headers = [sheet.cell(8, col_index).value for col_index in range(0, 10)]

    listparser = ListParser(workbook, sheet, 8, 2)
    headers = listparser.get_headers()



    # headers = ['LineNumber', 'PartNumber', 'Description', 'Item Type', 'Price']
    print(headers)

    items = listparser.get_items_in_list()
    print(items)

    # iterate over all the data in the tabular part of the sheet
    for r in range(row, row + MAX_ROW):
       index = col - 2  # as the table is shifted 2 columns to the right
       eachitem = {}
       for c in range(col, col + MAX_COL - 1):
           val = get_cell_data(r, c)


           # get the header
           c_header = headers[index]


           eachitem[c_header] = val


           index = index + 1


       items_list.append(eachitem)


    mydict['Items'] = items_list


    # format the dictionary in the required manner
    # making it a json object helps transmission over HTTP
    json_file = json.dumps(mydict, indent=4, sort_keys=False)


    print(json_file)