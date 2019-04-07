import xlrd
import re
from src.common.datacleaner import DataCleaner
from collections import OrderedDict

class ListParser(object):
    def __init__(self, workbook, sheet, start_row, start_col=2, n_rows=None, n_cols=None):
        self.workbook = workbook
        self.sheet = sheet
        self.start_row = start_row
        self.start_col = start_col
        self.n_rows = n_rows
        self.n_cols = n_cols

        self.headers = []

    # get the headers in the sheet
    def set_headers(self, headers):
        self.headers = headers

    def get_items_in_list(self):
        # initiate the column parsing procedure
        row, col = self.start_row, self.start_col
        MAX_COL = len(self.headers)
        MAX_ROW = self.sheet.nrows

        items_list = []  # list to add items

        # to get the headers for each column. This can be done dynamically also
        headers = self.headers

        # iterate over all the data in the tabular part of the sheet
        # (row +1) as the first row is the header row
        for r in range(row+1, self.sheet.nrows):

            if col - self.start_col > 0:    # to make sure that the index never goes below zero
                index = col - self.start_col  # as the table is shifted those many columns to the right
            else:
                index = -1                  # to ignore the value of index while searching for header

            eachitem = OrderedDict()
            for c in range(0, col + MAX_COL):
                data = self.sheet.cell(r, c)

                # clean the data into proper format
                val = DataCleaner.format_data(data, self.workbook)
                print val

                if c >= col:

                    # if val is atleast 10 dashes '----------', break
                    pattern = '----------'
                    if re.match(pattern=pattern, string=str(val)):
                        print '10 or more dashes found. Exiting!'
                        break

                    # get the header for that column
                    c_header = headers[index]

                    # if the values are not empty, add them to the dictionary
                    if val is not u'' and c_header is not u'':
                        eachitem[c_header] = val

                    index = index + 1
            if len(eachitem) != 0:
                items_list.append(eachitem)

        return items_list


if __name__ == '__main__':
    # standalone unit test for the listparser module
    # PLEASE IGNORE

    workbook = xlrd.open_workbook('../../ToParse_Python.xlsx')
    sheet = workbook.sheet_by_index(0)      # getting the first sheet
    l = ListParser(workbook, sheet, start_row=8)        # 8 only for testing
    l.set_headers()

    # headers = ['LineNumber', 'PartNumber', 'Description', 'Item Type', 'Price']
    # for our case
    items = l.get_items_in_list()
    print items
