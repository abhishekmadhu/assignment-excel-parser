from termcolor import colored
import xlrd
import re
from src.common.datacleaner import DataCleaner
from collections import OrderedDict


class ListParser(object):
    def __init__(self, workbook, sheet, required_headers, start_row=None,
                 start_col=None, n_rows=None, n_cols=None):
        self.workbook = workbook
        self.sheet = sheet
        self.start_row = start_row
        self.start_col = start_col
        self.n_rows = n_rows
        self.n_cols = n_cols
        self.required_headers = required_headers

        self.headers = []

    # get the headers in the sheet
    def get_headers(self):
        for j in range(0, self.sheet.ncols):
            x = self.sheet.cell(self.start_row, j)

            # if the cell is not empty, add it to headers
            if str(x.value) is not '':
                self.headers.append(str(x.value))

        # return the headers for testing purposes
        return self.headers

    def verify_headers(self, required_headers):
        for rh in required_headers:
            if rh not in self.headers:
                print rh, ' column has not been found in the list. ' \
                          'Please consult the administrator if you think it is present in the list.'

    def get_start_row_col(self):
        for r in range(self.sheet.nrows):
            for c in range(self.sheet.ncols):

                # A bold assumption that the first col of the list will always be 'LineNumber'
                # change this later by passing it from the user
                if self.sheet.cell(r, c).value == u'LineNumber':
                    # print 'Start of the List Found!', r, c

                    # if the start_row or start_col has not been passed by user, set them
                    if self.start_row is None:
                        self.start_row = r
                    if self.start_col is None:
                        self.start_col = c

                    # return (r, c) tuple for testing purposes
                    return r, c

    def get_items_in_list(self):
        # initiate the column parsing procedure

        MAX_COL = len(self.headers)
        MAX_ROW = self.sheet.nrows

        # if self.start_row is None or self.start_col is None:
        #     found_start_row, found_start_col = self.get_start_row_col()
        #     # handle the case where the col and row cannot be detected
        #     if found_start_row is None:
        #         print 'Could not detect the starting row of the list. Is "LineNumber" column present?'
        #         print 'Please pass the "start_row" parameter'
        #
        #     if found_start_col is None:
        #         print 'Could not detect the starting column of the list. Is "LineNumber" column present?'
        #         print 'Please pass the "start_col" parameter'
        #
        #     else:
        #         self.start_row, self.start_col = found_start_row, found_start_col

        row, col = self.start_row, self.start_col

        items_list = []  # list to add items

        # to get the headers for each column. This can be done dynamically also
        headers = self.headers

        # iterate over all the data in the tabular part of the sheet
        # (row +1) as the first row is the header row
        for r in range(row+1, self.sheet.nrows):
            index = 0  # as the table is shifted those many columns to the right

            eachitem = OrderedDict()

            for c in range(col, col + MAX_COL):
                data = self.sheet.cell(r, c)

                # clean the data into proper format
                val = DataCleaner.format_data(data, self.workbook)

                # if val is atleast 10 dashes '----------', return the list and stop
                pattern = '----------'
                if re.match(pattern=pattern, string=str(val)):
                    print '\nTen or more consecutive dashes found. Exiting!'
                    return items_list

                # get the header for that column
                c_header = headers[index]

                if c_header in self.required_headers:

                    # if the values are not empty, add them to the dictionary
                    if val is not u'' and c_header is not u'':
                        eachitem[c_header] = val

                    else:
                        print colored('Critical Error: ', 'red')
                        print 'Empty value found at cell(', r, ',', c, ')! ' \
                            'No entry for the', c_header, 'column!'

                        pass

                index = index + 1
            if len(eachitem) != 0:
                items_list.append(eachitem)

        return items_list


if __name__ == '__main__':
    # standalone unit test for the listparser module
    # PLEASE IGNORE

    workbook = xlrd.open_workbook('../../ToParse_Python.xlsx')
    sheet = workbook.sheet_by_index(0)      # getting the first sheet
    required_headers = ['LineNumber', 'PartNumber', 'Description', 'Price']
    # init a new listparser
    listparser = ListParser(workbook, sheet, required_headers=required_headers)

    # get the starting rows and cols of the list
    list_start_row, list_start_col = listparser.get_start_row_col()

    # get all the headers available in the sheet
    found_headers = listparser.get_headers()

    # headers = ['LineNumber', 'PartNumber', 'Description', 'Item Type', 'Price']

    # check if all the required headers have been found or not
    listparser.verify_headers(required_headers)

    # get the items for the 'Items:' key of the dict
    items = listparser.get_items_in_list()
    print items
