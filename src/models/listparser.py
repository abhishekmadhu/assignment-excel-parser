import xlrd

from gcd import get_cell_data


class ListParser(object):

    def __init__(self,workbook, sheet, start_row, start_col=2, n_rows=None, n_cols=None):
        self.workbook = workbook
        self.sheet = sheet
        self.start_row = start_row
        self.start_col = start_col
        self.n_rows = n_rows
        self.n_cols = n_cols

        self.headers = []


    def format_data(self, data):
        '''
            :param r: row
            :param c: column
            :return: the data, appropriately modified, for each cell
            '''

        val = data.value
        type = data.ctype
        if type == 3:  # implies that the data is a date
            val = xlrd.xldate.xldate_as_datetime(val, self.workbook.datemode)
            val = val.strftime('%Y-%m-%d')  # change it to the required format
        return val

    # get the headers in the sheet
    def get_headers(self):
        for j in range(0, self.sheet.ncols):
            x = self.sheet.cell(self.start_row, j)

            # if the cell is not empty, add it to headers
            if str(x.value) is not '':
                self.headers.append(str(x.value))

        # return the headers for testing purposes
        return self.headers

    def get_items_in_list(self):
        # initiate the column parsing procedure
        row, col = self.start_row, self.start_col
        MAX_COL = len(self.headers)
        MAX_ROW = self.sheet.nrows

        items_list = []  # list to add items

        # to get the headers for each column. This can be done dynamically also
        headers = self.headers

        # iterate over all the data in the tabular part of the sheet
        for r in range(row, self.sheet.nrows):
            index = col - 2  # as the table is shifted 2 columns to the right
            eachitem = {}
            for c in range(col, col + MAX_COL - 1):
                data = self.sheet.cell(r, c)
                val = self.format_data(data)
                # get the header
                c_header = headers[index]

                eachitem[c_header] = val

                index = index + 1

            items_list.append(eachitem)

