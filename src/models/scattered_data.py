import xlrd
import re
from listparser import ListParser
from src.common.datacleaner import DataCleaner


class ScatteredData(object):
    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet
        self.scattered_dict = {}


    def get_scattered_data(self):

        sheet = self.sheet
        # making a list of predefined (in the problem statement) except the name
        keys_to_search = ['Quote Number', 'Date', 'Ship To', 'Ship From']
        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                key = sheet.cell(r,c)
                # print key.value
                if key.value in keys_to_search:
                    # print 'Key found: ', key
                    next_col = c+1
                    val = sheet.cell(r, next_col)
                    # print val.value

                    key = DataCleaner.format_data(key, self.workbook)
                    val = DataCleaner.format_data(val, self.workbook)

                    self.scattered_dict[str(key)] = str(val)

                else:
                    # match name as a regexp
                    pattern = r'^Name:'

                    if re.match(pattern=pattern, string=str(key.value)):
                        name_cell = key.value

                        # split the string 'Name' and the actual name, and take the name
                        name = key.value.split(':')[1]

                        self.scattered_dict['Name'] = name

        return self.scattered_dict


if __name__ == '__main__':
    # unit test for only the scattereddata module
    # PLEASE IGNORE

    workbook = xlrd.open_workbook('../../ToParse_Python.xlsx')
    sheet = workbook.sheet_by_index(0)  # getting the first sheet
    s = ScatteredData(workbook, sheet)
    scattered_dict = s.get_scattered_data()
    print scattered_dict


