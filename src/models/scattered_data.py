import xlrd
import re

from termcolor import colored

from src.common.datacleaner import DataCleaner
from src.models.errors import DataNotBesideLabelError


class ScatteredData(object):
    def __init__(self, workbook, sheet, keys_to_search, strict):
        self.workbook = workbook
        self.sheet = sheet
        self.strict = strict

        self.keys_to_search = keys_to_search
        self.scattered_dict = {}


    def get_scattered_data(self):

        sheet = self.sheet
        # making a list of predefined (in the problem statement) except the name

        for r in range(sheet.nrows):
            for c in range(sheet.ncols):
                key = sheet.cell(r,c)
                # print key.value
                if key.value in self.keys_to_search:
                    # print 'Key found: ', key
                    next_col = c+1
                    val = sheet.cell(r, next_col)

                    # format the data in the proper manner required
                    key = DataCleaner.format_data(key, self.workbook)
                    val = DataCleaner.format_data(val, self.workbook)

                    if val != '':
                        # add the key val pair to the dictionary
                        self.scattered_dict[str(key)] = str(val)

                        # remove the found key from the searchable list
                        self.keys_to_search.remove(key)

                    else:
                        # handle the situation where the value is not located
                        # to the right of the label

                        # if strict mode is on
                        if self.strict:
                            raise DataNotBesideLabelError('Data is not beside the label! Pass "strict=False" to ignore this')



                else:
                    # match name as a regexp
                    pattern = r'^Name:'

                    if re.match(pattern=pattern, string=str(key.value)):
                        name_cell = key.value

                        # split the string 'Name' and the actual name, and take the name
                        name = key.value.split(':')[1]

                        self.scattered_dict['Name'] = name

        return self.scattered_dict

    def are_all_keys_found(self):
        if len(self.keys_to_search) > 0:
            print colored('WARNING: Required headers not found!', 'yellow')
            print 'They are: '
            for item in self.keys_to_search:
                print '\'', item, '\''
        else:

            # all keys found
            pass


if __name__ == '__main__':
    # unit test for only the scattereddata module
    # PLEASE IGNORE

    workbook = xlrd.open_workbook('../../ToParse_Python.xlsx')
    sheet = workbook.sheet_by_index(0)  # getting the first sheet
    keys_to_search = ['Quote Number', 'Date', 'Ship To', 'Ship From']
    s = ScatteredData(workbook, sheet, keys_to_search=keys_to_search,
                      strict=False)
    scattered_dict = s.get_scattered_data()
    s.are_all_keys_found()



