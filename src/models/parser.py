import json
from collections import OrderedDict
import xlrd

from src.models.listparser2 import ListParser
from src.models.scattered_data import ScatteredData


class Parser(object):
    def __init__(self, path_to_file, sheet_number, list_start_row, headers=None, strict=True):

        self.path_to_file = path_to_file
        self.sheet_number = sheet_number
        self.list_start_row = list_start_row
        self.strict = strict
        self.headers = headers

        self.workbook = xlrd.open_workbook(self.path_to_file)       # getting the workbook
        self.sheet = self.workbook.sheet_by_index(self.sheet_number)     # getting the sheet

        # specified in the problem. A different getHeaders function is also there to gram self.headers
        # if it is None
        self.headers = ['LineNumber', 'PartNumber', 'Description', 'Price']

    def get_data(self):
        # initialising an ordered dict to store the final dictionary
        mydict = OrderedDict()

        workbook = xlrd.open_workbook(self.path_to_file)
        sheet = workbook.sheet_by_index(0)  # getting the first sheet
        s = ScatteredData(workbook, sheet, strict=self.strict)
        scattered_dict = s.get_scattered_data()
        mydict.update(scattered_dict)


        listparser = ListParser(workbook, sheet, start_row=self.list_start_row)
        listparser.set_headers(self.headers)

        # headers = ['LineNumber', 'PartNumber', 'Description', 'Item Type', 'Price']

        items = listparser.get_items_in_list()

        mydict['Items'] = items if items is not None else 'No items detected in this sheet'

        # format the dictionary in the required manner
        # making it a json object helps transmission over HTTP
        json_file = json.dumps(mydict, indent=4, sort_keys=False)

        return json_file


if __name__ == '__main__':
    # the ToParse_python.xlsx is located one directory
    # above the current file in this system
    path_to_file = '../../ToParse_Python.xlsx'

    # unit test (specifying the starting row of the list) , and
    # "strictness" of the parser as parameter 'strict'
    # [Ignore Empty Labels: False]

    p = Parser(path_to_file=path_to_file, sheet_number=0, list_start_row=8, strict=False)

    data = p.get_data()
    print(data)

