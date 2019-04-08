import json
from collections import OrderedDict
import xlrd

from src.models.listparser import ListParser
from src.models.scattered_data import ScatteredData


class Parser(object):
    def __init__(self, path_to_file, sheet_number, keys_to_search,
                 list_start_row=None, required_headers=None,strict=True):

        self.path_to_file = path_to_file
        self.sheet_number = sheet_number
        self.list_start_row = list_start_row
        self.strict = strict
        self.required_headers = required_headers
        self.keys_to_search = keys_to_search

        self.workbook = xlrd.open_workbook(self.path_to_file)           # getting the workbook
        self.sheet = self.workbook.sheet_by_index(self.sheet_number)    # getting the sheet

    def get_data(self):
        # initialising an ordered dict to store the final dictionary
        mydict = OrderedDict()

        workbook = xlrd.open_workbook(self.path_to_file)
        sheet = workbook.sheet_by_index(0)  # getting the first sheet

        # init a ScatteredData object
        s = ScatteredData(workbook, sheet, keys_to_search=self.keys_to_search,
                          strict=self.strict)

        # gets the scattered data in the sheet andkeys_to_search = ['Quote Number', 'Date', 'Ship To', 'Ship From']
        # stores the data returned in a dictionary
        scattered_dict = s.get_scattered_data()
        s.are_all_keys_found()

        # update the dictionary with the returned data by merging them
        mydict.update(scattered_dict)

        # init a new listparser
        listparser = ListParser(workbook, sheet, required_headers=self.required_headers)

        #get the starting rows and cols of the list
        list_start_row, list_start_col = listparser.get_start_row_col()

        #get all the headers available in the sheet
        found_headers = listparser.get_headers()

        # headers = ['LineNumber', 'PartNumber', 'Description', 'Item Type', 'Price']

        # check if all the required headers have been found or not
        listparser.verify_headers(self.required_headers)

        # get the items for the 'Items:' key of the dict
        items = listparser.get_items_in_list()

        # add it to the list if there is atleast one item in the list, else add error msg
        mydict['Items'] = items if items is not None else 'No items detected in this sheet'

        # format the dictionary in the required manner
        # making it a json object helps transmission over HTTP
        json_file = json.dumps(mydict, indent=4, sort_keys=False)

        return json_file


if __name__ == '__main__':
    # unit test for the parser class
    # Please ignore

    # the ToParse_python.xlsx is located one directory
    # above the current file in this system
    path_to_file = '../../ToParse_Python.xlsx'

    # unit test (specifying the starting row of the list) , and
    # "strictness" of the parser as parameter 'strict'
    # [Ignore Empty Labels: False]

    keys_to_search = ['Quote Number', 'Date', 'Ship To', 'Ship From']
    required_headers = ['LineNumber', 'PartNumber', 'Description', 'Price']

    # create a parser
    p = Parser(path_to_file=path_to_file, sheet_number=0,
               keys_to_search=keys_to_search,
               required_headers=required_headers,
               strict=False)

    data = p.get_data()
    print(data)

