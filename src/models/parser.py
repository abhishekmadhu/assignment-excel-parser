import json
from collections import OrderedDict
import xlrd

from src.models.listparser import ListParser
from src.models.scattered_data import ScatteredData


class Parser(object):
    def __init__(self, path_to_file, sheet_number, list_start_row, strict=True):

        self.path_to_file = path_to_file
        self.sheet_number = sheet_number
        self.list_start_row = list_start_row
        self.strict = strict

        self.workbook = xlrd.open_workbook(self.path_to_file)       # getting the workbook
        self.sheet = self.workbook.sheet_by_index(self.sheet_number)     # getting the sheet


    def get_data(self):
        # initialising an ordered dict to store the final dictionary
        mydict = OrderedDict()

        workbook = xlrd.open_workbook('../ToParse_Python.xlsx')
        sheet = workbook.sheet_by_index(0)  # getting the first sheet
        s = ScatteredData(workbook, sheet, strict=self.strict)
        scattered_dict = s.get_scattered_data()
        mydict.update(scattered_dict)


        listparser = ListParser(workbook, sheet, start_row=self.list_start_row)
        headers = listparser.get_headers()

        # headers = ['LineNumber', 'PartNumber', 'Description', 'Item Type', 'Price']

        items = listparser.get_items_in_list()

        mydict['Items'] = items if items is not None else 'No items detected in this sheet'

        # format the dictionary in the required manner
        # making it a json object helps transmission over HTTP
        json_file = json.dumps(mydict, indent=4, sort_keys=False)

        return json_file
