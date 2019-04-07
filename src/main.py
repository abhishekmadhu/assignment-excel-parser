from __future__ import print_function
# for using print as a command, not as a command

from models.parser import Parser


if __name__ == '__main__':
    # unit test for the parser class
    # Please ignore

    # the ToParse_python.xlsx is located one directory
    # above the current file in this system
    path_to_file = '../ToParse_Python.xlsx'

    # unit test (specifying the starting row of the list) , and
    # "strictness" of the parser as parameter 'strict'
    # [Ignore Empty Labels: False]

    # required headers are the columnnames that we are to look for
    # (as in problem, others will be ignored)
    required_headers = ['LineNumber', 'PartNumber', 'Description', 'Price']

    # create a parser
    p = Parser(path_to_file=path_to_file, sheet_number=0,
               required_headers=required_headers,
               strict=False)

    # get data from the parser
    data = p.get_data()

    # since there is a possibility of warnings (when list values are missing),
    # the data part is separated from the errors for readability
    print('\n\n\n--- The data follows: ---\n')
    print(data)
