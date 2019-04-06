from __future__ import print_function
# for using print as a command, not as a command

from models.parser import Parser


if __name__ == '__main__':
    # the ToParse_python.xlsx is located one directory
    # above the current file in this system
    path_to_file = '../ToParse_Python.xlsx'

    # unit test (specifying the starting row of the list) , and
    # "strictness" of the parser as parameter 'strict'
    # [Ignore Empty Labels: False]

    p = Parser(path_to_file=path_to_file, sheet_number=0, list_start_row=8, strict=False)

    data = p.get_data()
    print(data)
