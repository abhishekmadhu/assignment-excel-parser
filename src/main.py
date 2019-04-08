from __future__ import print_function
# for using print as a function, not as a command

from models.parser import Parser


if __name__ == '__main__':
    # the ToParse_python.xlsx is located one directory
    # above the current file in this system
    path_to_file = '../ToParse_Python.xlsx'

    # UNIT TEST (specifying the starting row of the list) , and
    # "strictness" of the parser as parameter 'strict'
    # [Ignore Empty Labels: False]

    # keys to be searched in the top portion (above the list)
    keys_to_search = ['Quote Number', 'Date', 'Ship To', 'Ship From']

    # required headers are the column-names that we are to look for
    # (as in problem, others will be ignored)
    # change this if you want to look for other headers.
    # The headers in the sheet will be automatically parsed and extracted separately,
    # and out of them, only these (the following) will be searched for
    required_headers = ['LineNumber', 'PartNumber', 'Description', 'Price']

    # create a parser
    p = Parser(path_to_file=path_to_file, sheet_number=0,
               keys_to_search=keys_to_search,
               required_headers=required_headers,
               strict=False)

    # get data from the parser
    data = p.get_data()

    # since there is a possibility of warnings (when list values are missing),
    # the data part is separated from the errors for readability
    print('\n\n\n--- The data follows: ---\n')
    print(data)
