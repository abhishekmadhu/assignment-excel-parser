import xlrd


def get_cell_data(sheet, r, c):
    '''
    :param r: row
    :param c: column
    :return: the data, appropriately modified, for each cell
    '''


    cell_value = sheet.cell(r, c)
    data = cell_value.value
    type = cell_value.ctype




    if type == 3:       # implies that the data is a date
       data = xlrd.xldate.xldate_as_datetime(data, datafile.datemode)
       data = data.strftime('%Y-%m-%d')    # change it to the required format


    return data

