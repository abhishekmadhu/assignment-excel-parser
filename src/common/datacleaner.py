import xlrd


class DataCleaner(object):
    def __init__(self):
        pass

    @staticmethod
    def format_data(data, workbook):
        '''
            :param r: row of the data
            :param c: column of the data
            :return: the data, appropriately modified, for each cell
            '''

        val = data.value
        type = data.ctype

        if type == 3:      # implies that the data is a date

            # generate a datetime object that is at par with the workbook date mode
            val = xlrd.xldate.xldate_as_datetime(val, workbook.datemode)

            # change the datetime to a string of the required format
            val = val.strftime('%Y-%m-%d')
        return val



