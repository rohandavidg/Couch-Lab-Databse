#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

"""

This scripts validates each field for each
induvidual excel table generated from
excel_parse.py and after changes have been
made by Dr. Couch's staff

"""


import xlwt
import sys
import logging
import pprint
import ctypes
import win32com.client as win32
from excel_parse import configure_logger
import datetime
import re
import os
import xlrd
import math

required_field_list = ['CAST Barcode', 'CAST Plate/Box', 'Date Received', 'Sub_Contact ID', 'Sub_Contact Person',
                       'Sub_Contact E-mail', 'Sub_Project Type', 'Sub_Plate Name', 'CARRIERS ID',
                       'Sub_Sample Name', 'Sub_Individual ID', 'Sub_Gender', 'Sub_Sample Status']

unique_field_list = ['CAST Barcode', 'Sub_Plate Name', 'CARRIERS ID', 'Sub_Sample Name', 'Sub_Coord']
sample_status_accepted_values = [ 'Case', 'Control', 'Proband', 'Family Member']
sub_race_values = ['Hispanic or Latino', 'Unknown', 'Non-Hispanic/Latino']
sub_gender_values = ['M', 'F']
sub_ethnicity_values = ['White', 'Black or African American', 'Asian', 'Unknown']
sub_sample_blank_values = ['null', 'Null', '']
sub_project_type_values = ['Family', 'Case-Control']
cast_plate_box_values = ['Plate', 'Box']


class record(object):
    """
    using built in dictionary and grabbing
    each row
    """
    def __init__(self, mapping):
        self.__dict__.update(mapping)


    def __repr__(self):
        return '%s(%r)' % (self.__class__.__name__, self.__dict__)
                                    


class logfile(object):
    
    def __init__(self):
        print "hello"

        
    def create_log(self, logname):
        current_data = datetime.date.today()
        logger_filename = logname + "_" +  str(current_data) + ".log"
        return logger_filename
        

    def clean_log(self, log):
        current_dir = os.path.dirname(os.path.realpath('__file__'))
        if os.path.isfile(log):
            MessageBox = ctypes.windll.user32.MessageBoxA
            MessageBox(None, 'Removing old log file', 'Couch Lab Database',0)
            os.remove(log)
        else:
            print "nothing to remove"


class excel:

    def __init__(self, index):
        self.index = int(index)

    def get_sheet(self, fname, index):
        self.fname = fname
        workbook = xlrd.open_workbook(fname,formatting_info=True)
        sheet_number = workbook.nsheets
        sheet_names = workbook.sheet_names()
        sheet = workbook.sheet_by_index(index)
        return sheet

    
    def headers(self, sheet):
        self.sheet = sheet
        header_dict = {col_index:sheet.cell(0, col_index).value.encode('ascii','ignore') 
                       for col_index in xrange(sheet.ncols)}
        return header_dict

    def __repr__(self):
        return self.encode('ascii','ignore')


    def excel_value_list(self, sheet, header_dict, x, y):
        dict_list = []
        self.sheet = sheet
        self.x = x
        self.y = y
        for row_index in range(x,y):
            d = {header_dict[col_index]: sheet.cell(row_index, col_index).value.encode('ascii','ignore')
                 for col_index in xrange(sheet.ncols)}
            dict_list.append(d)
        row = [record(i) for i in dict_list]
        return dict_list, row 
                                

    def table_last_row(self, sheet):
        self.sheet = sheet
        it = iter(range(sheet.nrows))
        last = next(it)
        for val in it:
            yield last, True
            last = val
        yield last, False


    def chop_table(self, sheet, index):
        table_coord = []
        value = excel.table_last_row(self, sheet)
        self.sheet = sheet
        self.index = index
        for i, t in value:
            if t == False:
                for x in  xrange(0, i, index):
                    parts = int(round(float(i)/float(index)))
                    table_coord = [(x, x+index) for x in xrange(0, int(i))][::index]
        return table_coord
                
        
    def missing_fields(self, sheet):
        self.sheet = sheet
        for r in xrange(0,sheet.nrows):
            for c in xrange(0, sheet.ncols):
                if sheet.cell(r,c).ctype == 0 or  sheet.cell(r,c).ctype == 6:
                    print "Missing Information on row %s, column name : %s," %(r, c)

                    
class validate_excel:

    def __init__(self, sheet, header):
        self.sheet = sheet
        self.index = header


    def is_important(self, headers, headers_list):
        for head in headers:
            if head in required_field_list:
                req = [head]
                return req

            
    def check_int(self,sheet, index):
        try:
            n = str(float(s))
            if n == "nan" or n=="inf" or n=="-inf" :
                return False
        except ValueError:
            try:
                complex(s)
            except ValueError:
                return False
        return True

        
    def check_unique(self, sheet, column, x, y):
        unique_list = []
        self.sheet = sheet
        self.column = column
        self.x = x
        self.y = y
        for r in xrange(x +1, y):
            value = sheet.cell(r,column).value.encode('ascii','ignore')
            if value not in unique_list:
                unique_list.append(value)
            else:
                print "Non-unique value found :%s in row %s" % (value, r) 

                
    def check_in_dropdown(self, sheet, drop_down_list, column):
        self.sheet = sheet
        self.column = column
        for r in xrange(1, sheet.nrows):
            value = sheet.cell(r,column).value.encode('ascii','ignore')
            if value in drop_down_list:
                pass
            else:
                print "value not is lookup :%s in row %s " % (value, r) 
                                                                                                

    def convert_date(self, sheet, column):
        self.sheet = sheet
        self.column = column
        for x, i in enumerate(sheet.col(column)):
            if x > 0:
                print i
#                py_date = xlrd.xldate.xldate_from_date_tuple(i.value, 3)
#            else:
#                print "Date recieved not correct : found %s" % (i.value)

                
                    
if __name__ == '__main__':
    mylog = logfile()
    t = mylog.create_log("hello")
    n = mylog.clean_log("Three_table_tube_info-2016-01-08.log")
    sheet = excel(0)
    carrier_sheet = sheet.get_sheet("input/carriers_TEST2_2015-12-22.xls", 0)
    carrier_header = sheet.headers(carrier_sheet)
    carrier_value = sheet.excel_value_list(carrier_sheet, carrier_header, 1, carrier_sheet.nrows)
    box_coord = sheet.chop_table(carrier_sheet, 96)
    validate_carrier = validate_excel(carrier_sheet, 0)
    unique_value_check = validate_carrier.check_unique(carrier_sheet, 0, 0, 96)
    check_required = validate_carrier.is_important(carrier_header, required_field_list)
    carrier_check_dropdown = validate_carrier.check_in_dropdown(carrier_sheet, sub_gender_values, 3)
    cast_plate_sheet = sheet.get_sheet("input/carriers_TEST2_2015-12-22.xls", 2)
    validate_cast_plate = validate_excel(cast_plate_sheet,0)
    cast_plate_header = sheet.headers(cast_plate_sheet)
    print cast_plate_header
    get_date = validate_cast_plate.convert_date(cast_plate_sheet, 2)
    print get_date
    #    carrier_value =  validate_carrier.table_value(carrier_sheet, carrier_header, 'CARRIERS ID')
    
