#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

import xlrd
import xlwt
import sys
import logging
import pprint
import xlsxwriter
from ctypes import *
import win32com.client as win32
from excel_parse import configure_logger
from excel_parse import record
from excel_parse import get_each_data_fields


sample_status_accepted_values = [ 'Case', 'Control', 'Proband', 'Family member']
sub_race_values = ['Hispanic or Latino', 'Unknown', 'Non-Hispanic/Latino']


def main():
    logger = configure_logger()
    excel_parsed_manifest = sys.argv[1]
    run(excel_parsed_manifest, logger, sample_status_accepted_values, sub_race_values)

    
def run(excel_parsed_manifest, logger, sample_status_accepted_values, sub_race_values):
    carrier_id_sheet, caid_cast_sheet, cast_plate_sheet = get_excel_sheet(excel_parsed_manifest, logger)
    carrier_id_field_type = check_fields(carrier_id_sheet, logger)
    carrier_id_record = index_row_header(carrier_id_sheet, logger)
    validate_carrier_record = carrier_id_validate(carrier_id_record, logger, sample_status_accepted_values, sub_race_values)
    
    
def get_excel_sheet(manifest, logger):
    workbook = xlrd.open_workbook(manifest,formatting_info=True)
    sheet_number = workbook.nsheets
    sheet_names = workbook.sheet_names()
    CARRIERS_ID_sheet = workbook.sheet_by_index(0)
    CAID_CAST_sheet = workbook.sheet_by_index(1)
    CAST_plate_sheet = workbook.sheet_by_index(2)
    logger.info('%s -> CARRIERS ID table',sheet_names[0])
    logger.info('%s -> CAID_CAST table',sheet_names[1])
    logger.info('%s -> CAST_plate ID table',sheet_names[2])    
    return (CARRIERS_ID_sheet, CAID_CAST_sheet, CAST_plate_sheet)


def check_fields(sheet, logger):
    for rowx in xrange(sheet.nrows):
        for colx in xrange(sheet.ncols):
            missing_field_check(sheet, rowx, colx, logger)
            

def missing_field_check(sheet, rowx, colx, logger):
    c = sheet.cell(rowx, colx)
    header = [sheet.cell(0, col_index).value for col_index in xrange(sheet.ncols)]
    xf = sheet.book.xf_list[c.xf_index]
    fmt_obj = sheet.book.format_map[xf.format_key]
#    print rowx, colx, unicode(repr(c.value)), c.ctype, \
#        fmt_obj.type, fmt_obj.format_key, fmt_obj.format_str
    if c.ctype == 0:
        logger.warn("table %s - Missing Information on row %s, column name : %s",sheet, rowx+1, header[colx])



def index_row_header(sheet, logger):
    header = [sheet.cell(0, col_index).value for col_index in xrange(sheet.ncols)]
    new_header = [i.replace(" ", "_") for i in header]    
    dict_list = get_each_data_fields(sheet, new_header, 1, sheet.nrows)
    target = [record(i) for i in dict_list]
    return target


def carrier_id_validate(target, logger, sample_status_accepted_values, sub_race_values):
    for l in target:
        try:
            sample_status_drop_down = check_lookup_values(l.Sub_Sample_Status.encode('ascii','ignore'), sample_status_accepted_values, logger)
        except AttributeError:
            logger.warn("Number found in text field: %s", l.Sub_Sample_Status)

        try:
            sub_race_drop_down = check_lookup_values(l.Sub_Race.encode('ascii','ignore'), sub_race_values, logger) 
        except AttributeError:
            logger.warn("Number found in text field: %s", l.Sub_Race)
            

def check_lookup_values(sample_value, sample_value_lookup, logger):
    for i in xrange(len(sample_value_lookup)):
        if sample_value != sample_value_lookup[i]:
            logger.warn("Not an accepted value: %s", sample_value)
        else:
            pass

        
        
def is_number(s):
    try:
        n = str(float(s))
        if n == "nan" or n=="inf" or n=="-inf" :
            return False
    except ValueError:
        try:
            complex(s) # for complex
        except ValueError:
            return False
    return True
                                
        
        

#def check_digits(value):
#    if value.isdigit():
#        return int(value)
#    else:
#        pass
        

if __name__ == '__main__':
    main()
