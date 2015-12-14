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


logger_filename = "validation_info.log"
sample_status_accepted_values = [ 'Case', 'Control', 'Proband', 'Family Member']
sub_race_values = ['Hispanic or Latino', 'Unknown', 'Non-Hispanic/Latino']
sub_gender_values = ['M', 'F']
sub_ethnicity_values = ['White', 'Black or African American', 'Asian']


def main():
    excel_parsed_manifest = sys.argv[1]
    run(excel_parsed_manifest, sample_status_accepted_values,
        sub_race_values, sub_gender_values, sub_ethnicity_values)

    
def run(excel_parsed_manifest, sample_status_accepted_values,
        sub_race_values, sub_gender_values, sub_ethnicity_values):
    logger = configure_logger(logger_filename)
    carrier_id_sheet, caid_cast_sheet, cast_plate_sheet = get_excel_sheet(excel_parsed_manifest, logger)
    carrier_id_field_type = check_fields(carrier_id_sheet, logger)
    carrier_id_record = index_row_header(carrier_id_sheet, logger)
    validate_carrier_record = carrier_id_validate(carrier_id_record, logger, sample_status_accepted_values,
                                                  sub_race_values, sub_gender_values, sub_ethnicity_values)
    
    
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


def carrier_id_validate(target, logger, sample_status_accepted_values, sub_race_values,
                        sub_gender_values, sub_ethnicity_values):
    for l, x in enumerate(target):
        sample_status_drop_down = check_drop_down(x.Sub_Sample_Status, sample_status_accepted_values, logger, l+2, "sample_status")
        sub_race_drop_down = check_drop_down(x.Sub_Race, sub_race_values, logger, l+2, "Sub_race") 
        sub_gender_drop_down = check_drop_down(x.Sub_Gender, sub_gender_values, logger, l+2, "Sub_Gender")
        sub_ethnicity_drop_down = check_drop_down(x.Sub_Ethnicity, sub_ethnicity_values, logger, l+2, "Sub_Ethnicity")
        Disease_type = check_is_number(x.Sub_Disease_Type, logger, l+2, "Sub_Disease_Type")        
        induvidual_id = check_is_number(x.Sub_Individual_ID, logger, l+2, "Sub_Induvidual_ID")
        sample_name = check_is_number(x.Sub_Sample_Name, logger, l+2, "Sub_Sample_Name")
        mother_id = check_is_number(x.Sub_Mother_ID, logger, l+2, "Sub_Mother_ID")
        father_id = check_is_number(x.Sub_Father_ID, logger, l+2, "Sub_Father_ID")
        pedigree_id = is_number(x.Sub_Pedigree)
        if pedigree_id != True:
            logger.debug("Number found in text field: %s : in row %s : Coulumn name Sub_Pedigree", x.Sub_Pedigree, l+2)

            
def check_drop_down(value, look_up_list, logger, number, column_name):
    try:
        field_drop_down = check_lookup_values(value.encode('ascii','ignore'), look_up_list, logger, number, column_name)
    except AttributeError:
        logger.debug("Number found in text field: %s : in row %s : Column name : %s", value, number, column_name)


        
def check_is_number(value, logger, number, column_name):
    field = is_number(value.encode('ascii','ignore'))
    if field != False:
        logger.debug("Number found in text field: %s : in row %s : Coulumn name %s", value, number, column_name)
    else:
        pass


def check_lookup_values(sample_value, sample_value_lookup, logger, row_number, column_name):
    if sample_value not in sample_value_lookup:
        logger.debug("value not in lookup: %s : in row number %s : Column name %s", sample_value, row_number, column_name)
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
                                
        
        

if __name__ == '__main__':
    main()
