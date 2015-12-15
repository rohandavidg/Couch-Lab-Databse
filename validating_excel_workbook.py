#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

"""
This scripts validates each field in the each
induvidual excel tables 
"""

import xlrd
import xlwt
import sys
import logging
import pprint
from ctypes import *
import win32com.client as win32
from excel_parse import configure_logger
from excel_parse import record
from excel_parse import get_each_data_fields
import datetime 
import re

current_date = datetime.date.today()

logger_filename = "validation_info-" + str(current_date) + ".log"
sample_status_accepted_values = [ 'Case', 'Control', 'Proband', 'Family Member']
sub_race_values = ['Hispanic or Latino', 'Unknown', 'Non-Hispanic/Latino']
sub_gender_values = ['M', 'F']
sub_ethnicity_values = ['White', 'Black or African American', 'Asian']
sub_sample_blank_values = ['Yes', 'null', 'Null', 'yes']
sub_project_type_values = ['Family', 'case-control']
cast_plate_box_values = ['Plate', 'Box']


def main():
    excel_parsed_manifest = sys.argv[1]
    run(excel_parsed_manifest, sample_status_accepted_values,
        sub_race_values, sub_gender_values, sub_ethnicity_values,
        sub_sample_blank_values, sub_project_type_values, cast_plate_box_values)

    
def run(excel_parsed_manifest, sample_status_accepted_values,
        sub_race_values, sub_gender_values, sub_ethnicity_values,
        sub_sample_blank_values, sub_project_type_values, cast_plate_box_values):
    logger = configure_logger(logger_filename)
    workbook, carrier_id_sheet, caid_cast_sheet, cast_plate_sheet = get_excel_sheet(excel_parsed_manifest, logger)
    req_carrier_id_headers, req_caid_cast_headers, req_cast_plate_sheet = required_headers(carrier_id_sheet, caid_cast_sheet, cast_plate_sheet)
    carrier_id_empty_record = check_fields(carrier_id_sheet, logger, req_carrier_id_headers)
    carrier_id_record = index_row_header(carrier_id_sheet, logger)
    validate_carrier_record = carrier_id_validate(carrier_id_record, logger, sample_status_accepted_values,
                                                  sub_race_values, sub_gender_values, sub_ethnicity_values)
    caid_cast_empty_record = check_fields(caid_cast_sheet, logger, req_caid_cast_headers)
    caid_cast_record = index_row_header(caid_cast_sheet, logger)
    validate_caid_cast_record = caid_cast_validate(caid_cast_record, logger, sub_sample_blank_values)
    cast_plate_empty_record = check_fields(cast_plate_sheet, logger, req_cast_plate_sheet)
    cast_plate_record = index_row_header(cast_plate_sheet, logger)
    solve_date = get_date_field(workbook, cast_plate_sheet, logger)
    validate_cast_plate_record = cast_plate_validate(cast_plate_record, logger, sub_project_type_values, cast_plate_box_values)
    
    
def get_excel_sheet(manifest, logger):
    workbook = xlrd.open_workbook(manifest,formatting_info=True)
    sheet_number = workbook.nsheets
    sheet_names = workbook.sheet_names()
    CARRIERS_ID_sheet = workbook.sheet_by_index(0)
    CAID_CAST_sheet = workbook.sheet_by_index(1)
    CAST_plate_sheet = workbook.sheet_by_index(2)
    sys.stdout.write('%s -> CARRIERS ID table \n ' % (sheet_names[0]))
    sys.stdout.write('%s -> CAID_CAST table \n' % (sheet_names[1]))
    sys.stdout.write('%s -> CAST_plate ID table ' % (sheet_names[2]))    
    return (workbook,CARRIERS_ID_sheet, CAID_CAST_sheet, CAST_plate_sheet)


def required_headers(CARRIERS_ID_sheet, CAID_CAST_sheet, CAST_plate_sheet):
    required_carrier_header = [CARRIERS_ID_sheet.cell(0, col_index).value for col_index in xrange(CARRIERS_ID_sheet.ncols)][:-1]
    CAID_CAST_header = [CAID_CAST_sheet.cell(0, col_index).value for col_index in xrange(CAID_CAST_sheet.ncols)][:6]
    CAST_plate_header = [CAST_plate_sheet.cell(0, col_index).value for col_index in xrange(CAST_plate_sheet.ncols)][:7]
    return (required_carrier_header, CAID_CAST_header, CAST_plate_header)
    

def check_fields(sheet, logger, header):
    for rowx in xrange(sheet.nrows):
        for colx in xrange(sheet.ncols):
            missing_field_check(sheet, rowx, colx, header, logger)
            

def missing_field_check(sheet, rowx, colx, header, logger):
    c = sheet.cell(rowx, colx)
    col_index = [i for i in xrange(0,len(header))]
    xf = sheet.book.xf_list[c.xf_index]
    fmt_obj = sheet.book.format_map[xf.format_key]
#    print rowx, colx, unicode(repr(c.value)), c.ctype, \
#        fmt_obj.type, fmt_obj.format_key, fmt_obj.format_str
#    print unicode(repr(c.value))
    if c.ctype == 0 or c.ctype == 6:
        if colx in col_index:
            logger.warn("table %s - Missing Information on row %s, column name : %s",sheet, rowx+1, header[colx])



def index_row_header(sheet, logger):
    header = [sheet.cell(0, col_index).value for col_index in xrange(sheet.ncols)]
    header = [re.sub(r'[?|*|.|!|(|)|/|-]',r'',i).strip() for i in header]
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
        pedigree_id = check_is_number_true(x.Sub_Pedigree, logger, l+2, "Sub_Pedigree")

            
def check_drop_down(value, look_up_list, logger, number, column_name):
    try:
        field_drop_down = check_lookup_values(value.encode('ascii','ignore'), look_up_list, logger, number, column_name)
    except AttributeError:
        logger.debug("Number found in text field: %s : in row %s : Column name : %s", value, number, column_name)

        
def check_is_number(value, logger, number, column_name):
    try:
        field = is_number(value.encode('ascii','ignore'))
        if field != False:
            logger.debug("Number found in text field: %s : in row %s : Coulumn name %s", value, number, column_name)
        else:
            pass
    except AttributeError:
        logger.debug("Number found in text field: %s : in row %s : Coulumn name %s", value, number, column_name)

        
def check_is_number_true(value, logger, number, column_name):
    try:
        field = is_number(value.encode('ascii', 'ignore'))
        if field != True:
            logger.debug("Number not found in field: %s : in row %s : Column name %s", value, number, column_name)
        else:
            pass
    except AttributeError:
        logger.debug("Number not found in  field: %s : in row %s : Column name %s", value, number, column_name)
        

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
                                

def caid_cast_validate(target, logger, sub_sample_blank_values):
    for l, x in enumerate(target):
        cast_barcode = check_is_number(x.CAST_Barcode, logger, l+2, "CAST_Barcode")
        value_cast_barcode = check_cast_barcode(x.CAST_Barcode,logger,  l+2, "CAST_Barcode")
        sub_coord = check_is_number(x.Sub_Coord, logger, l+2, "Sub_Coord")
        sub_conc = check_is_number_true(x.Sub_Conc, logger, l+2, "Sub_Conc")
        sub_vol = check_is_number_true(x.Sub_Vol, logger, l+2, "Sub_Vol")
        sub_alias = check_is_number(x.Sub_Alias, logger, l+2, "Sub_Alias")


def check_cast_barcode(cast_barcode,logger,  number, column_name):
    try:
        cast_barcode_list = [ i.encode('ascii','ignore') for i in cast_barcode]
        four_mer = "".join(cast_barcode_list[:4])
        if four_mer != 'CAST':
            logger.debug("Cast barcode has error: %s : in row %s : column name %s", cast_barcode, number, column_name) 
    except TypeError:
        logger.debug("Number found in text field: %s : in row %s : Column name %s", cast_barcode, number, column_name)
        

def get_date_field(workbook, CAST_plate_sheet, logger):
    for x, i in enumerate(CAST_plate_sheet.col(2)):
        if x > 0:
            if str(i).split(':')[0] == 'xldate':
                py_date = xlrd.xldate.xldate_as_datetime(i.value, workbook.datemode)
            else:
                logger.debug("Date recieved not correct : found %s : in row %s : Column name: Date Received", i.value, x+1)

                
def check_email(email, logger, row, column_name):
    try:
        email_integ = re.search("@", email.encode('ascii','ignore'))
        if email_integ:
            pass
        else:
            logger.debug("email not in correct  format: %s : in row %s : Column name %s", email, row, column_name)
    except AttributeError:
        logger.debug("email not in correct format: %s : in row %s : column name %s", email, row, column_name)

        
def cast_plate_validate(target, logger, sub_project_type_values, cast_plate_box_values):
    for l, x in enumerate(target):
        cast_barcode = check_cast_barcode(x.CAST_Barcode, logger, l+2, "CAST_Barcode")
        contact_ID = check_is_number(x.Sub_Contact_ID, logger, l+2, "Sub_Contact_ID")
        plate_name = check_is_number(x.Sub_Plate_Name, logger, l+2, "Sub_Plate_Name")
        cast_plate_box = check_lookup_values(x.CAST_PlateBox, cast_plate_box_values, logger, l+2, "CAST_Plate/Box")
        contact_person = check_is_number(x.Sub_Contact_Person, logger, l+2, "Sub_contact_person")
        project_type = check_lookup_values(x.Sub_Project_Type, sub_project_type_values, logger, l+2, "Sub_Project_type")
        contact_email = check_is_number(x.Sub_Contact_Email, logger, l+2, "Sub_Contact_E-mail")
        check_email_integ = check_email(x.Sub_Contact_Email, logger, l+2, "Sub_Contact_E-mail")


    
if __name__ == '__main__':
    main()
