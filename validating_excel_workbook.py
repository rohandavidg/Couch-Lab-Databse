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
import ctypes
import win32com.client as win32
from excel_parse import configure_logger
from excel_parse import record
from excel_parse import get_each_data_fields
import datetime 
import re
import os



current_dir = os.path.dirname(os.path.realpath(__file__))
current_date = datetime.date.today()
logger_filename = "validation_info-" + str(current_date) + ".log"
sample_status_accepted_values = [ 'Case', 'Control', 'Proband', 'Family Member']
sub_race_values = ['Hispanic or Latino', 'Unknown', 'Non-Hispanic/Latino']
sub_gender_values = ['M', 'F']
sub_ethnicity_values = ['White', 'Black or African American', 'Asian', 'Unknown']
sub_sample_blank_values = ['null', 'Null', '']
sub_project_type_values = ['Family', 'Case-Control']
cast_plate_box_values = ['Plate', 'Box']
required_field_list = ['CAST Barcode', 'CAST Plate/Box', 'Date Received', 'Sub_Contact ID', 'Sub_Contact Person',
                       'Sub_Contact E-mail', 'Sub_Project Type', 'Sub_Plate Name', 'CARRIERS ID',
                       'Sub_Sample Name', 'Sub_Individual ID', 'Sub_Gender', 'Sub_Sample Status']

unique_field_list = ['CAST Barcode', 'Sub_Plate Name', 'CARRIERS ID', 'Sub_Sample Name', 'Sub_Coord']                     
                     

def main():
    excel_parsed_manifest = sys.argv[1]
    run(excel_parsed_manifest, sample_status_accepted_values,
        sub_race_values, sub_gender_values, sub_ethnicity_values,
        sub_sample_blank_values, sub_project_type_values,
        cast_plate_box_values, required_field_list)

    
def run(excel_parsed_manifest, sample_status_accepted_values,
        sub_race_values, sub_gender_values, sub_ethnicity_values,
        sub_sample_blank_values, sub_project_type_values, cast_plate_box_values,
        required_field_list):
    check_log = clean_logfile(current_dir, logger_filename)
    logger = configure_logger(logger_filename)
    workbook, carrier_id_sheet, caid_cast_sheet, cast_plate_sheet = get_excel_sheet(excel_parsed_manifest,
                                                                                    logger)
    req_carrier_id_headers, unique_carrier_id_headers, req_caid_cast_headers,\
    unique_caid_cast_headers, req_cast_plate_sheet, unique_cast_plate_sheet = required_headers(carrier_id_sheet,
                                                                                               caid_cast_sheet, cast_plate_sheet)
    carrier_id_empty_record = check_fields(carrier_id_sheet, logger, req_carrier_id_headers, "CARRIER ID")
    carrier_id_list, carrier_id_record = index_row_header(carrier_id_sheet, logger, "CARRIER ID")
    carrier_unique_field = count_fields(carrier_id_list, unique_carrier_id_headers, logger, "CARRIERS ID")
    modified_required_field = get_modified_required_list(required_field_list)
    validate_carrier_record = carrier_id_validate(carrier_id_record, logger, sample_status_accepted_values,
                                                  sub_race_values, sub_gender_values, sub_ethnicity_values,
                                                  "CARRIER ID", modified_required_field)
    caid_cast_empty_record = check_fields(caid_cast_sheet, logger, req_caid_cast_headers, "CAID_CAST")
    caid_cast_list, caid_cast_record = index_row_header(caid_cast_sheet, logger, "CAID_CAST")
    caid_cast_unique_field = count_fields(caid_cast_list, unique_caid_cast_headers, logger, "CAID_CAST")
    validate_caid_cast_record = caid_cast_validate(caid_cast_record, logger, sub_sample_blank_values, "CAID_CAST", modified_required_field)
    cast_plate_empty_record = check_fields(cast_plate_sheet, logger, req_cast_plate_sheet, "CAST Plate")
    cast_plate_list, cast_plate_record = index_row_header(cast_plate_sheet, logger, "CAST Plate")
    cat_plate_unique_field = count_fields(cast_plate_list, unique_cast_plate_sheet, logger, "CAST Plate")
    solve_date = get_date_field(workbook, cast_plate_sheet, logger, "CAST Plate")
    validate_cast_plate_record = cast_plate_validate(cast_plate_record, logger,
                                                     sub_project_type_values, cast_plate_box_values, "CAST Plate", modified_required_field)

def clean_logfile(current_dir, logger_filename):
    log_file = os.path.join(current_dir, logger_filename)
    if os.path.isfile(log_file):
        MessageBox = ctypes.windll.user32.MessageBoxA 
        MessageBox(None, 'Removing old log file', 'Couch Lab Database',0)
        os.remove(log_file)
    else:
        pass
    
    
def get_excel_sheet(manifest, logger):
    workbook = xlrd.open_workbook(manifest,formatting_info=True)
    sheet_number = workbook.nsheets
    sheet_names = workbook.sheet_names()
    CARRIERS_ID_sheet = workbook.sheet_by_index(0)
    CAID_CAST_sheet = workbook.sheet_by_index(1)
    CAST_plate_sheet = workbook.sheet_by_index(2)
    sys.stdout.write('%s -> CARRIERS ID table \n ' % (sheet_names[0]))
    sys.stdout.write('%s -> CAID_CAST table \n' % (sheet_names[1]))
    sys.stdout.write('%s -> CAST_plate ID table \n ' % (sheet_names[2]))    
    return (workbook,CARRIERS_ID_sheet, CAID_CAST_sheet, CAST_plate_sheet)


def required_headers(CARRIERS_ID_sheet, CAID_CAST_sheet, CAST_plate_sheet):
    required_carrier_header = [CARRIERS_ID_sheet.cell(0, col_index).value for col_index in xrange(CARRIERS_ID_sheet.ncols)][:-1]
    unique_carrier_header = [header for header in required_carrier_header if header in unique_field_list]
    CAID_CAST_header = [CAID_CAST_sheet.cell(0, col_index).value for col_index in xrange(CAID_CAST_sheet.ncols)][0:6]
    unique_cast_header = [header for header in CAID_CAST_header if header in unique_field_list]
    CAST_plate_header = [CAST_plate_sheet.cell(0, col_index).value for col_index in xrange(CAST_plate_sheet.ncols)][:8]
    unique_cast_plate_header = [header for header in CAST_plate_header if header in unique_field_list]
    return (required_carrier_header, unique_carrier_header, CAID_CAST_header, unique_cast_header,
            CAST_plate_header, unique_cast_plate_header)
    

def check_fields(sheet, logger, header, table_name):
    for rowx in xrange(sheet.nrows):
        for colx in xrange(sheet.ncols):
            missing_field_check(sheet, rowx, colx, header, logger, table_name)



def missing_field_check(sheet, rowx, colx, header, logger, table_name):
    c = sheet.cell(rowx, colx)
    col_index = [i for i in xrange(0,len(header))]
    xf = sheet.book.xf_list[c.xf_index]
    fmt_obj = sheet.book.format_map[xf.format_key]
    if c.ctype == 0 or c.ctype == 6:
        if colx in col_index:
            if header[colx] == "CAID_CAST_KEY":
                pass
            else:
                logger.debug("Missing Information on row %s, column name : %s, Table Name %s",rowx+1, header[colx], table_name)
            
            

def index_row_header(sheet, logger, table_name):
    header = [sheet.cell(0, col_index).value for col_index in xrange(sheet.ncols)]
    new_header = modify_header(header)
    dict_list = get_each_data_fields(sheet, new_header, 1, sheet.nrows)
    target = [record(i) for i in dict_list]
    return (dict_list, target)


def modify_header(header):
    header = [re.sub(r'[?|*|.|!|(|)|/|-]',r'',i).strip() for i in header]
    new_header = [i.replace(" ", "_") for i in header]
    return new_header


def count_fields(dict_list, unique_header, logger, table_name):
    new_unique_header = modify_header(unique_header)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 0, 95)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 95, 191)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 191, 287)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 287, 383)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 383, 479)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 479, 575)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 575, 671)
    check_unique_instance(dict_list, new_unique_header, logger, table_name, 671, 767)

    
def check_unique_instance(dict_list, new_unique_header, logger, table_name, value1, value2):
    unique_list = []
    for x, d in enumerate(dict_list):
        if x > value1 and x < value2:
            for i in new_unique_header:
                if d[i] not in unique_list:
                    unique_list.append(d[i])
                else:
                    if table_name =='CAID_CAST' and i == "CAST_Barcode":
                        pass
                    else:
                        logger.debug("Non-unique value found : %s : in row %s : Column name %s : Table Name %s", d[i], x+1, i, table_name)
                        
                    
def get_modified_required_list(required_field_list):
    modified_required_list = modify_header(required_field_list)
    return modified_required_list


def carrier_id_validate(target, logger, sample_status_accepted_values, sub_race_values,
                        sub_gender_values, sub_ethnicity_values, table_name, modified_required_list):
    for l, x in enumerate(target):
        sample_status_drop_down = check_drop_down(x.Sub_Sample_Status, sample_status_accepted_values,
                                                  logger, l+2, "Sub_Sample_Status", table_name, modified_required_list)
        sub_race_drop_down = check_drop_down(x.Sub_Race, sub_race_values, logger, l+2, "Sub_race", table_name, modified_required_list) 
        sub_gender_drop_down = check_drop_down(x.Sub_Gender, sub_gender_values, logger, l+2, "Sub_Gender", table_name,
                                               modified_required_list)
        sub_ethnicity_drop_down = check_drop_down(x.Sub_Ethnicity, sub_ethnicity_values, logger, l+2, "Sub_Ethnicity",
                                                  table_name, modified_required_list)
        Disease_type = check_is_number(x.Sub_Disease_Type, logger, l+2, "Sub_Disease_Type", table_name, modified_required_list)        
        induvidual_id = check_is_number(x.Sub_Individual_ID, logger, l+2, "Sub_Individual_ID", table_name, modified_required_list)
        sample_name = check_is_number(x.Sub_Sample_Name, logger, l+2, "Sub_Sample_Name", table_name, modified_required_list)
        mother_id = check_is_number(x.Sub_Mother_ID, logger, l+2, "Sub_Mother_ID", table_name, modified_required_list)
        father_id = check_is_number(x.Sub_Father_ID, logger, l+2, "Sub_Father_ID", table_name, modified_required_list)
        pedigree_id = check_is_number_true(x.Sub_Pedigree, logger, l+2, "Sub_Pedigree", table_name)

            
def check_drop_down(value, look_up_list, logger, number, column_name, table_name, modified_required_list):
    try:
        field_drop_down = check_lookup_values(value.encode('ascii','ignore'), look_up_list, logger, number,
                                              column_name, table_name, modified_required_list)
    except AttributeError:
        log_required_field(modified_required_list,value, logger, number, column_name, table_name)


def log_required_field(modified_required_list,value, logger, number, column_name, table_name):
    if column_name in modified_required_list:
        logger.debug("Number found in text field: %s : in row %s : Column name %s : Table Name %s",value, number, column_name, table_name)
    else:
        logger.warn("Number found in text field: %s : in row %s : Column name %s : Table Name %s", value, number, column_name, table_name)
        

        
def check_is_number(value, logger, number, column_name, table_name, modified_required_list):
    try:
        field = is_number(value.encode('ascii','ignore'))
        if field != False:
            log_required_field(modified_required_list, value, logger, number, column_name, table_name)
        else:
            pass
    except AttributeError:
        log_required_field(modified_required_list, value, logger, number, column_name, table_name)

        
def check_is_number_true(value, logger, number, column_name, table_name):
    try:
        field = is_number(value.encode('ascii', 'ignore'))
        if field != True:
            logger.warn("Number not found in field: %s : in row %s : Column name %s : Table Name %s", value, number, column_name, table_name)
        else:
            pass
    except AttributeError:
        logger.warn("Number not found in  field: %s : in row %s : Column name %s : Table Name %s", value, number, column_name, table_name)
        

def check_lookup_values(sample_value, sample_value_lookup, logger, row_number, column_name, table_name, modified_required_list):
    if sample_value not in sample_value_lookup:
        if column_name in modified_required_list:
            logger.debug("value not in lookup: %s : in row number %s : Column name %s : Table Name %s", sample_value, row_number, column_name, table_name)            
        else:
            logger.warn("value not in lookup: %s : in row number %s : Column name %s : Table Name %s", sample_value, row_number, column_name, table_name)            
    else:
        pass
    
        
def is_number(s):
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


def caid_cast_validate(target, logger, sub_sample_blank_values, table_name, modified_required_list):
    for l, x in enumerate(target):
        cast_barcode = check_is_number(x.CAST_Barcode, logger, l+2, "CAST_Barcode", table_name, modified_required_list)
        value_cast_barcode = check_cast_barcode(x.CAST_Barcode,logger,  l+2, "CAST_Barcode", table_name)
        sub_coord = check_is_number(x.Sub_Coord, logger, l+2, "Sub_Coord", table_name, modified_required_list)
        sub_conc = check_is_number_true(x.Sub_Conc, logger, l+2, "Sub_Conc", table_name)
        sub_vol = check_is_number_true(x.Sub_Vol, logger, l+2, "Sub_Vol", table_name)
        sub_alias = check_is_number(x.Sub_Alias, logger, l+2, "Sub_Alias", table_name, modified_required_list)
        sub_sample_blank = check_drop_down(x.Sub_Sample_Blank, sub_sample_blank_values, logger, l+2, "Sub_Sample_Blank",
                                           table_name, modified_required_list) 
        

def check_cast_barcode(cast_barcode,logger,  number, column_name, table_name):
    try:
        cast_barcode_list = [ i.encode('ascii','ignore') for i in cast_barcode]
        four_mer = "".join(cast_barcode_list[:4])
        if four_mer != 'CAST':
            logger.debug("Cast barcode has error: %s : in row %s : column name %s : Table Name %s", cast_barcode, number, column_name, table_name) 
    except TypeError:
        logger.debug("Number found in text field: %s : in row %s : Column name %s : Table Name %s", cast_barcode, number, column_name, table_name)
        

def get_date_field(workbook, CAST_plate_sheet, logger, table_name):
    for x, i in enumerate(CAST_plate_sheet.col(2)):
        if x > 0:
            if str(i).split(':')[0] == 'xldate':
                py_date = xlrd.xldate.xldate_as_datetime(i.value, workbook.datemode)
            else:
                logger.debug("Date recieved not correct : found %s : in row %s : Column name: Date Received : Table Name %s", i.value, x+1, table_name)

                
def check_email(email, logger, row, column_name, table_name):
    try:
        email_integ = re.search("@", email.encode('ascii','ignore'))
        if email_integ:
            pass
        else:
            logger.debug("email not in correct  format: %s : in row %s : Column name %s : Table Name %s", email, row, column_name, table_name)
    except AttributeError:
        logger.debug("email not in correct format: %s : in row %s : column name %s : Table Name %s", email, row, column_name, table_name)

        
def cast_plate_validate(target, logger, sub_project_type_values, cast_plate_box_values, table_name, modified_required_list):
    for l, x in enumerate(target):
        cast_barcode = check_cast_barcode(x.CAST_Barcode, logger, l+2, "CAST_Barcode", table_name)
        contact_ID = check_is_number(x.Sub_Contact_ID, logger, l+2, "Sub_Contact_ID", table_name, modified_required_list)
        plate_name = check_is_number(x.Sub_Plate_Name, logger, l+2, "Sub_Plate_Name", table_name, modified_required_list)
        cast_plate_box = check_lookup_values(x.CAST_PlateBox, cast_plate_box_values, logger, l+2,
                                             "CAST_PlateBox", table_name, modified_required_list)
        contact_person = check_is_number(x.Sub_Contact_Person, logger, l+2, "Sub_Contact_Person", table_name, modified_required_list)
        project_type = check_lookup_values(x.Sub_Project_Type, sub_project_type_values,
                                           logger, l+2, "Sub_Project_Type", table_name, modified_required_list)
        contact_email = check_is_number(x.Sub_Contact_Email, logger, l+2, "Sub_Contact_Email", table_name, modified_required_list)
        check_email_integ = check_email(x.Sub_Contact_Email, logger, l+2, "Sub_Contact_E-mail", table_name)


    
if __name__ == '__main__':
    main()
