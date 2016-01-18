#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

"""
This script is going to do magic
"""

import logging
import re
import pprint 
import xlrd
import xlwt
import win32com.client  as win32
import pyodbc
import sys
import datetime
import ctypes
from settings import *
from excel_parse import record
from excel_parse import configure_logger
from excel_parse import get_excel_sheet
from excel_parse import get_each_data_fields
from excel_parse import show_odbc_sources
from excel_parse import connect_db
from excel_parse import generate_carrier_id
from excel_parse import generate_excel_output


current_date = datetime.date.today()
logger_filename = "Three_table_tube_info-" + str(current_date) + ".log"
regex = r'[?|*|.|!|(|)|/|-]'

def main():
    submission_manifest = sys.argv[1]
    run(submission_manifest)
    

def run(submission_manifest):
    logger = configure_logger(logger_filename)
    manifest = get_excel_sheet(submission_manifest, logger)
    data_headers, data_row = get_data_headers(manifest, regex)    
    carrier_table_headers, caid_cast_headers = get_confusing_header(data_headers, carrier_id_table_name_fields,
                                                 caid_plate_table_fields)
    get_data_index = data_fields_index(manifest, data_row)
    carriers_dict, caid_plate_dict, carrier_sample  = sample_data(manifest,
                                                                  data_headers,
                                                                  get_data_index,
                                                                  carrier_table_headers,
                                                                  caid_cast_headers)
    pprint.pprint(carriers_dict)
    cast_tuple_output, contact_id = cast_plate_parse(manifest)
    carrier_id_number = connect_db(logger)
    create_carrier_ID, sample_number = generate_carrier_id(carrier_sample, logger, carrier_id_number)
    carrier_tuple_output, caid_tuple_output =  create_tuple_output(carriers_dict, caid_plate_dict,
                                                                   create_carrier_ID, sample_number)

    excel_output = generate_excel_output(carrier_output_table, carrier_headers, carrier_tuple_output,
                                         caid_cast_output_table, caid_cast_header,
                                         caid_tuple_output, cast_plate_output_table,
                                         cast_plate_header, cast_tuple_output, current_date, contact_id)
    
def get_data_headers(sheet, regex):
    """
    getting only unique Coord, as each sheet has multiple
    boxes
    """
    data_headers = []
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        if row[3].value == "Sample Name *":
            data_start = rowidx +1
            for colx in range(sheet.ncols):
                data_headers.append(re.sub(regex,r'',row[colx].value.encode('ascii', 'ignore')).strip().replace(" ", "_"))
    return (data_headers, data_start)


def data_fields_index(sheet, data_start):
    data_index = []
    for rowidx in range(data_start, sheet.nrows):
        row = sheet.row(rowidx)
        data_index.append((rowidx, sheet.nrows))
    return data_index


def get_confusing_header(data_headers, carrier_id_table_name_fields,
                         caid_plate_table_fields):
    carrier_table_headers = []
    caid_cast_table_headers = []
    for i in data_headers:
        if i in carrier_id_table_name_fields:
            carrier_table_headers.append(i)
        if i in caid_plate_table_fields:
            caid_cast_table_headers.append(i)
    return (carrier_table_headers, caid_cast_table_headers)



def sample_data(sheet,new_header, data_index,
                carrier_table_headers, caid_cast_table_headers):
    """
    going throw each row and making a dictionary out
    of it creating table specific dictionaries
    """
    carrierID_dict = {}
    caid_plate_dict = {}
    Sample_Name = []
    needed = data_index[0]
    data_list = get_each_data_fields(sheet, new_header, needed[0], needed[1])
    target = [record(r) for r in data_list]
    for l in target:
        if l.Sample_Name != '':
            Sample_Name.append(l.Sample_Name)
            carrierID_dict[l.Sample_Name] = [v for k, v in l.__dict__.iteritems() if k in carrier_table_headers]
            caid_plate_dict[l.Sample_Name] = [v for k, v in l.__dict__.iteritems() if k in caid_cast_table_headers]
    return(carrierID_dict, caid_plate_dict, Sample_Name)


def cast_plate_parse(sheet):
    contact_id = []
    contact_person = []
    contact_email = []
    project_type = []
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        if row[0].value == "Contact ID (Study Acronym)*":
            contact_id.append(row[2].value.encode('ascii', 'ignore'))
        elif row[0].value == "Contact Person*":
            contact_person.append(row[2].value.encode('ascii', 'ignore'))
        elif row[0].value == "Contact Email*":
            contact_email.append(row[2].value.encode('ascii', 'ignore'))
        elif row[0].value == "Project Type*":
            project_type.append(row[2].value.encode('ascii', 'ignore'))
        else:
            pass
    cast_plate_tuple = zip(contact_id, contact_person, contact_email, project_type)
    collaborator_name = contact_id[0]
    return (cast_plate_tuple, collaborator_name)
            

def create_tuple_output(carrierID_dict, caid_plate_dict, sample_name_carrier_id, set_length):
    """
    generating a list of each row for each table
    """
    carrier_id_table = []
    caid_id_table = []
    cast_id_table = []
    for carrier_key, carrier_value in  carrierID_dict.items():
        if sample_name_carrier_id[carrier_key]:
            carrier_id_string = ((str(sample_name_carrier_id[carrier_key]),
                                  str(carrier_key), str(",".join(str(i) for i in carrier_value[:])),''))
            carr_tup = carrier_id_string[:2] + tuple(carrier_id_string[2].split(","))
            carrier_id_table.append(carr_tup)


    for caid_key, caid_value in caid_plate_dict.items():
        if sample_name_carrier_id[caid_key]:
            caid_id_string = ('', str(sample_name_carrier_id[caid_key]),
                              '', ','.join(str(i) for i in caid_value))
            caid_tup = caid_id_string[:3] + tuple(caid_id_string[-1].split(",")[:-2])\
                       + tuple([caid_id_string[-1].split(',')]) + tuple('')
            caid_id_table.append(caid_tup)

    return(carrier_id_table, caid_id_table)
                                                    
    
if __name__ == '__main__':
    main()
