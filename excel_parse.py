#!/users/bin/python

"""
This codes parses the submission manifest and creates a ID for each unique Sample_Status
and create three excel files based on the Jenna's database schema
"""

import xlrd
import xlwt
from ctypes import *
from xlrd.sheet import ctype_text
import win32com.client as win32
import platform
import pyodbc
import pprint
import logging
import logging.config
import re
import sys

a =0
carrier_headers = ('CARRIERS ID', 'sub_Sample Name', 'sub_individual ID', 'sub Gender', 'sub_Sample status','sub_pedigree', 'Sub_Mother ID', 'Sub_Father ID', 'Sub_Disease Type', 'SubRace', 'Sub_Ethnicity', 'CARRIERS ID comment')
carrier_output_table = "carrier_ID"

caid_cast_header = ('CAID_CAST_KEY', 'CARRIERS_ID', 'CAST_Barcode', 'Sub_coord', 'Sub_vol', 'Sub_conc', 'Sub_Alias', 'Sub_site of origin', 'Sub_Tissue source', 'Sub_Sample Blank', 'Sub_default control', 'CAID_cast comment')
caid_cast_output_table = 'CAID_CAST'

cast_plate_header = ('CAST Barcode', 'CAST Plate/Box', 'Date Received', 'Sub_contact ID', 'Sub_contact_person', 'Sub_Plate Name', 'Sub_Plate Description', 'Manifest File', 'FileData', 
                        'FileFlags', 'FileTimeStamp', 'FileType', 'FileURL', 'CAST Plate Location', 'CAST Plate Comment')
cast_plate_output_table = 'CAST Plate'


class record(object):
    def __init__(self, mapping):
        self.__dict__.update(mapping)

    def __repr__(self):
        return '%s(%r)' % (self.__class__.__name__, self.__dict__) 


def configure_logger():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    handler = logging.FileHandler('parsing_excel.log')
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger


def main():
    logger = configure_logger()
    submission_manifest = sys.argv[1]
    run(submission_manifest, logger, carrier_headers, carrier_output_table, caid_cast_header, caid_cast_output_table, cast_plate_header, cast_plate_output_table)


def run(submission_manifest, logger, carrier_headers, carrier_output_table, caid_cast_header, caid_cast_output_table, cast_plate_header, cast_plate_output_table):
    manifest = get_excel_sheet(submission_manifest, logger)
    cast_plate_dict = create_cast_plate(manifest)
    fields_header = get_data_fields_header(manifest) 
    index = data_fields_index(manifest)
    carrierID_field, caid_plate_field, carrier_sample =  sample_data(manifest, fields_header, index)
    create_carrier_ID, sample_number = generate_carrier_id(carrier_sample, logger, a)
    carrier_tuple_output, caid_tuple_output, cast_tuple_output = create_tuple_output(carrierID_field, caid_plate_field, cast_plate_dict, create_carrier_ID, sample_number)
    excel_output = generate_excel_output(carrier_output_table,carrier_headers, carrier_tuple_output, caid_cast_output_table, caid_cast_header,
                                        caid_tuple_output, cast_plate_output_table, cast_plate_header, cast_tuple_output)


def get_excel_sheet(submission_manifest, logger):
    workbook = xlrd.open_workbook(submission_manifest, encoding_override='cp1252')
    sheet_number = workbook.nsheets
    sheet_names = workbook.sheet_names()
    sheet = workbook.sheet_by_index(0)
    logger.info('%s -> using this sheet to get info',sheet_names[0])
    return sheet


def create_index(sheet, field, x, y, z):
    cast_plate = []
    if field in sheet.cell(x,y).value:
        field_acronym = sheet.cell(x,y).value
        field_acronym = re.sub(r'[?|*|.|!|(|)|/]',r'',field_acronym).strip()
        new_field_acronym = field_acronym.replace(" ", "_")
        try:
            field = sheet.cell(x,z).value
        except ValueError:
            logger.debug('%s has no value',field)
            field = ''
        cast_plate.append((new_field_acronym, field))
    else:
        print "missing"
    return cast_plate


def create_cast_plate(sheet):
    cast_plate_dict = {}
    Plate_Barcode = create_index(sheet, "Plate Barcode", 21,0,2)
    contact_id = create_index(sheet,"Contact ID (Study Acronym)*", 15,0,2)
    contact_person = create_index(sheet, "Contact Person", 16,0,2)
    contact_email = create_index(sheet, "Contact Email*", 17,0,2)
    plate_name = create_index(sheet,"Plate Name *", 22,0,2)
    plate_description =  create_index(sheet,"Plate Description", 23,0,2)
    comb = contact_id + contact_person + Plate_Barcode + contact_email + plate_name + plate_description
    cast_plate_dict = dict(comb)
    return cast_plate_dict


def get_data_fields_header(sheet):
    header = [sheet.cell(26, col_index).value for col_index in xrange(sheet.ncols)]
    new_header = []
    for i  in header:
        nstr = re.sub(r'[?|*|.|!|(|)|/]',r'',i).strip()            
        new_value = nstr.replace(" ","_")
        new_header.append(new_value)
    return new_header



def data_fields_index(sheet):
    data_index = []
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        if row[0].value =="Coord *":
            data_index.append((rowidx+1, int(rowidx)+ 97))
    return data_index



def sample_data(sheet,new_header, data_index):
    carrierID_dict = {}
    caid_plate_dict = {}
    Sample_Name = []
    for i in data_index:
        a = get_each_data_fields(sheet, new_header, i[0], i[1])
        target = [record(r) for r in a]
        for l in target:
            if l.Sample_Name != '':
                Sample_Name.append(l.Sample_Name)
                carrierID_dict[l.Sample_Name] = [l.Individual_ID, l.Gender, l.Sample_Status, l.Pedigree, l.Mother_ID,
                                                 l.Father_ID, l.Disease_Type, l.Race, l.Ethnicity, l.Comment]
                caid_plate_dict[l.Sample_Name] = [l.Coord, l.Vol_ul, l.Conc_ngul, l.Alias, l.Site_of_Origin, l.Tissue_Source, 
                                                  l.Sample_Blank, l.Default_Control, l.Comment]
     
    return(carrierID_dict, caid_plate_dict, Sample_Name)



def get_each_data_fields(sheet, new_header, x, y):
    dict_list = []
    for row_index in range(x,y):
        d = {new_header[col_index]: sheet.cell(row_index, col_index).value for col_index in xrange(sheet.ncols)}
        dict_list.append(d)
    return dict_list



def generate_carrier_id(Sample_Name, logger, a):
    list_length  = len(Sample_Name)
    set_length = len(set(Sample_Name))
    assert list_length == set_length
    logger.info("number of samples in manifest %s", list_length)
    logger.info("number of unique samples in manifest %s", set_length)
    carrier_trunk = ["{0:08}".format(num) for num in xrange(a+1, set_length +1)]
    carrier_id = [''.join("CA" + str(i)) for i in carrier_trunk]
    sample_name_carrier_id = dict(zip(Sample_Name,carrier_id))
    return sample_name_carrier_id, set_length



def create_tuple_output(carrierID_dict, caid_plate_dict, cast_plate_dict, sample_name_carrier_id, set_length):
    carrier_id_table = []
    caid_id_table = []
    cast_id_table = []
    for carrier_key, carrier_value in  carrierID_dict.items():
        if sample_name_carrier_id[carrier_key]:
            carrier_id_string = ((str(sample_name_carrier_id[carrier_key]),  str(carrier_key), str(",".join(carrier_value[:]))))
            carr_tup = carrier_id_string[:2] + tuple(carrier_id_string[2].split(","))
            carrier_id_table.append(carr_tup)


    for caid_key, caid_value in caid_plate_dict.items():
        if sample_name_carrier_id[caid_key]:
            caid_id_string = ('', str(sample_name_carrier_id[caid_key]), '', ','.join(str(i) for i in caid_value))
            caid_tup = caid_id_string[:3] + tuple(caid_id_string[3].split(","))
            caid_id_table.append(caid_tup)

    cast_plate_row = (cast_plate_dict['Plate_Barcode'],'' , '', str(cast_plate_dict['Contact_ID_Study_Acronym']), str(cast_plate_dict['Contact_Person']), str(cast_plate_dict['Plate_Name']), str(cast_plate_dict['Plate_Description'])
                        ,"","","","","","","","")
    for i in range(set_length):
        cast_tup = tuple(cast_plate_row)
        cast_id_table.append(cast_tup)

    return(carrier_id_table, caid_id_table, cast_id_table)


def generate_excel_output(carrier_output_table, carrier_headers, carrier_id_table, caid_cast_output_table,caid_cast_header,
                            caid_id_table, cast_plate_output_table, cast_plate_header, cast_id_table):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(carrier_output_table)

    heading_xf = xlwt.easyxf('font: bold on; align: wrap on, vert centre, horiz center')
    rowx = 0
    for colx, value in enumerate(carrier_headers):
        ws.write(rowx, colx, value, heading_xf)

    for i, row in enumerate(carrier_id_table):
        for j, colx in enumerate(row):
            ws.write(i+1, j, colx)
    
    ws.col(0).width = 256 * max([len(row[0]) for row in carrier_id_table])
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(rowx+1)
    ws.set_remove_splits(True)
 #   wb.save("carriers.xls")
    ws = wb.add_sheet(caid_cast_output_table)
#   ws = wb.add_sheet(cast_input)
    heading_xf = xlwt.easyxf('font: bold on; align: wrap on, vert centre, horiz center')
    rowx = 0
    for colx, value in enumerate(caid_cast_header):
        ws.write(rowx, colx, value, heading_xf)

    for i, row in enumerate(caid_id_table):
        for j, colx in enumerate(row):
            ws.write(i+1, j, colx)
    
    ws.col(0).width = 256 * max([len(row[0]) for row in caid_id_table])
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(rowx+1)
    ws.set_remove_splits(True)

    ws = wb.add_sheet(cast_plate_output_table)
    heading_xf = xlwt.easyxf('font: bold on; align: wrap on, vert centre, horiz center')
    rowx = 0
    for colx, value in enumerate(cast_plate_header):
        ws.write(rowx, colx, value, heading_xf)

    for i, row in enumerate(cast_id_table):
        for j, colx in enumerate(row):
            ws.write(i+1, j, colx)
    
    ws.col(0).width = 256 * max([len(row[0]) for row in cast_id_table])
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(rowx+1)
    ws.set_remove_splits(True)

    wb.save("carriers.xls")

if __name__ == '__main__':
    main()
