#!/usr/bin/python
# -*- coding: utf-8 -*-

'''
formats the submission manifests into
a redCAp compatiable format
'''

import sys
print sys.getdefaultencoding()
sys.setrecursionlimit(1500)
import xlrd
import xlwt
import logging
import re
import datetime
import csv
import itertools
from collections import defaultdict
import collections
import pandas as pd
import pprint
import os
reload(sys)
sys.setdefaultencoding('utf-8')


version = "v1.7"
date = datetime.date.today()
plate_manifest = os.path.join(os.path.dirname("__file__"), 'plate_stock_manifest.txt')
#plate_manifest = 'c:/Users/m149947/Desktop/couch/CARRIERS/database/test/plate_stock_manifest.txt'
#box_manifest = 'c:/Users/m149947/Desktop/couch/CARRIERS/database/test/box_stock_manifest.txt' 
box_manifest = os.path.join(os.path.dirname("__file__"), 'box_stock_manifest.txt')
logger_filename = "RedCap_formatter-" + str(date) + ".log"
regex = r'[?|*|.|!|(|)|/|-]'
#redcap_mapping_file = 'C:/Users/m149947/Desktop/couch/CARRIERS/database/test/redcap_mapping_file.txt'
redcap_mapping_file = os.path.join(os.path.dirname("__file__"), 'redcap_mapping_file.txt')

complete_plate_dict = {'caqc_conc_complete':'0', 'caqc_dnaqual_complete':'0', 'cawk_complete':'0', 'capc_complete': '0', 'casq_complete':'0', 'caqc_vol':'0', 'cawk_plate_vol_stock': '0', 'capc_plate_vol_cawk': '0', 'capc_plate_vol_final':'0', 'capc_vol_qc':'0', 'caup_plate_vol':'0', 'cabp_vol_capc':'0', 'cabp_complete':'0','bio_complete':'0', 'db_script_version': version}
complete_box_dict = {'caqc_conc_complete':'0', 'caqc_dnaqual_complete':'0', 'cawk_complete':'0', 'capc_complete': '0', 'casq_complete':'0', 'caqc_vol':'0', 'cawk_plate_vol_stock': '0', 'capc_plate_vol_cawk': '0', 'capc_plate_vol_final':'0', 'capc_qc_complete': '0', 'capc_vol_qc':'0', 'caup_plate_vol':'0','cabp_vol_capc':'0', 'cabp_complete':'0', 'cast_tube_transfer': '1','bio_complete':'0', 'db_script_version': version}

#print complete_plate_dict

def main():
    submission_manifest = sys.argv[1]
    run(submission_manifest)


def run(submission_manifest):
    logger = configure_logger(logger_filename)
    manifest = get_excel_sheet(submission_manifest, logger)
    excel_headers, data_start_int = get_data_headers(manifest, regex)
    check_manifest = fork_by_headers(excel_headers)
    get_data_index = data_fields_index(manifest, data_start_int)
    header_mapper_dict = check_headers(check_manifest, redcap_mapping_file, regex, logger)
    sample_name, annotate_excel_file_dict, divide = sample_data(manifest, excel_headers,
                                                        get_data_index, header_mapper_dict)
    red_cap_empty_fields, out_headers = empty_dict(check_manifest, box_manifest, plate_manifest)
#    print red_cap_empty_fields #out_headers
    plate_headers_create = plate_headers_dict(check_manifest, manifest, sample_name, divide)
    contact_dict = normalize_all_dict(manifest, sample_name, red_cap_empty_fields)
    combination_dict = combine_contact_annotate(check_manifest, annotate_excel_file_dict,
                                                contact_dict, plate_headers_create)
    create_merged_list = merge_dict(combination_dict)
    out_tsv = write_out_tsv(create_merged_list, out_headers)
    sort_out = format_output_tsv(check_manifest)
    

class record(object):
    """
    using built in dictionary and grabbing
    each row
    """
    def __init__(self, mapping):
        self.__dict__.update(mapping)


    def __repr__(self):
        return '%s(%r)' % (self.__class__.__name__, self.__dict__)


def configure_logger(logger_filename):
    """
    setting up logging
    """
    logger = logging.getLogger('Couchlab')
    logger.setLevel(logging.DEBUG)
    handler = logging.FileHandler(logger_filename)
    handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter("%(asctime)s'\t'%(name)s'\t'%(levelname)s'\t'%(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger


class headers:

    def __init__(self, header_file):
        self.header_file = header_file
        self.headers = []
        self.plate_dict = {}
        self.box_dict = {}
        self.acs_dict = {}
        
    def index_headers(self):
        with open(self.header_file, 'r') as f:
            for raw_line in f:
                line = raw_line.strip()
                self.headers.append(line)
        return self.headers

    
    def get_plate_dict(self, regex):
        with open(self.header_file, 'r') as f:
            for raw_line in f:
                line = raw_line.strip().split("\t")
                value_1 = re.sub(regex,r'',
                                 line[0].encode('ascii', 'ignore')).strip().replace(" ", "_")
                self.plate_dict[value_1] = line[1]
        return self.plate_dict

    
    def get_box_dict(self, regex):
        with open(self.header_file, 'r') as f:
            for raw_line in f:
                line = raw_line.strip().split("\t")
                value_2 = re.sub(regex,r'',
                                  line[2].encode('ascii', 'ignore')).strip().replace(" ", "_")
                self.box_dict[value_2] = line[3]
        return self.box_dict

    
    def get_acs_dict(self, regex):
        with open(self.header_file, 'r') as f:
            for raw_line in f:
                line = raw_line.strip().split("\t")
                value_3 = re.sub(regex,r'',
                                 line[4].encode('ascii', 'ignore')).strip().replace(" ", "_")
                self.acs_dict[value_3] = line[5]
        return self.acs_dict
        

    
def get_excel_sheet(submission_manifest, logger):
    """
    Parsing excel workbook
    """
    workbook = xlrd.open_workbook(submission_manifest, encoding_override='cp1252')
    sheet_number = workbook.nsheets
    sheet_names = workbook.sheet_names()
    sheet = workbook.sheet_by_index(0)
    logger.info("version 1.0 recap_formatter.py")
    logger.info('%s -> using this sheet to get info',sheet_names[0])
    return sheet


def get_data_headers(sheet, regex):
    """
    getting only unique Coord, as each sheet has multiple
    boxes
    """
    data_headers = []
    data_start = []
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        if row[3].value == "Sample Name *" or row[1].value == "Sample Name *":
            data_start.append(rowidx +1)
            for colx in range(sheet.ncols):
                data_headers.append(re.sub(regex,r'',
                                           row[colx].value.encode('ascii',
                                                                  'ignore')).strip().replace(" ", "_"))
    return data_headers, data_start 


def fork_by_headers(data_headers):
    if "Box_Description" in data_headers:
        return "box_manifest"
    if "Box_Coordinate" in data_headers:
        return "acs_manifest"
    else:
        return "plate_manifest"
        

def check_headers(manifest, mapping_file,regex, logger):
    mapping = headers(redcap_mapping_file)
    if manifest == "box_manifest":
        mapping_dict = mapping.get_box_dict(regex)
        return mapping_dict
    elif manifest == "plate_manifest":
        mapping_dict = mapping.get_plate_dict(regex)
        return mapping_dict
    elif manifest == "acs_manifest":
        mapping_dict = mapping.get_acs_dict(regex)
        return mapping_dict
    else:
        logger.debug("%s headers not either box or plate", manifest)
        

        
def data_fields_index(sheet, data_start):
    data_index = []
    start = data_start[0]
    for rowidx in range(start, sheet.nrows):
        row = sheet.row(rowidx)
        data_index.append((rowidx, sheet.nrows))
    return data_index

    
class contact:

    def __init__(self, sheet):
        self.sheet = sheet
        self.contact_id = {}
        self.person = {}
        self.email = {}
        self.project_type = {}
        self.plate_name = {}
        self.plate_desc = {}

        
    def get_contact_id(self):
        sheet = self.sheet
        for rowidx in range(1, sheet.nrows):
            if sheet.cell_value(rowidx, 0) == 'Contact ID (Study Acronym)*':
                self.contact_id['sub_study_id'] = sheet.cell_value(rowidx, 2)
        return self.contact_id

    
    def get_contact_person(self):
        sheet = self.sheet
        for rowidx in range(1, sheet.nrows):        
            if sheet.cell_value(rowidx, 0) == 'Contact Person*':
                self.person['sub_contact'] = sheet.cell_value(rowidx, 2)
        return self.person

    
    def get_contact_email(self):            
        sheet = self.sheet
        for rowidx in range(1, sheet.nrows):
            if sheet.cell_value(rowidx, 0) == 'Contact Email*':
                self.email['sub_contact_email'] = sheet.cell_value(rowidx, 2)
        return self.email

    
    def get_project_type(self):
        sheet = self.sheet
        for rowidx in range(1, sheet.nrows):
            if sheet.cell_value(rowidx, 0) == 'Project Type*':
                self.project_type['sub_project'] = sheet.cell_value(rowidx, 2)
        return self.project_type


    def get_plate_name(self):
        sheet = self.sheet
        for rowidx in range(1, sheet.nrows):
            if sheet.cell_value(rowidx,0) == 'Plate Name *':
                self.plate_name.setdefault('sub_plate_name', []).append(sheet.cell_value(rowidx,2))
        return self.plate_name


    def get_plate_description(self):
        sheet = self.sheet
        for rowidx in range(1, sheet.nrows):
            if sheet.cell_value(rowidx, 0) == 'Plate Description':
                self.plate_desc.setdefault('sub_plate_desc', []).append(sheet.cell_value(rowidx,2))
        return self.plate_desc

    
def get_each_data_fields(sheet, header, x, y):
    """
    indexing headers with field values
    """
    dict_list = []
    for row_index in range(x,y):
        d = {header[col_index]: sheet.cell(row_index,
                                           col_index).value for col_index in xrange(sheet.ncols)}
        dict_list.append(d)
    return dict_list
                                    

def sample_data(sheet, header, data_index, header_dict):
    """
    going throw each row and making a dictionary out
    of it creating table specific dictionaries
    """
    abl_data_dict = {}
    sample_name = []
    needed = data_index[0]
    data_list = get_each_data_fields(sheet, header, needed[0], needed[1])
    target = [record(r) for r in data_list]
    chop_data = True
    for l in target:
        if l.Sample_Name != '' and l.Sample_Name != 'Sample Name *':
            sample_name.append(l.Sample_Name)
            abl_data_dict[l.Sample_Name] = [{header_dict[k]:v} for k, v in l.__dict__.iteritems() if k in header_dict.keys()]
        try:
            if l.Coord == 'H12':
                if not l.Sample_Name:
                    chop_data = True
                else:
                    pass
            else:
                pass
        except AttributeError:
            pass
                
    return sample_name, abl_data_dict, chop_data


def empty_dict(fork, box_manifest, plate_manifest):
    if fork == "box_manifest" or fork == "acs_manifest":
        header = headers(box_manifest)
        box_headers = header.index_headers()
        empty_box_headers = box_headers[:5] + ["cast_box_barcode"] + ['cast_buffer'] + box_headers[-24:]
        redcap_empty_list = [{k:complete_box_dict[k]} if k in complete_box_dict.keys() else {k:""} for k in empty_box_headers]
#        pprint.pprint(redcap_empty_list)
        return redcap_empty_list, box_headers
    else:
        fork == "plate_manifest"
        header = headers(plate_manifest)
        plate_headers = header.index_headers()
        empty_plate_headers = plate_headers[:5] + ['cast_plate_barcode'] + ['cast_buffer'] + plate_headers[-24:]
        redcap_empty_list = [{k:complete_plate_dict[k]} if k in complete_plate_dict.keys() else {k:""} for k in empty_plate_headers]
        return redcap_empty_list, plate_headers


def plate_headers_dict(fork, sheet, sample_name, chop_data):
    end = [int(95) if chop_data == True else int(96)]
    plate_head_dict = {}
    if fork == "plate_manifest":
        info = contact(sheet)    
        plate_name = info.get_plate_name()
        plate_desc = info.get_plate_description()
        ranges = [(n, min(n+end[0], len(sample_name))) for n in xrange(0, len(sample_name), end[0])]
        for i, x in enumerate(ranges):
            for y in xrange(x[0], x[1]):
                plate_head_dict[sample_name[y]] = [{'sub_plate_name': plate_name['sub_plate_name'][i]},
                                                   {'sub_plate_desc': plate_desc['sub_plate_desc'][i]}]
    else:
        pass
    return plate_head_dict


def normalize_all_dict(sheet, sample_name, redcap_empty_list):
    info = contact(sheet)
    study_contact_id = info.get_contact_id()
    study_person = info.get_contact_person()
    study_email = info.get_contact_email()
    study_project_type = info.get_project_type()
    sample_contact_dict = {}
    for sample in sample_name:
        sample_contact_dict[sample] = [study_contact_id, study_person, study_email,
                                       study_project_type] +  redcap_empty_list
    return sample_contact_dict


def combine_contact_annotate(fork, abl_data_dict, sample_contact_dict, plate_head_dict):
    dd = defaultdict(list)
    if fork == "plate_manifest":
        for d in (abl_data_dict, sample_contact_dict, plate_head_dict):
            for key, value in d.iteritems():
                dd[key].extend(value)
    else:
        for d in (abl_data_dict, sample_contact_dict):
            for key, value in d.iteritems():
                dd[key].extend(value)
    return dd

    
def merge_dict(dd):
    out_list = []
    for key, value in dd.items():
        new_value = dict(kv for d in value for kv in d.iteritems())
        out_list.append(new_value)
    return out_list


def write_out_tsv(out_list, headers):
    with open("test_out.tsv", 'wb') as fout:
        fieldnames = headers
        writer = csv.DictWriter(fout, fieldnames=fieldnames, extrasaction='ignore', delimiter="\t")
        writer.writeheader()
        for row in out_list:
            writer.writerow(row)


def format_output_tsv(fork):
    if fork == "plate_manifest":
        df = pd.read_table('test_out.tsv')
        df = df.sort_values(by=['sub_plate_name', 'sub_plate_coordinate'], ascending=[True, True])
        df.to_csv('output_sorted.csv', index=False)
#        os.remove('test_out.tsv')
    else:
        df = pd.read_csv('test_out.tsv', delimiter="\t")
        df = df.sort_values(by=['sub_box_name', 'sub_box_coordinate'], ascending=[True, True])
        df.to_csv('output_sorted.csv', index=False)
#        os.remove('test_out.tsv')
        
if __name__ ==  '__main__':
    main()
