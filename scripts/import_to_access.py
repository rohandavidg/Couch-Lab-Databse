#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

"""
This script imports the validated induvidual excel files
into the couch lab access database
"""

import pyodbc
import pprint
import logging
import win32com.client as win32
import ctypes
import sys
from excel_parse import configure_logger
import datetime
import xlrd
import os
from validating_excel_workbook import get_excel_sheet
import traceback


current_date = datetime.date.today()

logger_filename = "run_import-" + str(current_date) + ".log"

current_dir = os.path.dirname(os.path.realpath('__file__'))
 

def main():
    excel_workbook = sys.argv[1]
    run(excel_workbook, logger_filename)

    
def run(excel_workbook, logger_filename):
    logger = configure_logger(logger_filename)
    check_log = check_validation_log(current_dir)
    workbook, carrier_id_sheet, caid_cast_sheet, cast_plate_sheet = get_excel_sheet(excel_workbook, logger)
    import_carrier_id_table = connect_to_access(carrier_id_sheet, caid_cast_sheet, cast_plate_sheet)


def check_validation_log(current_dir):
    """
    checking the validation log to see if there are 
    still debug messages
    """
    items = os.listdir(current_dir)
    log_file = []
    for names in items:
        if names.startswith("validation") and names.endswith(".log"):
            log_file.append(os.path.join(current_dir, names))

    for path in log_file:
        with open(path) as pin:
            r = pin.read()
            if "DEBUG" in r:
                MessageBox = ctypes.windll.user32.MessageBoxA
                MessageBox(None, 'validation log file has "DEBUG" messages', 'Couch Lab Database',0)
                sys.exit("files did not validate")
            else:
                pass
    
    
def connect_to_access(carrier_id_sheet, caid_cast_sheet, cast_plate_sheet):
    """
    connect to database and import each table
    """
    DBfile = 'c:\\Users\m149947\Desktop\couch\CARRIERS\database\CARRIERS_SubManifestOnly.accdb'
    conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+DBfile)
    cursor = conn.cursor()
    try:
        carriers_id_table = import_carriers_id_table(cursor, carrier_id_sheet)
        cast_plate_table = import_cast_plate(cursor, cast_plate_sheet)
        caid_cast_table = import_caid_cast_table(cursor, caid_cast_sheet)
        MessageBox = ctypes.windll.user32.MessageBoxA
        MessageBox(None, 'SUCCESS!!!Tables have been imported', 'Couch Lab Database',0)
        cursor.close()
        conn.commit()
        conn.close()
    except pyodbc.IntegrityError:
        traceback.print_exc()
        MessageBox = ctypes.windll.user32.MessageBoxA
        MessageBox(None, "Microsoft error: Check log file" , 'Couch Lab Database',0)
        sys.exit(1)


    
def import_carriers_id_table(cursor, sheet):
    """
    importing carrier id table
    """
    SQL = """
    INSERT INTO "CARRIERS ID" ([CARRIERS ID], [Sub_Sample Name], [Sub_Individual ID], [Sub_Gender], [Sub_Sample Status], [Sub_Pedigree], [Sub_Mother ID], [Sub_Father ID], [Sub_Disease Type], [Sub_Race], [Sub_Ethnicity], [CARRIERS ID Comment]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """

    for r in xrange(1, sheet.nrows):
        CARRIERS_ID = sheet.cell(r,0).value
        Sub_Sample_Name = sheet.cell(r,1).value
        Sub_Individual_ID = sheet.cell(r,2).value
        Sub_Gender = sheet.cell(r,3).value
        Sub_Sample_Status = sheet.cell(r,4).value
        Sub_Pedigree = sheet.cell(r,5).value
        Sub_Mother_ID = sheet.cell(r,6).value
        Sub_Father_ID = sheet.cell(r,7).value
        Sub_Disease_Type = sheet.cell(r,8).value
        Sub_Race = sheet.cell(r,9).value
        Sub_Ethnicity= sheet.cell(r,10).value
        CARRIERS_ID_Comment = sheet.cell(r,11).value

        values = (CARRIERS_ID, Sub_Sample_Name, Sub_Individual_ID, Sub_Gender, Sub_Sample_Status,
                  Sub_Pedigree, Sub_Mother_ID, Sub_Father_ID, Sub_Disease_Type, Sub_Race, Sub_Ethnicity, CARRIERS_ID_Comment)
        
        cursor.execute(SQL, values)

def import_caid_cast_table(cursor, sheet):
    """
    import caid cast table
    """
    SQL =  """
    INSERT INTO "CAID_CAST" ([CARRIERS ID], [CAST Barcode], [Sub_Coord],[Sub_Vol], [Sub_Conc], [Sub_Alias], [Sub_Site of Origin], [Sub_Tissue Source],[Sub_Sample Blank], [Sub_Comment], [CAID_CAST Comment]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """
    for r in xrange(1, sheet.nrows):
#        CAID_CAST_KEY = sheet.cell(r,0).value
        CARRIERS_ID = sheet.cell(r,1).value
        CAST_BARCODE = sheet.cell(r,2).value
        Sub_Coord = sheet.cell(r,3).value
        Sub_Vol = sheet.cell(r,4).value
        Sub_Conc = sheet.cell(r,5).value
        Sub_Alias = sheet.cell(r,6).value
        Sub_Site_of_origin = sheet.cell(r,7).value
        Sub_Tissue_Source = sheet.cell(r,8).value
        Sub_Sample_Blank = sheet.cell(r,9).value
        Sub_Comment = sheet.cell(r,10).value
        CAID_CAST_Comment = sheet.cell(r,11).value
        
        values = (CARRIERS_ID, CAST_BARCODE, Sub_Coord, Sub_Vol, Sub_Conc, Sub_Alias, Sub_Site_of_origin,
                  Sub_Tissue_Source, Sub_Sample_Blank, Sub_Comment, CAID_CAST_Comment)

        cursor.execute(SQL, values)

def import_cast_plate(cursor, sheet):
    """
    import cast plate table
    """
    SQL = """
    INSERT INTO "CAST Plate" ([CAST Barcode], [CAST Plate/Box], [Date Received], [Sub_Contact ID], [Sub_Contact Person],[Sub_Contact E-mail], [Sub_Project Type], [Sub_Plate Name], [Sub_Plate Description],[CAST Plate Location], [CAST Plate comment]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
    """

    for r in xrange(1, sheet.nrows):
        CAST_Barcode = sheet.cell(r,0).value
        CAST_plate_box = sheet.cell(r,1).value
        Date_received = sheet.cell(r,2).value
        sub_contact_id = sheet.cell(r,3).value
        sub_contact_person = sheet.cell(r,4).value
        sub_contact_email = sheet.cell(r,5).value
        sub_project_type = sheet.cell(r,6).value
        sub_plate_name = sheet.cell(r,7).value
        sub_plate_description = sheet.cell(r,8).value
        cast_plate_location = sheet.cell(r,9).value
        cast_plate_comment = sheet.cell(r,10).value

        values = (CAST_Barcode, CAST_plate_box, Date_received, sub_contact_id, sub_contact_person, sub_contact_email,
                  sub_project_type, sub_plate_name, sub_plate_description, cast_plate_location, cast_plate_comment)
        cursor.execute(SQL, values)
    
if __name__ == '__main__':
    main()
