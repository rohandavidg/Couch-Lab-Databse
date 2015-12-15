#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

"""
This script imports the validated induvidual excel files
into the couch lab access database
"""

import pyodbc
import pprint
import logging
import win32com.client as win32
import sys
from excel_parse import configure_logger
import datetime
import xlrd
from validating_excel_workbook import get_excel_sheet


current_date = datetime.date.today()

logger_filename = "run_import-" + str(current_date) + ".log"

#carrier_headers = 

def main():
    excel_workbook = sys.argv[1]
    run(excel_workbook, logger_filename)

    
def run(excel_workbook, logger_filename):
    logger = configure_logger(logger_filename)
    workbook, carrier_id_sheet, caid_cast_sheet, cast_plate_sheet = get_excel_sheet(excel_workbook, logger)
    import_carrier_id_table = connect_to_access(carrier_id_sheet)

    
def connect_to_access(sheet):
    DBfile = 'c:\\Users\m149947\Desktop\couch\CARRIERS\database\CARRIERS_SubManifestOnly.accdb'
    conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ='+DBfile)
    cursor = conn.cursor()
    SQL = """
INSERT INTO "CARRIERS ID" ([CARRIERS ID], [Sub_Sample Name], [Sub_Individual ID], Sub_Gender, [Sub_Sample Status], Sub_Pedigree, [Sub_Mother ID], [Sub_Father ID], [Sub_Disease Type], Sub_Race, Sub_Ethnicity, [CARRIERS ID Comment]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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

        values = (CARRIERS_ID, Sub_Sample_Name, Sub_Individual_ID, Sub_Gender, Sub_Sample_Status, Sub_Pedigree, Sub_Mother_ID, Sub_Father_ID, Sub_Disease_Type, Sub_Race, Sub_Ethnicity, CARRIERS_ID_Comment)
        
        cursor.execute(SQL, values)
        
    cursor.close()
    conn.close()
    
if __name__ == '__main__':
    main()
