#!/cygdrive/c/Users/m149947/AppData/Local/Continuum/Anaconda2-32/python

import pickle


#def init():
carrier_headers = ('CARRIERS ID', 'Sub_Sample Name', 'Sub_Individual ID', 'Sub_Gender',
                   'Sub_Disease Type', 'Sub_Race', 'Sub_Ethnicity', 'CARRIERS ID Comment')
carrier_output_table = "carrier_ID"

caid_cast_header = ('CAID_CAST_KEY', 'CARRIERS ID', 'CAST Barcode', 'Sub_Coord', 'Sub_Vol',
                    'Sub_Conc', 'Sub_Alias', 'Sub_Site of Origin', 'Sub_Tissue Source',
                    'Sub_Sample Blank', 'Sub_Comment', 'CAID_CAST Comment')
caid_cast_output_table = 'CAID_CAST'
    
cast_plate_header = ('CAST Barcode', 'CAST Plate/Box', 'Date Received', 'Sub_Contact ID'
                         'Sub_Contact_Person','Sub_Contact E-mail', 'Sub_Project Type', 'Sub_Plate Name',
                         'Sub_Plate Description', 'CAST Plate Location', 'CAST Plate comment')
cast_plate_output_table = 'CAST Plate'

manifest_fields = ['Box_Barcode', 'Box_Name', 'Box_Coordinate', 'Sample_Name',
                   'Individual_ID', 'Gender', 'Sample_Status', 'ACS_Match_Set',
                   'Pedigree', 'Mother_ID', 'Father_ID', 'Vol_ul', 'Conc_ngul',
                   'Alias', 'Site_of_Origin', 'Tissue_Source', 'Disease_Type', 'Ethnicity', 'Race',
                   'Comment', 'Plate_Name', 'Plate_Description', 'Coord',  'Sample_Blank', 'Default_Control']

carrier_id_table_name_fields = ['Individual_ID', 'Sample_Status', 'Gender',
                                'Pedigree', 'ACS_Match_Set', 'Mother_ID', 'Father_ID',
                                'Disease_Type', 'Race', 'Ethnicity']

caid_plate_table_fields = ['coord', 'Box_Barcode', 'Plate_Name',
                           'Box_Name', 'Plate_Description',
                           'Box_Coordinate', 'Vol_ul', 'Conc_ngul',
                           'Alias', 'Site_of_Origin',
                           'Tissue_Source', 'Comment', 'Box_Coordinate', 'Sample_Blank']
          
