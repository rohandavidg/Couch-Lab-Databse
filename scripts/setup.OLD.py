from distutils.core import setup
import py2exe
setup(windows=[{"script":"excel_parse.py"}], options={"py2exe":{"includes":["sip"]}})
setup(windows=[{"script":"validating_excel_workbook.py"}], options={"py2exe":{"includes":["sip"]}})
setup(windows=[{"script":"import_to_access.py"}], options={"py2exe":{"includes":["sip"]}})
