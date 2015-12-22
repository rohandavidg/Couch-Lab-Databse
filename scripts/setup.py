from distutils.core import setup
import py2exe
setup(windows=[{"script":"excel_parse.py"}], options={"py2exe":{"includes":["sip"]}})
