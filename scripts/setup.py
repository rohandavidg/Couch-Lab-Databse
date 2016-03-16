from distutils.core import setup
import sys
import numpy
import pandas
sys.setrecursionlimit(5000)
import py2exe
import matplotlib
import glob
import os
from distutils.filelist import findall
import zmq.libzmq

matplotlibdatadir = matplotlib.get_data_path()
matplotlibdata = findall(matplotlibdatadir)
matplotlibdata_files = []
#for f in matplotlibdata:
#        print f
#        dirname = os.path.join('matplotlibdata', f[len(matplotlibdatadir)+1:])
#        matplotlibdata_files.append((os.path.split(dirname)[0], [f]))
 


setup(options = {
    "py2exe":
        {
            "includes":['sip'],
            "compressed": False,
            "excludes": ["zmq.libzmq"], 
            "dll_excludes": ["MSVCP90.dll","HID.DLL", "w9xpopen.exe", "libzmq.pyd"]
        }
    },
    data_files=matplotlib.get_py2exe_datafiles(),
    windows=[{'script': 'redcap_formatter.py'}]
    )
