"""

Created by: Nathan Starkweather
Created on: 02/10/2016
Created in: PyCharm Community Edition


"""
__author__ = 'Nathan Starkweather'

from .xlcom import *
from .xladdress import *
from . import xlcom
from . import xladdress
import win32com
import win32com.client

class Constants():
    def __init__(self):
        # based on win32com.client.constants.__getattr__
        c = win32com.client.constants
        for d in c.__dicts__:
            for k in d:
                if k.startswith("xl"):
                    setattr(self, k, d[k])

    def __getattr__(self, attr):
        val = getattr(win32com.client.constants, attr)
        setattr(self, attr, val)
        return val

xlc = Constants()