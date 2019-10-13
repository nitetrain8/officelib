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
        pass

    def __getattr__(self, attr):
        val = getattr(win32com.client.constants, attr)
        setattr(self, attr, val)
        return val

xlc = Constants()