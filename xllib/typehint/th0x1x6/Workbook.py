# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct  7 13:28:00 2013
'Microsoft Excel 12.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30302f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{00020813-0000-0000-C000-000000000046}')
MajorVersion = 1
MinorVersion = 6
LibraryFlags = 8
LCID = 0x0

from win32com.client import CoClassBaseClass
import sys
__import__('win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x6.WorkbookEvents')
WorkbookEvents = sys.modules['win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x6.WorkbookEvents'].WorkbookEvents
__import__('win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x6._Workbook')
_Workbook = sys.modules['win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x6._Workbook']._Workbook
class Workbook(CoClassBaseClass): # A CoClass
	CLSID = IID('{00020819-0000-0000-C000-000000000046}')
	coclass_sources = [
		WorkbookEvents,
	]
	default_source = WorkbookEvents
	coclass_interfaces = [
		_Workbook,
	]
	default_interface = _Workbook

win32com.client.CLSIDToClass.RegisterCLSID( "{00020819-0000-0000-C000-000000000046}", Workbook )
