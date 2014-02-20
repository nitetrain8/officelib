# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Thu Dec 12 15:12:13 2013
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

from win32com.client import DispatchBaseClass
class Border(DispatchBaseClass):
	CLSID = IID('{00020854-0000-0000-C000-000000000046}')
	coclass_clsid = None

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"Color": (99, 2, (12, 0), (), "Color", None),
		"ColorIndex": (97, 2, (12, 0), (), "ColorIndex", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		"LineStyle": (119, 2, (12, 0), (), "LineStyle", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"ThemeColor": (2365, 2, (12, 0), (), "ThemeColor", None),
		"TintAndShade": (2366, 2, (12, 0), (), "TintAndShade", None),
		"Weight": (120, 2, (12, 0), (), "Weight", None),
	}
	_prop_map_put_ = {
		"Color": ((99, LCID, 4, 0),()),
		"ColorIndex": ((97, LCID, 4, 0),()),
		"LineStyle": ((119, LCID, 4, 0),()),
		"ThemeColor": ((2365, LCID, 4, 0),()),
		"TintAndShade": ((2366, LCID, 4, 0),()),
		"Weight": ((120, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00020854-0000-0000-C000-000000000046}", Border )
