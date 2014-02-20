# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Tue Oct 15 11:41:36 2013
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
class PlotArea(DispatchBaseClass):
	CLSID = IID('{000208CB-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ClearFormats(self):
		return self._ApplyTypes_(112, 1, (12, 0), (), 'ClearFormats', None,)

	def Select(self):
		return self._ApplyTypes_(235, 1, (12, 0), (), 'Select', None,)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		# Method 'Border' returns object of type 'Border'
		"Border": (128, 2, (9, 0), (), "Border", '{00020854-0000-0000-C000-000000000046}'),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'Fill' returns object of type 'ChartFillFormat'
		"Fill": (1663, 2, (9, 0), (), "Fill", '{00024435-0000-0000-C000-000000000046}'),
		# Method 'Format' returns object of type 'ChartFormat'
		"Format": (116, 2, (9, 0), (), "Format", '{000244B2-0000-0000-C000-000000000046}'),
		"Height": (123, 2, (5, 0), (), "Height", None),
		"InsideHeight": (1670, 2, (5, 0), (), "InsideHeight", None),
		"InsideLeft": (1667, 2, (5, 0), (), "InsideLeft", None),
		"InsideTop": (1668, 2, (5, 0), (), "InsideTop", None),
		"InsideWidth": (1669, 2, (5, 0), (), "InsideWidth", None),
		# Method 'Interior' returns object of type 'Interior'
		"Interior": (129, 2, (9, 0), (), "Interior", '{00020870-0000-0000-C000-000000000046}'),
		"Left": (127, 2, (5, 0), (), "Left", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"Position": (133, 2, (3, 0), (), "Position", None),
		"Top": (126, 2, (5, 0), (), "Top", None),
		"Width": (122, 2, (5, 0), (), "Width", None),
		"_InsideHeight": (2657, 2, (5, 0), (), "_InsideHeight", None),
		"_InsideLeft": (2654, 2, (5, 0), (), "_InsideLeft", None),
		"_InsideTop": (2655, 2, (5, 0), (), "_InsideTop", None),
		"_InsideWidth": (2656, 2, (5, 0), (), "_InsideWidth", None),
	}
	_prop_map_put_ = {
		"Height": ((123, LCID, 4, 0),()),
		"InsideHeight": ((1670, LCID, 4, 0),()),
		"InsideLeft": ((1667, LCID, 4, 0),()),
		"InsideTop": ((1668, LCID, 4, 0),()),
		"InsideWidth": ((1669, LCID, 4, 0),()),
		"Left": ((127, LCID, 4, 0),()),
		"Position": ((133, LCID, 4, 0),()),
		"Top": ((126, LCID, 4, 0),()),
		"Width": ((122, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000208CB-0000-0000-C000-000000000046}", PlotArea )
