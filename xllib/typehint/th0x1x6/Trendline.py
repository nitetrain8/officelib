# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Thu Oct 24 13:53:50 2013
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
class Trendline(DispatchBaseClass):
	CLSID = IID('{000208BE-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ClearFormats(self):
		return self._ApplyTypes_(112, 1, (12, 0), (), 'ClearFormats', None,)

	def Delete(self):
		return self._ApplyTypes_(117, 1, (12, 0), (), 'Delete', None,)

	def Select(self):
		return self._ApplyTypes_(235, 1, (12, 0), (), 'Select', None,)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"Backward": (185, 2, (3, 0), (), "Backward", None),
		"Backward2": (2650, 2, (5, 0), (), "Backward2", None),
		# Method 'Border' returns object of type 'Border'
		"Border": (128, 2, (9, 0), (), "Border", '{00020854-0000-0000-C000-000000000046}'),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'DataLabel' returns object of type 'DataLabel'
		"DataLabel": (158, 2, (9, 0), (), "DataLabel", '{000208B2-0000-0000-C000-000000000046}'),
		"DisplayEquation": (190, 2, (11, 0), (), "DisplayEquation", None),
		"DisplayRSquared": (189, 2, (11, 0), (), "DisplayRSquared", None),
		# Method 'Format' returns object of type 'ChartFormat'
		"Format": (116, 2, (9, 0), (), "Format", '{000244B2-0000-0000-C000-000000000046}'),
		"Forward": (191, 2, (3, 0), (), "Forward", None),
		"Forward2": (2651, 2, (5, 0), (), "Forward2", None),
		"Index": (486, 2, (3, 0), (), "Index", None),
		"Intercept": (186, 2, (5, 0), (), "Intercept", None),
		"InterceptIsAuto": (187, 2, (11, 0), (), "InterceptIsAuto", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		"NameIsAuto": (188, 2, (11, 0), (), "NameIsAuto", None),
		"Order": (192, 2, (3, 0), (), "Order", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"Period": (184, 2, (3, 0), (), "Period", None),
		"Type": (108, 2, (3, 0), (), "Type", None),
	}
	_prop_map_put_ = {
		"Backward": ((185, LCID, 4, 0),()),
		"Backward2": ((2650, LCID, 4, 0),()),
		"DisplayEquation": ((190, LCID, 4, 0),()),
		"DisplayRSquared": ((189, LCID, 4, 0),()),
		"Forward": ((191, LCID, 4, 0),()),
		"Forward2": ((2651, LCID, 4, 0),()),
		"Intercept": ((186, LCID, 4, 0),()),
		"InterceptIsAuto": ((187, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"NameIsAuto": ((188, LCID, 4, 0),()),
		"Order": ((192, LCID, 4, 0),()),
		"Period": ((184, LCID, 4, 0),()),
		"Type": ((108, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000208BE-0000-0000-C000-000000000046}", Trendline )
