# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Tue Nov 26 15:13:46 2013
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
class ChartArea(DispatchBaseClass):
	CLSID = IID('{000208CC-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Clear(self):
		return self._ApplyTypes_(111, 1, (12, 0), (), 'Clear', None,)

	def ClearContents(self):
		return self._ApplyTypes_(113, 1, (12, 0), (), 'ClearContents', None,)

	def ClearFormats(self):
		return self._ApplyTypes_(112, 1, (12, 0), (), 'ClearFormats', None,)

	def Copy(self):
		return self._ApplyTypes_(551, 1, (12, 0), (), 'Copy', None,)

	def Select(self):
		return self._ApplyTypes_(235, 1, (12, 0), (), 'Select', None,)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"AutoScaleFont": (1525, 2, (12, 0), (), "AutoScaleFont", None),
		# Method 'Border' returns object of type 'Border'
		"Border": (128, 2, (9, 0), (), "Border", '{00020854-0000-0000-C000-000000000046}'),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'Fill' returns object of type 'ChartFillFormat'
		"Fill": (1663, 2, (9, 0), (), "Fill", '{00024435-0000-0000-C000-000000000046}'),
		# Method 'Font' returns object of type 'Font'
		"Font": (146, 2, (9, 0), (), "Font", '{0002084D-0000-0000-C000-000000000046}'),
		# Method 'Format' returns object of type 'ChartFormat'
		"Format": (116, 2, (9, 0), (), "Format", '{000244B2-0000-0000-C000-000000000046}'),
		"Height": (123, 2, (5, 0), (), "Height", None),
		# Method 'Interior' returns object of type 'Interior'
		"Interior": (129, 2, (9, 0), (), "Interior", '{00020870-0000-0000-C000-000000000046}'),
		"Left": (127, 2, (5, 0), (), "Left", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"RoundedCorners": (619, 2, (11, 0), (), "RoundedCorners", None),
		"Shadow": (103, 2, (11, 0), (), "Shadow", None),
		"Top": (126, 2, (5, 0), (), "Top", None),
		"Width": (122, 2, (5, 0), (), "Width", None),
	}
	_prop_map_put_ = {
		"AutoScaleFont": ((1525, LCID, 4, 0),()),
		"Height": ((123, LCID, 4, 0),()),
		"Left": ((127, LCID, 4, 0),()),
		"RoundedCorners": ((619, LCID, 4, 0),()),
		"Shadow": ((103, LCID, 4, 0),()),
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

win32com.client.CLSIDToClass.RegisterCLSID( "{000208CC-0000-0000-C000-000000000046}", ChartArea )
