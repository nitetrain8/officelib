# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 16:46:57 2013
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
class ChartObject(DispatchBaseClass):
	CLSID = IID('{000208CF-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Activate(self):
		return self._ApplyTypes_(304, 1, (12, 0), (), 'Activate', None,)

	def BringToFront(self):
		return self._ApplyTypes_(602, 1, (12, 0), (), 'BringToFront', None,)

	def Copy(self):
		return self._ApplyTypes_(551, 1, (12, 0), (), 'Copy', None,)

	def CopyPicture(self, Appearance=2, Format=-4147):
		return self._ApplyTypes_(213, 1, (12, 0), ((3, 49), (3, 49)), 'CopyPicture', None,Appearance
			, Format)

	def Cut(self):
		return self._ApplyTypes_(565, 1, (12, 0), (), 'Cut', None,)

	def Delete(self):
		return self._ApplyTypes_(117, 1, (12, 0), (), 'Delete', None,)

	def Duplicate(self):
		ret = self._oleobj_.InvokeTypes(1039, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'Duplicate', None)
		return ret

	def Select(self, Replace=defaultNamedOptArg):
		return self._ApplyTypes_(235, 1, (12, 0), ((12, 17),), 'Select', None,Replace
			)

	def SendToBack(self):
		return self._ApplyTypes_(605, 1, (12, 0), (), 'SendToBack', None,)

	def _Copy(self):
		return self._ApplyTypes_(2609, 1, (12, 0), (), '_Copy', None,)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		# Method 'Border' returns object of type 'Border'
		"Border": (128, 2, (9, 0), (), "Border", '{00020854-0000-0000-C000-000000000046}'),
		# Method 'BottomRightCell' returns object of type 'Range'
		"BottomRightCell": (615, 2, (9, 0), (), "BottomRightCell", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Chart' returns object of type 'Chart'
		"Chart": (7, 2, (13, 0), (), "Chart", '{00020821-0000-0000-C000-000000000046}'),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		"Enabled": (600, 2, (11, 0), (), "Enabled", None),
		"Height": (123, 2, (5, 0), (), "Height", None),
		"Index": (486, 2, (3, 0), (), "Index", None),
		# Method 'Interior' returns object of type 'Interior'
		"Interior": (129, 2, (9, 0), (), "Interior", '{00020870-0000-0000-C000-000000000046}'),
		"Left": (127, 2, (5, 0), (), "Left", None),
		"Locked": (269, 2, (11, 0), (), "Locked", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		"OnAction": (596, 2, (8, 0), (), "OnAction", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"Placement": (617, 2, (12, 0), (), "Placement", None),
		"PrintObject": (618, 2, (11, 0), (), "PrintObject", None),
		"ProtectChartObject": (1529, 2, (11, 0), (), "ProtectChartObject", None),
		"RoundedCorners": (619, 2, (11, 0), (), "RoundedCorners", None),
		"Shadow": (103, 2, (11, 0), (), "Shadow", None),
		# Method 'ShapeRange' returns object of type 'ShapeRange'
		"ShapeRange": (1528, 2, (9, 0), (), "ShapeRange", '{0002443B-0000-0000-C000-000000000046}'),
		"Top": (126, 2, (5, 0), (), "Top", None),
		# Method 'TopLeftCell' returns object of type 'Range'
		"TopLeftCell": (620, 2, (9, 0), (), "TopLeftCell", '{00020846-0000-0000-C000-000000000046}'),
		"Visible": (558, 2, (11, 0), (), "Visible", None),
		"Width": (122, 2, (5, 0), (), "Width", None),
		"ZOrder": (622, 2, (3, 0), (), "ZOrder", None),
	}
	_prop_map_put_ = {
		"Enabled": ((600, LCID, 4, 0),()),
		"Height": ((123, LCID, 4, 0),()),
		"Left": ((127, LCID, 4, 0),()),
		"Locked": ((269, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"OnAction": ((596, LCID, 4, 0),()),
		"Placement": ((617, LCID, 4, 0),()),
		"PrintObject": ((618, LCID, 4, 0),()),
		"ProtectChartObject": ((1529, LCID, 4, 0),()),
		"RoundedCorners": ((619, LCID, 4, 0),()),
		"Shadow": ((103, LCID, 4, 0),()),
		"Top": ((126, LCID, 4, 0),()),
		"Visible": ((558, LCID, 4, 0),()),
		"Width": ((122, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000208CF-0000-0000-C000-000000000046}", ChartObject )
