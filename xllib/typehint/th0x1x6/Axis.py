# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 17:24:22 2013
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
class Axis(DispatchBaseClass):
	CLSID = IID('{00020848-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Delete(self):
		return self._ApplyTypes_(117, 1, (12, 0), (), 'Delete', None,)

	def Select(self):
		return self._ApplyTypes_(235, 1, (12, 0), (), 'Select', None,)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"AxisBetweenCategories": (45, 2, (11, 0), (), "AxisBetweenCategories", None),
		"AxisGroup": (47, 2, (3, 0), (), "AxisGroup", None),
		# Method 'AxisTitle' returns object of type 'AxisTitle'
		"AxisTitle": (82, 2, (9, 0), (), "AxisTitle", '{0002084A-0000-0000-C000-000000000046}'),
		"BaseUnit": (1647, 2, (3, 0), (), "BaseUnit", None),
		"BaseUnitIsAuto": (1648, 2, (11, 0), (), "BaseUnitIsAuto", None),
		# Method 'Border' returns object of type 'Border'
		"Border": (128, 2, (9, 0), (), "Border", '{00020854-0000-0000-C000-000000000046}'),
		"CategoryNames": (156, 2, (12, 0), (), "CategoryNames", None),
		"CategoryType": (1651, 2, (3, 0), (), "CategoryType", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		"Crosses": (42, 2, (3, 0), (), "Crosses", None),
		"CrossesAt": (43, 2, (5, 0), (), "CrossesAt", None),
		"DisplayUnit": (1886, 2, (3, 0), (), "DisplayUnit", None),
		"DisplayUnitCustom": (1887, 2, (5, 0), (), "DisplayUnitCustom", None),
		# Method 'DisplayUnitLabel' returns object of type 'DisplayUnitLabel'
		"DisplayUnitLabel": (1889, 2, (9, 0), (), "DisplayUnitLabel", '{0002084C-0000-0000-C000-000000000046}'),
		# Method 'Format' returns object of type 'ChartFormat'
		"Format": (116, 2, (9, 0), (), "Format", '{000244B2-0000-0000-C000-000000000046}'),
		"HasDisplayUnitLabel": (1888, 2, (11, 0), (), "HasDisplayUnitLabel", None),
		"HasMajorGridlines": (24, 2, (11, 0), (), "HasMajorGridlines", None),
		"HasMinorGridlines": (25, 2, (11, 0), (), "HasMinorGridlines", None),
		"HasTitle": (54, 2, (11, 0), (), "HasTitle", None),
		"Height": (123, 2, (5, 0), (), "Height", None),
		"Left": (127, 2, (5, 0), (), "Left", None),
		"LogBase": (2646, 2, (5, 0), (), "LogBase", None),
		# Method 'MajorGridlines' returns object of type 'Gridlines'
		"MajorGridlines": (89, 2, (9, 0), (), "MajorGridlines", '{000208C3-0000-0000-C000-000000000046}'),
		"MajorTickMark": (26, 2, (3, 0), (), "MajorTickMark", None),
		"MajorUnit": (37, 2, (5, 0), (), "MajorUnit", None),
		"MajorUnitIsAuto": (38, 2, (11, 0), (), "MajorUnitIsAuto", None),
		"MajorUnitScale": (1649, 2, (3, 0), (), "MajorUnitScale", None),
		"MaximumScale": (35, 2, (5, 0), (), "MaximumScale", None),
		"MaximumScaleIsAuto": (36, 2, (11, 0), (), "MaximumScaleIsAuto", None),
		"MinimumScale": (33, 2, (5, 0), (), "MinimumScale", None),
		"MinimumScaleIsAuto": (34, 2, (11, 0), (), "MinimumScaleIsAuto", None),
		# Method 'MinorGridlines' returns object of type 'Gridlines'
		"MinorGridlines": (90, 2, (9, 0), (), "MinorGridlines", '{000208C3-0000-0000-C000-000000000046}'),
		"MinorTickMark": (27, 2, (3, 0), (), "MinorTickMark", None),
		"MinorUnit": (39, 2, (5, 0), (), "MinorUnit", None),
		"MinorUnitIsAuto": (40, 2, (11, 0), (), "MinorUnitIsAuto", None),
		"MinorUnitScale": (1650, 2, (3, 0), (), "MinorUnitScale", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"ReversePlotOrder": (44, 2, (11, 0), (), "ReversePlotOrder", None),
		"ScaleType": (41, 2, (3, 0), (), "ScaleType", None),
		"TickLabelPosition": (28, 2, (3, 0), (), "TickLabelPosition", None),
		"TickLabelSpacing": (29, 2, (3, 0), (), "TickLabelSpacing", None),
		"TickLabelSpacingIsAuto": (2647, 2, (11, 0), (), "TickLabelSpacingIsAuto", None),
		# Method 'TickLabels' returns object of type 'TickLabels'
		"TickLabels": (91, 2, (9, 0), (), "TickLabels", '{000208C9-0000-0000-C000-000000000046}'),
		"TickMarkSpacing": (31, 2, (3, 0), (), "TickMarkSpacing", None),
		"Top": (126, 2, (5, 0), (), "Top", None),
		"Type": (108, 2, (3, 0), (), "Type", None),
		"Width": (122, 2, (5, 0), (), "Width", None),
	}
	_prop_map_put_ = {
		"AxisBetweenCategories": ((45, LCID, 4, 0),()),
		"BaseUnit": ((1647, LCID, 4, 0),()),
		"BaseUnitIsAuto": ((1648, LCID, 4, 0),()),
		"CategoryNames": ((156, LCID, 4, 0),()),
		"CategoryType": ((1651, LCID, 4, 0),()),
		"Crosses": ((42, LCID, 4, 0),()),
		"CrossesAt": ((43, LCID, 4, 0),()),
		"DisplayUnit": ((1886, LCID, 4, 0),()),
		"DisplayUnitCustom": ((1887, LCID, 4, 0),()),
		"HasDisplayUnitLabel": ((1888, LCID, 4, 0),()),
		"HasMajorGridlines": ((24, LCID, 4, 0),()),
		"HasMinorGridlines": ((25, LCID, 4, 0),()),
		"HasTitle": ((54, LCID, 4, 0),()),
		"LogBase": ((2646, LCID, 4, 0),()),
		"MajorTickMark": ((26, LCID, 4, 0),()),
		"MajorUnit": ((37, LCID, 4, 0),()),
		"MajorUnitIsAuto": ((38, LCID, 4, 0),()),
		"MajorUnitScale": ((1649, LCID, 4, 0),()),
		"MaximumScale": ((35, LCID, 4, 0),()),
		"MaximumScaleIsAuto": ((36, LCID, 4, 0),()),
		"MinimumScale": ((33, LCID, 4, 0),()),
		"MinimumScaleIsAuto": ((34, LCID, 4, 0),()),
		"MinorTickMark": ((27, LCID, 4, 0),()),
		"MinorUnit": ((39, LCID, 4, 0),()),
		"MinorUnitIsAuto": ((40, LCID, 4, 0),()),
		"MinorUnitScale": ((1650, LCID, 4, 0),()),
		"ReversePlotOrder": ((44, LCID, 4, 0),()),
		"ScaleType": ((41, LCID, 4, 0),()),
		"TickLabelPosition": ((28, LCID, 4, 0),()),
		"TickLabelSpacing": ((29, LCID, 4, 0),()),
		"TickLabelSpacingIsAuto": ((2647, LCID, 4, 0),()),
		"TickMarkSpacing": ((31, LCID, 4, 0),()),
		"Type": ((108, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00020848-0000-0000-C000-000000000046}", Axis )
