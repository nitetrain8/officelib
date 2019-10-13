# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 17:07:14 2013
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
class Series(DispatchBaseClass):
	CLSID = IID('{0002086B-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ApplyCustomType(self, ChartType=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1401, LCID, 1, (24, 0), ((3, 1),),ChartType
			)

	def ApplyDataLabels(self, Type=2, LegendKey=defaultNamedOptArg, AutoText=defaultNamedOptArg, HasLeaderLines=defaultNamedOptArg
			, ShowSeriesName=defaultNamedOptArg, ShowCategoryName=defaultNamedOptArg, ShowValue=defaultNamedOptArg, ShowPercentage=defaultNamedOptArg, ShowBubbleSize=defaultNamedOptArg
			, Separator=defaultNamedOptArg):
		return self._ApplyTypes_(1922, 1, (12, 0), ((3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)), 'ApplyDataLabels', None,Type
			, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName
			, ShowValue, ShowPercentage, ShowBubbleSize, Separator)

	def ClearFormats(self):
		return self._ApplyTypes_(112, 1, (12, 0), (), 'ClearFormats', None,)

	def Copy(self):
		return self._ApplyTypes_(551, 1, (12, 0), (), 'Copy', None,)

	def DataLabels(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(157, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'DataLabels', None)
		return ret

	def Delete(self):
		return self._ApplyTypes_(117, 1, (12, 0), (), 'Delete', None,)

	def ErrorBar(self, Direction=defaultNamedNotOptArg, Include=defaultNamedNotOptArg, Type=defaultNamedNotOptArg, Amount=defaultNamedOptArg
			, MinusValues=defaultNamedOptArg):
		return self._ApplyTypes_(152, 1, (12, 0), ((3, 1), (3, 1), (3, 1), (12, 17), (12, 17)), 'ErrorBar', None,Direction
			, Include, Type, Amount, MinusValues)

	def Paste(self):
		return self._ApplyTypes_(211, 1, (12, 0), (), 'Paste', None,)

	def Points(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(70, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Points', None)
		return ret

	def Select(self):
		return self._ApplyTypes_(235, 1, (12, 0), (), 'Select', None,)

	def Trendlines(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(154, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Trendlines', None)
		return ret

	def _ApplyDataLabels(self, Type=2, LegendKey=defaultNamedOptArg, AutoText=defaultNamedOptArg, HasLeaderLines=defaultNamedOptArg):
		return self._ApplyTypes_(151, 1, (12, 0), ((3, 49), (12, 17), (12, 17), (12, 17)), '_ApplyDataLabels', None,Type
			, LegendKey, AutoText, HasLeaderLines)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"ApplyPictToEnd": (1661, 2, (11, 0), (), "ApplyPictToEnd", None),
		"ApplyPictToFront": (1660, 2, (11, 0), (), "ApplyPictToFront", None),
		"ApplyPictToSides": (1659, 2, (11, 0), (), "ApplyPictToSides", None),
		"AxisGroup": (47, 2, (3, 0), (), "AxisGroup", None),
		"BarShape": (1403, 2, (3, 0), (), "BarShape", None),
		# Method 'Border' returns object of type 'Border'
		"Border": (128, 2, (9, 0), (), "Border", '{00020854-0000-0000-C000-000000000046}'),
		"BubbleSizes": (1664, 2, (12, 0), (), "BubbleSizes", None),
		"ChartType": (1400, 2, (3, 0), (), "ChartType", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'ErrorBars' returns object of type 'ErrorBars'
		"ErrorBars": (159, 2, (9, 0), (), "ErrorBars", '{000208CE-0000-0000-C000-000000000046}'),
		"Explosion": (182, 2, (3, 0), (), "Explosion", None),
		# Method 'Fill' returns object of type 'ChartFillFormat'
		"Fill": (1663, 2, (9, 0), (), "Fill", '{00024435-0000-0000-C000-000000000046}'),
		# Method 'Format' returns object of type 'ChartFormat'
		"Format": (116, 2, (9, 0), (), "Format", '{000244B2-0000-0000-C000-000000000046}'),
		"Formula": (261, 2, (8, 0), (), "Formula", None),
		"FormulaLocal": (263, 2, (8, 0), (), "FormulaLocal", None),
		"FormulaR1C1": (264, 2, (8, 0), (), "FormulaR1C1", None),
		"FormulaR1C1Local": (265, 2, (8, 0), (), "FormulaR1C1Local", None),
		"Has3DEffect": (1665, 2, (11, 0), (), "Has3DEffect", None),
		"HasDataLabels": (78, 2, (11, 0), (), "HasDataLabels", None),
		"HasErrorBars": (160, 2, (11, 0), (), "HasErrorBars", None),
		"HasLeaderLines": (1394, 2, (11, 0), (), "HasLeaderLines", None),
		# Method 'Interior' returns object of type 'Interior'
		"Interior": (129, 2, (9, 0), (), "Interior", '{00020870-0000-0000-C000-000000000046}'),
		"InvertIfNegative": (132, 2, (11, 0), (), "InvertIfNegative", None),
		# Method 'LeaderLines' returns object of type 'LeaderLines'
		"LeaderLines": (1666, 2, (9, 0), (), "LeaderLines", '{00024437-0000-0000-C000-000000000046}'),
		"MarkerBackgroundColor": (73, 2, (3, 0), (), "MarkerBackgroundColor", None),
		"MarkerBackgroundColorIndex": (74, 2, (3, 0), (), "MarkerBackgroundColorIndex", None),
		"MarkerForegroundColor": (75, 2, (3, 0), (), "MarkerForegroundColor", None),
		"MarkerForegroundColorIndex": (76, 2, (3, 0), (), "MarkerForegroundColorIndex", None),
		"MarkerSize": (231, 2, (3, 0), (), "MarkerSize", None),
		"MarkerStyle": (72, 2, (3, 0), (), "MarkerStyle", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"PictureType": (161, 2, (3, 0), (), "PictureType", None),
		"PictureUnit": (162, 2, (3, 0), (), "PictureUnit", None),
		"PictureUnit2": (2649, 2, (5, 0), (), "PictureUnit2", None),
		"PlotOrder": (228, 2, (3, 0), (), "PlotOrder", None),
		"Shadow": (103, 2, (11, 0), (), "Shadow", None),
		"Smooth": (163, 2, (11, 0), (), "Smooth", None),
		"Type": (108, 2, (3, 0), (), "Type", None),
		"Values": (164, 2, (12, 0), (), "Values", None),
		"XValues": (1111, 2, (12, 0), (), "XValues", None),
	}
	_prop_map_put_ = {
		"ApplyPictToEnd": ((1661, LCID, 4, 0),()),
		"ApplyPictToFront": ((1660, LCID, 4, 0),()),
		"ApplyPictToSides": ((1659, LCID, 4, 0),()),
		"AxisGroup": ((47, LCID, 4, 0),()),
		"BarShape": ((1403, LCID, 4, 0),()),
		"BubbleSizes": ((1664, LCID, 4, 0),()),
		"ChartType": ((1400, LCID, 4, 0),()),
		"Explosion": ((182, LCID, 4, 0),()),
		"Formula": ((261, LCID, 4, 0),()),
		"FormulaLocal": ((263, LCID, 4, 0),()),
		"FormulaR1C1": ((264, LCID, 4, 0),()),
		"FormulaR1C1Local": ((265, LCID, 4, 0),()),
		"Has3DEffect": ((1665, LCID, 4, 0),()),
		"HasDataLabels": ((78, LCID, 4, 0),()),
		"HasErrorBars": ((160, LCID, 4, 0),()),
		"HasLeaderLines": ((1394, LCID, 4, 0),()),
		"InvertIfNegative": ((132, LCID, 4, 0),()),
		"MarkerBackgroundColor": ((73, LCID, 4, 0),()),
		"MarkerBackgroundColorIndex": ((74, LCID, 4, 0),()),
		"MarkerForegroundColor": ((75, LCID, 4, 0),()),
		"MarkerForegroundColorIndex": ((76, LCID, 4, 0),()),
		"MarkerSize": ((231, LCID, 4, 0),()),
		"MarkerStyle": ((72, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"PictureType": ((161, LCID, 4, 0),()),
		"PictureUnit": ((162, LCID, 4, 0),()),
		"PictureUnit2": ((2649, LCID, 4, 0),()),
		"PlotOrder": ((228, LCID, 4, 0),()),
		"Shadow": ((103, LCID, 4, 0),()),
		"Smooth": ((163, LCID, 4, 0),()),
		"Type": ((108, LCID, 4, 0),()),
		"Values": ((164, LCID, 4, 0),()),
		"XValues": ((1111, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{0002086B-0000-0000-C000-000000000046}", Series )
