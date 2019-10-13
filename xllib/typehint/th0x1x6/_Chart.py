# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 17:06:58 2013
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
class _Chart(DispatchBaseClass):
	CLSID = IID('{000208D6-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020821-0000-0000-C000-000000000046}')

	def Activate(self):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0), (),)

	def ApplyChartTemplate(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2507, LCID, 1, (24, 0), ((8, 1),),Filename
			)

	def ApplyCustomType(self, ChartType=defaultNamedNotOptArg, TypeName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1401, LCID, 1, (24, 0), ((3, 1), (12, 17)),ChartType
			, TypeName)

	def ApplyDataLabels(self, Type=2, LegendKey=defaultNamedOptArg, AutoText=defaultNamedOptArg, HasLeaderLines=defaultNamedOptArg
			, ShowSeriesName=defaultNamedOptArg, ShowCategoryName=defaultNamedOptArg, ShowValue=defaultNamedOptArg, ShowPercentage=defaultNamedOptArg, ShowBubbleSize=defaultNamedOptArg
			, Separator=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1922, LCID, 1, (24, 0), ((3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Type
			, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName
			, ShowValue, ShowPercentage, ShowBubbleSize, Separator)

	def ApplyLayout(self, Layout=defaultNamedNotOptArg, ChartType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2500, LCID, 1, (24, 0), ((3, 1), (12, 17)),Layout
			, ChartType)

	def Arcs(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(760, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Arcs', None)
		return ret

	def AreaGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(9, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'AreaGroups', None)
		return ret

	def AutoFormat(self, Gallery=defaultNamedNotOptArg, Format=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(114, LCID, 1, (24, 0), ((3, 1), (12, 17)),Gallery
			, Format)

	def Axes(self, Type=defaultNamedNotOptArg, AxisGroup=1):
		ret = self._oleobj_.InvokeTypes(23, LCID, 1, (9, 0), ((12, 17), (3, 49)),Type
			, AxisGroup)
		if ret is not None:
			ret = Dispatch(ret, 'Axes', None)
		return ret

	def BarGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(10, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'BarGroups', None)
		return ret

	def Buttons(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(557, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Buttons', None)
		return ret

	def ChartGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(8, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ChartGroups', None)
		return ret

	def ChartObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1060, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ChartObjects', None)
		return ret

	def ChartWizard(self, Source=defaultNamedOptArg, Gallery=defaultNamedOptArg, Format=defaultNamedOptArg, PlotBy=defaultNamedOptArg
			, CategoryLabels=defaultNamedOptArg, SeriesLabels=defaultNamedOptArg, HasLegend=defaultNamedOptArg, Title=defaultNamedOptArg, CategoryTitle=defaultNamedOptArg
			, ValueTitle=defaultNamedOptArg, ExtraTitle=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(196, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Source
			, Gallery, Format, PlotBy, CategoryLabels, SeriesLabels
			, HasLegend, Title, CategoryTitle, ValueTitle, ExtraTitle
			)

	def CheckBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(824, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'CheckBoxes', None)
		return ret

	def CheckSpelling(self, CustomDictionary=defaultNamedOptArg, IgnoreUppercase=defaultNamedOptArg, AlwaysSuggest=defaultNamedOptArg, SpellLang=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(505, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17)),CustomDictionary
			, IgnoreUppercase, AlwaysSuggest, SpellLang)

	def ClearToMatchStyle(self):
		return self._oleobj_.InvokeTypes(2510, LCID, 1, (24, 0), (),)

	def ColumnGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(11, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ColumnGroups', None)
		return ret

	def Copy(self, Before=defaultNamedOptArg, After=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(551, LCID, 1, (24, 0), ((12, 17), (12, 17)),Before
			, After)

	def CopyChartBuild(self):
		return self._oleobj_.InvokeTypes(1404, LCID, 1, (24, 0), (),)

	def CopyPicture(self, Appearance=1, Format=-4147, Size=2):
		return self._oleobj_.InvokeTypes(213, LCID, 1, (24, 0), ((3, 49), (3, 49), (3, 49)),Appearance
			, Format, Size)

	def CreatePublisher(self, Edition=defaultNamedNotOptArg, Appearance=1, Size=1, ContainsPICT=defaultNamedOptArg
			, ContainsBIFF=defaultNamedOptArg, ContainsRTF=defaultNamedOptArg, ContainsVALU=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(458, LCID, 1, (24, 0), ((12, 17), (3, 49), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17)),Edition
			, Appearance, Size, ContainsPICT, ContainsBIFF, ContainsRTF
			, ContainsVALU)

	def Delete(self):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (24, 0), (),)

	def Deselect(self):
		return self._oleobj_.InvokeTypes(1120, LCID, 1, (24, 0), (),)

	def DoughnutGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(14, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'DoughnutGroups', None)
		return ret

	def DrawingObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(88, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'DrawingObjects', None)
		return ret

	def Drawings(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(772, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Drawings', None)
		return ret

	def DropDowns(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(836, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'DropDowns', None)
		return ret

	def Evaluate(self, Name=defaultNamedNotOptArg):
		return self._ApplyTypes_(1, 1, (12, 0), ((12, 1),), 'Evaluate', None,Name
			)

	def Export(self, Filename=defaultNamedNotOptArg, FilterName=defaultNamedOptArg, Interactive=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1414, LCID, 1, (11, 0), ((8, 1), (12, 17), (12, 17)),Filename
			, FilterName, Interactive)

	def ExportAsFixedFormat(self, Type=defaultNamedNotOptArg, Filename=defaultNamedOptArg, Quality=defaultNamedOptArg, IncludeDocProperties=defaultNamedOptArg
			, IgnorePrintAreas=defaultNamedOptArg, From=defaultNamedOptArg, To=defaultNamedOptArg, OpenAfterPublish=defaultNamedOptArg, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2493, LCID, 1, (24, 0), ((3, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Type
			, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From
			, To, OpenAfterPublish, FixedFormatExtClassPtr)

	def GetChartElement(self, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg, ElementID=defaultNamedNotOptArg, Arg1=defaultNamedNotOptArg
			, Arg2=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1409, LCID, 1, (24, 0), ((3, 1), (3, 1), (16387, 1), (16387, 1), (16387, 1)),x
			, y, ElementID, Arg1, Arg2)

	# The method GetHasAxis is actually a property, but must be used as a method to correctly pass the arguments
	def GetHasAxis(self, Index1=defaultNamedOptArg, Index2=defaultNamedOptArg):
		return self._ApplyTypes_(52, 2, (12, 0), ((12, 17), (12, 17)), 'GetHasAxis', None,Index1
			, Index2)

	def GroupBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(834, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'GroupBoxes', None)
		return ret

	def GroupObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1113, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'GroupObjects', None)
		return ret

	def Labels(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(841, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Labels', None)
		return ret

	def LineGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(12, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'LineGroups', None)
		return ret

	def Lines(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(767, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Lines', None)
		return ret

	def ListBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(832, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ListBoxes', None)
		return ret

	# Result is of type Chart
	def Location(self, Where=defaultNamedNotOptArg, Name=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1397, LCID, 1, (13, 0), ((3, 1), (12, 17)),Where
			, Name)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Location', '{00020821-0000-0000-C000-000000000046}')
		return ret

	def Move(self, Before=defaultNamedOptArg, After=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(637, LCID, 1, (24, 0), ((12, 17), (12, 17)),Before
			, After)

	def OLEObjects(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(799, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'OLEObjects', None)
		return ret

	def OptionButtons(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(826, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'OptionButtons', None)
		return ret

	def Ovals(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(801, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Ovals', None)
		return ret

	def Paste(self, Type=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(211, LCID, 1, (24, 0), ((12, 17),),Type
			)

	def Pictures(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(771, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Pictures', None)
		return ret

	def PieGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(13, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'PieGroups', None)
		return ret

	def PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2361, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def PrintPreview(self, EnableChanges=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(281, LCID, 1, (24, 0), ((12, 17),),EnableChanges
			)

	def Protect(self, Password=defaultNamedOptArg, DrawingObjects=defaultNamedOptArg, Contents=defaultNamedOptArg, Scenarios=defaultNamedOptArg
			, UserInterfaceOnly=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2029, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Password
			, DrawingObjects, Contents, Scenarios, UserInterfaceOnly)

	def RadarGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(15, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'RadarGroups', None)
		return ret

	def Rectangles(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(774, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Rectangles', None)
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(1417, LCID, 1, (24, 0), (),)

	def SaveAs(self, Filename=defaultNamedNotOptArg, FileFormat=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg
			, ReadOnlyRecommended=defaultNamedOptArg, CreateBackup=defaultNamedOptArg, AddToMru=defaultNamedOptArg, TextCodepage=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg
			, Local=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1925, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup
			, AddToMru, TextCodepage, TextVisualLayout, Local)

	def SaveChartTemplate(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2508, LCID, 1, (24, 0), ((8, 1),),Filename
			)

	def ScrollBars(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(830, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'ScrollBars', None)
		return ret

	def Select(self, Replace=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(235, LCID, 1, (24, 0), ((12, 17),),Replace
			)

	def SeriesCollection(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(68, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'SeriesCollection', None)
		return ret

	def SetBackgroundPicture(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1188, LCID, 1, (24, 0), ((8, 1),),Filename
			)

	def SetDefaultChart(self, Name=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(219, LCID, 1, (24, 0), ((12, 1),),Name
			)

	def SetElement(self, Element=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2502, LCID, 1, (24, 0), ((3, 1),),Element
			)

	# The method SetHasAxis is actually a property, but must be used as a method to correctly pass the arguments
	def SetHasAxis(self, Index1=defaultNamedNotOptArg, Index2=defaultNamedOptArg, arg2=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(52, LCID, 4, (24, 0), ((12, 17), (12, 17), (12, 1)),Index1
			, Index2, arg2)

	def SetSourceData(self, Source=defaultNamedNotOptArg, PlotBy=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1413, LCID, 1, (24, 0), ((9, 1), (12, 17)),Source
			, PlotBy)

	def Spinners(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(838, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Spinners', None)
		return ret

	def TextBoxes(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(777, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'TextBoxes', None)
		return ret

	def Unprotect(self, Password=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(285, LCID, 1, (24, 0), ((12, 17),),Password
			)

	def XYGroups(self, Index=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(16, LCID, 1, (9, 0), ((12, 17),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'XYGroups', None)
		return ret

	def _ApplyDataLabels(self, Type=2, LegendKey=defaultNamedOptArg, AutoText=defaultNamedOptArg, HasLeaderLines=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(151, LCID, 1, (24, 0), ((3, 49), (12, 17), (12, 17), (12, 17)),Type
			, LegendKey, AutoText, HasLeaderLines)

	def _Evaluate(self, Name=defaultNamedNotOptArg):
		return self._ApplyTypes_(-5, 1, (12, 0), ((12, 1),), '_Evaluate', None,Name
			)

	def _PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1772, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def _Protect(self, Password=defaultNamedOptArg, DrawingObjects=defaultNamedOptArg, Contents=defaultNamedOptArg, Scenarios=defaultNamedOptArg
			, UserInterfaceOnly=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(282, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Password
			, DrawingObjects, Contents, Scenarios, UserInterfaceOnly)

	def _SaveAs(self, Filename=defaultNamedNotOptArg, FileFormat=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg
			, ReadOnlyRecommended=defaultNamedOptArg, CreateBackup=defaultNamedOptArg, AddToMru=defaultNamedOptArg, TextCodepage=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(284, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup
			, AddToMru, TextCodepage, TextVisualLayout)

	def _PrintOut_(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(905, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		# Method 'Area3DGroup' returns object of type 'ChartGroup'
		"Area3DGroup": (17, 2, (9, 0), (), "Area3DGroup", '{00020859-0000-0000-C000-000000000046}'),
		"AutoScaling": (107, 2, (11, 0), (), "AutoScaling", None),
		# Method 'BackWall' returns object of type 'Walls'
		"BackWall": (2506, 2, (9, 0), (), "BackWall", '{000208C8-0000-0000-C000-000000000046}'),
		# Method 'Bar3DGroup' returns object of type 'ChartGroup'
		"Bar3DGroup": (18, 2, (9, 0), (), "Bar3DGroup", '{00020859-0000-0000-C000-000000000046}'),
		"BarShape": (1403, 2, (3, 0), (), "BarShape", None),
		# Method 'ChartArea' returns object of type 'ChartArea'
		"ChartArea": (80, 2, (9, 0), (), "ChartArea", '{000208CC-0000-0000-C000-000000000046}'),
		"ChartStyle": (2509, 2, (12, 0), (), "ChartStyle", None),
		# Method 'ChartTitle' returns object of type 'ChartTitle'
		"ChartTitle": (81, 2, (9, 0), (), "ChartTitle", '{00020849-0000-0000-C000-000000000046}'),
		"ChartType": (1400, 2, (3, 0), (), "ChartType", None),
		"CodeName": (1373, 2, (8, 0), (), "CodeName", None),
		# Method 'Column3DGroup' returns object of type 'ChartGroup'
		"Column3DGroup": (19, 2, (9, 0), (), "Column3DGroup", '{00020859-0000-0000-C000-000000000046}'),
		# Method 'Corners' returns object of type 'Corners'
		"Corners": (79, 2, (9, 0), (), "Corners", '{000208C0-0000-0000-C000-000000000046}'),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'DataTable' returns object of type 'DataTable'
		"DataTable": (1395, 2, (9, 0), (), "DataTable", '{00020843-0000-0000-C000-000000000046}'),
		"DepthPercent": (48, 2, (3, 0), (), "DepthPercent", None),
		"DisplayBlanksAs": (93, 2, (3, 0), (), "DisplayBlanksAs", None),
		"Elevation": (49, 2, (3, 0), (), "Elevation", None),
		# Method 'Floor' returns object of type 'Floor'
		"Floor": (83, 2, (9, 0), (), "Floor", '{000208C7-0000-0000-C000-000000000046}'),
		"GapDepth": (50, 2, (3, 0), (), "GapDepth", None),
		"HasAxis": (52, 2, (12, 0), ((12, 17), (12, 17)), "HasAxis", None),
		"HasDataTable": (1396, 2, (11, 0), (), "HasDataTable", None),
		"HasLegend": (53, 2, (11, 0), (), "HasLegend", None),
		"HasPivotFields": (1815, 2, (11, 0), (), "HasPivotFields", None),
		"HasTitle": (54, 2, (11, 0), (), "HasTitle", None),
		"HeightPercent": (55, 2, (3, 0), (), "HeightPercent", None),
		# Method 'Hyperlinks' returns object of type 'Hyperlinks'
		"Hyperlinks": (1393, 2, (9, 0), (), "Hyperlinks", '{00024430-0000-0000-C000-000000000046}'),
		"Index": (486, 2, (3, 0), (), "Index", None),
		# Method 'Legend' returns object of type 'Legend'
		"Legend": (84, 2, (9, 0), (), "Legend", '{000208CD-0000-0000-C000-000000000046}'),
		# Method 'Line3DGroup' returns object of type 'ChartGroup'
		"Line3DGroup": (20, 2, (9, 0), (), "Line3DGroup", '{00020859-0000-0000-C000-000000000046}'),
		# Method 'MailEnvelope' returns object of type 'MsoEnvelope'
		"MailEnvelope": (2021, 2, (13, 0), (), "MailEnvelope", '{0006F01A-0000-0000-C000-000000000046}'),
		"Name": (110, 2, (8, 0), (), "Name", None),
		"Next": (502, 2, (9, 0), (), "Next", None),
		"OnDoubleClick": (628, 2, (8, 0), (), "OnDoubleClick", None),
		"OnSheetActivate": (1031, 2, (8, 0), (), "OnSheetActivate", None),
		"OnSheetDeactivate": (1081, 2, (8, 0), (), "OnSheetDeactivate", None),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (998, 2, (9, 0), (), "PageSetup", '{000208B4-0000-0000-C000-000000000046}'),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"Perspective": (57, 2, (3, 0), (), "Perspective", None),
		# Method 'Pie3DGroup' returns object of type 'ChartGroup'
		"Pie3DGroup": (21, 2, (9, 0), (), "Pie3DGroup", '{00020859-0000-0000-C000-000000000046}'),
		# Method 'PivotLayout' returns object of type 'PivotLayout'
		"PivotLayout": (1814, 2, (9, 0), (), "PivotLayout", '{0002444A-0000-0000-C000-000000000046}'),
		# Method 'PlotArea' returns object of type 'PlotArea'
		"PlotArea": (85, 2, (9, 0), (), "PlotArea", '{000208CB-0000-0000-C000-000000000046}'),
		"PlotBy": (202, 2, (3, 0), (), "PlotBy", None),
		"PlotVisibleOnly": (92, 2, (11, 0), (), "PlotVisibleOnly", None),
		"Previous": (503, 2, (9, 0), (), "Previous", None),
		"ProtectContents": (292, 2, (11, 0), (), "ProtectContents", None),
		"ProtectData": (1406, 2, (11, 0), (), "ProtectData", None),
		"ProtectDrawingObjects": (293, 2, (11, 0), (), "ProtectDrawingObjects", None),
		"ProtectFormatting": (1405, 2, (11, 0), (), "ProtectFormatting", None),
		"ProtectGoalSeek": (1407, 2, (11, 0), (), "ProtectGoalSeek", None),
		"ProtectSelection": (1408, 2, (11, 0), (), "ProtectSelection", None),
		"ProtectionMode": (1159, 2, (11, 0), (), "ProtectionMode", None),
		"RightAngleAxes": (58, 2, (12, 0), (), "RightAngleAxes", None),
		"Rotation": (59, 2, (12, 0), (), "Rotation", None),
		# Method 'Scripts' returns object of type 'Scripts'
		"Scripts": (1816, 2, (9, 0), (), "Scripts", '{000C0340-0000-0000-C000-000000000046}'),
		# Method 'Shapes' returns object of type 'Shapes'
		"Shapes": (1377, 2, (9, 0), (), "Shapes", '{0002443A-0000-0000-C000-000000000046}'),
		"ShowDataLabelsOverMaximum": (2504, 2, (11, 0), (), "ShowDataLabelsOverMaximum", None),
		"ShowWindow": (1399, 2, (11, 0), (), "ShowWindow", None),
		# Method 'SideWall' returns object of type 'Walls'
		"SideWall": (2505, 2, (9, 0), (), "SideWall", '{000208C8-0000-0000-C000-000000000046}'),
		"SizeWithWindow": (94, 2, (11, 0), (), "SizeWithWindow", None),
		"SubType": (109, 2, (3, 0), (), "SubType", None),
		# Method 'SurfaceGroup' returns object of type 'ChartGroup'
		"SurfaceGroup": (22, 2, (9, 0), (), "SurfaceGroup", '{00020859-0000-0000-C000-000000000046}'),
		# Method 'Tab' returns object of type 'Tab'
		"Tab": (1041, 2, (9, 0), (), "Tab", '{00024469-0000-0000-C000-000000000046}'),
		"Type": (108, 2, (3, 0), (), "Type", None),
		"Visible": (558, 2, (3, 0), (), "Visible", None),
		# Method 'Walls' returns object of type 'Walls'
		"Walls": (86, 2, (9, 0), (), "Walls", '{000208C8-0000-0000-C000-000000000046}'),
		"WallsAndGridlines2D": (210, 2, (11, 0), (), "WallsAndGridlines2D", None),
		"_CodeName": (-2147418112, 2, (8, 0), (), "_CodeName", None),
	}
	_prop_map_put_ = {
		"AutoScaling": ((107, LCID, 4, 0),()),
		"BarShape": ((1403, LCID, 4, 0),()),
		"ChartStyle": ((2509, LCID, 4, 0),()),
		"ChartType": ((1400, LCID, 4, 0),()),
		"DepthPercent": ((48, LCID, 4, 0),()),
		"DisplayBlanksAs": ((93, LCID, 4, 0),()),
		"Elevation": ((49, LCID, 4, 0),()),
		"GapDepth": ((50, LCID, 4, 0),()),
		"HasAxis": ((52, LCID, 4, 0),()),
		"HasDataTable": ((1396, LCID, 4, 0),()),
		"HasLegend": ((53, LCID, 4, 0),()),
		"HasPivotFields": ((1815, LCID, 4, 0),()),
		"HasTitle": ((54, LCID, 4, 0),()),
		"HeightPercent": ((55, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"OnDoubleClick": ((628, LCID, 4, 0),()),
		"OnSheetActivate": ((1031, LCID, 4, 0),()),
		"OnSheetDeactivate": ((1081, LCID, 4, 0),()),
		"Perspective": ((57, LCID, 4, 0),()),
		"PlotBy": ((202, LCID, 4, 0),()),
		"PlotVisibleOnly": ((92, LCID, 4, 0),()),
		"ProtectData": ((1406, LCID, 4, 0),()),
		"ProtectFormatting": ((1405, LCID, 4, 0),()),
		"ProtectGoalSeek": ((1407, LCID, 4, 0),()),
		"ProtectSelection": ((1408, LCID, 4, 0),()),
		"RightAngleAxes": ((58, LCID, 4, 0),()),
		"Rotation": ((59, LCID, 4, 0),()),
		"ShowDataLabelsOverMaximum": ((2504, LCID, 4, 0),()),
		"ShowWindow": ((1399, LCID, 4, 0),()),
		"SizeWithWindow": ((94, LCID, 4, 0),()),
		"SubType": ((109, LCID, 4, 0),()),
		"Type": ((108, LCID, 4, 0),()),
		"Visible": ((558, LCID, 4, 0),()),
		"WallsAndGridlines2D": ((210, LCID, 4, 0),()),
		"_CodeName": ((-2147418112, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D6-0000-0000-C000-000000000046}", _Chart )
# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 17:06:58 2013
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

_Chart_vtables_dispatch_ = 1
_Chart_vtables_ = [
	(( 'Application' , 'RHS' , ), 148, (148, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'RHS' , ), 149, (149, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'RHS' , ), 150, (150, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Activate' , 'lcid' , ), 304, (304, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Copy' , 'Before' , 'After' , 'lcid' , ), 551, (551, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , 'lcid' , ), 117, (117, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'CodeName' , 'RHS' , ), 1373, (1373, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( '_CodeName' , 'RHS' , ), -2147418112, (-2147418112, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 1024 , )),
	(( '_CodeName' , 'RHS' , ), -2147418112, (-2147418112, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 60 , (3, 0, None, None) , 1024 , )),
	(( 'Index' , 'lcid' , 'RHS' , ), 486, (486, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Move' , 'Before' , 'After' , 'lcid' , ), 637, (637, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 68 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'RHS' , ), 110, (110, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'RHS' , ), 110, (110, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'Next' , 'RHS' , ), 502, (502, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'OnDoubleClick' , 'lcid' , 'RHS' , ), 628, (628, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 64 , )),
	(( 'OnDoubleClick' , 'lcid' , 'RHS' , ), 628, (628, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 100 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 64 , )),
	(( 'PageSetup' , 'RHS' , ), 998, (998, (), [ (16393, 10, None, "IID('{000208B4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( 'Previous' , 'RHS' , ), 503, (503, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( '__PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'lcid' , ), 905, (905, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 7 , 116 , (3, 0, None, None) , 1088 , )),
	(( 'PrintPreview' , 'EnableChanges' , 'lcid' , ), 281, (281, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 120 , (3, 0, None, None) , 0 , )),
	(( '_Protect' , 'Password' , 'DrawingObjects' , 'Contents' , 'Scenarios' , 
			 'UserInterfaceOnly' , 'lcid' , ), 282, (282, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 5 , 124 , (3, 0, None, None) , 1088 , )),
	(( 'ProtectContents' , 'lcid' , 'RHS' , ), 292, (292, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'ProtectDrawingObjects' , 'lcid' , 'RHS' , ), 293, (293, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'ProtectionMode' , 'lcid' , 'RHS' , ), 1159, (1159, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( '_Dummy23' , ), 65559, (65559, (), [ ], 1 , 1 , 4 , 0 , 140 , (24, 0, None, None) , 1089 , )),
	(( '_SaveAs' , 'Filename' , 'FileFormat' , 'Password' , 'WriteResPassword' , 
			 'ReadOnlyRecommended' , 'CreateBackup' , 'AddToMru' , 'TextCodepage' , 'TextVisualLayout' , 
			 'lcid' , ), 284, (284, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 144 , (3, 0, None, None) , 1088 , )),
	(( 'Select' , 'Replace' , 'lcid' , ), 235, (235, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 148 , (3, 0, None, None) , 0 , )),
	(( 'Unprotect' , 'Password' , 'lcid' , ), 285, (285, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'Visible' , 'lcid' , 'RHS' , ), 558, (558, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Shapes' , 'RHS' , ), 1377, (1377, (), [ (16393, 10, None, "IID('{0002443A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( '_ApplyDataLabels' , 'Type' , 'LegendKey' , 'AutoText' , 'HasLeaderLines' , 
			 'lcid' , ), 151, (151, (), [ (3, 49, '2', None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 3 , 168 , (3, 0, None, None) , 1088 , )),
	(( 'Arcs' , 'Index' , 'lcid' , 'RHS' , ), 760, (760, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 172 , (3, 0, None, None) , 64 , )),
	(( 'Area3DGroup' , 'lcid' , 'RHS' , ), 17, (17, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020859-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 64 , )),
	(( 'AreaGroups' , 'Index' , 'lcid' , 'RHS' , ), 9, (9, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 180 , (3, 0, None, None) , 64 , )),
	(( 'AutoFormat' , 'Gallery' , 'Format' , ), 114, (114, (), [ (3, 1, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 184 , (3, 0, None, None) , 64 , )),
	(( 'AutoScaling' , 'lcid' , 'RHS' , ), 107, (107, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'AutoScaling' , 'lcid' , 'RHS' , ), 107, (107, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Axes' , 'Type' , 'AxisGroup' , 'lcid' , 'RHS' , 
			 ), 23, (23, (), [ (12, 17, None, None) , (3, 49, '1', None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'SetBackgroundPicture' , 'Filename' , ), 1188, (1188, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'Bar3DGroup' , 'lcid' , 'RHS' , ), 18, (18, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020859-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 64 , )),
	(( 'BarGroups' , 'Index' , 'lcid' , 'RHS' , ), 10, (10, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 208 , (3, 0, None, None) , 64 , )),
	(( 'Buttons' , 'Index' , 'lcid' , 'RHS' , ), 557, (557, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 212 , (3, 0, None, None) , 64 , )),
	(( 'ChartArea' , 'lcid' , 'RHS' , ), 80, (80, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{000208CC-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'ChartGroups' , 'Index' , 'lcid' , 'RHS' , ), 8, (8, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 220 , (3, 0, None, None) , 0 , )),
	(( 'ChartObjects' , 'Index' , 'lcid' , 'RHS' , ), 1060, (1060, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ChartTitle' , 'lcid' , 'RHS' , ), 81, (81, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020849-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'ChartWizard' , 'Source' , 'Gallery' , 'Format' , 'PlotBy' , 
			 'CategoryLabels' , 'SeriesLabels' , 'HasLegend' , 'Title' , 'CategoryTitle' , 
			 'ValueTitle' , 'ExtraTitle' , 'lcid' , ), 196, (196, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 11 , 232 , (3, 0, None, None) , 0 , )),
	(( 'CheckBoxes' , 'Index' , 'lcid' , 'RHS' , ), 824, (824, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 236 , (3, 0, None, None) , 64 , )),
	(( 'CheckSpelling' , 'CustomDictionary' , 'IgnoreUppercase' , 'AlwaysSuggest' , 'SpellLang' , 
			 'lcid' , ), 505, (505, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 4 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Column3DGroup' , 'lcid' , 'RHS' , ), 19, (19, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020859-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 244 , (3, 0, None, None) , 64 , )),
	(( 'ColumnGroups' , 'Index' , 'lcid' , 'RHS' , ), 11, (11, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 248 , (3, 0, None, None) , 64 , )),
	(( 'CopyPicture' , 'Appearance' , 'Format' , 'Size' , 'lcid' , 
			 ), 213, (213, (), [ (3, 49, '1', None) , (3, 49, '-4147', None) , (3, 49, '2', None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'Corners' , 'lcid' , 'RHS' , ), 79, (79, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{000208C0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 64 , )),
	(( 'CreatePublisher' , 'Edition' , 'Appearance' , 'Size' , 'ContainsPICT' , 
			 'ContainsBIFF' , 'ContainsRTF' , 'ContainsVALU' , 'lcid' , ), 458, (458, (), [ 
			 (12, 17, None, None) , (3, 49, '1', None) , (3, 49, '1', None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 4 , 260 , (3, 0, None, None) , 64 , )),
	(( 'DataTable' , 'RHS' , ), 1395, (1395, (), [ (16393, 10, None, "IID('{00020843-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'DepthPercent' , 'lcid' , 'RHS' , ), 48, (48, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'DepthPercent' , 'lcid' , 'RHS' , ), 48, (48, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'Deselect' , 'lcid' , ), 1120, (1120, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 276 , (3, 0, None, None) , 64 , )),
	(( 'DisplayBlanksAs' , 'lcid' , 'RHS' , ), 93, (93, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'DisplayBlanksAs' , 'lcid' , 'RHS' , ), 93, (93, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 284 , (3, 0, None, None) , 0 , )),
	(( 'DoughnutGroups' , 'Index' , 'lcid' , 'RHS' , ), 14, (14, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 288 , (3, 0, None, None) , 64 , )),
	(( 'Drawings' , 'Index' , 'lcid' , 'RHS' , ), 772, (772, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 292 , (3, 0, None, None) , 64 , )),
	(( 'DrawingObjects' , 'Index' , 'lcid' , 'RHS' , ), 88, (88, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 296 , (3, 0, None, None) , 64 , )),
	(( 'DropDowns' , 'Index' , 'lcid' , 'RHS' , ), 836, (836, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 300 , (3, 0, None, None) , 64 , )),
	(( 'Elevation' , 'lcid' , 'RHS' , ), 49, (49, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'Elevation' , 'lcid' , 'RHS' , ), 49, (49, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'Evaluate' , 'Name' , 'lcid' , 'RHS' , ), 1, (1, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( '_Evaluate' , 'Name' , 'lcid' , 'RHS' , ), -5, (-5, (), [ 
			 (12, 1, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 316 , (3, 0, None, None) , 1024 , )),
	(( 'Floor' , 'lcid' , 'RHS' , ), 83, (83, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{000208C7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'GapDepth' , 'lcid' , 'RHS' , ), 50, (50, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'GapDepth' , 'lcid' , 'RHS' , ), 50, (50, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'GroupBoxes' , 'Index' , 'lcid' , 'RHS' , ), 834, (834, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 332 , (3, 0, None, None) , 64 , )),
	(( 'GroupObjects' , 'Index' , 'lcid' , 'RHS' , ), 1113, (1113, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 336 , (3, 0, None, None) , 64 , )),
	(( 'HasAxis' , 'Index1' , 'Index2' , 'lcid' , 'RHS' , 
			 ), 52, (52, (), [ (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 2 , 340 , (3, 0, None, None) , 0 , )),
	(( 'HasAxis' , 'Index1' , 'Index2' , 'lcid' , 'RHS' , 
			 ), 52, (52, (), [ (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , (12, 1, None, None) , ], 1 , 4 , 4 , 2 , 344 , (3, 0, None, None) , 0 , )),
	(( 'HasDataTable' , 'RHS' , ), 1396, (1396, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 348 , (3, 0, None, None) , 0 , )),
	(( 'HasDataTable' , 'RHS' , ), 1396, (1396, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'HasLegend' , 'lcid' , 'RHS' , ), 53, (53, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'HasLegend' , 'lcid' , 'RHS' , ), 53, (53, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'HasTitle' , 'lcid' , 'RHS' , ), 54, (54, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( 'HasTitle' , 'lcid' , 'RHS' , ), 54, (54, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'HeightPercent' , 'lcid' , 'RHS' , ), 55, (55, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'HeightPercent' , 'lcid' , 'RHS' , ), 55, (55, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'Hyperlinks' , 'RHS' , ), 1393, (1393, (), [ (16393, 10, None, "IID('{00024430-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'Labels' , 'Index' , 'lcid' , 'RHS' , ), 841, (841, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 384 , (3, 0, None, None) , 64 , )),
	(( 'Legend' , 'lcid' , 'RHS' , ), 84, (84, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{000208CD-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 388 , (3, 0, None, None) , 0 , )),
	(( 'Line3DGroup' , 'lcid' , 'RHS' , ), 20, (20, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020859-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 64 , )),
	(( 'LineGroups' , 'Index' , 'lcid' , 'RHS' , ), 12, (12, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 396 , (3, 0, None, None) , 64 , )),
	(( 'Lines' , 'Index' , 'lcid' , 'RHS' , ), 767, (767, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 400 , (3, 0, None, None) , 64 , )),
	(( 'ListBoxes' , 'Index' , 'lcid' , 'RHS' , ), 832, (832, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 404 , (3, 0, None, None) , 64 , )),
	(( 'Location' , 'Where' , 'Name' , 'RHS' , ), 1397, (1397, (), [ 
			 (3, 1, None, None) , (12, 17, None, None) , (16397, 10, None, "IID('{00020821-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 408 , (3, 0, None, None) , 0 , )),
	(( 'OLEObjects' , 'Index' , 'lcid' , 'RHS' , ), 799, (799, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 412 , (3, 0, None, None) , 0 , )),
	(( 'OptionButtons' , 'Index' , 'lcid' , 'RHS' , ), 826, (826, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 416 , (3, 0, None, None) , 64 , )),
	(( 'Ovals' , 'Index' , 'lcid' , 'RHS' , ), 801, (801, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 420 , (3, 0, None, None) , 64 , )),
	(( 'Paste' , 'Type' , 'lcid' , ), 211, (211, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 424 , (3, 0, None, None) , 0 , )),
	(( 'Perspective' , 'lcid' , 'RHS' , ), 57, (57, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 428 , (3, 0, None, None) , 0 , )),
	(( 'Perspective' , 'lcid' , 'RHS' , ), 57, (57, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'Pictures' , 'Index' , 'lcid' , 'RHS' , ), 771, (771, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 436 , (3, 0, None, None) , 64 , )),
	(( 'Pie3DGroup' , 'lcid' , 'RHS' , ), 21, (21, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020859-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 64 , )),
	(( 'PieGroups' , 'Index' , 'lcid' , 'RHS' , ), 13, (13, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 444 , (3, 0, None, None) , 64 , )),
	(( 'PlotArea' , 'lcid' , 'RHS' , ), 85, (85, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{000208CB-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'PlotVisibleOnly' , 'lcid' , 'RHS' , ), 92, (92, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 452 , (3, 0, None, None) , 0 , )),
	(( 'PlotVisibleOnly' , 'lcid' , 'RHS' , ), 92, (92, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'RadarGroups' , 'Index' , 'lcid' , 'RHS' , ), 15, (15, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 460 , (3, 0, None, None) , 64 , )),
	(( 'Rectangles' , 'Index' , 'lcid' , 'RHS' , ), 774, (774, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 464 , (3, 0, None, None) , 64 , )),
	(( 'RightAngleAxes' , 'lcid' , 'RHS' , ), 58, (58, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 468 , (3, 0, None, None) , 0 , )),
	(( 'RightAngleAxes' , 'lcid' , 'RHS' , ), 58, (58, (), [ (3, 5, None, None) , 
			 (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'Rotation' , 'lcid' , 'RHS' , ), 59, (59, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 476 , (3, 0, None, None) , 0 , )),
	(( 'Rotation' , 'lcid' , 'RHS' , ), 59, (59, (), [ (3, 5, None, None) , 
			 (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'ScrollBars' , 'Index' , 'lcid' , 'RHS' , ), 830, (830, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 484 , (3, 0, None, None) , 64 , )),
	(( 'SeriesCollection' , 'Index' , 'lcid' , 'RHS' , ), 68, (68, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 488 , (3, 0, None, None) , 0 , )),
	(( 'SizeWithWindow' , 'lcid' , 'RHS' , ), 94, (94, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 492 , (3, 0, None, None) , 64 , )),
	(( 'SizeWithWindow' , 'lcid' , 'RHS' , ), 94, (94, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 496 , (3, 0, None, None) , 64 , )),
	(( 'ShowWindow' , 'RHS' , ), 1399, (1399, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 500 , (3, 0, None, None) , 64 , )),
	(( 'ShowWindow' , 'RHS' , ), 1399, (1399, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 504 , (3, 0, None, None) , 64 , )),
	(( 'Spinners' , 'Index' , 'lcid' , 'RHS' , ), 838, (838, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 508 , (3, 0, None, None) , 64 , )),
	(( 'SubType' , 'lcid' , 'RHS' , ), 109, (109, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 64 , )),
	(( 'SubType' , 'lcid' , 'RHS' , ), 109, (109, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 516 , (3, 0, None, None) , 64 , )),
	(( 'SurfaceGroup' , 'lcid' , 'RHS' , ), 22, (22, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020859-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 520 , (3, 0, None, None) , 64 , )),
	(( 'TextBoxes' , 'Index' , 'lcid' , 'RHS' , ), 777, (777, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 524 , (3, 0, None, None) , 64 , )),
	(( 'Type' , 'lcid' , 'RHS' , ), 108, (108, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 64 , )),
	(( 'Type' , 'lcid' , 'RHS' , ), 108, (108, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 532 , (3, 0, None, None) , 64 , )),
	(( 'ChartType' , 'RHS' , ), 1400, (1400, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'ChartType' , 'RHS' , ), 1400, (1400, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'ApplyCustomType' , 'ChartType' , 'TypeName' , ), 1401, (1401, (), [ (3, 1, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 544 , (3, 0, None, None) , 64 , )),
	(( 'Walls' , 'lcid' , 'RHS' , ), 86, (86, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{000208C8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 548 , (3, 0, None, None) , 0 , )),
	(( 'WallsAndGridlines2D' , 'lcid' , 'RHS' , ), 210, (210, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 552 , (3, 0, None, None) , 64 , )),
	(( 'WallsAndGridlines2D' , 'lcid' , 'RHS' , ), 210, (210, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 556 , (3, 0, None, None) , 64 , )),
	(( 'XYGroups' , 'Index' , 'lcid' , 'RHS' , ), 16, (16, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16393, 10, None, None) , ], 1 , 1 , 4 , 1 , 560 , (3, 0, None, None) , 64 , )),
	(( 'BarShape' , 'RHS' , ), 1403, (1403, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 564 , (3, 0, None, None) , 0 , )),
	(( 'BarShape' , 'RHS' , ), 1403, (1403, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'PlotBy' , 'RHS' , ), 202, (202, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'PlotBy' , 'RHS' , ), 202, (202, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'CopyChartBuild' , ), 1404, (1404, (), [ ], 1 , 1 , 4 , 0 , 580 , (3, 0, None, None) , 64 , )),
	(( 'ProtectFormatting' , 'RHS' , ), 1405, (1405, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'ProtectFormatting' , 'RHS' , ), 1405, (1405, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 588 , (3, 0, None, None) , 0 , )),
	(( 'ProtectData' , 'RHS' , ), 1406, (1406, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'ProtectData' , 'RHS' , ), 1406, (1406, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 596 , (3, 0, None, None) , 0 , )),
	(( 'ProtectGoalSeek' , 'RHS' , ), 1407, (1407, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 600 , (3, 0, None, None) , 64 , )),
	(( 'ProtectGoalSeek' , 'RHS' , ), 1407, (1407, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 604 , (3, 0, None, None) , 64 , )),
	(( 'ProtectSelection' , 'RHS' , ), 1408, (1408, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'ProtectSelection' , 'RHS' , ), 1408, (1408, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 612 , (3, 0, None, None) , 0 , )),
	(( 'GetChartElement' , 'x' , 'y' , 'ElementID' , 'Arg1' , 
			 'Arg2' , ), 1409, (1409, (), [ (3, 1, None, None) , (3, 1, None, None) , (16387, 1, None, None) , 
			 (16387, 1, None, None) , (16387, 1, None, None) , ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'SetSourceData' , 'Source' , 'PlotBy' , ), 1413, (1413, (), [ (9, 1, None, "IID('{00020846-0000-0000-C000-000000000046}')") , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 620 , (3, 0, None, None) , 0 , )),
	(( 'Export' , 'Filename' , 'FilterName' , 'Interactive' , 'RHS' , 
			 ), 1414, (1414, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 2 , 624 , (3, 0, None, None) , 0 , )),
	(( 'Refresh' , ), 1417, (1417, (), [ ], 1 , 1 , 4 , 0 , 628 , (3, 0, None, None) , 0 , )),
	(( 'PivotLayout' , 'RHS' , ), 1814, (1814, (), [ (16393, 10, None, "IID('{0002444A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'HasPivotFields' , 'RHS' , ), 1815, (1815, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 636 , (3, 0, None, None) , 64 , )),
	(( 'HasPivotFields' , 'RHS' , ), 1815, (1815, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 640 , (3, 0, None, None) , 64 , )),
	(( 'Scripts' , 'RHS' , ), 1816, (1816, (), [ (16393, 10, None, "IID('{000C0340-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 644 , (3, 0, None, None) , 64 , )),
	(( '_PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'lcid' , 
			 ), 1772, (1772, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 648 , (3, 0, None, None) , 1088 , )),
	(( 'Tab' , 'RHS' , ), 1041, (1041, (), [ (16393, 10, None, "IID('{00024469-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 652 , (3, 0, None, None) , 0 , )),
	(( 'MailEnvelope' , 'RHS' , ), 2021, (2021, (), [ (16397, 10, None, "IID('{0006F01A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'ApplyDataLabels' , 'Type' , 'LegendKey' , 'AutoText' , 'HasLeaderLines' , 
			 'ShowSeriesName' , 'ShowCategoryName' , 'ShowValue' , 'ShowPercentage' , 'ShowBubbleSize' , 
			 'Separator' , 'lcid' , ), 1922, (1922, (), [ (3, 49, '2', None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 9 , 660 , (3, 0, None, None) , 0 , )),
	(( 'SaveAs' , 'Filename' , 'FileFormat' , 'Password' , 'WriteResPassword' , 
			 'ReadOnlyRecommended' , 'CreateBackup' , 'AddToMru' , 'TextCodepage' , 'TextVisualLayout' , 
			 'Local' , ), 1925, (1925, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 9 , 664 , (3, 0, None, None) , 0 , )),
	(( 'Protect' , 'Password' , 'DrawingObjects' , 'Contents' , 'Scenarios' , 
			 'UserInterfaceOnly' , ), 2029, (2029, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 5 , 668 , (3, 0, None, None) , 0 , )),
	(( 'ApplyLayout' , 'Layout' , 'ChartType' , ), 2500, (2500, (), [ (3, 1, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 672 , (3, 0, None, None) , 0 , )),
	(( 'SetElement' , 'Element' , ), 2502, (2502, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 676 , (3, 0, None, None) , 0 , )),
	(( 'ShowDataLabelsOverMaximum' , 'RHS' , ), 2504, (2504, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'ShowDataLabelsOverMaximum' , 'RHS' , ), 2504, (2504, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 684 , (3, 0, None, None) , 0 , )),
	(( 'SideWall' , 'RHS' , ), 2505, (2505, (), [ (16393, 10, None, "IID('{000208C8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'BackWall' , 'RHS' , ), 2506, (2506, (), [ (16393, 10, None, "IID('{000208C8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 692 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'lcid' , 
			 ), 2361, (2361, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 696 , (3, 0, None, None) , 0 , )),
	(( 'ApplyChartTemplate' , 'Filename' , ), 2507, (2507, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 700 , (3, 0, None, None) , 0 , )),
	(( 'SaveChartTemplate' , 'Filename' , ), 2508, (2508, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'SetDefaultChart' , 'Name' , ), 219, (219, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 708 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat' , 'Type' , 'Filename' , 'Quality' , 'IncludeDocProperties' , 
			 'IgnorePrintAreas' , 'From' , 'To' , 'OpenAfterPublish' , 'FixedFormatExtClassPtr' , 
			 ), 2493, (2493, (), [ (3, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 8 , 712 , (3, 0, None, None) , 0 , )),
	(( 'ChartStyle' , 'RHS' , ), 2509, (2509, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 716 , (3, 0, None, None) , 0 , )),
	(( 'ChartStyle' , 'RHS' , ), 2509, (2509, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'ClearToMatchStyle' , ), 2510, (2510, (), [ ], 1 , 1 , 4 , 0 , 724 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000208D6-0000-0000-C000-000000000046}", _Chart )
