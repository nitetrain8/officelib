# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 16:38:45 2013
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
class Shapes(DispatchBaseClass):
	CLSID = IID('{0002443A-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Shape
	def AddCallout(self, Type=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
			, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1713, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Type
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddCallout', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddCanvas(self, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(2177, LCID, 1, (9, 0), ((4, 1), (4, 1), (4, 1), (4, 1)),Left
			, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddCanvas', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddChart(self, XlChartType=defaultNamedOptArg, Left=defaultNamedOptArg, Top=defaultNamedOptArg, Width=defaultNamedOptArg
			, Height=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2665, LCID, 1, (9, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),XlChartType
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddChart', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddConnector(self, Type=defaultNamedNotOptArg, BeginX=defaultNamedNotOptArg, BeginY=defaultNamedNotOptArg, EndX=defaultNamedNotOptArg
			, EndY=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1714, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Type
			, BeginX, BeginY, EndX, EndY)
		if ret is not None:
			ret = Dispatch(ret, 'AddConnector', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddCurve(self, SafeArrayOfPoints=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1719, LCID, 1, (9, 0), ((12, 1),),SafeArrayOfPoints
			)
		if ret is not None:
			ret = Dispatch(ret, 'AddCurve', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddDiagram(self, Type=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
			, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(2176, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Type
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddDiagram', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddFormControl(self, Type=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
			, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1738, LCID, 1, (9, 0), ((3, 1), (3, 1), (3, 1), (3, 1), (3, 1)),Type
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddFormControl', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddLabel(self, Orientation=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
			, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1721, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Orientation
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddLabel', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddLine(self, BeginX=defaultNamedNotOptArg, BeginY=defaultNamedNotOptArg, EndX=defaultNamedNotOptArg, EndY=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1722, LCID, 1, (9, 0), ((4, 1), (4, 1), (4, 1), (4, 1)),BeginX
			, BeginY, EndX, EndY)
		if ret is not None:
			ret = Dispatch(ret, 'AddLine', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddOLEObject(self, ClassType=defaultNamedOptArg, Filename=defaultNamedOptArg, Link=defaultNamedOptArg, DisplayAsIcon=defaultNamedOptArg
			, IconFileName=defaultNamedOptArg, IconIndex=defaultNamedOptArg, IconLabel=defaultNamedOptArg, Left=defaultNamedOptArg, Top=defaultNamedOptArg
			, Width=defaultNamedOptArg, Height=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1739, LCID, 1, (9, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),ClassType
			, Filename, Link, DisplayAsIcon, IconFileName, IconIndex
			, IconLabel, Left, Top, Width, Height
			)
		if ret is not None:
			ret = Dispatch(ret, 'AddOLEObject', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddPicture(self, Filename=defaultNamedNotOptArg, LinkToFile=defaultNamedNotOptArg, SaveWithDocument=defaultNamedNotOptArg, Left=defaultNamedNotOptArg
			, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1723, LCID, 1, (9, 0), ((8, 1), (3, 1), (3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Filename
			, LinkToFile, SaveWithDocument, Left, Top, Width
			, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddPicture', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddPolyline(self, SafeArrayOfPoints=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1726, LCID, 1, (9, 0), ((12, 1),),SafeArrayOfPoints
			)
		if ret is not None:
			ret = Dispatch(ret, 'AddPolyline', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddShape(self, Type=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
			, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1727, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Type
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddShape', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddTextEffect(self, PresetTextEffect=defaultNamedNotOptArg, Text=defaultNamedNotOptArg, FontName=defaultNamedNotOptArg, FontSize=defaultNamedNotOptArg
			, FontBold=defaultNamedNotOptArg, FontItalic=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1728, LCID, 1, (9, 0), ((3, 1), (8, 1), (8, 1), (4, 1), (3, 1), (3, 1), (4, 1), (4, 1)),PresetTextEffect
			, Text, FontName, FontSize, FontBold, FontItalic
			, Left, Top)
		if ret is not None:
			ret = Dispatch(ret, 'AddTextEffect', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def AddTextbox(self, Orientation=defaultNamedNotOptArg, Left=defaultNamedNotOptArg, Top=defaultNamedNotOptArg, Width=defaultNamedNotOptArg
			, Height=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1734, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1), (4, 1), (4, 1)),Orientation
			, Left, Top, Width, Height)
		if ret is not None:
			ret = Dispatch(ret, 'AddTextbox', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type FreeformBuilder
	def BuildFreeform(self, EditingType=defaultNamedNotOptArg, X1=defaultNamedNotOptArg, Y1=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(1735, LCID, 1, (9, 0), ((3, 1), (4, 1), (4, 1)),EditingType
			, X1, Y1)
		if ret is not None:
			ret = Dispatch(ret, 'BuildFreeform', '{0002443F-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Shape
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(170, LCID, 1, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{00024439-0000-0000-C000-000000000046}')
		return ret

	# Result is of type ShapeRange
	# The method Range is actually a property, but must be used as a method to correctly pass the arguments
	def Range(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(197, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Range', '{0002443B-0000-0000-C000-000000000046}')
		return ret

	def SelectAll(self):
		return self._oleobj_.InvokeTypes(1737, LCID, 1, (24, 0), (),)

	# Result is of type Shape
	def _Default(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '_Default', '{00024439-0000-0000-C000-000000000046}')
		return ret

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"Count": (118, 2, (3, 0), (), "Count", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is '_Default'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00024439-0000-0000-C000-000000000046}')
		return ret

	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,2,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, '{00024439-0000-0000-C000-000000000046}')
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		return self._get_good_object_(self._oleobj_.Invoke(*(170, LCID, 1, 1, key)), "Item", '{00024439-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(118, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002443A-0000-0000-C000-000000000046}", Shapes )
