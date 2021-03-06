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

from win32com.client import DispatchBaseClass
class Workbooks(DispatchBaseClass):
	CLSID = IID('{000208DB-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Workbook
	def Add(self, Template=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(181, LCID, 1, (13, 0), ((12, 17),),Template
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Add', '{00020819-0000-0000-C000-000000000046}')
		return ret

	def CanCheckOut(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2070, LCID, 1, (11, 0), ((8, 1),),Filename
			)

	def CheckOut(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2069, LCID, 1, (24, 0), ((8, 1),),Filename
			)

	def Close(self):
		return self._oleobj_.InvokeTypes(277, LCID, 1, (24, 0), (),)

	# Result is of type Workbook
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(170, LCID, 2, (13, 0), ((12, 1),),Index
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Item', '{00020819-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Workbook
	def Open(self, Filename=defaultNamedNotOptArg, UpdateLinks=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, Format=defaultNamedOptArg
			, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg, IgnoreReadOnlyRecommended=defaultNamedOptArg, Origin=defaultNamedOptArg, Delimiter=defaultNamedOptArg
			, Editable=defaultNamedOptArg, Notify=defaultNamedOptArg, Converter=defaultNamedOptArg, AddToMru=defaultNamedOptArg, Local=defaultNamedOptArg
			, CorruptLoad=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1923, LCID, 1, (13, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, UpdateLinks, ReadOnly, Format, Password, WriteResPassword
			, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify
			, Converter, AddToMru, Local, CorruptLoad)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'Open', '{00020819-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Workbook
	def OpenDatabase(self, Filename=defaultNamedNotOptArg, CommandText=defaultNamedOptArg, CommandType=defaultNamedOptArg, BackgroundQuery=defaultNamedOptArg
			, ImportDataAs=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2067, LCID, 1, (13, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, CommandText, CommandType, BackgroundQuery, ImportDataAs)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'OpenDatabase', '{00020819-0000-0000-C000-000000000046}')
		return ret

	def OpenText(self, Filename=defaultNamedNotOptArg, Origin=defaultNamedNotOptArg, StartRow=defaultNamedNotOptArg, DataType=defaultNamedNotOptArg
			, TextQualifier=1, ConsecutiveDelimiter=defaultNamedOptArg, Tab=defaultNamedOptArg, Semicolon=defaultNamedOptArg, Comma=defaultNamedOptArg
			, Space=defaultNamedOptArg, Other=defaultNamedOptArg, OtherChar=defaultNamedOptArg, FieldInfo=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg
			, DecimalSeparator=defaultNamedOptArg, ThousandsSeparator=defaultNamedOptArg, TrailingMinusNumbers=defaultNamedOptArg, Local=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1924, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter
			, Tab, Semicolon, Comma, Space, Other
			, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator
			, TrailingMinusNumbers, Local)

	# Result is of type Workbook
	def OpenXML(self, Filename=defaultNamedNotOptArg, Stylesheets=defaultNamedOptArg, LoadOption=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2280, LCID, 1, (13, 0), ((8, 1), (12, 17), (12, 17)),Filename
			, Stylesheets, LoadOption)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, 'OpenXML', '{00020819-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Workbook
	# The method _Default is actually a property, but must be used as a method to correctly pass the arguments
	def _Default(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (13, 0), ((12, 1),),Index
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, '_Default', '{00020819-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Workbook
	def _Open(self, Filename=defaultNamedNotOptArg, UpdateLinks=defaultNamedOptArg, ReadOnly=defaultNamedOptArg, Format=defaultNamedOptArg
			, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg, IgnoreReadOnlyRecommended=defaultNamedOptArg, Origin=defaultNamedOptArg, Delimiter=defaultNamedOptArg
			, Editable=defaultNamedOptArg, Notify=defaultNamedOptArg, Converter=defaultNamedOptArg, AddToMru=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(682, LCID, 1, (13, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, UpdateLinks, ReadOnly, Format, Password, WriteResPassword
			, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify
			, Converter, AddToMru)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, '_Open', '{00020819-0000-0000-C000-000000000046}')
		return ret

	def _OpenText(self, Filename=defaultNamedNotOptArg, Origin=defaultNamedNotOptArg, StartRow=defaultNamedNotOptArg, DataType=defaultNamedNotOptArg
			, TextQualifier=1, ConsecutiveDelimiter=defaultNamedOptArg, Tab=defaultNamedOptArg, Semicolon=defaultNamedOptArg, Comma=defaultNamedOptArg
			, Space=defaultNamedOptArg, Other=defaultNamedOptArg, OtherChar=defaultNamedOptArg, FieldInfo=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg
			, DecimalSeparator=defaultNamedOptArg, ThousandsSeparator=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1773, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter
			, Tab, Semicolon, Comma, Space, Other
			, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator
			)

	# Result is of type Workbook
	def _OpenXML(self, Filename=defaultNamedNotOptArg, Stylesheets=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(2071, LCID, 1, (13, 0), ((8, 1), (12, 17)),Filename
			, Stylesheets)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, '_OpenXML', '{00020819-0000-0000-C000-000000000046}')
		return ret

	def _OpenText_(self, Filename=defaultNamedNotOptArg, Origin=defaultNamedNotOptArg, StartRow=defaultNamedNotOptArg, DataType=defaultNamedNotOptArg
			, TextQualifier=1, ConsecutiveDelimiter=defaultNamedOptArg, Tab=defaultNamedOptArg, Semicolon=defaultNamedOptArg, Comma=defaultNamedOptArg
			, Space=defaultNamedOptArg, Other=defaultNamedOptArg, OtherChar=defaultNamedOptArg, FieldInfo=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(683, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter
			, Tab, Semicolon, Comma, Space, Other
			, OtherChar, FieldInfo, TextVisualLayout)

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
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (13, 0), ((12, 1),),Index
			)
		if ret is not None:
			# See if this IUnknown is really an IDispatch
			try:
				ret = ret.QueryInterface(pythoncom.IID_IDispatch)
			except pythoncom.error:
				return ret
			ret = Dispatch(ret, '__call__', '{00020819-0000-0000-C000-000000000046}')
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
		return win32com.client.util.Iterator(ob, '{00020819-0000-0000-C000-000000000046}')
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		return self._get_good_object_(self._oleobj_.Invoke(*(170, LCID, 2, 1, key)), "Item", '{00020819-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(118, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{000208DB-0000-0000-C000-000000000046}", Workbooks )
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

Workbooks_vtables_dispatch_ = 1
Workbooks_vtables_ = [
	(( 'Application' , 'RHS' , ), 148, (148, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'RHS' , ), 149, (149, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'RHS' , ), 150, (150, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'Add' , 'Template' , 'lcid' , 'RHS' , ), 181, (181, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 40 , (3, 0, None, None) , 0 , )),
	(( 'Close' , 'lcid' , ), 277, (277, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( 'Count' , 'RHS' , ), 118, (118, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'RHS' , ), 170, (170, (), [ (12, 1, None, None) , 
			 (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( '_NewEnum' , 'RHS' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 1024 , )),
	(( '_Open' , 'Filename' , 'UpdateLinks' , 'ReadOnly' , 'Format' , 
			 'Password' , 'WriteResPassword' , 'IgnoreReadOnlyRecommended' , 'Origin' , 'Delimiter' , 
			 'Editable' , 'Notify' , 'Converter' , 'AddToMru' , 'lcid' , 
			 'RHS' , ), 682, (682, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 12 , 60 , (3, 0, None, None) , 1088 , )),
	(( '__OpenText' , 'Filename' , 'Origin' , 'StartRow' , 'DataType' , 
			 'TextQualifier' , 'ConsecutiveDelimiter' , 'Tab' , 'Semicolon' , 'Comma' , 
			 'Space' , 'Other' , 'OtherChar' , 'FieldInfo' , 'TextVisualLayout' , 
			 'lcid' , ), 683, (683, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 49, '1', None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 9 , 64 , (3, 0, None, None) , 1088 , )),
	(( '_Default' , 'Index' , 'RHS' , ), 0, (0, (), [ (12, 1, None, None) , 
			 (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 1024 , )),
	(( '_OpenText' , 'Filename' , 'Origin' , 'StartRow' , 'DataType' , 
			 'TextQualifier' , 'ConsecutiveDelimiter' , 'Tab' , 'Semicolon' , 'Comma' , 
			 'Space' , 'Other' , 'OtherChar' , 'FieldInfo' , 'TextVisualLayout' , 
			 'DecimalSeparator' , 'ThousandsSeparator' , 'lcid' , ), 1773, (1773, (), [ (8, 1, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 49, '1', None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 11 , 72 , (3, 0, None, None) , 1088 , )),
	(( 'Open' , 'Filename' , 'UpdateLinks' , 'ReadOnly' , 'Format' , 
			 'Password' , 'WriteResPassword' , 'IgnoreReadOnlyRecommended' , 'Origin' , 'Delimiter' , 
			 'Editable' , 'Notify' , 'Converter' , 'AddToMru' , 'Local' , 
			 'CorruptLoad' , 'lcid' , 'RHS' , ), 1923, (1923, (), [ (8, 1, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , 
			 (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 14 , 76 , (3, 0, None, None) , 0 , )),
	(( 'OpenText' , 'Filename' , 'Origin' , 'StartRow' , 'DataType' , 
			 'TextQualifier' , 'ConsecutiveDelimiter' , 'Tab' , 'Semicolon' , 'Comma' , 
			 'Space' , 'Other' , 'OtherChar' , 'FieldInfo' , 'TextVisualLayout' , 
			 'DecimalSeparator' , 'ThousandsSeparator' , 'TrailingMinusNumbers' , 'Local' , 'lcid' , 
			 ), 1924, (1924, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 49, '1', None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 13 , 80 , (3, 0, None, None) , 0 , )),
	(( 'OpenDatabase' , 'Filename' , 'CommandText' , 'CommandType' , 'BackgroundQuery' , 
			 'ImportDataAs' , 'RHS' , ), 2067, (2067, (), [ (8, 1, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 4 , 84 , (3, 0, None, None) , 0 , )),
	(( 'CheckOut' , 'Filename' , ), 2069, (2069, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'CanCheckOut' , 'Filename' , 'RHS' , ), 2070, (2070, (), [ (8, 1, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( '_OpenXML' , 'Filename' , 'Stylesheets' , 'RHS' , ), 2071, (2071, (), [ 
			 (8, 1, None, None) , (12, 17, None, None) , (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 96 , (3, 0, None, None) , 1088 , )),
	(( 'OpenXML' , 'Filename' , 'Stylesheets' , 'LoadOption' , 'RHS' , 
			 ), 2280, (2280, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16397, 10, None, "IID('{00020819-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 100 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000208DB-0000-0000-C000-000000000046}", Workbooks )
