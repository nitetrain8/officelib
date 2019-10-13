# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 17:09:40 2013
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
class SeriesCollection(DispatchBaseClass):
	CLSID = IID('{0002086C-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Series
	def Add(self, Source=defaultNamedNotOptArg, Rowcol=-4105, SeriesLabels=defaultNamedOptArg, CategoryLabels=defaultNamedOptArg
			, Replace=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(181, LCID, 1, (9, 0), ((12, 1), (3, 49), (12, 17), (12, 17), (12, 17)),Source
			, Rowcol, SeriesLabels, CategoryLabels, Replace)
		if ret is not None:
			ret = Dispatch(ret, 'Add', '{0002086B-0000-0000-C000-000000000046}')
		return ret

	def Extend(self, Source=defaultNamedNotOptArg, Rowcol=defaultNamedOptArg, CategoryLabels=defaultNamedOptArg):
		return self._ApplyTypes_(227, 1, (12, 0), ((12, 1), (12, 17), (12, 17)), 'Extend', None,Source
			, Rowcol, CategoryLabels)

	# Result is of type Series
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(170, LCID, 1, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{0002086B-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Series
	def NewSeries(self):
		ret = self._oleobj_.InvokeTypes(1117, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'NewSeries', '{0002086B-0000-0000-C000-000000000046}')
		return ret

	def Paste(self, Rowcol=-4105, SeriesLabels=defaultNamedOptArg, CategoryLabels=defaultNamedOptArg, Replace=defaultNamedOptArg
			, NewSeries=defaultNamedOptArg):
		return self._ApplyTypes_(211, 1, (12, 0), ((3, 49), (12, 17), (12, 17), (12, 17), (12, 17)), 'Paste', None,Rowcol
			, SeriesLabels, CategoryLabels, Replace, NewSeries)

	# Result is of type Series
	def _Default(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '_Default', '{0002086B-0000-0000-C000-000000000046}')
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
			ret = Dispatch(ret, '__call__', '{0002086B-0000-0000-C000-000000000046}')
		return ret

	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, '{0002086B-0000-0000-C000-000000000046}')
	#This class has Item property/method which allows indexed access with the object[key] syntax.
	#Some objects will accept a string or other type of key in addition to integers.
	#Note that many Office objects do not use zero-based indexing.
	def __getitem__(self, key):
		return self._get_good_object_(self._oleobj_.Invoke(*(170, LCID, 1, 1, key)), "Item", '{0002086B-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(118, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002086C-0000-0000-C000-000000000046}", SeriesCollection )
