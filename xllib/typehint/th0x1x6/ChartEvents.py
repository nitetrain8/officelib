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

class ChartEvents:
	CLSID = CLSID_Sink = IID('{0002440F-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020821-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		     1534 : "OnBeforeRightClick",
		     1530 : "OnDeactivate",
		1610678275 : "OnInvoke",
		     1531 : "OnMouseDown",
		1610612736 : "OnQueryInterface",
		     1536 : "OnDragOver",
		1610612737 : "OnAddRef",
		     1532 : "OnMouseUp",
		      256 : "OnResize",
		     1533 : "OnMouseMove",
		      235 : "OnSelect",
		1610678272 : "OnGetTypeInfoCount",
		1610612738 : "OnRelease",
		     1535 : "OnDragPlot",
		      304 : "OnActivate",
		     1537 : "OnBeforeDoubleClick",
		      279 : "OnCalculate",
		     1538 : "OnSeriesChange",
		1610678274 : "OnGetIDsOfNames",
		1610678273 : "OnGetTypeInfo",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnBeforeRightClick(self, Cancel=defaultNamedNotOptArg):
#	def OnDeactivate(self):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnMouseDown(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnDragOver(self):
#	def OnAddRef(self):
#	def OnMouseUp(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#	def OnResize(self):
#	def OnMouseMove(self, Button=defaultNamedNotOptArg, Shift=defaultNamedNotOptArg, x=defaultNamedNotOptArg, y=defaultNamedNotOptArg):
#	def OnSelect(self, ElementID=defaultNamedNotOptArg, Arg1=defaultNamedNotOptArg, Arg2=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnRelease(self):
#	def OnDragPlot(self):
#	def OnActivate(self):
#	def OnBeforeDoubleClick(self, ElementID=defaultNamedNotOptArg, Arg1=defaultNamedNotOptArg, Arg2=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnCalculate(self):
#	def OnSeriesChange(self, SeriesIndex=defaultNamedNotOptArg, PointIndex=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):


win32com.client.CLSIDToClass.RegisterCLSID( "{0002440F-0000-0000-C000-000000000046}", ChartEvents )
