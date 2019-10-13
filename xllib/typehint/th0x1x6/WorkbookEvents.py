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

class WorkbookEvents:
	CLSID = CLSID_Sink = IID('{00024412-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020819-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		     1556 : "OnWindowActivate",
		     2157 : "OnSheetPivotTableUpdate",
		     2159 : "OnPivotTableOpenConnection",
		     1552 : "OnAddinInstall",
		     1550 : "OnNewSheet",
		     1546 : "OnBeforeClose",
		     1560 : "OnSheetBeforeRightClick",
		     1564 : "OnSheetChange",
		     2287 : "OnBeforeXmlExport",
		     2283 : "OnBeforeXmlImport",
		     2266 : "OnSync",
		1610678274 : "OnGetIDsOfNames",
		     1559 : "OnSheetBeforeDoubleClick",
		     1923 : "OnOpen",
		1610678273 : "OnGetTypeInfo",
		     1554 : "OnWindowResize",
		     2610 : "OnRowsetComplete",
		     2285 : "OnAfterXmlImport",
		     1563 : "OnSheetCalculate",
		     1854 : "OnSheetFollowHyperlink",
		     1557 : "OnWindowDeactivate",
		     1561 : "OnSheetActivate",
		     1530 : "OnDeactivate",
		     1549 : "OnBeforePrint",
		     1553 : "OnAddinUninstall",
		     1558 : "OnSheetSelectionChange",
		1610678272 : "OnGetTypeInfoCount",
		     2158 : "OnPivotTableCloseConnection",
		1610678275 : "OnInvoke",
		1610612736 : "OnQueryInterface",
		1610612738 : "OnRelease",
		      304 : "OnActivate",
		     1562 : "OnSheetDeactivate",
		     1547 : "OnBeforeSave",
		     2288 : "OnAfterXmlExport",
		1610612737 : "OnAddRef",
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
#	def OnWindowActivate(self, Wn=defaultNamedNotOptArg):
#	def OnSheetPivotTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnPivotTableOpenConnection(self, Target=defaultNamedNotOptArg):
#	def OnAddinInstall(self):
#	def OnNewSheet(self, Sh=defaultNamedNotOptArg):
#	def OnBeforeClose(self, Cancel=defaultNamedNotOptArg):
#	def OnSheetBeforeRightClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnBeforeXmlExport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnBeforeXmlImport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSync(self, SyncEventType=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnSheetBeforeDoubleClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnOpen(self):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnWindowResize(self, Wn=defaultNamedNotOptArg):
#	def OnRowsetComplete(self, Description=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
#	def OnAfterXmlImport(self, Map=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnSheetCalculate(self, Sh=defaultNamedNotOptArg):
#	def OnSheetFollowHyperlink(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Wn=defaultNamedNotOptArg):
#	def OnSheetActivate(self, Sh=defaultNamedNotOptArg):
#	def OnDeactivate(self):
#	def OnBeforePrint(self, Cancel=defaultNamedNotOptArg):
#	def OnAddinUninstall(self):
#	def OnSheetSelectionChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnPivotTableCloseConnection(self, Target=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnRelease(self):
#	def OnActivate(self):
#	def OnSheetDeactivate(self, Sh=defaultNamedNotOptArg):
#	def OnBeforeSave(self, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnAfterXmlExport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnAddRef(self):


win32com.client.CLSIDToClass.RegisterCLSID( "{00024412-0000-0000-C000-000000000046}", WorkbookEvents )
