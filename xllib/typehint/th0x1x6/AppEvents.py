# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct  7 13:27:55 2013
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

class AppEvents:
	CLSID = CLSID_Sink = IID('{00024413-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00024500-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		     2161 : "OnWorkbookPivotTableOpenConnection",
		     1567 : "OnWorkbookOpen",
		1610678274 : "OnGetIDsOfNames",
		1610612736 : "OnQueryInterface",
		     2293 : "OnWorkbookAfterXmlExport",
		1610678275 : "OnInvoke",
		     1559 : "OnSheetBeforeDoubleClick",
		     1575 : "OnWorkbookAddinUninstall",
		     2157 : "OnSheetPivotTableUpdate",
		     2290 : "OnWorkbookBeforeXmlImport",
		     2611 : "OnWorkbookRowsetComplete",
		     1556 : "OnWindowActivate",
		     1558 : "OnSheetSelectionChange",
		     2292 : "OnWorkbookBeforeXmlExport",
		     1565 : "OnNewWorkbook",
		     1568 : "OnWorkbookActivate",
		     1569 : "OnWorkbookDeactivate",
		1610612738 : "OnRelease",
		     2160 : "OnWorkbookPivotTableCloseConnection",
		     1571 : "OnWorkbookBeforeSave",
		     1574 : "OnWorkbookAddinInstall",
		     1557 : "OnWindowDeactivate",
		     1573 : "OnWorkbookNewSheet",
		     1554 : "OnWindowResize",
		1610678272 : "OnGetTypeInfoCount",
		     1563 : "OnSheetCalculate",
		     1561 : "OnSheetActivate",
		     2289 : "OnWorkbookSync",
		     1564 : "OnSheetChange",
		1610678273 : "OnGetTypeInfo",
		     2291 : "OnWorkbookAfterXmlImport",
		     1562 : "OnSheetDeactivate",
		1610612737 : "OnAddRef",
		     1572 : "OnWorkbookBeforePrint",
		     2612 : "OnAfterCalculate",
		     1560 : "OnSheetBeforeRightClick",
		     1854 : "OnSheetFollowHyperlink",
		     1570 : "OnWorkbookBeforeClose",
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
#	def OnWorkbookPivotTableOpenConnection(self, Wb=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookOpen(self, Wb=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnWorkbookAfterXmlExport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnSheetBeforeDoubleClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWorkbookAddinUninstall(self, Wb=defaultNamedNotOptArg):
#	def OnSheetPivotTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookBeforeXmlImport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnWorkbookRowsetComplete(self, Wb=defaultNamedNotOptArg, Description=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnSheetSelectionChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookBeforeXmlExport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnNewWorkbook(self, Wb=defaultNamedNotOptArg):
#	def OnWorkbookActivate(self, Wb=defaultNamedNotOptArg):
#	def OnWorkbookDeactivate(self, Wb=defaultNamedNotOptArg):
#	def OnRelease(self):
#	def OnWorkbookPivotTableCloseConnection(self, Wb=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookBeforeSave(self, Wb=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWorkbookAddinInstall(self, Wb=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnWorkbookNewSheet(self, Wb=defaultNamedNotOptArg, Sh=defaultNamedNotOptArg):
#	def OnWindowResize(self, Wb=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnSheetCalculate(self, Sh=defaultNamedNotOptArg):
#	def OnSheetActivate(self, Sh=defaultNamedNotOptArg):
#	def OnWorkbookSync(self, Wb=defaultNamedNotOptArg, SyncEventType=defaultNamedNotOptArg):
#	def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnWorkbookAfterXmlImport(self, Wb=defaultNamedNotOptArg, Map=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnSheetDeactivate(self, Sh=defaultNamedNotOptArg):
#	def OnAddRef(self):
#	def OnWorkbookBeforePrint(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnAfterCalculate(self):
#	def OnSheetBeforeRightClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetFollowHyperlink(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnWorkbookBeforeClose(self, Wb=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00024413-0000-0000-C000-000000000046}", AppEvents )
