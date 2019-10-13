# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Tue Nov  5 11:29:32 2013
'Microsoft Word 12.0 Object Library'
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

CLSID = IID('{00020905-0000-0000-C000-000000000046}')
MajorVersion = 8
MinorVersion = 4
LibraryFlags = 8
LCID = 0x0

class ApplicationEvents3:
	CLSID = CLSID_Sink = IID('{00020A00-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{000209FF-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        2 : "OnQuit",
		        7 : "OnDocumentBeforePrint",
		        4 : "OnDocumentOpen",
		        3 : "OnDocumentChange",
		       15 : "OnEPostagePropertyDialog",
		1610678275 : "OnInvoke",
		1610612736 : "OnQueryInterface",
		       13 : "OnWindowBeforeRightClick",
		       19 : "OnMailMergeBeforeMerge",
		       21 : "OnMailMergeDataSourceLoad",
		1610678274 : "OnGetIDsOfNames",
		1610678273 : "OnGetTypeInfo",
		1610612738 : "OnRelease",
		        1 : "OnStartup",
		       14 : "OnWindowBeforeDoubleClick",
		       24 : "OnMailMergeWizardStateChange",
		1610612737 : "OnAddRef",
		       22 : "OnMailMergeDataSourceValidate",
		       10 : "OnWindowActivate",
		       18 : "OnMailMergeAfterRecordMerge",
		        9 : "OnNewDocument",
		       12 : "OnWindowSelectionChange",
		        8 : "OnDocumentBeforeSave",
		       23 : "OnMailMergeWizardSendToCustom",
		       11 : "OnWindowDeactivate",
		        6 : "OnDocumentBeforeClose",
		       16 : "OnEPostageInsert",
		       20 : "OnMailMergeBeforeRecordMerge",
		       25 : "OnWindowSize",
		1610678272 : "OnGetTypeInfoCount",
		       17 : "OnMailMergeAfterMerge",
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
#	def OnQuit(self):
#	def OnDocumentBeforePrint(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnDocumentOpen(self, Doc=defaultNamedNotOptArg):
#	def OnDocumentChange(self):
#	def OnEPostagePropertyDialog(self, Doc=defaultNamedNotOptArg):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnWindowBeforeRightClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeBeforeMerge(self, Doc=defaultNamedNotOptArg, StartRecord=defaultNamedNotOptArg, EndRecord=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceLoad(self, Doc=defaultNamedNotOptArg):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnRelease(self):
#	def OnStartup(self):
#	def OnWindowBeforeDoubleClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeWizardStateChange(self, Doc=defaultNamedNotOptArg, FromState=defaultNamedNotOptArg, ToState=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnAddRef(self):
#	def OnMailMergeDataSourceValidate(self, Doc=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnMailMergeAfterRecordMerge(self, Doc=defaultNamedNotOptArg):
#	def OnNewDocument(self, Doc=defaultNamedNotOptArg):
#	def OnWindowSelectionChange(self, Sel=defaultNamedNotOptArg):
#	def OnDocumentBeforeSave(self, Doc=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeWizardSendToCustom(self, Doc=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnDocumentBeforeClose(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnEPostageInsert(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeBeforeRecordMerge(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowSize(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnMailMergeAfterMerge(self, Doc=defaultNamedNotOptArg, DocResult=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00020A00-0000-0000-C000-000000000046}", ApplicationEvents3 )
