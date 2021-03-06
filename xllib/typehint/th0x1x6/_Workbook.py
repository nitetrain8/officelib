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
class _Workbook(DispatchBaseClass):
	CLSID = IID('{000208DA-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020819-0000-0000-C000-000000000046}')

	def AcceptAllChanges(self, When=defaultNamedOptArg, Who=defaultNamedOptArg, Where=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1466, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),When
			, Who, Where)

	def Activate(self):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0), (),)

	def AddToFavorites(self):
		return self._oleobj_.InvokeTypes(1476, LCID, 1, (24, 0), (),)

	def ApplyTheme(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2534, LCID, 1, (24, 0), ((8, 1),),Filename
			)

	def BreakLink(self, Name=defaultNamedNotOptArg, Type=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2047, LCID, 1, (24, 0), ((8, 1), (3, 1)),Name
			, Type)

	def CanCheckIn(self):
		return self._oleobj_.InvokeTypes(2053, LCID, 1, (11, 0), (),)

	def ChangeFileAccess(self, Mode=defaultNamedNotOptArg, WritePassword=defaultNamedOptArg, Notify=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(989, LCID, 1, (24, 0), ((3, 1), (12, 17), (12, 17)),Mode
			, WritePassword, Notify)

	def ChangeLink(self, Name=defaultNamedNotOptArg, NewName=defaultNamedNotOptArg, Type=1):
		return self._oleobj_.InvokeTypes(802, LCID, 1, (24, 0), ((8, 1), (8, 1), (3, 49)),Name
			, NewName, Type)

	def CheckIn(self, SaveChanges=defaultNamedOptArg, Comments=defaultNamedOptArg, MakePublic=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2051, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),SaveChanges
			, Comments, MakePublic)

	def CheckInWithVersion(self, SaveChanges=defaultNamedOptArg, Comments=defaultNamedOptArg, MakePublic=defaultNamedOptArg, VersionType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2517, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17)),SaveChanges
			, Comments, MakePublic, VersionType)

	def Close(self, SaveChanges=defaultNamedOptArg, Filename=defaultNamedOptArg, RouteWorkbook=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(277, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),SaveChanges
			, Filename, RouteWorkbook)

	def DeleteNumberFormat(self, NumberFormat=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(397, LCID, 1, (24, 0), ((8, 1),),NumberFormat
			)

	def Dummy16(self):
		return self._oleobj_.InvokeTypes(2048, LCID, 1, (24, 0), (),)

	def Dummy17(self, calcid=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2044, LCID, 1, (24, 0), ((3, 1),),calcid
			)

	def EnableConnections(self):
		return self._oleobj_.InvokeTypes(2537, LCID, 1, (24, 0), (),)

	def EndReview(self):
		return self._oleobj_.InvokeTypes(2058, LCID, 1, (24, 0), (),)

	def ExclusiveAccess(self):
		return self._oleobj_.InvokeTypes(1168, LCID, 1, (11, 0), (),)

	def ExportAsFixedFormat(self, Type=defaultNamedNotOptArg, Filename=defaultNamedOptArg, Quality=defaultNamedOptArg, IncludeDocProperties=defaultNamedOptArg
			, IgnorePrintAreas=defaultNamedOptArg, From=defaultNamedOptArg, To=defaultNamedOptArg, OpenAfterPublish=defaultNamedOptArg, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2493, LCID, 1, (24, 0), ((3, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Type
			, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From
			, To, OpenAfterPublish, FixedFormatExtClassPtr)

	def FollowHyperlink(self, Address=defaultNamedNotOptArg, SubAddress=defaultNamedOptArg, NewWindow=defaultNamedOptArg, AddHistory=defaultNamedOptArg
			, ExtraInfo=defaultNamedOptArg, Method=defaultNamedOptArg, HeaderInfo=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1470, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Address
			, SubAddress, NewWindow, AddHistory, ExtraInfo, Method
			, HeaderInfo)

	def ForwardMailer(self):
		return self._oleobj_.InvokeTypes(973, LCID, 1, (24, 0), (),)

	# The method GetColors is actually a property, but must be used as a method to correctly pass the arguments
	def GetColors(self, Index=defaultNamedOptArg):
		return self._ApplyTypes_(286, 2, (12, 0), ((12, 17),), 'GetColors', None,Index
			)

	# Result is of type WorkflowTasks
	def GetWorkflowTasks(self):
		ret = self._oleobj_.InvokeTypes(2522, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'GetWorkflowTasks', '{000CD901-0000-0000-C000-000000000046}')
		return ret

	# Result is of type WorkflowTemplates
	def GetWorkflowTemplates(self):
		ret = self._oleobj_.InvokeTypes(2523, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'GetWorkflowTemplates', '{000CD903-0000-0000-C000-000000000046}')
		return ret

	def HighlightChangesOptions(self, When=defaultNamedOptArg, Who=defaultNamedOptArg, Where=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1458, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),When
			, Who, Where)

	def LinkInfo(self, Name=defaultNamedNotOptArg, LinkInfo=defaultNamedNotOptArg, Type=defaultNamedOptArg, EditionRef=defaultNamedOptArg):
		return self._ApplyTypes_(807, 1, (12, 0), ((8, 1), (3, 1), (12, 17), (12, 17)), 'LinkInfo', None,Name
			, LinkInfo, Type, EditionRef)

	def LinkSources(self, Type=defaultNamedOptArg):
		return self._ApplyTypes_(808, 1, (12, 0), ((12, 17),), 'LinkSources', None,Type
			)

	def LockServerFile(self):
		return self._oleobj_.InvokeTypes(2520, LCID, 1, (24, 0), (),)

	def MergeWorkbook(self, Filename=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1446, LCID, 1, (24, 0), ((12, 1),),Filename
			)

	# Result is of type Window
	def NewWindow(self):
		ret = self._oleobj_.InvokeTypes(280, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'NewWindow', '{00020893-0000-0000-C000-000000000046}')
		return ret

	def OpenLinks(self, Name=defaultNamedNotOptArg, ReadOnly=defaultNamedOptArg, Type=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(803, LCID, 1, (24, 0), ((8, 1), (12, 17), (12, 17)),Name
			, ReadOnly, Type)

	# Result is of type PivotCaches
	def PivotCaches(self):
		ret = self._oleobj_.InvokeTypes(1449, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'PivotCaches', '{0002441D-0000-0000-C000-000000000046}')
		return ret

	def PivotTableWizard(self, SourceType=defaultNamedOptArg, SourceData=defaultNamedOptArg, TableDestination=defaultNamedOptArg, TableName=defaultNamedOptArg
			, RowGrand=defaultNamedOptArg, ColumnGrand=defaultNamedOptArg, SaveData=defaultNamedOptArg, HasAutoFormat=defaultNamedOptArg, AutoPage=defaultNamedOptArg
			, Reserved=defaultNamedOptArg, BackgroundQuery=defaultNamedOptArg, OptimizeCache=defaultNamedOptArg, PageFieldOrder=defaultNamedOptArg, PageFieldWrapCount=defaultNamedOptArg
			, ReadData=defaultNamedOptArg, Connection=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(684, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),SourceType
			, SourceData, TableDestination, TableName, RowGrand, ColumnGrand
			, SaveData, HasAutoFormat, AutoPage, Reserved, BackgroundQuery
			, OptimizeCache, PageFieldOrder, PageFieldWrapCount, ReadData, Connection
			)

	def Post(self, DestName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1166, LCID, 1, (24, 0), ((12, 17),),DestName
			)

	def PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg, IgnorePrintAreas=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2361, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName, IgnorePrintAreas)

	def PrintPreview(self, EnableChanges=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(281, LCID, 1, (24, 0), ((12, 17),),EnableChanges
			)

	def Protect(self, Password=defaultNamedOptArg, Structure=defaultNamedOptArg, Windows=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2029, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),Password
			, Structure, Windows)

	def ProtectSharing(self, Filename=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg, ReadOnlyRecommended=defaultNamedOptArg
			, CreateBackup=defaultNamedOptArg, SharingPassword=defaultNamedOptArg, FileFormat=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2543, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, SharingPassword
			, FileFormat)

	def PurgeChangeHistoryNow(self, Days=defaultNamedNotOptArg, SharingPassword=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1464, LCID, 1, (24, 0), ((3, 1), (12, 17)),Days
			, SharingPassword)

	def RecheckSmartTags(self):
		return self._oleobj_.InvokeTypes(2065, LCID, 1, (24, 0), (),)

	def RefreshAll(self):
		return self._oleobj_.InvokeTypes(1452, LCID, 1, (24, 0), (),)

	def RejectAllChanges(self, When=defaultNamedOptArg, Who=defaultNamedOptArg, Where=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1467, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),When
			, Who, Where)

	def ReloadAs(self, Encoding=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1821, LCID, 1, (24, 0), ((3, 1),),Encoding
			)

	def RemoveDocumentInformation(self, RemoveDocInfoType=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2514, LCID, 1, (24, 0), ((3, 1),),RemoveDocInfoType
			)

	def RemoveUser(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1453, LCID, 1, (24, 0), ((3, 1),),Index
			)

	def Reply(self):
		return self._oleobj_.InvokeTypes(977, LCID, 1, (24, 0), (),)

	def ReplyAll(self):
		return self._oleobj_.InvokeTypes(978, LCID, 1, (24, 0), (),)

	def ReplyWithChanges(self, ShowMessage=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2057, LCID, 1, (24, 0), ((12, 17),),ShowMessage
			)

	def ResetColors(self):
		return self._oleobj_.InvokeTypes(1468, LCID, 1, (24, 0), (),)

	def Route(self):
		return self._oleobj_.InvokeTypes(946, LCID, 1, (24, 0), (),)

	def RunAutoMacros(self, Which=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(634, LCID, 1, (24, 0), ((3, 1),),Which
			)

	def Save(self):
		return self._oleobj_.InvokeTypes(283, LCID, 1, (24, 0), (),)

	def SaveAs(self, Filename=defaultNamedNotOptArg, FileFormat=defaultNamedNotOptArg, Password=defaultNamedNotOptArg, WriteResPassword=defaultNamedNotOptArg
			, ReadOnlyRecommended=defaultNamedNotOptArg, CreateBackup=defaultNamedNotOptArg, AccessMode=1, ConflictResolution=defaultNamedOptArg, AddToMru=defaultNamedOptArg
			, TextCodepage=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg, Local=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1925, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup
			, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout
			, Local)

	def SaveAsXMLData(self, Filename=defaultNamedNotOptArg, Map=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2278, LCID, 1, (24, 0), ((8, 1), (9, 1)),Filename
			, Map)

	def SaveCopyAs(self, Filename=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(175, LCID, 1, (24, 0), ((12, 17),),Filename
			)

	def SendFaxOverInternet(self, Recipients=defaultNamedOptArg, Subject=defaultNamedOptArg, ShowMessage=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2267, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),Recipients
			, Subject, ShowMessage)

	def SendForReview(self, Recipients=defaultNamedOptArg, Subject=defaultNamedOptArg, ShowMessage=defaultNamedOptArg, IncludeAttachment=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2054, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17)),Recipients
			, Subject, ShowMessage, IncludeAttachment)

	def SendMail(self, Recipients=defaultNamedNotOptArg, Subject=defaultNamedOptArg, ReturnReceipt=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(947, LCID, 1, (24, 0), ((12, 1), (12, 17), (12, 17)),Recipients
			, Subject, ReturnReceipt)

	def SendMailer(self, FileFormat=defaultNamedNotOptArg, Priority=-4143):
		return self._oleobj_.InvokeTypes(980, LCID, 1, (24, 0), ((12, 17), (3, 49)),FileFormat
			, Priority)

	# The method SetColors is actually a property, but must be used as a method to correctly pass the arguments
	def SetColors(self, Index=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(286, LCID, 4, (24, 0), ((12, 17), (12, 1)),Index
			, arg1)

	def SetLinkOnData(self, Name=defaultNamedNotOptArg, Procedure=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(809, LCID, 1, (24, 0), ((8, 1), (12, 17)),Name
			, Procedure)

	def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=defaultNamedOptArg, PasswordEncryptionAlgorithm=defaultNamedOptArg, PasswordEncryptionKeyLength=defaultNamedOptArg, PasswordEncryptionFileProperties=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(2062, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17)),PasswordEncryptionProvider
			, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties)

	def ToggleFormsDesign(self):
		return self._oleobj_.InvokeTypes(2279, LCID, 1, (24, 0), (),)

	def Unprotect(self, Password=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(285, LCID, 1, (24, 0), ((12, 17),),Password
			)

	def UnprotectSharing(self, SharingPassword=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1455, LCID, 1, (24, 0), ((12, 17),),SharingPassword
			)

	def UpdateFromFile(self):
		return self._oleobj_.InvokeTypes(995, LCID, 1, (24, 0), (),)

	def UpdateLink(self, Name=defaultNamedOptArg, Type=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(804, LCID, 1, (24, 0), ((12, 17), (12, 17)),Name
			, Type)

	def WebPagePreview(self):
		return self._oleobj_.InvokeTypes(1818, LCID, 1, (24, 0), (),)

	def XmlImport(self, Url=defaultNamedNotOptArg, ImportMap=pythoncom.Missing, Overwrite=defaultNamedOptArg, Destination=defaultNamedOptArg):
		return self._ApplyTypes_(2270, 1, (3, 0), ((8, 1), (16393, 2), (12, 17), (12, 17)), 'XmlImport', None,Url
			, ImportMap, Overwrite, Destination)

	def XmlImportXml(self, Data=defaultNamedNotOptArg, ImportMap=pythoncom.Missing, Overwrite=defaultNamedOptArg, Destination=defaultNamedOptArg):
		return self._ApplyTypes_(2277, 1, (3, 0), ((8, 1), (16393, 2), (12, 17), (12, 17)), 'XmlImportXml', None,Data
			, ImportMap, Overwrite, Destination)

	def _PrintOut(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg, PrToFileName=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1772, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate, PrToFileName)

	def _Protect(self, Password=defaultNamedOptArg, Structure=defaultNamedOptArg, Windows=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(282, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17)),Password
			, Structure, Windows)

	def _ProtectSharing(self, Filename=defaultNamedOptArg, Password=defaultNamedOptArg, WriteResPassword=defaultNamedOptArg, ReadOnlyRecommended=defaultNamedOptArg
			, CreateBackup=defaultNamedOptArg, SharingPassword=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1450, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, SharingPassword
			)

	def _SaveAs(self, Filename=defaultNamedNotOptArg, FileFormat=defaultNamedNotOptArg, Password=defaultNamedNotOptArg, WriteResPassword=defaultNamedNotOptArg
			, ReadOnlyRecommended=defaultNamedNotOptArg, CreateBackup=defaultNamedNotOptArg, AccessMode=1, ConflictResolution=defaultNamedOptArg, AddToMru=defaultNamedOptArg
			, TextCodepage=defaultNamedOptArg, TextVisualLayout=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(284, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (3, 49), (12, 17), (12, 17), (12, 17), (12, 17)),Filename
			, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup
			, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout
			)

	def _PrintOut_(self, From=defaultNamedOptArg, To=defaultNamedOptArg, Copies=defaultNamedOptArg, Preview=defaultNamedOptArg
			, ActivePrinter=defaultNamedOptArg, PrintToFile=defaultNamedOptArg, Collate=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(905, LCID, 1, (24, 0), ((12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17), (12, 17)),From
			, To, Copies, Preview, ActivePrinter, PrintToFile
			, Collate)

	def sblt(self, s=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1826, LCID, 1, (24, 0), ((8, 1),),s
			)

	_prop_map_get_ = {
		"AcceptLabelsInFormulas": (1441, 2, (11, 0), (), "AcceptLabelsInFormulas", None),
		# Method 'ActiveChart' returns object of type 'Chart'
		"ActiveChart": (183, 2, (13, 0), (), "ActiveChart", '{00020821-0000-0000-C000-000000000046}'),
		"ActiveSheet": (307, 2, (9, 0), (), "ActiveSheet", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"Author": (574, 2, (8, 0), (), "Author", None),
		"AutoUpdateFrequency": (1442, 2, (3, 0), (), "AutoUpdateFrequency", None),
		"AutoUpdateSaveChanges": (1443, 2, (11, 0), (), "AutoUpdateSaveChanges", None),
		"BuiltinDocumentProperties": (1176, 2, (9, 0), (), "BuiltinDocumentProperties", None),
		"CalculationVersion": (1806, 2, (3, 0), (), "CalculationVersion", None),
		"ChangeHistoryDuration": (1444, 2, (3, 0), (), "ChangeHistoryDuration", None),
		# Method 'Charts' returns object of type 'Sheets'
		"Charts": (121, 2, (9, 0), (), "Charts", '{000208D7-0000-0000-C000-000000000046}'),
		"CheckCompatibility": (2528, 2, (11, 0), (), "CheckCompatibility", None),
		"CodeName": (1373, 2, (8, 0), (), "CodeName", None),
		"Colors": (286, 2, (12, 0), ((12, 17),), "Colors", None),
		# Method 'CommandBars' returns object of type 'CommandBars'
		"CommandBars": (1439, 2, (13, 0), (), "CommandBars", '{55F88893-7708-11D1-ACEB-006008961DA5}'),
		"Comments": (575, 2, (8, 0), (), "Comments", None),
		"ConflictResolution": (1175, 2, (3, 0), (), "ConflictResolution", None),
		# Method 'Connections' returns object of type 'Connections'
		"Connections": (2513, 2, (9, 0), (), "Connections", '{00024486-0000-0000-C000-000000000046}'),
		"ConnectionsDisabled": (2536, 2, (11, 0), (), "ConnectionsDisabled", None),
		"Container": (1190, 2, (9, 0), (), "Container", None),
		# Method 'ContentTypeProperties' returns object of type 'MetaProperties'
		"ContentTypeProperties": (2512, 2, (9, 0), (), "ContentTypeProperties", '{000C038E-0000-0000-C000-000000000046}'),
		"CreateBackup": (287, 2, (11, 0), (), "CreateBackup", None),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		"CustomDocumentProperties": (1177, 2, (9, 0), (), "CustomDocumentProperties", None),
		# Method 'CustomViews' returns object of type 'CustomViews'
		"CustomViews": (1456, 2, (9, 0), (), "CustomViews", '{00024422-0000-0000-C000-000000000046}'),
		# Method 'CustomXMLParts' returns object of type 'CustomXMLParts'
		"CustomXMLParts": (2530, 2, (13, 0), (), "CustomXMLParts", '{000CDB0C-0000-0000-C000-000000000046}'),
		"Date1904": (403, 2, (11, 0), (), "Date1904", None),
		"DefaultPivotTableStyle": (2527, 2, (12, 0), (), "DefaultPivotTableStyle", None),
		"DefaultTableStyle": (2526, 2, (12, 0), (), "DefaultTableStyle", None),
		# Method 'DialogSheets' returns object of type 'Sheets'
		"DialogSheets": (764, 2, (9, 0), (), "DialogSheets", '{000208D7-0000-0000-C000-000000000046}'),
		"DisplayDrawingObjects": (404, 2, (3, 0), (), "DisplayDrawingObjects", None),
		"DisplayInkComments": (2276, 2, (11, 0), (), "DisplayInkComments", None),
		"DoNotPromptForConvert": (2541, 2, (11, 0), (), "DoNotPromptForConvert", None),
		# Method 'DocumentInspectors' returns object of type 'DocumentInspectors'
		"DocumentInspectors": (2521, 2, (9, 0), (), "DocumentInspectors", '{000C0392-0000-0000-C000-000000000046}'),
		# Method 'DocumentLibraryVersions' returns object of type 'DocumentLibraryVersions'
		"DocumentLibraryVersions": (2274, 2, (9, 0), (), "DocumentLibraryVersions", '{000C0388-0000-0000-C000-000000000046}'),
		"EnableAutoRecover": (2049, 2, (11, 0), (), "EnableAutoRecover", None),
		"EncryptionProvider": (2540, 2, (8, 0), (), "EncryptionProvider", None),
		"EnvelopeVisible": (1824, 2, (11, 0), (), "EnvelopeVisible", None),
		# Method 'Excel4IntlMacroSheets' returns object of type 'Sheets'
		"Excel4IntlMacroSheets": (581, 2, (9, 0), (), "Excel4IntlMacroSheets", '{000208D7-0000-0000-C000-000000000046}'),
		# Method 'Excel4MacroSheets' returns object of type 'Sheets'
		"Excel4MacroSheets": (579, 2, (9, 0), (), "Excel4MacroSheets", '{000208D7-0000-0000-C000-000000000046}'),
		"Excel8CompatibilityMode": (2535, 2, (11, 0), (), "Excel8CompatibilityMode", None),
		"FileFormat": (288, 2, (3, 0), (), "FileFormat", None),
		"Final": (2531, 2, (11, 0), (), "Final", None),
		"ForceFullCalculation": (2542, 2, (11, 0), (), "ForceFullCalculation", None),
		"FullName": (289, 2, (8, 0), (), "FullName", None),
		"FullNameURLEncoded": (1927, 2, (8, 0), (), "FullNameURLEncoded", None),
		# Method 'HTMLProject' returns object of type 'HTMLProject'
		"HTMLProject": (1823, 2, (9, 0), (), "HTMLProject", '{000C0356-0000-0000-C000-000000000046}'),
		"HasMailer": (976, 2, (11, 0), (), "HasMailer", None),
		"HasPassword": (290, 2, (11, 0), (), "HasPassword", None),
		"HasRoutingSlip": (950, 2, (11, 0), (), "HasRoutingSlip", None),
		"HasVBProject": (2529, 2, (11, 0), (), "HasVBProject", None),
		"HighlightChangesOnScreen": (1461, 2, (11, 0), (), "HighlightChangesOnScreen", None),
		# Method 'IconSets' returns object of type 'IconSets'
		"IconSets": (2539, 2, (9, 0), (), "IconSets", '{0002449C-0000-0000-C000-000000000046}'),
		"InactiveListBorderVisible": (2275, 2, (11, 0), (), "InactiveListBorderVisible", None),
		"IsAddin": (1445, 2, (11, 0), (), "IsAddin", None),
		"IsInplace": (1769, 2, (11, 0), (), "IsInplace", None),
		"KeepChangeHistory": (1462, 2, (11, 0), (), "KeepChangeHistory", None),
		"Keywords": (577, 2, (8, 0), (), "Keywords", None),
		"ListChangesOnNewSheet": (1463, 2, (11, 0), (), "ListChangesOnNewSheet", None),
		# Method 'Mailer' returns object of type 'Mailer'
		"Mailer": (979, 2, (9, 0), (), "Mailer", '{000208D1-0000-0000-C000-000000000046}'),
		# Method 'Modules' returns object of type 'Sheets'
		"Modules": (582, 2, (9, 0), (), "Modules", '{000208D7-0000-0000-C000-000000000046}'),
		"MultiUserEditing": (1169, 2, (11, 0), (), "MultiUserEditing", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		# Method 'Names' returns object of type 'Names'
		"Names": (442, 2, (9, 0), (), "Names", '{000208B8-0000-0000-C000-000000000046}'),
		"OnSave": (1178, 2, (8, 0), (), "OnSave", None),
		"OnSheetActivate": (1031, 2, (8, 0), (), "OnSheetActivate", None),
		"OnSheetDeactivate": (1081, 2, (8, 0), (), "OnSheetDeactivate", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		"Password": (429, 2, (8, 0), (), "Password", None),
		"PasswordEncryptionAlgorithm": (2060, 2, (8, 0), (), "PasswordEncryptionAlgorithm", None),
		"PasswordEncryptionFileProperties": (2063, 2, (11, 0), (), "PasswordEncryptionFileProperties", None),
		"PasswordEncryptionKeyLength": (2061, 2, (3, 0), (), "PasswordEncryptionKeyLength", None),
		"PasswordEncryptionProvider": (2059, 2, (8, 0), (), "PasswordEncryptionProvider", None),
		"Path": (291, 2, (8, 0), (), "Path", None),
		# Method 'Permission' returns object of type 'Permission'
		"Permission": (2264, 2, (9, 0), (), "Permission", '{000C0376-0000-0000-C000-000000000046}'),
		"PersonalViewListSettings": (1447, 2, (11, 0), (), "PersonalViewListSettings", None),
		"PersonalViewPrintSettings": (1448, 2, (11, 0), (), "PersonalViewPrintSettings", None),
		"PrecisionAsDisplayed": (405, 2, (11, 0), (), "PrecisionAsDisplayed", None),
		"ProtectStructure": (588, 2, (11, 0), (), "ProtectStructure", None),
		"ProtectWindows": (295, 2, (11, 0), (), "ProtectWindows", None),
		# Method 'PublishObjects' returns object of type 'PublishObjects'
		"PublishObjects": (1819, 2, (9, 0), (), "PublishObjects", '{00024443-0000-0000-C000-000000000046}'),
		"ReadOnly": (296, 2, (11, 0), (), "ReadOnly", None),
		"ReadOnlyRecommended": (2005, 2, (11, 0), (), "ReadOnlyRecommended", None),
		"RemovePersonalInformation": (2050, 2, (11, 0), (), "RemovePersonalInformation", None),
		# Method 'Research' returns object of type 'Research'
		"Research": (2532, 2, (9, 0), (), "Research", '{000244AC-0000-0000-C000-000000000046}'),
		"RevisionNumber": (1172, 2, (3, 0), (), "RevisionNumber", None),
		"Routed": (951, 2, (11, 0), (), "Routed", None),
		# Method 'RoutingSlip' returns object of type 'RoutingSlip'
		"RoutingSlip": (949, 2, (9, 0), (), "RoutingSlip", '{000208AA-0000-0000-C000-000000000046}'),
		"SaveLinkValues": (406, 2, (11, 0), (), "SaveLinkValues", None),
		"Saved": (298, 2, (11, 0), (), "Saved", None),
		# Method 'ServerPolicy' returns object of type 'ServerPolicy'
		"ServerPolicy": (2519, 2, (9, 0), (), "ServerPolicy", '{000C0390-0000-0000-C000-000000000046}'),
		# Method 'ServerViewableItems' returns object of type 'ServerViewableItems'
		"ServerViewableItems": (2524, 2, (9, 0), (), "ServerViewableItems", '{000244A4-0000-0000-C000-000000000046}'),
		# Method 'SharedWorkspace' returns object of type 'SharedWorkspace'
		"SharedWorkspace": (2265, 2, (9, 0), (), "SharedWorkspace", '{000C0385-0000-0000-C000-000000000046}'),
		# Method 'Sheets' returns object of type 'Sheets'
		"Sheets": (485, 2, (9, 0), (), "Sheets", '{000208D7-0000-0000-C000-000000000046}'),
		"ShowConflictHistory": (1171, 2, (11, 0), (), "ShowConflictHistory", None),
		"ShowPivotChartActiveFields": (2538, 2, (11, 0), (), "ShowPivotChartActiveFields", None),
		"ShowPivotTableFieldList": (2046, 2, (11, 0), (), "ShowPivotTableFieldList", None),
		# Method 'Signatures' returns object of type 'SignatureSet'
		"Signatures": (2516, 2, (9, 0), (), "Signatures", '{000C0410-0000-0000-C000-000000000046}'),
		# Method 'SmartDocument' returns object of type 'SmartDocument'
		"SmartDocument": (2273, 2, (9, 0), (), "SmartDocument", '{000C0377-0000-0000-C000-000000000046}'),
		# Method 'SmartTagOptions' returns object of type 'SmartTagOptions'
		"SmartTagOptions": (2064, 2, (9, 0), (), "SmartTagOptions", '{00024464-0000-0000-C000-000000000046}'),
		# Method 'Styles' returns object of type 'Styles'
		"Styles": (493, 2, (9, 0), (), "Styles", '{00020853-0000-0000-C000-000000000046}'),
		"Subject": (953, 2, (8, 0), (), "Subject", None),
		# Method 'Sync' returns object of type 'Sync'
		"Sync": (2266, 2, (9, 0), (), "Sync", '{000C0386-0000-0000-C000-000000000046}'),
		# Method 'TableStyles' returns object of type 'TableStyles'
		"TableStyles": (2525, 2, (9, 0), (), "TableStyles", '{000244A8-0000-0000-C000-000000000046}'),
		"TemplateRemoveExtData": (1457, 2, (11, 0), (), "TemplateRemoveExtData", None),
		# Method 'Theme' returns object of type 'OfficeTheme'
		"Theme": (2533, 2, (9, 0), (), "Theme", '{000C03A0-0000-0000-C000-000000000046}'),
		"Title": (199, 2, (8, 0), (), "Title", None),
		"UpdateLinks": (864, 2, (3, 0), (), "UpdateLinks", None),
		"UpdateRemoteReferences": (411, 2, (11, 0), (), "UpdateRemoteReferences", None),
		"UserControl": (1210, 2, (11, 0), (), "UserControl", None),
		"UserStatus": (1173, 2, (12, 0), (), "UserStatus", None),
		"VBASigned": (1828, 2, (11, 0), (), "VBASigned", None),
		# Method 'VBProject' returns object of type 'VBProject'
		"VBProject": (1469, 2, (13, 0), (), "VBProject", '{0002E169-0000-0000-C000-000000000046}'),
		# Method 'WebOptions' returns object of type 'WebOptions'
		"WebOptions": (1820, 2, (9, 0), (), "WebOptions", '{00024449-0000-0000-C000-000000000046}'),
		# Method 'Windows' returns object of type 'Windows'
		"Windows": (430, 2, (9, 0), (), "Windows", '{00020892-0000-0000-C000-000000000046}'),
		# Method 'Worksheets' returns object of type 'Sheets'
		"Worksheets": (494, 2, (9, 0), (), "Worksheets", '{000208D7-0000-0000-C000-000000000046}'),
		"WritePassword": (1128, 2, (8, 0), (), "WritePassword", None),
		"WriteReserved": (299, 2, (11, 0), (), "WriteReserved", None),
		"WriteReservedBy": (300, 2, (8, 0), (), "WriteReservedBy", None),
		# Method 'XmlMaps' returns object of type 'XmlMaps'
		"XmlMaps": (2269, 2, (9, 0), (), "XmlMaps", '{0002447C-0000-0000-C000-000000000046}'),
		# Method 'XmlNamespaces' returns object of type 'XmlNamespaces'
		"XmlNamespaces": (2268, 2, (9, 0), (), "XmlNamespaces", '{00024477-0000-0000-C000-000000000046}'),
		"_CodeName": (-2147418112, 2, (8, 0), (), "_CodeName", None),
		"_ReadOnlyRecommended": (297, 2, (11, 0), (), "_ReadOnlyRecommended", None),
	}
	_prop_map_put_ = {
		"AcceptLabelsInFormulas": ((1441, LCID, 4, 0),()),
		"Author": ((574, LCID, 4, 0),()),
		"AutoUpdateFrequency": ((1442, LCID, 4, 0),()),
		"AutoUpdateSaveChanges": ((1443, LCID, 4, 0),()),
		"ChangeHistoryDuration": ((1444, LCID, 4, 0),()),
		"CheckCompatibility": ((2528, LCID, 4, 0),()),
		"Colors": ((286, LCID, 4, 0),()),
		"Comments": ((575, LCID, 4, 0),()),
		"ConflictResolution": ((1175, LCID, 4, 0),()),
		"Date1904": ((403, LCID, 4, 0),()),
		"DefaultPivotTableStyle": ((2527, LCID, 4, 0),()),
		"DefaultTableStyle": ((2526, LCID, 4, 0),()),
		"DisplayDrawingObjects": ((404, LCID, 4, 0),()),
		"DisplayInkComments": ((2276, LCID, 4, 0),()),
		"DoNotPromptForConvert": ((2541, LCID, 4, 0),()),
		"EnableAutoRecover": ((2049, LCID, 4, 0),()),
		"EncryptionProvider": ((2540, LCID, 4, 0),()),
		"EnvelopeVisible": ((1824, LCID, 4, 0),()),
		"Final": ((2531, LCID, 4, 0),()),
		"ForceFullCalculation": ((2542, LCID, 4, 0),()),
		"HasMailer": ((976, LCID, 4, 0),()),
		"HasRoutingSlip": ((950, LCID, 4, 0),()),
		"HighlightChangesOnScreen": ((1461, LCID, 4, 0),()),
		"InactiveListBorderVisible": ((2275, LCID, 4, 0),()),
		"IsAddin": ((1445, LCID, 4, 0),()),
		"KeepChangeHistory": ((1462, LCID, 4, 0),()),
		"Keywords": ((577, LCID, 4, 0),()),
		"ListChangesOnNewSheet": ((1463, LCID, 4, 0),()),
		"OnSave": ((1178, LCID, 4, 0),()),
		"OnSheetActivate": ((1031, LCID, 4, 0),()),
		"OnSheetDeactivate": ((1081, LCID, 4, 0),()),
		"Password": ((429, LCID, 4, 0),()),
		"PersonalViewListSettings": ((1447, LCID, 4, 0),()),
		"PersonalViewPrintSettings": ((1448, LCID, 4, 0),()),
		"PrecisionAsDisplayed": ((405, LCID, 4, 0),()),
		"ReadOnlyRecommended": ((2005, LCID, 4, 0),()),
		"RemovePersonalInformation": ((2050, LCID, 4, 0),()),
		"SaveLinkValues": ((406, LCID, 4, 0),()),
		"Saved": ((298, LCID, 4, 0),()),
		"ShowConflictHistory": ((1171, LCID, 4, 0),()),
		"ShowPivotChartActiveFields": ((2538, LCID, 4, 0),()),
		"ShowPivotTableFieldList": ((2046, LCID, 4, 0),()),
		"Subject": ((953, LCID, 4, 0),()),
		"TemplateRemoveExtData": ((1457, LCID, 4, 0),()),
		"Title": ((199, LCID, 4, 0),()),
		"UpdateLinks": ((864, LCID, 4, 0),()),
		"UpdateRemoteReferences": ((411, LCID, 4, 0),()),
		"UserControl": ((1210, LCID, 4, 0),()),
		"WritePassword": ((1128, LCID, 4, 0),()),
		"_CodeName": ((-2147418112, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000208DA-0000-0000-C000-000000000046}", _Workbook )
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

_Workbook_vtables_dispatch_ = 1
_Workbook_vtables_ = [
	(( 'Application' , 'RHS' , ), 148, (148, (), [ (16397, 10, None, "IID('{00024500-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'RHS' , ), 149, (149, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'RHS' , ), 150, (150, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( 'AcceptLabelsInFormulas' , 'RHS' , ), 1441, (1441, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 64 , )),
	(( 'AcceptLabelsInFormulas' , 'RHS' , ), 1441, (1441, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 44 , (3, 0, None, None) , 64 , )),
	(( 'Activate' , 'lcid' , ), 304, (304, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( 'ActiveChart' , 'RHS' , ), 183, (183, (), [ (16397, 10, None, "IID('{00020821-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( 'ActiveSheet' , 'RHS' , ), 307, (307, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Author' , 'lcid' , 'RHS' , ), 574, (574, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 64 , )),
	(( 'Author' , 'lcid' , 'RHS' , ), 574, (574, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 64 , )),
	(( 'AutoUpdateFrequency' , 'RHS' , ), 1442, (1442, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( 'AutoUpdateFrequency' , 'RHS' , ), 1442, (1442, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'AutoUpdateSaveChanges' , 'RHS' , ), 1443, (1443, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( 'AutoUpdateSaveChanges' , 'RHS' , ), 1443, (1443, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'ChangeHistoryDuration' , 'RHS' , ), 1444, (1444, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( 'ChangeHistoryDuration' , 'RHS' , ), 1444, (1444, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'BuiltinDocumentProperties' , 'RHS' , ), 1176, (1176, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( 'ChangeFileAccess' , 'Mode' , 'WritePassword' , 'Notify' , 'lcid' , 
			 ), 989, (989, (), [ (3, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 96 , (3, 0, None, None) , 0 , )),
	(( 'ChangeLink' , 'Name' , 'NewName' , 'Type' , 'lcid' , 
			 ), 802, (802, (), [ (8, 1, None, None) , (8, 1, None, None) , (3, 49, '1', None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( 'Charts' , 'RHS' , ), 121, (121, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Close' , 'SaveChanges' , 'Filename' , 'RouteWorkbook' , 'lcid' , 
			 ), 277, (277, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 3 , 108 , (3, 0, None, None) , 0 , )),
	(( 'CodeName' , 'RHS' , ), 1373, (1373, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( '_CodeName' , 'RHS' , ), -2147418112, (-2147418112, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 1024 , )),
	(( '_CodeName' , 'RHS' , ), -2147418112, (-2147418112, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 1024 , )),
	(( 'Colors' , 'Index' , 'lcid' , 'RHS' , ), 286, (286, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 2 , 4 , 1 , 124 , (3, 0, None, None) , 0 , )),
	(( 'Colors' , 'Index' , 'lcid' , 'RHS' , ), 286, (286, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (12, 1, None, None) , ], 1 , 4 , 4 , 1 , 128 , (3, 0, None, None) , 0 , )),
	(( 'CommandBars' , 'RHS' , ), 1439, (1439, (), [ (16397, 10, None, "IID('{55F88893-7708-11D1-ACEB-006008961DA5}')") , ], 1 , 2 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( 'Comments' , 'lcid' , 'RHS' , ), 575, (575, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 64 , )),
	(( 'Comments' , 'lcid' , 'RHS' , ), 575, (575, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 140 , (3, 0, None, None) , 64 , )),
	(( 'ConflictResolution' , 'RHS' , ), 1175, (1175, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'ConflictResolution' , 'RHS' , ), 1175, (1175, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( 'Container' , 'RHS' , ), 1190, (1190, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'CreateBackup' , 'lcid' , 'RHS' , ), 287, (287, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( 'CustomDocumentProperties' , 'RHS' , ), 1177, (1177, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Date1904' , 'lcid' , 'RHS' , ), 403, (403, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( 'Date1904' , 'lcid' , 'RHS' , ), 403, (403, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'DeleteNumberFormat' , 'NumberFormat' , 'lcid' , ), 397, (397, (), [ (8, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 172 , (3, 0, None, None) , 0 , )),
	(( 'DialogSheets' , 'RHS' , ), 764, (764, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 64 , )),
	(( 'DisplayDrawingObjects' , 'lcid' , 'RHS' , ), 404, (404, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( 'DisplayDrawingObjects' , 'lcid' , 'RHS' , ), 404, (404, (), [ (3, 5, None, None) , 
			 (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'ExclusiveAccess' , 'lcid' , 'RHS' , ), 1168, (1168, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( 'FileFormat' , 'lcid' , 'RHS' , ), 288, (288, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'ForwardMailer' , 'lcid' , ), 973, (973, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( 'FullName' , 'lcid' , 'RHS' , ), 289, (289, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'HasMailer' , 'lcid' , 'RHS' , ), 976, (976, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 204 , (3, 0, None, None) , 64 , )),
	(( 'HasMailer' , 'lcid' , 'RHS' , ), 976, (976, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 208 , (3, 0, None, None) , 64 , )),
	(( 'HasPassword' , 'lcid' , 'RHS' , ), 290, (290, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( 'HasRoutingSlip' , 'lcid' , 'RHS' , ), 950, (950, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 64 , )),
	(( 'HasRoutingSlip' , 'lcid' , 'RHS' , ), 950, (950, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 220 , (3, 0, None, None) , 64 , )),
	(( 'IsAddin' , 'RHS' , ), 1445, (1445, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'IsAddin' , 'RHS' , ), 1445, (1445, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( 'Keywords' , 'lcid' , 'RHS' , ), 577, (577, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 64 , )),
	(( 'Keywords' , 'lcid' , 'RHS' , ), 577, (577, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 236 , (3, 0, None, None) , 64 , )),
	(( 'LinkInfo' , 'Name' , 'LinkInfo' , 'Type' , 'EditionRef' , 
			 'lcid' , 'RHS' , ), 807, (807, (), [ (8, 1, None, None) , (3, 1, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 2 , 240 , (3, 0, None, None) , 0 , )),
	(( 'LinkSources' , 'Type' , 'lcid' , 'RHS' , ), 808, (808, (), [ 
			 (12, 17, None, None) , (3, 5, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 1 , 244 , (3, 0, None, None) , 0 , )),
	(( 'Mailer' , 'RHS' , ), 979, (979, (), [ (16393, 10, None, "IID('{000208D1-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'MergeWorkbook' , 'Filename' , ), 1446, (1446, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 252 , (3, 0, None, None) , 0 , )),
	(( 'Modules' , 'RHS' , ), 582, (582, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 64 , )),
	(( 'MultiUserEditing' , 'lcid' , 'RHS' , ), 1169, (1169, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'RHS' , ), 110, (110, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'Names' , 'RHS' , ), 442, (442, (), [ (16393, 10, None, "IID('{000208B8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 268 , (3, 0, None, None) , 0 , )),
	(( 'NewWindow' , 'lcid' , 'RHS' , ), 280, (280, (), [ (3, 5, None, None) , 
			 (16393, 10, None, "IID('{00020893-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'OnSave' , 'lcid' , 'RHS' , ), 1178, (1178, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 276 , (3, 0, None, None) , 64 , )),
	(( 'OnSave' , 'lcid' , 'RHS' , ), 1178, (1178, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 284 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetActivate' , 'lcid' , 'RHS' , ), 1031, (1031, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 288 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 292 , (3, 0, None, None) , 64 , )),
	(( 'OnSheetDeactivate' , 'lcid' , 'RHS' , ), 1081, (1081, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 64 , )),
	(( 'OpenLinks' , 'Name' , 'ReadOnly' , 'Type' , 'lcid' , 
			 ), 803, (803, (), [ (8, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 300 , (3, 0, None, None) , 0 , )),
	(( 'Path' , 'lcid' , 'RHS' , ), 291, (291, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'PersonalViewListSettings' , 'RHS' , ), 1447, (1447, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( 'PersonalViewListSettings' , 'RHS' , ), 1447, (1447, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'PersonalViewPrintSettings' , 'RHS' , ), 1448, (1448, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( 'PersonalViewPrintSettings' , 'RHS' , ), 1448, (1448, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'PivotCaches' , 'RHS' , ), 1449, (1449, (), [ (16393, 10, None, "IID('{0002441D-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( 'Post' , 'DestName' , 'lcid' , ), 1166, (1166, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 328 , (3, 0, None, None) , 0 , )),
	(( 'PrecisionAsDisplayed' , 'lcid' , 'RHS' , ), 405, (405, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 332 , (3, 0, None, None) , 0 , )),
	(( 'PrecisionAsDisplayed' , 'lcid' , 'RHS' , ), 405, (405, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( '__PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'lcid' , ), 905, (905, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 7 , 340 , (3, 0, None, None) , 1088 , )),
	(( 'PrintPreview' , 'EnableChanges' , 'lcid' , ), 281, (281, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 344 , (3, 0, None, None) , 0 , )),
	(( '_Protect' , 'Password' , 'Structure' , 'Windows' , ), 282, (282, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 348 , (3, 0, None, None) , 1088 , )),
	(( '_ProtectSharing' , 'Filename' , 'Password' , 'WriteResPassword' , 'ReadOnlyRecommended' , 
			 'CreateBackup' , 'SharingPassword' , ), 1450, (1450, (), [ (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 6 , 352 , (3, 0, None, None) , 1088 , )),
	(( 'ProtectStructure' , 'RHS' , ), 588, (588, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 356 , (3, 0, None, None) , 0 , )),
	(( 'ProtectWindows' , 'RHS' , ), 295, (295, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'ReadOnly' , 'lcid' , 'RHS' , ), 296, (296, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 364 , (3, 0, None, None) , 0 , )),
	(( '_ReadOnlyRecommended' , 'lcid' , 'RHS' , ), 297, (297, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 368 , (3, 0, None, None) , 1088 , )),
	(( 'RefreshAll' , ), 1452, (1452, (), [ ], 1 , 1 , 4 , 0 , 372 , (3, 0, None, None) , 0 , )),
	(( 'Reply' , 'lcid' , ), 977, (977, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'ReplyAll' , 'lcid' , ), 978, (978, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 380 , (3, 0, None, None) , 0 , )),
	(( 'RemoveUser' , 'Index' , ), 1453, (1453, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'RevisionNumber' , 'lcid' , 'RHS' , ), 1172, (1172, (), [ (3, 5, None, None) , 
			 (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 388 , (3, 0, None, None) , 0 , )),
	(( 'Route' , 'lcid' , ), 946, (946, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 392 , (3, 0, None, None) , 64 , )),
	(( 'Routed' , 'lcid' , 'RHS' , ), 951, (951, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 396 , (3, 0, None, None) , 64 , )),
	(( 'RoutingSlip' , 'RHS' , ), 949, (949, (), [ (16393, 10, None, "IID('{000208AA-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 64 , )),
	(( 'RunAutoMacros' , 'Which' , 'lcid' , ), 634, (634, (), [ (3, 1, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 404 , (3, 0, None, None) , 0 , )),
	(( 'Save' , 'lcid' , ), 283, (283, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( '_SaveAs' , 'Filename' , 'FileFormat' , 'Password' , 'WriteResPassword' , 
			 'ReadOnlyRecommended' , 'CreateBackup' , 'AccessMode' , 'ConflictResolution' , 'AddToMru' , 
			 'TextCodepage' , 'TextVisualLayout' , 'lcid' , ), 284, (284, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 49, '1', None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 4 , 412 , (3, 0, None, None) , 1088 , )),
	(( 'SaveCopyAs' , 'Filename' , 'lcid' , ), 175, (175, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 416 , (3, 0, None, None) , 0 , )),
	(( 'Saved' , 'lcid' , 'RHS' , ), 298, (298, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 420 , (3, 0, None, None) , 0 , )),
	(( 'Saved' , 'lcid' , 'RHS' , ), 298, (298, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'SaveLinkValues' , 'lcid' , 'RHS' , ), 406, (406, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 428 , (3, 0, None, None) , 0 , )),
	(( 'SaveLinkValues' , 'lcid' , 'RHS' , ), 406, (406, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'SendMail' , 'Recipients' , 'Subject' , 'ReturnReceipt' , 'lcid' , 
			 ), 947, (947, (), [ (12, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 436 , (3, 0, None, None) , 0 , )),
	(( 'SendMailer' , 'FileFormat' , 'Priority' , 'lcid' , ), 980, (980, (), [ 
			 (12, 17, None, None) , (3, 49, '-4143', None) , (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'SetLinkOnData' , 'Name' , 'Procedure' , 'lcid' , ), 809, (809, (), [ 
			 (8, 1, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 444 , (3, 0, None, None) , 0 , )),
	(( 'Sheets' , 'RHS' , ), 485, (485, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'ShowConflictHistory' , 'lcid' , 'RHS' , ), 1171, (1171, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 452 , (3, 0, None, None) , 0 , )),
	(( 'ShowConflictHistory' , 'lcid' , 'RHS' , ), 1171, (1171, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'Styles' , 'RHS' , ), 493, (493, (), [ (16393, 10, None, "IID('{00020853-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 460 , (3, 0, None, None) , 0 , )),
	(( 'Subject' , 'lcid' , 'RHS' , ), 953, (953, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 64 , )),
	(( 'Subject' , 'lcid' , 'RHS' , ), 953, (953, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 468 , (3, 0, None, None) , 64 , )),
	(( 'Title' , 'lcid' , 'RHS' , ), 199, (199, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 64 , )),
	(( 'Title' , 'lcid' , 'RHS' , ), 199, (199, (), [ (3, 5, None, None) , 
			 (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 476 , (3, 0, None, None) , 64 , )),
	(( 'Unprotect' , 'Password' , 'lcid' , ), 285, (285, (), [ (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 1 , 480 , (3, 0, None, None) , 0 , )),
	(( 'UnprotectSharing' , 'SharingPassword' , ), 1455, (1455, (), [ (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 484 , (3, 0, None, None) , 0 , )),
	(( 'UpdateFromFile' , 'lcid' , ), 995, (995, (), [ (3, 5, None, None) , ], 1 , 1 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'UpdateLink' , 'Name' , 'Type' , 'lcid' , ), 804, (804, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 2 , 492 , (3, 0, None, None) , 0 , )),
	(( 'UpdateRemoteReferences' , 'lcid' , 'RHS' , ), 411, (411, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'UpdateRemoteReferences' , 'lcid' , 'RHS' , ), 411, (411, (), [ (3, 5, None, None) , 
			 (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 500 , (3, 0, None, None) , 0 , )),
	(( 'UserControl' , 'RHS' , ), 1210, (1210, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 504 , (3, 0, None, None) , 64 , )),
	(( 'UserControl' , 'RHS' , ), 1210, (1210, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 508 , (3, 0, None, None) , 64 , )),
	(( 'UserStatus' , 'lcid' , 'RHS' , ), 1173, (1173, (), [ (3, 5, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'CustomViews' , 'RHS' , ), 1456, (1456, (), [ (16393, 10, None, "IID('{00024422-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 516 , (3, 0, None, None) , 0 , )),
	(( 'Windows' , 'RHS' , ), 430, (430, (), [ (16393, 10, None, "IID('{00020892-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'Worksheets' , 'RHS' , ), 494, (494, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 524 , (3, 0, None, None) , 0 , )),
	(( 'WriteReserved' , 'lcid' , 'RHS' , ), 299, (299, (), [ (3, 5, None, None) , 
			 (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'WriteReservedBy' , 'lcid' , 'RHS' , ), 300, (300, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 532 , (3, 0, None, None) , 0 , )),
	(( 'Excel4IntlMacroSheets' , 'RHS' , ), 581, (581, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'Excel4MacroSheets' , 'RHS' , ), 579, (579, (), [ (16393, 10, None, "IID('{000208D7-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 540 , (3, 0, None, None) , 0 , )),
	(( 'TemplateRemoveExtData' , 'RHS' , ), 1457, (1457, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'TemplateRemoveExtData' , 'RHS' , ), 1457, (1457, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 548 , (3, 0, None, None) , 0 , )),
	(( 'HighlightChangesOptions' , 'When' , 'Who' , 'Where' , ), 1458, (1458, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 552 , (3, 0, None, None) , 0 , )),
	(( 'HighlightChangesOnScreen' , 'RHS' , ), 1461, (1461, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 556 , (3, 0, None, None) , 0 , )),
	(( 'HighlightChangesOnScreen' , 'RHS' , ), 1461, (1461, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'KeepChangeHistory' , 'RHS' , ), 1462, (1462, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 564 , (3, 0, None, None) , 0 , )),
	(( 'KeepChangeHistory' , 'RHS' , ), 1462, (1462, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'ListChangesOnNewSheet' , 'RHS' , ), 1463, (1463, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 572 , (3, 0, None, None) , 0 , )),
	(( 'ListChangesOnNewSheet' , 'RHS' , ), 1463, (1463, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'PurgeChangeHistoryNow' , 'Days' , 'SharingPassword' , ), 1464, (1464, (), [ (3, 1, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 580 , (3, 0, None, None) , 0 , )),
	(( 'AcceptAllChanges' , 'When' , 'Who' , 'Where' , ), 1466, (1466, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 584 , (3, 0, None, None) , 0 , )),
	(( 'RejectAllChanges' , 'When' , 'Who' , 'Where' , ), 1467, (1467, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 588 , (3, 0, None, None) , 0 , )),
	(( 'PivotTableWizard' , 'SourceType' , 'SourceData' , 'TableDestination' , 'TableName' , 
			 'RowGrand' , 'ColumnGrand' , 'SaveData' , 'HasAutoFormat' , 'AutoPage' , 
			 'Reserved' , 'BackgroundQuery' , 'OptimizeCache' , 'PageFieldOrder' , 'PageFieldWrapCount' , 
			 'ReadData' , 'Connection' , 'lcid' , ), 684, (684, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (3, 5, None, None) , ], 1 , 1 , 4 , 16 , 592 , (3, 0, None, None) , 64 , )),
	(( 'ResetColors' , ), 1468, (1468, (), [ ], 1 , 1 , 4 , 0 , 596 , (3, 0, None, None) , 0 , )),
	(( 'VBProject' , 'RHS' , ), 1469, (1469, (), [ (16397, 10, None, "IID('{0002E169-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'FollowHyperlink' , 'Address' , 'SubAddress' , 'NewWindow' , 'AddHistory' , 
			 'ExtraInfo' , 'Method' , 'HeaderInfo' , ), 1470, (1470, (), [ (8, 1, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 6 , 604 , (3, 0, None, None) , 0 , )),
	(( 'AddToFavorites' , ), 1476, (1476, (), [ ], 1 , 1 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'IsInplace' , 'RHS' , ), 1769, (1769, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 612 , (3, 0, None, None) , 0 , )),
	(( '_PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'lcid' , 
			 ), 1772, (1772, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 8 , 616 , (3, 0, None, None) , 1088 , )),
	(( 'WebPagePreview' , ), 1818, (1818, (), [ ], 1 , 1 , 4 , 0 , 620 , (3, 0, None, None) , 0 , )),
	(( 'PublishObjects' , 'RHS' , ), 1819, (1819, (), [ (16393, 10, None, "IID('{00024443-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'WebOptions' , 'RHS' , ), 1820, (1820, (), [ (16393, 10, None, "IID('{00024449-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 628 , (3, 0, None, None) , 0 , )),
	(( 'ReloadAs' , 'Encoding' , ), 1821, (1821, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'HTMLProject' , 'RHS' , ), 1823, (1823, (), [ (16393, 10, None, "IID('{000C0356-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 636 , (3, 0, None, None) , 64 , )),
	(( 'EnvelopeVisible' , 'RHS' , ), 1824, (1824, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'EnvelopeVisible' , 'RHS' , ), 1824, (1824, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 644 , (3, 0, None, None) , 0 , )),
	(( 'CalculationVersion' , 'RHS' , ), 1806, (1806, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'Dummy17' , 'calcid' , ), 2044, (2044, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 652 , (3, 0, None, None) , 64 , )),
	(( 'sblt' , 's' , ), 1826, (1826, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 656 , (3, 0, None, None) , 64 , )),
	(( 'VBASigned' , 'RHS' , ), 1828, (1828, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 660 , (3, 0, None, None) , 0 , )),
	(( 'ShowPivotTableFieldList' , 'RHS' , ), 2046, (2046, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'ShowPivotTableFieldList' , 'RHS' , ), 2046, (2046, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 668 , (3, 0, None, None) , 0 , )),
	(( 'UpdateLinks' , 'RHS' , ), 864, (864, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'UpdateLinks' , 'RHS' , ), 864, (864, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 676 , (3, 0, None, None) , 0 , )),
	(( 'BreakLink' , 'Name' , 'Type' , ), 2047, (2047, (), [ (8, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'Dummy16' , ), 2048, (2048, (), [ ], 1 , 1 , 4 , 0 , 684 , (3, 0, None, None) , 64 , )),
	(( 'SaveAs' , 'Filename' , 'FileFormat' , 'Password' , 'WriteResPassword' , 
			 'ReadOnlyRecommended' , 'CreateBackup' , 'AccessMode' , 'ConflictResolution' , 'AddToMru' , 
			 'TextCodepage' , 'TextVisualLayout' , 'Local' , 'lcid' , ), 1925, (1925, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 49, '1', None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 5 , 688 , (3, 0, None, None) , 0 , )),
	(( 'EnableAutoRecover' , 'RHS' , ), 2049, (2049, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 692 , (3, 0, None, None) , 0 , )),
	(( 'EnableAutoRecover' , 'RHS' , ), 2049, (2049, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'RemovePersonalInformation' , 'RHS' , ), 2050, (2050, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 700 , (3, 0, None, None) , 0 , )),
	(( 'RemovePersonalInformation' , 'RHS' , ), 2050, (2050, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'FullNameURLEncoded' , 'lcid' , 'RHS' , ), 1927, (1927, (), [ (3, 5, None, None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 708 , (3, 0, None, None) , 0 , )),
	(( 'CheckIn' , 'SaveChanges' , 'Comments' , 'MakePublic' , ), 2051, (2051, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 712 , (3, 0, None, None) , 0 , )),
	(( 'CanCheckIn' , 'RHS' , ), 2053, (2053, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 716 , (3, 0, None, None) , 0 , )),
	(( 'SendForReview' , 'Recipients' , 'Subject' , 'ShowMessage' , 'IncludeAttachment' , 
			 ), 2054, (2054, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 4 , 720 , (3, 0, None, None) , 0 , )),
	(( 'ReplyWithChanges' , 'ShowMessage' , ), 2057, (2057, (), [ (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 724 , (3, 0, None, None) , 0 , )),
	(( 'EndReview' , ), 2058, (2058, (), [ ], 1 , 1 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'Password' , 'RHS' , ), 429, (429, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 732 , (3, 0, None, None) , 0 , )),
	(( 'Password' , 'RHS' , ), 429, (429, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'WritePassword' , 'RHS' , ), 1128, (1128, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 740 , (3, 0, None, None) , 0 , )),
	(( 'WritePassword' , 'RHS' , ), 1128, (1128, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionProvider' , 'RHS' , ), 2059, (2059, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 748 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionAlgorithm' , 'RHS' , ), 2060, (2060, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionKeyLength' , 'RHS' , ), 2061, (2061, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 756 , (3, 0, None, None) , 0 , )),
	(( 'SetPasswordEncryptionOptions' , 'PasswordEncryptionProvider' , 'PasswordEncryptionAlgorithm' , 'PasswordEncryptionKeyLength' , 'PasswordEncryptionFileProperties' , 
			 ), 2062, (2062, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 4 , 760 , (3, 0, None, None) , 0 , )),
	(( 'PasswordEncryptionFileProperties' , 'RHS' , ), 2063, (2063, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 764 , (3, 0, None, None) , 0 , )),
	(( 'ReadOnlyRecommended' , 'RHS' , ), 2005, (2005, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'ReadOnlyRecommended' , 'RHS' , ), 2005, (2005, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 772 , (3, 0, None, None) , 0 , )),
	(( 'Protect' , 'Password' , 'Structure' , 'Windows' , ), 2029, (2029, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 776 , (3, 0, None, None) , 0 , )),
	(( 'SmartTagOptions' , 'RHS' , ), 2064, (2064, (), [ (16393, 10, None, "IID('{00024464-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 780 , (3, 0, None, None) , 0 , )),
	(( 'RecheckSmartTags' , ), 2065, (2065, (), [ ], 1 , 1 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'Permission' , 'RHS' , ), 2264, (2264, (), [ (16393, 10, None, "IID('{000C0376-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 788 , (3, 0, None, None) , 0 , )),
	(( 'SharedWorkspace' , 'RHS' , ), 2265, (2265, (), [ (16393, 10, None, "IID('{000C0385-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'Sync' , 'RHS' , ), 2266, (2266, (), [ (16393, 10, None, "IID('{000C0386-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 796 , (3, 0, None, None) , 0 , )),
	(( 'SendFaxOverInternet' , 'Recipients' , 'Subject' , 'ShowMessage' , ), 2267, (2267, (), [ 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 3 , 800 , (3, 0, None, None) , 0 , )),
	(( 'XmlNamespaces' , 'RHS' , ), 2268, (2268, (), [ (16393, 10, None, "IID('{00024477-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 804 , (3, 0, None, None) , 0 , )),
	(( 'XmlMaps' , 'RHS' , ), 2269, (2269, (), [ (16393, 10, None, "IID('{0002447C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'XmlImport' , 'Url' , 'ImportMap' , 'Overwrite' , 'Destination' , 
			 'RHS' , ), 2270, (2270, (), [ (8, 1, None, None) , (16393, 2, None, "IID('{0002447B-0000-0000-C000-000000000046}')") , (12, 17, None, None) , 
			 (12, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 812 , (3, 0, None, None) , 0 , )),
	(( 'SmartDocument' , 'RHS' , ), 2273, (2273, (), [ (16393, 10, None, "IID('{000C0377-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'DocumentLibraryVersions' , 'RHS' , ), 2274, (2274, (), [ (16393, 10, None, "IID('{000C0388-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 820 , (3, 0, None, None) , 0 , )),
	(( 'InactiveListBorderVisible' , 'RHS' , ), 2275, (2275, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 824 , (3, 0, None, None) , 0 , )),
	(( 'InactiveListBorderVisible' , 'RHS' , ), 2275, (2275, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 828 , (3, 0, None, None) , 0 , )),
	(( 'DisplayInkComments' , 'RHS' , ), 2276, (2276, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'DisplayInkComments' , 'RHS' , ), 2276, (2276, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 836 , (3, 0, None, None) , 0 , )),
	(( 'XmlImportXml' , 'Data' , 'ImportMap' , 'Overwrite' , 'Destination' , 
			 'RHS' , ), 2277, (2277, (), [ (8, 1, None, None) , (16393, 2, None, "IID('{0002447B-0000-0000-C000-000000000046}')") , (12, 17, None, None) , 
			 (12, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 840 , (3, 0, None, None) , 0 , )),
	(( 'SaveAsXMLData' , 'Filename' , 'Map' , ), 2278, (2278, (), [ (8, 1, None, None) , 
			 (9, 1, None, "IID('{0002447B-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 844 , (3, 0, None, None) , 0 , )),
	(( 'ToggleFormsDesign' , ), 2279, (2279, (), [ ], 1 , 1 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'ContentTypeProperties' , 'RHS' , ), 2512, (2512, (), [ (16393, 10, None, "IID('{000C038E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 852 , (3, 0, None, None) , 0 , )),
	(( 'Connections' , 'RHS' , ), 2513, (2513, (), [ (16393, 10, None, "IID('{00024486-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'RemoveDocumentInformation' , 'RemoveDocInfoType' , ), 2514, (2514, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 860 , (3, 0, None, None) , 0 , )),
	(( 'Signatures' , 'RHS' , ), 2516, (2516, (), [ (16393, 10, None, "IID('{000C0410-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'CheckInWithVersion' , 'SaveChanges' , 'Comments' , 'MakePublic' , 'VersionType' , 
			 ), 2517, (2517, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 4 , 868 , (3, 0, None, None) , 0 , )),
	(( 'ServerPolicy' , 'RHS' , ), 2519, (2519, (), [ (16393, 10, None, "IID('{000C0390-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'LockServerFile' , ), 2520, (2520, (), [ ], 1 , 1 , 4 , 0 , 876 , (3, 0, None, None) , 0 , )),
	(( 'DocumentInspectors' , 'RHS' , ), 2521, (2521, (), [ (16393, 10, None, "IID('{000C0392-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'GetWorkflowTasks' , 'RHS' , ), 2522, (2522, (), [ (16393, 10, None, "IID('{000CD901-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 884 , (3, 0, None, None) , 0 , )),
	(( 'GetWorkflowTemplates' , 'RHS' , ), 2523, (2523, (), [ (16393, 10, None, "IID('{000CD903-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'PrintOut' , 'From' , 'To' , 'Copies' , 'Preview' , 
			 'ActivePrinter' , 'PrintToFile' , 'Collate' , 'PrToFileName' , 'IgnorePrintAreas' , 
			 'lcid' , ), 2361, (2361, (), [ (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (3, 5, None, None) , ], 1 , 1 , 4 , 9 , 892 , (3, 0, None, None) , 0 , )),
	(( 'ServerViewableItems' , 'RHS' , ), 2524, (2524, (), [ (16393, 10, None, "IID('{000244A4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'TableStyles' , 'RHS' , ), 2525, (2525, (), [ (16393, 10, None, "IID('{000244A8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 900 , (3, 0, None, None) , 0 , )),
	(( 'DefaultTableStyle' , 'RHS' , ), 2526, (2526, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 904 , (3, 0, None, None) , 1024 , )),
	(( 'DefaultTableStyle' , 'RHS' , ), 2526, (2526, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 908 , (3, 0, None, None) , 1024 , )),
	(( 'DefaultPivotTableStyle' , 'RHS' , ), 2527, (2527, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 912 , (3, 0, None, None) , 1024 , )),
	(( 'DefaultPivotTableStyle' , 'RHS' , ), 2527, (2527, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 916 , (3, 0, None, None) , 1024 , )),
	(( 'CheckCompatibility' , 'RHS' , ), 2528, (2528, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'CheckCompatibility' , 'RHS' , ), 2528, (2528, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 924 , (3, 0, None, None) , 0 , )),
	(( 'HasVBProject' , 'RHS' , ), 2529, (2529, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'CustomXMLParts' , 'RHS' , ), 2530, (2530, (), [ (16397, 10, None, "IID('{000CDB0C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 932 , (3, 0, None, None) , 0 , )),
	(( 'Final' , 'RHS' , ), 2531, (2531, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 936 , (3, 0, None, None) , 0 , )),
	(( 'Final' , 'RHS' , ), 2531, (2531, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 940 , (3, 0, None, None) , 0 , )),
	(( 'Research' , 'RHS' , ), 2532, (2532, (), [ (16393, 10, None, "IID('{000244AC-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'Theme' , 'RHS' , ), 2533, (2533, (), [ (16393, 10, None, "IID('{000C03A0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 948 , (3, 0, None, None) , 0 , )),
	(( 'ApplyTheme' , 'Filename' , ), 2534, (2534, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 952 , (3, 0, None, None) , 0 , )),
	(( 'Excel8CompatibilityMode' , 'RHS' , ), 2535, (2535, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 956 , (3, 0, None, None) , 0 , )),
	(( 'ConnectionsDisabled' , 'RHS' , ), 2536, (2536, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 960 , (3, 0, None, None) , 0 , )),
	(( 'EnableConnections' , ), 2537, (2537, (), [ ], 1 , 1 , 4 , 0 , 964 , (3, 0, None, None) , 0 , )),
	(( 'ShowPivotChartActiveFields' , 'RHS' , ), 2538, (2538, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
	(( 'ShowPivotChartActiveFields' , 'RHS' , ), 2538, (2538, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 972 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat' , 'Type' , 'Filename' , 'Quality' , 'IncludeDocProperties' , 
			 'IgnorePrintAreas' , 'From' , 'To' , 'OpenAfterPublish' , 'FixedFormatExtClassPtr' , 
			 ), 2493, (2493, (), [ (3, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , ], 1 , 1 , 4 , 8 , 976 , (3, 0, None, None) , 0 , )),
	(( 'IconSets' , 'RHS' , ), 2539, (2539, (), [ (16393, 10, None, "IID('{0002449C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 980 , (3, 0, None, None) , 0 , )),
	(( 'EncryptionProvider' , 'RHS' , ), 2540, (2540, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
	(( 'EncryptionProvider' , 'RHS' , ), 2540, (2540, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 988 , (3, 0, None, None) , 0 , )),
	(( 'DoNotPromptForConvert' , 'RHS' , ), 2541, (2541, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 992 , (3, 0, None, None) , 0 , )),
	(( 'DoNotPromptForConvert' , 'RHS' , ), 2541, (2541, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 996 , (3, 0, None, None) , 0 , )),
	(( 'ForceFullCalculation' , 'RHS' , ), 2542, (2542, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1000 , (3, 0, None, None) , 0 , )),
	(( 'ForceFullCalculation' , 'RHS' , ), 2542, (2542, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1004 , (3, 0, None, None) , 0 , )),
	(( 'ProtectSharing' , 'Filename' , 'Password' , 'WriteResPassword' , 'ReadOnlyRecommended' , 
			 'CreateBackup' , 'SharingPassword' , 'FileFormat' , ), 2543, (2543, (), [ (12, 17, None, None) , 
			 (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , (12, 17, None, None) , 
			 (12, 17, None, None) , ], 1 , 1 , 4 , 7 , 1008 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000208DA-0000-0000-C000-000000000046}", _Workbook )
