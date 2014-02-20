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

from win32com.client import CoClassBaseClass
import sys
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents3')
ApplicationEvents3 = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents3'].ApplicationEvents3
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents4')
ApplicationEvents4 = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents4'].ApplicationEvents4
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents2')
ApplicationEvents2 = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents2'].ApplicationEvents2
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents')
ApplicationEvents = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4.ApplicationEvents'].ApplicationEvents
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4._Application')
_Application = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x4._Application']._Application
# This CoClass is known by the name 'Word.Application.12'
class Application(CoClassBaseClass): # A CoClass
	CLSID = IID('{000209FF-0000-0000-C000-000000000046}')
	coclass_sources = [
		ApplicationEvents3,
		ApplicationEvents4,
		ApplicationEvents2,
		ApplicationEvents,
	]
	default_source = ApplicationEvents4
	coclass_interfaces = [
		_Application,
	]
	default_interface = _Application

win32com.client.CLSIDToClass.RegisterCLSID( "{000209FF-0000-0000-C000-000000000046}", Application )
