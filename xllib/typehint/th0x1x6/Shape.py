# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.3.2 (v3.3.2:d047928ae3f6, May 16 2013, 00:03:43) [MSC v.1600 32 bit (Intel)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Mon Oct 14 16:39:11 2013
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
class Shape(DispatchBaseClass):
	CLSID = IID('{00024439-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Apply(self):
		return self._oleobj_.InvokeTypes(1675, LCID, 1, (24, 0), (),)

	def CanvasCropBottom(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2175, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def CanvasCropLeft(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2172, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def CanvasCropRight(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2174, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def CanvasCropTop(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(2173, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def Copy(self):
		return self._oleobj_.InvokeTypes(551, LCID, 1, (24, 0), (),)

	def CopyPicture(self, Appearance=defaultNamedOptArg, Format=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(213, LCID, 1, (24, 0), ((12, 17), (12, 17)),Appearance
			, Format)

	def Cut(self):
		return self._oleobj_.InvokeTypes(565, LCID, 1, (24, 0), (),)

	def Delete(self):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (24, 0), (),)

	# Result is of type Shape
	def Duplicate(self):
		ret = self._oleobj_.InvokeTypes(1039, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'Duplicate', '{00024439-0000-0000-C000-000000000046}')
		return ret

	def Flip(self, FlipCmd=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1676, LCID, 1, (24, 0), ((3, 1),),FlipCmd
			)

	def IncrementLeft(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1678, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def IncrementRotation(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1680, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def IncrementTop(self, Increment=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1681, LCID, 1, (24, 0), ((4, 1),),Increment
			)

	def PickUp(self):
		return self._oleobj_.InvokeTypes(1682, LCID, 1, (24, 0), (),)

	def RerouteConnections(self):
		return self._oleobj_.InvokeTypes(1683, LCID, 1, (24, 0), (),)

	def ScaleHeight(self, Factor=defaultNamedNotOptArg, RelativeToOriginalSize=defaultNamedNotOptArg, Scale=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1684, LCID, 1, (24, 0), ((4, 1), (3, 1), (12, 17)),Factor
			, RelativeToOriginalSize, Scale)

	def ScaleWidth(self, Factor=defaultNamedNotOptArg, RelativeToOriginalSize=defaultNamedNotOptArg, Scale=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1688, LCID, 1, (24, 0), ((4, 1), (3, 1), (12, 17)),Factor
			, RelativeToOriginalSize, Scale)

	def Select(self, Replace=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(235, LCID, 1, (24, 0), ((12, 17),),Replace
			)

	def SetShapesDefaultProperties(self):
		return self._oleobj_.InvokeTypes(1689, LCID, 1, (24, 0), (),)

	# Result is of type ShapeRange
	def Ungroup(self):
		ret = self._oleobj_.InvokeTypes(244, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'Ungroup', '{0002443B-0000-0000-C000-000000000046}')
		return ret

	def ZOrder(self, ZOrderCmd=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(622, LCID, 1, (24, 0), ((3, 1),),ZOrderCmd
			)

	_prop_map_get_ = {
		# Method 'Adjustments' returns object of type 'Adjustments'
		"Adjustments": (1691, 2, (9, 0), (), "Adjustments", '{000C0310-0000-0000-C000-000000000046}'),
		"AlternativeText": (1891, 2, (8, 0), (), "AlternativeText", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (148, 2, (13, 0), (), "Application", '{00024500-0000-0000-C000-000000000046}'),
		"AutoShapeType": (1693, 2, (3, 0), (), "AutoShapeType", None),
		"BackgroundStyle": (2661, 2, (3, 0), (), "BackgroundStyle", None),
		"BlackWhiteMode": (1707, 2, (3, 0), (), "BlackWhiteMode", None),
		# Method 'BottomRightCell' returns object of type 'Range'
		"BottomRightCell": (615, 2, (9, 0), (), "BottomRightCell", '{00020846-0000-0000-C000-000000000046}'),
		# Method 'Callout' returns object of type 'CalloutFormat'
		"Callout": (1694, 2, (9, 0), (), "Callout", '{000C0311-0000-0000-C000-000000000046}'),
		# Method 'CanvasItems' returns object of type 'CanvasShapes'
		"CanvasItems": (2171, 2, (9, 0), (), "CanvasItems", '{000C0371-0000-0000-C000-000000000046}'),
		# Method 'Chart' returns object of type 'Chart'
		"Chart": (7, 2, (13, 0), (), "Chart", '{00020821-0000-0000-C000-000000000046}'),
		"Child": (2169, 2, (3, 0), (), "Child", None),
		"ConnectionSiteCount": (1695, 2, (3, 0), (), "ConnectionSiteCount", None),
		"Connector": (1696, 2, (3, 0), (), "Connector", None),
		# Method 'ConnectorFormat' returns object of type 'ConnectorFormat'
		"ConnectorFormat": (1697, 2, (9, 0), (), "ConnectorFormat", '{0002443E-0000-0000-C000-000000000046}'),
		# Method 'ControlFormat' returns object of type 'ControlFormat'
		"ControlFormat": (1709, 2, (9, 0), (), "ControlFormat", '{00024440-0000-0000-C000-000000000046}'),
		"Creator": (149, 2, (3, 0), (), "Creator", None),
		# Method 'Diagram' returns object of type 'Diagram'
		"Diagram": (2167, 2, (9, 0), (), "Diagram", '{0002446F-0000-0000-C000-000000000046}'),
		# Method 'DiagramNode' returns object of type 'DiagramNode'
		"DiagramNode": (2165, 2, (9, 0), (), "DiagramNode", '{000C0370-0000-0000-C000-000000000046}'),
		"DrawingObject": (1708, 2, (9, 0), (), "DrawingObject", None),
		# Method 'Fill' returns object of type 'FillFormat'
		"Fill": (1663, 2, (9, 0), (), "Fill", '{000C0314-0000-0000-C000-000000000046}'),
		"FormControlType": (1712, 2, (3, 0), (), "FormControlType", None),
		# Method 'Glow' returns object of type 'GlowFormat'
		"Glow": (2663, 2, (9, 0), (), "Glow", '{000C03BD-0000-0000-C000-000000000046}'),
		# Method 'GroupItems' returns object of type 'GroupShapes'
		"GroupItems": (1698, 2, (9, 0), (), "GroupItems", '{0002443C-0000-0000-C000-000000000046}'),
		"HasChart": (2658, 2, (3, 0), (), "HasChart", None),
		"HasDiagram": (2168, 2, (3, 0), (), "HasDiagram", None),
		"HasDiagramNode": (2166, 2, (3, 0), (), "HasDiagramNode", None),
		"Height": (123, 2, (4, 0), (), "Height", None),
		"HorizontalFlip": (1699, 2, (3, 0), (), "HorizontalFlip", None),
		# Method 'Hyperlink' returns object of type 'Hyperlink'
		"Hyperlink": (1706, 2, (9, 0), (), "Hyperlink", '{00024431-0000-0000-C000-000000000046}'),
		"ID": (570, 2, (3, 0), (), "ID", None),
		"Left": (127, 2, (4, 0), (), "Left", None),
		# Method 'Line' returns object of type 'LineFormat'
		"Line": (817, 2, (9, 0), (), "Line", '{000C0317-0000-0000-C000-000000000046}'),
		# Method 'LinkFormat' returns object of type 'LinkFormat'
		"LinkFormat": (1710, 2, (9, 0), (), "LinkFormat", '{00024442-0000-0000-C000-000000000046}'),
		"LockAspectRatio": (1700, 2, (3, 0), (), "LockAspectRatio", None),
		"Locked": (269, 2, (11, 0), (), "Locked", None),
		"Name": (110, 2, (8, 0), (), "Name", None),
		# Method 'Nodes' returns object of type 'ShapeNodes'
		"Nodes": (1701, 2, (9, 0), (), "Nodes", '{000C0319-0000-0000-C000-000000000046}'),
		# Method 'OLEFormat' returns object of type 'OLEFormat'
		"OLEFormat": (1711, 2, (9, 0), (), "OLEFormat", '{00024441-0000-0000-C000-000000000046}'),
		"OnAction": (596, 2, (8, 0), (), "OnAction", None),
		"Parent": (150, 2, (9, 0), (), "Parent", None),
		# Method 'ParentGroup' returns object of type 'Shape'
		"ParentGroup": (2170, 2, (9, 0), (), "ParentGroup", '{00024439-0000-0000-C000-000000000046}'),
		# Method 'PictureFormat' returns object of type 'PictureFormat'
		"PictureFormat": (1631, 2, (9, 0), (), "PictureFormat", '{000C031A-0000-0000-C000-000000000046}'),
		"Placement": (617, 2, (3, 0), (), "Placement", None),
		# Method 'Reflection' returns object of type 'ReflectionFormat'
		"Reflection": (2664, 2, (9, 0), (), "Reflection", '{000C03BE-0000-0000-C000-000000000046}'),
		"Rotation": (59, 2, (4, 0), (), "Rotation", None),
		# Method 'Script' returns object of type 'Script'
		"Script": (1892, 2, (9, 0), (), "Script", '{000C0341-0000-0000-C000-000000000046}'),
		# Method 'Shadow' returns object of type 'ShadowFormat'
		"Shadow": (103, 2, (9, 0), (), "Shadow", '{000C031B-0000-0000-C000-000000000046}'),
		"ShapeStyle": (2660, 2, (3, 0), (), "ShapeStyle", None),
		# Method 'SoftEdge' returns object of type 'SoftEdgeFormat'
		"SoftEdge": (2662, 2, (9, 0), (), "SoftEdge", '{000C03BC-0000-0000-C000-000000000046}'),
		# Method 'TextEffect' returns object of type 'TextEffectFormat'
		"TextEffect": (1702, 2, (9, 0), (), "TextEffect", '{000C031F-0000-0000-C000-000000000046}'),
		# Method 'TextFrame' returns object of type 'TextFrame'
		"TextFrame": (1692, 2, (9, 0), (), "TextFrame", '{0002443D-0000-0000-C000-000000000046}'),
		# Method 'TextFrame2' returns object of type 'TextFrame2'
		"TextFrame2": (2659, 2, (9, 0), (), "TextFrame2", '{000C0398-0000-0000-C000-000000000046}'),
		# Method 'ThreeD' returns object of type 'ThreeDFormat'
		"ThreeD": (1703, 2, (9, 0), (), "ThreeD", '{000C0321-0000-0000-C000-000000000046}'),
		"Top": (126, 2, (4, 0), (), "Top", None),
		# Method 'TopLeftCell' returns object of type 'Range'
		"TopLeftCell": (620, 2, (9, 0), (), "TopLeftCell", '{00020846-0000-0000-C000-000000000046}'),
		"Type": (108, 2, (3, 0), (), "Type", None),
		"VerticalFlip": (1704, 2, (3, 0), (), "VerticalFlip", None),
		"Vertices": (621, 2, (12, 0), (), "Vertices", None),
		"Visible": (558, 2, (3, 0), (), "Visible", None),
		"Width": (122, 2, (4, 0), (), "Width", None),
		"ZOrderPosition": (1705, 2, (3, 0), (), "ZOrderPosition", None),
	}
	_prop_map_put_ = {
		"AlternativeText": ((1891, LCID, 4, 0),()),
		"AutoShapeType": ((1693, LCID, 4, 0),()),
		"BackgroundStyle": ((2661, LCID, 4, 0),()),
		"BlackWhiteMode": ((1707, LCID, 4, 0),()),
		"Height": ((123, LCID, 4, 0),()),
		"Left": ((127, LCID, 4, 0),()),
		"LockAspectRatio": ((1700, LCID, 4, 0),()),
		"Locked": ((269, LCID, 4, 0),()),
		"Name": ((110, LCID, 4, 0),()),
		"OnAction": ((596, LCID, 4, 0),()),
		"Placement": ((617, LCID, 4, 0),()),
		"Rotation": ((59, LCID, 4, 0),()),
		"ShapeStyle": ((2660, LCID, 4, 0),()),
		"Top": ((126, LCID, 4, 0),()),
		"Visible": ((558, LCID, 4, 0),()),
		"Width": ((122, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00024439-0000-0000-C000-000000000046}", Shape )
