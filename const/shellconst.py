"""snippet used to generate file:

import ctypes.wintypes
from win32com.shell import shell, shellcon
import types

with open("C:\\Users\\PBS Biotech\\Documents\\Personal\\PBS_Office\\MSOffice\\officelib\\shellconst.py", 'w') as f:

    for var in dir(shellcon):

        if not var.startswith('__'):
            value = getattr(shellcon, var)
#             print(value)

            if isinstance(value, int):
                value = str(value)
            elif isinstance(value, str):
                value = value.replace('\'', '\\\'')
                value = ''.join(['\"', value, '\"'])
            elif isinstance(value, tuple) and len(value) == 1:
                value = str(value[0])
            else:
                continue

            f.write(''.join([var, ' = ', value, '\n']))
"""


ABE_BOTTOM = 3
ABE_LEFT = 0
ABE_RIGHT = 2
ABE_TOP = 1
ABM_ACTIVATE = 6
ABM_GETAUTOHIDEBAR = 7
ABM_GETSTATE = 4
ABM_GETTASKBARPOS = 5
ABM_NEW = 0
ABM_QUERYPOS = 2
ABM_REMOVE = 1
ABM_SETAUTOHIDEBAR = 8
ABM_SETPOS = 3
ABM_WINDOWPOSCHANGED = 9
ABN_FULLSCREENAPP = 2
ABN_POSCHANGED = 1
ABN_STATECHANGE = 0
ABN_WINDOWARRANGE = 3
ABS_ALWAYSONTOP = 2
ABS_AUTOHIDE = 1
ADDURL_SILENT = 1
ADLT_FREQUENT = 1
ADLT_RECENT = 0
AD_APPLY_ALL = 7
AD_APPLY_BUFFERED_REFRESH = 16
AD_APPLY_DYNAMICREFRESH = 32
AD_APPLY_FORCE = 8
AD_APPLY_HTMLGEN = 2
AD_APPLY_REFRESH = 4
AD_APPLY_SAVE = 1
ASSOCCLASS_APP_KEY = 5
ASSOCCLASS_APP_STR = 6
ASSOCCLASS_CLSID_KEY = 3
ASSOCCLASS_CLSID_STR = 4
ASSOCCLASS_FOLDER = 8
ASSOCCLASS_PROGID_KEY = 1
ASSOCCLASS_PROGID_STR = 2
ASSOCCLASS_SHELL_KEY = 0
ASSOCCLASS_STAR = 9
ASSOCCLASS_SYSTEM_STR = 7
ASSOCDATA_EDITFLAGS = 5
ASSOCDATA_HASPERUSERASSOC = 4
ASSOCDATA_MSIDESCRIPTOR = 1
ASSOCDATA_NOACTIVATEHANDLER = 2
ASSOCDATA_QUERYCLASSSTORE = 3
ASSOCDATA_VALUE = 6
ASSOCF_IGNOREBASECLASS = 512
ASSOCF_INIT_BYEXENAME = 2
ASSOCF_INIT_DEFAULTTOFOLDER = 8
ASSOCF_INIT_DEFAULTTOSTAR = 4
ASSOCF_INIT_NOREMAPCLSID = 1
ASSOCF_NOFIXUPS = 256
ASSOCF_NOTRUNCATE = 32
ASSOCF_NOUSERSETTINGS = 16
ASSOCF_OPEN_BYEXENAME = 2
ASSOCF_REMAPRUNDLL = 128
ASSOCF_VERIFY = 64
ASSOCKEY_APP = 2
ASSOCKEY_BASECLASS = 4
ASSOCKEY_CLASS = 3
ASSOCKEY_SHELLEXECCLASS = 1
ASSOCSTR_COMMAND = 1
ASSOCSTR_CONTENTTYPE = 14
ASSOCSTR_DDEAPPLICATION = 9
ASSOCSTR_DDECOMMAND = 7
ASSOCSTR_DDEIFEXEC = 8
ASSOCSTR_DDETOPIC = 10
ASSOCSTR_DEFAULTICON = 15
ASSOCSTR_EXECUTABLE = 2
ASSOCSTR_FRIENDLYAPPNAME = 4
ASSOCSTR_FRIENDLYDOCNAME = 3
ASSOCSTR_INFOTIP = 11
ASSOCSTR_NOOPEN = 5
ASSOCSTR_QUICKTIP = 12
ASSOCSTR_SHELLEXTENSION = 16
ASSOCSTR_SHELLNEWVALUE = 6
ASSOCSTR_TILEINFO = 13
BFFM_ENABLEOK = 1125
BFFM_INITIALIZED = 1
BFFM_SELCHANGED = 2
BFFM_SETSELECTION = 1126
BFFM_SETSELECTIONA = 1126
BFFM_SETSELECTIONW = 1127
BFFM_SETSTATUSTEXT = 1124
BFFM_SETSTATUSTEXTA = 1124
BFFM_SETSTATUSTEXTW = 1128
BFFM_VALIDATEFAILED = 3
BFFM_VALIDATEFAILEDA = 3
BFFM_VALIDATEFAILEDW = 4
BFO_ADD_IE_TOCAPTIONBAR = 512
BFO_BOTH_OPTIONS = 4
BFO_BROWSER_PERSIST_SETTINGS = 1
BFO_BROWSE_NO_IN_NEW_PROCESS = 16
BFO_ENABLE_HYPERLINK_TRACKING = 32
BFO_GO_HOME_PAGE = 16384
BFO_NONE = 0
BFO_NO_PARENT_FOLDER_SUPPORT = 4096
BFO_NO_REOPEN_NEXT_RESTART = 8192
BFO_PREFER_IEPROCESS = 32768
BFO_QUERY_ALL = -1
BFO_RENAME_FOLDER_OPTIONS_TOINTERNET = 2
BFO_SHOW_NAVIGATION_CANCELLED = 65536
BFO_SUBSTITUE_INTERNET_START_PAGE = 128
BFO_USE_DIALUP_REF = 1024
BFO_USE_IE_LOGOBANDING = 256
BFO_USE_IE_OFFLINE_SUPPORT = 64
BFO_USE_IE_TOOLBAR = 2048
BIF_BROWSEFORCOMPUTER = 4096
BIF_BROWSEFORPRINTER = 8192
BIF_BROWSEINCLUDEFILES = 16384
BIF_DONTGOBELOWDOMAIN = 2
BIF_EDITBOX = 16
BIF_PREFER_INTERNET_SHORTCUT = 8
BIF_RETURNFSANCESTORS = 8
BIF_RETURNONLYFSDIRS = 1
BIF_STATUSTEXT = 4
BIF_VALIDATE = 32
CDBOSC_KILLFOCUS = 1
CDBOSC_RENAME = 3
CDBOSC_SELCHANGE = 2
CDBOSC_SETFOCUS = 0
CFSTR_AUTOPLAY_SHELLIDLISTS = "Autoplay Enumerated IDList Array"
CFSTR_DRAGCONTEXT = "DragContext"
CFSTR_FILECONTENTS = "FileContents"
CFSTR_FILEDESCRIPTOR = "FileGroupDescriptor"
CFSTR_FILEDESCRIPTORA = "FileGroupDescriptor"
CFSTR_FILEDESCRIPTORW = "FileGroupDescriptorW"
CFSTR_FILENAME = "FileName"
CFSTR_FILENAMEA = "FileName"
CFSTR_FILENAMEMAP = "FileNameMap"
CFSTR_FILENAMEMAPA = "FileNameMap"
CFSTR_FILENAMEMAPW = "FileNameMapW"
CFSTR_FILENAMEW = "FileNameW"
CFSTR_INDRAGLOOP = "InShellDragLoop"
CFSTR_INETURLA = "UniformResourceLocator"
CFSTR_INETURLW = "UniformResourceLocatorW"
CFSTR_LOGICALPERFORMEDDROPEFFECT = "Logical Performed DropEffect"
CFSTR_MOUNTEDVOLUME = "MountedVolume"
CFSTR_NETRESOURCES = "Net Resource"
CFSTR_PASTESUCCEEDED = "Paste Succeeded"
CFSTR_PERFORMEDDROPEFFECT = "Performed DropEffect"
CFSTR_PERSISTEDDATAOBJECT = "PersistedDataObject"
CFSTR_PREFERREDDROPEFFECT = "Preferred DropEffect"
CFSTR_PRINTERGROUP = "PrinterFriendlyName"
CFSTR_SHELLIDLIST = "Shell IDList Array"
CFSTR_SHELLIDLISTOFFSET = "Shell Object Offsets"
CFSTR_SHELLURL = "UniformResourceLocator"
CFSTR_TARGETCLSID = "TargetCLSID"
CLSID_ExplorerBrowser = "{71f96385-ddd6-48d3-a0c1-ae06e8b055fb}"
CMDSTR_NEWFOLDER = "NewFolder"
CMDSTR_NEWFOLDERA = "NewFolder"
CMDSTR_VIEWDETAILS = "ViewDetails"
CMDSTR_VIEWDETAILSA = "ViewDetails"
CMDSTR_VIEWLIST = "ViewList"
CMDSTR_VIEWLISTA = "ViewList"
CMF_CANRENAME = 16
CMF_DEFAULTONLY = 1
CMF_EXPLORE = 4
CMF_INCLUDESTATIC = 64
CMF_NODEFAULT = 32
CMF_NORMAL = 0
CMF_NOVERBS = 8
CMF_RESERVED = -65536
CMF_VERBSONLY = 2
CMIC_MASK_ASYNCOK = 1048576
CMIC_MASK_FLAG_NO_UI = 1024
CMIC_MASK_HOTKEY = 32
CMIC_MASK_ICON = 16
CMIC_MASK_NO_CONSOLE = 32768
CMIC_MASK_PTINVOKE = 536870912
CMIC_MASK_UNICODE = 16384
COMPONENT_DEFAULT_LEFT = 65535
COMPONENT_DEFAULT_TOP = 65535
COMPONENT_TOP = 1073741823
COMP_ELEM_ALL = 32767
COMP_ELEM_CHECKED = 2
COMP_ELEM_CURITEMSTATE = 16384
COMP_ELEM_DIRTY = 4
COMP_ELEM_FRIENDLYNAME = 1024
COMP_ELEM_NOSCROLL = 8
COMP_ELEM_ORIGINAL_CSI = 4096
COMP_ELEM_POS_LEFT = 16
COMP_ELEM_POS_TOP = 32
COMP_ELEM_POS_ZINDEX = 256
COMP_ELEM_RESTORED_CSI = 8192
COMP_ELEM_SIZE_HEIGHT = 128
COMP_ELEM_SIZE_WIDTH = 64
COMP_ELEM_SOURCE = 512
COMP_ELEM_SUBSCRIBEDURL = 2048
COMP_ELEM_TYPE = 1
COMP_TYPE_CFHTML = 4
COMP_TYPE_CONTROL = 3
COMP_TYPE_HTMLDOC = 0
COMP_TYPE_MAX = 4
COMP_TYPE_PICTURE = 1
COMP_TYPE_WEBSITE = 2
CSIDL_ADMINTOOLS = 48
CSIDL_ALTSTARTUP = 29
CSIDL_APPDATA = 26
CSIDL_BITBUCKET = 10
CSIDL_CDBURN_AREA = 59
CSIDL_COMMON_ADMINTOOLS = 47
CSIDL_COMMON_ALTSTARTUP = 30
CSIDL_COMMON_APPDATA = 35
CSIDL_COMMON_DESKTOPDIRECTORY = 25
CSIDL_COMMON_DOCUMENTS = 46
CSIDL_COMMON_FAVORITES = 31
CSIDL_COMMON_MUSIC = 53
CSIDL_COMMON_OEM_LINKS = 58
CSIDL_COMMON_PICTURES = 54
CSIDL_COMMON_PROGRAMS = 23
CSIDL_COMMON_STARTMENU = 22
CSIDL_COMMON_STARTUP = 24
CSIDL_COMMON_TEMPLATES = 45
CSIDL_COMMON_VIDEO = 55
CSIDL_COMPUTERSNEARME = 61
CSIDL_CONNECTIONS = 49
CSIDL_CONTROLS = 3
CSIDL_COOKIES = 33
CSIDL_DESKTOP = 0
CSIDL_DESKTOPDIRECTORY = 16
CSIDL_DRIVES = 17
CSIDL_FAVORITES = 6
CSIDL_FONTS = 20
CSIDL_HISTORY = 34
CSIDL_INTERNET = 1
CSIDL_INTERNET_CACHE = 32
CSIDL_LOCAL_APPDATA = 28
CSIDL_MYDOCUMENTS = 12
CSIDL_MYMUSIC = 13
CSIDL_MYPICTURES = 39
CSIDL_MYVIDEO = 14
CSIDL_NETHOOD = 19
CSIDL_NETWORK = 18
CSIDL_PERSONAL = 5
CSIDL_PRINTERS = 4
CSIDL_PRINTHOOD = 27
CSIDL_PROFILE = 40
CSIDL_PROGRAMS = 2
CSIDL_PROGRAM_FILES = 38
CSIDL_PROGRAM_FILESX86 = 42
CSIDL_PROGRAM_FILES_COMMON = 43
CSIDL_PROGRAM_FILES_COMMONX86 = 44
CSIDL_RECENT = 8
CSIDL_RESOURCES = 56
CSIDL_RESOURCES_LOCALIZED = 57
CSIDL_SENDTO = 9
CSIDL_STARTMENU = 11
CSIDL_STARTUP = 7
CSIDL_SYSTEM = 37
CSIDL_SYSTEMX86 = 41
CSIDL_TEMPLATES = 21
CSIDL_WINDOWS = 36
DBIF_VIEWMODE_FLOATING = 2
DBIF_VIEWMODE_NORMAL = 0
DBIF_VIEWMODE_TRANSPARENT = 4
DBIF_VIEWMODE_VERTICAL = 1
DBIMF_BKCOLOR = 64
DBIMF_DEBOSSED = 32
DBIMF_NORMAL = 0
DBIMF_VARIABLEHEIGHT = 8
DBIM_ACTUAL = 8
DBIM_BKCOLOR = 64
DBIM_INTEGRAL = 4
DBIM_MAXSIZE = 2
DBIM_MINSIZE = 1
DBIM_MODEFLAGS = 32
DBIM_TITLE = 16
DROPEFFECT_COPY = 1
DROPEFFECT_LINK = 4
DROPEFFECT_MOVE = 2
DROPEFFECT_NONE = 0
DROPEFFECT_SCROLL = -2147483648
DSFT_DETECT = 1
DSFT_PRIVATE = 2
DSFT_PUBLIC = 3
DTI_ADDUI_DEFAULT = 0
DTI_ADDUI_DISPSUBWIZARD = 1
DTI_ADDUI_POSITIONITEM = 2
DVASPECT_SHORTNAME = 2
DWFAF_HIDDEN = 1
DWFRF_DELETECONFIGDATA = 1
DWFRF_NORMAL = 0
EBF_NODROPTARGET = 512
EBF_NONE = 0
EBF_SELECTFROMDATAOBJECT = 256
EBO_ALWAYSNAVIGATE = 4
EBO_NAVIGATEONCE = 1
EBO_NONE = 0
EBO_NOTRAVELLOG = 8
EBO_NOWRAPPERWINDOW = 16
EBO_SHOWFRAMES = 2
ECF_HASLUASHIELD = 16
ECF_HASSPLITBUTTON = 2
ECF_HASSUBCOMMANDS = 1
ECF_HIDELABEL = 4
ECF_ISSEPARATOR = 8
ECS_CHECKBOX = 4
ECS_CHECKED = 8
ECS_DISABLED = 1
ECS_ENABLED = 0
ECS_HIDDEN = 2
EVCCBF_LASTNOTIFICATION = 1
EVCF_DONTSHOWIFZERO = 16
EVCF_ENABLEBYDEFAULT = 2
EVCF_ENABLEBYDEFAULT_AUTO = 8
EVCF_HASSETTINGS = 1
EVCF_OUTOFDISKSPACE = 64
EVCF_REMOVEFROMLIST = 4
EVCF_SETTINGSMODE = 32
EXP_DARWIN_ID_SIG = 2684354566
EXP_LOGO3_ID_SIG = 2684354567
EXP_SPECIAL_FOLDER_SIG = 2684354565
EXP_SZ_ICON_SIG = 2684354567
EXP_SZ_LINK_SIG = 2684354561
FCIDM_BROWSERFIRST = 40960
FCIDM_BROWSERLAST = 48896
FCIDM_GLOBALFIRST = 32768
FCIDM_GLOBALLAST = 40959
FCIDM_MENU_EDIT = 32832
FCIDM_MENU_EXPLORE = 33104
FCIDM_MENU_FAVORITES = 33136
FCIDM_MENU_FILE = 32768
FCIDM_MENU_FIND = 33088
FCIDM_MENU_HELP = 33024
FCIDM_MENU_TOOLS = 32960
FCIDM_MENU_TOOLS_SEP_GOTO = 32961
FCIDM_MENU_VIEW = 32896
FCIDM_MENU_VIEW_SEP_OPTIONS = 32897
FCIDM_SHVIEWFIRST = 0
FCIDM_SHVIEWLAST = 32767
FCIDM_STATUS = 40961
FCIDM_TOOLBAR = 40960
FCT_ADDTOEND = 4
FCT_CONFIGABLE = 2
FCT_MERGE = 1
FCW_INTERNETBAR = 6
FCW_PROGRESS = 8
FCW_STATUS = 1
FCW_TOOLBAR = 2
FCW_TREE = 3
FD_ACCESSTIME = 16
FD_ATTRIBUTES = 4
FD_CLSID = 1
FD_CREATETIME = 8
FD_FILESIZE = 64
FD_LINKUI = 32768
FD_PROGRESSUI = 16384
FD_SIZEPOINT = 2
FD_WRITESTIME = 32
FFFP_EXACTMATCH = 0
FFFP_NEARESTPARENTMATCH = 1
FOF_ALLOWUNDO = 64
FOF_CONFIRMMOUSE = 2
FOF_FILESONLY = 128
FOF_MULTIDESTFILES = 1
FOF_NOCONFIRMATION = 16
FOF_NOCONFIRMMKDIR = 512
FOF_NOCOPYSECURITYATTRIBS = 2048
FOF_NOERRORUI = 1024
FOF_RENAMEONCOLLISION = 8
FOF_SILENT = 4
FOF_SIMPLEPROGRESS = 256
FOF_WANTMAPPINGHANDLE = 32
FO_COPY = 2
FO_DELETE = 3
FO_MOVE = 1
FO_RENAME = 4
FVM_DETAILS = 4
FVM_FIRST = 1
FVM_ICON = 1
FVM_LIST = 3
FVM_SMALLICON = 2
FVM_THUMBNAIL = 5
FVM_THUMBSTRIP = 7
FVM_TILE = 6
FVSIF_CANVIEWIT = 1073741824
FVSIF_NEWFAILED = 134217728
FVSIF_NEWFILE = -2147483648
FVSIF_PINNED = 2
FVSIF_RECT = 1
FWF_ABBREVIATEDNAMES = 2
FWF_ALIGNLEFT = 2048
FWF_AUTOARRANGE = 1
FWF_BESTFITWINDOW = 16
FWF_CHECKSELECT = 262144
FWF_DESKTOP = 32
FWF_HIDEFILENAMES = 131072
FWF_NOCLIENTEDGE = 512
FWF_NOICONS = 4096
FWF_NOSCROLL = 1024
FWF_NOSUBFOLDERS = 128
FWF_NOVISIBLE = 16384
FWF_NOWEBVIEW = 65536
FWF_OWNERDATA = 8
FWF_SHOWSELALWAYS = 8192
FWF_SINGLECLICKACTIVATE = 32768
FWF_SINGLESEL = 64
FWF_SNAPTOGRID = 4
FWF_TRANSPARENT = 256
GADOF_DIRTY = 1
GCS_HELPTEXT = 1
GCS_HELPTEXTA = 1
GCS_HELPTEXTW = 5
GCS_UNICODE = 4
GCS_VALIDATE = 2
GCS_VALIDATEA = 2
GCS_VALIDATEW = 6
GCS_VERB = 0
GCS_VERBA = 0
GCS_VERBW = 4
GIL_ASYNC = 32
GIL_CHECKSHIELD = 512
GIL_DEFAULTICON = 64
GIL_DONTCACHE = 16
GIL_FORCENOSHIELD = 1024
GIL_FORSHELL = 2
GIL_FORSHORTCUT = 128
GIL_NOTFILENAME = 8
GIL_OPENICON = 1
GIL_PERCLASS = 4
GIL_PERINSTANCE = 2
GIL_SHIELD = 512
GIL_SIMULATEDOC = 1
GPS_BESTEFFORT = 64
GPS_DEFAULT = 0
GPS_DELAYCREATION = 32
GPS_FASTPROPERTIESONLY = 8
GPS_HANDLERPROPERTIESONLY = 1
GPS_MASK_VALID = 127
GPS_OPENSLOWITEM = 16
GPS_READWRITE = 2
GPS_TEMPORARY = 4
IDC_OFFLINE_HAND = 103
ID_PSREBOOTSYSTEM = 3
ID_PSRESTARTWINDOWS = 2
ISIOI_ICONFILE = 1
ISIOI_ICONINDEX = 2
ISIOI_SYSIMAGELISTINDEX = 4
ISOLATION_AWARE_BUILD_STATIC_LIBRARY = 0
ISOLATION_AWARE_USE_STATIC_LIBRARY = 0
IS_FULLSCREEN = 2
IS_NORMAL = 1
IS_SPLIT = 4
IS_VALIDSIZESTATEBITS = 7
IS_VALIDSTATEBITS = 3221225479
IURL_INVOKECOMMAND_FL_ALLOW_UI = 1
IURL_INVOKECOMMAND_FL_DDEWAIT = 4
IURL_INVOKECOMMAND_FL_USE_DEFAULT_VERB = 2
IURL_SETURL_FL_GUESS_PROTOCOL = 1
IURL_SETURL_FL_USE_DEFAULT_PROTOCOL = 2
KDC_FREQUENT = 1
KDC_RECENT = 2
KF_CATEGORY_COMMON = 3
KF_CATEGORY_FIXED = 2
KF_CATEGORY_PERUSER = 4
KF_CATEGORY_VIRTUAL = 1
KF_FLAG_CREATE = 32768
KF_FLAG_DEFAULT_PATH = 1024
KF_FLAG_DONT_UNEXPAND = 8192
KF_FLAG_DONT_VERIFY = 16384
KF_FLAG_INIT = 2048
KF_FLAG_NOT_PARENT_RELATIVE = 512
KF_FLAG_NO_ALIAS = 4096
KF_FLAG_SIMPLE_IDLIST = 256
KF_REDIRECTION_CAPABILITIES_ALLOW_ALL = 255
KF_REDIRECTION_CAPABILITIES_DENY_ALL = 1048320
KF_REDIRECTION_CAPABILITIES_DENY_PERMISSIONS = 1024
KF_REDIRECTION_CAPABILITIES_DENY_POLICY = 512
KF_REDIRECTION_CAPABILITIES_DENY_POLICY_REDIRECTED = 256
KF_REDIRECTION_CAPABILITIES_REDIRECTABLE = 1
KF_REDIRECT_CHECK_ONLY = 16
KF_REDIRECT_COPY_CONTENTS = 512
KF_REDIRECT_COPY_SOURCE_DACL = 2
KF_REDIRECT_DEL_SOURCE_CONTENTS = 1024
KF_REDIRECT_EXCLUDE_ALL_KNOWN_SUBFOLDERS = 2048
KF_REDIRECT_OWNER_USER = 4
KF_REDIRECT_PIN = 128
KF_REDIRECT_SET_OWNER_EXPLICIT = 8
KF_REDIRECT_UNPIN = 64
KF_REDIRECT_USER_EXCLUSIVE = 1
KF_REDIRECT_WITH_UI = 32
LFF_ALLITEMS = 3
LFF_FORCEFILESYSTEM = 1
LFF_STORAGEITEMS = 2
LOF_DEFAULT = 0
LOF_MASK_ALL = 1
LOF_PINNEDTONAVPANE = 1
LSF_FAILIFTHERE = 0
LSF_MAKEUNIQUENAME = 2
LSF_OVERRIDEEXISTING = 1
MAXPROPPAGES = 100
NIF_ICON = 2
NIF_MESSAGE = 1
NIF_TIP = 4
NIM_ADD = 0
NIM_DELETE = 2
NIM_MODIFY = 1
NSTCGNI_CHILD = 5
NSTCGNI_FIRSTVISIBLE = 6
NSTCGNI_LASTVISIBLE = 7
NSTCGNI_NEXT = 0
NSTCGNI_NEXTVISIBLE = 1
NSTCGNI_PARENT = 4
NSTCGNI_PREV = 2
NSTCGNI_PREVVISIBLE = 3
NSTCIS_BOLD = 4
NSTCIS_DISABLED = 8
NSTCIS_EXPANDED = 2
NSTCIS_NONE = 0
NSTCIS_SELECTED = 1
NSTCRS_EXPANDED = 2
NSTCRS_HIDDEN = 1
NSTCRS_VISIBLE = 0
NSTCS_ALLOWJUNCTIONS = 268435456
NSTCS_AUTOHSCROLL = 1048576
NSTCS_BORDER = 32768
NSTCS_CHECKBOXES = 8388608
NSTCS_DIMMEDCHECKBOXES = 67108864
NSTCS_DISABLEDRAGDROP = 4096
NSTCS_EMPTYTEXT = 4194304
NSTCS_EVENHEIGHT = 1024
NSTCS_EXCLUSIONCHECKBOXES = 33554432
NSTCS_FADEINOUTEXPANDOS = 2097152
NSTCS_FAVORITESMODE = 524288
NSTCS_FULLROWSELECT = 8
NSTCS_HASEXPANDOS = 1
NSTCS_HASLINES = 2
NSTCS_HORIZONTALSCROLL = 32
NSTCS_NOEDITLABELS = 65536
NSTCS_NOINDENTCHECKS = 134217728
NSTCS_NOINFOTIP = 512
NSTCS_NOORDERSTREAM = 8192
NSTCS_NOREPLACEOPEN = 2048
NSTCS_PARTIALCHECKBOXES = 16777216
NSTCS_RICHTOOLTIP = 16384
NSTCS_ROOTHASEXPANDO = 64
NSTCS_SHOWDELETEBUTTON = 1073741824
NSTCS_SHOWREFRESHBUTTON = -2147483648
NSTCS_SHOWSELECTIONALWAYS = 128
NSTCS_SHOWTABSBUTTON = 536870912
NSTCS_SINGLECLICKEXPAND = 4
NSTCS_SPRINGEXPAND = 16
NSTCS_TABSTOP = 131072
NT_CONSOLE_PROPS_SIG = 2684354562
NT_FE_CONSOLE_PROPS_SIG = 2684354564
PIDASI_AVG_DATA_RATE = 4
PIDASI_CHANNEL_COUNT = 7
PIDASI_COMPRESSION = 10
PIDASI_FORMAT = 2
PIDASI_SAMPLE_RATE = 5
PIDASI_SAMPLE_SIZE = 6
PIDASI_STREAM_NAME = 9
PIDASI_STREAM_NUMBER = 8
PIDASI_TIMELENGTH = 3
PIDDI_THUMBNAIL = 2
PIDDRSI_DESCRIPTION = 3
PIDDRSI_PLAYCOUNT = 4
PIDDRSI_PLAYEXPIRES = 6
PIDDRSI_PLAYSTARTS = 5
PIDDRSI_PROTECTED = 2
PIDDSI_BYTECOUNT = 4
PIDDSI_CATEGORY = 2
PIDDSI_COMPANY = 15
PIDDSI_DOCPARTS = 13
PIDDSI_HEADINGPAIR = 12
PIDDSI_HIDDENCOUNT = 9
PIDDSI_LINECOUNT = 5
PIDDSI_LINKSDIRTY = 16
PIDDSI_MANAGER = 14
PIDDSI_MMCLIPCOUNT = 10
PIDDSI_NOTECOUNT = 8
PIDDSI_PARCOUNT = 6
PIDDSI_PRESFORMAT = 3
PIDDSI_SCALE = 11
PIDDSI_SLIDECOUNT = 7
PIDISF_CACHEDSTICKY = 2
PIDISF_CACHEIMAGES = 16
PIDISF_FOLLOWALLLINKS = 32
PIDISF_RECENTLYCHANGED = 1
PIDISM_DONTWATCH = 2
PIDISM_GLOBAL = 0
PIDISM_WATCH = 1
PIDMSI_COPYRIGHT = 11
PIDMSI_EDITOR = 2
PIDMSI_OWNER = 8
PIDMSI_PRODUCTION = 10
PIDMSI_PROJECT = 6
PIDMSI_RATING = 9
PIDMSI_SEQUENCE_NO = 5
PIDMSI_SOURCE = 4
PIDMSI_STATUS = 7
PIDMSI_SUPPLIER = 3
PIDSI_ALBUM = 4
PIDSI_APPNAME = 18
PIDSI_ARTIST = 2
PIDSI_AUTHOR = 4
PIDSI_CHARCOUNT = 16
PIDSI_COMMENT = 6
PIDSI_COMMENTS = 6
PIDSI_CREATE_DTM = 12
PIDSI_DOC_SECURITY = 19
PIDSI_EDITTIME = 10
PIDSI_GENRE = 11
PIDSI_KEYWORDS = 5
PIDSI_LASTAUTHOR = 8
PIDSI_LASTPRINTED = 11
PIDSI_LASTSAVE_DTM = 13
PIDSI_LYRICS = 12
PIDSI_PAGECOUNT = 14
PIDSI_REVNUMBER = 9
PIDSI_SONGTITLE = 3
PIDSI_SUBJECT = 3
PIDSI_TEMPLATE = 7
PIDSI_THUMBNAIL = 17
PIDSI_TITLE = 2
PIDSI_TRACK = 7
PIDSI_WORDCOUNT = 15
PIDSI_YEAR = 5
PIDVSI_COMPRESSION = 10
PIDVSI_DATA_RATE = 8
PIDVSI_FRAME_COUNT = 5
PIDVSI_FRAME_HEIGHT = 4
PIDVSI_FRAME_RATE = 6
PIDVSI_FRAME_WIDTH = 3
PIDVSI_SAMPLE_SIZE = 9
PIDVSI_STREAM_NAME = 2
PIDVSI_STREAM_NUMBER = 11
PIDVSI_TIMELENGTH = 7
PID_BEHAVIOR = -2147483645
PID_CODEPAGE = 1
PID_COMPUTERNAME = 5
PID_CONTROLPANEL_CATEGORY = 2
PID_DESCRIPTIONID = 2
PID_DICTIONARY = 0
PID_DISPLACED_DATE = 3
PID_DISPLACED_FROM = 2
PID_DISPLAY_PROPERTIES = 0
PID_FINDDATA = 0
PID_FIRST_NAME_DEFAULT = 4095
PID_FIRST_USABLE = 2
PID_HTMLINFOTIPFILE = 5
PID_ILLEGAL = -1
PID_INTROTEXT = 1
PID_INTSITE_AUTHOR = 3
PID_INTSITE_CODEPAGE = 18
PID_INTSITE_COMMENT = 8
PID_INTSITE_CONTENTCODE = 11
PID_INTSITE_CONTENTLEN = 10
PID_INTSITE_DESCRIPTION = 7
PID_INTSITE_FLAGS = 9
PID_INTSITE_LASTMOD = 5
PID_INTSITE_LASTVISIT = 4
PID_INTSITE_RECURSE = 12
PID_INTSITE_SUBSCRIPTION = 14
PID_INTSITE_TITLE = 16
PID_INTSITE_TRACKING = 19
PID_INTSITE_URL = 15
PID_INTSITE_VISITCOUNT = 6
PID_INTSITE_WATCH = 13
PID_INTSITE_WHATSNEW = 2
PID_IS_AUTHOR = 11
PID_IS_COMMENT = 13
PID_IS_DESCRIPTION = 12
PID_IS_HOTKEY = 6
PID_IS_ICONFILE = 9
PID_IS_ICONINDEX = 8
PID_IS_NAME = 4
PID_IS_SHOWCMD = 7
PID_IS_URL = 2
PID_IS_WHATSNEW = 10
PID_IS_WORKINGDIR = 5
PID_LINK_TARGET = 2
PID_LOCALE = -2147483648
PID_MAX_READONLY = -1073741825
PID_MIN_READONLY = -2147483648
PID_MISC_ACCESSCOUNT = 3
PID_MISC_OWNER = 4
PID_MISC_PICS = 6
PID_MISC_STATUS = 2
PID_MODIFY_TIME = -2147483647
PID_NETRESOURCE = 1
PID_NETWORKLOCATION = 4
PID_QUERY_RANK = 2
PID_SECURITY = -2147483646
PID_SHARE_CSC_STATUS = 2
PID_SYNC_COPY_IN = 2
PID_VOLUME_CAPACITY = 3
PID_VOLUME_FILESYSTEM = 4
PID_VOLUME_FREE = 2
PID_WHICHFOLDER = 3
PO_DELETE = 19
PO_PORTCHANGE = 32
PO_RENAME = 20
PO_REN_PORT = 52
PRINTACTION_DOCUMENTDEFAULTS = 6
PRINTACTION_NETINSTALL = 2
PRINTACTION_NETINSTALLLINK = 3
PRINTACTION_OPEN = 0
PRINTACTION_OPENNETPRN = 5
PRINTACTION_PROPERTIES = 1
PRINTACTION_SERVERPROPERTIES = 7
PRINTACTION_TESTPAGE = 4
PROPSETFLAG_ANSI = 2
PROPSETFLAG_CASE_SENSITIVE = 8
PROPSETFLAG_DEFAULT = 0
PROPSETFLAG_NONSIMPLE = 1
PROPSETFLAG_UNBUFFERED = 4
PROPSET_BEHAVIOR_CASE_SENSITIVE = 1
PROP_LG_CXDLG = 252
PROP_LG_CYDLG = 218
PROP_MED_CXDLG = 227
PROP_MED_CYDLG = 215
PROP_SM_CXDLG = 212
PROP_SM_CYDLG = 188
PRSPEC_INVALID = -1
PRSPEC_LPWSTR = 0
PRSPEC_PROPID = 1
PSBTN_APPLYNOW = 4
PSBTN_BACK = 0
PSBTN_CANCEL = 5
PSBTN_FINISH = 2
PSBTN_HELP = 6
PSBTN_MAX = 6
PSBTN_NEXT = 1
PSBTN_OK = 3
PSCB_BUTTONPRESSED = 3
PSCB_INITIALIZED = 1
PSCB_PRECREATE = 2
PSH_DEFAULT = 0
PSH_HASHELP = 512
PSH_HEADER = 524288
PSH_MODELESS = 1024
PSH_NOAPPLYNOW = 128
PSH_NOCONTEXTHELP = 33554432
PSH_PROPSHEETPAGE = 8
PSH_PROPTITLE = 1
PSH_RTLREADING = 2048
PSH_STRETCHWATERMARK = 262144
PSH_USECALLBACK = 256
PSH_USEHBMHEADER = 1048576
PSH_USEHBMWATERMARK = 65536
PSH_USEHICON = 2
PSH_USEHPLWATERMARK = 131072
PSH_USEICONID = 4
PSH_USEPAGELANG = 2097152
PSH_USEPSTARTPAGE = 64
PSH_WATERMARK = 32768
PSH_WIZARD = 32
PSH_WIZARD97 = 16777216
PSH_WIZARDCONTEXTHELP = 4096
PSH_WIZARDHASFINISH = 16
PSH_WIZARD_LITE = 4194304
PSNRET_INVALID = 1
PSNRET_INVALID_NOCHANGEPAGE = 2
PSNRET_MESSAGEHANDLED = 3
PSNRET_NOERROR = 0
PSPCB_ADDREF = 0
PSPCB_CREATE = 2
PSPCB_RELEASE = 1
PSP_DEFAULT = 0
PSP_DLGINDIRECT = 1
PSP_HASHELP = 32
PSP_HIDEHEADER = 2048
PSP_PREMATURE = 1024
PSP_RTLREADING = 16
PSP_USECALLBACK = 128
PSP_USEFUSIONCONTEXT = 16384
PSP_USEHEADERSUBTITLE = 8192
PSP_USEHEADERTITLE = 4096
PSP_USEHICON = 2
PSP_USEICONID = 4
PSP_USEREFPARENT = 64
PSP_USETITLE = 8
PSWIZB_BACK = 1
PSWIZB_DISABLEDFINISH = 8
PSWIZB_FINISH = 4
PSWIZB_NEXT = 2
QIF_CACHED = 1
QIF_DONTEXPANDFOLDER = 2
SBSP_ABSOLUTE = 0
SBSP_ALLOW_AUTONAVIGATE = 65536
SBSP_DEFBROWSER = 0
SBSP_DEFMODE = 0
SBSP_EXPLOREMODE = 32
SBSP_INITIATEDBYHLINKFRAME = -2147483648
SBSP_NAVIGATEBACK = 16384
SBSP_NAVIGATEFORWARD = 32768
SBSP_NEWBROWSER = 2
SBSP_NOAUTOSELECT = 67108864
SBSP_OPENMODE = 16
SBSP_PARENT = 8192
SBSP_REDIRECT = 1073741824
SBSP_RELATIVE = 4096
SBSP_SAMEBROWSER = 1
SBSP_WRITENOHISTORY = 134217728
SCHEME_CREATE = 128
SCHEME_DISPLAY = 1
SCHEME_DONOTUSE = 64
SCHEME_EDIT = 2
SCHEME_GLOBAL = 8
SCHEME_LOCAL = 4
SCHEME_REFRESH = 16
SCHEME_UPDATE = 32
SEE_MASK_ASYNCOK = 1048576
SEE_MASK_CLASSKEY = 3
SEE_MASK_CLASSNAME = 1
SEE_MASK_CONNECTNETDRV = 128
SEE_MASK_DOENVSUBST = 512
SEE_MASK_FLAG_DDEWAIT = 256
SEE_MASK_FLAG_NO_UI = 1024
SEE_MASK_HMONITOR = 2097152
SEE_MASK_HOTKEY = 32
SEE_MASK_ICON = 16
SEE_MASK_IDLIST = 4
SEE_MASK_INVOKEIDLIST = 12
SEE_MASK_NOCLOSEPROCESS = 64
SEE_MASK_NO_CONSOLE = 32768
SEE_MASK_UNICODE = 16384
SE_ERR_ACCESSDENIED = 5
SE_ERR_ASSOCINCOMPLETE = 27
SE_ERR_DDEBUSY = 30
SE_ERR_DDEFAIL = 29
SE_ERR_DDETIMEOUT = 28
SE_ERR_DLLNOTFOUND = 32
SE_ERR_FNF = 2
SE_ERR_NOASSOC = 31
SE_ERR_OOM = 8
SE_ERR_PNF = 3
SE_ERR_SHARE = 26
SFGAO_BROWSABLE = 134217728
SFGAO_CANCOPY = 1
SFGAO_CANDELETE = 32
SFGAO_CANLINK = 4
SFGAO_CANMONIKER = 4194304
SFGAO_CANMOVE = 2
SFGAO_CANRENAME = 16
SFGAO_CAPABILITYMASK = 375
SFGAO_COMPRESSED = 67108864
SFGAO_CONTENTSMASK = -2147483648
SFGAO_DISPLAYATTRMASK = 983040
SFGAO_DROPTARGET = 256
SFGAO_FILESYSANCESTOR = 268435456
SFGAO_FILESYSTEM = 1073741824
SFGAO_FOLDER = 536870912
SFGAO_GHOSTED = 524288
SFGAO_HASPROPSHEET = 64
SFGAO_HASSTORAGE = 4194304
SFGAO_HASSUBFOLDER = -2147483648
SFGAO_HIDDEN = 524288
SFGAO_LINK = 65536
SFGAO_NEWCONTENT = 2097152
SFGAO_NONENUMERATED = 1048576
SFGAO_READONLY = 262144
SFGAO_REMOVABLE = 33554432
SFGAO_SHARE = 131072
SFGAO_STORAGE = 8
SFGAO_STORAGEANCESTOR = 8388608
SFGAO_STORAGECAPMASK = 1891958792
SFGAO_STREAM = 4194304
SFGAO_VALIDATE = 16777216
SFVM_ADDOBJECT = 3
SFVM_GETSELECTEDOBJECTS = 9
SFVM_REARRANGE = 1
SFVM_REMOVEOBJECT = 6
SFVM_SETCLIPBOARD = 16
SFVM_SETITEMPOS = 14
SFVM_SETPOINTS = 23
SFVM_UPDATEOBJECT = 7
SHARD_APPIDINFO = 4
SHARD_APPIDINFOIDLIST = 5
SHARD_APPIDINFOLINK = 7
SHARD_LINK = 6
SHARD_PATH = 2
SHARD_PATHA = 2
SHARD_PATHW = 3
SHARD_PIDL = 1
SHARD_SHELLITEM = 8
SHCIDS_ALLFIELDS = -2147483648
SHCIDS_BITMASK = -65536
SHCIDS_CANONICALONLY = 268435456
SHCIDS_COLUMNMASK = 65535
SHCNEE_ORDERCHANGED = 2
SHCNE_ALLEVENTS = 2147483647
SHCNE_ASSOCCHANGED = 134217728
SHCNE_ATTRIBUTES = 2048
SHCNE_CREATE = 2
SHCNE_DELETE = 4
SHCNE_DISKEVENTS = 145439
SHCNE_DRIVEADD = 256
SHCNE_DRIVEADDGUI = 65536
SHCNE_DRIVEREMOVED = 128
SHCNE_EXTENDED_EVENT = 67108864
SHCNE_FREESPACE = 262144
SHCNE_GLOBALEVENTS = 201687520
SHCNE_INTERRUPT = -2147483648
SHCNE_MEDIAINSERTED = 32
SHCNE_MEDIAREMOVED = 64
SHCNE_MKDIR = 8
SHCNE_NETSHARE = 512
SHCNE_NETUNSHARE = 1024
SHCNE_RENAMEFOLDER = 131072
SHCNE_RENAMEITEM = 1
SHCNE_RMDIR = 16
SHCNE_SERVERDISCONNECT = 16384
SHCNE_UPDATEDIR = 4096
SHCNE_UPDATEIMAGE = 32768
SHCNE_UPDATEITEM = 8192
SHCNF_DWORD = 3
SHCNF_FLUSH = 4096
SHCNF_FLUSHNOWAIT = 8192
SHCNF_IDLIST = 0
SHCNF_PATH = 1
SHCNF_PATHA = 1
SHCNF_PATHW = 5
SHCNF_PRINTER = 2
SHCNF_PRINTERA = 2
SHCNF_PRINTERW = 6
SHCNF_TYPE = 255
SHCNRF_InterruptLevel = 1
SHCNRF_NewDelivery = 32768
SHCNRF_RecursiveInterrupt = 4096
SHCNRF_ShellLevel = 2
SHCOLSTATE_EXTENDED = 64
SHCOLSTATE_HIDDEN = 256
SHCOLSTATE_ONBYDEFAULT = 16
SHCOLSTATE_PREFER_VARCMP = 512
SHCOLSTATE_SECONDARYUI = 128
SHCOLSTATE_SLOW = 32
SHCOLSTATE_TYPEMASK = 15
SHCOLSTATE_TYPE_DATE = 3
SHCOLSTATE_TYPE_INT = 2
SHCOLSTATE_TYPE_STR = 1
SHCONTF_FOLDERS = 32
SHCONTF_INCLUDEHIDDEN = 128
SHCONTF_INIT_ON_FIRST_NEXT = 256
SHCONTF_NETPRINTERSRCH = 512
SHCONTF_NONFOLDERS = 64
SHCONTF_SHAREABLE = 1024
SHCONTF_STORAGE = 2048
SHDID_COMPUTER_CDROM = 10
SHDID_COMPUTER_DRIVE35 = 5
SHDID_COMPUTER_DRIVE525 = 6
SHDID_COMPUTER_FIXED = 8
SHDID_COMPUTER_NETDRIVE = 9
SHDID_COMPUTER_OTHER = 12
SHDID_COMPUTER_RAMDISK = 11
SHDID_COMPUTER_REMOVABLE = 7
SHDID_FS_DIRECTORY = 3
SHDID_FS_FILE = 2
SHDID_FS_OTHER = 4
SHDID_NET_DOMAIN = 13
SHDID_NET_OTHER = 17
SHDID_NET_RESTOFNET = 16
SHDID_NET_SERVER = 14
SHDID_NET_SHARE = 15
SHDID_ROOT_REGITEM = 1
SHERB_NOCONFIRMATION = 1
SHERB_NOPROGRESSUI = 2
SHERB_NOSOUND = 4
SHGDFIL_DESCRIPTIONID = 3
SHGDFIL_FINDDATA = 1
SHGDFIL_NETRESOURCE = 2
SHGDN_FORADDRESSBAR = 16384
SHGDN_FOREDITING = 4096
SHGDN_FORPARSING = 32768
SHGDN_INCLUDE_NONFILESYS = 8192
SHGDN_INFOLDER = 1
SHGDN_NORMAL = 0
SHGFI_ATTRIBUTES = 2048
SHGFI_ATTR_SPECIFIED = 131072
SHGFI_DISPLAYNAME = 512
SHGFI_EXETYPE = 8192
SHGFI_ICON = 256
SHGFI_ICONLOCATION = 4096
SHGFI_LARGEICON = 0
SHGFI_LINKOVERLAY = 32768
SHGFI_OPENICON = 2
SHGFI_PIDL = 8
SHGFI_SELECTED = 65536
SHGFI_SHELLICONSIZE = 4
SHGFI_SMALLICON = 1
SHGFI_SYSICONINDEX = 16384
SHGFI_TYPENAME = 1024
SHGFI_USEFILEATTRIBUTES = 16
SHGNLI_NOUNIQUE = 4
SHGNLI_PIDL = 1
SHGNLI_PREFIXNAME = 2
SHGVSPB_ALLFOLDERS = 8
SHGVSPB_ALLUSERS = 2
SHGVSPB_FOLDER = 5
SHGVSPB_FOLDERNODEFAULTS = 2147483653
SHGVSPB_GLOBALDEAFAULTS = 10
SHGVSPB_INHERIT = 16
SHGVSPB_NOAUTODEFAULTS = 2147483648
SHGVSPB_PERFOLDER = 4
SHGVSPB_PERUSER = 1
SHGVSPB_ROAM = 32
SHGVSPB_USERDEFAULTS = 9
SIATTRIBFLAGS_AND = 1
SIATTRIBFLAGS_APPCOMPAT = 3
SIATTRIBFLAGS_MASK = 3
SIATTRIBFLAGS_OR = 2
SICHINT_ALLFIELDS = -2147483648
SICHINT_CANONICAL = 268435456
SICHINT_DISPLAY = 0
SIGDN_DESKTOPABSOLUTEEDITING = -2147172352
SIGDN_DESKTOPABSOLUTEPARSING = -2147319808
SIGDN_FILESYSPATH = -2147123200
SIGDN_NORMALDISPLAY = 0
SIGDN_PARENTRELATIVE = -2146959359
SIGDN_PARENTRELATIVEEDITING = -2147282943
SIGDN_PARENTRELATIVEFORADDRESSBAR = -2146975743
SIGDN_PARENTRELATIVEPARSING = -2147385343
SIGDN_URL = -2147057664
SLDF_FORCE_NO_LINKINFO = 256
SLDF_FORCE_UNCNAME = 65536
SLDF_HAS_ARGS = 32
SLDF_HAS_DARWINID = 4096
SLDF_HAS_EXP_ICON_SZ = 16384
SLDF_HAS_EXP_SZ = 512
SLDF_HAS_ICONLOCATION = 64
SLDF_HAS_ID_LIST = 1
SLDF_HAS_LINK_INFO = 2
SLDF_HAS_LOGO3ID = 2048
SLDF_HAS_NAME = 4
SLDF_HAS_RELPATH = 8
SLDF_HAS_WORKINGDIR = 16
SLDF_NO_PIDL_ALIAS = 32768
SLDF_RESERVED = 2147483648
SLDF_RUNAS_USER = 8192
SLDF_RUN_IN_SEPARATE = 1024
SLDF_RUN_WITH_SHIMLAYER = 131072
SLDF_UNICODE = 128
SSF_DESKTOPHTML = 512
SSF_DONTPRETTYPATH = 2048
SSF_DOUBLECLICKINWEBVIEW = 128
SSF_HIDEICONS = 16384
SSF_MAPNETDRVBUTTON = 4096
SSF_NOCONFIRMRECYCLE = 32768
SSF_SHOWALLOBJECTS = 1
SSF_SHOWATTRIBCOL = 256
SSF_SHOWCOMPCOLOR = 8
SSF_SHOWEXTENSIONS = 2
SSF_SHOWINFOTIP = 8192
SSF_SHOWSYSFILES = 32
SSF_WIN95CLASSIC = 1024
SSM_CLEAR = 0
SSM_REFRESH = 2
SSM_SET = 1
SSM_UPDATE = 4
STRRET_CSTR = 2
STRRET_OFFSET = 1
STRRET_WSTR = 0
STR_AVOID_DRIVE_RESTRICTION_POLICY = "Avoid Drive Restriction Policy"
STR_BIND_DELEGATE_CREATE_OBJECT = "Delegate Object Creation"
STR_BIND_FOLDERS_READ_ONLY = "Folders As Read Only"
STR_BIND_FOLDER_ENUM_MODE = "Folder Enum Mode"
STR_BIND_FORCE_FOLDER_SHORTCUT_RESOLVE = "Force Folder Shortcut Resolve"
STR_DONT_PARSE_RELATIVE = "Don\'t Parse Relative"
STR_DONT_RESOLVE_LINK = "Don\'t Resolve Link"
STR_FILE_SYS_BIND_DATA = "File System Bind Data"
STR_GET_ASYNC_HANDLER = "GetAsyncHandler"
STR_GPS_BESTEFFORT = "GPS_BESTEFFORT"
STR_GPS_DELAYCREATION = "GPS_DELAYCREATION"
STR_GPS_FASTPROPERTIESONLY = "GPS_FASTPROPERTIESONLY"
STR_GPS_HANDLERPROPERTIESONLY = "GPS_HANDLERPROPERTIESONLY"
STR_GPS_NO_OPLOCK = "GPS_NO_OPLOCK"
STR_GPS_OPENSLOWITEM = "GPS_OPENSLOWITEM"
STR_IFILTER_FORCE_TEXT_FILTER_FALLBACK = "Always bind persistent handlers"
STR_IFILTER_LOAD_DEFINED_FILTER = "Only bind registered persistent handlers"
STR_INTERNAL_NAVIGATE = "Internal Navigation"
STR_INTERNETFOLDER_PARSE_ONLY_URLMON_BINDABLE = "Validate URL"
STR_ITEM_CACHE_CONTEXT = "ItemCacheContext"
STR_NO_VALIDATE_FILENAME_CHARS = "NoValidateFilenameChars"
STR_PARSE_ALLOW_INTERNET_SHELL_FOLDERS = "Allow binding to Internet shell folder handlers and negate STR_PARSE_PREFER_WEB_BROWSING"
STR_PARSE_AND_CREATE_ITEM = "ParseAndCreateItem"
STR_PARSE_DONT_REQUIRE_VALIDATED_URLS = "Do not require validated URLs"
STR_PARSE_EXPLICIT_ASSOCIATION_SUCCESSFUL = "ExplicitAssociationSuccessful"
STR_PARSE_PARTIAL_IDLIST = "ParseOriginalItem"
STR_PARSE_PREFER_FOLDER_BROWSING = "Parse Prefer Folder Browsing"
STR_PARSE_PREFER_WEB_BROWSING = "Do not bind to Internet shell folder handlers"
STR_PARSE_PROPERTYSTORE = "DelegateNamedProperties"
STR_PARSE_SHELL_PROTOCOL_TO_FILE_OBJECTS = "Parse Shell Protocol To File Objects"
STR_PARSE_SHOW_NET_DIAGNOSTICS_UI = "Show network diagnostics UI"
STR_PARSE_SKIP_NET_CACHE = "Skip Net Resource Cache"
STR_PARSE_TRANSLATE_ALIASES = "Parse Translate Aliases"
STR_PARSE_WITH_EXPLICIT_ASSOCAPP = "ExplicitAssociationApp"
STR_PARSE_WITH_EXPLICIT_PROGID = "ExplicitProgid"
STR_PARSE_WITH_PROPERTIES = "ParseWithProperties"
STR_SKIP_BINDING_CLSID = "Skip Binding CLSID"
STR_TRACK_CLSID = "Track the CLSID"
SVGIO_ALLVIEW = 2
SVGIO_BACKGROUND = 0
SVGIO_CHECKED = 3
SVGIO_FLAG_VIEWORDER = -2147483648
SVGIO_SELECTION = 1
SVGIO_TYPE_MASK = 15
SVSI_DESELECT = 0
SVSI_DESELECTOTHERS = 4
SVSI_EDIT = 3
SVSI_ENSUREVISIBLE = 8
SVSI_FOCUSED = 16
SVSI_SELECT = 1
SVSI_TRANSLATEPT = 32
SVUIA_ACTIVATE_FOCUS = 2
SVUIA_ACTIVATE_NOFOCUS = 1
SVUIA_DEACTIVATE = 0
SVUIA_INPLACEACTIVATE = 3
WIZ_BODYCX = 184
WIZ_BODYX = 92
WIZ_CXBMP = 80
WIZ_CXDLG = 276
WIZ_CYDLG = 140
WM_USER = 1024
WPSTYLE_CENTER = 0
WPSTYLE_MAX = 3
WPSTYLE_STRETCH = 2
WPSTYLE_TILE = 1
