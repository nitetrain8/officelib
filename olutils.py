"""
Created on Dec 2, 2013

@author: PBS Biotech
"""

from os import name as os_name, listdir as _listdir
import ctypes.wintypes

from os.path import normpath as _normpath, exists as _exists, expanduser as _expanduser, \
    splitext as _splitext, split as _split, splitdrive as _splitdrive


# noinspection PyUnusedLocal
def __v_print_none(*args, **kwargs):
    pass


def __v_print(*args, **kwargs):
    print(*args, sep='', **kwargs)

v_print = __v_print_none


def echo_on():
    global v_print
    v_print = __v_print


def echo_off():
    global v_print
    v_print = __v_print_none


class OfficeLibError(Exception):
    """Base Exception for Officelib Errors"""
    pass


CSIDL_PERSONAL = 5  # My Docs
CSIDL_COMMON_DOCUMENTS = 46
SHGFP_TYPE_CURRENT = 0  # Current value, not default value


def getWorkDir():

    """Get current user's Documents folder
    @rtype: str
    """

    # OS-specific attempt for Windows
    if os_name == 'nt':

        try:
            return getWinUserDocs()
        except OSError:
            pass

        try:
            return getWinCommonDocs()
        except OSError:
            pass

    user = _expanduser("~")

    docs = 'Documents'
    mydocs = 'My Documents'
    folders = (docs, mydocs)

    for folder in folders:
        workdir = ''.join((user, folder))
        if _exists(workdir):
            return workdir.replace('/', '\\')

    return None


def getWinUserDocs():
    """C/P from
    http://stackoverflow.com/questions/3858851/python-get-windows-special-folders-for-currently-logged-in-user#3859336

    Convenience function to get current user's documents folder.

    @return: user's docs folder.
    @rtype: str
    """

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)

    hresult = ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_CURRENT, buf)
    if hresult != 0:  # SHGetFolderPathW returned error
        raise OSError("Failed to find user's Documents folder")

    return _normpath(buf.value)


def getWinCommonDocs():
    """Convenience function to return the windows common docs folder.
    Same source as above.

    @return: common docs folder, or raise OSError.
    @rtype: str
    """

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    hresult = ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_COMMON_DOCUMENTS, 0, SHGFP_TYPE_CURRENT, buf)
    if hresult != 0:
        raise OSError("Failed to find common Documents folder")

    return _normpath(buf.value)


def getDownloadDir():
    """
    @return: filepath of download dir
    @rtype: str

    Todo- figure out a non stupid way to do this. There should be a
    special OS folder designated as default folder.

    Update:
    http://msdn.microsoft.com/en-us/library/windows/desktop/bb762188(v=vs.85).aspx

    Use SHGetKnownFolderPath with the correct GUID
    issue- GUID is written as a string, but needs to be sent to
    function as a struct(?)
    need to make custom c structure.
    """
    try:
        user = _expanduser("~")
    except:
        # Todo- figure out how to find dl folder on mac?
        raise

    dl_dir = '\\'.join([user, "Downloads"])
    if not _exists(dl_dir):
        raise FileNotFoundError("Couldn't find downloads folder")

    return dl_dir.replace('/', '\\')


from weakref import WeakValueDictionary, ref


# noinspection PyAttributeOutsideInit
class SingletonWrapper(type):
    """Metaclass that keeps track of its instances"""
    __instances__ = WeakValueDictionary()

    def __call__(cls, *args):

        #  First check class to see if it implemented its own self_ reference
        try:
            self = cls._selfref()
            if self is not None and self.__class__ is cls:
                return self
        except AttributeError:
            pass

        # Self reference not found in class, check metaclass
        try:
            return SingletonWrapper.__instances__[cls]
        except KeyError:
            self = type.__call__(cls, *args)
            cls._selfref = ref(self)
            SingletonWrapper.__instances__[cls] = self
            return self


class Singleton(metaclass=SingletonWrapper):
    """ Allow singleton by subclassing instead of metaclassing.
        Include selfref- will be weakref.ref
    """
    _selfref = lambda: None
    pass


def __dir_scan_nosplit(filename, directory, listdir=_listdir):
    """
    Internal recursive helper function for _get_lib_path_no_basename,
    Identical to above, but doesn't split the extension off of filepaths from listdir.

    @param filename: str
    @type filename: str
    @param directory: str
    @type directory: str
    @return: filename, or none
    @rtype: str | None
    """
    try:
        dirs = listdir(directory)
    except OSError:
        # Catch NotADirectoryError, PermissionError
        return None
    for path in dirs:
        if filename == path:
            return '\\'.join((directory, path))
        result = __dir_scan_nosplit(filename, '\\'.join((directory, path)))
        if result:
            return result
    return None


def _get_lib_path_parital_qualname(name, base, search_dirs, splitext=_splitext):
    """
    Internal function to search for partially qualified names
    @param name: filename
    @type name: str
    @param base: parially qualified base to join with name
    @type base: str
    @param search_dirs: search dirs
    @type search_dirs: list[str] or tuple[str]
    @return: str
    @rtype: str
    """
    ext = splitext(name)[1]
    if ext:
        for fldr in search_dirs:
            path = fldr + base
            for file in _listdir(path):
                if file == name:
                    return '\\'.join((path, name))

    else:
        for fldr in search_dirs:
            path = fldr + base
            for file in _listdir(path):
                file, ext = splitext(file)
                # test presence of ext to exclude dirs
                if ext and file == name:
                    return ''.join((path, "\\", name, ext))

    raise FileNotFoundError("Partially qualified name %s not found" % '\\'.join((base, name)))


def _get_lib_path_no_basename(filename, search_dirs, scanner=__dir_scan_nosplit):
    """Internal function to find a lib path given filename with extension.

    Iterate through directory, if matching filename found, return it.

    @param filename: a filename with extension but no basename
    @type filename: str
    @param search_dirs: list of dirs to search
    @type search_dirs: list[str]
    @return: first valid, existing filename found in directory.
    @rtype: str
    """
    for fldr in search_dirs:
        result = scanner(filename, fldr)
        if result:
            return result
    raise FileNotFoundError("File '%s' not found via no-basename search:" % filename)


def _get_lib_path_no_extension(filepath, splitext=_splitext):
    """Internal function to find file when given basepath but
    no file extension.

    Scan the directory of the file and try to match the head of
    the filepath to an entry in that directory.

    @param filepath: a filename with extension but no basename
            To make function not require xl param, just use
            xl = newExcel(False, False) and close at the end.
    @type filepath: str
    @rtype: str
    """
    base, head = _split(filepath)
    for entry in _listdir(base):
        filename, ext = splitext(entry)
        # if not ext, then we have a dir, not filename
        # test ext first to short circuit non-files
        if ext and filename == head:
            return '\\'.join((base, entry))
    raise FileNotFoundError("File '%s' not found via no-extension search" % filepath)


def __dir_scan(filename, directory, listdir=_listdir, splitext=_splitext):
    """
    Internal recursive helper function for _get_lib_path_no_ctxt

    @param filename: str
    @type filename: str
    @param directory: str
    @type directory: str
    @return: filename, or none
    @rtype: str | None
    """
    try:
        dirs = listdir(directory)
    except OSError:
        # Catch NotADirectoryError, and PermissionError
        return None

    for path in dirs:
        name, ext = splitext(path)

        # check ext -> if empty, we have a directory
        if ext and (name == filename):
            return '\\'.join((directory, path))

        result = __dir_scan(filename, '\\'.join((directory, path)))
        if result:
            return result
    return None


def _get_lib_path_no_ctxt(filename, search_dirs, scanner=__dir_scan):
    """Internal function to find file when given neither basepath
    nor extension (no context)

    @param filename: a filename with extension but no basename
    @type filename: str
    @type search_dirs: list[str]
    @rtype: str
    """

    for fldr in search_dirs:
        result = scanner(filename, fldr)
        if result:
            return result
    raise FileNotFoundError("File %s not found via no-context search." % filename)


def __get_work_dirs():
    """
    @return: Set of folders corresponding to user work dirs (my docs, downloads...)
    @rtype: set
    """
    wrk = set()

    get_funcs = (
                getWinUserDocs,
                getWinCommonDocs,
                getDownloadDir
                )

    for func in get_funcs:
        try:
            fldr = func()
        except:
            pass
        else:
            wrk.add(fldr)

    return {f.replace('/', '\\') for f in wrk}


def _lib_path_search_dir_list_builder(folder_hint=None, *folder_hints, workdirs=__get_work_dirs()):
    """Helper function to build the list of folders in which
    to search for getFullFilename() function.

    @param folder_hint: folder to look in
    @type folder_hint: str | None
    @type folder_hints: list[str]

    @return: list of folders to search
    """

    fldrs = []
    if folder_hint and folder_hint not in workdirs:
        fldrs.append(folder_hint.replace('/', '\\'))

    for folder in folder_hints:
        if folder not in workdirs:
            fldrs.append(folder.replace('/', '\\'))

    fldrs.extend(workdirs)

    return fldrs


def getFullFilename(path, hint=None):
    """ Function to get full library path. Figure out
    what's in the path iteratively, based on 3 common scenarios.

    @param path: a filepath or filename
    @type path: str
    @param hint: the first directory tree in which to search for the file
    @type hint: str
    @return: full library path to existing file.
    @rtype: str

    Try to find the path by checking for three common cases:

    1. filename with extension
        - base
        - no base
    2. filename with only base
        2.1 partially qualified directory name
        2.2 fully qualified directory name
    3. neither one

    Build list of folders to search by calling first helper function.

    Update 1/16/2014- xl nonsense gone

    I moved the algorithm for executing the search to an inlined dispatch function
    that receives all the relative args, for the sake of making this function
    cleaner, but I'm not sure if that level of indirection just makes
    everything even worse. Having it defined within this function allows
    it to access path, etc variables without having to explicitly call them.
    In all, there is much less text in the areas in which the dispatch is called.
    """

    path = path.replace('/', '\\')  # Normalize sep type

    # Was path already good?
    if _exists(path):
        return path

    # Begin process of finding file 
    search_dirs = _lib_path_search_dir_list_builder(hint)
    base, name = _split(path)
    ext = _splitext(name)[1]

    # Most likely- given extension.
    # no need to check for case of fully qualified basename.
    # an existing file with a fully qualified base name and extension would
    # be caught by earlier _exists()
    if ext:
        if base:
            v_print('\nPartially qualified filename \'', path, "\' given, searching for file...")
            return _get_lib_path_parital_qualname(name, base, search_dirs)

        # else
        v_print("\nNo directory given for \'", path, "\', scanning for file...")
        return _get_lib_path_no_basename(path, search_dirs)

    # Next, given filename with base, but no extension
    elif base:

        drive, _tail = _splitdrive(base)

        # fully qualified base, just check the dir for matching name
        if drive:
            v_print("\nNo file extension given for \'", path, "\', scanning for file...")
            return _get_lib_path_no_extension(path)

        # partially qualified base, search dirs. I don't think this works well (at all).
        # I don't think I managed to get a working unittest for it.
        else:
            v_print("\nAttempting to find partially qualified name \'", path, "\' ...")
            return _get_lib_path_parital_qualname(name, base, search_dirs)

    # Finally, user gave no context- no base or extension. 
    # Try really hard to find it anyway.
    else:
        v_print("\nNo context given for filename, scanning for file.\nIf you give a full filepath, you wouldn't \nhave to wait for the long search.")
        return _get_lib_path_no_ctxt(path, search_dirs)

    # noinspection PyUnreachableCode
    raise SystemExit("Unreachable code reached: fix module olutils")


def ListFullDir(dirname):
    """Sometimes it is inconvenient to have to
    type '\\'.join(basename, filename) when calling listdir
    to get the full path, so here's a shortcut func.

    """
    return ['\\'.join((dirname, filename)) for filename in _listdir(dirname)]


# from tkinter import Tk, Frame, E as tkE, W as tkW, StringVar, Toplevel
# from tkinter.ttk import Label, Button, Entry, LabelFrame
#
#
# class SimplePrompt():
#     """ Simple helper class to open a window
#     with two buttons (ok and cancel), and a text entry
#
#     """
#
#     def __init__(self):
#         self.display_text = None
#         self._callback = lambda _: None
#         self.root = None
#         self.frame = None
#         self._complete = False
#         self.text_var = None
#         self._return_text = None
#
#     def ask(self, display_text='SimplePrompt', result_callback=lambda _: None):
#         """
#         @param display_text: the text to display on the labelframe.
#         @return: the text entered by user, or None if dialog was canceled.
#
#         """
#         self.display_text = display_text
#         self._callback = result_callback
#         self._complete = False
#
#         root = Toplevel()
#         frame = LabelFrame(root, text=self.display_text)
#
#         ok_btn = Button(frame, text="Ok", command=self.ok_action)
#         cancel_btn = Button(frame, text="Cancel", command=self.cancel_action)
#
#         text_var = StringVar()
#         text_entry = Entry(frame, width=40, textvariable=text_var)
#
#         frame.bind_all("<Key>", self.key_event_handler)
#
#         text_entry.grid(columnspan=4, row=0)
#         ok_btn.grid(column=2, row=1, sticky=tkW)
#         cancel_btn.grid(column=1, row=1, sticky=tkE)
#
#         self.root = root
#         self.frame = frame
#         self.text_var = text_var
#
#         frame.bind("<Destroy>", self.destroy_action)
#         text_entry.focus()
#         self.frame.grid()
#         self.root.grab_set()
#         self.root.wait_window(self.root)
#         return self._return_text
#
#     def key_event_handler(self, event):
#
#         ENTER = 13
#         ESC = 27
#         keycode = event.keycode
#
#         if keycode == ENTER:
#             self.ok_action()
#
#         elif keycode == ESC:
#             self.cancel_action()
#
#     # noinspection PyUnusedLocal
#     def destroy_action(self, _event):
#         if self._complete:
#             self._callback(self._return_text)
#
#     def ok_action(self):
#         text = self.text_var.get()
#         if text:
#             self._return_text = text
#             self._complete = True
#             self.root.destroy()
#         else:
#             self.root.bell()
#
#     def cancel_action(self):
#         self.root.destroy()


if __name__ == '__main__':

    #Todo: more

#     a = SimplePrompt()
#     print(a.ask("HelloWorld!"))
#
#     filename1 = r"C:\Users\Public\Documents\PBSSS\PBS 3 RTD cal template.xlsx"
#     filename2 = r"mytest.xlsx"
#     filename3 = r"C:\Users\Public\Documents\PBSSS\PBS 3 RTD cal template"
#     filename4 = r"3L"
#     from officelib import xllib
#     xl = xllib.Excel(visible=True)
#     xlFolder = xl.DefaultFilePath
#     try:
#         print(getFullFilename('mytest', xlFolder))
#         print(getFullFilename(filename1, xlFolder))
#         print(getFullFilename(filename2, xlFolder))
#         print(getFullFilename(filename3, xlFolder))
#         print(getFullFilename(filename4))
#     finally:
#         xl.Quit()
#
#
    mytuple = list(range(1, 6))
    print(mytuple)
    print(list(reversed(mytuple)))
    print(len(mytuple) - 1 - next(i for i, v in enumerate(reversed(mytuple)) if v < 3))
    
    
    
    
