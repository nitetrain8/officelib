"""
Created on Dec 2, 2013

@author: PBS Biotech
"""

import os.path
import ctypes.wintypes

#wintypes const values
CSIDL_PERSONAL = 5  # My Docs
CSIDL_COMMON_DOCUMENTS = 46
SHGFP_TYPE_CURRENT = 0  # Current value, not default value


def getWorkDir():

    """Get current user's Documents folder
    @rtype: str
    """

    # OS-specific attempt for Windows
    if os.name == 'nt':

        try:
            return getWinUserDocs()
        except OSError:
            pass

        try:
            return getWinCommonDocs()
        except OSError:
            pass

    user = os.path.expanduser("~")

    docs = 'Documents'
    mydocs = 'My Documents'
    folders = (docs, mydocs)

    for folder in folders:
        workdir = ''.join((user, folder))
        if os.path.exists(workdir):
            return workdir

    return None


def getWinUserDocs():
    """C/P from
    http://stackoverflow.com/questions/3858851/python-get-windows-special-folders-for-currently-logged-in-user#3859336

    Convenience function to get current user's documents folder.

    @return: user's docs folder.
    @rtype: str
    """

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)

    bSuccess = ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_PERSONAL, 0, SHGFP_TYPE_CURRENT, buf)
    if bSuccess != 0:  # SHGetFolderPathW returned error
        raise OSError("Failed to find user's Documents folder")

    return buf.value


def getWinCommonDocs():
    """Convenience function to return the windows common docs folder.
    Same source as above.

    @return: common docs folder, or raise OSError.
    @rtype: str
    """

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    query = ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_COMMON_DOCUMENTS, 0, SHGFP_TYPE_CURRENT, buf)
    if query != 0:
        raise OSError("Failed to find common Documents folder")

    return buf.value


def getDownloadDir():
    """
    @return: filepath of download dir
    @rtype: str

    Todo- figure out a non stupid way to do this. There should be a
    special OS folder designated as default folder.
    """
    try:
        user = os.path.expanduser("~")
    except:
        # Todo- figure out how to find dl folder on mac?
        raise

    dl_dir = '\\'.join([user, "Downloads"])
    if not os.path.exists(dl_dir):
        raise FileNotFoundError("Couldn't find downloads folder")

    return dl_dir


def getFileExtension(filepath, split=os.path.splitext):
    """
    @param filepath: filepath
    @type filepath: str
    @return: extension or None (or is it ''?)
    @rtype: str | None
    """
    base, ext = split(filepath)
    if base and not ext:
        if base[0] == '.' and '\\' not in base and '/' not in base:
            return base
        else:
            return ''
    else:
        return ext


def getUniqueName(filename):
    """ Give a base filename plus extension -> get unique filename

    @param filename- the base filename minus extension
    @type filename: str
    """

    root, extension = os.path.splitext(filename)

    exists = os.path.exists

    unique_name = filename
    exists_counter = 0
    while exists(unique_name):
        exists_counter += 1
        unique_name = ''.join((
                            root,
                            "(",
                            str(exists_counter),
                            ")",
                            extension
                            ))

    return unique_name

from weakref import WeakValueDictionary, ref


# noinspection PyAttributeOutsideInit
class SingletonWrapper(type):
    """Metaclass that keeps track of its instances"""
    __instances__ = WeakValueDictionary()

    def __call__(cls, *args):

        #  First check class to see if it implemented its own self_ reference
        try:
            self_ = cls._selfref()
            if self_ is not None and self_.__class__ is cls:
                return self_
        except AttributeError:
            pass

        # Self reference not found in class, check metaclass
        try:
            return SingletonWrapper.__instances__[cls]
        except KeyError:
            self_ = type.__call__(cls, *args)
            cls._selfref = ref(self_)
            SingletonWrapper.__instances__[cls] = self_
            return self_


class Singleton(metaclass=SingletonWrapper):
    """ Allow singleton by subclassing instead of metaclassing.
        Include selfref- will be weakref.ref

    """
    _selfref = lambda: None
    pass


def _get_lib_path_no_basename(filename, target_folder):
    """Internal function to find a lib path given filename with extension.
    Search directory currently only targets Excel.Application.DefaultFilePath
    aka library.

    Iterate through directory, if matching filename found, return it.

    @param filename: a filename with extension but no basename
    @param target_folder: an active excel instance to use to query for
                DefaultFilePath, to avoid needing to open a new
                one.

            To make function not require xl param, just use
            xl = newExcel(False, False) and close at the end.
    @type filename: str
    @type target_folder: str

    @return: first valid, existing filename found in directory.
    @rtype: str
    """

    for dirpath, _dirnames, filenames in os.walk(target_folder):
        if filename in filenames:
            return "\\".join((dirpath, filename))

    raise NameError('''Couldn't find file %s in user's default library
    after scanning all files in %s.''' % (filename, target_folder))


def _get_lib_path_no_extension(filepath, splitext=os.path.splitext):
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
    base, head = os.path.split(filepath)

    for entry in os.listdir(base):

        filename, ext = splitext(entry)

        if ext and filename == head:  # if not ext, then we have a dir, not filename
            return '\\'.join((base, entry))

    raise NameError("No file extension given.\nUnable guess proper extension for file %s" % filepath)


def _get_lib_path_no_ctxt(filename, target_folder, listdir=os.listdir, splitext=os.path.splitext):
    """Internal function to find file when given neither basepath
    nor extension (no context)

    @param filename: a filename with extension but no basename
    @type filename: str
    @type target_folder: str
    @rtype: str

    """
    def dir_scan(directory, filename=filename, listdir=listdir, splitext=splitext):
        for entry in listdir(directory):

            name, ext = splitext(entry)

            # check ext -> if empty, we have a directory, move on
            # had problems with os.path.isdir in the past
            if ext and (name == filename):
                return '\\'.join((directory, entry))

            # catch errors from os.listdir (NotADirectoryError)
            # as well as own own at the end of the loop,
            # which allows us to unwind the current stack. 
            try:
                return dir_scan('\\'.join((directory, entry)))
            except:
                pass

        raise StopIteration

    try:
        return dir_scan(target_folder)
    except:
        raise NameError("Couldn't find file \'%s\' in any library folder.\nEnter a valid file path with extension" % filename)


def _lib_path_search_dir_list_builder(folder_hint=None, *folder_hints):
    """Helper function to build the list of folders in which
    to search for getFullLibraryPath() function.

    @param folder_hint(s): folder(s) to include in search.
                            pass in a list of strings.
    @type folder_hint: str
    @type folder_hints: list[str]

    @return: list of folders to search
    """

    folders = []

    if folder_hint:
        folders.append(folder_hint)

    if folder_hints:
        folders.extend(folder_hints)

    # These next functions throw error on failure to return,
    # so capture in individual try blocks.
    # also make sure they're not already in there 

    try:
        user_docs = getWinUserDocs()
        if user_docs not in folders:
            folders.append(user_docs)
    except:
        pass

    try:
        common_docs = getWinCommonDocs()
        if common_docs not in folders:
            folders.append(common_docs)
    except:
        pass

    try:
        dl_folder = getDownloadDir()
        if dl_folder not in folders:
            folders.append(dl_folder)
    except:
        pass

    return folders


def getFullLibraryPath(path, hint=None, *, verbose=True):
    """ Function to get full library path. Figure out
    what's in the path iteratively, based on 3 common scenarios.

    @param path: a filepath or filename
    @type path: str
    @param hint: the first directory tree in which to search for the file
    @type hint: str
    @return: full library path to existing file.
    @rtype: str

    Try to find the path by checking for three common cases:

    1. filename with extension but no base.
    2. filename with base but no extension
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

    if not path:
        raise NameError("Enter a valid filepath")

    # Was path already good?
    if os.path.exists(path):
        return path

    if verbose:
        v_print = print
    else:
        v_print = lambda *_: None

    path = path.replace('/', '\\')  # Normalize sep type

    # Begin process of finding file 
    search_dirs = _lib_path_search_dir_list_builder(hint)
    basename, filename = os.path.split(path)
    ext = os.path.splitext(filename)[1]

    #helper function for cases 1, 2.1, 3.
    def do_dispatch(search_func, path=path, v_print=v_print):
        for directory in search_dirs:
            v_print("Searching %s" % directory)
            try:
                return search_func(path, directory)
            except:
                pass

        else:
            err_msg = '\n'.join(("Couldn't find \'%s\' in the following places:\n" % path,
                                '\n'.join(search_dirs)))
            raise NameError(err_msg)

    # Most likely- given a filename with no base, but with extension
    if (not basename) and ext:
        v_print("\nNo directory given for \'%s\', scanning for file..." % path)
        return do_dispatch(_get_lib_path_no_basename)

    # Next, given filename with base, but no extension
    elif basename and (not ext):

        drive, _tail = os.path.splitdrive(basename)

        if not drive:  # partially qualified base, search dirs
            v_print("\nAttempting to find partially qualified name \'%s\' ..." % path)
            return do_dispatch(_get_lib_path_no_basename)

        else:  # fully qualified base, just check the dir for matching name
            v_print("No file extension given for \'%s\', scanning for file..." % path)
            return _get_lib_path_no_extension(path)

    # Finally, user gave no context- no base or extension. 
    # Try really hard to find it anyway.
    elif (not basename) and (not ext):
        v_print("\nNo context given for filename, scanning for file.\nIf you give a full filepath, you wouldn't \nhave to wait for the long search.")
        return do_dispatch(_get_lib_path_no_ctxt)

    else:
        raise NameError("Unable to find file %s" % path)


def ListFullDir(dirname, os_listdir=os.listdir):
    """Sometimes it is inconvenient to have to
    type '\\'.join(basename, filename) when calling os.listdir
    to get the full path, so here's a shortcut func.

    """
    return ['\\'.join((dirname, filename)) for filename in os_listdir(dirname)]


from tkinter import Tk, Frame, E as tkE, W as tkW, StringVar, Toplevel
from tkinter.ttk import Label, Button, Entry, LabelFrame


def MsgBox(msg='', *args):
    """Imitate OpenOffice Basic msgbox command.
    OO msgbox is a simple dialog communication thing."""
    if args:
        msg = ' '.join((msg, ' '.join(str(arg) for arg in args)))
    root = Tk()
    frame = Frame(root)
    root.resizable(width=False, height=False)
    okbtn = Button(frame, text="Ok", command=root.destroy)
    txt = Label(frame)
    frame.grid()
    txt.grid(sticky=(tkE, tkW))
    okbtn.grid()
    txt.config(text=msg)
    root.mainloop()


class SimplePrompt():
    """ Simple helper class to open a window
    with two buttons (ok and cancel), and a text entry

    """

    def __init__(self):
        self.display_text = None
        self._callback = lambda _: None
        self.root = None
        self.frame = None
        self._complete = False
        self.text_var = None
        self._return_text = None

    def ask(self, display_text='SimplePrompt', result_callback=lambda _: None):
        """
        @param display_text: the text to display on the labelframe.
        @return: the text entered by user, or None if dialog was canceled.

        """
        self.display_text = display_text
        self._callback = result_callback
        self._complete = False

        root = Toplevel()
        frame = LabelFrame(root, text=self.display_text)

        ok_btn = Button(frame, text="Ok", command=self.ok_action)
        cancel_btn = Button(frame, text="Cancel", command=self.cancel_action)

        text_var = StringVar()
        text_entry = Entry(frame, width=40, textvariable=text_var)

        frame.bind_all("<Key>", self.key_event_handler)

        text_entry.grid(columnspan=4, row=0)
        ok_btn.grid(column=2, row=1, sticky=tkW)
        cancel_btn.grid(column=1, row=1, sticky=tkE)

        self.root = root
        self.frame = frame
        self.text_var = text_var

        frame.bind("<Destroy>", self.destroy_action)
        text_entry.focus()
        self.frame.grid()
        self.root.grab_set()
        self.root.wait_window(self.root)
        return self._return_text

    def key_event_handler(self, event):

        ENTER = 13
        ESC = 27
        keycode = event.keycode

        if keycode == ENTER:
            self.ok_action()

        elif keycode == ESC:
            self.cancel_action()

    # noinspection PyUnusedLocal
    def destroy_action(self, _event):
        if self._complete:
            self._callback(self._return_text)

    def ok_action(self):
        text = self.text_var.get()
        if text:
            self._return_text = text
            self._complete = True
            self.root.destroy()
        else:
            self.root.bell()

    def cancel_action(self):
        self.root.destroy()

# Alias
msgbox = MsgBox


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
#         print(getFullLibraryPath('mytest', xlFolder))
#         print(getFullLibraryPath(filename1, xlFolder))
#         print(getFullLibraryPath(filename2, xlFolder))
#         print(getFullLibraryPath(filename3, xlFolder))
#         print(getFullLibraryPath(filename4))
#     finally:
#         xl.Quit()
#
#
    mytuple = list(range(1, 6))
    print(mytuple)
    print(list(reversed(mytuple)))
    print(len(mytuple) - 1 - next(i for i,v in enumerate(reversed(mytuple)) if v < 3))
    
    
    
    
