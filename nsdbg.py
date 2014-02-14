"""
Created on Nov 7, 2013

@author: PBS Biotech


Note 2/13/2014: This module needs refactoring so bad wow.
"""

from types import FunctionType, ModuleType
import os
from time import perf_counter
from collections import OrderedDict
from subprocess import Popen
import tkinter as tk
import tkinter.ttk as ttk


debug_file = "C:/dbg.txt"
debug_file_viewer = "C:/Program Files/Notepad++/notepad++.exe"
debug_flags = "-nosession"
debug_view_cmd = (debug_file_viewer, debug_flags)
debug_active_fh = None


def sys_exit():
    import sys
    sys.exit()


def npp_open(filename=debug_file):

    args = ''.join((' '.join(debug_view_cmd), ' "', filename, '"'))
    try:
        Popen(args)
    except:
        print("failed to open debug file with %s, %s" % debug_view_cmd)


def close_fh():
    global debug_active_fh
    try:
        debug_active_fh.close()
    except:
        pass
import atexit
atexit.register(close_fh)


def view_dbg():
    global debug_active_fh

    try:
        f = debug_active_fh.name
    except AttributeError:  # no active filehandle
        f = debug_file

    editor, flags = debug_view_cmd

    try:
        Popen(' '.join([editor, flags, f]))
    except:
        print("failed to open debug file with %s, %s" % (editor, flags))


def awrite_dbg(data, debug_file=debug_file):
    """data should be a list of primitive types"""

    with open(debug_file, 'a') as f:
        f.write(data)
    view_dbg()


def purge_dbg_file(debug_file=debug_file):
    # noinspection PyUnusedLocal
    with open(debug_file, 'w') as _f:
        pass


def write_dbg(fh=None, data=None, mode='a'):

    global debug_active_fh

    if data is None:
        return

    if fh is None:
        if debug_active_fh is None:
            fh = open_dbg(mode=mode)
        else:
            fh = debug_active_fh
    try:
        fh.write(data + '\n')
    except TypeError:
        fh.write('\n'.join(str(x) for x in data))


def open_dbg(debug_file=debug_file, mode='w'):

    global debug_active_fh

    if debug_active_fh is not None:
        debug_active_fh.close()
    debug_active_fh = open(debug_file, mode)

    return debug_active_fh


def close_dbg_file(f=None):

    global debug_active_fh

    if f is None:
        f = debug_active_fh

    try:
        f.close()
    except:
        pass

    debug_active_fh = None


def dbg_dump(data):
    purge_dbg_file()
    awrite_dbg(data)


# noinspection PyGlobalUndefined
def close_all_xl():

    global client
    import win32com.client as client

    uh_oh_counter = 0
    uh_oh_max = 30
    while True:
        try:
            xl = client.GetActiveObject('Excel.Application')
            xl.DisplayAlerts = False
            xl.Quit()
            xl.DisplayAlerts = True
            xl = None
            del xl
            uh_oh_counter += 1
            print("Excel %d closed" % uh_oh_counter)
        except Exception as e:
            #this is the COM error code corresponding to attempting to
            #close non-existent object, or something.
            if e.__str__()[0] != -2147023170:
                print(e)
                raise e
            break
        #in case something crazy happens and no error
        #is thrown, but the thing gets stuck in an otherwise
        #infinite loop
        if uh_oh_counter > uh_oh_max:
            print("uh_oh!")
            break


def printDir(obj):
    for var in dir(obj):
        print(var, ": ", getattr(obj, var))


def _echo_write_dbg(*msgs):
    data = ' '.join(msgs)
#     print(*(x.strip() for x in msgs))
    print(trace_scope * ' ', data.strip())
#     write_dbg(data=' '.join(msgs), mode='a')

    # inline write_dbg bc of performance criticality
    global debug_active_fh
    fh = debug_active_fh
    if debug_active_fh is None:
            fh = open_dbg(mode='a')

    fh.write(''.join((data, '\n')))

TFC_echo_write_dbg = _echo_write_dbg


def _no_echo_write_dbg(*msgs):
    write_dbg(data=' '.join(msgs), mode='a')


def dbg_echo_off():
    global TFC_echo_write_dbg
    TFC_echo_write_dbg = _no_echo_write_dbg


def dbg_echo_on():
    global TFC_echo_write_dbg
    TFC_echo_write_dbg = _echo_write_dbg


def _echo_only_print(*msgs):
    print(*(x.strip() for x in msgs))


def dbg_echo_only():
    global TFC_echo_write_dbg
    TFC_echo_write_dbg = _echo_only_print

# VC_COORD_WIDGET = 'Label'            
# VC_COORD_WIDGET = 'LabelFrame'
# VC_COORD_WIDGET = 'Button'
# VC_COORD_WIDGET = 'Frame'
VC_COORD_WIDGET = 'Canvas'
VC_COORD_FONT = ("Arial", 7)
VC_COORD_RELIEF = 'groove'

#Note to self-rename this something easier


class TkViewCoordsCallError(Exception):
    """Used internally to distinguish between 
    errors raised due to improper use of TkVC-related
    functions and other types of errors"""
    pass


def TkViewCoords(*args):

    """redirect function for debugging grid-based tkinter
    dialogs. 
    
    If one arg passed, assume function is being used
    as a class decorator, return view_cls_coords function. 
    usage: @TkViewCoords('tk_frame_widget')
           class myclass(...):
           
    name of tk frame widget is a string corresponding to the name of 
    the instance or class attribute which contains an object reference
    to the frame widget to examine. ie cls_instance.__dict__['tk_frame_widget']...
    
    note: user can also pass TkViewCoords(obj.tk_frame_widget) and it [should] work
    
    If two args passed, assume function is being used to examine a single
    instance. usage: obj = TkViewCoords(obj, 'tk_frame_widget')
    
    Throw TkViewCoordsCallError everywhere if function is used incorrectly. 
    
    Todo: badly need of refactoring (esp params)
    """

    # Begin the process of figuring out what the hell the user wants
    if len(args) > 2 or len(args) < 1:
        raise TkViewCoordsCallError("TkViewCoords expects from 1-2 arguments, %d given." % len(args))

    # One arg given. Function is being used as a class decorator, or used incorrectly.
    if len(args) == 1:
        m_name = args[0]

        if isinstance(m_name, str):
            return view_cls_coords(m_name)

            # used incorrectly. passed (obj.master) instead of (obj, obj.master) 
            # or (obj, 'master'), but might still be able to figure out how to 
            # add VC button anyway"""
        elif is_good_tk_master(m_name):
            obj = m_name
            m_name = None
            return view_instance_coords(obj, m_name)

        else:
            raise TkViewCoordsCallError("View coords passed single argument of type \"%s\".\n \
                         Must pass args of types (instance, frame_name) or use as class decorator:\n \
                    TkViewCoords(\"frame_to_examine\")\n \
                         class Dialog():\n \
                         or:\n \
                   TkViewCoords = view_coords(dialog, \"frame_to_examine\")" % repr(m_name))

    elif len(args) == 2:
        obj = args[0]
        m_name = args[1]

        # Allow user to pass the frame by attribute eg obj.name or by name eg 'name'
        if isinstance(m_name, str):
            try:
                master = getattr(obj, m_name)
            except:
                master = None  # fall through to next if, trigger same error
        else:
            master = m_name
            # reverse dict lookup to get instance attr 
            # associated with the Tk master frame
            for k, v in obj.__dict__.items():
                if v is master:
                    m_name = k
                    break
            else:
                m_name = None

        if not is_good_tk_master(master):
            # This is mostly unnecessary, but was a fun exercise in using
            # the traceback and re modules for runtime reflection.

            import traceback
            import re

            #magically get the name and args of the function that called view_dbg. 
            #if this starts showing the wrong function call, change the value of the first line below.
            calling_func_tb_index = -2
            tbfunc = [x.strip() for x in traceback.format_stack()[calling_func_tb_index].split('\n')[:-1]][1]

            #magically get var names sent to func that called view_dbg, yay regex. 
            args = [arg.strip() for arg in re.findall(re.compile(r"(?<=\()(.*)(?=\))"), tbfunc)[0].split(',')]
            for i, arg in enumerate(args):
                # noinspection PyTypeChecker
                if arg.strip('\'') == m_name:
                    # noinspection PyTypeChecker
                    args[i] = arg.strip('\'')
            args.append(__name__)

            #how do you multi-line these damn strings without up the formatting???
            message = '\'{1}\' is not a valid Tk Master\n    Check identity of \'{0}.{1}\' or check function \'{2}.is_good_tk_master()\' to ensure\n\tthe list there is complete.'.format(*args)
            raise TkViewCoordsCallError(message)

        #if everything falls through, do this
        return view_instance_coords(obj, m_name)

    #this should never happen, but if something horrible goes wrong, fall through
    #and raise arcane error. 
    raise TkViewCoordsCallError("Unknown error. Life sucks.")


def view_cls_coords(m_name):
    """Decorator function for use with debugging tkinter dialogs
    ex usage: 
    
    # @TkViewCoords("master")
      class myDialog():
          ....
           
    m_name- the name of the class instance attribute corresponding to
    the master widget within which to display coordinates. The upper-
    rightmost cell will contain a button "Coords" which will toggle 
    the showing of cell coordinates within the master. Yay. 
    
    Works by adding new class methods to the decorated class, and 
    hijacking the __init__ function to automagically set up all 
    the stuff necessary for the view-coords button to show up and work 
    """

    if not isinstance(m_name, str):
        raise TkViewCoordsCallError("TkViewCoords expects argument of type \"str\", not %s" % m_name.__class__)

    def _view_coords(cls):

        if not isinstance(cls, type):
            raise TkViewCoordsCallError("Decorator function expects argument of type \'class\' (type)")
        init = cls.__init__

        def new_init(self, *args, m_name=m_name, **kwargs):
            init(self, *args, **kwargs)

            #in case a dialog is made where 
            #frame is assigned during class creation
            #instead of object initialization

            if m_name.startswith("__"):
                #name was mangled by python
                m_name = "_%s%s" % (cls.__name__, m_name)

            ##is this even necessary?
            try:
                master = getattr(self, m_name)
            except AttributeError:
                master = getattr(cls, m_name)
                setattr(self, m_name, master)

            cols, _rows = master.grid_size()

            self.coordbtn = ttk.Button(master, command=self.toggle_cell_coords, text="Toggle Coords")
            self.coordbtn.grid(column=cols, row=0, sticky=tk.N)

            #call this after making coord button so that coord btn row
            #not included in coords view if user re-calculates coord list
            #after coord button is gridded
            self.coords = self.make_cell_coords(master)

        cls.__init__ = new_init
        cls.make_cell_coords = __make_cell_coords
        cls.toggle_cell_coords = __toggle_cell_coords
        cls.show_cell_coords = __show_cell_coords
        cls.hide_cell_coords = __hide_cell_coords

        #inject shallow copy into nsdbg global namespace
        global VC_class
        VC_class = cls

        return cls

    return _view_coords


def view_instance_coords(old_obj, m_name):
    """similar to the above function, but operates on instances 
    of dialog classes instead.
    
    Create a new class "VC_class" as a wonky snapshot of the passed
    instance. Tricky because tkinter wraps the underlying TCL engine, 
    which doesn't play nicely with standard Python class magic. 
    
    """

    try:
        master = getattr(old_obj, m_name)
    except TypeError:
        master = old_obj

    old_attrs = dict(old_obj.__dict__)

    #get root
    try:
        root_name, root = __find_root(obj=old_obj)
        old_attrs.pop(root_name)

    except TkViewCoordsCallError:
        root_name, root = __find_root(slave=master)

    if root_name is None:
        root_name = "__new_root_"

    VC_init = __build_VC_init_(m_name=m_name, root_name=root_name)
    VC_dict = __build_VC_class_dict(init=VC_init, root=root, master=master, root_name=root_name, m_name=m_name, **old_attrs)

    #might as well save a copy of the newly created class to the module namespace
    #just in case
    global VC_class
    VC_class = type('VC_class', (old_obj.__class__,), VC_dict)

    obj = VC_class(m_name, root_name)

    # This is the old code to simply use methodtype to bind
    # buttons directly onto the same running instance.
    #
    # Uncomment here and comment everything above to get old
    # functionality back

#     from types import MethodType
#     MethodType creates bound method
#     obj=old_obj
#     master=getattr(obj, m_name)
#     obj.make_cell_coords = MethodType(__make_cell_coords, obj)
#     obj.toggle_cell_coords = MethodType(__toggle_cell_coords, obj)
#     obj.show_cell_coords = MethodType(__show_cell_coords, obj)
#     obj.hide_cell_coords = MethodType(__hide_cell_coords, obj)
#     cols, rows = master.grid_size()
#     obj.coordbtn = ttk.Button(master, command=obj.toggle_cell_coords, text="Toggle Coords")
#     obj.coordbtn.grid(column=cols, row=0, sticky=tk.N)
#     obj.coords = obj.make_cell_coords(master)

    return obj


def __build_VC_class_dict(init=None, root=None, master=None, root_name=None, m_name=None, **kwargs):
    """return the dict for a new VC_class object"""

    new_dict = {
            '__init__' : init,
            'make_cell_coords' : __make_cell_coords,
            'toggle_cell_coords' : __toggle_cell_coords,
            'show_cell_coords' : __show_cell_coords,
            'hide_cell_coords' : __hide_cell_coords,
            '__VC_root_ref_' : root,
            '__VC_master_ref_' : master,
            '__VC_root_name_' : root_name,
            '__VC_m_name_' : m_name
            }
    new_dict.update(kwargs)
    return new_dict


def __build_VC_init_(m_name=None, root_name=None):
    """Magical function which builds the init function for 
    VC_class"""

    ####Begin Magic
    def __init__(self, m_name=m_name, root_name=root_name):

        #nested helper function #1
        def copy_widget(widget, new_master, self):
            """Deep copy widget based on extracting non-default
            config options, and all grid_info options. 
            
            Currently doesn't copy button commands."""

            # Reference to self is not properly updated
            # if used as a default parameter, for some reason
            # so need to explicitly pass it.
            # Recalculate class each loop just in case. 

            grid_info = dict(widget.grid_info())
            try:
                del grid_info['in']
            except:
                pass

            try:
                # widget.config() returns weird info, so use .cget(k) to get the relevent info
                config_info = {k : widget.cget(k) for k in widget.config().keys() if widget.cget(k) != ''}
                if is_good_tk_master(widget) and 'text' in widget.config().keys():
                    config_info['text'] += " copy"
            except:
                raise TkViewCoordsCallError("error occurred copying widget %s config_info" % (widget.__repr__()))

            try:
                new = widget.__class__(new_master, **config_info)
                new.grid(**grid_info)
            except:
                raise TkViewCoordsCallError("error occurred copying widget %s" % widget.__repr__(), config_info, grid_info)

            # search all attributes to see if self should contain attribute 
            # reference to widget 
            # if so, insert it into self

            #first check class dict           
            for k, v in self.__class__.__dict__.items():
                if widget is v:
                    setattr(self, k, new)
                    return new

            #now check self dict
            for k, v in self.__dict__.items():
                if widget is v:
                    setattr(self, k, new)
                    return new

            #else pass
            return new

        # Recursively (depth-first) copy ALL 
        # widgets by recreating them through copy_widget.

        # Todo: make everything shallow copies that are gridded inside a new
        # Tk.Tk() frame, so that VC_class window is a realtime
        # reflection of the copied dialog (is this even possible?)

        def copy_slaves(old_master, new_master, self):
            for slave in old_master.grid_slaves():
                new = copy_widget(slave, new_master, self)
                copy_slaves(slave, new, self)

        root = tk.Tk()
        copy_slaves(self.__VC_root_ref_, root, self)

        #attribute cleanup and finalization
        #copies of root and root_name for easy lookup 
        #(not currently used)
        #also copy of master name

        setattr(self, root_name, root)  # compatibility with objects own methods
        self.__new_root_ = root  # stored reference to new root

        master = getattr(self, m_name)

        cols, _rows = master.grid_size()
        #sometimes coords appear in wrong window for some reason
        #if you just pass in root
        #probably due to underlying Tcl architecture using unordered hash map

        self.coordbtn = ttk.Button(master.master, command=self.toggle_cell_coords, text="Toggle Coords")
        self.coordbtn.grid(column=cols, row=0, sticky=tk.N)
        self.coords = self.make_cell_coords(master)

########End new __init__ magic
    return __init__


def is_good_tk_master(obj):

    return isinstance(obj, tk.Tk) or \
           isinstance(obj, tk.Frame) or \
           isinstance(obj, ttk.Frame) or \
           isinstance(obj, ttk.LabelFrame)


# These functions are added to instance or classes passed to 
# TkViewCoords, so they can be used to generate coords


def __make_cell_coords(self, master, widget_type=VC_COORD_WIDGET):
    """Coords is a list of 3-tuples
    Each tuple has the form (row, column, tkWidget)

    widget_cb is the function used to construct widgets
    and set the widget gridops dict to the correct settings
    Use
    """
    font = VC_COORD_FONT
    if widget_type == 'Label':
        self.widget_gridops = {'padx' : (0, 0), 'pady' : (0, 0)}

        def make_coord(row, col, master=master):
            return ttk.Label(master, text="(r%d,c%d)" % (row, col), anchor=tk.CENTER, relief=VC_COORD_RELIEF, font=font)

    elif widget_type == 'LabelFrame':
        self.widget_gridops = {'sticky' : (tk.N, tk.E, tk.W, tk.S)}

        def make_coord(row, col, master=master):
            return ttk.LabelFrame(master, text="(r%d,c%d)" % (row, col))

    elif widget_type == 'Button':
        self.widget_gridops = {'sticky' : (tk.N, tk.E, tk.W, tk.S)}

        def make_coord(row, col, master=master):
            return ttk.Button(master, font=font, text="(r%d,c%d)" % (row, col),
                              width=0, command=lambda: print("(r%d,c%d)" % (row, col)))

    elif widget_type == 'Frame':
        self.widget_gridops = {'sticky' : (tk.N, tk.E, tk.W, tk.S)}

        def make_coord(row, col, master=master):
            widget = ttk.Frame(master, relief=tk.GROOVE)
            label_widget = ttk.Label(widget, font=font,
                                     text="(r%d,c%d)" % (row, col), anchor=tk.CENTER)
            label_widget.grid()
            return widget

    elif widget_type == 'Canvas':
        self.widget_gridops = {'sticky' : (tk.N, tk.E, tk.W, tk.S)}

        def make_coord(row, col, master=master):
            widget = tk.Canvas(master, width=0, height=0,
                               relief=VC_COORD_RELIEF, borderwidth=0)

            widget.create_text((1, 0), anchor=tk.NW, font=font, text="(r%d,c%d)" % (row, col))
            return widget

    else:
        raise TkViewCoordsCallError("error- %s is not a valid widget for coordinate view" % widget_type)

    self.coords_visible = False
    columns, rows = master.grid_size()

    if self.coordbtn.master is master:
        columns -= 1

    coords = [(row, col, make_coord(row, col)) for row in range(rows) for col in range(columns)]

    return coords


def __toggle_cell_coords(self):
    if self.coords_visible:
        self.hide_cell_coords()
    else:
        self.show_cell_coords()
    self.coords_visible = not self.coords_visible


def __show_cell_coords(self):
    ops = self.widget_gridops
    for row, col, widget in self.coords:
        widget.grid(column=col, row=row, **ops)


def __hide_cell_coords(self):
    for _row, _col, button in self.coords:
        button.grid_forget()


def __find_root(slave=None, obj=None):

    """Get root. Try to get it with three strategies:
    try#1: follow master.master chain until master=None
    try#2: scan class dict directly to try to find tk.Tk
    try#3: scan any masters in class dict through master.master chain"""
    if slave is not None:
        #try #1
        root = slave.master
        for _try_count in range(100):
            if isinstance(root, tk.Tk):
                return None, root
            root = root.master

        else:
            raise TkViewCoordsCallError("unable to find root for widget %s" % slave.__repr__())

    elif obj is not None:
        for k, v in obj.__dict__.items():
            if isinstance(v, tk.Tk):
                return k, v
        else:
            raise TkViewCoordsCallError("unable to find root for widget %s" % slave.__repr__())

    return None, None


# The following are debugging functions designed to be used to
# inspect arguments passed to function calls, similar to inspect module
# functionality. These were written before I discovered the inspect module
# but they still work nicely, and provide slightly different functionality.
#
# Each function attaches a printFuncCalls function wrapper to each function
# found according to the rules of each function (ie, module finds all functions
# in itself, as well as all functions of classes and modules in its __dict__, etc).


class CallTraceError(Exception):
    """Throw this to user on bad call."""
    pass


class ClassCallTraceError(CallTraceError):
    """Identify specific errors"""
    pass


class ModuleCallTraceError(CallTraceError):
    """Identify specific errors"""
    pass


class FunctionCallTraceError(CallTraceError):
    """Identify specific errors"""
    pass


# This is the path to the python33/Lib folder
pylibpath = '\\'.join(os.__file__.split("\\")[:-1])

trace_scope = 0
trace_spacer = '    '
# 4 spaces (tab) width

# List of default modules to ignore when recursively printFuncCalls-ing
# modules. Internal use, don't touch at runtime!

__TMC_EXCLUDE_DEFAULTS = [
               'importlib._bootstrap',
               '_frozen_importlib',
               '_weakrefset',
               'codecs',
               'ntpath',
               'STAT',
               'os',
                'win32com',
                'win32com.client',
                're'
               ]

__TFC_EXCLUDE_DEFAULTS = [

                        ]

__TFC_IGNORE_DEFAULTS = [

                        ]
#if OrderedDict is called, then PFC will throw a 
#runtime recursion depth error since it uses
#orderedDict internally for arg parsing. 
__TCC_EXCLUDE_DEFAULTS = [
                          'OrderedDict'
                        ]


# Modify this at runtime
traceModuleCallsExclude = __TMC_EXCLUDE_DEFAULTS[:]
traceFunctionCallsExclude = __TFC_EXCLUDE_DEFAULTS[:]
traceClassCallsExclude = __TCC_EXCLUDE_DEFAULTS[:]
traceFunctionCallsIgnore = __TFC_IGNORE_DEFAULTS[:]


# Reset if it gets too messed up


def resetModuleCallsExclude():
    global traceModuleCallsExclude
    traceModuleCallsExclude = __TMC_EXCLUDE_DEFAULTS[:]


def addModuleCallsExclude(*modules):
    global traceModuleCallsExclude
    traceModuleCallsExclude.extendFiles(modules)


def resetFunctionCallsExclude():
    global traceFunctionCallsExclude
    traceFunctionCallsExclude = __TFC_EXCLUDE_DEFAULTS[:]


def addFunctionCallsExclude(*Functions):
    global traceFunctionCallsExclude
    traceFunctionCallsExclude.extendFiles(Functions)


def resetClassCallsExclude():
    global traceClassCallsExclude
    traceClassCallsExclude = __TCC_EXCLUDE_DEFAULTS[:]


def addClassCallsExclude(*Classes):
    global traceClassCallsExclude
    traceClassCallsExclude.extendFiles(Classes)


def resetFunctionCallsIgnore():
    global traceFunctionCallsIgnore
    traceFunctionCallsIgnore = __TFC_IGNORE_DEFAULTS[:]


def addFunctionCallsIgnore(*Functions):
    global traceFunctionCallsIgnore
    traceFunctionCallsIgnore.extendFiles(Functions)


def traceModuleCalls(module, *, checked=None):
    """Operate on module variables to attach printFuncCalls wrapper
    to everything found in its __dict__. For functions, attach wrapper.
    For classes, attach wrapper to each class function. For modules, 
    recursively call itself, using checked parameter as a dict to avoid 
    infinite recursion. Probably better done by just scanning sys.modules
    or w/e the list of imported modules is, but oh well. 
    
    ToDo: scan for instances of non-builtin classes. 
    Todo: use predicate callback to allow more runtime customization
    of traces.  
    """

    #first pass- create checked dict
    if checked is None:
        checked = {}

    m_id = id(module)

    #already checked module? return
    if checked.get(m_id) is not None:
        return module

    #if not, declare module checked.
    checked[m_id] = True

    #if user passes in non-module, abort
    if not isinstance(module, ModuleType):
        raise CallTraceError("function expects argument of type ModuleType")

    #iterate over dict
    for k, v in module.__dict__.items():

        #wrap functions
        if isinstance(v, FunctionType) and \
            v.__module__ not in traceModuleCallsExclude:
            try:
                module.__dict__[k] = traceFuncCalls(v)
            except FunctionCallTraceError:
                pass

        #wrap classes 
        elif isinstance(v, type) and \
            v.__module__ not in traceModuleCallsExclude:

            try:
                module.__dict__[k] = traceClassCalls(v)
            except ClassCallTraceError:
                pass

        # Left here for use as example code, but depreciated.
        # To trace module calls, scan sys.modules instead.
#         elif isinstance(v, ModuleType) and v.__name__ not in traceModuleCallsExclude:
#
#             #skip if module already checked (redundant with prev code))
#             if checked.get(id(v)) is None:
#                 try:
#                     module.__dict__[k] = traceModuleCalls(v, checked=checked)
#                 except:
#                     #debug
#                     print(k, v)
#                     print(dir(v))
#                     raise

    return module


def traceClassCalls(cls):

    """handle wrapping all functions found in classes"""
    if cls.__name__ in traceClassCallsExclude:
        raise CallTraceError("Attempted to trace functions in excluded class.")

    cls_msg = "class method <{}> of class <%s.%s>" % (cls.__module__, cls.__name__)

    for k, v in cls.__dict__.items():

        #skip builtins
        if (not k.startswith("__")) and (not k.endswith("__")):
            setattr(cls, k, traceFuncCalls(v, annotation=cls_msg.format(k)))

    return cls


def traceFuncCalls(func, annotation=None):

    """Wrap the function by defining a decorator to print
    information about all arguments passed by comparing *args **kwargs
    with the information found in the function's bound code object
    
    Do note that this function is slow as hell due to the amount of extra work
    done on dicts and such.
    
    Update 1/3/2014:
        Function is going to be really ugly due to optimization and legacy code
        left in. Because python is dynamically compiled, inlining function calls
        and even variable assignments will speed up the interpreter. Legacy code
        will show the methodology/provide a name to strange representations of things
        
        Todo: the absolute fastest way to do this will be to create the wrapper entirely
        within the scope of an eval()'d string, to avoid having to unpack and pack 
        arguments, and to absolutely minimize the amount of processing necessary. 
        
        Eventually, the wrapper needs to actually become a class with configurable options
        to control display over single args and such, which calls the function directly
        through __call__.
    """

    #if function was already wrapped, abort
    #also, avoid recursion if function is called on itself
    #because toys are too fun to not play with, and it might 
    #be useful to track calls to and from debug library. 

    if func.__name__ in traceFunctionCallsExclude:
        raise FunctionCallTraceError("Attemped to wrap excluded function")

    try:
        if func.__module__ == 'nsdbg':
            if 'printFuncCalls' in func.__name__ or 'dbg' in func.__name__ or 'trace' in func.__name__:
                raise FunctionCallTraceError("Attempted to wrap internal use debug function")
    except AttributeError:
        pass

    #get data
    v = func
    _c = v.__code__
    arg_only_count = _c.co_argcount
    kw_only = _c.co_kwonlyargcount
    arg_names = _c.co_varnames
    defaults = func.__kwdefaults__ or {}
    arg_only = arg_names[:arg_only_count]

    if annotation is None:
        annotation = "function: <%s.%s>" % (func.__module__, func.__name__)

    tc = perf_counter

    if func.__name__ in traceFunctionCallsIgnore:
        # Wrap function, but don't trace args or return value
        def printFuncCalls(*args, **kwargs):
            # noinspection PyGlobalUndefined
            global trace_scope
            # noinspection PyTypeChecker
            _tss = '\n' + trace_spacer * trace_scope

            TFC_echo_write_dbg(''.join((
                                        _tss,
                                        "Call to ",
                                        annotation
                                        ))
                           )

            trace_scope += 1

            # outer try/finally necessary to ensure scope trace is updated
            # properly when errors are caught in caller's try/except during func call
            t_call = tc()
            try:
                _return = func(*args, **kwargs)
            finally:  # print regardless of raised exception
                t_return = tc()
                trace_scope -= 1

                TFC_echo_write_dbg(''.join((
                                              _tss,
                                              "Return from ",
                                              annotation,
                                              _tss,
                                              '   return time: %f seconds' % (t_return - t_call)
                                          ))
                                   )

            return _return

    else:
        # noinspection PyGlobalUndefined
        def printFuncCalls(*args, **kwargs):

            #list of positional only
            passed = OrderedDict(zip(arg_only, args))

            #the rest of the list of args must have been passed as *args
            passed["args"] = args[arg_only_count:]

            kwcopy = dict(kwargs)
            #add default kwargs (faster with list comprehension using if/else?)
            keys = kwcopy.keys()
            passed.update({k : kwcopy.pop(k) if k in keys else defaults[k] for k in arg_names[arg_only_count:arg_only_count + kw_only]})

            #only list unnamed, passed kwargs in kwargs 

            passed["kwargs"] = kwcopy

            global trace_scope
            global trace_spacer
            _tss = trace_scope * trace_spacer

            if 'self' in passed.keys():
                passed['self'] = 'self'

            TFC_echo_write_dbg(''.join((
                                     '\n',
                                     _tss,
                                    "Call to ",
                                    annotation,
                                    '\n',
                                    _tss,
                                    '   parameters: (',  # inline fcall starting here
                                    ', '.join("%s=%s" % (param, val) for param, val in passed.items()),
                                    ')'
                                    ))
                           )

            trace_scope += 1
            # outer try/finally necessary to ensure scope trace is updated
            # properly when errors are caught in caller's try/except during func call
            t_call = tc()
            try:
                _return = func(*args, **kwargs)
            finally:
                t_return = tc()
                trace_scope -= 1
                TFC_echo_write_dbg(''.join((
                                         "\n",
                                         _tss,
                                         "Return from ",
                                         annotation
                                         ))
                               )

                # this try block is necessary to prevent accidental errors from
                # being raised when working with library modules that use internal function
                # calls on objects that may not print info properly
                #
                # _return is placed on an empty line to force CPython to attempt
                # to look up the value, and throw exception quickly if _return
                # was not defined.
                #

                try:
                    # noinspection PyStatementEffect,PyUnboundLocalVariable
                    _return
                    TFC_echo_write_dbg(_tss + '   return value: {}'.format(_return))
                except:
                    TFC_echo_write_dbg(_tss + '   return value: <Exception>')

                TFC_echo_write_dbg(_tss + '   return time: %f seconds' % (t_return - t_call))

            return _return

    #store reference to function because why not
    printFuncCalls._func = func

    return printFuncCalls

traceFunctionCalls = traceFuncCalls  # alias

# Debugging stuff


def VerboseEmptyMethod(func):

    def EmptyMethodWrapper(self, *args, **kwargs):
        print("Non-implemented method <%s.%s> accessed" % (self.__class__.__name__, func.__name__))
        return func(self, *args, **kwargs)

    return EmptyMethodWrapper


def __empty():
    pass

# noinspection PyUnresolvedReferences
__empty_code = __empty.__code__.co_code


def VerboseEmptyMethodMeta(name, bases, kwargs, _empty=__empty_code):
    for k, v in kwargs.items():
        isinstance(v, FunctionType)
        if isinstance(v, FunctionType) and v.__code__.co_code == _empty:
            kwargs[k] = VerboseEmptyMethod(k)
    return name, bases, kwargs

OverrideIgnoreList = [
                    '__module__',
                    '__qualname__',
                    '__doc__',
                    '__slots__',
                    '__init__',
                    '__new__',
                    'EmptyMethodWrapper'
                    ]


def OverrideWarningMeta(name, bases, kwargs):

    """Primarily a debugging metaclass for checking
    for Overrideed properties and functions"""

    ignore = OverrideIgnoreList

    # Override Warnings
    for base in bases:
        basekeys = base.__dict__.keys()
        for k in kwargs:
            if k not in ignore and k in basekeys:
                print("Override Warning!")
                print("<%s.%s> overrides <%s.%s>" % (
                    name, k,
                    base.__name__, k
                ))
    return name, bases, kwargs


slots_notice = """Slots override notice-
<%s.%s> __dict__ overriding <%s.%s> __slots__
"""


def SlotsNoticeMeta(name, bases, kwargs, notice=slots_notice):

    bases_with_slots = []
    bases_with_dicts = []
    for base in bases:
        if hasattr(base, '__slots__'):
            bases_with_slots.append(base)
        elif hasattr(base, '__dict__'):
            bases_with_dicts.append(base)

    if bases_with_slots:

        if '__slots__' not in kwargs.keys():
            base = bases_with_slots[0]
            print(notice % (
                            kwargs['__module__'],
                            name,
                            base.__module__,
                            base.__name__
                            )
                  )

        if bases_with_dicts:
            dict_notice = "Note: bases also without slots:"
            dicts_list = '\n'.join("<%s.%s>" % (base.__module__, base.__name__)
                                                for base in bases_with_dicts)
            print(dict_notice, dicts_list)

    return name, bases, kwargs


def ExplicitVariableDeclarationMeta(name, bases, kwargs, force=True):

    if '__slots__' not in kwargs.keys():
        print("Explicit variable warning- <%s.%s>__slots__ not found!" % (name, kwargs['__module__']))
        if force:
            print("Forcing __slots__.")
            kwargs['__slots__'] = []

    return name, bases, kwargs


def dbgPrintEmptyClsDecorator(cls, *, __empty=__empty_code):
    for attr in dir(cls):
        val = getattr(cls, attr)
        try:
            if val.__code__.co_code == __empty:
                setattr(cls, attr, VerboseEmptyMethod(val))
        except AttributeError:
            pass
