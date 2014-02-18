"""
Created on Oct 7, 2013

@author: PBS Biotech

Contains functions/etc for working with Excel nicely
bundled into one module. Functions are essentially wrappers
around win32com and win32com.client dispatches, returning
win32com dispatch objects.


Update 1/16/2014

Lots of constant update over the months!

I felt this was important to point out:

    if verbose:
        v_print = print
    else:
        v_print = __v_print_none  # override print if not verbose

This snippet appears a lot in this library.

This provides an easy runtime way of determining whether to echo
progress to console.

Originally I just used:

    if not verbose:
        print = __v_print_none

but the IDE complains about overriding the built-in method.

Update 1/29/2014:
    Moved to new xllib module, file renamed xlcom.
    Reduce size of file, begin separating out code that is not relevant
    to the python <-> COM server communication process.
"""

from win32com.client import DispatchEx
from win32com.client.gencache import EnsureModule, GetModuleForCLSID, EnsureDispatch
from win32com.client.CLSIDToClass import GetClass
# noinspection PyUnresolvedReferences
from pythoncom import com_error as py_com_error
from tkinter.filedialog import askopenfilenames
from datetime import datetime
from os.path import split as path_split
from officelib.olutils import getFullLibraryPath
from officelib.const import xlLinear, xlByRows, xlDiagonalUp, xlContinuous, \
                                        xlDiagonalDown, xlNone, xlEdgeTop, \
                                        xlEdgeBottom, xlEdgeRight, xlEdgeLeft, xlInsideHorizontal, \
                                        xlInsideVertical, xlXYScatter, xlPrimary, xlSecondary, xlCategory, xlValue
from officelib import OfficeLibError

from xllib._wincom_type_hint import update_typehints


class xllibDefaultArg():
    """Create a new class so that 'None' can be
    distinguished from "Did not pass an arg"

    This should only be needed where we want wrap a series
    of wincom function calls through a single wrapper, but want
    to avoid a function call, as opposed to sending "None" or a default.

    eg, don't change chart Title, rather than resetting it to "None".
    """
    pass


class xlLibError(OfficeLibError):
    """Base Exception for xllib errors"""
    pass


class xlDateFormatError(xlLibError):
    """Pass"""
    pass


# misc constants
PYLIST_TO_XL_ROW = 2
XLTIME_TO_SEC = 86400
XL_POINT_TO_PIXEL = 24 / 18
XL_PIXEL_TO_POINT = 18 / 24
XL_POINT_TO_INCH = 0.25 / 18
XL_INCH_TO_POINT = 18 / 0.25
XL_ROW_HEIGHT = 15  # points?
XL_COL_WIDTH = 8.34  # units?


# noinspection PyUnusedLocal
def __v_print_none(*args, **kwargs):
    pass


def AddTrendlines(xlchart, linetype=xlLinear):

    sc = xlchart.SeriesCollection()
    for i in range(1, sc.Count + 1):
        series = sc(i)
        trendline = series.Trendlines().Add()
        trendline.Type = linetype
        trendline.DisplayEquation = True
        trendline.DisplayRSquared = True


# Exists just to use as a reference
def find_cell_by_text(cells, text, SearchOrder=xlByRows, startRow=1, startCol=1):
    return cells.Find(What=text, After=cells(startRow, startCol), SearchOrder=SearchOrder)
    

# series of internal helper functions to determine
# emptiness of worksheets, workbooks, and excel applications
    
def __ws_is_empty(ws):
    """ Return True if ws appears
    to be empty.
    @param ws: win32com.gen_py Worksheet
    """
    used_range = ws.UsedRange
    count = used_range.Count
    if count > 1:
        return False
    else:  # worksheet shows count of 1 for both empty and only single cell
        return bool(used_range.Columns) and bool(used_range.Rows)


def __wb_is_empty(wb):
    """ Return True if workbook
    appears to be empty.
    Return False otherwise.
    @param wb: win32com.gen_py workbook.
    """
    
    for ws in wb.Worksheets:
        if not __ws_is_empty(ws):
            return False
    return True
    
    
def __find_empty_wb(xl):
    """ Scan xl to see if
    there are any empty workbooks.

    @param xl: win32com gen_py Excel.Application
    """

    for wb in xl.Workbooks:
        if __wb_is_empty(wb):
            return wb
    return None
            

def __ensure_wb(xl):
    """ Internal use.
    @param xl: excel application from win32com.

    Make sure a new workbook is returned for
    dispatching functions.
    """
    
    wb = __find_empty_wb(xl)
    if wb is not None:
        return wb
        
    return xl.Workbooks.Add()
    
    
def __ensure_ws(wb):
    
    if not wb.Worksheets.Count:
        return wb.Worksheets.Add()
    else:
        return wb.Worksheets(1)
    
    
def Excel(new=False, visible=True, verbose=True, v_print=__v_print_none):
    """Get running Excel instance if possible, else
    return new instance.
    """
    
    if verbose:
        v_print = print

    if new:
        xl = EnsureNewDispatch("Excel.Application")
        v_print("New Excel instance created, returning object.")
    else:
        xl = EnsureDispatch("Excel.Application")
    
    xl.Visible = visible

    return xl
    
    
def xlBook(filepath=None, new_xl=False, visible=True, verbose=True):
    """Get win32com workbook object from filepath.
    If workbook is open, get active object.
    If workbook is not open, create a new instance of
    xl and open the workbook in that instance.
    If filename is not found, see if user specified
    default filename error behavior as returning new
    workbook. If not, raise error. If so, return new workbook

    Warning: returns error in some circumstances if dialogs or
    certain areas like formula bar in the desired Excel instance
    have focus.

    @param filepath: valid filepath
    @param visible: xl instance visible to user?
                    turn off to do heavy processing before showing
    @param new_xl: open in a new window
    @param verbose- echo progress to console

    @return: = the newly opened xl workbook instance

    Update 1/15/2014- Lots of refactoring to make it really clean and such.
    Or so I tried.

    Update 1/29/2014- rewote and moved most logic to new function
    xlBook2. This function now supplies an identical interface to old xlBook, for backward
    compatibility with existing code.
    """

    _xl, wb = xlBook2(filepath, new_xl, visible, verbose)
    return wb
    

def xlBook2(filepath=None, new_xl=False, visible=True, verbose=True):
    """Get win32com workbook object from filepath.
    If workbook is open, get active object.
    If workbook is not open, create a new instance of
    xl and open the workbook in that instance.
    If filename is not found, see if user specified
    default filename error behavior as returning new
    workbook. If not, raise error. If so, return new workbook

    Warning: returns error in some circumstances if dialogs or
    certain areas like formula bar in the desired Excel instance
    have focus.

    @param filepath: valid filepath
    @type filepath: str
    @param visible: xl instance visible to user?
                    turn off to do heavy processing before showing
    @type visible: bool
    @param new_xl: open in a new window
    @type new_xl: bool
    @param verbose: echo progress to console
    @type verbose: bool

    @return: the newly opened xl workbook instance

    Update 1/15/2014- Lots of refactoring to make it really clean and such.
    Or so I tried.

    Update 1/29/2014- this function is now converted to abstract internal function.
    Interfaced moved to new function with same name.

    This function still contains logic.

    Update 1/31/2014- renamed function xlBook2, now public.
    """
    
    if verbose:
        v_print = print
    else:
        v_print = __v_print_none
        
    xl = Excel(new=new_xl, visible=visible, verbose=verbose)
        
    if not filepath:
        wb = __ensure_wb(xl)
        return xl, wb 
    
    # First try to see if passed name of open workbook
    try:
        _base, name = path_split(filepath)
        wb = xl.Workbooks(name)
        wb.Activate()
        v_print("\'%s\' found, returning existing workbook." % filepath)
        return xl, wb
    except:
        pass
    
    # Workbook wasn't open, get filepath and open it.
    try:
        filepath = getFullLibraryPath(filepath, hint=xl.DefaultFilePath, verbose=verbose)
    except:
        if new_xl:
            xl.Quit()
        else:
            xl.Visible = True
        raise xlLibError("Couldn't find path specified, check that it is correct.")
        
    v_print("Attempting to create new workbook for \"%s\"." % filepath)
    try:
        wb = xl.Workbooks.Open(filepath, Notify=False)
    except py_com_error as e:
        raise xlLibError("Unknown error occurred.") from e

    v_print("Filename \'%s\' found.\nReturning newly opened workbook." % filepath)
    wb.Activate()
    return xl, wb
        
    # This is unreachable, but will catch anything falling through
    # if the above block is refactored. 

    # noinspection PyUnreachableCode
    raise xlLibError("Unknown error occurred. \nCheck filename, if the target file is open, ensure\nno dialogs are open.")


def xlObjs(filename=None, new=False, visible=True, verbose=True):
    """easy return of excel app object,
    workbook object, worksheet object , cells
    object in one func

    Update 1/15/2014-
    After excessive refactoring of the "get excel stuff" family of functions,
    this should be the main programming interface for opening instances of excel.

    Ask for a filename (or none), get all the objects. Yay.

    @param filename: the filename to open. New excel if None.
    @param new: open a new excel application window. sometimes doesn't work.
    @param visible: make the excel application visible before returning.
                    set to false to do heavy computation before showing.
    @param verbose: echo actions to console
    @return 4-tuple: of (xlApplication, xlWorkbook, xlWorksheet, worksheet cells)
    """
    
    if verbose:
        v_print = print
    else:
        v_print = __v_print_none
        
    # get the workbook by calling the xlBook function
    # get other objects directly and return them as a tuple
        
    if filename is not None:
        xl, wb = xlBook2(filename,
                            new_xl=new,
                            visible=visible,  
                            verbose=verbose)

        ws = __ensure_ws(wb)
        cells = ws.Cells
        v_print("Returning Excel instance objects.")
        
    # Same as above, but get a fresh workbook

    else:
        
        xl = Excel(new=new, visible=visible, verbose=verbose)
        wb = __ensure_wb(xl)
        ws = __ensure_ws(wb)
        cells = ws.Cells
        v_print("Returning new Excel instance objects.")

    return xl, wb, ws, cells


# noinspection PyProtectedMember,PyUnusedLocal
def EnsureNewDispatch(prog_id, bForDemand=1):  # New fn, so we default the new demand feature to on!

    # This whole function stolen from win32com.client.gencache
    # only modification from EnsureDispatch is the use of DispatchEx
    # instead of Dispatch, plus adjusting pathnames as necessary
    # allows creation of new instance of com object
    # while still generating makepy python class

    """Given a COM prog_id, return an object that is using makepy support, building if necessary"""
    disp = DispatchEx(prog_id)
    if not disp.__dict__.get("CLSID"):  # Eeek - no makepy support - try and build it.
        try:
            ti = disp._oleobj_.GetTypeInfo()
            disp_clsid = ti.GetTypeAttr()[0]
            tlb, _index = ti.GetContainingTypeLib()
            tla = tlb.GetLibAttr()
            _mod = EnsureModule(tla[0], tla[1], tla[3], tla[4], bForDemand=bForDemand)
            GetModuleForCLSID(disp_clsid)
            # Get the class from the module.
            disp_class = GetClass(str(disp_clsid))
            disp = disp_class(disp._oleobj_)
        except py_com_error:
            raise TypeError("This COM object can not automate the makepy process - please run makepy manually for this object")
    return disp


def ChangeBorders(RemoveRange=None, AddRange=None, BorderType=xlContinuous):
    """Expanding borders in excel is REALLY ugly.
    @param: RemoveRange
        cell range to RemoveRange borders from
        enter a cell range object
    @param: AddRange
        same thing, but where to AddRange cells.
    @param: BorderType
        xl enum value corresponding to the BorderType for "AddRange"
    """

    # mostly copy/paste from excel code. yay.
    if RemoveRange is not None:
        RemoveRange.Select()
        xl = RemoveRange.Application
        Selection = xl.Selection
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        Selection.Borders(xlEdgeLeft).LineStyle = xlNone
        Selection.Borders(xlEdgeTop).LineStyle = xlNone
        Selection.Borders(xlEdgeBottom).LineStyle = xlNone
        Selection.Borders(xlEdgeRight).LineStyle = xlNone
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    if AddRange is not None:
        AddRange.Select()
        xl = AddRange.Application
        Selection = xl.Selection
        Selection.Borders(xlEdgeLeft).LineStyle = BorderType
        Selection.Borders(xlEdgeTop).LineStyle = BorderType
        Selection.Borders(xlEdgeBottom).LineStyle = BorderType
        Selection.Borders(xlEdgeRight).LineStyle = BorderType
        Selection.Borders(xlInsideVertical).LineStyle = BorderType
        Selection.Borders(xlInsideHorizontal).LineStyle = BorderType


def prompt_files(multiple=True):
    allowedfiletypes = ["{Text, csv} {.txt .csv}", "{Excel} {.xlsx .xls}", "{All} {.*}"]
#     initialdir = "C:/Users/Public/Documents/PBSSS"
    initialdir = "C:/Users/PBS Biotech/Downloads"
    files = askopenfilenames(filetypes=allowedfiletypes, multiple=multiple, initialdir=initialdir).strip("{}").split("} {")
    if isinstance(files, str):
        files = [files]
    return files


def xl_date_to_float(date_strings, date_fmt="%m/%d/%Y %I:%M:%S %p"):

    """Give list of dates and times (dates w/o time are assumed at midnight
    in any date_fmt, with corresponding (and correct) date date_fmt string, get
    a list back that gives the dates in units of days since Dec 31, 1899.
    This is how xl stores dates as floats."""

    # See python docs on datetime module for interpretation of
    # date_fmt options. TL;DR: default date_fmt is month/day/year hour minute
    # second AM/PM

    strptime = datetime.strptime

    # For clarity
    def timedelta_to_float(timedelta):
        sec_per_day = 86400
        return timedelta.days + timedelta.seconds / sec_per_day

    # datetime object set to an Excel floating point date time value of '0'
    xlStartDateTime = strptime('12/31/1899', '%m/%d/%Y')
    try:
        return [timedelta_to_float(strptime(date_string, date_fmt) - xlStartDateTime)
                                for date_string in date_strings if date_string != '']
    except ValueError:
        raise xlDateFormatError("Invalid date date_fmt")
        

def col_to_csv(*lists):
    """Turn data from excel Range.Values into exportable format for csv
    Basically, invert rows/columns.

    Speculation based on name (docstring written months later)."""
    return '\n'.join([str(x).strip("()") for x in list(zip(*lists))])


CHART_WIDTH_DEFAULT = 300
CHART_HEIGHT_DEFAULT = 180


def CreateChart(worksheet,
                ChartType=xlXYScatter,
                Left=None,
                Top=None,
                Width=CHART_WIDTH_DEFAULT,
                Height=CHART_HEIGHT_DEFAULT):
    """Nothing special here. Adding an excel chart involves some required parameters,
    so might as well create a python function that takes care of setting defaults if
    not otherwise specified.

    @param worksheet: worksheet instance
    @type worksheet: win32com.gen_py.typehint0x1x6._Worksheet._Worksheet
    @param ChartType: type of chart as xl enum
    @type ChartType: int
    @param Left: offset from left edge of sheet in points
    @type Left: int | float
    @param Top: offset from top edge of sheet in points
    @type Top: int | float
    @param Width: width of chart in points
    @type Width: int | float
    @param Height: height of chart in points
    @type Height: int | float
    """

    chart_count = worksheet.ChartObjects().Count

    # These two snippets, when used to create a bunch of charts,
    # will create each new chart in two columns going down at
    # the beginning of the worksheet.
    if Left is None:
        Left = 20 + chart_count * (20 + Width) * (chart_count % 2)

    if Top is None:
        Top = 50 + chart_count // 2 * (50 + Height)

    chartobj = worksheet.ChartObjects().Add(Left=Left, Top=Top, Width=Width, Height=Height)
    chart = chartobj.Chart
    chart.ChartType = ChartType

    PurgeSeriesCollection(chart)
    
    return chart


def FormatChart(chart,
                SourceData=None,
                ChartTitle=None,
                xAxisTitle=None,
                yAxisTitle=None,
                Trendline=None,
                Legend=None):
    """ Similar to create chart function, to take care of all the
     annoying formatting I'd have to type out otherwise.

     None = Do nothing. Otherwise, use bool or str.

    @param chart: chart object
    @type chart: win32com.gen_py.typehint0x1x6._Chart._Chart
    @param SourceData: source data as an address string with both X and Y values
    @type SourceData: str | None
    @param ChartTitle: chart title
    @type ChartTitle: str | None
    @param xAxisTitle: x axis title
    @type xAxisTitle: str | None
    @param yAxisTitle: yaxis title
    @type yAxisTitle: str | None
    @param Trendline: type of trendline (xl enum)
    @type Trendline: int | bool | None
    @param Legend: show legend on chart
    @type Legend: bool | None
     """

    if SourceData is not None:
        chart.SetSourceData(SourceData)

    if Legend is not None:
        chart.HasLegend = Legend

    if ChartTitle is not None:
        chart.HasTitle = True
        if not isinstance(ChartTitle, str):
            ChartTitle = str(ChartTitle)  # allow objects to be used to identify charts
        chart.ChartTitle.Text = ChartTitle

    axes = chart.Axes(AxisGroup=xlPrimary)
    xAxis = axes(1)
    yAxis = axes(2)

    if xAxisTitle is not None:
        xAxis.HasTitle = True
        if not isinstance(xAxisTitle, str):
            xAxisTitle = str(xAxisTitle)  # allow objects to be used to identify x axis
        xAxis.AxisTitle.Text = xAxisTitle

    if yAxisTitle is not None:
        yAxis.HasTitle = True
        if not isinstance(yAxisTitle, str):
            yAxisTitle = str(yAxisTitle)  # allow objects to be used to identify y axis
        yAxis.AxisTitle.Text = yAxisTitle

    if Trendline is not None:
        if isinstance(Trendline, bool):
            Trendline = xlLinear
        AddTrendlines(chart, Trendline)


def FormatAxesScale(chart, XAxisMin=None, XAxisMax=None, Y1AxisMin=None,
                            Y1AxisMax=None, Y2AxisMin=None, Y2AxisMax=None):
    """
    @param chart: chart
    @type chart: win32com.gen_py.typehint0x1x6._Chart._Chart
    @param XAxisMin: minimum x axis
    @type XAxisMin: float | int | None
    @param XAxisMax: maximum y axis
    @type XAxisMax: float | int | None
    @param Y1AxisMin: minimum y1 (primary) axis
    @type Y1AxisMin: float | int | None
    @param Y1AxisMax: maximum y1 (primary) axis
    @type Y1AxisMax: float | int | None
    @param Y2AxisMin: minimum y2 (secondary) axis
    @type Y2AxisMin: float | int | None
    @param Y2AxisMax: maximum y2 (secondary) axis
    @type Y2AxisMax: float | int | None
    @return: None
    @rtype: None

    Excel records accessing chart axes as
    chart(Type, AxisGroup). Axis group defaults to
    primary, included here for clarity.

    Only access axes if parameter is passed: otherwise,
    exception may be thrown if accessing non-existent axis.
    """

    if XAxisMin or XAxisMax:
        xAxis = chart.Axes(xlCategory, xlPrimary)
        if XAxisMin:
            xAxis.MinimumScale = XAxisMin
        if XAxisMax:
            xAxis.MaximumScale = XAxisMax

    if Y1AxisMin or Y1AxisMax:
        yAxis1 = chart.Axes(xlValue, xlPrimary)
        if Y1AxisMin:
            yAxis1.MinimumScale = Y1AxisMin
        if Y1AxisMax:
            yAxis1.MaximumScale = Y1AxisMax

    if Y2AxisMin or Y2AxisMax:
        yAxis2 = chart.Axes(xlValue, xlSecondary)
        if Y2AxisMin:
            yAxis2.MinimumScale = Y2AxisMin
        if Y2AxisMax:
            yAxis2.MaximumScale = Y2AxisMax


def CreateDataSeries(chart,
                     XValues,
                     YValues,
                     Name=None,
                     SeriesLabels=None,
                     CategoryLabels=None):
    """
    @param chart: chart to create data series for
    @type chart: win32com.gen_py.typehint0x1x6._Chart._Chart
    @param XValues: Address string in format "=SheetName!XStart:XEnd"
    @type XValues: str
    @param YValues: Address string in format "=SheetName!YStart:YEnd"
    @type YValues: str
    @param Name: Str or Address string in format "=SheetName!Cell"
    @type Name: str
    @param SeriesLabels: NotImplemented
    @type SeriesLabels: NotImplemented
    @param CategoryLabels: NotImplemented
    @type CategoryLabels: NotImplemented
    @return: New Data Series
    @rtype: win32com.gen_py.typehint0x1x6.Series.Series
    """

    if SeriesLabels is not None or CategoryLabels is not None:  # todo
        raise NotImplemented

    SeriesCollection = chart.SeriesCollection()

    Series = SeriesCollection.NewSeries()

    Series.XValues = XValues
    Series.Values = YValues

    if Name and type(Name) is str:
        Series.Name = Name
    else:
        Series.Name = str(Name)

    return Series


def PurgeSeriesCollection(chart):
    for series in chart.SeriesCollection():
        series.Delete()


class HiddenXl():
    """ Excel works much faster when the application is hidden,
    because it doesn't have to draw updates to the screen.
    This simple context manager hides the excel window
    upon entering, and automatically shows it again upon
    exiting, regardless of errors thrown during context.
    """
    def __init__(self, xl):
        self.xl = xl

    def __enter__(self):
        self.xl.Visible = False

    # noinspection PyUnusedLocal
    def __exit__(self, *_args):
        self.xl.Visible = True
        return False


if __name__ == '__main__':
    
    # insert unit tests here (?)
    # xl = EnsureDispatch("Excel.Application")
    # xl.Visible = True
    # wb = xl.Workbooks(1)
    # chart = wb.Charts("Off to Auto")
    #
    # FormatAxesScale(chart, *[1 for i in range(6)])
    update_typehints()



    # ws = wb.Worksheets(1)
    # r = ws.UsedRange
    # print(r)
    # for ws in wb.Worksheets:
    #     print(ws)

