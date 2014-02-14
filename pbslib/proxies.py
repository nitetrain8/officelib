"""
Created on Jan 16, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Proxies for objects like Times, PVs, Parameters, etc

"""
from datetime import datetime

from officelib.pbslib.batchbase import BatchBase, BatchError
from officelib.pbslib.batchutil import ExtractDataReport, GroupHeaderData, ParseDateFormat, \
                                        default_strptime_fmt, TimedeltaDays
from officelib.olutils import getFullLibraryPath
from collections import Counter, OrderedDict
from itertools import zip_longest, islice, takewhile


class BatchProxyError(BatchError):
    """Exception class for proxy types"""
    pass
    
    
class BatchFileError(BatchProxyError):
    pass


class InvalidDateStringError(BatchProxyError):
    """
    @param string: problematic datestring
    """
    def __init__(self, string):
        self.args = "Invalid date string encountered: %s" % string


class _EmptyColumn():
    """ Simple placeholder class for
    representing empty columns.
    """
    def __iter__(self):
        return iter((None,))
        
        
EmptyColumn = _EmptyColumn()


def _not_empty(other: str) -> int:
    """
    @param other: str to check
    @return: bool
    Simple helper func to pass as predicate
    to itertools.takewhile() when reading
    batch file.
    """
    return other != ''


class ProxyListBase(list, BatchBase):
    """Base for Times/PVs and any
    additional Columns.

    """
    
    @property
    def Mode(self):
        return Counter(self).most_common(1)[0][0]  # value of most common thing
    
    mode = Mode
    
    @property
    def Median(self):
        sorted_self = sorted(self)
        len_self = len(self)
        
        if len_self % 2:
            return sorted_self[(len_self - 1) / 2]
        else:
            middle = len_self / 2
            return (sorted_self[middle] + sorted_self[middle - 1]) / 2
    
    median = Median
    
    @property
    def Mean(self):
        return sum(self) / len(self)
        
    mean = Mean
    average = mean
    Average = mean


def ParseBadDates(raw, fmt=default_strptime_fmt):
    """
    @param raw: list of datestrings to parse
    @type raw: collections.Iterable[str]
    @param fmt: optional first format to try.
    @type fmt: str
    @return: datetime generator. Use generator to avoid building the list twice
                (once here, once when calling super().__init__())
    @rtype: __generator[str]
    """

    strptime = datetime.strptime
    parse = ParseDateFormat

    for date in raw:
        try:
            parsed = strptime(date, fmt)
        except ValueError:
            fmt = parse(date)
            parsed = strptime(date, fmt)
        yield parsed


class Times(ProxyListBase):
    """ Proxy representation of the times of a parameter.
    Initialize from raw data. Allow times class to handle its own 
    processing. 
    
    Note on times-as-floats: all times are relative to 12/31/1899
    This is the reference point for Excel, so it is what we use.

    @type _raw: list[str]
    @type _fmt: str
    @type _floats: list[float]
    @type _elapsed_times: list[float]
    @type self: list[datetime]
    """

    def __init__(self, times):
        """
        @param times: list of raw datestrings
        @type times: collections.Iterable[str]
        @return: None
        @rtype: None
        """
        # Similar to values.__init__, optimize this initialization
        # by chaining generators. Stripped lazily strips the whitespace
        # off of each line.
        # Raw is generated from a takewhile generator. Because
        # the takewhile aborts on predicate failure, and stripped
        # iterated one line at a time by takewhile, finding an
        # empty date string (which indicates end of column)
        # stops the takewhile generator, which stops the stripped
        # generator. Because str.strip() was consuming a significant
        # amount of time according to cProfile, this significantly
        # speeds up the initialization.

        # Predicate passed to takewhile is bool. Because strings,
        # and not numbers, are passed to the bool function,
        # bool('0') is true, and the only false return value
        # is for an actual empty string.

        stripped = (time.strip() for time in times)
        raw = list(takewhile(bool, stripped))

        fmt = ParseDateFormat(raw[0])
        strptime = datetime.strptime

        try:
            super().__init__(strptime(date, fmt) for date in raw)
        except ValueError:
            self.clear()
            fixed = ParseBadDates(raw, fmt)
            super().__init__(fixed)

        self._raw = raw
        self._fmt = fmt
        self._floats = None 
        self._elapsed_times = None

    @property
    def Datestrings(self):
        return self._raw
    
    @property
    def Floats(self):
        """ Return list of values as floats.
        Lazy evaluation.
        """
        if self._floats is None:
            self._floats = self.as_floats()
        return self._floats
                                  
    def as_floats(self, xl_zero=datetime.strptime('12/31/1899', '%m/%d/%Y')):
        """ Calculate datetimes into floats.
        """
        
        return [TimedeltaDays(dt - xl_zero) for dt in self]
        
    @property
    def ElapsedTimes(self):
        
        """ Elapsed time is a common operation for Batch
        Analysis, so build the method right into the times 
        class.
        
        Lazy evaluation.
        @rtype: list of elapsed times in seconds
        """
        
        if self._elapsed_times is None:
            self._elapsed_times = self.as_elapsed_seconds()
        return self._elapsed_times
        
    def as_elapsed_seconds(self):
        """
        @rtype: list of elapsed times in seconds
        """
                
        dt0 = self[0]
        return [(dt - dt0).total_seconds() for dt in self]
        

class Values(ProxyListBase):
    """ Proxy representation of the values of a parameter.
    Initialize from raw data. Allow values class to handle its own 
    processing. 
    """
    
    def __init__(self, values):
        """
        @param values: list of values
        @type values: collections.Iterable[str]
        @return: None
        @rtype: None
        """

        # Similar to times.__init__, optimize this initialization
        # by chaining generators. Stripped lazily strips the whitespace
        # off of each line.
        # Raw is generated from a takewhile generator. Because
        # the takewhile aborts on predicate failure, and stripped
        # iterated one line at a time by takewhile, finding an
        # empty date string (which indicates end of column)
        # stops the takewhile generator, which stops the stripped
        # generator. Because str.strip() was consuming a significant
        # amount of time according to cProfile, this significantly
        # speeds up the initialization.

        # Predicate passed to takewhile is bool. Because strings,
        # and not numbers, are passed to the bool function,
        # bool('0') is true, and the only false return value
        # is for an actual empty string.

        gen = (v.strip() for v in values)
        has_value = takewhile(bool, gen)
        super().__init__(float(value) for value in has_value)


class Parameter(BatchBase):
    """Proxy representation of a batch_file
    parameter.

    @type _header: str
    @type _times: Times
    @type _values: Values
    @type _xlheaders: list[str]
    @type _xldata: list
    """
                     
    def __init__(self, header, times, values):
        """
        @param header: name of parameter eg TempPV(C)
        @type header: str
        @param times: list of raw date strings
        @type times: collections.Iterable[str]
        @param values: list of values
        @type values: collections.Iterable[str]
        @return: None
        @rtype: None
        """
        
        self._header = header
        self._times = Times(times)
        self._values = Values(values)
        self._xlheaders = []
        self._xldata = []
        
        self.AppendColumn(header, self._times.Datestrings)
        self.AppendColumn(header, self._values)
        self.AppendColumn("", EmptyColumn)
        
    @property
    def Header(self):
        return self._header
    
    header = Header
    parameter = header
    type = header
    name = header
    Name = header

    def __getitem__(self, index):
        return self._times[index], self._values[index]

    @property
    def ColumnCount(self):
        return len(self._xldata)
        
    @property
    def RowCount(self):
        return max(len(column) for column in self._xldata)
    
    @property
    def Times(self):
        return self._times
        
    times = Times
    time = times
    Time = times
    dates = times
    Dates = times
    
    @property
    def Values(self):
        return self._values
    
    values = Values
    pvs = values
    Pvs = values
    PVs = values
    
    def xlHeaders(self):
        return self._xlheaders
    
    def xlColumns(self):
        return self._xldata
        
    def AddColumn(self, index, header, column):
        """ Add a column
        @param index: index to insert column
        @param header: str to use as header
        @param column: list to insert as a column
        """
        
        self._xlheaders.insert(index, header)
        self._xldata.insert(index, column)
        
    def AppendColumn(self, header, column):
        """ Append a column to the list of 
        columns.
        @param header: str to use as header
        @param column: column to append
        """
        self._xlheaders.append(header)
        self._xldata.append(column)
    
    def itertimes(self):
        for time in self._times:
            yield time
            
    def itervalues(self):
        for value in self._values:
            yield value

    def __getslice__(self, slice_arg):
        return list(zip(self._times[slice_arg], self._values[slice_arg]))

    def __iter__(self):

        for data_point in zip(self._times, self._values):
            yield data_point
    
    def __reprdata__(self):
        param_repr = ''.join(''.join((
                              '\n',
                              str(t), 
                              '   ', 
                              str(v))) for t, v in zip(self._times, self._values))
        return param_repr
    
    def ShowData(self):
        """
        @return: human readable presentation of data
        """
        return "\n%s: %s%s" % (self.__class__, self._header, self.__reprdata__())
        
    def __str__(self):
        return self._header
        
    def __len__(self):
        return len(self._times)
        
    @property
    def ElapsedTimes(self):
        return self._times.ElapsedTimes
        
    elapsedtimes = ElapsedTimes
    elapsedTimes = elapsedtimes
    elapsed_times = elapsedtimes
    
    def AverageStartValues(self, n=5):
        sum_of = sum(islice(self.Values, 0, n, 1))
        return sum_of / n
        
    def AverageEndValues(self, n=5):
        sum_of = sum(islice(reversed(self.Values), 0, n, 1))
        return sum_of / n
                          
        
class BatchFile(OrderedDict, BatchBase):
    
    """ Proxy representation of data in a batch
    file.

    batch = BatchFile(filename) -> processed batch file

    Access parameters by dict-style lookup: batch[Parameter]
    Due to complexity of names, lookup by most relevant match.
    If multiple matches, return a list of matches and print
    announcement to console.

    Parameters are instances of class Parameter.

    """
    def __init__(self, filename=None):
        super().__init__()
        if filename is not None:
            self.ProcessFile(filename)
        self._filename = filename

    def ProcessFile(self, filename):
        """
        @param filename: filename to process
        @type filename: str
        @return: None
        @rtype: None
        """
        filename = getFullLibraryPath(filename)
        self._filename = filename
        self._create_data()

    def get(self, key, default=None):
        """

        @param key: key to get
        @type key: str
        @param default: value to return if key not found
        @return: value stored for key
        """
        try:
            return self[key]
        except KeyError:
            return default

    def __getitem__(self, key, _dict_getitem_=OrderedDict.__getitem__):
        
        try:
            # Try exact match first
            return _dict_getitem_(self, key)
        except KeyError:
            pass

        return self.__getitems__(key)

    def __getitems__(self, key):
        """
        @param key: key to get items for
        @type key: str
        @return: match or list of matches

        Batch file parameter names can be funky, so provide
        a way for users to almost get the parameter name right,
        but not quite. 
        """

        # No exact match, build a list of all partials
        # We know param are all unique, so we can treat
        # Them as case insensitive.

        matches = []
        lower_key = key.lower()
        for name, parameter in self.items():
            if lower_key in name.lower():
                matches.append(parameter)
                
        # If only one match, let user use directly.
        # Otherwise, inform user and return a list of
        # matches. This is unnecessary for a script,
        # but useful for command-line.
            
        if not matches:
            raise KeyError("No parameter %s found in %s" % (key, self._filename))
                
        elif len(matches) == 1:
            return matches[0]
        else:
            print("\nMultiple matches found:")
            print('\n'.join(match.header for match in matches))
            return matches  
            
    def _create_data(self):

        # Dispatch handling of input and data
        filename = self._filename
        headers, raw_data = ExtractDataReport(filename)
        
        for header, (times, pvs, _empty) in GroupHeaderData(headers, raw_data):

            try:
                param = Parameter(header, times, pvs)
            except Exception as e:
                # Catch errors during creation to reraise with filename of problem. 
                raise BatchFileError("Error occurred trying to make %s in batch file " % header + self.Filename) from e
            
            self[header] = param

    @classmethod
    def fromMapping(cls, mapping: dict):
        new = cls.__new__(cls)
        OrderedDict.__init__(new)
        for key in mapping:
            new[key] = mapping[key]

        return new
         
    @property
    def Filename(self):
        return self._filename

    def xlColumns(self):
        """ Build the list of headers 
        and data together in a format that can easily
        be pasted into excel. 
        
        This mostly exists as a snippet example of how to
        do this correctly, in conjunction with xlData().
        
        Functions are separate because switching between
        row- ordered and column- ordered data sets requires
        a list(zip(*)) on the entire list. Since columns are
        the most common addition but data needs to be list of 
        rows for xl, it is convenient to have these functions 
        be separate. 
        
        This function returns column-ordered data. 
        """
        
        headers = []
        data = []
        for parameter in self.values():
            
            headers.extend(parameter.xlHeaders())
            data.extend(parameter.xlColumns())
        
        return headers, data
        
    def xlData(self):
        
        headers, data = self.xlColumns()
        data = list(zip_longest(*data))
        
        return headers, data
        
    def toWorksheet(self, ws):
        """
        @param ws: Excel Worksheet object from win32com 
        
        Put the data into an excel worksheet. Currently assumes 
        it is the only test in the worksheet. 
        """

        headers, data = self.xlData()

        self.plotdata(ws, headers, data)
        
    def plotdata(self, ws, headers, data):
        
        columns = len(headers)
        rows = len(data)
        
        cells = ws.Cells
        cell_range = cells.Range
        
        header_range = cell_range(cells(1, 1),
                                  cells(1, columns))
        data_range = cell_range(cells(2, 1),
                                cells(rows + 1, columns))
                                  
        header_range.Value = headers
        data_range.Value = data
        
    def py2xlColumn(self, param_name):
        name_index = next(i for i, key in enumerate(self.keys()) if key == param_name)
        return name_index * 3 + 1

    @property
    def Parameters(self):
        return self.values()
    
    @property
    def ColumnCount(self):
        return sum(p.ColumnCount for p in self.Parameters)

    # Support pickling.
    # pickle calls init of dict subclasses, but this causes ours
    # to screw up a bit.

    def __reduce__(self):
        dict_copy = self.__dict__.copy()
        return _BatchFile_pickle_init, tuple(), dict_copy, None, ((k, self[k]) for k in self)


def _BatchFile_pickle_init():
    """
    Because BatchFile inherits dict (via OrderedDict)
    but does funky stuff in its constructor,
    it needs a special pickling protocol, because
    dict subclasses' __init__ methods are called by pickle.

    This function returns an empty batch file that bypasses
    the normal batch file init method.

    @return: empty batch file
    @rtype: BatchFile
    """
    batch = BatchFile.__new__(BatchFile)
    OrderedDict.__init__(batch)
    return batch

if __name__ == '__main__':
    testfile = 'C:/Users/PBS Biotech/Downloads/tpidinsulationp40i3stopat36.9.csv'
    
    b = BatchFile(testfile)
#     for param in b:
#         print(param)
    
    test = b
    #
    # from xllib.xlcom import xlObjs  # @UnresolvedImport
    # xl, wb, ws, cells = xlObjs()
    # b.toWorksheet(ws)
        
    
    

