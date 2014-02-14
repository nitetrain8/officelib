"""

Created by: Nathan Starkweather
Created on: 02/14/2014
Created in: PyCharm Community Edition


Proxies for the elements which make up a batch file.

"""
from datetime import datetime

from officelib.pbslib.batchbase import BatchBase, BatchError
# noinspection PyUnresolvedReferences
from officelib.pbslib.batchutil import ExtractDataReport, GroupHeaderData, ParseDateFormat, \
                                        default_strptime_fmt, TimedeltaDays
# noinspection PyUnresolvedReferences
from officelib.olutils import getFullLibraryPath
# noinspection PyUnresolvedReferences
from collections import Counter, OrderedDict

from itertools import islice, takewhile
from weakref import ref as wref


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


class BatchProxyError(BatchError):
    """Exception class for proxy types"""
    pass


class InvalidDateStringError(BatchProxyError):
    """
    @param string: problematic datestring
    """
    def __init__(self, string):
        self.args = "Invalid date string encountered: %s" % string


class ProxyListBase(list, BatchBase):
    """Base for Times/PVs and any
    additional Columns.

    @type _parent: wref
    """

    @property
    def Mode(self):
        return Counter(self).most_common(1)[0][0]  # value of most common thing

    mode = Mode

    @property
    def Parent(self):
        """
        @return: weakrefable parent
        @rtype: BatchBase
        """
        return self._parent()

    @Parent.setter
    def Parent(self, parent):
        """
        @param parent: weakrefable parent
        @type parent: BatchBase
        @return: None
        @rtype: None
        """
        self._parent = wref(parent)

    @property
    def Median(self):
        sorted_self = sorted(self)
        len_self = len(self)

        if len_self % 2:
            return sorted_self[(len_self - 1) / 2]
        else:
            middle = len_self / 2
            return (sorted_self[middle] + sorted_self[middle - 1]) / 2

    @property
    def Mean(self):
        return sum(self) / len(self)

    def AverageStart(self, n=5):
        sum_of = sum(islice(self, 0, n, 1))
        return sum_of / n

    def AverageEnd(self, n=5):
        sum_of = sum(islice(reversed(self), 0, n, 1))
        return sum_of / n

    median = Median
    mean = Mean
    average = mean
    Average = mean


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

        stripped = (v.strip() for v in values)
        has_value = takewhile(bool, stripped)
        super().__init__(float(value) for value in has_value)


class Parameter(BatchBase):
    """Proxy representation of a batch_file
    parameter.

    @type _header: str
    @type _times: Times
    @type _values: Values
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

        if len(self._times) != len(self._values):
            raise BatchProxyError("Malformed batch data.")

    @property
    def Header(self):
        return self._header

    def __getitem__(self, index):
        return self._times[index], self._values[index]

    @property
    def Times(self):
        return self._times

    @property
    def Values(self):
        return self._values

    def __getslice__(self, slice_arg):
        return list(zip(self._times[slice_arg], self._values[slice_arg]))

    def __iter__(self):

        for pair in zip(self._times, self._values):
            yield pair

    def __str__(self):
        return self._header

    def __len__(self):
        return len(self._times)

    def AverageStartValues(self, n=5):
        return self.Values.AverageStart(n)

    def AverageEndValues(self, n=5):
        return self.Values.AverageEnd(n)

    # Aliases

    header = Header
    parameter = header
    type = header
    name = header
    Name = header

    times = Times
    time = times
    Time = times
    dates = times
    Dates = times

    values = Values
    pvs = values
    Pvs = values
    PVs = values
