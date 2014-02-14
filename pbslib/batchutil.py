"""
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Misc utility functions as necessary

"""

from itertools import zip_longest, islice, chain
from _strptime import _regex_cache, _TimeRE_cache
from officelib.pbslib.batchbase import PBSLibError


class DateStringError(PBSLibError, ValueError):
    def __init__(self, string):
        self.args = "Invalid date string encountered: %s" % string


def grouper(iterable, n):
    """Itertools group recipe"""
    args = [iter(iterable)] * n
    return zip_longest(*args, fillvalue=None)
    
    
def FilterIndexRange(data, cb=bool, islice=islice):
    """ Return first occurrence range in data where cb(x) is true.
    Index of first time true, until first time not true.
    Ignore subsequent dips into and out of range.

    Stupid function do not use.
    """
    try:
        start = next(i for i, v in enumerate(data) if cb(v))
    except StopIteration:
        start = None
        try:
            end = next(i for i, v in enumerate(islice(data, 0, None), 0) if not cb(v))
        except StopIteration:
            end = None
    else:
        end = next(i for i, v in enumerate(islice(data, start, None), start) if not cb(v))
    
    return start, end    


def ExtractDataReport(filename: str) -> tuple:
    """
    @param filename: name of data report to open
    @type filename: str
    @return: headers and data from data report
    @rtype: (list[str], list[str])
    """

    with open(filename, 'r') as f:
        headers = f.readline().split(',')
        data = f.readlines()

    data = zip(*[x.split(',') for x in data])
    
    return headers, data
    
    
def GroupHeaderData(headers, data):
    return zip(headers[::3], grouper(data, 3))    

# Ordered (somewhat) in order of likeliness of occurring.
known_strptime_fmts = [
                "%m/%d/%Y %I:%M:%S %p",
                "%m/%d/%Y",
                "%m/%d/%y %I:%M %p",
                "%m/%d/%Y %H:%M",
                ]

default_strptime_fmt = known_strptime_fmts[0]  # This is the most common batch date format


def __fast_parse_date(date_string: str, fmt: str, regex_cache=_regex_cache, TimeRE_cache=_TimeRE_cache) -> int:
    """
    @param date_string: date string to scan
    @type date_string: str
    @param fmt: strptime format use for scan
    @type fmt: str
    @return: True/False
    @rtype: bool

    Delve into the innards of _strptime.py to find the logic
    for determining whether a datestring is valid or not.
    """

    format_regex = regex_cache.get(fmt)
    if not format_regex:
        try:
            format_regex = TimeRE_cache.compile(fmt)
        except:
            return False
        regex_cache[fmt] = format_regex

    return bool(format_regex.match(date_string))


def ParseDateFormat(date, guess=None, known=known_strptime_fmts, parse=__fast_parse_date):
    """ Parse date format, return the format the string is in.
    Test from a known list of date formats.

    @param date: date string to scan
    @type date: str
    @param guess: optional date string to try first
    @type guess: str
    @param known: known list of possible formats to try
    @type known: collections.Iterable[str]
    @param parse: inline reference to the fast parse function
    """
    fmt = guess
    if fmt:
        if parse(date, fmt):
            return fmt

    for fmt in known:
        if parse(date, fmt):
            return fmt

    raise DateStringError("""Invalid date string format:
    '%s' does not match format '%s'.""" % (date, fmt))


def flatten(iterables, chain_from_iterable=chain.from_iterable):
    """
    @param iterables: iterable of iterables
    @type iterables: collections.Iterable[collections.Iterable]
    @return: flattened iterable
    @rtype: itertools.chain
    """
    return chain_from_iterable(iterables)


def TimedeltaDays(td):
    """Simple helper function to calculate timedeltas
    @param td: timedelta to convert
    @type td: datetime.timedelta
    @return: timedelta represented as a float in units of days
    @rtype: float
    """
    return td.days + td.seconds / 86400



