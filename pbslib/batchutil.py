"""
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Misc utility functions as necessary

"""

from itertools import zip_longest, islice
from datetime import datetime

def grouper(iterable, n):
    """Itertools group recipe. Swallow pride."""
    args = [iter(iterable)] * n
    return zip_longest(*args, fillvalue=None)
    
    
def FilterIndexRange(data, cb=bool, islice=islice):
    """ Return first occurrence range in data where cb(x) is true.
    Index of first time true, until first time not true.
    Ignore subsequent dips into and out of range.
    """
    try:
        start = next(i for i,v in enumerate(data) if cb(v))
    except StopIteration:
        start = None
        try:
            end = next(i for i,v in enumerate(islice(data, 0, None), 0) if not cb(v))
        except StopIteration:
            end = None
    else:
        end = next(i for i,v in enumerate(islice(data, start, None), start) if not cb(v))
    
    return start, end    


def ExtractCSV(filename):
    

    with open(filename, 'r') as f:
        headers = f.readline().split(',')
        data = f.readlines()

    data = zip(*[x.split(',') for x in data])
    
    return headers, data
    
    
def groupHeaderData(headers, data):
    return zip(headers[::3], grouper(data, 3))    


known_date_fmts = [
                "%m/%d/%Y %I:%M:%S %p",
                '%m/%d/%y %I:%M %p',
                "%m/%d/%Y %H:%M"
                ]


from _strptime import _regex_cache, _TimeRE_cache


def __fast_parse_date(date_string: str, fmt: str, regex_cache=_regex_cache, TimeRE_cache=_TimeRE_cache) -> int:
    """
    @param date_string: date string to scan
    @param fmt: format use for scan
    @return: True/False

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


def ParseDateFormat(date, guess: str=None, known: list=known_date_fmts, parse=__fast_parse_date) -> str:

    """Parse date format, return the format the string is in.
    Test from a known list of date formats.
    """

    if guess:
        if parse(date, guess):
            return guess

    for fmt in known:
        if parse(date, fmt):
            return fmt

    raise ValueError("Invalid date string format : '%s'" % date)



    
    
    
