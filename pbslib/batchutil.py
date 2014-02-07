'''
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Misc utility functions as necessary

'''

from itertools import zip_longest, islice
from officelib.pbslib.batchbase import BatchError  # @UnresolvedImport


def grouper(iterable, n):
    '''Itertools group recipe. Swallow pride.'''
    args = [iter(iterable)] * n
    return zip_longest(*args, fillvalue=None)
    
    
def FilterIndexRange(data, cb=bool, islice=islice):
    ''' Return first occurrence range in data where cb(x) is true.
    Index of first time true, until first time not true. 
    Ignore subsequent dips into and out of range. 
    ''' 
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
    
    try:
        with open(filename, 'r') as f:
            headers = f.readline().split(',')
            data = f.readlines()
    except TypeError:
        raise BatchError("Invalid filename or filetype")
        
    data = zip(*[x.split(',') for x in data])
    
    return headers, data
    
    
def groupHeaderData(headers, data):
    return zip(headers[::3], grouper(data, 3))    






    
    
    
    