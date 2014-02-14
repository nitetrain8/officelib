"""
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Base classes/exceptions for batch handler
to avoid cluttering main batchhandler.py file.
"""

BATCH_DEBUG = False


from officelib.pbslib.pbslibexception import PBSLibError
from officelib.pbslib.batchutil import ParseDateFormat
# from officelib.olutils import getFullLibraryPath
from datetime import datetime, timedelta


# Base types and classes


class PBSBatchType(type):
    """Type of PBS batch classes. Currently unneeded."""
    pass  


if BATCH_DEBUG:
    class _PBSBatchDebugType(PBSBatchType):
        """Metaclass for debugging all of PBSlib classes (?)
        
        Pseudo metaclasses take name, bases, kwargs
        and return them after (possibly) modifying them.
        This makes it easy to make specific behaviors
        more modular, and easier to work with. 
        
        Creating property aliases is a bit tricker, since we
        need to make sure that we don't override inherited 
        attributes, 
        '''
        """
        
        from officelib.nsdbg import OverrideWarningMeta, VerboseEmptyMethodMeta, \
                        SlotsNoticeMeta  # , ExplicitVariableDeclarationMeta
        pseudo_meta_list = (
                            OverrideWarningMeta,
                            VerboseEmptyMethodMeta,
                            SlotsNoticeMeta,
#                             ExplicitVariableDeclarationMeta
                            )
        
        def __new__(mcs, name, bases, kwargs):
            
            for pmeta in mcs.pseudo_meta_list:
                name, bases, kwargs = pmeta(name, bases, kwargs)
                
            new_cls = super().__new__(mcs, name, bases, kwargs)
            
            return new_cls
            
    PBSBatchType = _PBSBatchDebugType
    
    
class BatchError(PBSLibError):
    pass
    
    
class InvalidDateStringError(BatchError):
    def __init__(self, string):
        self.args = "Invalid date string encountered: %s" % string


def _timedelta_to_float(td: timedelta):
    """Simple helper function to calculate timedeltas
    @param: timedelta- a datetime module datetime object
    @return: timedelta represented as a float in units of days
    """
    sec_per_day = 86400
    return td.days + td.seconds / sec_per_day


class BatchBase(metaclass=PBSBatchType):
    """
    Base for proxy classes.
    Has some common utility functions.
    """
    
    xl_date_fmts = [
                    "%m/%d/%Y %I:%M:%S %p",
                    '%m/%d/%y %I:%M %p',
                    "%m/%d/%Y %H:%M"
                    ]
    xl_date_preferred_fmt = "%m/%d/%Y %I:%M:%S %p"
    __slots__ = []  # Allow subclasses to support slots

    parse_date_fmt = staticmethod(ParseDateFormat)
    _td2float = staticmethod(_timedelta_to_float)


class DataHandler(BatchBase):
    
    """Abstract implementation of common data handling methods,
    including data extraction, file handling, date parsing, etc.

    Update 1/13/2014: now uses a better algorithm for extracting data
    from CSV. Algorithm was split into multiple functions anyway,
    just in case specific functionality needs to change later.

    Update 1/16/2014: moved the actual building of the sub-dict
    of datadict to a new function, to allow easier access by inheritance
    to building of a single parameter's dict of times/values.

    Also wanted to clarify- this class should hold pretty much
    all implementations of raw data extraction from batch data,
    to make subclassing easier. Just make sure to call the correct
    function from the interface.

    Update 1/29/2014- Really this class is more like a "DataExtracter",
    since its job is to extract the raw CSV. basically. Since moving
    to use of proxy classes to handle batch files, almost all
    of these class has been removed. The remaining methods were static,
    and so removed to batchutil.py.

    """
    xl_date_preferred_fmt = None

        
        
        
        
        
