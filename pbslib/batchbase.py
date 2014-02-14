"""
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Base classes/exceptions for batch handler
to avoid cluttering main batchhandler.py file.
"""
from officelib import OfficeLibError


BATCH_DEBUG = False


# Base types and classes

class PBSLibError(OfficeLibError):
    """pbslib exception class"""
    pass


class BatchError(PBSLibError):
    pass


class PBSBatchType(type):
    """Type of PBS batch classes. Currently unneeded."""
    pass  


if BATCH_DEBUG:
    from officelib.pbslib.test.debug_type import make_debug_type
    PBSBatchType = make_debug_type(PBSBatchType)


class BatchBase(metaclass=PBSBatchType):
    """
    Base for proxy classes.
    Has some common utility functions.

    Update 2/14/2014: ALl of everything
    moved to places where they make more sense.
    """

    __slots__ = ()  # Allow subclasses to support slots

        
        
        
        
        
