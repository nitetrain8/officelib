'''
Created on Jan 10, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather
'''

from officelib import OfficeLibError


class PBSLibError(OfficeLibError):
    '''pbslib exception class'''
    pass
    
    
class InternalException(PBSLibError):
    '''Tried to access non-public class or constructor'''
    pass

