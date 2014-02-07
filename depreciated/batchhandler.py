'''
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather
'''

from os.path import exists as path_exists
# from officelib.xllib import cellRangeStr, \
#                             xlByRows, \
#                             xlDown
                            
from datetime import datetime
                            
from cachetypes import DataCache  # @UnresolvedImport
from pbslibexception import BatchHandlerError  # @UnresolvedImport
from _batchhandlerbase import BatchBase  # @UnresolvedImport
from batchutil import grouper  # @UnresolvedImport
from officelib.olutils import getFullLibraryPath


class AbstractDataHandler(BatchBase):
    
    '''Abstract implementation of common data handling methods,
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
    
    '''
    
    xl_date_fmts = [
                "%m/%d/%Y %I:%M:%S %p",
                '%m/%d/%y %I:%M %p',
                "%m/%d/%Y %H:%M",
                ]
    xl_date_preferred_fmt = None
    #preserve slots just in case
    __slots__ = []
    
    # Maybe add constructor guard here?
                
    def _extract_from_csv(self, filename):
        
        '''Type checking of filename'''
        try:
            with open(filename, 'r') as f:
                headers = f.readline().split(',')
                data = f.readlines()
        except TypeError:
            raise BatchHandlerError("Invalid filename or filetype")
            
        data = list(zip(*[x.split(',') for x in data]))
        
        return headers, data
        
    def _build_dict_headers(self, headers, data):
        ''' Do not access this directly!! This function
        relies on the data pulled straight from _extract_from_csv
        is well-formed, and that headers are in order and correspond
        to the unlabeled columns from data
        
        @param: headers- list of headers pulled from the top
                        of the batch file. Note that this function
                        assumes header size has already been chopped
                        down to one header per 3 data rows 
                        aka one header per time, pv, spacer column
        @param: data- the time, pv, spacer data pulled from csv
        @return: return- dict of data parsed into date strings (time column)
                        and float values (pv column)
        
        '''
        datadict = {}
        _build_param_dict = self._build_param_dict
        for header, (times, pvs, _) in zip(headers, grouper(data, 3)):
            datadict[header] = _build_param_dict(times, pvs, True)
            
        return datadict
        
    def _build_param_dict(self, times, pvs, as_dict=True):
        '''The actual transformation of columns of data from
        raw strings to a 'Time', 'PV' dict of pre-processed values
        for an individual header/batch file parameter. 
        
        @param: times- column of times data
        @param: pvs- column of pv data
        @param: as_dict- return data as dict. If false, return as 2-tuple of lists.
        
        @return: individual dict entry
        
        be sure times correspond to pvs!
        
        ''' 
        if as_dict:
            sub_data = {
                        'Time' : [time.strip() for time in times if time.strip()],
                        'PV' : [float(pv.strip()) for pv in pvs if pv.strip() != '']  # make sure not to exclude 0's
                        }
        else:
            sub_data = (
                        ([time.strip() for time in times if time.strip()],
                        [float(pv.strip()) for pv in pvs if pv.strip() != ''])
                        )
            
        return sub_data
            
    def _extract_full_datadict(self, filename):
        ''' These functions are closely related. 
            They could be inlined together, but separated
            out just in case they need to be accessed independently
            or separately later.
            
        @param: filename- filename. passed to getFullLibraryPath to
                        get the full path.
        @param: return- the fully built datadict.
        '''        
        
        filename = getFullLibraryPath(filename)
        
        headers, data = self._extract_from_csv(filename)
        datadict = self._build_dict_headers(headers, data)
        _parse_dates = self._parse_dates

        for parameter in datadict.values():
            parameter['Time'] = _parse_dates(parameter['Time'])
            
        return datadict
 
    def _parse_date_fmt(self, date):
        
        '''Parse based on assumption that all batch file dates
        use the same format. May revisit if I find this is a poor
        assumption.'''
        strptime = datetime.strptime
        
        try:
            strptime(date, self.xl_date_preferred_fmt)  # check stored fmt first
            return self.xl_date_preferred_fmt
        except:
            pass
        
        for fmt in self.xl_date_fmts:
            try:
                strptime(date, fmt)  # throw error if wrong format
                return fmt  # if we found fmt, return
            except:
                pass
                
        # if we didn't find fmt, raise error
        raise BatchHandlerError("Couldn't parse dates- improperly formatted data.")
            
    def _parse_dates(self, dates):
        
        ''' Date parsing implementation.
            Parse dates by determining the format of the first
            date in the column, then applying it to the rest of 
            the column. On error, apply date-format conversion on a
            by-case basis. 
            
            First loop- catch error thrown by strptime, move to slow loop.
            
            Second loop- let error raise, which indicates date could
                         not be parsed. 
        '''
            
        _fmt = self._parse_date_fmt(dates[0])
            
        try:
            dates = self._xl_dates_to_float(dates, _fmt)
            self.xl_date_preferred_fmt = _fmt
            return dates
        except ValueError:
            # Uh oh, handle?
            raise
            pass
            
        # parse line by line. this is probably really slow. 
        # I didn't bother to test it, you really shouldn't be here. 
        _parse_date_fmt = self._parse_date_fmt
        _xl_single_date_to_float = self._xl_single_date_to_float
        for i, date in enumerate(dates):
            _fmt = _parse_date_fmt(date)
            dates[i] = _xl_single_date_to_float(date, _fmt)
            
        return dates
            
    def _xl_single_date_to_float(self, date_string, date_fmt="%m/%d/%Y %I:%M:%S %p"): 
        '''Parse dates to floats individually (slow) '''
        
        xlStartDateTime = datetime.strptime('12/31/1899', '%m/%d/%Y')
        return self._timedelta_to_float(
                                datetime.strptime(
                                                date_string, 
                                                date_fmt
                                                ) - xlStartDateTime)
    
    def _xl_dates_to_float(self, date_strings, date_fmt="%m/%d/%Y %I:%M:%S %p"):

        ''' Inline the xllib function here just in case it needs to change
        due to local needs.
            
        xllib.xl_date_to_float
    
        Give list of dates and times (dates w/o time are assumed at midnight 
        in any date_fmt, with corresponding (and correct) date date_fmt string, get 
        a list back that gives the dates in units of days since Dec 31, 1899. 
        This is how xl stores dates as floats.
         
        @param: date_strings- list of date strings eg Aug 25, 1945 to parse
        @param: date_fmt- the date format to feed to strptime to parse strings
        
        @return- list of parsed strings converted to floats.   
            
        '''
    
        # See python docs on datetime module for interpretation of 
        # date_fmt options. TL;DR: default date_fmt is month/day/year hour minute 
        # second AM/PM'''
    
        strptime = datetime.strptime
        td_to_float = self._timedelta_to_float
        
        # datetime object set to an Excel floating point date time value of '0'
        xlStartDateTime = strptime('12/31/1899', '%m/%d/%Y')
        
        return [td_to_float(strptime(date_string, date_fmt) - xlStartDateTime) \
                            for date_string in date_strings if date_string]

    def _timedelta_to_float(self, timedelta):  
        '''Simple helper function to calculate timedeltas
        
        @param: timedelta- a datetime module datetime object
        @return: timedelta represented as a float in units of days
        
        '''  
        sec_per_day = 86400
        return timedelta.days + timedelta.seconds / sec_per_day
    
    
class BatchHandler(AbstractDataHandler):
    
    ''' Class to encapsulate manipulation of batch data outside of 
        Excel.
        
        Feed in a list of files (complete with pathname!), call functions to
        analyze data.
        
        Primary interface should be through dictionaries which return 
        dicts of Parameter: [Time][PV] based on filename.
         
        Nested dictionaries, yay. 
        
        Since columns are always in order of Time, PV, it may be easier to
        just return them as lists or tuples instead of bothering to make time
        and pv a nested dict. 
        
        Data is cached and purged on a last-called-first-purged basis,
        to prevent having to repeatedly process lots of batch files. However,
        data can add up, so gotta beware.  
        
        Since usually we're only interested in contents of a few sets of 
        data, need to add an interface for accessing only those contents. 
        
        Update 1/9/2014:
        Most implementation details moved to superclass AbstractDataHandler.
        
        Update 1/13/2014:
        Superclass, class, and cache class re-written. New module as well. 
        
        This is the public interface of batch handling.
        '''
        
    # max number of dicts full of data to keep around.
    _max_cache = 10
        
    def __init__(self, files):
        super().__init__()
        
        if isinstance(files, str) and path_exists(files):
            # Single file passed as string
            self._files = [files]
            
        elif isinstance(files, list):
            self._files = files
        
        else:
            raise TypeError("Files must be filenames or a list of files.")
            
        self._cache = DataCache()
        
        self._cache.setMaxCache(self._max_cache)
        self.xl_date_preferred_fmt = self.xl_date_fmts[0]
                             
    def appendFile(self, File):
        self._files.append(File)
        
    def extendFiles(self, filelist):
        self._files.extend(filelist)
                         
    def __iter__(self):
        return iter(self._files)   
        
    def getDataDict(self, batch_file):
    
        '''Main public interface for accessing data from Handler's 
        batch files. 
        @param: batch_file- filename to get batch file from
        @return: dict full of data 
        '''
        
        try:
            return self._cache[batch_file]
        except KeyError:
            pass
            
        return self.buildDataDict(batch_file)

    def buildDataDict(self, batch_file):
        '''Implementation of building data dict thingy'''
        
        data = self._extract_full_datadict(batch_file)
        self._push_cache(batch_file, data)
        
        return data
         
    def handleFiles(self):
        for f in self._files:
            self.buildDataDict(f)
            
    handle = handleFiles
        
    def iterDataDicts(self):
        for f in self._files:
            yield self.getDataDict(f)     
        
    def _push_cache(self, filename, data):
        self._dataDictCache[filename] = data
             
    def reset(self):
        self._files = []
        self._cache = DataCache()    
        
    def clearCache(self):
        self._cache = DataCache()

    def setMaxCache(self, value):
        self._cache.setMaxCache(value)
        self._max_cache = value

if __name__ == '__main__':
    b = BatchHandler("C:\dbg.txt")













 
