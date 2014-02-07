'''
Created on Jan 10, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather
'''


from os.path import exists
from officelib.xllib import xlDateFormatError, \
                            cellRangeStr, \
                            xl_date_to_float as _xllib_date_to_float, \
                            xlByRows, \
                            xlDown
                            

try:
    from cachetypes import LimitedDataCache
    from pbsbase import BatchHandlerError, BatchBase

except ImportError:
    from .cachetypes import LimitedDataCache 
    from .pbsbase import BatchHandlerError


class BatchTime(BatchBase):
    
    __slots__ = []
    
    '''Represent the time column of batch parameter'''
    
    def __init__(self, times, parameter=None):
        self._times = times
        self._parameter = parameter
    
    
    
'''Proxy/virtual representation container classes'''
class Parameter(BatchBase):
    '''Anything needed here for special parameters?
    '''
    def __init__(self, name, xvalues, yvalues):
        self._name = name
        self._xvalues = xvalues
        self._yvalues = yvalues
        
    @property
    def XValues(self):
        return self._xvalues
        
    '''property aliases'''
    Time = XValues
    xvalues = XValues
    time = xvalues
    times = XValues
    Times = xvalues
    
    
    @property
    def YValues(self):
        return self._yvalues
        
    '''property aliases'''
    Values = YValues
    values = Values
    yvalues = values
    pv = values
    pvs = values
    PVs = values
    PVS = values
    PresentValues = values
    presentvalues = values

    
    

        

class BatchFile(AbstractDataHandler):
    
    '''Interface to accessing data of a batch file
    
    Since this interface is its own window into the batch file,
    no need to set up any special caching scheme, other than
    a reference to the single dict. 
    
    Todo: incorporate into BatchHandler class as default container?
            then use class as cache object instead of dict.
    '''
    
    def __init__(self, filename, Handler=None):
        
        self._cache = None
        self._filename = filename
        self._handler = Handler
        
    def getDataDict(self):
        
        if self._cache is not None:
            return self._cache
        else:
            data = self._make_data_dict(self.filename)
            self._push_cache(data)
            
    def _push_cache(self, data):
        
        self._cache = data
        try:
            self.Handler._push_cache(self.Filename, data)
        except AttributeError:
            pass
            
    def _syncHandler(self):
        try:
            self.Handler._push_cache(self.Filename, self._cache)
        except AttributeError:
            raise BatchHandlerError("Error when attempting to sync handler cache")
            
    @property
    def Handler(self):
        return self._handler
        
    @Handler.setter
    def Handler(self, Handler):
        self._handler = Handler
        self._syncHandler()
    
    handler = Handler
            
    @property
    def Filename(self):
        return self._filename

    @Filename.setter
    def Filename(self, filename):
        self._filename = filename
        self._cache = None
    
    filename = Filename
    
    def __str__(self):
        return self._filename
        
    
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
        
        '''
        
    #max number of dicts full of data to keep around.
    max_cache = 10
    
    def __init__(self, files):
        
        if isinstance(files, str) and exists(files):
            '''Single file passed as string'''
            self._files = [files]
        else:  
            self._files = files
            
        self._cache = LimitedDataCache()
        
        self.setMaxCache = self._cache.setMaxCache
        
        self.setMaxCache(self.max_cache)
        self.xl_date_preferred_fmt = self.xl_date_fmts[0]
                             
    def appendFile(self, File):
        self._files.append(File)
        
    def extendFiles(self, filelist):
        self._files.extend(filelist)
                         
    def __iter__(self):
        return iter(self._files)   
        
    def getDataDict(self, batch_file):
    
        '''Main public interface for accessing data from Handler's 
        batch files. '''
        try:
            return self._cache[batch_file]
        except KeyError:
            pass
        
        data = self._make_data_dict(batch_file)
        self._cache[batch_file] = data
        
        return data
        
    def rebuildDataDict(self, batch_file):
        '''Implementation of building data dict thingy'''
        
        data = self._make_data_dict(batch_file)
        self._cache[batch_file] = data

        return data
         
    def handleFiles(self):
        for f in self._files:
            self.buildDataDict(f)
        
    def iterDataDicts(self):
        for f in self._files:
            yield self.getDataDict(f)     
        
    def _push_cache(self, filename, data):
        self._dataDictCache[filename] = data
             
    def reset(self):
        self._files = []
        self._cache = LimitedDataCache()    
        
    def clearCache(self):
        self._cache = LimitedDataCache()
       
        
def verbose(func):
    def VerboseWrapper(*args, **kwargs):
        print("%s called!" % func.__name__)
        return func(*args, **kwargs)
    return VerboseWrapper

FUNC_TEST_FOLDER = "C:\\Users\\Public\Documents\\PBSSS\\Functional Testing"


#this function is poorly named
def get_logger_data(cells, parameter, *, return_structure=list):
    '''given cells object (search space),
    search for string parameter, return date 
    and value column data
    '''
    date_cell = cells.Find(What=parameter, After=cells(1,1), SearchOrder=xlByRows)
    pv_cell = cells.Find(What=parameter, After=cells(1,date_cell.Column), SearchOrder=xlByRows)
       
    date_col = date_cell.Column 
    row_start = date_cell.Row + 1
    row_end = date_cell.End(xlDown).Row
    pv_col = pv_cell.Column
    
    if pv_col - date_col == 1:
        
        '''Almost always'''    
        target = cellRangeStr(
                             (row_start, date_col),
                             (row_end, pv_col)
                             )
    
    #     data = tuple(zip(zip(*cells.Range(target).Value2)))
        data = dict(zip(['Time', 'PV'], return_structure(zip(*cells.Range(target).Value2))))
    
    else:
        x_target = cellRangeStr(
                                (row_start, date_col),
                                (row_end, date_col)
                                )
                                
        y_target = cellRangeStr(
                                (row_start, pv_col),
                                (row_end, pv_col)
                                )
        
        x_data = return_structure(zip(*cells.Range(x_target).Value2))
        y_data = return_structure(zip(*cells.Range(y_target).Value2))
                 
        data = {'Time' : x_data, 'PV' : y_data}            

    return data
    


if __name__ == '__main__':
    '''Tests and such'''
    print(type(BatchBase))
    print(type(BatchFile))
    print(type(BatchHandler))
    print(type(dict))
