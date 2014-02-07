'''
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather


classes in proxies were beginning to try too hard to
represent themselves as both batch file information
and excel spreadsheet information, while pretending
that excel didn't exist yet. 

This module will now inherit from proxies objects to 
be used to implement spreadsheet-specific behavior. 

In particular, adding proxy references from children to 
parents will be very useful.
'''

from officelib.pbslib.batchbase import DataHandler  # @UnresolvedImport
from officelib.pbslib.proxies import BatchFile, Parameter, Times, Values, BatchFileError, EmptyColumn  # @UnresolvedImport
from itertools import zip_longest
from weakref import ref as wref


class xlBatchError(BatchFileError):
    pass


class xlTimes(Times):
    ''' Proxy representation of the times of a parameter.
    Initialize from raw data. Allow times class to handle its own 
    processing. 
    
    Note on times-as-floats: all times are relative to 12/31/1899
    This is the reference point for Excel, so it is what we use. 
    '''
    
    def __init__(self, parent, times):
        super().__init__(times)
        self._parent = wref(parent)
        
    @property
    def Parent(self):
        return self._parent()
                

class xlValues(Values):
    ''' Proxy representation of the values of a parameter.
    Initialize from raw data. Allow values class to handle its own 
    processing. 
    '''
    
    def __init__(self, parent, values):
        super().__init__(values)
        self._parent = wref(parent)
        
    @property
    def Parent(self):
        return self._parent()


class xlParameter(Parameter):
    '''Proxy representation of a batch_file
    parameter.
    
    '''
                     
    def __init__(self, parent, header, times, values):
        
        self._header = header
        self._times = xlTimes(self, times)
        self._values = xlValues(self, values)
        self._xlheaders = []
        self._xldata = []
        self._topleft = None
        
        self.AppendColumn(header, self._times.Datestrings)
        self.AppendColumn(header, self._values)
        self.AppendColumn("", EmptyColumn)
        
        self._parent = wref(parent)
        
    @property
    def Parent(self):
        return self._parent()

    @property
    def ColumnCount(self):
        return len(self._xldata)
        
    @property
    def RowCount(self):
        return max(len(column) for column in self._xldata)

    def xlHeaders(self):
        return self._xlheaders
    
    def xlColumns(self):
        return self._xldata
        
    def AddColumn(self, index, header, column):
        ''' Add a column.
        @param index: index to insert column
        @param header: str to use as header
        @param column: list to insert as a column
        '''
        
        self._xlheaders.insert(index, header)
        self._xldata.insert(index, column)
        
    def AppendColumn(self, header, column):
        ''' Append a column to the list of 
        columns.
        @param header: str to use as header
        @param column: column to append
        '''
        self._xlheaders.append(header)
        self._xldata.append(column)

    __len__ = RowCount

        
class xlBatchFile(BatchFile):
    
    ''' Proxy representation of data in a spreadsheet.
    
    batch = xlBatchFile(filename) -> processed spreadsheet
    that can communicate with excel.
    
    Access parameters by dict-style lookup: batch[Parameter]
    Due to complexity of names, lookup by most relevant match.
    If multiple matches, return a list of matches and print
    announcement to console. 
    
    Parameters are instances of class Parameter.
    
    '''
    def __init__(self, ws, filename=None):
        
        super().__init__(filename)
        self._ws = ws
        
    def _create_data(self):

        # Dispatch handling of input and data
        filename = self._filename
        headers, raw_data = DataHandler.ExtractCSV(filename)
        
        for i, (header, (times, pvs, _empty)) \
                in enumerate(DataHandler.groupHeaderData(headers, raw_data), start=1):
            
            param = xlParameter(self, header, times, pvs)
            self[header] = param
        
    def xlColumns(self):
        ''' Build the list of headers 
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
        '''
        
        headers = []
        data = []
        for parameter in self.values():
            
            headers.extend(parameter.xlHeaders())
            data.extend(parameter.xlColumns())  # iterable None for zip_longest
            
        return headers, data
        
    def xlData(self):
        
        headers, data = self.xlColumns()
        data = list(zip_longest(*data))
        
        return headers, data
        
    def toWorksheet(self, ws):
        '''
        @param ws: Excel Worksheet object from win32com 
        
        Put the data into an excel worksheet. Currently assumes 
        it is the only test in the worksheet. 
        '''

        headers, data = self.xlData()

        self.plotdata(ws, headers, data)
        
    def plotdata(self, ws, headers, data):
        
        columns = len(headers)
        rows = len(data)
        
        cells = ws.Cells
        cell_range = cells.Range
        
        header_range = cell_range(cells(1,1),
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








 
