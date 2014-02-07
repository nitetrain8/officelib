'''
Created on Jan 21, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Some helper classes/proxies for 
making excel easier. Probably won't
be used much.
'''

class DataSeries():
    ''' Simple proxy to make it easier to
    set a chart's data series.
    
    '''
    
    def __init__(self, XValues=None, Values=None, Name=None):
        
        self._XValues = XValues
        self._Values = Values
        self._Name = Name
        
        raise NotImplemented
        
    @property 
    def XValues(self):
        return self._XValues
        
    @property
    def Values(self):
        return self._YValues
        
    @property
    def Name(self):
        return self._Name