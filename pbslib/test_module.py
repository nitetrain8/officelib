'''
Created on Jan 13, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather
'''

from cachetypes import DataCache  # @UnresolvedImport
from collections import OrderedDict
import types


class UT_builder(type):
    
    def __new__(cls, name, bases, kwargs):
        
        tests = []
        for k, v in kwargs.items():
            if isinstance(v, types.FunctionType) \
                and v.__name__.startswith("test"):
                    tests.append((k, v))
                    
        kwargs['unit_test_list'] = tests
        
        return super().__new__(cls, name, bases, kwargs)
        
        
        

class CacheTest(metaclass=UT_builder):
    
    
    def test_simple_set_get_item(self, max_cache, keys, values):
        
        cache = DataCache()
        reference = OrderedDict()
        
        cache.setMaxCache(max_cache)
            
        for i, (key, value) in enumerate(zip(keys, values)):
            
            if i > max_cache:
                break
            
            cache[key] = value
            reference[key] = value
            
        for key in keys:
            assert reference[key] == cache[key]
            
    def test_pop_item_cache(self, max_cache, keys, values):
        
        cache = DataCache()
        reference = OrderedDict()
        
        
        
        
        
        
        
        
        
#     try:
#         print("pop empty cache", cache.popitem())
#     except ValueError:
#         pass
#     else:
#         raise AssertionError
#         
#     for i in range(10):
#         cache[test_key % i] = test_value % i
#         
#     print("Iteration Test")
#     for i in cache:
#         print(i)
#     
#     print("Iterating after adding too many links")
#     for i in range(100):
#         cache[test_key % i] = test_value % i
#         
#     for i in cache:
#         print(i)
#         
#     print("Inserting existing item")
#     cache[test_key % 95] = test_value % 95
#     
#     for i in cache:
#         print(i)
    
    test_key = "testkey %d"
    test_value = "testvalue %d"
        

