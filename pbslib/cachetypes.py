"""Here for completion. A standard doubly linked list
        appends and pops by index instead of keys.''''''
Created on Jan 10, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather


Module to hold special container types used
for PBS lib caching schemes.

These abstract the implementation details of
how the cache decides to push/flush/pop/purge/clear etc
away from the classes themselves.

"""
from sys import getsizeof as sizeof


class _Link():
    __slots__ = [
                 '_next',
                 '_prev',
                 '_key',
                 '_value'
                ]
                
                
class LinkedListBase(_Link):
    __slots__ = [
                 '_num_links'
                ]

    def __init__(self):
        
        self._num_links = 0
        
        self._next = self
        self._prev = self
        self._key = self
        self._value = self
        
    def _insert_after(self, link, prev_link):
        
        """Insert the given link object after the other
        given link"""
        
        next_link = prev_link._next
        
        prev_link._next = link
        link._prev = prev_link
        
        link._next = next_link
        next_link._prev = link
        
        self._num_links += 1
    
    def _remove_link(self, link):
        
        """ Decrement link counter, shuffle references.
            remove link's references"""
        
        prev_link = link._prev
        next_link = link._next
        
        link._next = None
        link._prev = None

        prev_link._next = next_link
        next_link._prev = prev_link
        
        self._num_links -= 1
        
    def _move_link_to_front(self, link):
        """remove link and move after self"""
        self._remove_link(link)
        self._insert_after(link, self)
        
    def _move_link_to_end(self, link):
        """remove link and move before self"""
        self._remove_link(link)
        self._insert_after(link, self._prev)
    
    def __len__(self):
        return self._num_links
        
    def clear(self):
        
        current = self._next
        self._next = None
        
        while current is not self:
            
            next_link = current._next
            next_link._prev = None

            current._next = None
            current._prev = None
            current._key = None
            current._value = None
            
            del current
            current = next_link
            
        self._next = self
        self._prev = self
        
    def __sizeof__(self):
        n = len(self)  # number of links
        size = sizeof(_Link()) * n + sizeof(self.__slots__)
        return size
        

class DoublyLinkedList(LinkedListBase):
    
    """Since I made a doubly linked list, might as well move the
    implementation out of the Data Cache, so I can finish it properly later"""
    
    def __init__(self, iterable=None):
        
        super().__init__()
        
        if iterable:
            for _i, value in enumerate(iterable):
                self.pushfirst(value)
                
    def __iter__(self):
        current = self._next
        while current is not self:
            yield current._value
            current = current._next
        
    def popitem(self, index=None):
        """Pop index, or return popfirst"""
        if not self._num_links:
            raise ValueError("Can't pop empty list")
        
        if index is None:
            return self.popfirst()
        
        link = self._get_link_by_index(index)
        self._remove_link(link)
        return link._value
            
    def popfirst(self):
        """Pop the first item in the list"""
        if not self._num_links:
            raise ValueError("Can't pop empty list")
            
        link = self._next
        self._remove_link(link)

        return link._value
        
    def poplast(self):
        """Pop the last item in the list"""
        if not self._num_links:
            raise ValueError("Can't pop empty list")
            
        link = self._prev
        self._remove_link(link)
        return link._value
    
    def __delitem__(self, index):
        link = self._get_link_by_index(index)
        self._remove_link(link)
    
    def __setitem__(self, index, item):
        
        index = self._check_index(index)
        
        link = self._get_link_by_index(index)
        link._value = item
    
    def __getitem__(self, index):
        
        self._check_index(index)
            
        link = self._get_link_by_index(index)
        return link._value
    
    def _check_index(self, index):
        if index < 0:
            index = self._num_links - 1 - index
            
        if index < 0 or index > self._num_links - 1:
            raise IndexError("Index out of range")
            
        return index
    
    def _get_link_by_index(self, index):
        """ Get a link based on the index.
            Start scan at head of the list. """
        current = self._next
        for _i in range(index):
            current = current._next
        return current
        
    def _remove_link_by_index(self, index):
        """Remove and return a link by index"""
        link = self._get_link_by_index(index)
        self._remove_link(link)
        return link

    def __contains__(self, value):
        """Iterate through self and return true if keys is found"""
        current = self._next
        while current is not self:
            if value == current._value:
                return True
            current = current._next
            
        return False
    
    def move_to_front(self, index):
        link = self._get_link_by_index(index)
        self._move_link_to_front(link)

    def move_to_end(self, index):
        """Move item to the end of the list"""
        link = self._get_link_by_index(index)
        self._move_link_to_end(link)
            
    def pushfirst(self, item):

        link = _Link()
        link._value = item
        self._insert_after(link, self)
        
    def pushlast(self, item):

        link = _Link()
        link._value = item
        self._insert_after(link, self._prev)
                
        
class DoublyLinkedListDict(DoublyLinkedList):
    """1/16/2014- I've realized there's no reason
    to have a datacache class without also having the
    base dict class so that I have access to both if I need.

    Just have DataCacheOld inherit and override the one
    or two methods in which it checks its len.

    """
    
    __slots__ = []
    
    def __init__(self, **mappings):
        for k, v in mappings:
            self[k] = v
        
        super().__init__()
    
    def __iter__(self):
        current = self._next
        while current is not self:
            yield (current._key, current._value)
            current = current._next
            
    def items(self):
        return iter(self)
            
    def keys(self):
        current = self._next 
        while current is not self:
            yield current._key
            current = current._next
        
    def values(self):
        current = self._next
        while current is not self:
            yield current._value
            current = current._next
        
    def __delitem__(self, keys):
        link = self._remove_link_by_key(keys)
        link._key = None
        link._value = None
        del link
        
    def __getitem__(self, keys):
        """ Instead of accessing list by index, access by keys.
            while this is O(n), it allows lookup by keys easily."""
        link = self._get_link_by_key(keys)
        return link._value
        
    def __setitem__(self, keys, value):
        """ Instead of accessing list by index, access by keys.
            while this is O(n), it allows lookup by keys easily."""
            
        ''' If keys in self, get the link and set its value,
            and return. '''
        try:
            link = self._get_link_by_key(keys)
        except KeyError:
            pass
        else:
            link._value = value
            return
            
        ''' If link not in self, create a new link'''
        link = _Link()
        link._key = keys
        link._value = value
        self._insert_after(link, self)
  
    def get(self, keys, default=None):
        try:
            return self[keys]
        except KeyError:
            return default
                    
    def popitem(self, keys=None):
        """Pop keys-value pair by keys, or return popfirst"""
        if not self._num_links:
            raise ValueError("Can't pop empty list")
        
        if keys is None:
            return self.popfirst()
        
        link = self._get_link_by_key(keys)
        self._remove_link(link)
        return (link._key, link._value)
            
    def popfirst(self):
        """Pop the first item in the list"""
        if not self._num_links:
            raise ValueError("Can't pop empty list")
            
        link = self._next
        self._remove_link(link)

        return (link._key, link._value)
        
    def poplast(self):
        """Pop the last item in the list"""
        if not self._num_links:
            raise ValueError("Can't pop empty list")
            
        link = self._prev
        self._remove_link(link)
        return (link._key, link._value)
    
    def _get_link_by_key(self, keys):
        """ Used internally to return the link object (instead of just
            value, keys, or keys-value pair."""
        current = self._next
        while current is not self:
            if keys == current._key:
                return current
            current = current._next
        raise KeyError("%s not found" % keys)
        
    def __contains__(self, keys):
        """Iterate through self and return true if keys is found"""
        current = self._next
        while current is not self:
            if keys == current._key:
                return True
            current = current._next
            
        return False
        
    def _remove_link_by_key(self, keys):
        link = self._get_link_by_key(keys)
        self._remove_link(link)
        
        return link
        
    def move_to_front(self, keys):
        link = self._get_link_by_key(keys)
        self._move_link_to_front(link)

    def move_to_end(self, keys):
        """Move item to the end of the list"""
        link = self._get_link_by_key(keys)
        self._move_link_to_end(link)
        
    def __str__(self):
        super_repr = super().__repr__()
        return ''.join((super_repr, '\n', "DoublyLinkedList:", str(list(self)), '\n'))
        
        
class DataCacheOld(DoublyLinkedListDict):
    """Simple subclass to automatically implement limited caching
       behavior.

       If items < 10? set item. Items > 10? Pop the least-recently
       accessed item.

       Todo: heuristic for determining access?
       last lookup, # of lookups, recent lookup bias

       append/pop : right side == first accessed
       appendleft/popleft : left side == last accessed

       creating a doubly linked list because it should use less
       mem than a dict (by a lot)

       Update 1/13/2014:
       instead of a dictionary, this is now a doubly linked list. This
       should save a significant amount of memory.
    """
    
    __default_max_cache = 10
    __slots__ = [
                 '_max_cache'
                 ]
    
    def __init__(self, **mappings):
        self._max_cache = self.__default_max_cache
        super().__init__(**mappings)
    
    def setMaxCache(self, max_val):
        
        if max_val < 0:
            raise ValueError("Cache value must be positive")
            
        self._max_cache = max_val
        
        while self._num_links > self._max_cache:
            self.poplast()
        
    def __getitem__(self, keys):
        """ Instead of accessing list by index, access by keys.
            while this is O(n), it allows lookup by keys easily."""
        link = self._get_link_by_key(keys)
        self._move_link_to_front(link)
        return link._value
        
    def __setitem__(self, keys, value):
        """ Instead of accessing list by index, access by keys.
            while this is O(n), it allows lookup by keys easily."""
            
        ''' If keys in self, get the link and set its value,
            and return. '''
        try:
            link = self._get_link_by_key(keys)
        except KeyError:
            pass
        else:
            self._move_link_to_front(link)
            link._value = value
            return
            
        ''' If link not in self, create a new link'''
        link = _Link()
        link._key = keys
        link._value = value
        self._insert_after(link, self)
        
    @property
    def max_cache(self):
        return self._max_cache

    def _check_len(self):
        while self._num_links > self._max_cache:
            self.poplast()

    def _insert_after(self, link, prev_link):
        
        """Insert the given link object after the other
        given link"""
        
        next_link = prev_link._next
        
        prev_link._next = link
        link._prev = prev_link
        link._next = next_link
        next_link._prev = link
        
        self._num_links += 1
        self._check_len()

from collections import OrderedDict


class DataCache(OrderedDict):
    
    """Ordered dict only uses ~ 3x the memory
    of DoublyLinkedListDict, but has instant
    access to keys instead of some horrible lookup
    time. Screw this experiment, just use ordereddict.

    """
    
    __default_max_cache = 10
    
    def __init__(self, **mappings):
        super().__init__(mappings)
        self._max_cache = self.__default_max_cache
    
    def setMaxCache(self, value):
        self._max_cache = value
        
    def __getitem__(self, key, _dict_get_=dict.__getitem__):
        self.move_to_end(key, False)
        result = _dict_get_(self, key)

        return result
        
    def __setitem__(self, key, value, _dict_set_=OrderedDict.__setitem__):
        _dict_set_(self, key, value)
        self._check_len()
            
    def _check_len(self, _dict_popitem_=OrderedDict.popitem):
        while len(self) > self._max_cache:
            _dict_popitem_(self, True)
            
    def __getnomove__(self, key, _dict_get_=dict.__getitem__):
        """Provide a way to get an item without moving to end"""
        
        return _dict_get_(self, key)
            
    def items(self):
        """Iterating through keys, items or values gets screwed
        up into an infinite loop, because it forces a call
        to getitem, which causes iter->next to always call
        the recently accessed item instead of the actual next.

        So, items(), keys(), and values() need to all be
        reimplemented here to get a full list of keys
        before attempting to return them.
        """
        
        keys = tuple(iter(self))
        for key in keys:
            yield (key, self.__getnomove__(key))
            
    def values(self):
        keys = tuple(iter(self))
        for key in keys:
            yield self.__getnomove__(key)
            
    def keys(self):
        keys = tuple(iter(self))
        for key in keys:
            yield key
            
    @property
    def max_cache(self):
        return self._max_cache
        

        
if __name__ == '__main__':
    from collections import OrderedDict
    test_key = "testkey %d"
    test_value = "testvalue %d"
    
    def cache_access_test():

        test = DataCache()
        print("base test", test)
        test['Foo'] = 'Foovalue'
        print("setitem test", test)
        print("pop item test", test.popitem())
        try:
            print("pop empty test", test.popitem())
        except KeyError:
            pass
        else:
            raise AssertionError
            
        for i in range(10):
            test[test_key % i] = test_value % i
            
        print("Iteration Test")
        for i in test:
            print(i)
        
        print("Iterating after adding too many links")
        for i in range(100):
            test[test_key % i] = test_value % i
            
        for i in test:
            print(i)
            
        print("Inserting existing item")
        test[test_key % 95] = test_value % 95
        
        for i in test.items():
            print(i)
        
        test.setMaxCache(20)
        
        ref = OrderedDict()
        
        for i in range(20):
            test[test_key % i] = test_value % i
            ref[test_key % i] = test_value % i
        
        
        print("Iter test")
        print("basic iter")
        for k in test:
            print(k)
            
        print("Values")

        for v in test.values():
            print(v)
        
        print("Keys")
        for k in test.keys():
            print(k)
            assert test[k] == ref[k]
            
        for k, v in test.items():
            print(k, v)
            
        print("Size Test")
        print(sizeof(ref))
        print(sizeof(test))
        print(sizeof(_Link()))
        
    cache_access_test()
    
    print("Access Time Test")
    from time import perf_counter as timer
    from random import randrange
    
    def time_test(container, key_count, key_range, randrange=randrange, timer=timer):
        t1 = timer()
        for _i in range(key_count):
            keys = test_key % randrange(key_range)
            container[keys]
        t2 = timer()
        
        return (t2 - t1)
        
    key_count = (10, 100, 1000, 10000, 1000000)
    
    cache_results = []
    linked_dict_results = []
    ref_results = []
    
    for keys in key_count:
        
        cache = DataCache() 
        LinkedDict = DoublyLinkedListDict()
        ref = OrderedDict()
        cache.setMaxCache(keys)
        print("\n")
        print("Access Time Test %d lookups" % keys)
        print("Setting up containers...")
        test_list = []
        for i in range(keys):
            cache[test_key % i] = test_value % i
#             LinkedDict[test_key % i] = test_value % i
            ref[test_key % i] = test_value % i
#             test_list.append(test_value % i)
        
        print("")
        print("Doing Test...")
        cache_result = time_test(cache, 1000, keys)
        cache_results.append(cache_result)
        
#         linked_dict_result = time_test(LinkedDict, 1000, keys)
#         linked_dict_results.append(linked_dict_result)
        
        ref_result = time_test(ref, 1000, keys)
        ref_results.append(ref_result)
        
        print("cache: ", cache_result, int(cache_result/ref_result))
#         print("LinkedDict", linked_dict_result, int(linked_dict_result/ref_result))
        print("OrderedDict", ref_result)
        
        
        print(sizeof(cache))
#         print(sizeof(LinkedDict))
        print(sizeof(ref))
        print(sizeof(test_list))
        print(sizeof(set(test_list)))
        print(sizeof(frozenset(test_list)))
        
    
    
    
    
    
    
    
        
