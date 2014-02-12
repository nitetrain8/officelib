"""
Created on Jan 10, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather


Created to handle managing of automatically
generating property aliases.

This can be done by hand, but this module
provides an easier way to create them
while ensuring safety, and auto-management of
weak/strong aliases when developing interfaces.
"""


class AliasError(Exception):
    """Base exception for alias errors"""
    pass


class AliasExistsError(AliasError):
    alias_notice = "Attribute '%s' already exists for property '%s.%s'"
    
    def __init__(self, alias, klass, prop):
        msg = self.alias_notice % (alias, klass, prop)
        super().__init__(msg)


class StrongAliasError(AliasError):
    
    def __init__(self, alias, prop):
        notice = "Error: alias already exists trying to create alias %s for %s"
        msg = notice % (alias, prop)
        super().__init__(msg)
    
     
class AliasCreationError(AliasError):
    pass


def _createPropertyAlias(cls, properties, namerule=None, raise_on_conflict=True):
    
    """Operate on a class to make aliases.

    Cls- Class object to operate on.

    properties- Map attribute name to a list of aliases.

                If namerule is None, must be a dict which maps
                attribute names to an iterable of aliases to use.

                If namerule is not None, properties is accessed
                as an iterable of properties to pass to namerule. If
                a dict is passed, attribute names are pulled from
                properties.keys()

    namerule-   Optional callback function that takes an attribute name
                and returns an iterable of aliases to use.

                Default is None. If None, properties is accessed as a dict
                to map names to aliases.

                If not None, any mapping of attribue names to aliases
                will ignored.

    raise_on_confict-
                If a naming conflict arises, raise AliasExistsError.
                If set to false, silently ignore conflict and move on.


    """

    if namerule is None:
        attrs = properties.keys()
        
        for attr in attrs:
            value = getattr(cls, attr)
            
            for alias in properties[attr]:
                
                if hasattr(cls, alias):
                    
                    if raise_on_conflict:
                        raise AliasExistsError(alias, cls.__name__, attr)
                
                else:
                    setattr(cls, alias, value)
    else:
        
        #get attrs as keys of dict, or any iterable
        try:
            attrs = properties.keys()
        
        except AttributeError:
            attrs = properties
            
        for attr in attrs:
            
            value = getattr(cls, attr)
            alias_list = namerule(attr)
            
            for alias in alias_list:
                
                if hasattr(cls, alias):
                    
                    if raise_on_conflict:
                        raise AliasExistsError(alias, cls.__name__, attr)
                
                else:
                    setattr(cls, alias, value)   


class AliasManager():
    """ Class to encapsulate management of alias
        creation

        This is kind of tricky, since we can't use a metaclass
        for applying aliases, but we have to pop all alias
        references from the class dict before creating the class
        so this requires cooperation from the interface.

    """

    def __init__(self, *, weak_alias=None, 
                            weak_alias_namerule=None, 
                            strong_alias=None,
                            strong_alias_namerule=None):
        
        _g = self._generate_from_rule
        
        if weak_alias_namerule:
            self.weak_alias_map = _g(weak_alias,
                                    weak_alias_namerule)
            self.weak_alias_namerule = weak_alias_namerule
        else:
            self.weak_alias_map = weak_alias or {}
        
        if strong_alias_namerule:
            self.strong_alias_map = _g(strong_alias, 
                                        strong_alias_namerule)
                                        
            self.strong_alias_namerule = strong_alias_namerule
        else:
            self.strong_alias_map = strong_alias or {}
        
    def weak_alias(self, prop, alias_or_list):
        
        if self.weak_alias_map.get(prop, None) is None:
            self.weak_alias_map[prop] = []
        
        try:
            self.weak_alias_map[prop].update(alias for alias in alias_or_list)
        except TypeError:
            self.weak_alias_map[prop].append(alias_or_list)
            
    def strong_alias(self, prop, alias_or_list):
        
        if self.strong_alias_map.get(prop, None) is None:
            self.strong_alias_map[prop] = []
        
        try:
            self.strong_alias_map[prop].update(alias for alias in alias_or_list)
        except TypeError:
            self.strong_alias_map[prop].append(alias_or_list)

    def _generate_from_rule(self, properties, rule):
        return {prop : rule(prop) for prop in properties}
   
  
class AliasDict(dict):
    
    def __init__(self, *args, **kwargs):
        dict.__init__(self, *args, **kwargs)
        self._alias_map = {}
    
    def __setitem__(self, keys, value, dict_setitem_=dict.__setitem__):
        
        try:
            good_key = self._alias_map[keys]
        except KeyError:
            good_key = keys
        return dict_setitem_(self, good_key, value)
        
    def __getitem__(self, keys, dict_getitem_=dict.__getitem__):
        try:
            good_key = self._alias_map[keys]
        except KeyError:
            good_key = keys
        return dict_getitem_(self, good_key)
    
    def update_map(self, mapping):
        self._alias_map.update(mapping)
        
    def set_map(self, mapping):
        self._alias_map = mapping
        
        
class AbstractAliasHandler(type):
    
    def __new__(cls, name, bases, kwargs):
        
        managers = []
        for k,v in kwargs.items():
            if isinstance(v, AliasManager):
                managers.append(kwargs.pop(k))
                
        manager = managers[0]
        extras = managers[1:]
        for extra in extras:
            manager.weak_alias_map.update(extra.weak_alias_map)
            manager.strong_alias_map.update(extra.strong_alias_map)
            
        new_cls = super().__new__(cls, name, bases, kwargs)
        
        weak = manager.weak_alias_map
        strong = manager.strong_alias_map
        
        _createPropertyAlias(new_cls, weak, None, False)
        _createPropertyAlias(new_cls, strong, None, True)
        
        return new_cls
            
    
    
            
    
    
    
    
    
    
    
    
    
