from officelib.wordlib import *

def add_ref(rng, refnum, typ=c.wdRefTypeNumberedItem, kind=c.wdNumberNoContext, hyperlink=True, includepos=False, separate_numbers = False, sep_string = " "):
    rng.InsertCrossReference(typ, kind, refnum, hyperlink, includepos, separate_numbers, sep_string)
   
   
from collections import OrderedDict

def GetNumberedRefs(doc, reftyp=c.wdRefTypeNumberedItem, num_only=True):
    items = doc.GetCrossReferenceItems(reftyp)
    res = OrderedDict()
    for i, it in enumerate(items, 1):
        s = it.strip()
        num, *_ = s.split()
        if num_only:
            res[num] = i
        else:
            res[s] = i
    return res
    
def insert_row(s,n=1):
    s.InsertRowsBelow(n)

def _callback(k, v):
    return k[0].isdigit()

def insert_cross_references(d, callback):
    w = d.Application
    for k, v in GetNumberedRefs(d).items():
        if callback(k, v):
            r = w.Selection.Range
            add_ref(r,v)
            insert_row(w.Selection)
    
    