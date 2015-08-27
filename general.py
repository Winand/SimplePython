# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 18:44:38 2015

@author: Winand
"""

SOURCEDIR, DEF_MODULE = "source", "default"
macro_tree, modules = {}, {}
comobj_cache = {}

from win32com.client import Dispatch
import pythoncom, string, re, datetime
from functools import wraps

def _typename(obj):
    name = obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    if name.startswith("_"): name = name[1:] #FIXME: dirty hack?
    return name
    
def _datevalue(s, f):
    try: return datetime.datetime.strptime(s, f).date()
    except: pass
    
ctx_doc_list = ["Range", "Cells"] #active sheet context
ctx_app_list = ["Selection", "Intersect", "ActiveSheet", "ActiveWindow", "ActiveCell"]
ctx_custom = {"TypeName": _typename, "CreateObject": Dispatch, "msoFalse": 0,
              "xlSortOnValues": 0, "xlDescending": 2, "xlSortNormal": 0,
              "xlGuess": 0, "xlTopToBottom": 1, "xlPinYin": 1,
              "ppSlideSizeOnScreen": 1, "DateValue": _datevalue}
              
COL = {} #dict of column names
for i in string.ascii_uppercase:
    COL[i] = ord(i)-ord("A")+1
    COL["A"+i] = 26+ord(i)-ord("A")+1
    
def Like(s, p):
    return re.compile(p+r"\Z").match(s) is not None

def macro(func):
    if func.__module__.startswith(SOURCEDIR+"."):
        module = func.__module__[len(SOURCEDIR+"."):]
    else: module = func.__module__
    if module not in macro_tree:
        macro_tree[module] = []
    if func.__name__ not in macro_tree[module]: #in case of duplicate macro names
        macro_tree[module].append(func.__name__)
    @wraps(func)
    def wrapper(doc):
        doc_obj = getOpenedFileObject(doc)
        if doc_obj:
            with context(doc_obj, modules[module]):
                try:
                    return func(doc_obj)
                except Exception as e:
                    print("Macro '%s' failed with exception: %s"%(func.__name__, e))
        else: print("Opened document '%s' not found"%doc)
        
    return wrapper
    
def getOpenedFileObject(name):
    if name in comobj_cache:
        try:
            o = comobj_cache[name]
            o._oleobj_.GetTypeInfoCount()
            return o
        except: del comobj_cache[name]
    ctx, rot = pythoncom.CreateBindCtx(0), pythoncom.GetRunningObjectTable()
    for i in rot:
        if i.GetDisplayName(ctx, None) == name:
            comobj_cache[name] = Dispatch(rot.GetObject(i).QueryInterface(pythoncom.IID_IDispatch))
            return comobj_cache[name]
    
def getMacroList():
    return [f if m==DEF_MODULE else m+"."+f for m in macro_tree for f in macro_tree[m]]
    
class context():
    def __init__(self, doc, module):
        self.doc, self.module = doc, module
        
    def __enter__(self):
        active, app = self.doc.ActiveSheet, self.doc.Parent
        for i in ctx_doc_list:
            attr = getattr(active, i, None)
            if attr: setattr(self.module, i, attr)
        for i in ctx_app_list:
            attr = getattr(app, i, None)
            if attr: setattr(self.module, i, attr)
        for i in ctx_custom:
            setattr(self.module, i, ctx_custom[i])
        
    def __exit__(self, exc_type, exc_value, traceback):
        for i in ctx_doc_list+ctx_app_list+list(ctx_custom.keys()):
            if hasattr(self.module, i): delattr(self.module, i)
            