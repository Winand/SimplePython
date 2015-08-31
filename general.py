# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 18:44:38 2015

@author: Winand
"""

SOURCEDIR, DEF_MODULE = "source", "default"
macro_tree, modules = {}, {}
comobj_cache = {}

from win32com.client import Dispatch
import pythoncom, string, re, sys, datetime
from functools import wraps
import context
from threaded_ui import Widget
    
ctx_app_list = ["Selection", "ActiveSheet", "ActiveWindow", "ActiveCell"]
              
COL = {} #dict of column names
for i in string.ascii_uppercase:
    COL[i] = ord(i)-ord("A")+1
    COL["A"+i] = 26+ord(i)-ord("A")+1
    
import builtins
__print_def = builtins.print
def __print(*args, **kwargs):
    "print with timestamp"
    __print_def(datetime.datetime.now().strftime("%M:%S|"), *args, **kwargs)
builtins.print = __print

def Like(s, p):
    "Check string to match RegExp"
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
            with Context(doc_obj, modules[module]):
                try:
                    return func()
                except Exception as e:
                    frame = sys.exc_info()[2].tb_next
                    if not frame: frame = sys.exc_info()[2]
                    print("Exception in macro '%s': %s[%s:%d] %s" %
                        (func.__name__, type(e).__name__, module, frame.tb_lineno, e))
                finally:
                    context.context_app.ScreenUpdating = True #Turn back updating!
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
    
class Context():
    def __init__(self, doc, module):
        self.doc, self.module = doc, module
        
    def __enter__(self):
        app = self.doc.Parent
        context.context_app, context.context_wb, context.context_sh = app, self.doc, self.doc.ActiveSheet
        for i in ctx_app_list:
            attr = getattr(app, i, None)
            if attr: setattr(self.module, i, attr)
        
    def __exit__(self, exc_type, exc_value, traceback):
        for i in ctx_app_list:
            if hasattr(self.module, i): delattr(self.module, i)

context.macro = macro #so macro can be imported from context
            