# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 18:44:38 2015

@author: Winand
"""

SOURCEDIR, DEF_MODULE = "source", "default"
macro_tree, modules = {}, {}
comobj_cache = {}

from win32com.client import Dispatch, gencache
import pythoncom, string, sys, datetime, builtins
from functools import wraps
import context
from threaded_ui import Widget
import cProfile, pstats, io #Profiling
              
COL = {} #dict of column names
for i in string.ascii_uppercase:
    COL[i] = ord(i)-ord("A")+1
    COL["A"+i] = 26+ord(i)-ord("A")+1
    
def print(*args, **kwargs):
    "print with timestamp"
    builtins.print(datetime.datetime.now().strftime("%M:%S|"), *args, **kwargs)

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
            context.context(doc_obj, modules[module])
#            with Profile():            
            try:
                return func()
            except Exception as e:
                frame = sys.exc_info()[2].tb_next
                if not frame: frame = sys.exc_info()[2]
                print("Exception in macro '%s': %s[%s:%d] %s" %
                    (func.__name__, type(e).__name__, module, frame.tb_lineno, e))
            finally:
                context.App.ScreenUpdating = True #Turn back updating!
                
        else: print("Opened document '%s' not found"%doc)
    return wrapper

context.macro = macro #so macro can be imported from context
    
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
            gencache.EnsureDispatch(comobj_cache[name]) #FIXME: is it good way to build cache?
            return comobj_cache[name]
    
def getMacroList():
    return [f if m==DEF_MODULE else m+"."+f for m in macro_tree for f in macro_tree[m]]
    
class Profile():        
    def __enter__(self):
        self.pr = cProfile.Profile()
        self.pr.enable()
        
    def __exit__(self, exc_type, exc_value, traceback):
        self.pr.disable()
        s = io.StringIO()
        pstats.Stats(self.pr, stream=s).sort_stats('cumulative').print_stats()
        print(s.getvalue())
