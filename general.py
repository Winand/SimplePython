# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 18:44:38 2015

@author: Winand
"""

SOURCEDIR, DEF_MODULE = "source", "default"
macro_tree, modules = {}, {}
comobj_cache = {}

from win32com.client import gencache
import pythoncom, string, sys, datetime, builtins, threading
from functools import wraps
import context
import cProfile, pstats, io #Profiling
from threaded_ui import app

run_lock = threading.Semaphore(2) #for macro interruption
              
COL = {} #dict of column names
for i in string.ascii_uppercase:
    COL[i] = ord(i)-ord("A")+1
    COL["A"+i] = 26+ord(i)-ord("A")+1
    
def print(*args, **kwargs):
    "print with timestamp"
    builtins.print(datetime.datetime.now().strftime("%M:%S|"), *args, **kwargs)

def optional_arguments(f):
    "Allows to use any decorator with or w/o arguments \
    http://stackoverflow.com/a/14412901/1119602"
    @wraps(f)
    def new_dec(*args, **kwargs):
        if len(args) == 1 and len(kwargs) == 0 and callable(args[0]):
            return f(args[0])
        else: return lambda realf: f(realf, *args, **kwargs)
    return new_dec
    
@optional_arguments    
def macro(func, for_=context.Office):
    if func.__module__.startswith(SOURCEDIR+"."):
        module = func.__module__[len(SOURCEDIR+"."):]
    else: module = func.__module__
    if module not in macro_tree:
        macro_tree[module] = {}
    if func.__name__ not in macro_tree[module]: #in case of duplicate macro names
        macro_tree[module][func.__name__] = for_
        
    @wraps(func)
    def wrapper(doc):
        doc_obj = getOpenedFileObject(doc)
        if doc_obj:
            context.context(doc_obj, modules[module])
            try:
                with run_lock:
#                    with Profile():
                    ret = func()
                    while not run_lock._value: pass #prevent interruption outside try block
                    return ret
            except KeyboardInterrupt:
                print("Macro '%s' interrupted"%func.__name__)
            except Exception as e:
                frame = sys.exc_info()[2].tb_next
                if not frame: frame = sys.exc_info()[2]
                print("Exception in macro '%s': %s[%s:%d] %s" %
                    (func.__name__, type(e).__name__, module, frame.tb_lineno, e))
                showConsole() #Show exception to user
            finally:
                try: context.App.ScreenUpdating = True #Turn back updating!
                except: print("Failed to turn on screen updating!")
        else: print("Opened document '%s' not found"%doc)
    return wrapper
    
def getOpenedFileObject(name):
    if name in comobj_cache:
        try:
            #http://stackoverflow.com/questions/3500503/check-com-interface-still-alive
            o = comobj_cache[name]
            o._oleobj_.GetTypeInfoCount()
            return o
        except: del comobj_cache[name]
    ctx, rot = pythoncom.CreateBindCtx(0), pythoncom.GetRunningObjectTable()
    for i in rot:
        if i.GetDisplayName(ctx, None) == name:
            comobj_cache[name] = gencache.EnsureDispatch(
                    rot.GetObject(i).QueryInterface(pythoncom.IID_IDispatch))
            return comobj_cache[name]
    
def getMacroList(app):
    l = [f if m==DEF_MODULE else m+"."+f for m in macro_tree for f in macro_tree[m] if macro_tree[m][f] in (app, context.Office)]
    return sorted(l, key=lambda s: s.lower())
    
class Profile():        
    def __enter__(self):
        self.pr = cProfile.Profile()
        self.pr.enable()
        
    def __exit__(self, exc_type, exc_value, traceback):
        self.pr.disable()
        s = io.StringIO()
        pstats.Stats(self.pr, stream=s).sort_stats('cumulative').print_stats()
        print(s.getvalue())
        
def showConsole():
    app().form.showWindow(console=True)

context.macro = macro #so macro can be imported from context
context.print = lambda *args, **kwargs: print(">", *args, **kwargs) #macro prints are marked with ">" sign