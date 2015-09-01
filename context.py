# -*- coding: utf-8 -*-
"""
Created on Thu Aug 27 10:45:29 2015

@author: МакаровАС
"""

import datetime, dateutil.parser, re
from win32com.client import Dispatch

#context_app, context_wb, context_sh = None, None, None
#macro = None

#EXCEL
#Constants
xlTopToBottom = 1
#--
xlSortNormal, xlSortTextAsNumbers = 0, 1 #XlSortDataOption
xlPinYin, xlStroke = 1, 2 #XlSortMethod
xlSortOnValues, xlSortOnCellColor, xlSortOnFontColor, xlSortOnIcon = range(4) #XlSortOn
xlAscending, xlDescending = 1, 2 #XlSortOrder
xlGuess, xlYes, xlNo = 0, 1, 2 #XlYesNoGuess

#POWERPOINT
#PpSlideSizeType
ppSlideSizeOnScreen = 1

#MS OFFICE
#MsoTriState
msoCTrue = 1
msoFalse = 0
msoTriStateMixed = -2
msoTriStateToggle = -3
msoTrue = -1

#Interaction
CreateObject = Dispatch

#Information
def TypeName(obj):
    name = obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    if name.startswith("_"): name = name[1:] #FIXME: dirty hack?
    return name
    
#DateTime
def DateValue(s):
    "Returns datetime.date from string or pywintypes.datetime"
    try:
        return dateutil.parser.parse(s.strip()).date() if type(s) is str \
            else datetime.date(s.year, s.month, s.day)
    except: pass

#VBA
def Like(s, p):
    "Check string to match RegExp"
    return re.compile(p+r"\Z").match(s) is not None
    
class OfficeApp():
    def __getattr__(self, name): return context_app.__getattr__(name)
    def __call__(self, *args, **kwargs): return context_app.__call__(*args, **kwargs)
    def __setattr__(self, name, value): return context_app.__setattr__(name, value)
Application = OfficeApp()
App = Application #short for Application

def context(doc, module):
    excel_app_ctx = "Selection", "ActiveSheet", "ActiveWindow", "ActiveCell", "Range", "Cells", "Intersect"
    app_ctxs = {"Microsoft Excel": excel_app_ctx}
    global context_app
    context_app = doc.Parent
    app_ctx = app_ctxs[context_app.Name]
    for i in app_ctx:
        setattr(module, i, getattr(context_app, i))
