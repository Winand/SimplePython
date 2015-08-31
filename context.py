# -*- coding: utf-8 -*-
"""
Created on Thu Aug 27 10:45:29 2015

@author: МакаровАС
"""

import datetime, dateutil.parser
from win32com.client import Dispatch

context_app, context_wb, context_sh = None, None, None
macro = None

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

CreateObject = Dispatch

def TypeName(obj):
    name = obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    if name.startswith("_"): name = name[1:] #FIXME: dirty hack?
    return name
    
def DateValue(s):
    "Returns datetime.date from string or pywintypes.datetime"
    try:
        return dateutil.parser.parse(s.strip()).date() if type(s) is str \
            else datetime.date(s.year, s.month, s.day)
    except: pass
    
def Range(*args, **kwargs):
    return context_sh.Range(*args, **kwargs)
    
def Cells(*args, **kwargs):
    return context_sh.Cells(*args, **kwargs)
    
def Intersect(*args, **kwargs):
    return context_app.Intersect(*args, **kwargs)
