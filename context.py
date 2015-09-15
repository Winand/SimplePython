# -*- coding: utf-8 -*-
"""
Created on Thu Aug 27 10:45:29 2015

@author: МакаровАС
"""

import datetime, dateutil.parser, re
from win32com.client import gencache
from threaded_ui import Dialog, QtGui, QtCore

def ExtractValues(func):
    "Decorator: find arguments with Value attribute and replace them with values"
    def func_wrapper(*args, **kwargs):
        args = [i.Value if hasattr(i, "Value") else i for i in args]
        kwargs = {i: kwargs[i].Value if hasattr(kwargs[i], "Value") else kwargs[i] for i in kwargs}
        return func(*args, **kwargs)
    return func_wrapper

#EXCEL
#Constants
xlTopToBottom = 1
#--
xlSortNormal, xlSortTextAsNumbers = 0, 1 #XlSortDataOption
xlPinYin, xlStroke = 1, 2 #XlSortMethod
xlSortOnValues, xlSortOnCellColor, xlSortOnFontColor, xlSortOnIcon = range(4) #XlSortOn
xlAscending, xlDescending = 1, 2 #XlSortOrder
xlGuess, xlYes, xlNo = 0, 1, 2 #XlYesNoGuess
vbInformation, vbExclamation, vbQuestion, vbCritical, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel = 64, 48, 32, 16, 1, 2, 3, 4, 5 #VbMsgBoxStyle
vbAbort, vbCancel, vbIgnore, vbNo, vbOK, vbRetry, vbYes = 3, 2, 5, 7, 1, 4, 6

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
CreateObject = gencache.EnsureDispatch

#Information
def TypeName(obj):
    name = obj.__class__.__name__ #Cache must be built
    return name[name.startswith("_"):]

@ExtractValues
def RGB(Red, Green, Blue):
    return Blue<<16 | Green<<8 | Red
    
#DateTime
@ExtractValues
def DateValue(s):
    "Returns datetime.date from string or pywintypes.datetime"
    try:
        return dateutil.parser.parse(s.strip()).date() if type(s) is str \
            else datetime.date(s.year, s.month, s.day)
    except: pass

#VBA
@ExtractValues
def Like(s, p):
    "Check string to match RegExp"
    return re.compile(p+r"\Z").match(s) is not None
    
def UserForm(Base):
    "Show user dialog"
    return Dialog(Base, ontop=True)
    
def short(v):
    "Convert int to signed short"
    return v-0x10000 if v>>15 else v
    
_ = QtGui.QMessageBox
__ic = {vbQuestion: _.Question, vbInformation: _.Information,
        vbExclamation: _.Warning, vbCritical: _.Critical}
__bt = {vbOKCancel: _.Ok|_.Cancel, vbYesNo: _.Yes|_.No, vbRetryCancel: _.Retry|_.Cancel,
        vbYesNoCancel: _.Yes|_.No|_.Cancel, vbAbortRetryIgnore: _.Abort|_.Retry|_.Ignore}   
__retl = {_.Abort: vbAbort, _.Cancel: vbCancel, _.Ignore: vbIgnore, _.No: vbNo,
            _.Ok: vbOK, _.Retry: vbRetry, _.Yes: vbYes}        
def MsgBox(Prompt, Buttons=0, Title="SimplePython"):
    ic = __ic.get(Buttons&0xf0, _.NoIcon)
    bt = __bt.get(Buttons&0xf, _.Ok)
    return __retl.get(_(ic, Title, Prompt, bt, flags=QtCore.Qt.WindowStaysOnTopHint).exec(), None)

def context(doc, module):
    "Sets needed global variables to a /module/, sets App variable to itself"
    excel_app_ctx = "Selection", "ActiveSheet", "ActiveWorkbook", "ActiveWindow", "ActiveCell", "Range", "Cells", "Intersect", "Workbooks"
    app_ctxs = {"Microsoft Excel": excel_app_ctx}
    global App
    App = doc.Parent
    app_ctx = app_ctxs[App.Name]
    for i in app_ctx:
        try: attr = getattr(App, i, None)
        except: attr = None
        setattr(module, i, attr)
    setattr(module, "App", App); setattr(module, "Application", App)
