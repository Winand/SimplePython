# -*- coding: utf-8 -*-
"""
Created on Thu Aug 27 10:45:29 2015

@author: МакаровАС
"""

import datetime, dateutil.parser, re
from win32com.client import gencache
from threaded_ui import Dialog, QtGui

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
#    name = obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    name = obj.__class__.__name__ #Cache must be built
    return name[name.startswith("_"):]
    
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
    
def UserForm(Base):
    "Show user dialog"
    return Dialog(Base, ontop=True)
    
def short(v):
    "Convert int to signed short"
    return v-0x10000 if v>>15 else v
#    return ctypes.c_short(v).value
    
_ = QtGui.QMessageBox
__ic = {vbQuestion: _.Question, vbInformation: _.Information,
        vbExclamation: _.Warning, vbCritical: _.Critical}
__bt = {vbOKCancel: _.Ok|_.Cancel, vbYesNo: _.Yes|_.No, vbRetryCancel: _.Retry|_.Cancel,
        vbYesNoCancel: _.Yes|_.No|_.Cancel, vbAbortRetryIgnore: _.Abort|_.Retry|_.Ignore}   
__retl = {_.Abort: vbAbort, _.Cancel: vbCancel, _.Ignore: vbIgnore, _.No: vbNo,
            _.Ok: vbOK, _.Retry: vbRetry, _.Yes: vbYes}        
def MsgBox(Prompt, Buttons=None, Title=None):
    ic = __ic.get(Buttons&0xf0, None)
    bt = __bt.get(Buttons&0xf, None)
    return __retl.get(_(ic, Title, Prompt, bt).exec(), None)
    
class OfficeApp():
    def __getattr__(self, name): return context_app.__getattr__(name)
    def __call__(self, *args, **kwargs): return context_app.__call__(*args, **kwargs)
    def __setattr__(self, name, value): return context_app.__setattr__(name, value)
Application = OfficeApp()
App = Application #short for Application

def context(doc, module):
    excel_app_ctx = "Selection", "ActiveSheet", "ActiveWorkbook", "ActiveWindow", "ActiveCell", "Range", "Cells", "Intersect", "Workbooks"
    app_ctxs = {"Microsoft Excel": excel_app_ctx}
    global context_app
    context_app = doc.Parent
    app_ctx = app_ctxs[context_app.Name]
    for i in app_ctx:
        try: attr = getattr(context_app, i, None)
        except: attr = None
        setattr(module, i, attr)
