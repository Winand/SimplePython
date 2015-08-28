# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 13:21:53 2015

@author: Winand
"""
    
from context import *

@macro
def unmerge_and_fill(wb):
    "разгруппировать все ячейки в Selection и ячейки каждой \
    бывшей группы заполнить значениями из их первых ячеек"
    if TypeName(Selection) != "Range": return
    elif Selection.Cells.Count == 1: return
    for cell in Intersect(Selection, ActiveSheet.UsedRange).Cells:
        if cell.MergeCells:
            Address = cell.MergeArea.Address
            cell.UnMerge()
            Range(Address).Value = cell.Value

def fitFactor(ws, hs, wd, hd):
    f1, f2 = wd / ws, hd / hs
    return f2 if f2 < f1 else f1
            
@macro
def ЭкспортВПрезентацию(wb):
    SLIDE_MARGIN = 8
    pp = CreateObject("PowerPoint.Application")
    pr = pp.Presentations.Add()
    pr.PageSetup.SlideSize = ppSlideSizeOnScreen #4:3
    blank = pr.Designs(1).SlideMaster.CustomLayouts(7)
    for i in ActiveWindow.SelectedSheets:
        if TypeName(i) == "Chart":
            i.ChartArea.Copy()
            sl = pr.Slides.AddSlide(pr.Slides.count + 1, blank)
            sl.Select()
            pp.ActiveWindow.View.PasteSpecial(Link=msoFalse)
            sh = sl.Shapes(1)
            sh.LinkFormat.BreakLink()
            f = fitFactor(sh.Width, sh.Height, 
                          pr.PageSetup.SlideWidth - SLIDE_MARGIN * 2,
                          pr.PageSetup.SlideHeight - SLIDE_MARGIN * 2)
            sh.Left = sh.Left + sh.Width * (1 - f) / 2
            sh.Top = sh.Top + sh.Height * (1 - f) / 2
            sh.ScaleHeight(f, msoFalse)
            sh.ScaleWidth(f, msoFalse)
    
def clean(org):
    org = org.replace('"', ' ').replace("'", ' ').replace("\x1a", ' ') #
    org = org.replace('(', ' ').replace(")", ' ').replace("/", ' ')
    org = org.replace('\\', ' ').replace("-", ' ').replace(".", ' ')
    org = org.replace(',', ' ').replace("`", ' ')
    while org.count("  "): org = org.replace("  ", ' ')
    org = org.replace(' ж д', '').replace('ж д ', '')
    if org.endswith(" жд"): org = org[:-3]
    return org.strip()    
    
@macro
def я_asrb_format_org(wb):
    'Вспомогательный макрос, форматирует имя организации для добавления в GenerateOrgs'
    def copy(text=None):
        import win32clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        if text: win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
    copy(clean(str(ActiveCell.Value)))
    