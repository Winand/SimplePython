# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 13:20:59 2015

@author: Winand
"""
from general import getMacroList, DEF_MODULE, SOURCEDIR, macro_tree, modules, pythoncom, print
from general import run_lock, interrupt_lock
import struct, json, pathlib
import importlib, _thread
import win32file, win32con, winnt
from threaded_ui import QtApp, invoke, QtCore, QtGui, app, isConsoleApp
import sys, threading, time

import socketserver
SHARED_SERVER_ADDR, SHARED_SERVER_PORT = "127.0.0.1", 50002
appPath = pathlib.Path(__file__).absolute().parent
            
class TCPServer(socketserver.TCPServer):
    allow_reuse_address = True
    
def macro_caller(macro, *args, **kwargs):
    pythoncom.CoInitialize()
    macro(*args, **kwargs)
    
class Handler(socketserver.BaseRequestHandler):
    def sendString(self, s):
        if s:
            msg = s.encode("utf-8")
            self.request.sendall(struct.pack('i', len(msg)))
            self.request.sendall(msg)
        
    def sendByte(self, b):
        self.request.sendall(struct.pack('b', b))
        
    def recvString(self):
        msgl = struct.unpack('i', self.request.recv(4))[0]
        return self.request.recv(msgl).decode("utf-8")
        
    def handle(self):
        msg = json.loads("{%s}"%self.recvString())
        self.sendString(self.handle_wrap(msg["Type"], msg))
        
    def handle_wrap(self, msg, args):
        if msg in ("Call", "Doc"):
            path = args["Macro"].rsplit(".", 1)
            if len(path)==1: path.insert(0, DEF_MODULE)
            if path[0] in macro_tree and path[1] in macro_tree[path[0]]:
                macro = getattr(modules[path[0]], path[1], None)
            else: macro = None
            if not macro: return "SERVER_NO_MACRO"
            elif msg == "Doc":
                print("Help on '%s' requested"%args["Macro"])
                return macro.__doc__ or ""
            else: #Caller must 'Test' if server is 'OK' before this
                print("Macro '%s' called "%path[1])
                invoke(app().form.startMacro, macro, args["Workbook"])
                return "OK"
        elif msg == "Request":
            print("Macro list requested")
            return "|".join(getMacroList())
        elif msg == "Test":
            print("Connection checked")
            return "Busy" if run_lock.locked() else "OK"
        elif msg == "Interrupt":
            print("Request to interrupt macro")
            with interrupt_lock: #FIXME: It's possible to make several requests before macro is interrupted. Is it ok?
                if run_lock.locked():
                    _thread.interrupt_main()
                else: return "Not busy"
            return "OK"
        else:
            print("Unknown message")
            return "Unknown"
       
watch, loader_lock = {}, threading.RLock()
def initModuleLoader():
    src_path = str(appPath.joinpath("source"))
    def watcher():
        def readChanges(h, flags):
            return win32file.ReadDirectoryChangesW(h, 1024, True, flags, None, None)
        actions = {1: "add", 2: "del", 3: "add", 4: "del", 5: "add"} #3 - update; 4, 5 - rename
        hDir = win32file.CreateFile(src_path, winnt.FILE_LIST_DIRECTORY,
            win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE,
            None, win32con.OPEN_EXISTING, win32con.FILE_FLAG_BACKUP_SEMANTICS, None)
        while True:
            for action, file in readChanges(hDir, win32con.FILE_NOTIFY_CHANGE_FILE_NAME|
                                            win32con.FILE_NOTIFY_CHANGE_LAST_WRITE):
                filename = pathlib.Path(file)
                if filename.suffix == ".py" and not filename.stem.startswith("__"):
                    with loader_lock:
                        watch[filename.stem] = actions[action]
    def unload(mod_name):
        if mod_name in modules:
#            m = modules[mod_name]
            del modules[mod_name], macro_tree[mod_name], sys.modules[SOURCEDIR+"."+mod_name]
#            print(sys.getrefcount(m))
            
    def import_mod(mod_name):
        try:
            macro_tree[mod_name] = []
            modules[mod_name] = importlib.import_module(SOURCEDIR+"."+mod_name)
            print("Module '%s' updated"%mod_name)
        except Exception as e:
            print("Failed to update '%s' module: %s: %s"%(mod_name, type(e).__name__, e))
            del macro_tree[mod_name]
            
    def reloader():
        while True:
            with loader_lock:
                for i in watch:
                    if watch[i] == "add":
                        if i in modules: unload(i)
                        import_mod(i)
                    elif watch[i] == "del":
                        unload(i)
                        print("Module '%s' unloaded"%i)
                    invoke(app().form.updateMacroTree)
                watch.clear()
            time.sleep(1)
    for i in [f for f in pathlib.Path(src_path).iterdir() if f.is_file() and
                            f.suffix == ".py" and not f.stem.startswith("__")]:
        watch[i.stem] = "add"
    for i in threading.Thread(target=watcher), \
            threading.Thread(target=reloader):
        i.daemon = True
        i.start()

class TempTrayIcon:
    "Temporary changes tray icon"
    def __init__(self, systray, tempicon):
        f = QtGui.qApp.style().standardIcon \
            if type(tempicon) is QtGui.QStyle.StandardPixmap else QtGui.QIcon
        self.tray = systray
        self.oldicon, self.newicon = self.tray.icon(), f(tempicon)
        
    def __enter__(self):
        self.tray.setIcon(self.newicon)
        
    def __exit__(self, *args):
        self.tray.setIcon(self.oldicon)
        
class SimplePython(QtGui.QWidget):
    def __init__(self):
        initModuleLoader()
        self.server = TCPServer((SHARED_SERVER_ADDR, SHARED_SERVER_PORT), Handler)
        threading.Thread(target=self.server.serve_forever).start()
        self.tray.addMenuItem("Exit", self.btnExit_clicked)
        self.terminated.connect(self.btnExit_clicked)

    def tray_activated(self, reason):
        if reason == QtGui.QSystemTrayIcon.Trigger:
            self.showWindow()
            
    def showWindow(self, console=False):
        if console and isConsoleApp():
            return
        self.show()
        self.activateWindow()
        if console:
            self.tabs.setCurrentIndex(0)
            
    def btnClear_clicked(self):
        self.txtConsole.clear()
        
    @QtCore.pyqtSlot()
    def updateMacroTree(self):
#        d=QtGui.QDialog()
#        d.setWindowModality(1)
#        d.exec()
        ret = ""
        for m in macro_tree:
            ret += "» "+m+"\n"
            for i, j in enumerate(macro_tree[m]):
                ret += ("└ " if i==len(macro_tree[m])-1 else "├ ")+j+"\n"
            ret += "\n"
        self.txtModules.setPlainText(ret)
        
    def btnExit_clicked(self):
        self.server.shutdown()
        self.server.server_close()
        QtGui.qApp.quit()
    
    def closeEvent(self, event):
        if event.type() == event.Close:
            event.ignore()
            self.hide()
            
    @QtCore.pyqtSlot(object, object)
    def startMacro(self, macro, wb):
        with TempTrayIcon(app().form.tray,
                          QtGui.QStyle.SP_ArrowRight):
            macro(wb)
            
stdout = None if isConsoleApp() else "txtConsole" #redirect output if no console
QtApp(SimplePython, ontop=True, hidden=True, stdout=stdout, 
      tray={"icon": str(appPath.joinpath(r"res\icon.png")), "tip": "SimplePython Server"})
