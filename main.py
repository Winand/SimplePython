# -*- coding: utf-8 -*-
"""
Created on Thu Aug 20 13:20:59 2015

@author: Winand
"""
from general import getMacroList, DEF_MODULE, SOURCEDIR, macro_tree, modules, pythoncom
import struct, json, pathlib
import importlib
import win32file, win32con, winnt
from threaded_ui import QtApp, QtCore

import socketserver
SHARED_SERVER_ADDR, SHARED_SERVER_PORT = "127.0.0.1", 50002
appPath = pathlib.Path(__file__).absolute().parent

import sys, threading, time
from PyQt4 import QtGui
            
class TCPServer(socketserver.TCPServer):
    allow_reuse_address = True
    
class Handler(socketserver.BaseRequestHandler):
    def sendString(self, s):
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
        if msg["Type"] in ("Call", "Doc"):
            path = msg["Macro"].rsplit(".", 1)
            if len(path)==1: path.insert(0, DEF_MODULE)
            if path[0] in macro_tree and path[1] in macro_tree[path[0]]:
                macro = getattr(modules[path[0]], path[1], None)
            else: macro = None
            if not macro: self.sendString("SERVER_NO_MACRO")
            elif msg["Type"]=="Doc":
                print("Help on '%s' requested"%msg["Macro"])
                self.sendString(macro.__doc__ or "")
            else:
                print("Macro '%s' called"%path[1])
                self.sendString("OK")
                macro(msg["Workbook"])
        elif msg["Type"] == "Request":
            print("Macro list requested")
            self.sendString("|".join(getMacroList()))
        elif msg["Type"] == "Test":
            print("Connection established")
            self.sendString("OK")
        else: print("Unknown message")
       
watch, lock = {}, threading.RLock()
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
                    with lock:
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
            print("Module '%s' (re)loaded"%mod_name)
        except Exception as e:
            print("Failed to (re)load '%s' module: %s: %s"%(mod_name, type(e).__name__, e))
            del macro_tree[mod_name]
            
            
    def reloader():
        while True:
            with lock:
                for i in watch:
                    if watch[i] == "add":
                        if i in modules: unload(i)
                        import_mod(i)
                    elif watch[i] == "del":
                        unload(i)
                        print("Module '%s' unloaded"%i)
                watch.clear()
            time.sleep(1)
    for i in [f for f in pathlib.Path(src_path).iterdir() if f.is_file() and
                            f.suffix == ".py" and not f.stem.startswith("__")]:
        watch[i.stem] = "add"
    for i in threading.Thread(target=watcher), \
            threading.Thread(target=reloader):
        i.daemon = True
        i.start()

class SimplePython(QtGui.QWidget):
    def __init__(self):
        initModuleLoader()
        self.server = TCPServer((SHARED_SERVER_ADDR, SHARED_SERVER_PORT), Handler)
        threading.Thread(target=lambda:(pythoncom.CoInitialize(), self.server.serve_forever())).start()
        def activated(reason):
            if reason == QtGui.QSystemTrayIcon.Trigger:
                self.show()
                self.activateWindow()
        self.tray.activated.connect(activated)
        self.tray.contextMenu().addAction("Exit").triggered.connect(self.quit_)
        self.btnExit.clicked.connect(self.quit_)
        self.b.clicked.connect(self.update)
        
    def update(self):
        ret = ""
        for m in macro_tree:
            ret += "» "+m+"\n"
            for i, j in enumerate(macro_tree[m]):
                ret += ("└ " if i==len(macro_tree[m])-1 else "├ ")+j+"\n"
            ret += "\n"
        self.txtModules.setPlainText(ret)
        
    def quit_(self):
        self.server.shutdown()
        self.server.server_close()            
        QtGui.qApp.quit()
    
    def closeEvent(self, event):
        if event.type() == event.Close:
            event.ignore()
            self.hide()
        
QtApp(SimplePython, hidden=True, flags=QtCore.Qt.WindowStaysOnTopHint, stdout="txtConsole",
      tray={"icon": QtGui.QStyle.SP_ArrowRight, "tip": "SimplePython Server"})
