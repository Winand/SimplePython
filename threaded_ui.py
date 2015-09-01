__author__ = 'МакаровАС'

from PyQt4 import QtCore, QtGui, uic
import sys, queue, pythoncom, types, pathlib, win32con, win32gui

DEBUG = False
print_def = lambda *args: not DEBUG or print(*args, file=sys.__stdout__)

#QtUtils
#https://bitbucket.org/philipstarkey/qtutils

class Caller(QtCore.QObject):
    """An event handler which calls the function held within a CallEvent."""
    def event(self, event):
        event.accept()
        exception = None
        try:
            result = event.fn(*event.args, **event.kwargs)
        except Exception:
            # Store for re-raising the exception in the calling thread:
            exception = sys.exc_info()
            result = None
            if event._exceptions_in_main:
                # Or, if nobody is listening for this exception,
                # better raise it here so it doesn't pass
                # silently:
                raise
        finally:
            event._returnval.put([result,exception])
        return True
caller = Caller()

def inmain(fn, *args, **kwargs):
    class CallEvent(QtCore.QEvent):
        """An event containing a request for a function call."""
        EVENT_TYPE = QtCore.QEvent.Type(QtCore.QEvent.registerEventType())
        def __init__(self, queue, exceptions_in_main, fn, *args, **kwargs):
            QtCore.QEvent.__init__(self, self.EVENT_TYPE)
            self.fn = fn
            self.args = args
            self.kwargs = kwargs
            self._returnval = queue
            # Whether to raise exceptions in the main thread or store them
            # for raising in the calling thread:
            self._exceptions_in_main = exceptions_in_main

    def in_main_later(fn, exceptions_in_main, *args, **kwargs):
        """Asks the mainloop to call a function when it has time. Immediately
        returns the queue that was sent to the mainloop.  A call to queue.get()
        will return a list of [result,exception] where exception=[type,value,traceback]
        of the exception.  Functions are guaranteed to be called in the order
        they were requested."""
        q = queue.Queue()
        QtCore.QCoreApplication.postEvent(caller, CallEvent(q, exceptions_in_main, fn, *args, **kwargs))
        return q

    def get_inmain_result(queue):
        result,exception = queue.get()
        if exception is not None:
            type, value, traceback = exception
            raise value.with_traceback(traceback)
        return result
    return fn(*args, **kwargs) if isMainThread() else get_inmain_result(in_main_later(fn,False,*args,**kwargs))

def bind(func, to):
    "Bind function to instance, unbind if needed"
    return types.MethodType(func.__func__ if hasattr(func, "__self__") else func, to)

class prx():
    "Proxies object, automatically calls methods in GUI thread"
    GETATTR, CALL = range(2)
    builtin = str, bool, int, type(None), complex, bytes, dict
    def __init__(self, client, *args, atts={}, **kwargs):
        self.__dict__['client'] = client
        for k in atts:
            self.__dict__[k] = atts[k]
    def proxy(self, t, *args, **kwargs):
        if t == self.GETATTR:
            print_def("THD_UI GET:", self.client, self.client.__class__)
            ret = getattr(self.client, args[0])
        else:
            if hasattr(self.client, "__self__"):
                if self.client.__self__.__module__.endswith("QtGui"):
                    #Call QtGui stuff in main thread
                    print_def("THD_UI CALL IN MAIN:", self.client.__name__)
                    ret = inmain(self.client, *args, **kwargs)
                else: #Call other stuff in the same thread, pass proxied /self/
                    print_def("THD_UI CALL:", self.client.__name__)
                    ret = bind(self.client, prx(self.client.__self__))(*args, **kwargs)
            else: #Call unbound stuff
                print_def("THD_UI CALL UNBOUND:", self.client.__name__)
                ret = self.client(*args, **kwargs)
        return ret if type(ret) in self.builtin else prx(ret) #if type(ret) != types.MethodType else ret
    def __getattr__(self, name): return self.proxy(self.GETATTR, name)
    def __call__(self, *args, **kwargs): return self.proxy(self.CALL, *args, **kwargs)
    def __setattr__(self, name, value): return setattr(self.client, name, value)
    def __str__(self): return "<Proxied %s>" % self.client
    def __eq__(self, other): return self.client is other.client

threads = [] #stores thread refs
class GenericThread(QtCore.QThread):
    def __init__(self, client_self, function, *args, **kwargs):
        QtCore.QThread.__init__(self)
        self.client_self = prx(client_self, \
                            atts={"sender": lambda s=client_self.sender(): s})
        self.function, self.args, self.kwargs = function, args, kwargs
        self.finished.connect(self.finished_)
        threads.append(self)
        self.start()
    def __del__(self):
        if self.isRunning():
            print_def("Thread %s is still running. Waiting..." % self)
            self.wait() #block until finished
    def finished_(self):
        del threads[threads.index(self)]
    def run(self):
        pythoncom.CoInitialize()
        self.function(self.client_self, *self.args, **self.kwargs)

def isMainThread():
    if not QtCore.QCoreApplication.instance():
        print_def("THD_UI ERROR (isMainThread): app instance is None!")
        return True
    return QtCore.QThread.currentThread() is QtCore.QCoreApplication.instance().thread()

def pyqtThreadedSlot(*args, **kwargs):
    def threaded_int(func):
        @QtCore.pyqtSlot(*args, name=func.__name__, **kwargs)
        def wrap_func(self, *args1, **kwargs1):
            GenericThread(self, func, *args1, **kwargs1)
        return wrap_func
    return threaded_int
    
def module_path(cls):
    "Get module folder path from class"
    return pathlib.Path(sys.modules[cls.__module__].__file__).absolute().parent
    
class QApplication(QtGui.QApplication):
    terminated = QtCore.pyqtSignal()
    def winEventFilter(self, message):
        if message.message == win32con.WM_DESTROY:
            if win32gui.GetClassName(int(message.hwnd)
                        ).startswith("QEventDispatcherWin32_Internal_Widget"):
                print("Application terminated.")
                self.terminated.emit()
        return QtGui.QApplication.winEventFilter(self, message)

#Widget events are connected to appropriate defs - <widget>_<signal>()
#To catch terminated signal (QProcess.terminate) connect it manually
def QtApp(Form, *args, flags=QtCore.Qt.WindowType(), ui=None, stdout=None, tray=None, hidden=False, ontop=False, **kwargs):
    "Create new QApplication and specified window"
    app = QApplication(sys.argv)
    class Form_(Form):
        def __init__(self, flags, ui):
            super(Form, self).__init__(flags=flags|(QtCore.Qt.WindowStaysOnTopHint if ontop else 0))
            uic.loadUi(str(ui or module_path(Form).joinpath(Form.__name__.lower()))+".ui", self)
            if stdout: redirect_stdout(getattr(self, stdout))
            if type(tray) is dict:
                f = app.style().standardIcon \
                    if type(tray["icon"]) is QtGui.QStyle.StandardPixmap else QtGui.QIcon
                self.tray = QtGui.QSystemTrayIcon(f(tray["icon"]))
                if "tip" in tray:
                    self.tray.setToolTip(tray["tip"])
                self.tray.setContextMenu(QtGui.QMenu())
                self.tray.show()
                self.tray.addMenuItem = bind(addMenuItem, self.tray)
                QtGui.qApp.setQuitOnLastWindowClosed(False) #important! open qdialog, hide main window, close qdialog: trayicon stops working
            widgets, members = super(Form, self).__dict__, Form.__dict__
            for i in widgets:
                for m in [j for j in members if j.startswith(i+"_")]:
                    signal = getattr(widgets[i], m[len(i)+1:], None)
                    if signal: signal.connect(bind(members[m], self))
                    else: print("Signal '%s' of '%s' not found" % (m[len(i)+1:], i))
            self.terminated = app.terminated
            if "__init__" in members:
                super().__init__(*args, **kwargs)
    form = Form_(flags, ui)
    if not hidden:
        form.show()
    def aboutToQuit(): #cleanup before exit
        del form.tray.addMenuItem
    app.aboutToQuit.connect(aboutToQuit)
    sys.exit(app.exec_())
    
def addMenuItem(self, *args):
    for i in range(0, len(args), 2):
        self.contextMenu().addAction(args[i]).triggered.connect(args[i+1])
    
def Widget(Form, *args, flags=QtCore.Qt.WindowType(), ui=None, exec_=False, ontop=False, **kwargs):
    class Form_(Form):
        def __init__(self, flags, ui):
            super(Form, self).__init__(flags=flags|(QtCore.Qt.WindowStaysOnTopHint if ontop else 0))
            uic.loadUi(str(ui or module_path(Form).joinpath(Form.__name__.lower()))+".ui", self)
            if "__init__" in Form.__dict__:
                super().__init__(*args, **kwargs)
    form = inmain(Form_, flags, ui)
#    form.setWindowModality(1)
    if exec_:
        return inmain(form.exec), form.getResult()
    else: return form

def redirect_stdout(wgt):
    """Redirect standard output to the specified widget"""
    classes = wgt.metaObject().className(), wgt.metaObject().superClass().className()
    if "QPlainTextEdit" in classes:
        def write(self, txt):
            self.moveCursor(QtGui.QTextCursor.End)
            self.insertPlainText(txt)
        wgt.write = bind(write, wgt)
        wgt.flush = bind(lambda self: None, wgt)
    else:
        print_def("THD_UI ERROR (redirect_stdout): cannot redirect output to unsupported "+classes[0])
        return
    parent = wgt.parent()
    def closeEvent(e, orig_ce=parent.closeEvent):
        sys.stdout = sys.__stdout__
        orig_ce(e)
    parent.closeEvent = closeEvent
    sys.stdout = prx(wgt)