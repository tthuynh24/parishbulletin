"""
******************************************************************************
 NAME: MyThread.py
******************************************************************************
"""
from Utils.QtUtils import *

from Utils.PyUtils import getTrace, pause

##############################################################################
class MyThread(QThread):
    threadStarted = pyqtSignal(int, str)
    threadComplt = pyqtSignal(int, str)
    threadError = pyqtSignal(int, str)
    threadTerminated = pyqtSignal(int, str)

    def __init__(self, parent=None, procId=0):
        super().__init__(parent)

        self.procId = procId
        self.setTerminationEnabled(True)

        self.exiting = False
        self.alive = True

    def __del__(self):
        self.exiting = True

    def run(self):
        try:
            self.threadStarted.emit(self.procId, 'Started')
            self.alive = True
            pause(0.5, 0.01)
            self.myProc(*self.procArgs)
            self.threadComplt.emit(self.procId, 'Completed')
            self.alive = False
        except:
            self.obtainTrace()

    def render(self, myProc, procArgs):
        self.myProc, self.procArgs = myProc, procArgs
        self.start()

    def obtainTrace(self):
        self.alive = False
        self.threadError.emit(self.procId, getTrace(retEMsg=True))

    def stopProc(self):
        self.threadTerminated.emit(self.procId, 'Terminated')
        pass

    def done(self):
        try:
            if self.isRunning():
                self.terminate()
        except:
            pass
        self.deleteLater()

##############################################################################
# End of File
##############################################################################
