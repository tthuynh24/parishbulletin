#!/usr/bin/env python
##############################################################################
# Module: AppShared                                                          #
##############################################################################
from __future__ import absolute_import
import os
import time
import platform
import logging
from logging import WARNING, ERROR, CRITICAL, INFO
from queue import Queue

from PyQt5.QtWidgets import QWidget
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import QThread, pyqtSignal, Qt

from . import icons_rc

_fromUtf8 = lambda s: s

LOG_STAT, LOG_MSG, PROG_BAR = 60, 70, 80
PROG_New, PROG_Stat, PROG_End = range(3)

##############################################################################
class QtQueue(QWidget):
    QueAdded = pyqtSignal()

    ##########################################################################
    def __init__ (self):
        super().__init__(None)
        self.qtQue = Queue()

    ##########################################################################
    def put(self, addItem, block=False):
        self.qtQue.put(addItem, block)
        self.QueAdded.emit()

    ##########################################################################
    def get(self, block=False):
        return self.qtQue.get(block)

    ##########################################################################
    def isEmpty(self):
        return self.qtQue.empty()

    ##########################################################################
    def qsize(self):
        return self.qtQue.qsize()

    ##########################################################################
    def full(self):
        return self.qtQue.full()


###########################################################################
def createIcon(iconRc, pixmap=False):
    selMethod = QPixmap if pixmap else QIcon
    return selMethod(_fromUtf8(':/%s' % iconRc))

##############################################################################
# ShareData
##############################################################################
##########################################################################
def setSharedObj(applNm, appl):
    setattr(SharedData, applNm, appl)

##########################################################################
def getSharedObj(getAttr):
    return getattr(SharedData, getAttr, None)

##########################################################################
def runSharedApp(selApp):
    app = getattr(SharedData, selApp, None)
    if app is not None:
        return app()

##########################################################################
def isReady():
    return SharedData.isReady

##########################################################################
def setReady(readyStat):
    SharedData.isReady = readyStat

##########################################################################
def getAppl():
    return SharedData.appl

##############################################################################
def setPath(pathNm, newPath):
    SharedData.appl.setPath(pathNm, newPath)
    return newPath

##########################################################################
def isAbsPath(selPath):
    return os.path.isabs(selPath)
def getApplObj(getAttr):
    return getattr(SharedData.appl, getAttr, None)


##############################################################################
##############################################################################
class MyLog(logging.Logger):

    def __init__(self, name):
        """
            This is the constructor for class Log.
        py 3
        super(self).__init__(name)
        py 2
        super(class).__init__(name)
        """
        super().__init__(name)

        try:
            logFn, self.appDir = prepLogFile()
            fh = logging.FileHandler(logFn)
        except:
            fh = logging.StreamHandler()
            self.appDir = os.getcwd()

        self.logDir = os.path.join(self.appDir, 'Logs')

        fh.setLevel(logging.DEBUG)
        fs = '%(asctime)s.%(msecs)-3d %(name)s: %(levelname)s %(message)s'
        fd = '%Y-%m-%d.%H:%M:%S'
        fmt = logging.Formatter(fs, fd)
        fh.setFormatter(fmt)
        self.addHandler(fh)

        self.resetStat()
        self.logger = {WARNING:  self.warning,
                       ERROR:    self.error,
                       CRITICAL: self.critical,
                       INFO:     self.info,
                      }

    def logMsg(self, msg, level):
        self.logger.get(level, self.info)(msg)
        return prtMsg(msg, level)

    def _log(self, level, msg, *args, **kwargs):
        self.lastReport = (level, msg, 1)

    def verStat(self, level):
        if level in (ERROR, CRITICAL): 
            self.isOk = False

    def resetStat(self, statOk=True):
        self.isOk = statOk

##############################################################################
##########################################################################
def setAppl(appl):
    MyLogger.appl = appl

##########################################################################
def getAppl():
    return MyLogger.appl

##########################################################################
def getApplObj(getAttr):
    return getattr(MyLogger.appl, getAttr, None)

##########################################################################
def setStat(statMsg):
    addMsg((statMsg, LOG_STAT))

def wrt2Trm(msg):
    addMsg((msg, LOG_MSG))

def reportStat(msg, statLvl=ERROR):
    MyLogger.verStat(statLvl)
    addMsg((msg, statLvl))

def addMsg(msgInfo):
    try:
        MyLogger.appl.addMsg(msgInfo)
    except Exception as e:
        print(e)

##########################################################################
def logMsg(msg, statLvl=ERROR):
    return MyLogger.logMsg(msg, statLvl)

##########################################################################
def isProcOk():
    return MyLogger.isOk

##########################################################################
def resetStat(statOk=True):
    MyLogger.resetStat(statOk)

##########################################################################
##########################################################################
def procEventQue():
    #start = time.time()
    #while time.time() - start < 1:
    SharedData.appl.processEvents()
    #    time.sleep(0.1)
def setPath(pathNm, newPath):
    MyLogger.appl.setPath(pathNm, newPath)
    return newPath

def getPath(pathNm):
    return MyLogger.appl.getPath(pathNm)

def getLogDir():
    return MyLogger.logDir

##########################################################################
def rtrvLogFn():
    for h in MyLogger.__dict__['handlers']:
        if h.__class__.__name__ == 'FileHandler':
            return h.baseFilename
    return 'Sorry - log file has not been create for your session.'

##########################################################################
def prtMsg(msg, statLvl):
    if statLvl!=CRITICAL:
        if statLvl==WARNING:
            statMsg = 'WARNING: %s' % msg
        elif statLvl==ERROR:
            statMsg = 'ERROR: %s' % msg
        else:
            statMsg = msg if statLvl is None else 'INFO: %s' % msg

        if getattr(MyLogger, 'appl', None) is not None:
            wrt2Trm(statMsg)

        return statMsg

##########################################################################
##########################################################################
def getUserHome():
    return os.path.expanduser('~')

##########################################################################
def verifyDir(verDir, createIfNotExist=True):
    if os.path.isdir(verDir):
        return True
    if createIfNotExist:
        try:
            os.mkdir(verDir)
        except:
            pass
    return os.path.isdir(verDir)

####################################################################################
def setTmpDir():
    tmpDir = os.path.join(MyLogger.appDir, 'tmp')
    verifyDir(tmpDir)
    return tmpDir

##############################################################################

MyLogger = MyLog('Parish')

##############################################################################
##############################################################################
class MsgEvents(QThread):
    NewMsg = pyqtSignal()

    def __init__(self, parent):
        super().__init__(parent)
        self.msgQue = QtQueue()
        self.msgQue.QueAdded.connect(self.onNewMsg)

    def run(self):
        self.msgHdl()

    def msgHdl(self):
        self.msgDia = QDialog(self)
        self.msgDia.hide()

    def onNewMsg(self):
        self.NewMsg.emit()

    def postMsg(self, msgInfo):
        self.NewMsg.emit()

    def __del__(self):
        self.deleteLater()
        
##############################################################################
SharedData = lambda cls: cls.__name__
setattr(SharedData, 'isReady', False)
MyLogger = MyLog('Parish')


##############################################################################
# End of module                                                              #
##############################################################################
