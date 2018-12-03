#!/usr/bin/env python
##############################################################################
# Module: MyAppl                                                            #
##############################################################################
from __future__ import absolute_import
import re
import fnmatch

from PyQt5.QtWidgets import QApplication
from PyQt5 import QtWebEngineWidgets
from PyQt5.QtCore import QUrl

from Lib.AppShared import *
from Lib.Window import Window
from Lib.Dialog import Dialog
from Lib.PrtDialog import PrtDialog
from Lib.StatusTrm import startEditor
from Lib.MyThread import MyThread
from Lib.MsWord import MsWord, W, H
from Lib.MyConfig import MyConfig
from Lib.TabManager import TabManager

from Utils.QtUtils import *
from Utils.PyUtils import *
from Utils.BookletUtils import *

SEP = None
ColourCode = {'ERROR':'red',
              'WARNING':'orange',
              'NOTE':'green',
              'DFLT':'blue'
              }
AdType = 0

####### Importing user defined packages ###############
AdsSpaces = newObj()
setattr(AdsSpaces, 'SGL',  (4, 1.50))
setattr(AdsSpaces, 'HAP',  (8, 4.56))
setattr(AdsSpaces, 'DBLA', (4, 2.50))
setattr(AdsSpaces, 'DBLB', (8, 1.25))
setattr(AdsSpaces, 'TRPA', (4, 3.00))
setattr(AdsSpaces, 'TRPB', (8, 2.00))
setattr(AdsSpaces, 'QTRA', (4, 4.00))
setattr(AdsSpaces, 'QTRB', (8, 2.50))
setattr(AdsSpaces, 'SPC',  (8, 1.25))
AdType = 0

STATUS = PDFVIEWER = 0
PDFJS = 'file:///./web/viewer.html'
PDF = 'file:///{}'

##############################################################################
class BulletinMFG(Window):

    def __init__(self):
        global STATUS
        super().__init__()
        self.setWindowTitle('Editor')
        setAppl(self)

        self.path = self.currDir = getDirname(absPath(__file__))
        
        self.menuBar = self.getMenuBar()
        self.mainLayout = self.getLayout()
        setSpacing(self.mainLayout, 4)
        self.addMenubar()

        AddWidget(self.mainLayout, CustLabel('Parish Bulletin Manufature'))
        self.statTrm = startEditor()
        self.addWidgets()
        
        self.tabInd = -1
        self.tabMgr = TabManager(self)
        AddWidget(self.mainLayout, self.tabMgr)
        
        self.statTrm = startEditor()
        STATUS = self.tabMgr.addNewTab(self.statTrm, 'Status')
        
        self.addPdfView()

        self.msgHdl = MsgEvents(self)
        self.msgHdl.NewMsg.connect(self.onNewMsg)
        
        self.tabMgr.setActTab(STATUS)

        self.resize(1100, 750)
        
        self.pageId = 0
        self.docNm = ''

    ##########################################################################
    def __del__(self):
        self.onCloseAppl()

    ##########################################################################
    def _saveDoc(self):
        try:
            ttmvFn = joinPaths(self.path, 'TTMV', self.docNm)
            self.msWord.saveDoc(ttmvFn.replace('/', '\\'))
        except Exception as eMsg:
            addMsg('====>> Unable to save the your document: {0}.'.format(eMsg))

    ##########################################################################
    def addMenubar(self):
        
        self.menuBar.addMenu(createIcon('Editor'), 'Logo')

        self.menuAct = {}
        
        # File menu
        menuLst = ('Print', SEP, 'Exit')
        actLst = (self.onPrtReq, '', self.onCloseAppl)
        dbMenu = AddSubMenu(self.menuBar, '&File', menuLst, actLst, self.menuAct)
        
        self.menuBar.triggered.connect(self.onMenuSel)
        self.menuBar.addSeparator()

    ##########################################################################
    def addWidgets(self):
        labelWidth = 160
        myLayout = BoxLayout('V', spacing=2, margins=(6, 0, 6, 0))

        self.cfgSel = Button(None, icon=createIcon('OpenFile'))
        txtGrp, self.selCfg = TextWidget('Ads Config: ', btnObj=self.cfgSel, textSize=12, labelWidth=labelWidth)
        AddWidget(myLayout, txtGrp)

        self.currDir = joinPaths(self.path, 'Advertisement')
        self.selCfg.setText(joinPaths(self.currDir, 'Ads.cfg'))

        self.selDate = Button(None, icon=createIcon('Calendar'))
        txtGrp, self.relDate = TextWidget('Release Date: ', btnObj=self.selDate, textSize=12, labelWidth=labelWidth)
        AddWidget(myLayout, txtGrp)

        self.selTTMV = Button(None, icon=createIcon('OpenFile'))
        txtGrp, self.fileNm = TextWidget('Bulletin File Name: ', btnObj=self.selTTMV, textSize=12, labelWidth=labelWidth)
        AddWidget(myLayout, txtGrp)

        AddWidget(myLayout, self.addPrtSelection(textSize=12))


        btnsLayout = BoxLayout('H', spacing=10, margins=(6, 0, 0, 0))
        buttons = (('Bulletin', 'Create Bulletin', self.onCreateBull),
                   ('Advertise', 'Add Advertisement', self.onConfigAds),
                   ('Save', 'Save TTMV', self.onSave),
                   ('Printer', 'Print TTMV', self.onPrintTTMV),
                   ('Print', 'Prints Spooler', self.onPrintSpooler),
            )
        self.actBtns = []
        for image, tip, hdl in buttons:
            btn = Button(None, label=tip, icon=createIcon(image),
                         enable=False, tooltip='Click on button to add new page')
            btn.clicked.connect(hdl)
            AddWidget(btnsLayout, btn)
            self.actBtns.append(btn)
        AddLayout(myLayout, btnsLayout)

        AddWidget(self.mainLayout, GroupBox(None, myLayout))
        
        self.cfgSel.clicked.connect(self.onSelCfg)
        self.selTTMV.clicked.connect(self.onSelTTMV)
        self.selDate.clicked.connect(self.onselDate)
        self.enableSel()
        
    ##########################################################################
    def addPdfView(self):
        global PDFVIEWER
        self.pdfViewer = QtWebEngineWidgets.QWebEngineView()
        self.pdfViewer.settings().setAttribute(QtWebEngineWidgets.QWebEngineSettings.PluginsEnabled, True)
        PDFVIEWER = self.tabMgr.addNewTab(self.pdfViewer, 'PdfViewer')
        
    ##########################################################################
    def addPrtSelection(self, textSize=16):
        layout = BoxLayout('H')

        prtTTMVSel = CheckBox()
        AddWidget(layout, prtTTMVSel, stretch=0)
        prtTTMVSel.clicked.connect(self.onPrintting)

        label = CustLabel('Printer:', textSize, fixedWidth=60, align='right')
        AddWidget(layout, label, stretch=0)
        self.myPrinter = ComboBox(editable=True, minWidth=175)
        AddWidget(layout, self.myPrinter, stretch=3)

        label = CustLabel('Paper:', textSize, fixedWidth=60, align='right')
        AddWidget(layout, label, stretch=0)
        self.prtPaper = ComboBox(editable=True, minWidth=175)
        AddWidget(layout, self.prtPaper, stretch=2)

        label = CustLabel('Copies:', textSize, fixedWidth=60, align='right')
        AddWidget(layout, label, stretch=0)
        self.prtCopies = ComboBox(editable=True, fixedWidth=50)
        AddWidget(layout, self.prtCopies, stretch=0)
        self.loadPrtSelection()
        return  GroupBox('', layout, fixedHeight=55*textSize/16)

    ##########################################################################
    def loadPrtSelection(self):
        printerLst = rtrvPrinters()
        self.mySelPrinter = self.selMyPrinter(printerLst)

        loadSelLst(self.myPrinter, printerLst, selVal=self.mySelPrinter)
        loadSelLst(self.prtPaper, ['Letter', 'Tabloid'], selVal='Tabloid')
        loadSelLst(self.prtCopies, ['1', '5', '10', '25', '50', '75', '100', '125', '150' ], sortLst=False)
        
    ##########################################################################
    def selMyPrinter(self, printerLst, dfltPrt='Xerox WorkCentre 5325 V4 PCL6'):
        if dfltPrt in printerLst:
            return dfltPrt
        for prt in printerLst:
            if prt in ('Send To OneNote 2013', 'Microsoft XPS Document Writer', 'Microsoft Print to PDF', 'Fax'): continue
            return prt
        
    ##########################################################################
    ##########################################################################
    def enableSel(self, selStat=None):
        enable = isNull(selStat) and len(self.fileNm.text())>0 and len(self.selCfg.text())>0
        for btn in self.actBtns[:-2]:
            setEnabled(btn, enable)
            
        enable = selStat in (None, False)
        for wdgt in (self.cfgSel, self.selCfg, self.selDate, self.relDate, self.fileNm):
            setEnabled(wdgt, enable)

        enable = not enable
        for wdgt in self.actBtns[-2:] + [self.selTTMV, self.myPrinter, self.prtPaper, self.prtCopies]:
            setEnabled(wdgt, enable)
    
    ##########################################################################
    def onEditorPage(self):
        fileNm = joinPaths(self.path, 'Advertisement', 'Ads', 'InsertTest.docx')
        self.msWord.addPage()
        self.msWord.addTextBox(self.msWord.numPages(), fileNm)
        self.rtrvNumPages()

    ##########################################################################
    def onImagPage(self):
        self.msWord.addPage()
        self.rtrvNumPages()

    ##########################################################################
    def onPrintting(self, isSel=False):
        if isSel:
            ttmvpath = joinPaths(self.path, 'TTMV')
            newestTTMV = getNewestFile(ttmvpath, "TTMV_*-*.pdf")
            self.fileNm.setText(newestTTMV)
            self.loadPdf(newestTTMV)
        else:
            self.fileNm.setText(self.docNm)
            self.tabMgr.setActTab(STATUS)
        self.enableSel(isSel)

    ##########################################################################
    def onPrintTTMV(self, isSel=False):
        selPrt = self.myPrinter.currentText()
        selTTMV = self.fileNm.text()
        selPaper = self.prtPaper.currentText()
        try:
            numCopies = int(self.prtCopies.currentText())
        except:
            numCopies = 1
        
        pdfBooklet = printTTMV(self.path, selPrt, selTTMV, selPaper, numCopies)
        self.loadPdf(pdfBooklet)
        
    ##########################################################################
    def setMyPrinter(self, selPrt):
        self.mySelPrinter = selPrt

    ##########################################################################
    def onPrintSpooler(self):
        try:
            showPrtQue(self.mySelPrinter)
        except Exception:
            pass


        
    ##########################################################################
    def addAdsContent__(self, myCfg):
        
        self.myCfg = MyConfig(selCfg)
        for sectn in self.myCfg.sections():
            self.msWord.addSection()
            adCfg, custAds = [], []
            for rowInd, opt in enumerate(self.myCfg.options(sectn)):
                rowInfo = self.myCfg.getOptVal(sectn, opt)
                if 'CST' in rowInfo:
                    custAds.append((tbl, rowInfo[1], joinPaths(self.path, 'Advertisement', 'Ads', rowInfo[-1])))
                    continue

                adInfo = self.getAdInfo(rowInfo)
                if any(adInfo):
                    adCfg.append((rowInd+1, self.compileAdsLst(rowInfo)))
                    if rowInd==0:
                        self.pageId += 1
                        if self.pageId>self.msWord.numPages():
                            self.msWord.addPage()
                        self.msWord.selectPage()
                        self.msWord.fmtAdsPages()
                        self.msWord.prepAdsPage(adInfo[H])
                        tbl = self.msWord.numTables()
                    else:
                        self.msWord.addRow(tbl, adInfo[H])

            self.msWord.addRow(tbl, 0.01)
            for rowId, adsLst in adCfg:
                self.msWord.loadAdsToRow(tbl, rowId, adsLst[1:], adsLst[0])
            self.msWord.delRow(tbl, rowId+1)
        
        
        
        
        
        
        
    ##########################################################################
    def loadPdf(self, pdfFile):
        self.pdfViewer.load(QUrl.fromUserInput('%s?file=%s' % (PDFJS, PDF.format(pdfFile))))
        self.tabMgr.setActTab(PDFVIEWER)

    ##########################################################################
    def onConfigAds(self):
        cfgFn = self.selCfg.text().strip()
        if len(cfgFn)>0:
            self.configAds(cfgFn)
            self.rtrvNumPages()
        self.tabMgr.setActTab(STATUS)

    ##########################################################################
    def onSave(self):
        self._saveDoc()

    ##########################################################################
    def onSelCfg(self):
        selCfg = SelectFile(self, self.currDir, '*.cfg',
                            title='Select Ads Config File')
        self.selCfg.setText(selCfg[0])
        self.enableSel()

    ##########################################################################
    def onSelTTMV(self):
        selCfg = SelectFile(self, self.currDir, '*.pdf',
                            title='Select TTMV to print')
        self.fileNm.setText(selCfg[0])

    ##########################################################################
    def onselDate(self):
        dia = DatePicker(self.setDate)
        dia.exec_()

    ##########################################################################
    def setDate(self, setDate):
        timeStamp = setDate.toString("yyyy-MM-dd")
        self.docNm = 'TTMV_{0}.docx'.format(timeStamp)
        self.relDate.setText(timeStamp)
        self.fileNm.setText(self.docNm)
        self.enableSel()

    ##########################################################################
    def startMsWord(self):
        if len(self.fileNm.text())>0 and len(self.selCfg.text())>0:
            self.enableSel()

            self.msWord.addSection()
            self.rtrvNumPages()
            self.msWord.selectPage(2)
            self.msWord.fmtBodyPages()

    ##########################################################################
    def addMsg(self, msgInfo):
        self.msgHdl.msgQue.put(msgInfo, True)

    ##########################################################################
    def onNewMsg(self):
        msgInfo = self.msgHdl.msgQue.get(True)

        msg, msgType = msgInfo if isValid(msgInfo, (tuple, list)) else (msgInfo, None)
        if isNull(msg): return
        if msgType==PROG_BAR:
            return self.newProgStat(msg)

        if not isinstance(msg, str):
            msg = 'INVALID Message (CODE): %s - %s' % (type(msgInfo), str(msgInfo))

        if msgType==LOG_STAT:
            return self.setStatus(msg)

        self.statTrm.write(msg.strip()+'\n')

    ##########################################################################
    ##########################################################################
    def onMenuSel(self, evt=None):
        selMenu, parentNm, selAct = self.getSelItem(evt)
        self.processEvents()
        eAction, eCaller = self.menuAct.get(selMenu, (None, None))
        if eCaller is not None:
            eCaller()

    ##########################################################################
    def getSelItem(self, evt):
        selMenu = str(evt.text()).replace('&', '')
        parentNm = str(evt.parentWidget().title())
        return selMenu, parentNm, parentNm + selMenu

    ##########################################################################
    def processEvents(self):
        self.setFocus(True)
        QCoreApplication.processEvents()

    ##########################################################################
    ##########################################################################
    def onCloseAppl(self, evt=None):
        #self._saveDoc()
        #self.msWord.quitApp()
        #self.msWord.closeDoc(saveDoc=False)

        print('\n====>> onCloseAppl')

        #self.__del__()
        #sys.exit(0)

    ##########################################################################
    ##########################################################################
    def onPrtReq(self, evt=None):
        dlg = PrtDialog(self.printReq)

    ##########################################################################
    ##########################################################################
    def startInThread(self, procId, proc, args):

        threadProc = MyThread(self, procId)
        threadProc.render(proc, args)
        
        threadProc.threadStarted.connect(self.threadStarted)
        threadProc.threadComplt.connect(self.threadStoped)
        threadProc.threadError.connect(self.threadStoped)

        return threadProc

    ##########################################################################
    def threadStarted(self, procId, msg):
        self.setStatus('Please wait, your request(s) is being processed...')
        self.progress = 1

    ##########################################################################
    def threadStoped(self, procId, msg):
        self.setStatus('Process completed...')

    ##########################################################################
    def setStatus(self, statMsg):
        self.statusBar().showMessage(statMsg)
        statColor = 'color: %s; background-color: #d3ecff' % self.setStatusColor(statMsg)
        self.statusBar().setStyleSheet(statColor)
        
    ##########################################################################
    def setStatusColor(self, statMsg):
        msgKey = statMsg.split(':',1)[0]
        return ColourCode[msgKey if msgKey in ColourCode else 'DFLT']

    ##########################################################################
    ##########################################################################
    def onCreateBull(self):
        self.msWord = MsWord()
        
        self.msWord.new()
        selID, adsPageId = 0, 3
        for sectn in range(4):
            selID = self.msWord.numPages()
            sectn = self.msWord.fmtSection(selID)
            
            if selID==1:
                self.msWord.fmtFrontPage(sectn)
                self.addFrontPageContent(selID)
            elif selID==2:
                self.msWord.fmt2ndPage(sectn)
                self.add2ndPageContent(selID)
            elif selID>2:
                adsTbl = self.msWord.fmtAdsPage(sectn, selID)
                if selID==3:
                    cfgFn = self.selCfg.text().strip()
                    myCfg = MyConfig(cfgFn)
                self.addAdsContent(selID, adsTbl, myCfg)

            if selID<4:
                self.msWord.addSection()
                
        self.verPages(selID)
        
    ##########################################################################
    def addFrontPageContent(self, pageId):
        self.msWord.addImageTheme(joinPaths(self.path, 'Data', 'NhaTho_Logo.png'))

        self.msWord.importFileToCell(pageId, 2, 1,
                                     joinPaths(self.path, 'Data', 'TTMV_LeftPanel.htm'))

        self.msWord.importFileToCell(pageId, 1, 2,
                                     joinPaths(self.path, 'Data', 'TTMV_TopPanel.htm'))
        
        self.msWord.importFileToCell(pageId, 2, 2,
                             joinPaths(self.path, 'Data', 'ContentFrontPage.docx'))

        
        print('====>> Front page DONE', self.msWord.numPages())
        self.rtrvNumPages()

    ##########################################################################
    def add2ndPageContent(self, pageId):
        
        self.msWord.importFileToCell(pageId, 1, 1,
                                     joinPaths(self.path, 'Data', 'Content2ndPage.docx'))
        
        self.msWord.importFileToCell(pageId, 2, 1,
                                     joinPaths(self.path, 'Data', 'TTMV_BotPanel.htm'))
        
    ##########################################################################
    def addAdsContent(self, pageId, adsTbl, myCfg):
        adsSectn = 'AdsPage{}'.format(pageId-2)
            
        adCfg, custAds = [], []
        for rowInd, opt in enumerate(myCfg.options(adsSectn)):
            rowInfo = myCfg.getOptVal(adsSectn, opt)
            if 'CST' in rowInfo:
                custAds.append((rowInfo[1], joinPaths(self.path, 'Advertisement', 'Ads', rowInfo[-1])))
                continue

            adInfo = self.getAdInfo(rowInfo)
            if any(adInfo):
                adCfg.append((rowInd+1, self.compileAdsLst(rowInfo)))
                if rowInd==0:
                    self.msWord.setRowHeight(adsTbl, 1, adInfo[H])
                else:
                    self.msWord.addRow(adsTbl, adInfo[H])

        for rowId, adsLst in adCfg:
            self.msWord.loadAds(adsTbl, rowId, adsLst[1:], adsLst[0])
            
        for adInfo, adFn in custAds:
            self.msWord.loadCustAds(adsTbl, adInfo, adFn)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ##########################################################################
    def onCreateBull__(self):
        self.msWord = MsWord()
        self.addFrontPage()
        self.addSecPage()
        self.rtrvNumPages()
        self.tabMgr.setActTab(STATUS)

    ##########################################################################
    def addFrontPage(self):
        self.pageId = 1
        self.msWord.new()

        self.msWord.fmtFrontPage()
        self.msWord.addFrontPage()

        imagFn = joinPaths(self.path, 'Data', 'NhaTho_Logo.png')
        self.msWord.addImageTheme(imagFn)

        sideBarFn = joinPaths(self.path, 'Data', 'TTMV_LeftPanel.htm')
        self.msWord.importFileToCell(self.pageId, 2, 1, sideBarFn)

        parishLogo = joinPaths(self.path, 'Data', 'TTMV_TopPanel.htm')
        self.msWord.importFileToCell(self.pageId, 1, 2, parishLogo)
        self.rtrvNumPages()

        print('====>> addFrontPage - Front page DONE', self.msWord.numSections(), self.msWord.numPages())

    ##########################################################################
    def addSecPage(self):
        self.pageId += 1
        self.msWord.addSection()
        self.msWord.selectPage(pageId=self.pageId)
        self.msWord.fmtBodyPages()
        
        Table = self.msWord.addTblPerPage(pageId=self.pageId, numRows=2, numCols=1)

        parishLogo = joinPaths(self.path, 'Data', 'TTMV_BotPanel.htm')
        self.msWord.importFileToCell(self.pageId, 2, 1, parishLogo)

    ##########################################################################
    def rtrvNumPages(self):
        addMsg('====>> Total number of pages in document: {0} pages.'.format(self.msWord.numPages()))

    ##########################################################################
    def configAds(self, selCfg):
        self.myCfg = MyConfig(selCfg)
        for sectn in self.myCfg.sections():
            self.msWord.addSection()
            adCfg, custAds = [], []
            for rowInd, opt in enumerate(self.myCfg.options(sectn)):
                rowInfo = self.myCfg.getOptVal(sectn, opt)
                if 'CST' in rowInfo:
                    custAds.append((tbl, rowInfo[1], joinPaths(self.path, 'Advertisement', 'Ads', rowInfo[-1])))
                    continue

                adInfo = self.getAdInfo(rowInfo)
                if any(adInfo):
                    adCfg.append((rowInd+1, self.compileAdsLst(rowInfo)))
                    if rowInd==0:
                        self.pageId += 1
                        if self.pageId>self.msWord.numPages():
                            self.msWord.addPage()
                        self.msWord.selectPage()
                        self.msWord.fmtAdsPages()
                        self.msWord.prepAdsPage(adInfo[H])
                        tbl = self.msWord.numTables()
                    else:
                        self.msWord.addRow(tbl, adInfo[H])

            self.msWord.addRow(tbl, 0.01)
            for rowId, adsLst in adCfg:
                self.msWord.loadAdsToRow(tbl, rowId, adsLst[1:], adsLst[0])
            self.msWord.delRow(tbl, rowId+1)

            for tbl, adInfo, adFn in custAds:
                self.msWord.loadCustAds(tbl, adInfo, adFn)
                
        self.verPages()
        self.rtrvNumPages()
        
        
    ##########################################################################
    def verPages(self, numPages):
        if numPages>self.msWord.numPages(): return
        self.msWord.selectPage(pageId=self.msWord.numPages())
        self.msWord.setFontSize(fontSize=1)
        
    ##########################################################################
    def compileAdsLst(self, adLst):
        return [adLst[0]] + [joinPaths(self.path, 'Advertisement', 'Ads', adFn) for adFn in adLst[1:]]

    ##########################################################################
    def getAdInfo(self, rowInfo):
        return getattr(AdsSpaces, rowInfo[AdType], (4, 1.50))

    ##########################################################################
    def LoadRowInfo(self, tbl, row, adLst):
        adsLst = []
        for ind, ad in enumerate(adLst):
            adFn = joinPaths(self.path, 'Advertisement', 'Ads', ad)
            adsLst.append(adFn)
        self.msWord.loadAdsToRow(tbl, row, adsLst)

##############################################################################
if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    win = BulletinMFG()
    win.show()
    sys.exit(app.exec_())

##############################################################################
# End of module                                                              #
##############################################################################
