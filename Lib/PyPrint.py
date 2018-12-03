#!/usr/bin/env python
##############################################################################
# Module: PyPrint                                                            #
##############################################################################
from __future__ import absolute_import
import sys
import os

from PyQt5.QtCore import *
from PyQt5.QtGui import *

##############################################################################
from QtUtils import *
from .PyUtils import *

##############################################################################
##############################################################################
class PyPrint(QWidget):
    
    ##########################################################################
    def __init__(self):
        super(self.__class__, self).__init__()
        self.mylabel = Label(self)

        global PreviewWidth, PreviewHeight, HD
        margins = 0.5
        self.myPrinter = QPrinter(QPrinter.PrinterResolution)
        self.myPrinter.setPaperSize(QPrinter.Letter)
        self.myPrinter.setPageMargins(margins, margins, margins, margins, QPrinter.Inch)

    ##########################################################################
    def printReq(self, inSrc):
        if isValid(inSrc, QPixmap):
            res = self.myPrinter.resolution()
            prtImage = inSrc.scaled(7.5*res, 10*res, Qt.KeepAspectRatio)
            self.mylabel.setFixedSize(prtImage.size())
            self.mylabel.setPixmap(prtImage)
            self.printPreview()

    ##########################################################################
    def printPreview(self):
        myPreview = QPrintPreviewDialog(self.myPrinter)
        myPreview.setWindowFlags(Qt.Window)
        self.connect(myPreview, SIGNAL("paintRequested(QPrinter*)"), self._setPreview)

        return myPreview.exec_()==QDialog.Accepted

    ##########################################################################
    def _setPreview(self, printer):
        painter = QPainter()
        painter.begin(printer)
        self.mylabel.render(painter)
        painter.end()

    ##########################################################################
    def _PreviewDoc(self, printer):
        if self.hasSelectedText():
            printDoc = QTextDocument(self.selectedText())
        else:
            printDoc = QTextDocument(self.text())
        printDoc.setDefaultFont(self.font())
        printDoc.print(printer)
        
##############################################################################
# End of module                                                              #
##############################################################################
