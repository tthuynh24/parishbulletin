#!/usr/bin/env python
##############################################################################
# Module: TitleBar                                                            #
##############################################################################
from __future__ import absolute_import

from PyQt5.QtWidgets import QMenuBar, QSizePolicy
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette
from .UI_Vars import *

##############################################################################
class TitleBar(QMenuBar):

    def __init__(self, win):
        super().__init__()
        css = """
QMenuBar {
    background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                      stop:0 lightblue, stop:1 darkblue);
    spacing: 11px; /* spacing between menu bar items */
    margin: 1px 1px;
    color: yellow;
    font: bold 18px;
}
QMenuBar::item {
    margin: 0px 0px;
    padding: 1px 1px;
    background: transparent;
    border-radius: 2px;
}
QMenuBar::item:selected {
    background: #C9C9C9;
}
QMenu {
    color: blue;
    font: bold 12px;
    font: bold 14px;
}
QMenu:selected {
    background: #FFFFD9;
}
"""

        self.setAutoFillBackground(True)
        self.setBackgroundRole(QPalette.Highlight)
        self.setStyleSheet(css) 
        self.setFixedHeight(BAR_H)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        self.mMoving = False
        self.win = win

    ##########################################################################
    def mousePressEvent(self, evt):
        super(TitleBar, self).mousePressEvent(evt)
        if evt.buttons() == Qt.LeftButton:
            self.mMoving = True
            self.setMouseTracking ( self.mMoving )
            self.mLastMousePosition = evt.globalPos()

    ##########################################################################
    def mouseMoveEvent(self, evt):
        super(TitleBar, self).mouseMoveEvent(evt)
        if self.mMoving and evt.buttons() == Qt.LeftButton:
            self.win.move(self.win.pos() + (evt.globalPos() - self.mLastMousePosition))
            self.mLastMousePosition = evt.globalPos()

    ##########################################################################
    def mouseReleaseEvent(self, evt):
        super(TitleBar, self).mouseReleaseEvent(evt)
        if evt.button() == Qt.LeftButton:
            self.mMoving = False
            self.setMouseTracking ( self.mMoving )

##############################################################################
# End of module                                                              #
##############################################################################
