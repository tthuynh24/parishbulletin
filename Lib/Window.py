#!/usr/bin/env python
##############################################################################
# Module: Window                                                            #
##############################################################################
from PyQt5.QtWidgets import QMainWindow

from .TitleBar import TitleBar

from Utils.QtUtils import (customizeWin, BoxLayout, Button, Widget, AddWidget,
                           GroupBox, createIcon, Menu)
from .UI_Vars import *

class Window(QMainWindow):

    def __init__(self):
        super(Window, self).__init__()
        customizeWin(self)

        mainWidget=Widget(self)
        self.setCentralWidget(mainWidget)
        
        self.__createMenu()
        self.__mainLayout = BoxLayout('V', spacing=0, margins=(0, 0, 0, 0))
        self.__mainLayout.setMenuBar(self.__menuBar)
        mainWidget.setLayout(self.__mainLayout)

    ##########################################################################
    def __addDlgBtn(self):
        self.__MaxIco = createIcon('Max')
        self.__NormIco = createIcon('Norm')
        self.__maxNormal=  False

        self.menuLayout = BoxLayout('H', spacing=6, margins=(0, 0, 0, 0))
        exitBtn = Button(icon=createIcon('Exit'), parent=self)
        exitBtn.setFixedSize(BTN_SIZE,BTN_SIZE)
        exitBtn.pressed.connect(self.__onExit)
        self.__maximize = Button(icon=self.__MaxIco, parent=self)
        self.__maximize.setFixedSize(BTN_SIZE,BTN_SIZE)
        self.__maximize.pressed.connect(self.__onMaximize)
        minBtn = Button(icon=createIcon('Min'), parent=self)
        minBtn.setFixedSize(BTN_SIZE,BTN_SIZE)
        minBtn.pressed.connect(self.__onMinimize)
        AddWidget(self.menuLayout, minBtn)
        AddWidget(self.menuLayout, self.__maximize)
        AddWidget(self.menuLayout, exitBtn)
        return GroupBox(None, self.menuLayout)

    def __createMenu(self):
        self.__menuBar = TitleBar(self)               # QmenuBar
        self.__menuBar.setFixedHeight(BAR_H)
        self.__menuBar.setCornerWidget(self.__addDlgBtn())

    def __onMinimize(self):
        self.showMinimized()

    def __onMaximize(self):
        if(self.__maxNormal):
            self.showNormal()
            self.__maxNormal= False
            self.__maximize.setIcon(self.__MaxIco)
        else:
            self.showMaximized()
            self.__maxNormal=  True
            self.__maximize.setIcon(self.__NormIco)

    def __onExit(self):
        self._saveDoc()
        self.close()
        
    def getLayout(self):
        return self.__mainLayout
        
    def getMenuBar(self):
        return self.__menuBar

    # To be implemented by inheritor class
    def _saveDoc(self):
        print('====>> _saveDoc...')
        pass

##############################################################################
# End of module                                                              #
##############################################################################
