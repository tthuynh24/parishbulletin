#!/usr/bin/env python
##############################################################################
# Module: Dialog                                                            #
##############################################################################
from PyQt5.QtWidgets import QDialog

from .TitleBar import TitleBar

from Utils.QtUtils import (customizeWin, BoxLayout, Button, AddWidget,
                           AddLayout, GroupBox, createIcon, Menu, Label)

from .UI_Vars import *

class Dialog(QDialog):

    def __init__(self):
        super().__init__()
        customizeWin(self)

        self.__createMenu()
        self.__mainLayout = BoxLayout('V', spacing=0, margins=(0, 0, 0, 0))
        AddWidget(self.__mainLayout, Label(None, maxHeight=11), align='Top')
        self.__mainLayout.setMenuBar(self.__menuBar)
        self.setLayout(self.__mainLayout)

    ##########################################################################
    def __addDlgBtn(self):
        self.__MaxIco = createIcon('Max')
        self.__NormIco = createIcon('Norm')
        self.__maxNormal=  False

        menuLo = BoxLayout('H', spacing=0, margins=(0, 0, 0, 0))
        exitBtn = Button(icon=createIcon('Exit'), parent=self)
        exitBtn.setFixedSize(BTN_SIZE,BTN_SIZE)
        exitBtn.pressed.connect(self.__onExit)
        self.__maximize = Button(icon=self.__MaxIco, parent=self)
        self.__maximize.setFixedSize(BTN_SIZE,BTN_SIZE)
        self.__maximize.pressed.connect(self.__onMaximize)
        minBtn = Button(icon=createIcon('Min'), parent=self)
        minBtn.setFixedSize(BTN_SIZE,BTN_SIZE)
        minBtn.pressed.connect(self.__onMinimize)
        AddWidget(menuLo, minBtn)
        AddWidget(menuLo, self.__maximize)
        AddWidget(menuLo, exitBtn)
        
        #self.menuLayout = BoxLayout('H', spacing=0, margins=(0, 0, 0, 0))
        #AddLayout(self.menuLayout, menuLo)
        
        return GroupBox(None, menuLo)

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
        self.close()
        
    def getLayout(self):
        return self.__mainLayout
    
    def getMenuBar(self):
        return self.__menuBar
    
        
##############################################################################
# End of module                                                              #
##############################################################################
