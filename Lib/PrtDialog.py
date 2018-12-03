#!/usr/bin/env python
##############################################################################
# Module: PrtDialog                                                            #
##############################################################################
from Utils.QtUtils import *

from Lib.Dialog import Dialog

class PrtDialog(Dialog):

    def __init__(self):
        super(PrtDialog, self).__init__()
        self.setWindowTitle('PrtDialog')
        self.setWindowIcon(createIcon('SeaFood'))

        self.mainLayout = BoxLayout('V', spacing=10, margin=(0, 0, 0, 0))
        
        AddWidget(self.mainLayout, self.__addMemWidgets())
        AddWidget(self.mainLayout, self.__addAddrWidgets())

        AddLayout(self.getLayout(), self.mainLayout)
        self.resize(725, 375)
        
    ##########################################################################
    def __addMemWidgets(self):
        hbox = BoxLayout('V', spacing=0, margin=(0, 0, 0, 0))
        self.frmFam = MyWidget(hbox, ComboBox, 'From Member', minWidth=125)[-1]
        self.toFam = MyWidget(hbox, ComboBox, 'To Member', minWidth=125)[-1]
        return GroupBox('Member(s) Selection', hbox)

        
    ##########################################################################
    def __addAddrWidgets(self):
        vbox = BoxLayout('V', spacing=0, margin=(0, 0, 0, 0))
        self.memNm = MyWidget(vbox, LineEdit, 'Member Name', minWidth=125)[-1]
        self.toFam = MyWidget(vbox, LineEdit, 'Address 1', minWidth=125)[-1]
        self.toFam = MyWidget(vbox, LineEdit, 'Address 2', minWidth=125)[-1]
        self.toFam = MyWidget(vbox, LineEdit, 'City', minWidth=125)[-1]
        self.toFam = MyWidget(vbox, LineEdit, 'State', minWidth=125)[-1]
        self.toFam = MyWidget(vbox, LineEdit, 'Zipcode', minWidth=125)[-1]
        return GroupBox('Member Address', vbox)
        
##############################################################################
# End of module                                                              #
##############################################################################
