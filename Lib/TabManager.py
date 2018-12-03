"""
******************************************************************************
 NAME: TabManager.py
******************************************************************************
 REVISION HISTORY:
  dd-mmm-yyyy BCR/PR  Created      Comments
  ----------- ------- ------------ -------------------------------------------
  05-July-2009         Hien Pham     Initial version
******************************************************************************
"""
##############################################################################
from .AppShared import *
from Utils.PyUtils import *
from Utils.QtUtils import *

STAT_TAB, ACT_TAB = 'Status Tab', 'Action Tab'

##############################################################################
##############################################################################
class TabManager(QTabWidget):
    CONN_STAT = {1:     'orange',
                 2:     'green',
                 3:     'red',
                 4:    'blue'
                 }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName('TabManager')
        self.tabBar = QTabBar(self)
        self.tabBar.setMovable(True)
        self.setTabBar(self.tabBar)

        self.setStyleSheet("""
                            QTabBar::tab {padding: 1px 1px 1px 1px;
                                          font-size: 10pt;
                                          font-weight: bold;
                                          width: 150px;
                                          background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                                                      stop: 0 #fafafa, stop: 0.4 #f4f4f4,
                                                                      stop: 0.5 #e7e7e7, stop: 1.0 #fafafa);
                                          }
                            QTabBar::tab:selected {background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                                                                               stop: 0 cyan, stop: 1.0 blue);
                                                   }
                            QTabBar::tab:selected {
                                                   color: yellow;
                                                   border-color: #9B9B9B;
                                                   border-bottom-color: #C2C7CB;
                                                   }
                            QTabBar::tab:!selected {
                                                   color: grey;
                                                   margin-top: 2px;
                                                   }
                            QTabBar::tab:selected {
                                                   margin-left: -4px;
                                                   margin-right: -4px;
                                                   }
                            """)

    ##########################################################################
    def hideTab(self, tabInd):
        self.removeTab(tabInd)

    ##########################################################################
    def addNewTab(self, tabWidget, tabLbl):

        tabIcon = createIcon(MyIco.get(tabLbl, 'SeaFood'))
        self.setUpdatesEnabled(False)
        self.addTab(tabWidget, tabIcon, tabLbl)
        self.setUpdatesEnabled(True)
        tabInd = self.tabBar.count()-1
        self.setCurrentIndex(tabInd)
        self.tabBar.setTabTextColor(tabInd, QColor('blue'))
        self.setCurrentIndex(tabInd)

        return tabInd

    ##########################################################################
    def delTab(self, tabId=None):
        tabInd = self.tabIndex(tabId) if tabId else self.currentIndex()
        if tabInd>=0:
            self.widget(tabInd).deleteLater()
            self.removeTab(tabInd)
            self.tabRemoved(tabInd)

    ##########################################################################
    def setTabTitle(self, tabInd, tabTitle):
        self.tabBar.setTabText(tabInd, tabTitle)

    ##########################################################################
    def setActTab(self, tabInd):
        if tabInd>=0:
            self.setCurrentIndex(tabInd)

    ##########################################################################
    def terminate(self):
        activeTab = self.currentWidget()
        while activeTab:
            tabInd = self.currentIndex()
            self.widget(tabInd).deleteLater()
            self.removeTab(tabInd)
            self.tabRemoved(tabInd)

            del activeTab
            activeTab = self.currentWidget()

##############################################################################
# End of File
##############################################################################
