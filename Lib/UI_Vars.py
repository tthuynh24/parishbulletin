"""

 NAME: UI_Vars.py

 REVISION HISTORY:
  dd-mmm-yyyy BCR/PR  Created      Comments
  ----------- ------- ------------ -------------------------------------------
******************************************************************************
"""
from PyQt5.QtGui import QColor

STAT_TAB, ADM_TAB = range(1000, 1002)

PrioColour = {1:QColor('#FFC9C9'),
              2:QColor('#F1F1B1'),
              3:QColor('#DFFFCA'),
              4:QColor('#A6DEEE'),
              5:QColor('#A6D000'),
              }

STATTAB    = 'Status'
ADMTAB     = 'Admin'

MyIco = {STAT_TAB:      'Church',
         ADM_TAB:       'Church',
        }

BAR_H = 30
ICON_SIZE = 26
BTN_SIZE = 28


##############################################################################
# End of File
##############################################################################
