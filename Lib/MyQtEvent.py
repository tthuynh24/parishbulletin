from Utils.QtUtils import *

from Utils.PyUtils import getTrace, pause

MYEVENT = QEvent.User
class MyQtEvent(QEvent):
    def __init__(self, data=None):
        super().__init__(MYEVENT)
        self.data = data

    def type(self):
        return MYEVENT

def emitMyEvt(reciever, evtData):
    QApplication.postEvent(reciever, MyQtEvent(evtData))

##############################################################################
# End of File
##############################################################################
