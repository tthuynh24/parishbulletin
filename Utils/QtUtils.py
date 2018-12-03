#!/usr/bin/env python
##############################################################################
# Module: QtUtils                                                            #
##############################################################################
from __future__ import absolute_import

import os
import re
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from datetime import datetime

from .PyUtils import joinPaths, isFile, isNixOs, addToLst, isValid, isNull, isNotNull
from Lib.AppShared import *
from Lib.UI_Vars import *

from. import bogo
TypeRules = bogo.core.get_vni_definition()

GuiFont = {}
for size in (8, 9, 10, 11, 12, 14, 16):
    GuiFont[size] = QFont('Helvetica [Cronyx]', size, QFont.Bold)

_fromUtf8 = lambda s: s

FontSize = 10

''' Creating Widgets Margins(left, top, right, bottom)'''
MSG_CRIT, MSG_ERROR, MSG_FAIL, MSG_CONF, MSG_WARN, MSG_INFO, MSG_DFLT = range(1,8)
MSG_RE = re.compile('(CRITICAL)|(ERROR)|(Sorry)|(NOTE)|(WARNING)|(INFO)', re.I)

IconDict = {MSG_DFLT:'Smiley_OK',
            MSG_CONF:'Curious',
            MSG_WARN:'Curious',
            MSG_FAIL:'Smiley_NOK',
            MSG_ERROR:'Smiley_NOK',
            MSG_CRIT:'Smiley_NOK',
            }

COLOUR = {MSG_DFLT:'blue',
          MSG_CONF:'orange',
          MSG_WARN:'orange',
          MSG_FAIL:'red',
          MSG_ERROR:'red',
          MSG_CRIT:'red',
         }

NM, ICO, HDL = range(3)
FontSize = 10 if isNixOs() else 8

PROMPT, CONFIRM, SAVECONF = range(3)

LOG_STAT, LOG_MSG = 60, 70

###########################################################################
def AddLayout(pLayout, cLayout, stretch=0.5):
    pLayout.addLayout(cLayout, stretch=stretch)

###########################################################################
def AddMenu(myMenu, menuLst, actLst, preFix=None, menuAct={}):
    for ind, menu in enumerate(menuLst):
        if isNull(menu):
            myMenu.addSeparator()
        else:
            act = addMenuItem(myMenu, menu)
            if isNotNull(preFix):
                menu = preFix + menu
            menuAct[menu] = (act, actLst[ind])

###########################################################################
def AddMenuSep(myMenu):
    myMenu.addSeparator()

###########################################################################
def AddSubMenu(menuBar, menuNm, menuLst, actLst, menuAct={}, preFixMenuNm=False):
    myMenu=QMenu(menuNm, menuBar)
    preFix = menuNm if preFixMenuNm else None
    AddMenu(myMenu, menuLst, actLst, preFix=preFix, menuAct=menuAct)
    try:
        image = UI_Icons.get(menuNm, None)
    except:
        image = None
    if isNotNull(image):
        myMenu.setIcon(createIcon(image))
    menuBar.addMenu(myMenu)
    return myMenu

###########################################################################
def AddWidget(layout, widget, align=None, stretch=0):
    align = align if isNull(align) else align.lower()
    if align in AlignOpts:
        layout.addWidget(widget, stretch=stretch, alignment=AlignOpts[align])
    else:
        layout.addWidget(widget, stretch=stretch)

###########################################################################
def BoxLayout(view, parent=None, **args):
    layout = QVBoxLayout(parent) if view.upper() in ('VIRTICAL', 'V') else QHBoxLayout(parent)
    return setWidgetOpts(layout, args)

###########################################################################
def BrowserWidget(parent, label, labelWidth=140, tip=None, selDir=True):
    if isNull(tip):
        tip = 'Navigate to select "%s"' % ('Directory' if selDir else 'File')
    layout = BoxLayout('H')
    label = CustLabel(label, 10, fixedWidth=labelWidth, align='right')
    selPath = LineEdit()
    browser = Button(label='Sel Dir' if selDir else 'Sel File',
                     fixedWidth=75, tooltip=tip)
    [layout.addWidget(w) for w in (label, selPath, browser)]
    parent.addWidget(GroupBox('', layout))

    return layout, label, selPath, browser

###########################################################################
def Button(parent=None, **args):
    label = args.get('label', '')
    icon = args.get('icon', None)
    pushButton = QPushButton(icon, label, parent=parent) if icon else QPushButton(label, parent=parent)
    return setWidgetOpts(pushButton, args)

###########################################################################
def CheckBox(parent=None, **args):
    chkBox = args.get('label', '')
    chkBox = QCheckBox(chkBox, parent=parent)
    return setWidgetOpts(chkBox, args)

###########################################################################
def ChkLstGrp(grpNm, chkItems, dflt, hdl=None, view='V'):
    chkGrpDct = {}
    layout = BoxLayout(view, spacing=21, margins=6)
    lbl = CustLabel(grpNm, size=10)
    AddWidget(layout, lbl, align='right')
    for chkItem, tip in chkItems:
        chkBtn = CheckBox(label=chkItem, tooltip=tip)
        chkBtn.setObjectName(chkItem)
        AddWidget(layout, chkBtn)
        if chkItem==dflt:
            setCheckState(chkBtn, True)
        if hdl:
            chkBtn.clicked.connect(hdl)
        chkGrpDct[chkItem] = chkBtn
    return (GroupBox('', layout), chkGrpDct)

###########################################################################
def ComboBox(parent=None, **args):
    combo = QComboBox(parent=parent)
    return setWidgetOpts(combo, args)



###########################################################################
def DateTime(parent=None, **args):
    datTime = QDateTime(parent=parent)
    return setWidgetOpts(datTime, args)




##########################################################################
def MyWidget(parentLo, wdgt, label, xWdgt=None, labelWidth=175, tip=None, addGrp=True, **args):
    G = 2
    if isNull(tip):
        tip = 'Navigate to select "%s"' % label
    layout = BoxLayout('H', spacing=G, margins=(G, G, G, G))
    label = CustLabel(label+':', 11, fixedWidth=labelWidth, align='right')
    selecTor = wdgt(**args)
    [layout.addWidget(w) for w in (label, selecTor, xWdgt) if w is not None]
    if addGrp:
        parentLo.addWidget(GroupBox('', layout))
    else:
        parentLo.addLayout(layout)
    
    return label, selecTor



###########################################################################
def ComboWidget(parent, label, labelWidth=140, tip=None, addGrp=True, **args):
    if isNull(tip):
        tip = 'Navigate to select "%s"' % label
    layout = BoxLayout('H')
    label = CustLabel(label, 10, fixedWidth=labelWidth, align='right')
    selecTor = ComboBox(**args)
    [layout.addWidget(w) for w in (label, selecTor)]
    if addGrp:
        parent.addWidget(GroupBox('', layout))
    else:
        parent.addWidget(label)
        parent.addWidget(selecTor)
    return label, selecTor

###########################################################################
def ConfirmDia(msg, colour='orange', caption='WARNING - WARNING - WARNING', dflt=True):
    dlg = MsgDialog(caption, 'Curious', msg, colour, CONFIRM)
    return dlg._getUserAns(CONFIRM)

###########################################################################
def CustLabel(label, size=16, colour='blue', parent=None, **args):
    label = QLabel(fmtLabel(label, colour=colour), parent=parent)
    label.setObjectName('Customized')
    args['font'] = QFont('Monaco', size, QFont.Bold)
    if isNull(args.get('align', None)):
        args['align'] = 'center'
    return setWidgetOpts(label, args)

###########################################################################
def DirSel(parent, label, oName, selHdl, width=115):
    return FileDirSel(parent, label, oName, selHdl, width, fileSel=False)

###########################################################################
def EditWidget(parent=None, **args):
    label = args.get('label', None)
    text = args.get('text', '')
    lineEdit = QLineEdit(text, parent=parent)
    setWidgetOpts(lineEdit, args)
    if isNull(label):
        return lineEdit

    layout = BoxLayout('H', spacing=4, margins=4)
    label = CustLabel(label, 10, fixedWidth=FontSize*len(label))
    AddWidget(layout, label)
    AddWidget(layout, lineEdit)

    return lineEdit, layout

###########################################################################
def FileCombo(parent, label, labelWidth=140, tip=None):
    tip = 'Navigate to select the "Source File Directory"'
    layout = BoxLayout('H')
    label = CustLabel(label, 10, fixedWidth=labelWidth)
    selFile = ComboBox(tooltip=tip)
    browser = Button(label='Sel Dir', fixedWidth=75, tooltip=tip)
    [layout.addWidget(w) for w in (label, selFile, browser)]
    parent.addWidget(GroupBox('', layout))

    return label, selFile, browser

###########################################################################
def FileDirSel(parent, label, oName, selHdl, width=115, fileSel=True):
    selLayout = BoxLayout('H')
    fmtLable = CustLabel(label, 10, fixedWidth=width, align='right')
    AddWidget(selLayout, fmtLable)
    dirSel = ComboBox(editable=True, minWidth=325)
    icon = 'OpenFile' if fileSel else 'OpenFolder'
    selButn = Button(icon=createIcon(icon),
                    fixedWidth=25,
                    tooltip='Navigate to select %s file' % label,
                    objectname=oName)
    selButn.clicked.connect(selHdl)
    AddWidget(selLayout, dirSel)
    AddWidget(selLayout, selButn)
    return dirSel, selLayout

###########################################################################
def FileSel(parent, label, oName, selHdl, width=115):
    return FileDirSel(parent, label, oName, selHdl, width)

###########################################################################
def FontCombo(parent=None, **args):
    combo = QFontComboBox(parent=parent)
    return setWidgetOpts(combo, args)

###########################################################################
def GetLstSels(qList, selItems=[]):
    for ind in range(qList.count()):
        if isCheckState(qList.item(ind), True):
            addToLst(selItems, qList.item(ind).text())

    if len(selItems)==0:
        for selItem in qList.selectedItems():
            addToLst(selItems, selItem.text())
    return selItems[:]

###########################################################################
def GroupBox(label, layout, parent=None, **args):
    grpBox = QGroupBox(label, parent)
    grpBox.setLayout(layout)
    return setWidgetOpts(grpBox, args)

###########################################################################
def Label(parent=None, **args):
    label = args.get('label', '')
    label = QLabel(label, parent=parent)
    return setWidgetOpts(label, args)

###########################################################################
def LineEdit(parent=None, **args):
    lineEdit = QLineEdit(args.get('text', ''), parent=parent)
    return setWidgetOpts(lineEdit, args)

###########################################################################
def ListGroup(grpNm, view='V', mSels=True, **args):
    layout = 'V' if view.upper() in ('VIRTICAL', 'V') else 'H'
    args.update({'selectionMode':'MultiSel' if mSels else 'SingleSel'})
    qList = ListWidget(None, **args)
    lo = BoxLayout(layout)
    AddWidget(lo, qList)
    return GroupBox(grpNm, lo), qList

###########################################################################
def ListItem(itemNm, **args):
    item = QListWidgetItem(itemNm)
    return setWidgetOpts(item, args)

###########################################################################
def ListWidget(parent=None, **args):
    qList = QListWidget(parent=parent)
    return setWidgetOpts(qList, args)

###########################################################################
def Menu(menuNm, parent=None, **args):
    myMenu=QMenu(menuNm, parent)
    return setWidgetOpts(myMenu, args)

###########################################################################
def MenuBar(parent=None, **args):
    menuBar = QMenuBar(parent=parent)
    return setWidgetOpts(menuBar, args)

###########################################################################
def ProgressBar(range, parent=None, **args):
    progressBar = QProgressBar(parent=parent)
    progressBar.setRange(range[0], range[1])
    return setWidgetOpts(progressBar, args)

###########################################################################
def RadioBox(grpNm, radioButtons, dflt, label=None, hdl=None, view='V', **args):
    radioDct = {}
    layout = BoxLayout(view, spacing=4, margins=4)
    if label is None:
        layout.addWidget(Label(fixedWidth=args.get('spacer', 10)))
    else:
        layout.addWidget(label)
    for btnNm in radioButtons:
        radioBtn = addRadioBtn(btnNm, objectName=btnNm)
        radioBtn.setObjectName(btnNm)
        layout.addWidget(radioBtn)
        if btnNm==dflt:
            radioBtn.setChecked(True)
        if hdl:
            radioBtn.clicked.connect(hdl)
        radioDct[btnNm] = radioBtn
    return (GroupBox(grpNm, layout) if isNotNull(grpNm) else layout, radioDct)

###########################################################################
def SaveConfirm(msg, caption='Save Confirm Dialog'):
    dlg = MsgDialog(caption, 'Curious', msg, 'blue', SAVECONF)
    return dlg._getUserAns()


###########################################################################
def SelColour(currColor=Qt.blue, parent=None):
    dlg = QColorDialog.getColor(currColor, parent=parent)
    return dlg.name() if dlg.isValid() else None

###########################################################################
def SelFont(currFont=GuiFont[10], parent=None):
    font, dlgOk = QFontDialog.getFont(currFont, parent=parent)
    return font if dlgOk else None

###########################################################################
def SetDialogStaysOnTop(dialog):
    dialog.setWindowFlags(Qt.Dialog | Qt.WindowStaysOnTopHint)

###########################################################################
def Splitter(view, parent=None, **args):
    splitter = QSplitter(parent=parent)
    splitter.setOrientation(Qt.Vertical if view.upper() in ('VIRTICAL', 'V') else Qt.Horizontal)
    return setWidgetOpts(splitter, args)



'''
Qt::Widget     0x00000000     This is the default type for QWidget. Widgets of this type are child widgets if they have a parent, and independent windows if they have no parent. See also Qt::Window and Qt::SubWindow.
Qt::Window     0x00000001     Indicates that the widget is a window, usually with a window system frame and a title bar, irrespective of whether the widget has a parent or not. Note that it is not possible to unset this flag if the widget does not have a parent.
Qt::Dialog     0x00000002 | Window     Indicates that the widget is a window that should be decorated as a dialog (i.e., typically no maximize or minimize buttons in the title bar). This is the default type for QDialog. If you want to use it as a modal dialog, it should be launched from another window, or have a parent and used with the QWidget::windowModality property. If you make it modal, the dialog will prevent other top-level windows in the application from getting any input. We refer to a top-level window that has a parent as a secondary window.
Qt::Sheet     0x00000004 | Window     Indicates that the window is a Macintosh sheet. Since using a sheet implies window modality, the recommended way is to use QWidget::setWindowModality(), or QDialog::open(), instead.
Qt::Drawer     0x00000006 | Window     Indicates that the widget is a Macintosh drawer.
Qt::Popup     0x00000008 | Window     Indicates that the widget is a pop-up top-level window, i.e. that it is modal, but has a window system frame appropriate for pop-up menus.
Qt::Tool     0x0000000a | Window     Indicates that the widget is a tool window. A tool window is often a small window with a smaller than usual title bar and decoration, typically used for collections of tool buttons. If there is a parent, the tool window will always be kept on top of it. If there isn't a parent, you may consider using Qt::WindowStaysOnTopHint as well. If the window system supports it, a tool window can be decorated with a somewhat lighter frame. It can also be combined with Qt::FramelessWindowHint.

On Mac OS X, tool windows correspond to the Floating class of windows. This means that the window lives on a level above normal windows; it impossible to put a normal window on top of it. By default, tool windows will disappear when the application is inactive. This can be controlled by the Qt::WA_MacAlwaysShowToolWindow attribute.
Qt::ToolTip     0x0000000c | Window     Indicates that the widget is a tooltip. This is used internally to implement tooltips.
Qt::SplashScreen     0x0000000e | Window     Indicates that the window is a splash screen. This is the default type for QSplashScreen.
Qt::Desktop     0x00000010 | Window     Indicates that this widget is the desktop. This is the type for QDesktopWidget.
Qt::SubWindow     0x00000012     Indicates that this widget is a sub-window, such as a QMdiSubWindow widget.

'''
###########################################################################
def Table(numRows, tblNm, maxWidth=0, maxHeight=0, toolTip=None, editable=True, parent=None):
    tableLayout = QHBoxLayout(parent=parent)
    tbl = QTableWidget(numRows, 1)
    tbl.verticalHeader().hide()
    tbl.setAlternatingRowColors(True)
    tbl.setFont(QFont('Monaco', 12, 75))
    tbl.setHorizontalHeaderLabels((tblNm,))
    tbl.setFont(QFont('Monaco', 12))
    if isValid(toolTip, str):
        tbl.setToolTip(toolTip)
    tbl.takeVerticalHeaderItem(0)
    if maxWidth:
        tbl.setMaximumWidth(maxWidth)
    if maxHeight:
        tbl.setMaximumHeight(maxHeight)
    tableLayout.addWidget(tbl)
    tbl.setColumnWidth(0, maxWidth-25)
    if not editable:
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
    for row in range(numRows):
        tbl.setRowHeight(row, 21)

    return tableLayout

###########################################################################
def TextWidget(label, btnObj=None, textSize=16, labelWidth=125, wdgtWidth=None, **args):
    ratio = textSize/16
    layout = BoxLayout('H')
    label = CustLabel(label, textSize, fixedWidth=labelWidth, align='right')
    lineEdit = LineEdit() if wdgtWidth is None else LineEdit(fixedWidth=wdgtWidth)
    setWidgetOpts(lineEdit, args)
    [layout.addWidget(w) for w in (Label(), label, lineEdit, Label(), btnObj) if w is not None]
    return  GroupBox('', layout, fixedHeight=55*ratio), lineEdit

###########################################################################
def Textedit(parent=None, **args):
    textedit = QTextEdit(parent=parent)
    return setWidgetOpts(textedit, args)

###########################################################################
def ToolButton(parent=None, **args):
    tButton = QToolButton(parent=parent)
    return setWidgetOpts(tButton, args)

###########################################################################
def TreeWidget(parent=None, **args):
    treeWidget = QTreeWidget(parent=parent)
    return setWidgetOpts(treeWidget, args)

###########################################################################
def TreeWidgetItem(parent, **args):
    treeWItem = QTreeWidgetItem(parent)
    return setWidgetOpts(treeWItem, args)

###########################################################################
def Widget(parent=None, **args):
    wdgt = QWidget(parent)
    return setWidgetOpts(wdgt, args)
    
###########################################################################
def addFileSel(parent, label, oName, selHdl, width=115):
    selLayout = BoxLayout('H')
    fmtLable = CustLabel(label, 10, fixedWidth=width, align='right')
    selLayout.addWidget(fmtLable)
    selCombo = ComboBox(editable=True, minWidth=325)
    selButn = Button(icon=createIcon('OpenFile'),
                     fixedWidth=25,
                     tooltip='Navigate to select %s file' % label,
                     objectname=oName)
    parent.connect(selButn, SIGNAL('clicked()'), selHdl)
    selLayout.addWidget(selCombo)
    selLayout.addWidget(selButn)
    return selLayout, fmtLable, selCombo

###########################################################################
def addItems(widget, items):
    widget.addItems(items)

###########################################################################
def addMenuItem(selMenu, menu):
    try:
        image = UI_Icons.get(menu, None)
    except:
        image = None
    if isNotNull(image):
        return selMenu.addAction(createIcon(image), str(menu))
    return selMenu.addAction(str(menu))

###########################################################################
def addRadioBtn(btnNm, **args):
    radioBtn = QRadioButton(btnNm)
    return setWidgetOpts(radioBtn, args)

###########################################################################
def buttonEvt(button, evtHdl, onEvt):
    evtHdl.connect(button, SIGNAL('clicked()'), onEvt)

###########################################################################
def checkButtonEvt(button, evtHdl, onEvt):
    evtHdl.connect(button, SIGNAL('stateChanged(int)'), onEvt)

###########################################################################
def chkCombo(parent, label, chkLabel='Check', labelWidth=125, tip=None):
    if isNull(tip):
        tip = 'Navigate to select "%s"' % label
    layout = BoxLayout('H')
    label1 = CustLabel(label, 10, align='left')
    chkBox = CheckBox(fixedWidth=25)
    label2 = CustLabel(chkLabel, 10, fixedWidth=labelWidth, align='right')
    selecTor = ComboBox()
    [layout.addWidget(w) for w in (label1, label2, chkBox, selecTor)]
    parent.addWidget(GroupBox('', layout))
    return label2, chkBox, selecTor

###########################################################################
def comboEvt(combo, evtHdl, onEvt):
    evtHdl.connect(combo, SIGNAL('activated(QString)'), onEvt)

###########################################################################
def createContextMenu(widget, contextMenusSpec, enableCmd):
    menu = QMenu()
    for contextSpec in contextMenusSpec:
        if contextSpec==None:
            menu.addSeparator()
        else:
            menu.addAction(create_action(widget, contextSpec[NM],
                                         icon=createIcon(UI_Icons[contextSpec[NM]]),
                                         triggered=contextSpec[HDL],
                                         enable=contextSpec[NM] in enableCmd))
    return menu

###########################################################################
def createFont(fontFamily, fSize, boldOn, italicOn):
    Qfont(fontFamily, int(fSize), (75 if boldOn else 50), italicOn)


''' Common Widget Option setting '''
###########################################################################
def createLine(view='H', parent=None):
    lineView = QFrame.VLine if view.upper() in ('VIRTICAL', 'V') else QFrame.HLine
    line = QFrame(parent=parent)
    line.setFrameShape(lineView)
    line.setFrameShadow(QFrame.Sunken)
    return line

###########################################################################
def create_action(widget, text, shortcut=None, icon=None, tip=None,
                  toggled=None, triggered=None, data=None,
                  window_context=True, enable=True):
    """Create a QAction"""
    action = QAction(text, widget)
    if isNotNull(triggered):
        widget.connect(action, SIGNAL("triggered()"), triggered)
    if isNotNull(toggled):
        widget.connect(action, SIGNAL("toggled(bool)"), toggled)
        action.setCheckable(True)
    if isNotNull(icon):
        if isValid(icon, str):
            icon = get_icon(icon)
        action.setIcon(icon)
    if isNotNull(shortcut):
        action.setShortcut(shortcut)
    if isNotNull(tip):
        action.setToolTip(tip)
        action.setStatusTip(tip)
    if isNotNull(data):
        action.setData(QVariant(data))
    if window_context:
        action.setShortcutContext(Qt.WindowShortcut)
    else:
        #TODO: Hard-code all shortcuts and choose window_context=False
        # (this will avoid calling shortcuts from another dockwidget
        #  since the context thing doesn't work quite well with these)
        action.setShortcutContext(Qt.WidgetShortcut)
    if enable is False:
        setEnabled(action, enable)
    return action

###########################################################################
def customizeDia(dlg):
    dlg.setWindowFlags(Qt.Window | Qt.CustomizeWindowHint | Qt.WindowTitleHint |
                       Qt.WindowSystemMenuHint | Qt.WindowStaysOnTopHint)

###########################################################################
def customizeWin(win):
    win.setWindowFlags(Qt.CustomizeWindowHint)             # Set window with no title bar

###########################################################################
def deleteLayout(layout):
    if isNotNull(layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if isNotNull(widget):
                widget.deleteLater()

###########################################################################
def fmtLabel(label, colour='blue', size=None):
    if isNull(size):
        return '<font color=\"%s\"><b>%s</b></font>' % (colour, label)
    return '<font color=\"%s\" size=\"%s\"><b>%s</b></font>' % (colour, size, label)

###########################################################################
def getActiveWidget(win):
    widget = win.focusWidget()
    if isNull(widget):
        return None, None
    widgetNm = str(widget.objectName())
    return widget, widgetNm

###########################################################################
def getAvalFonts():
    return QFontDatabase().families()

###########################################################################
def getCursorPos():
    return QMouseEvent.pos()

###########################################################################
def getFontInfo():
    myFont = QApplication.font()
    return myFont.family(), myFont.pointSize()

###########################################################################
def getMsgInd(**opts):
    msg = opts.get('msg')
    return opts.get('msgType') if isNull(msg) else setMsgType(msg)[-1]

###########################################################################
def getSelList(combo):
    return tuple([combo.itemText(i) for i in range(combo.count())])

###########################################################################
def get_icon( name, default=None ):
    """Return image inside a QIcon object"""
    if isNull(default):
        return QIcon(get_image_path(name))
    elif isValid(default, QIcon):
        icon_path = get_image_path(name, default=None)
        return default if isNull(icon_path) else QIcon(icon_path)
    else:
        return QIcon(get_image_path(name, default))

###########################################################################
def InsertWidget(layout, widget, loc=0, stretch=0, align='right'):
    layout.insertWidget(loc, widget, stretch=stretch, alignment=AlignOpts.get(align))
    return widget
   
###########################################################################
def get_image_path( imgNm, imgPath='icons' ):
    """Return image absolute path"""
    full_path = joinPaths(imgPath, imgNm)
    if isFile(full_path):
        return os.path.abspath(full_path)

###########################################################################
def isCheckState(widget, checkState):
    setState = Qt.Checked if checkState else Qt.Unchecked
    return widget.checkState()==setState

###########################################################################
def isErrMsg(msg):
    msgInd = setMsgType(msg)[-1]
    return msgInd <= MSG_FAIL

###########################################################################
def keybinding(attr):
    """Return keybinding"""
    ks = getattr(QKeySequence, attr)
    return QKeySequence.keyBindings(ks)[0].toString()

###########################################################################
def loadSelLst(selWdgt, selLst, selVal=0, isCombo=True, sortLst=True):
    if len(selLst)>0:
        if isNull(selVal):
            selVal = 0
        selWdgt.clear()

        selWdgt.addItems(sorted(list(selLst)) if sortLst else selLst)
        if isValid(selVal, int, chkVal=False):
            loader = selWdgt.setCurrentIndex if isCombo else selWdgt.setCurrentRow
            loader(selVal)
        elif isValid(selVal, str, chkVal=False):
            selWdgt.setEditText(selVal)
        selWdgt.update()

###########################################################################
def moveCenter(wdgt):
    screen, dialog = QDesktopWidget().availableGeometry(), wdgt.size()
    wdgt.move(screen.width()/2 - dialog.width()/2,
              screen.height()/2 - dialog.height()/2)
###########################################################################
def promptUser(msg, selCaption=None):
    caption, msgType = setMsgType(msg)
    if selCaption is not None:
        caption = selCaption
    dlg = MsgDialog(caption, setMsgIcon(msgType=msgType),
                    msg, setMsgColour(msgType=msgType), PROMPT)
    return dlg._getUserAns()

###########################################################################
def selFont(xOS):
    availFonts =  getAvalFonts()
    for fontNm in ('Monaco', 'Monospace', 'System', 'Arial', 'Courier'):
        if fontNm in availFonts:
            break
    else:
        fontNm = availFonts[4] if len(availFonts) > 4 else 'Monospace'

    applFont = QFont(fontNm, 9) if xOS else QFont(fontNm, 10)
    applFont.setFixedPitch(True)
    return applFont

###########################################################################
def selMenuAction(frMenu, actNm):
    for act in frMenu.actions():
        if act.text()==actNm:
            return act

###########################################################################
def setAlignment(widget, align):
    widget.setAlignment(AlignOpts[align.lower()])

###########################################################################
def setCaption(widget, caption):
    widget.setWindowTitle(caption)

###########################################################################
def setCheckState(widget, checkState):
    setState = Qt.Checked if checkState else Qt.Unchecked
    widget.setCheckState(setState)

###########################################################################
def setClour(widget, color):
    widget.setStyleSheet("background-color: rgb(color);")

GuiFont = {}
for size in (6, 7, 8, 9, 10, 11, 12, 14, 16):
    GuiFont[size] = QFont('Monaco', size, QFont.Bold)

###########################################################################
def setColumnCount(widget, columnCount):
    widget.setColumnCount(columnCount)

###########################################################################
def setColumnWidth(widget, colWidths):
    for colId, colWidth in enumerate(colWidths):
        widget.setColumnWidth(colId, colWidth)

###########################################################################
def setComboText(combo, txt):
    if txt:
        combo.setEditText(txt)
        combo.update()

###########################################################################
def setDefault(widget, default):
    widget.setDefault(default)

###########################################################################
def setDuplicatesEnabled(widget, enabled):
    widget.setDuplicatesEnabled(enabled)

###########################################################################
def setEditable(widget, editable):
    if isValid(widget, QLineEdit):
        setReadOnly(widget, not editable)
    else:
        widget.setEditable(editable)

###########################################################################
def setEnabled(widget, enable):
    widget.setEnabled(enable)

###########################################################################
def setFixedHeight(widget, fixedHeight):
    widget.setFixedHeight(fixedHeight)

###########################################################################
def setFixedSize(widget, size):
    widget.setFixedSize(QSize(size[0], size[1]))


###########################################################################
def setFixedWidth(widget, fixedWidth):
    widget.setFixedWidth(fixedWidth)

###########################################################################
def setFont(widget, font):
    widget.setFont(font)

###########################################################################
def setHeaderLabels(widget, headerLabels):
    widget.setHeaderLabels(headerLabels)

###########################################################################
def setIcon(widget, icon):
    widget.setWindowIcon(createIcon(icon))

###########################################################################
def setIconSize(widget, iconSize):
    widget.setIconSize(QSize(iconSize[0], iconSize[0]))

###########################################################################
def setLayout(widget, layout):
    widget.setLayout(layout)

###########################################################################
def setMargin(widget, margins):
    if isValid(margins, (list, tuple)):
        widget.setContentsMargins(margins[0],margins[1],margins[2],margins[3])
    elif isValid(margins, int):
        widget.setContentsMargins(margins, margins, margins, margins)

###########################################################################
def setMaxHeight(widget, maxHeight):
    widget.setMaximumHeight(maxHeight)

###########################################################################
def setMaxLength(widget, maxLength):
    widget.setMaxLength(maxLength)

###########################################################################
def setMaxWidth(widget, maxWidth):
    widget.setMaximumWidth(maxWidth)

###########################################################################
def setMaximumSize(widget, size):
    widget.setMaximumSize(QSize(size[0], size[1]))

###########################################################################
def setMinHeight(widget, minHeight):
    widget.setMinimumHeight(minHeight)

###########################################################################
def setMinWidth(widget, minWidth):
    widget.setMinimumWidth(minWidth)

###########################################################################
def setMinimumSize(widget, size):
    widget.setMinimumSize(QSize(size[0], size[1]))

###########################################################################
def setMsgColour(**opts):
    msgInd = getMsgInd(**opts)
    return COLOUR.get(msgInd, COLOUR[MSG_DFLT])

###########################################################################
def setMsgIcon(**opts):
    msgInd = getMsgInd(**opts)
    return IconDict.get(msgInd, IconDict[MSG_DFLT])

###########################################################################
def setMsgType(msg):
    srchRslt = MSG_RE.search(msg)
    if isNull(srchRslt): return 'Message', MSG_DFLT
    msgInd = srchRslt.lastindex
    return srchRslt.group(msgInd).upper(), msgInd


###########################################################################
def setObjectName(widget, objectName):
    widget.setObjectName(objectName)

###########################################################################
def setReadOnly(widget, readOnly):
    widget.setReadOnly(readOnly)

###########################################################################
def setResizeMode(widget, resizeMode):
    widget.setResizeMode(resizeMode)

###########################################################################
def setSelectionMode(widget, selectionMode='MultiSel'):
    mode = QAbstractItemView.SingleSelection if selectionMode=='SingleSel' else QAbstractItemView.ExtendedSelection
    widget.setSelectionMode(mode)

###########################################################################
def setSpacing(widget, spacing):
    widget.setSpacing(spacing)

###########################################################################
def setStatusTip(widget, statusTip):
    widget.setStatusTip(statusTip[0], statusTip[1])

###########################################################################
def setStyleSheet(widget, style):
    widget.setStyleSheet(style)

###########################################################################
def setText(widget, data):
    if isValid(data, (list, tuple)):
        return widget.setText(data[0], data[1])
    if isValid(widget, QComboBox):
        return setComboText(widget, data)
    widget.setText(data)

###########################################################################
def setToolTip(widget, toolTip):
    widget.setToolTip(toolTip)

###########################################################################
def setWidgetBgColor(widget, wType, color):
    if isValid(wType, str) and isValid(color, str):
        widget.setStyleSheet('%s { background-color: %s; }' % (wType, color))
    else:
        widget.setStyleSheet('')

###########################################################################
def setWidgetOpts(widget, args):
    for opt, val in list(args.items()):
        setOpt = opt.lower()
        if setOpt in WidgetOptSet:
            WidgetOptSet[setOpt](widget, val)
    return widget

###########################################################################
def setWordWrap(widget, wordWrap):
    widget.setWordWrap(wordWrap)

###########################################################################
def swapWidget(layout, frmWdgt, toWdgt, loc, stretch=0):
    layout.removeWidget(frmWdgt)
    layout.insertWidget(loc, toWdgt, stretch=stretch)
    frmWdgt.hide(); toWdgt.show()
    return toWdgt

###########################################################################
def translate(context, string):
    """Translation"""
    return QApplication.translate(context, string, None, QApplication.UnicodeUTF8)

###########################################################################
def updateColorWidget(widget, color):
    if isValid(color, str):
        widget.setStyleSheet('QWidget { background-color: %s; }' % color )
    else:
        widget.setStyleSheet('')

###############################################################################
''' Common Dialogs '''
##########################################################################
def SelectDir(parent, dfltDir=None, title='Open Directory'):
    ''' Select existing directory. '''
    dfltDir = dfltDir if dfltDir else os.curdir     
    return QFileDialog.getExistingDirectory(parent, title, dfltDir,
           QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks)
     
##########################################################################
def SelectFile(parent, fileDir, wildcard, title='Open file'):
    ''' Select existing file. '''
    return QFileDialog.getOpenFileName(parent, title, fileDir, wildcard)

##########################################################################
def SaveFile(parent, fileDir, wildcard, title='Save As', dfltFn=''):
    ''' Select or specify file to save to. '''
    fileSpec = joinPaths(fileDir, dfltFn) if fileDir else fileDir
    return QFileDialog.getSaveFileName(parent, title, fileSpec, wildcard)

##########################################################################
def SelColour(currColor=Qt.blue, parent=None):
    dlg = QColorDialog.getColor(currColor, parent)
    return dlg.name() if dlg.isValid() else None

##########################################################################
def SelFont(currFont=GuiFont[10], parent=None):
    font, dlgOk = QFontDialog.getFont(currFont, parent)
    return font if dlgOk else None

###############################################################################
class MyDialog(QDialog):

    ###########################################################################
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setWindowFlags(Qt.Dialog | Qt.WindowStaysOnTopHint | Qt.WindowTitleHint)

    ###########################################################################
    def closeEvent(self, event):
        event.ignore()

###############################################################################
###############################################################################
class CustCombo(QComboBox):
    ###########################################################################
    def __init__ (self, parent):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.InsertAlphabetically)

    ###########################################################################
    def getSpec(self):
        return self.currentText()

    ###########################################################################
    def loadSels (self, sels):
        loadSelLst(self, sels, selVal=0)

###############################################################################
###############################################################################
class DatePicker(MyDialog):

    ###########################################################################
    def __init__(self, callBack):
        super().__init__(None)
        
        self.callBack = callBack
        setCaption(self, 'Chọn Ngày')
        
        mainLayout = BoxLayout('V', spacing=10, margins=(2,2,2,2))
        self.setLayout(mainLayout)
        
        self.addCalendar()
        self.createDatesGroupBox()
        
        AddWidget(mainLayout, self.calendarGroupBox)
        AddWidget(mainLayout, self.datesGroupBox)
        
        self.saveBtn = Button(None, label='Apply')
        self.saveBtn.setStyleSheet("color: blue; background-color: yellow; font: bold 16px")
        self.saveBtn.clicked.connect(self.addApply)
        AddWidget(mainLayout, self.saveBtn, align='right')

    ###########################################################################
    def addApply(self):
        self.callBack(self.currentDateEdit.date())
        self.deleteLater()

    ###########################################################################
    def addCalendar(self):
        self.calendarGroupBox = QGroupBox("Lịch Tháng")
        self.calendar = QCalendarWidget()
        self.calendar.setMinimumDate(QDate(1900, 1, 1))
        self.calendar.setMaximumDate(QDate(3000, 1, 1))
        self.calendar.setGridVisible(True)
        self.calendar.selectionChanged.connect(self.onDateSel)
        self.previewLayout = QGridLayout()
        self.previewLayout.addWidget(self.calendar, 0, 0,
                Qt.AlignCenter)
        self.calendarGroupBox.setLayout(self.previewLayout)

    ###########################################################################
    def createDatesGroupBox(self):
        self.datesGroupBox = QGroupBox(self.tr("Dates"))
  
        self.currentDateEdit = QDateEdit()
        self.currentDateEdit.setDisplayFormat('MMM d yyyy')
        self.currentDateEdit.setDate(self.calendar.selectedDate())
  
        self.currentDateLabel = QLabel("&Current Date:")
        self.currentDateLabel.setBuddy(self.currentDateEdit)
  
        dateBoxLayout = QGridLayout()
        dateBoxLayout.addWidget(self.currentDateLabel, 1, 0)
        dateBoxLayout.addWidget(self.currentDateEdit, 1, 1)
        dateBoxLayout.setRowStretch(3, 1)
  
        self.datesGroupBox.setLayout(dateBoxLayout)            

    ###########################################################################
    def onDateSel(self):
        self.currentDateEdit.setDate(self.calendar.selectedDate())
            
###############################################################################
###############################################################################
class MsgDialog(MyDialog):

    ###########################################################################
    def __init__(self, caption, icon, msg, colour, diaType):
        super().__init__(None)
        setCaption(self, caption)
        self.setWindowIcon(createIcon('Parish'))

        self.mainLayout = BoxLayout('V', spacing=2, margins=25)
        msgLo = BoxLayout('H', spacing=25, margins=2)

        image = createIcon(icon, pixmap=True)   #.scaled(48,48)
        label = Label(); label.setPixmap(image)
        setMaxWidth(label, 72)
        AddWidget(msgLo, label)

        myMsg = CustLabel(msg, size=14, colour=colour)
        myMsg.setWordWrap(True)
        AddWidget(msgLo, myMsg)
        AddLayout(self.mainLayout, msgLo)

        setMinWidth(self, 725)
        self.setLayout(self.mainLayout)

        if diaType==PROMPT:
            buttonBox = QDialogButtonBox(QDialogButtonBox.Ok)
            AddWidget(self.mainLayout, buttonBox)
        elif diaType==CONFIRM:
            buttonBox = QDialogButtonBox(QDialogButtonBox.Yes | QDialogButtonBox.No)
            AddWidget(self.mainLayout, buttonBox)
        elif diaType==SAVECONF:
            buttonBox = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Discard | QDialogButtonBox.Cancel)
            AddWidget(self.mainLayout, buttonBox)
        else:
            buttonBox = None
            
        if buttonBox is not None:
            buttonBox.clicked.connect(self.onClick)
        self.exec_()

    ###########################################################################
    def _getUserAns(self, isYesNo=None):
        self.close()
        self.deleteLater()
        return self.rslt==QDialogButtonBox.Yes if isYesNo==CONFIRM else self.rslt

    ###########################################################################
    def onClick(self, btn):
        userSel = btn.text().replace('&', '')
        self.rslt = getattr(QDialogButtonBox, 'Ok' if userSel=='OK' else userSel)
        self.done(self.rslt)

###############################################################################
###############################################################################
class SelectUI(MyDialog):

    ###########################################################################
    def __init__(self, parent, caption, selLst, tips=None, multSel=False):
        super().__init__(parent)
        setCaption(self, caption)

        self.multSel = multSel
        selLst = list(selLst)
        selLst.sort()
        self.dialogLo = BoxLayout('V', spacing=10, margins=6)
        self.setLayout(self.dialogLo)

        self.dialogLo.addWidget(CustLabel(caption, 16, maxheight=45))

        if isNull(tips):
            tips = 'Click on item to select...'
        grpWidget, self.itemSel = ListGroup('', mSels=multSel, fixedHeight=175)
        self.dialogLo.addWidget(grpWidget)
        self.loadSels(selLst, tips)

        actLo = BoxLayout('H', spacing=10, margins=6)
        cancleBtn = Button(label='CANCEL')
        selBtn = Button(label='APPLY')
        AddWidget(actLo, selBtn)
        AddWidget(actLo, cancleBtn)
        AddWidget(self.dialogLo, GroupBox('', actLo))
        
        cancleBtn.clicked.connect(self.reject)
        selBtn.clicked.connect(self.accept)

    ###########################################################################
    def getSel(self):
        self.show()
        selItem = self.getSpec() if self.exec_() else None
        self.deleteLater()
        return selItem

    ###########################################################################
    def getSpec(self):
        if self.multSel:
            return [selItem.text() for selItem in self.itemSel.selectedItems()]
        return self.itemSel.currentItem().text()

    ###########################################################################
    def loadSels(self, selLst, tip):
        for item in selLst:
            self.itemSel.addItem( ListItem(item, tooltip=tip) )

###############################################################################
###############################################################################
class TimeSelector(QTimeEdit):

    ###########################################################################
    def __init__(self, minInc):
        super().__init__(None)
        self.minInc = minInc
        self.setDisplayFormat('hh:mm AP')
        self.setDate(datetime.now())
        
    ###########################################################################
    def stepBy(self, steps):
        if self.currentSection() == self.MinuteSection:
            QTimeEdit.stepBy(self, steps * self.minInc)
        else:
            QTimeEdit.stepBy(self, steps)
            
###############################################################################
###############################################################################
class UserInputUI(MyDialog):

    ###########################################################################
    def __init__(self, parent, label, caption):
        super().__init__(parent)
        setCaption(self, caption)

        self.dialogLo = BoxLayout('V', spacing=10, margins=6)
        self.setLayout(self.dialogLo)
        self.dialogLo.addWidget(CustLabel(caption, 16, maxheight=45))

        layout = BoxLayout('H')
        lbl, self.txtEdit = TextWidget(label, labelWidth=125, layout=layout)
        AddLayout(self.dialogLo, layout)

        actLo = BoxLayout('H', spacing=10, margins=6)
        doneBtn = Button(label='DONE')
        selBtn = Button(label='APPLY')
        AddWidget(actLo, selBtn)
        AddWidget(actLo, doneBtn)
        AddWidget(self.dialogLo, GroupBox('', actLo))

        doneBtn.clicked.connect(self.reject)
        selBtn.clicked.connect(self.accept)

        self.exec_()

    ###########################################################################
    def getUserInput(self):
        return self.txtEdit.text()



##############################################################################
###########################################################################
def VNTextEdit(parent=None, **args):
    textEdit = VNTextBox(parent=parent)
    return setWidgetOpts(textEdit, args)

##############################################################################
class VNTextBox(QTextEdit):

    def __init__(self, parent=None):
        super().__init__(parent)

        self.cursorPositionChanged.connect(self._cursor_position_changed)
        self.currLineNo = 0

    def insertCompletion(self, completion):
        if self._completer.widget() is not self:
            return

        tc = self.textCursor()
        extra = len(completion) - len(self._completer.completionPrefix())
        tc.movePosition(QTextCursor.Left)
        tc.movePosition(QTextCursor.EndOfWord)
        tc.insertText(completion[-extra:])
        self.setTextCursor(tc)

    def textUnderCursor(self, curs):
        curs.select(QTextCursor.WordUnderCursor)
        return curs.selectedText()

    def focusInEvent(self, e):

        super().focusInEvent(e)

    def keyPressEvent(self, e):
        global TypeRules

        if ((e.modifiers() & Qt.ControlModifier) and e.key()==Qt.Key_E): return

        super().keyPressEvent(e)

        curs = self.textCursor()
        completionPrefix = self.textUnderCursor(curs)

        if len(completionPrefix)>1 and e.text() in TypeRules:
            vnTxt = bogo.process_sequence(completionPrefix, rules=TypeRules)
            if vnTxt!=completionPrefix:
                curs.insertText(vnTxt)

    def _cursor_position_changed(self):
        lineNo = self.textCursor().blockNumber() + 1
        if self.currLineNo != lineNo:
            self.currLineNo = lineNo
            self.onCurrentLine()

    def onCurrentLine(self):
        lineColor = QColor(Qt.red).lighter(190)
        highlight = QTextEdit.ExtraSelection()
        highlight.cursor = self.textCursor()
        highlight.format.setProperty(QTextFormat.FullWidthSelection, True)
        highlight.format.setBackground(lineColor)

        extras = []
        extras.append(highlight)
        self.setExtraSelections( extras ) 


##############################################################################
###########################################################################
def VNLineEdit(parent=None, **args):
    textEdit = VNTextline(parent=parent)
    return setWidgetOpts(textEdit, args)

##############################################################################
class VNTextline(QLineEdit):

    def __init__(self, parent, separator = ' '):
        super().__init__(parent=parent)
        self._separator = separator

    def keyPressEvent(self, e):
        global TypeRules

        if ((e.modifiers() & Qt.ControlModifier) and e.key()==Qt.Key_E): return

        super().keyPressEvent(e)

        ind, completionPrefix = self.textUnderCursor()
        
        if len(completionPrefix)>1 and e.text() in TypeRules:
            vnTxt = bogo.process_sequence(completionPrefix, rules=TypeRules)
            if vnTxt!=completionPrefix:
                self.setVnText(vnTxt, ind, len(completionPrefix))
                
    def setVnText(self, vnTxt, ind, wrdLen):
        self.setSelection(ind, wrdLen)
        self.cut()
        self.insert(vnTxt)

    def textUnderCursor(self):
        text = self.text()
        textUnderCursor = ''
        i = self.cursorPosition() - 1
        while i >=0 and text[i] != self._separator:
            textUnderCursor = text[i] + textUnderCursor
            i -= 1
        return i+1, textUnderCursor


# All parameter keywords must be lower case in WidgetOptSet
WidgetOptSet = {'maxwidth': setMaxWidth,
                'minwidth': setMinWidth,
                'fixedwidth': setFixedWidth,
                'maxheight': setMaxHeight,
                'minheight': setMinHeight,
                'fixedheight': setFixedHeight,
                'spacing': setSpacing,
                'margins': setMargin,
                'maxlength': setMaxLength,
                'style': setStyleSheet,
                'objectname': setObjectName,
                'layout': setLayout,
                'iconsize': setIconSize,
                'resizemode': setResizeMode,
                'text': setText,
                'tooltip': setToolTip,
                'selectionmode': setSelectionMode,
                'statustip': setStatusTip,
                'font': setFont,
                'readonly': setReadOnly,
                'align': setAlignment,
                'size':setFixedSize,
                'editable':setEditable,
                'enable':setEnabled,
                'headerlabels':setHeaderLabels,
                'default':setDefault,
                'color':setClour,
                'columncount':setColumnCount,
                'colwidths':setColumnWidth,
                'wordwrap':setWordWrap,
                'checkstate':setCheckState,
                'additems':addItems,
                'minimumsize':setMinimumSize,
                'maximumsize':setMaximumSize,
                'fixedsize':setFixedSize,
                'duplicatesenabled':setDuplicatesEnabled,
               }

AlignOpts = {'center': Qt.AlignCenter,
             'hcenter': Qt.AlignHCenter,
             'vcenter': Qt.AlignVCenter,
             'right': Qt.AlignRight,
             'left': Qt.AlignLeft,
             'top': Qt.AlignTop,
             'bottom': Qt.AlignBottom,
            }

###############################################################################
#                                End of module                                #
###############################################################################
