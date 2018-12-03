"""
******************************************************************************
 NAME: StTrm.patusy
******************************************************************************
"""
import os
import sys
from time import sleep as wait
from multiprocessing.queues import SimpleQueue

##############################################################################
##############################################################################
from Utils.QtUtils import *
from .AppShared import *
from .MyThread import MyThread
from .MyQtEvent import emitMyEvt

from PyQt5.Qsci import QsciScintilla, QsciLexerCustom, QsciLexerHTML

STAT_NORM   = 1
STAT_WARN   = 2
STAT_ERROR  = 3


class StatusLexer(QsciLexerCustom):

    def __init__(self, parent):
        super().__init__(parent)
        self._trmLexer = {}
        self._styles = {}

        self.procStat = STAT_NORM
        self.rejCmds = 0

    def setProcStat(self, stat=STAT_NORM):
        self.procStat = stat

    def cfgLexer(self, trmLexer):
        self._trmLexer = trmLexer
        for styleNm, style in list(self._trmLexer['TX'].items()):

            self._styles[style] = styleNm
            setattr(self, styleNm, style)

        dfltStyle = 1
        self.setDefaultColor(QColor(self._trmLexer['Colour'][dfltStyle]))
        self.setDefaultFont(self._trmLexer['Font'][dfltStyle])
        self.setDefaultPaper(QColor(self._trmLexer['Colour'][-11]))

    def description(self, style):
        return self._styles.get(style, '')

    def defaultColor(self, style):
        if style in self._trmLexer['Colour']:
            return self._trmLexer['Colour'][style]
        return QsciLexerCustom.defaultColor(self, style)

    def defaultFont(self, style):
        if style in self._trmLexer['Font']:
            return self._trmLexer['Font'][style]
        return QsciLexerCustom.defaultFont(self, style)

    def defaultEolFill(self, style):
        # This allowed to colorize all the background of a line.
        if style in self._trmLexer['Font']:
            return True
        return QsciLexerCustom.defaultEolFill(self, style)


    def styleText(self, start, end):
        editor = self.editor()
        if isNull(editor): return
        textLines = self.getTextLines(editor, start, end)

        # the line index will also be needed to implement folding
        index = editor.SendScintilla(editor.SCI_LINEFROMPOSITION, start)
        if index > 0:
            # the previous state may be needed for multi-line styling
            pos = editor.SendScintilla(editor.SCI_GETLINEENDPOSITION, index - 1)
            state = editor.SendScintilla(editor.SCI_GETSTYLEAT, pos)
        else:
            state = self.Foreground

        #set_style = self.setStyling
        self.startStyling(start, 0x1f)

        # scintilla always asks to style whole lines
        for line in textLines:
            length = len(line)
            levelFolder = editor.SendScintilla(editor.SCI_GETFOLDLEVEL, index)
            procStat = STAT_NORM
            line = str(line)
            if 'FATAL:' in line or 'Traceback' in line:
                procStat = STAT_ERROR
                state = self.Fatal
            elif 'ERROR:' in line or 'error(s)' in line:
                procStat = STAT_ERROR
                state = self.Error
            elif 'INTERNAL:' in line or 'Internal' in line:
                state = self.Info
            elif 'WARNING:' in line or 'warning(s)' in line:
                procStat = STAT_WARN
                state = self.Warning
            elif 'STATUS:' in line:
                state = self.Info
            elif 'installed' in line:
                state = self.Installed
            elif 'successfully' in line:
                state = self.Success
            else:
                state = self.Foreground

            self.setProcStat(procStat)
            self.setStyling(length, state)
            index += 1

    def getTextLines(self, editor, start, end):
        if isNull(editor): return []

        source = ''
        if end > editor.length():
            end = editor.length()
        if end > start:
            if sys.hexversion >= 0x02060000:
                # faster when styling big files, but needs python 2.6
                source = bytearray(end - start)
                editor.SendScintilla(
                    editor.SCI_GETTEXTRANGE, start, end, source)
            else:
                source = str(editor.text()).encode('utf-8')[start:end]
        return source.splitlines(True) if source else []


##############################################################################
class Editor(QsciScintilla):
    ARROW_MARKER_NUM = 8

    ##########################################################################
    def __init__(self):
        super().__init__(None)

        self.setWrapMode(QsciScintilla.WrapWord)

        self.editable = True

        self.setLexer(QsciLexerHTML())

    ##########################################################################
    def enableMargin(self):
        # Set the default font
        font = QFont()
        font.setFamily('Monaco')
        font.setFixedPitch(True)
        font.setPointSize(10)
        self.setFont(font)
        self.setMarginsFont(font)
        # Margin 0 is used for line numbers
        fontmetrics = QFontMetrics(font)
        self.setMarginsFont(font)
        self.setMarginWidth(0, fontmetrics.width("00000") + 6)
        self.setMarginLineNumbers(0, True)
        self.setMarginsBackgroundColor(QColor("#cccccc"))

        # Clickable margin 1 for showing markers
        self.setMarginSensitivity(1, True)
        self.connect(self,
            SIGNAL('marginClicked(int, int, Qt::KeyboardModifiers)'),
            self.on_margin_clicked)
        self.markerDefine(QsciScintilla.RightArrow,
            self.ARROW_MARKER_NUM)
        self.setMarkerBackgroundColor(QColor("#ee1111"),
            self.ARROW_MARKER_NUM)

    ##########################################################################
    def write(self, data):
        self.append(data)

    ##########################################################################
    def setViewer(self):
        self.editable = False
        self.setReadOnly(True)

##############################################################################
def cfgLexer():
    fontNm = 'Courier New'
    lexerDct = {1:('Foreground', ((fontNm, 10, 50), '#000000')),
                2:('Info', ((fontNm, 12, 75), 'blue')),
                3:('Warning', ((fontNm, 12, 75), 'darkMagenta')),
                4:('Error', ((fontNm, 12, 75), 'red')),
                5:('Fatal', ((fontNm, 12, 75), 'red')),
                6:('Sorry', ((fontNm, 12, 75), 'red')),
                7:('Success', ((fontNm, 12, 75), 'blue')),
                8:('Internal', ((fontNm, 12, 75), 'darkMagenta')),
                9:('Installed', ((fontNm, 12, 75), 'blue')),
               }

    lexer = {'Colour':{}, 'Font':{}, 'TX':{}}
    for style, (styleNm, (lexerFont, lexerColour)) in list(lexerDct.items()):
        lexer['Font'][style] = QFont(lexerFont[0], pointSize=lexerFont[1], weight=lexerFont[2])
        lexer['Colour'][style] = QColor(lexerColour)
        lexer['TX'][styleNm] = style
    lexer['Colour'][-11] = QColor('#ffffeb')

    return lexer


##########################################################################
def startEditor():
    editor = Editor()
    editor.setUtf8(True)
    editor.setMinimumWidth(700)

    trmLexer = cfgLexer()
    lexer = StatusLexer(editor)
    lexer.cfgLexer(trmLexer)
    editor.setLexer(lexer)
    
    return editor


##############################################################################
# End of File
##############################################################################
