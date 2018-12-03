import sys
import re

from Utils.PyUtils import *

STAT_NORM   = 1
STAT_WARN   = 2
STAT_ERROR  = 3

class StatusLexer(Qsci.QsciLexerCustom):

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
        self.setDefaultColor(QtGui.QColor(self._trmLexer['Colour'][dfltStyle]))
        self.setDefaultFont(self._trmLexer['Font'][dfltStyle])
        self.setDefaultPaper(QtGui.QColor(self._trmLexer['Colour'][-11]))

    def description(self, style):
        return self._styles.get(style, '')

    def defaultColor(self, style):
        if style in self._trmLexer['Colour']:
            return self._trmLexer['Colour'][style]
        return Qsci.QsciLexerCustom.defaultColor(self, style)

    def defaultFont(self, style):
        if style in self._trmLexer['Font']:
            return self._trmLexer['Font'][style]
        return Qsci.QsciLexerCustom.defaultFont(self, style)

    def defaultEolFill(self, style):
        # This allowed to colorize all the background of a line.
        if style in self._trmLexer['Font']:
            return True
        return Qsci.QsciLexerCustom.defaultEolFill(self, style)


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
# End of File
##############################################################################
