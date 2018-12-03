import sys

from Utils.QtUtils import *
from Utils.PyUtils import *

SEL_DIR, SEL_FILE, SAVE_FILE, SEL_FILES = range(1,5)
FileMode = {SEL_DIR:    QFileDialog.DirectoryOnly,
            SAVE_FILE:  QFileDialog.AnyFile,
            SEL_FILE:   QFileDialog.ExistingFile,
            SEL_FILES:  QFileDialog.ExistingFiles,}

##############################################################################
##############################################################################
class DiaFileModel(QSortFilterProxyModel):                                                                                                                                  
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDynamicSortFilter(True)
        self.setSortCaseSensitivity(Qt.CaseInsensitive)


    def filterAcceptsRow(self, source_row, source_parent):
        source_model = self.sourceModel()
        source_index = source_model.index(source_row, 0, source_parent)
        
        fileSpec = source_model.filePath(source_index)
        return isDir(fileSpec) or isScenFile(fileSpec)

        #print('source_index: ', dir(source_parent))
        #print('source_index: ', dir(source_index.data())
        #print('\nsource_index data: ', source_index.data())
        #print('source_index filePath: ', source_index.model().filePath(source_index))
        #print('source_index mimeTypes: ', source_index.model().mimeTypes())
        #print('source_index mimeData: ', source_index.model().mimeData([source_index]))
        #print('account_info: ', dir(account_info)) isDir mimeData  mimeTypes 
            
        '''
        source_index:  ['__class__', '__delattr__', '__dict__', '__dir__', '__doc__', '__eq__', '__format__', 
        '__ge__', '__getattribute__', '__gt__', '__hash__', '__init__', '__le__', '__lt__', '__module__', '__ne__', 
        '__new__', '__reduce__', '__reduce_ex__', '__repr__', '__setattr__', '__sizeof__', '__str__', 
        '__subclasshook__', '__weakref__', 'child', 'column', 'data', 'flags', 'internalId', 'internalPointer', 
        'isValid', 'model', 'parent', 'row', 'sibling']

        '''

##############################################################################
class FileDialog(QDialog):
    ##########################################################################
    def __init__(self, parent, caption, directory=None, filter="Any File (*.*)", mode=SEL_FILE):
        super().__init__(parent)
        self.setWindowTitle('SeaFood ...')
        #self.setWindowIcon(createIcon('Seagull'))
        self.selMode = mode

        mainLayout = BoxLayout('V', spacing=2, margin=4)
        AddWidget(mainLayout, CustLabel(caption, minHeight=21, size=28))

        self.dlg = QFileDialog(None, None, directory=directory, filter=filter)
        self.dlg.setViewMode(QFileDialog.Detail)
        self.dlg.setFileMode(FileMode[mode])
        if mode==SAVE_FILE:
            self.dlg.setAcceptMode(QFileDialog.AcceptSave)
            self.dlg.setConfirmOverwrite(True)

        self.listView = self.dlg.findChild(QListView, 'listView')
        self.listView.setAlternatingRowColors(True)
        self.listView.setModelColumn(0)

        '''
        sizeHintForRow:  ['AboveItem', 'AllEditTriggers', 'AnimatingState', 'AnyKeyPressed', 
        'BelowItem', 'Box', 'CollapsingState', 'ContiguousSelection', 'CurrentChanged', 
        'CursorAction', 'Custom', 'DoubleClicked', 'DragDrop', 'DragDropMode', 'DragOnly', 
        'DragSelectingState', 'DraggingState', 'DrawChildren', 'DrawWindowBackground', 
        'DropIndicatorPosition', 'DropOnly', 'EditKeyPressed', 'EditTrigger', 'EditTriggers', 
        'EditingState', 'EnsureVisible', 'ExpandingState', 'ExtendedSelection', 'Fixed', 'HLine', 
        'IgnoreMask', 'Interactive', 'InternalMove', 'MoveDown', 'MoveEnd', 'MoveHome', 'MoveLeft', 
        'MoveNext', 'MovePageDown', 'MovePageUp', 'MovePrevious', 'MoveRight', 'MoveUp', 
        'MultiSelection', 'NoDragDrop', 'NoEditTriggers', 'NoFrame', 'NoSelection', 'NoState', 
        'OnItem', 'OnViewport', 'PaintDeviceMetric', 'Panel', 'PdmDepth', 'PdmDpiX', 'PdmDpiY', 
        'PdmHeight', 'PdmHeightMM', 'PdmNumColors', 'PdmPhysicalDpiX', 'PdmPhysicalDpiY', 'PdmWidth', 
        'PdmWidthMM', 'Plain', 'PositionAtBottom', 'PositionAtCenter', 'PositionAtTop', 'Raised', 
        'RenderFlag', 'RenderFlags', 'ResizeMode', 'ResizeToContents', 'ScrollHint', 'ScrollMode', 
        'ScrollPerItem', 'ScrollPerPixel', 'SelectColumns', 'SelectItems', 'SelectRows', 
        'SelectedClicked', 'SelectionBehavior', 'SelectionMode', 'Shadow', 'Shadow_Mask', 'Shape', 
        'Shape_Mask', 'SingleSelection', 'State', 'Stretch', 'StyleMask', 'StyledPanel', 'Sunken', 
        'VLine', 'WinPanel', '__class__', '__delattr__', '__dict__', '__dir__', '__doc__', '__eq__', 
        '__format__', '__ge__', '__getattr__', '__getattribute__', '__gt__', '__hash__', '__init__', 
        '__le__', '__len__', '__lt__', '__module__', '__ne__', '__new__', '__reduce__', 
        '__reduce_ex__', '__repr__', '__setattr__', '__sizeof__', '__str__', '__subclasshook__', 
        '__weakref__', 'acceptDrops', 'accessibleDescription', 'accessibleName', 'actionEvent', 
        'actions', 'activateWindow', 'activated', 'addAction', 'addActions', 'addScrollBarWidget', 
        'adjustSize', 'alternatingRowColors', 'autoFillBackground', 'autoScrollMargin', 
        'backgroundRole', 'baseSize', 'blockSignals', 'cascadingSectionResizes', 'changeEvent', 
        'childAt', 'childEvent', 'children', 'childrenRect', 'childrenRegion', 'clearFocus', 
        'clearMask', 'clearSelection', 'clicked', 'close', 'closeEditor', 'closeEvent', 
        'closePersistentEditor', 'colorCount', 'commitData', 'connect', 'connectNotify', 
        'contentsMargins', 'contentsRect', 'contextMenuEvent', 'contextMenuPolicy', 'cornerWidget', 
        'count', 'create', 'currentChanged', 'currentIndex', 'cursor', 'customContextMenuRequested', 
        'customEvent', 'dataChanged', 'defaultAlignment', 'defaultDropAction', 'defaultSectionSize', 
        'deleteLater', 'depth', 'destroy', 'destroyed', 'devType', 'dirtyRegionOffset', 'disconnect', 
        'disconnectNotify', 'doubleClicked', 'dragDropMode', 'dragDropOverwriteMode', 'dragEnabled', 
        'dragEnterEvent', 'dragLeaveEvent', 'dragMoveEvent', 'drawFrame', 'dropEvent', 
        'dropIndicatorPosition', 'dumpObjectInfo', 'dumpObjectTree', 'dynamicPropertyNames', 'edit', 
        'editTriggers', 'editorDestroyed', 'effectiveWinId', 'emit', 'enabledChange', 
        'ensurePolished', 'enterEvent', 'entered', 'event', 'eventFilter', 'executeDelayedItemsLayout', 
        'find', 'findChild', 'findChildren', 'focusInEvent', 'focusNextChild', 'focusNextPrevChild', 
        'focusOutEvent', 'focusPolicy', 'focusPreviousChild', 'focusProxy', 'focusWidget', 'font', 
        'fontChange', 'fontInfo', 'fontMetrics', 'foregroundRole', 'frameGeometry', 'frameRect', 
        'frameShadow', 'frameShape', 'frameSize', 'frameStyle', 'frameWidth', 'geometriesChanged', 
        'geometry', 'getContentsMargins', 'grabGesture', 'grabKeyboard', 'grabMouse', 'grabShortcut', 
        'graphicsEffect', 'graphicsProxyWidget', 'hasAutoScroll', 'hasFocus', 'hasMouseTracking', 
        'headerDataChanged', 'height', 'heightForWidth', 'heightMM', 'hiddenSectionCount', 'hide', 
        'hideEvent', 'hideSection', 'highlightSections', 'horizontalOffset', 'horizontalScrollBar', 
        'horizontalScrollBarPolicy', 'horizontalScrollMode', 'horizontalScrollbarAction', 
        'horizontalScrollbarValueChanged', 'horizontalStepsPerItem', 'iconSize', 'indexAt', 'indexWidget', 
        'inherits', 'initStyleOption', 'initialize', 'initializeSections', 'inputContext', 
        'inputMethodEvent', 'inputMethodHints', 'inputMethodQuery', 'insertAction', 'insertActions', 
        'installEventFilter', 'isActiveWindow', 'isAncestorOf', 'isClickable', 'isEnabled', 'isEnabledTo', 
        'isEnabledToTLW', 'isFullScreen', 'isHidden', 'isIndexHidden', 'isLeftToRight', 'isMaximized', 
        'isMinimized', 'isModal', 'isMovable', 'isRightToLeft', 'isSectionHidden', 'isSortIndicatorShown', 
        'isTopLevel', 'isVisible', 'isVisibleTo', 'isWidgetType', 'isWindow', 'isWindowModified', 
        'itemDelegate', 'itemDelegateForColumn', 'itemDelegateForRow', 'keyPressEvent', 'keyReleaseEvent', 
        'keyboardGrabber', 'keyboardSearch', 'killTimer', 'languageChange', 'layout', 'layoutDirection', 
        'leaveEvent', 'length', 'lineWidth', 'locale', 'logicalDpiX', 'logicalDpiY', 'logicalIndex', 
        'logicalIndexAt', 'lower', 'mapFrom', 'mapFromGlobal', 'mapFromParent', 'mapTo', 'mapToGlobal', 
        'mapToParent', 'mask', 'maximumHeight', 'maximumSize', 'maximumViewportSize', 'maximumWidth', 
        'metaObject', 'metric', 'midLineWidth', 'minimumHeight', 'minimumSectionSize', 'minimumSize', 
        'minimumSizeHint', 'minimumWidth', 'model', 'mouseDoubleClickEvent', 'mouseGrabber', 
        'mouseMoveEvent', 'mousePressEvent', 'mouseReleaseEvent', 'move', 'moveCursor', 'moveEvent', 
        'moveSection', 'moveToThread', 'nativeParentWidget', 'nextInFocusChain', 'normalGeometry', 
        'numColors', 'objectName', 'offset', 'openPersistentEditor', 'orientation', 'overrideWindowFlags', 
        'overrideWindowState', 'paintEngine', 'paintEvent', 'paintSection', 'paintingActive', 'palette', 
        'paletteChange', 'parent', 'parentWidget', 'physicalDpiX', 'physicalDpiY', 'pos', 'pressed', 
        'previousInFocusChain', 'property', 'pyqtConfigure', 'raise_', 'receivers', 'rect', 'releaseKeyboard',
         'releaseMouse', 'releaseShortcut', 'removeAction', 'removeEventFilter', 'render', 'repaint', 
         'reset', 'resetInputContext', 'resize', 'resizeEvent', 'resizeMode', 'resizeSection', 
         'resizeSections', 'restoreGeometry', 'restoreState', 'rootIndex', 'rowsAboutToBeRemoved', 
         'rowsInserted', 'saveGeometry', 'saveState', 'scheduleDelayedItemsLayout', 'scroll', 
         'scrollBarWidgets', 'scrollContentsBy', 'scrollDirtyRegion', 'scrollTo', 'scrollToBottom', 
         'scrollToTop', 'sectionAutoResize', 'sectionClicked', 'sectionCountChanged', 'sectionDoubleClicked', 
         'sectionEntered', 'sectionHandleDoubleClicked', 'sectionMoved', 'sectionPosition', 'sectionPressed', 
         'sectionResized', 'sectionSize', 'sectionSizeFromContents', 'sectionSizeHint', 
         'sectionViewportPosition', 'sectionsAboutToBeRemoved', 'sectionsHidden', 'sectionsInserted', 
         'sectionsMoved', 'selectAll', 'selectedIndexes', 'selectionBehavior', 'selectionChanged', 
         'selectionCommand', 'selectionMode', 'selectionModel', 'sender', 'senderSignalIndex', 
         'setAcceptDrops', 'setAccessibleDescription', 'setAccessibleName', 'setAlternatingRowColors', 
         'setAttribute', 'setAutoFillBackground', 'setAutoScroll', 'setAutoScrollMargin', 
         'setBackgroundRole', 'setBaseSize', 'setCascadingSectionResizes', 'setClickable', 
         'setContentsMargins', 'setContextMenuPolicy', 'setCornerWidget', 'setCurrentIndex', 'setCursor', 
         'setDefaultAlignment', 'setDefaultDropAction', 'setDefaultSectionSize', 'setDirtyRegion', 
         'setDisabled', 'setDragDropMode', 'setDragDropOverwriteMode', 'setDragEnabled', 
         'setDropIndicatorShown', 'setEditTriggers', 'setEnabled', 'setFixedHeight', 'setFixedSize', 
         'setFixedWidth', 'setFocus', 'setFocusPolicy', 'setFocusProxy', 'setFont', 'setForegroundRole', 
         'setFrameRect', 'setFrameShadow', 'setFrameShape', 'setFrameStyle', 'setGeometry', 
         'setGraphicsEffect', 'setHidden', 'setHighlightSections', 'setHorizontalScrollBar', 
         'setHorizontalScrollBarPolicy', 'setHorizontalScrollMode', 'setHorizontalStepsPerItem', 
         'setIconSize', 'setIndexWidget', 'setInputContext', 'setInputMethodHints', 'setItemDelegate', 
         'setItemDelegateForColumn', 'setItemDelegateForRow', 'setLayout', 'setLayoutDirection', 
         'setLineWidth', 'setLocale', 'setMask', 'setMaximumHeight', 'setMaximumSize', 'setMaximumWidth', 
         'setMidLineWidth', 'setMinimumHeight', 'setMinimumSectionSize', 'setMinimumSize', 'setMinimumWidth', 
         'setModel', 'setMouseTracking', 'setMovable', 'setObjectName', 'setOffset', 'setOffsetToLastSection', 
         'setOffsetToSectionPosition', 'setPalette', 'setParent', 'setProperty', 'setResizeMode', 'setRootIndex', 
         'setSectionHidden', 'setSelection', 'setSelectionBehavior', 'setSelectionMode', 'setSelectionModel', 
         'setShortcutAutoRepeat', 'setShortcutEnabled', 'setShown', 'setSizeIncrement', 'setSizePolicy', 
         'setSortIndicator', 'setSortIndicatorShown', 'setState', 'setStatusTip', 'setStretchLastSection', 
         'setStyle', 'setStyleSheet', 'setTabKeyNavigation', 'setTabOrder', 'setTextElideMode', 'setToolTip', 
         'setUpdatesEnabled', 'setVerticalScrollBar', 'setVerticalScrollBarPolicy', 'setVerticalScrollMode', 
         'setVerticalStepsPerItem', 'setViewport', 'setViewportMargins', 'setVisible', 'setWhatsThis', 
         'setWindowFilePath', 'setWindowFlags', 'setWindowIcon', 'setWindowIconText', 'setWindowModality', 
         'setWindowModified', 'setWindowOpacity', 'setWindowRole', 'setWindowState', 'setWindowTitle', 
         'setupViewport', 'show', 'showDropIndicator', 'showEvent', 'showFullScreen', 'showMaximized', 
         'showMinimized', 'showNormal', 'showSection', 'signalsBlocked', 'size', 'sizeHint', 'sizeHintForColumn', 
         'sizeHintForIndex', 'sizeHintForRow', 'sizeIncrement', 'sizePolicy', 'sortIndicatorChanged', 
         'sortIndicatorOrder', 'sortIndicatorSection', 'stackUnder', 'startDrag', 'startTimer', 'state', 
         'staticMetaObject', 'statusTip', 'stretchLastSection', 'stretchSectionCount', 'style', 'styleSheet', 
         'swapSections', 'tabKeyNavigation', 'tabletEvent', 'testAttribute', 'textElideMode', 'thread', 
         'timerEvent', 'toolTip', 'topLevelWidget', 'tr', 'trUtf8', 'underMouse', 'ungrabGesture', 
         'unsetCursor', 'unsetLayoutDirection', 'unsetLocale', 'update', 'updateEditorData', 
         'updateEditorGeometries', 'updateGeometries', 'updateGeometry', 'updateMicroFocus', 'updateSection', 
         'updatesEnabled', 'verticalOffset', 'verticalScrollBar', 'verticalScrollBarPolicy', 'verticalScrollMode', 
         'verticalScrollbarAction', 'verticalScrollbarValueChanged', 'verticalStepsPerItem', 'viewOptions', 
         'viewport', 'viewportEntered', 'viewportEvent', 'visibleRegion', 'visualIndex', 'visualIndexAt', 
         'visualRect', 'visualRegionForSelection', 'whatsThis', 'wheelEvent', 'width', 'widthMM', 'winEvent', 
         'winId', 'window', 'windowActivationChange', 'windowFilePath', 'windowFlags', 'windowIcon', 
         'windowIconText', 'windowModality', 'windowOpacity', 'windowRole', 'windowState', 'windowTitle', 
         'windowType', 'x', 'y']

        '''

        self.headerView = self.dlg.findChild(QHeaderView)
        #self.headerView.setSizePolicy(275)
        
        print('sizeHintForRow: ', dir(self.headerView))
        #self.headerView.setResizeMode(QHeaderView.Stretch)
        self.headerView.setResizeMode(QHeaderView.ResizeToContents)
        #self.headerView.setResizeMode(QHeaderView.Custom)
        
        #self.listView.setColumnWidthMode(0,QListView.Maximum)
        #print (dir(self.listView))

        proxyModel = DiaFileModel()
        self.dlg.setProxyModel(proxyModel)

        AddWidget(mainLayout, self.dlg)
        self.setLayout(mainLayout)

        self.show()
        
        #print('QTreeView: ', self.dlg.findChild(QTreeView))
        
    def sizeHint__(self):
        s = QSize()
        s.setHeight(self.listView.sizeHint().height())
        s.setWidth(275)
        return s        

    def exec_(self):
        if self.dlg.exec_()==QDialog.Accepted:
            #self.getList()
            self.close()
            return QDialog.Accepted
        return QDialog.Rejected

    def selectedFiles(self):
        return self.dlg.selectedFiles()


    def getList(self):
        
        fileTree = self.dlg.findChild(QListView)
        
        print(fileTree, fileTree.selectedIndexes())

        for i in fileTree.selectedIndexes():
            print(os.path.join(str(self.dlg.directory().absolutePath()),str(i.data())))
        
            
##########################################################################
''' Common Dialogs '''
##########################################################################
##########################################################################
def SelectDir(parent, dfltDir=None, title='Open Directory'):
    ''' Select existing directory. '''
    if dfltDir is None: dfltDir = QDir.currentPath()
    return getUseSel(FileDialog(parent, title, dfltDir, mode=SEL_DIR))

##########################################################################
def SelectFile(parent, dfltDir, wildcard, title='Open file', multSels=False):
    ''' Select existing file. '''
    if dfltDir is None: dfltDir = QDir.currentPath()
    return getUseSel(FileDialog(parent, title, dfltDir, wildcard, mode=SEL_FILES if multSels else SEL_FILE))

##########################################################################
def SaveFile(parent, fileSpec, wildcard, title='Specify file name to save...', dfltFn=None):
    ''' Select or specify file to save to. '''
    if isinstance(dfltFn, str):
        fileSpec = joinPaths(fileSpec, dfltFn)
    return getUseSel(FileDialog(parent, title, fileSpec, wildcard, mode=SAVE_FILE))

##########################################################################
def getUseSel(dlg):
    if dlg.exec_():
        sels = dlg.selectedFiles()
        mode = dlg.selMode
        if isinstance(sels, list):
            rslt = sels[:] if mode==SEL_FILES else sels[0]
    else:
        rslt = ''
    #dlg.close()
    return rslt            
    
##########################################################################