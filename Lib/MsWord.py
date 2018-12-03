####### Importing Python modules ###############
import os
import win32com.client, pythoncom, time
from Utils.PyUtils import *
from Utils.WordConstants import *

####### Importing user defined packages ###############
RetChar = chr(13)
W, H = 0,1

ColWidth = 4
TopSpacing = 0.04
AdsSpacing = 0.01
MergedDim = {True: (577.5, 108.0), False: (292.0, 216.0),
             'SGL': (291.5, 108.0),
             'SGLR': (585.0, 108.0),
             'HAPR': (585.0, 325.5)}

##############################################################################
##############################################################################
class MsWord:

    ##########################################################################
    def __init__(self, visible=False):
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        self.msWord = win32com.client.Dispatch("Word.Application")
        self.msWord.DisplayAlerts = False
        self.isSaved = False

    ##########################################################################
    def __del__(self):
        pythoncom.CoUninitialize()

    ##########################################################################
    def new(self, filename=None):
        """
        Create a new Word document. If 'filename' specified,
        use the file as a template.
        """
        self.actvDoc = self.msWord.Documents.Add()
        self.selection = self.msWord.Selection
        self.msWord.Visible = True

    ##########################################################################
    def openDoc(self, filename):
        """
        open an existing Word document for editing.
        """
        self.actvDoc = self.msWord.Documents.Open(filename)
        self.selection = self.msWord.Selection
        self.msWord.Visible = True

    ##########################################################################
    def setFontSize(self, fontSize=10):
        
        #self.actvDoc.Content.Font.Size = fontSize
        self.selection.Font.Size = fontSize
        
        #style = self.actvDoc.styles['Normal']
        #font = style.font
        
        #font.size = Pt(1)

    ##############################################################################
    # New formatting using 8.5 X 11
    ##############################################################################
    def fmtSection(self, sectionId, withBorder=False):
        
        sectn = self.actvDoc.Sections[sectionId-1]
        sectn.PageSetup.Orientation = wdOrientPortrait

        sectn.PageSetup.TopMargin    = BulletinMargin
        sectn.PageSetup.BottomMargin = BulletinMargin
        sectn.PageSetup.LeftMargin   = BulletinMargin
        sectn.PageSetup.RightMargin  = BulletinMargin
        sectn.PageSetup.Gutter       = ZeroPt

        sectn.PageSetup.HeaderDistance = ZeroPt
        sectn.PageSetup.FooterDistance = ZeroPt

        if withBorder:
            for s in (wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight):
                sectn.Borders(s).LineStyle = wdLineStyleDouble
                sectn.Borders(s).LineWidth = wdLineWidth025pt
        return sectn
    
    #########################################################################
    def fmtTable(self, sectn, numRows=2, numCols=2, isAdsPage=False):
        selTbl = self.actvDoc.Tables.Add(Range=sectn.Range, NumRows=numRows, NumColumns=numCols,
                                         DefaultTableBehavior=wdWord9TableBehavior,
                                         AutoFitBehavior=wdAutoFitFixed)
    
        selTbl.Style = "Table Grid"
        lineStyle = wdLineStyleSingle if isAdsPage else wdLineStyleDouble
        lineWidth = wdLineWidth075pt if isAdsPage else wdLineWidth025pt
        selTbl.Borders.OutsideLineStyle = lineStyle  
        selTbl.Borders.OutsideLineWidth = lineWidth 
        selTbl.Borders.InsideLineStyle = lineStyle 
        selTbl.Borders.InsideLineWidth = lineWidth 

        tblMargin = AdsTableMargin if isAdsPage else ZeroPt
        selTbl.TopPadding = tblMargin
        selTbl.BottomPadding = tblMargin
        selTbl.LeftPadding = tblMargin
        selTbl.RightPadding = tblMargin
        selTbl.Spacing = ZeroPt
        selTbl.AllowPageBreaks = False
        selTbl.AllowAutoFit = False

        selTbl.Borders.Shadow = False
        
        selTbl.Rows.HeightRule = wdRowHeightExactly
        selTbl.Rows.Alignment = wdAlignRowCenter
        if numCols>1:
            selTbl.Columns.PreferredWidthType = wdPreferredWidthPoints
        
        return selTbl
    
    ##########################################################################
    def fmtFrontPage(self, sectn):
        selTbl = self.fmtTable(sectn)
        
        selTbl.Rows(1).Height = InchesToPoints(1.245)
        selTbl.Rows(2).Height = InchesToPoints(9.40)
        
        selTbl.Columns(1).PreferredWidth = InchesToPoints(1.875)
        selTbl.Columns(2).PreferredWidth = InchesToPoints(6.3)
        
        return selTbl

    ##########################################################################
    def fmt2ndPage(self, sectn):
        selTbl = self.fmtTable(sectn, numCols=1)

        selTbl.Rows(1).Height = InchesToPoints(9.145)
        selTbl.Rows(2).Height = InchesToPoints(1.50)

        return selTbl

    ##########################################################################
    def fmtAdsPage(self, sectn, pageId):
        self.selectPage(pageId=pageId)
        self.setFontSize(fontSize=1)
        return self.fmtTable(sectn, numRows=1, isAdsPage=True)

    ##########################################################################
    def loadAds(self, adsTbl, row, adsLst, adSize=''):
        adsPerRow = len(adsLst)
        if adsPerRow==1:
            adsTbl.Rows(row).Cells.Merge()
            adSize += 'R'
        adsTbl.Cell(row, 1).Select()
        for ind, adFn in enumerate(adsLst):
            if isFile(adFn):
                inlineImag = self.selection.InlineShapes.AddPicture(adFn, False, True)
                addMsg('====>> Loaded ads: {0}'.format(adFn))
                self.fitImag2MergedCells(inlineImag, adSize)
            if ind==0 and adsPerRow==2:
                self.moveSelectionRightByCell()
                    
    ##########################################################################
    def loadCustAds(self, adsTbl, adInfo, adFn):
        row, col = int(adInfo[1]), int(adInfo[3])
        mergeCol = adInfo[-1]=='C'
        
        if mergeCol:
            adsTbl.Cell(row, col).Merge(adsTbl.Cell(row, col+1))
        else:
            adsTbl.Cell(row, col).Merge(adsTbl.Cell(row+1, col))
            
        adsTbl.Cell(row, col).Range.Select()
        inlineImag = self.selection.InlineShapes.AddPicture(adFn, False, True)
        self.fitImag2MergedCells(inlineImag, mergeCol)
                    
                    

    ##########################################################################
    #(PORTRAIT (0), 7772400, 10058400)  # (Inches(8.5), Inches(11)) WD_ORIENT.LANDSCAPE
    def setStdLayout(self):
        self.actvDoc.PageSetup.Orientation = wdOrientPortrait

        self.actvDoc.PageSetup.TopMargin    = PageMargin
        self.actvDoc.PageSetup.BottomMargin = PageMargin
        self.actvDoc.PageSetup.LeftMargin   = PageMargin
        self.actvDoc.PageSetup.RightMargin  = PageMargin

    ##############################################################################
    #(PORTRAIT (0), 7772400, 100584fmtFrontPage00)  # (Inches(8.5), Inches(11)) WD_ORIENT.LANDSCAPE
    def fmtFrontPage__(self):
        self.actvDoc.PageSetup.Orientation = wdOrientLandscape

        self.actvDoc.PageSetup.TopMargin    = BulletinMargin
        self.actvDoc.PageSetup.BottomMargin = BulletinMargin
        self.actvDoc.PageSetup.LeftMargin   = BulletinMargin
        self.actvDoc.PageSetup.RightMargin  = BulletinMargin
        self.actvDoc.PageSetup.Gutter       = ZeroPt
        self.actvDoc.PageSetup.GutterPos    = wdGutterPosLeft

        self.actvDoc.PageSetup.HeaderDistance = ZeroPt
        self.actvDoc.PageSetup.FooterDistance = ZeroPt

        self.actvDoc.PageSetup.PageWidth    = InchesToPoints(17)
        self.actvDoc.PageSetup.PageHeight   = InchesToPoints(11)

        self.actvDoc.PageSetup.VerticalAlignment = wdAlignVerticalTop
        self.actvDoc.PageSetup.BookFoldPrinting = True
        self.actvDoc.PageSetup.BookFoldPrintingSheets = 1
        
    ##############################################################################
    #(PORTRAIT (0), 7772400, 10058400)  # (Inches(8.5), Inches(11)) WD_ORIENT.LANDSCAPE
    def fmtBodyPages(self, sectId=1):
        for s in (wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight):
            self.selection.Sections(sectId).Borders(s).LineStyle = wdLineStyleDouble
            self.selection.Sections(sectId).Borders(s).LineWidth = wdLineWidth025pt

        self.selection.Sections(sectId).Borders.DistanceFrom = wdBorderDistanceFromPageEdge
        self.selection.Sections(sectId).Borders.DistanceFromTop = 12
        self.selection.Sections(sectId).Borders.DistanceFromBottom = 12
        self.selection.Sections(sectId).Borders.DistanceFromLeft = 12
        self.selection.Sections(sectId).Borders.DistanceFromRight = 12

        self.selection.Sections(sectId).Borders.AlwaysInFront = True
        self.selection.Sections(sectId).Borders.SurroundHeader = False
        self.selection.Sections(sectId).Borders.SurroundFooter = False
        self.selection.Sections(sectId).Borders.EnableFirstPageInSection = True
        self.selection.Sections(sectId).Borders.EnableOtherPagesInSection = True

    ##############################################################################
    #(PORTRAIT (0), 7772400, 10058400)  # (Inches(8.5), Inches(11)) WD_ORIENT.LANDSCAPE
    def fmtAdsPages(self, sectId=1):

        print('====>> fmtAdsPages: ', sectId, len(self.actvDoc.Sections))

        for s in (wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight):
            self.selection.Sections(sectId).Borders(s).LineStyle = wdLineStyleDouble
            self.selection.Sections(sectId).Borders(s).LineWidth = wdLineWidth025pt

        self.selection.Sections(sectId).Borders.DistanceFrom = wdBorderDistanceFromPageEdge
        self.selection.Sections(sectId).Borders.DistanceFromTop = 12
        self.selection.Sections(sectId).Borders.DistanceFromBottom = 12
        self.selection.Sections(sectId).Borders.DistanceFromLeft = 12
        self.selection.Sections(sectId).Borders.DistanceFromRight = 12

        self.selection.Sections(sectId).Borders.AlwaysInFront = True
        self.selection.Sections(sectId).Borders.SurroundHeader = False
        self.selection.Sections(sectId).Borders.SurroundFooter = False
        self.selection.Sections(sectId).Borders.EnableFirstPageInSection = True
        self.selection.Sections(sectId).Borders.EnableOtherPagesInSection = True

    ##########################################################################
    def addFrontPage__(self, sectId=1):
        selection = self.selection.Sections(sectId)
        Table = self.actvDoc.Tables.Add(Range=selection.Range, NumRows=2, NumColumns=2, DefaultTableBehavior=wdWord9TableBehavior, AutoFitBehavior=wdAutoFitFixed)


        print('====>> addFrontPage *1- Front page DONE', self.numSections(), self.numPages())

        Table.Style = "Table Grid"

        Table.Borders.OutsideLineStyle = wdLineStyleDouble  
        Table.Borders.OutsideLineWidth = wdLineWidth025pt 
        Table.Borders.InsideLineStyle = wdLineStyleDouble 
        Table.Borders.InsideLineWidth = wdLineWidth025pt 

        Table.TopPadding = ZeroPt
        Table.BottomPadding = ZeroPt
        Table.LeftPadding = ZeroPt
        Table.RightPadding = ZeroPt
        Table.Spacing = ZeroPt
        Table.AllowPageBreaks = True
        Table.AllowAutoFit = False

        Table.Borders.Shadow = False
        
        Table.Rows.HeightRule = wdRowHeightExactly
        Table.Rows(1).Height = InchesToPoints(1.22)
        Table.Rows(2).Height = InchesToPoints(9.38)
        
        Table.Columns.PreferredWidthType = wdPreferredWidthPoints
        Table.Columns(1).PreferredWidth = InchesToPoints(1.875)
        Table.Columns(2).PreferredWidth = InchesToPoints(6.3)
        
        # Align table to center of the page
        Table.Rows.Alignment = wdAlignRowCenter

        print('====>> addFrontPage *2- Front page DONE', self.numSections(), self.numPages())

    ##########################################################################
    def addTblPerPage(self, pageId=1, numRows=1, numCols=2):
        if pageId > self.numPages():
            self.addPage()
            
        self.selectPage(pageId)
        Table = self.actvDoc.Tables.Add(Range=self.selection.Range, NumRows=numRows, NumColumns=numCols,
                                        DefaultTableBehavior=wdWord9TableBehavior, AutoFitBehavior=wdAutoFitFixed)

        Table.Style = "Table Grid"
        Table.Borders.OutsideLineStyle = wdLineStyleNone  
        Table.Borders.InsideLineStyle = wdLineStyleNone 

        Table.TopPadding = ZeroPt
        Table.BottomPadding = ZeroPt
        Table.LeftPadding = ZeroPt
        Table.RightPadding = ZeroPt
        Table.Spacing = ZeroPt
        Table.AllowPageBreaks = False
        Table.AllowAutoFit = False
        
        Table.Columns.PreferredWidthType = wdPreferredWidthPoints

        Table.Rows.HeightRule = wdRowHeightExactly
        Table.Rows(1).Height = InchesToPoints(9.05)
        Table.Rows(2).Height = InchesToPoints(1.6)

        # Align table to center of the page
        Table.Rows.Alignment = wdAlignRowCenter
        
        return Table

    ##########################################################################
    def addTable(self, pageId=1, numRows=1, numCols=2):
        if pageId > self.numPages():
            self.addPage()

        self.selectPage(pageId)
        Table = self.actvDoc.Tables.Add(Range=self.selection.Range, NumRows=numRows, NumColumns=numCols,
                                        DefaultTableBehavior=wdWord9TableBehavior, AutoFitBehavior=wdAutoFitFixed)

        Table.Style = "Table Grid"
        Table.Borders.OutsideLineStyle = wdLineStyleSingle  
        Table.Borders.InsideLineStyle = wdLineStyleSingle 

        Table.TopPadding = ZeroPt
        Table.BottomPadding = ZeroPt
        Table.LeftPadding = ZeroPt
        Table.RightPadding = ZeroPt
        Table.Spacing = ZeroPt
        Table.AllowPageBreaks = False
        Table.AllowAutoFit = False

        Table.Columns.PreferredWidthType = wdPreferredWidthPoints

        Table.Rows.HeightRule = wdRowHeightExactly
        for r in range(numRows):
            Table.Rows(r+1).Height = InchesToPoints(1.25)

        # Align table to center of the page
        Table.Rows.Alignment = wdAlignRowCenter
        
        return Table

    ##########################################################################
    def addTextCtrl(self, pageId, selFn):
        objTemplate = Templates(1) 
        objBB = objTemplate.BuildingBlockEntries.Item(1) 
        self.actvDoc.ActiveWindow.Panes(1).Pages(pageId).Rectangles(1).Range.Select()
        #objBB.Appearance = wdContentControlHidden
        objBB.Insert(self.selection.Range)

        objBB.TopMargin    = BulletinMargin
        objBB.BottomMargin = BulletinMargin
        objBB.LeftMargin   = BulletinMargin
        objBB.RightMargin  = BulletinMargin

    ##########################################################################
    def addTextBox(self, pageId, selFn):
        self.actvDoc.ActiveWindow.Panes(1).Pages(pageId).Rectangles(1).Range.Select()
        objCC  = self.selection.Range.ContentControls.Add(wdContentControlRichText)
        objCC.Range.InsertFile(selFn, Range="", ConfirmConversions=False, Link=False, Attachment=False)
        #objCC.Selection.ParagraphFormat.VerticalAlignment = wdAlignVerticalTop
        #objCC.Appearance = wdContentControlHidden

        self.selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

        '''
        self.selection.Sections(1).Borders.DistanceFrom = wdBorderDistanceFromPageEdge
        self.selection.Sections(1).Borders.DistanceFromTop = 12
        self.selection.Sections(1).Borders.DistanceFromBottom = 12
        self.selection.Sections(1).Borders.DistanceFromLeft = 12
        self.selection.Sections(1).Borders.DistanceFromRight = 12
        '''

    ##########################################################################
    def selectSection(self, sectnId=0):
        if sectnId==0:
            self.actvDoc.Sections.Last()
        else:
            self.actvDoc.Sections(sectnId)

    ##########################################################################
    def selectPage(self, pageId=0):
        if pageId==0:  pageId = self.numPages()
        self.actvDoc.ActiveWindow.Panes(1).Pages(pageId).Rectangles(1).Range.Select()

    ##########################################################################
    def addImageTheme(self, imagFn):
        if not isFile(imagFn): return
        self.actvDoc.Tables(1).Cell(1,1).Select()
        self.selection.InlineShapes.AddPicture(imagFn, False, True)

        self.selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        self.selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter

    ##########################################################################
    def importFileToCell(self, tblId, row, col, sideBarFn):
        self.actvDoc.Tables(tblId).Cell(row, col).Select()
        self.selection.InsertFile(sideBarFn, Range="", ConfirmConversions=False, Link=False, Attachment=False)

    ##########################################################################
    def selectionInsertFile(self, selFn):
        self.selection.InsertFile(selFn, Range="", ConfirmConversions=False, Link=False, Attachment=False)

    ##########################################################################
    def addTxtBox(self, page):
        self.selection.Range.ContentControls.Add(wdContentControlRichText)

    ##########################################################################
    def cfgSidbar(self, tmpltFn):
        self.actvDoc.Tables(1).Cell(2,1).Select()
        self.myCfg = MyConfig(joinPaths(self.path, 'tmplates', 'template1.cfg'))

        for sectn in self.myCfg.sections():
            self.insertStyledParagraph(sectn, wdStyleHeader)
            for opt in self.myCfg.options(sectn):
                self.insertStyledParagraph(opt, wdStyleStrong)
                for info in self.myCfg.get(sectn, opt):
                    self.insertStyledParagraph(info, wdStyleNormalIndent)

    ##########################################################################
    ##########################################################################
    def addSection(self):
        self.actvDoc.Sections.Add()

    ##########################################################################
    def addPage(self):
        self.moveSelectionToEndDoc()
        self.selection.InsertNewPage()

    ##########################################################################
    # Not working
    def delPage(self, pageId):
        if pageId >= self.numPages():
            selPage = self.actvDoc.GoTo(wdGoToPage, pageId)
            selPage.Delete()
            self.collapseSelectionToEnd()

    ##########################################################################
    # Not working
    def delLastPage(self):
        self.collapseSelectionToEnd()
        self.selection.Delete()
        self.collapseSelectionToEnd()

    ##########################################################################
    # Not working
    def clearPageContent(self, pageId):
        selPage = self.actvDoc.GoTo(wdGoToPage, pageId)
        selPage.Delete()
        self.collapseSelectionToEnd()
        self.selection.InsertBreak(wdPageBreak)

    ##########################################################################
    def prepAdsPage(self, rowHeiht):
        Table = self.actvDoc.Tables.Add(Range=self.selection.Range, NumRows=1, NumColumns=2, DefaultTableBehavior=wdWord9TableBehavior, AutoFitBehavior=wdAutoFitFixed)

        Table.Style = "Table Grid"
        Table.Borders.OutsideLineStyle = wdLineStyleNone  
        Table.Borders.InsideLineStyle = wdLineStyleNone 

        addSpace = InchesToPoints(AdsSpacing)
        #Table.TopPadding = InchesToPoints(TopSpacing)
        Table.TopPadding = addSpace
        Table.BottomPadding = addSpace
        Table.LeftPadding = addSpace
        Table.RightPadding = addSpace
        Table.Spacing = addSpace
        Table.AllowPageBreaks = False
        Table.AllowAutoFit = False

        Table.Columns.PreferredWidthType = wdPreferredWidthPoints
        Table.Columns.PreferredWidth = InchesToPoints(ColWidth)

        Table.Rows.HeightRule = wdRowHeightExactly
        Table.Rows.Height = InchesToPoints(rowHeiht)

        # Align table to center of the page
        Table.Rows.Alignment = wdAlignRowCenter

        self.myTbl = Table

    ##########################################################################
    def addRow(self, selTbl, rowHeiht):
        newRow = selTbl.Rows.Add()
        newRow.Height = InchesToPoints(rowHeiht)
        
    ##########################################################################
    def setRowHeight(self, selTbl, rowId, rowHeiht):
        selTbl.Rows(rowId).Height = InchesToPoints(rowHeiht)
        
    ##########################################################################
    def delRow(self, tbl, row=0):
        rowId = row if row>0 else self.numRows(tbl)
        self.actvDoc.Tables(tbl).Rows(rowId).Delete()

    ##########################################################################
    def loadAdsToRow(self, tbl, row, adsLst, adSize=''):
        if len(adsLst)==1:
            self.actvDoc.Tables(tbl).Rows(row).Cells.Merge()
            adSize += 'R'
        self.actvDoc.Tables(tbl).Cell(row, 1).Select()
        for adFn in adsLst:
            if isFile(adFn):
                inlineImag = self.selection.InlineShapes.AddPicture(adFn, False, True)
                addMsg('====>> Loaded ads: {0}'.format(adFn))
                self.fitImag2MergedCells(inlineImag, adSize)
            self.moveSelectionRightByCell()


    ##########################################################################
    def loadCustAds__(self, tbl, adInfo, adFn):
        row, col = int(adInfo[1]), int(adInfo[3])
        mergeCol = adInfo[-1]=='C'
        myTbl = self.actvDoc.Tables(tbl)
        
        if mergeCol:
            myTbl.Cell(row, col).Merge(myTbl.Cell(row, col+1))
        else:
            myTbl.Cell(row, col).Merge(myTbl.Cell(row+1, col))
            
        myTbl.Cell(row, col).Range.Select()
        inlineImag = self.selection.InlineShapes.AddPicture(adFn, False, True)
        self.fitImag2MergedCells(inlineImag, mergeCol)

    ##########################################################################
    def fitImag2MergedCells(self, inlineImag, mergeCol):
        if mergeCol in MergedDim:
            inlineImag.Width = MergedDim[mergeCol][W]
            inlineImag.Height = MergedDim[mergeCol][H]

    ##########################################################################
    def loadCell(self, tbl, row, col, adFn, mergeCell):
        if not isFile(adFn):
            return print('====>> Unable to locate file: {0}'.format(adFn))

        if mergeCell:
            self.actvDoc.Tables(tbl).Rows(row).Cells.Merge()
        self.actvDoc.Tables(tbl).Cell(row, col).Select()
        self.selection.InlineShapes.AddPicture(adFn, False, True)

    ##########################################################################

    # Word File selecting and moving around...
    def collapseSelectionToStart(self):
        # ensures insertion point is at the beginning of the selection
        self.selection.Collapse(wdCollapseStart)

    def collapseSelectionToEnd(self):
        # ensures insertion point is at the end of the selection
        self.selection.Collapse(wdCollapseEnd)

    def moveSelectionToEndDoc(self):
        # moves insertion point is at the end of the document
        self.selection.EndOf(wdStory, wdMove)

    def moveSelectionToEndTable(self):
        # moves insertion point is at the end of the colunm
        self.selection.EndOf(wdTable, wdMove)

    def moveSelectionToEndColunm(self):
        # moves insertion point is at the end of the colunm
        self.selection.EndOf(wdColumn, wdMove)

    def moveSelectionLeftByChars(self, num=1):
        self.selection.MoveLeft(wdCharacter, num)

    def moveSelectionRightByChars(self, num=1):
        self.selection.MoveRight(wdCharacter, num)

    def moveSelectionRightByParagraphs(self, num=1):
        self.selection.MoveRight(wdParagraph, num)

    def moveSelectionRightByCell(self, num=1):
        self.selection.MoveRight(wdCell, num)

    def moveSelectionDownByCell(self, num=1):
        self.selection.MoveDown(wdCell, num)

    # Word insertions...
    def insertPageBreak(self):
        self.selection.InsertBreak(wdPageBreak)

    def insertSectionBreakNextPage(self) :
        self.selection.InsertBreak(wdSectionBreakNextPage)

    def insertSectionBreakContinuous(self) :
        self.selection.InsertBreak(wdSectionBreakContinuous)

    def insertLandscapeSection(self, left=1.0, right=1.0, top=1.0, bottom=1.0,
                        header=0.5, footer=0.5):
        # second, do a PageSetUp on this section
        self.setSectionPageSetUp('Land',left,right,top,bottom,header,footer)
        # first mark off the previous section and start a new page
        self.insertSectionBreakNextPage()

    def insertPortraitSection(self, left=1.0, right=1.0, top=1.0, bottom=1.0,
                        header=0.5, footer=0.5):
        # second, do a PageSetUp on this section
        self.setSectionPageSetUp('Port',left,right,top,bottom,header,footer)
        # first mark off the previous section and start a new page
        self.insertSectionBreakNextPage()


    # T E X T  insertions...
    def insertWordFile(self, filename):
        self.selection.InsertFile(filename)

    def insertText(self, text):
        self.insertAfter(text)

    def insertStyledText(self, text, stylename):
        self.insertAfter(text, stylename)

    def insertStyledParagraph(self, text, stylename):
        if text[-1] != '\n':
            text = text + '\n'
        self.insertAfter(text, stylename)

    def insertSectionHeader(self, text):
        if text[-1] != '\n':
            text = text + '\n'
        self.selection.InsertAfter(text)
        #self.selection.Style = wdStyleStrong
        self.selection.Style = wdStyleHeader
        self.moveSelectionToEndColunm()

    def insertAfter(self, text, stylename=None):
        self.selection.InsertAfter(text)
        if stylename is not None:
            self.selection.Style = stylename
        self.moveSelectionToEndColunm()
        #self.collapseSelectionToEnd()


    ##########################################################################
    def saveDoc(self, filename=None):
        fileNm = os.path.abspath(filename)
        if self.isSaved:
            self.actvDoc.Save(fileNm)
        else:
            self.actvDoc.SaveAs(fileNm)
        self.isSaved = True

    ##########################################################################
    def save_as(self, filename):
        self.actvDoc.SaveAs(filename)

    ##########################################################################
    def saveAsPdf(self, pdfFileNm):
        try:
            self.actvDoc.SaveAs(pdfFileNm,
                                FileFormat=WdFormatPDF,
                                EmbedTrueTypeFonts=True,
                                LockComments=False,
                                ReadOnlyRecommended=False)
        except:
            print('Unable to save: ', pdfFileNm)

    ##########################################################################
    def printFile(self):
        """
        Print the active document.
        """
        self.msWord.Application.PrintOut()
        time.sleep(2)

    ##########################################################################
    def printPages(self, frPage, toPage=None):
        """
        Print the active document.
        """
        if toPage is None:
            self.msWord.Application.PrintOut(frPage)
        else:
            self.msWord.Application.PrintOut(frPage, toPage)
        time.sleep(2)

    ##########################################################################
    def printPDF(self, pdfDoc):
        ie = client.Dispatch("InternetExplorer.Application")
        ie.Navigate(pdfDoc)

        if ie.Busy:
            time.sleep(1)

        ie.Document.printAll()
        time.sleep(2)

        ie.Quit()

    ##########################################################################
    def closeDoc(self, saveDoc=True):
        """
        Close the active document.
        """
        self.actvDoc.Close(saveDoc)

    ##########################################################################
    def quitApp(self):
        """
        Quit Word.
        """
        try:
            return self.msWord.Quit()
        except:
            pass

    ##########################################################################
    def find_and_replace(self, find_str, replace_str):
        """
        Find all occurances of 'find_str' and replace with 'replace_str'
        in the active document.
        """
        self.selection.HomeKey(Unit=wdStory)
        find = self.selection.Find
        find.Text = find_str
        while self.selection.Find.Execute():
            self.selection.TypeText(Text=replace_str)


    ##########################################################################
    def find_replace_in_canvas(self, find_str, replace_str):
        for shp in self.actvDoc.Shapes:
            try:
                myTxt = shp.TextFrame.TextRange.Text
                if find_str in myTxt:
                    shp.TextFrame.TextRange.Text = myTxt.replace(find_str, replace_str)
            except:
                pass

    ##########################################################################
    ##########################################################################
    def findIt(self):
        return self.actvDoc.Selection.Find

    ##########################################################################
    def gotoPage(self, pageId=0):
        if pageId==0:
            self.toLastPage()
        elif pageId <= self.numPages():
            self.actvDoc.GoTo(wdGoToPage, pageId).Select()

    ##########################################################################
    def toLastPage(self):
        self.actvDoc.GoTo(wdGoToPage, self.numPages()).Select()

    ##########################################################################
    ##########################################################################
    def numSections(self):
        return len(self.actvDoc.Sections)

    ##########################################################################
    def numPages(self):
        return self.actvDoc.ActiveWindow.Panes(1).Pages.Count

    ##########################################################################
    def numTables(self):
        return self.actvDoc.Tables.Count

    ##########################################################################
    def numRows(self, tblId):
        return self.actvDoc.Tables(tblId).Rows.Count

    ##########################################################################
    def numColumns(self, tblId):
        return self.actvDoc.Tables(tblId).Columns.Count

    ##########################################################################
    ##########################################################################
    def getTableOfContents(self):
        return self.actvDoc.TablesOfContents

    ##########################################################################
    def selSection(self, sectnId):
        return self.actvDoc.ActiveWindow.Panes(1).Sections(sectnId)

    ##########################################################################
    def selPage(self, pageId):
        return self.actvDoc.ActiveWindow.Panes(1).Pages(pageId)

    ##########################################################################
    def getParagraphs(self):
        return self.actvDoc.Paragraphs

    ##########################################################################
    def getTables(self):
        return self.actvDoc.Tables

    ##########################################################################
    def getRowsCount(self, tblId):
        return self.actvDoc.Tables(tblId).Rows.Count

    ##########################################################################
    def getTableSize(self, myTable):
        return (len(myTable.Columns), len(myTable.Rows))

    ##########################################################################
    def getCellData(self, tblObj, row, col):

        reqCell = tblObj.Cell(row, col).Range
        try: return str(reqCell.Text).replace('\x0c', '').split(RetChar)
        except: return ''

##############################################################################
##############################################################################
##############################################################################
def readTblHeader(tblObj, tableSize, limitCol=None):
    for row in range(1, min(2, tableSize[1])):
        numCols = min(limitCol, tableSize[0]) if isValid(limitCol, int, chkVal=False) else tableSize[0]
        rowData = getRowData(tblObj, row, numCols, offset=0)
        if any(rowData):
            return rowData
    return []

##############################################################################
def getCellData(cellObj):
    try:
        return str(cellObj.Text).split(RetChar)[0].replace('\x0c', '')
    except:
        return None

##############################################################################
def getRowData(tblObj, rowId, numCols, offset=1):
    rowData = []
    for col in range(numCols):
        try:
            cellData = getCellData(tblObj.Cell(rowId+offset, col+offset).Range)
            if isNull(cellData): break
            rowData.append(cellData)
        except:
            pass
    return rowData

##############################################################################
class converter:

    def __init__(self, factor, divisor=1):
        if factor != 1 and divisor != 1:
            self.factor = float(factor)
            self.divisor = float(divisor)
            self.operation = self.scale
        elif divisor != 1 :
            self.divisor = float(divisor)
            self.operation = self.divide
        elif factor != 1 :
            self.factor = float(factor)
            self.operation = self.multiply
        else:
            self.operation = self.noop
        
    def multiply(self, arg): return arg * self.factor
    def divide(self, arg):   return arg / self.divisor
    def scale(self, arg):    return arg * self.factor / self.divisor
    def noop(self, arg):     return arg+0.0
    def __call__(self, arg): return self.operation(arg)

PicasToPoints = converter(12)
PointsToPicas = converter(1, 12)
InchesToPoints = converter(72)
PointsToInches = converter(1, 72)
LinesToPoints = converter(12)
PointsToLines = converter(1, 12)
InchesToCentimeters = converter(254, 100)
CentimetersToInches = converter(100, 254)
CentimetersToPoints = converter(InchesToPoints(100), 254)
PointsToCentimeters = converter(254, InchesToPoints(100))
ZeroPt = InchesToPoints(0)
AdsTableMargin = 1
DfltSpacing = InchesToPoints(0.02)
DblSpacing = InchesToPoints(0.04)
PageMargin = InchesToPoints(0.24)
BulletinMargin = InchesToPoints(0.17)

##############################################################################
# End of File
##############################################################################
