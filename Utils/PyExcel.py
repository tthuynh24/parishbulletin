##############################################################################
# Module: DB_Admin                                                            #
##############################################################################

import os
import sys
from .openpyxl import Workbook, load_workbook
from .openpyxl.cell import get_column_letter
from .PyUtils import isNull, isNotNull

##########################################################################
def isValid(selData):
    dataType = type(selData)
    if dataType in (int, float): return selData>0
    if dataType in (str, list, tuple, dict):
        if isinstance(selData, str):
            selData = selData.strip()
        return len(selData)>0
    return isNotNull(selData)

##########################################################################
def isFile(selFn):
    return os.path.isfile(selFn) if isValid(selFn) else False

class PyExcel:
    ##########################################################################
    def __init__(self, excelFile):
        self.myDbDct = {}
        self.excelFile = excelFile
        self.xlsBook = None
        self.xlsSheet = None
        self.startExcel()

    ##########################################################################
    def startExcel(self):
        if isFile(self.excelFile):
            self.xlsBook = load_workbook(self.excelFile)
        else:
            self.xlsBook = Workbook()

    ##########################################################################
    def addPage(self, wkSheetNm=None, asLastPage=True):
        if asLastPage:
            wkSheet = self.xlsBook.create_sheet()     # Add as last page
        else:
            wkSheet =  self.xlsBook.create_sheet(0)    # Add as first page
        if isNotNull(wkSheetNm):
            wkSheet.title = wkSheetNm
        self.xlsSheet = wkSheet
        return wkSheet

    ##########################################################################
    def rennamePage(self, newNm, oldNm=None):
        if isNull(oldNm):
            self.xlsBook.active.title = newNm
        else:
            self.xlsBook.get_sheet_by_name(oldNm).title = newNm

    ##########################################################################
    def getAllWkSheets(self):
        return self.xlsBook.get_sheet_names()

    ##########################################################################
    def getWkSheet(self, wkSheetNm=None):
        try:
            xlsSheet = self.xlsBook.active if isNull(wkSheetNm) else self.xlsBook.get_sheet_by_name(wkSheetNm)
        except :
            return None
        self.xlsSheet = xlsSheet
        return self.xlsSheet
        
    ##########################################################################
    def delWorkSheet(self, wkSheetNm=None):
        wkSheet = self.getWkSheet(wkSheetNm)
        if isNotNull(wkSheet):
            self.xlsBook.remove_sheet(wkSheet)
        return True

    ##########################################################################
    def clearXls(self):
        while isNotNull(self.delWorkSheet()): pass

    ##########################################################################
    def clearWorkSheet(self, wkSheetNm=None):
        wkSheet = self.getWkSheet(wkSheetNm)
        if wkSheetNm is None:
            wkSheetNm = wkSheet.title
        if isNotNull(wkSheet):
            self.xlsBook.remove_sheet(wkSheet)
        wkSheet = self.addPage()
        wkSheet.title = wkSheetNm
        self.xlsSheet = wkSheet
        return wkSheet

    ##########################################################################
    def cleanUp(self, wkSheet):
        wkSheet.garbage_collect()

    ##########################################################################
    def getPageSize(self, wkSheetNm=None):
        wkSheet = self.getWkSheet(wkSheetNm)
        return len(wkSheet.rows), len(wkSheet.columns)

    ##########################################################################
    def iterXlsCols(self, wkSheetNm=None, offsetCol=0):
        # openpyxl.worksheet.Worksheet.columns() 
        for col in self.getWkSheet(wkSheetNm).columns[offsetCol:]:
            yield [cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in col]

    ##########################################################################
    def iterXlsRows(self, wkSheetNm=None, offset=0, dflt=None):
        for row in self.getWkSheet(wkSheetNm).rows[offset:]:
            if isNull(dflt):
                yield [cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in row]
            else:
                yield [dflt if isNull(cell.value) else cell.value.strip() if isinstance(cell.value, str) else cell.value for cell in row]

    ##########################################################################
    def doesWsExist(self, wkSheetNm):
        return wkSheetNm in self.getAllWkSheets()

    ##########################################################################
    def parseWkSheet(self):
        for row in self.iterXlsRows(): # it brings a new method: iter_rows()
            print(row[1:4])

    ##########################################################################
    def addRows(self, wkSheet):
        for i in range(11):
            sRang = i*6
            wkSheet.append([c for c in range(sRang, sRang+6, 1)])
        self.saveXls()

    ##########################################################################
    def addTblEntry(self, tblNm, tblEnryt):
        self.myDbDct[tblNm].append(tblEnryt)
        return 1

    ##########################################################################
    def geNumRecs(self, wkSheetNm=None):
        # openpyxl.worksheet.Worksheet.rows()
        return len(self.getWkSheet(wkSheetNm).rows[1:])

    ##########################################################################
    def saveXls(self, wkSheetNm=None):
        self.xlsBook.save(self.excelFile if isNull(wkSheetNm) else wkSheetNm)

    ##########################################################################
    ##########################################################################
    def getColNms(self, tblNm):

        print('====>> DB Table: ', tblNm, cnt, colNms)
        
        
    ##########################################################################
    def cellVal(self, row, column, wkSheet=None, value=None):
        if isNull(wkSheet):
            wkSheet = self.getWkSheet()
            if isNull(wkSheet):
                wkSheet =  self.xlsSheet
        if isNull(value):
            return wkSheet.cell(row = row, column = column).value
        wkSheet.cell(row = row, column = column).value = value
        
##############################################################################
##############################################################################
def parseRec(recData, parmLst):
    return [(parmLst[ind], data) for ind, data in enumerate(recData)]

##############################################################################
##############################################################################
def getValStr(ind, inVal):
    if isinstance(inVal, int):
        val = '%d' % inVal
    elif isNull(inVal) or len(inVal)==0:
        return (', ""') if ind>0 else '""'
    else:
        val = inVal
    return (', %s' % inVal) if ind>0 else val

##############################################################################
def getTblNms():
    return ['Parish_parish',
            'Members_family',
            'Members_member',
            'Ministry_liturgy',
            'Ministry_association',
            'Ministry_membership',
            'ReligiousLife_baptism',
            'ReligiousLife_firstcommunion',
            'ReligiousLife_confirmation',
            'ReligiousLife_holymatrimony',
            'ReligiousLife_holyorders',
            'ReligiousLife_deceased',
            'Ministry_ministrymembers',
            'Finance_familydonation',
            'Finance_bldgdonation',
            'TNTT_tnttmembers',
            'Catechism_ccdmembers',
            'VietNamese_vnmembers',
            'Finance_weeklydonation',
            'Finance_miscfun',
            'BibleVerses_bibleverses',
            'BibleVerses_bibleversesel',
            ]

##############################################################################
# End of File
##############################################################################
