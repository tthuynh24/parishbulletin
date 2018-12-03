##############################################################################
# Module: DB_Admin                                                            #
##############################################################################
from __future__ import absolute_import

import os
import sys

##########################################################################
def isValid(selData):
    dataType = type(selData)
    if dataType in (int, float): return selData>0
    if dataType in (str, list, tuple, dict):
        if isinstance(selData, str):
            selData = selData.strip()
        return len(selData)>0
    return selData is not None

##########################################################################
def isFile(selFn):
    return os.path.isfile(selFn) if isValid(selFn) else False

class MsXls:

    ##########################################################################
    def __init__(self, excelFile):
        self.sqlCursor = None

    ##########################################################################
    def getColNms(self, tblNm):
        tblCols = []
        self.sqlCursor.execute('PRAGMA table_info(%s)' % tblNm)
        return [colInfo[1] for colInfo in self.sqlCursor.fetchall()]

    ##########################################################################
    def alterTbl(self, tblNm):
        if tblNm not in self.cnrtTbls:
            cmd = 'ALTER TABLE %s CONVERT TO CHARACTER SET utf8mb4 COLLATE utf8mb4_bin;'
            self.sqlCursor.execute( cmd % tblNm )
            self.sqlConn.commit()
            self.cnrtTbls.append(tblNm)

    ##########################################################################
    ##########################################################################
    def procReq(self):
        for tblNm in getTblNms():
            colNms = self.getColNms(tblNm)
            self.myDbDct.update({tblNm: [colNms]})
            cnt = 0
            for rowData in self.sqlCursor.execute('SELECT * FROM %s' % tblNm):
                cnt += self.addTblEntry(tblNm, rowData)

            print('====>> DB Table: ', tblNm, cnt, colNms)

        pickleData(self.myDbDct, 'DucMeLaVangDB.pkl')

        print('====>> getPklData: ', list(getPklData('DucMeLaVangDB.pkl').keys()))


##############################################################################
##############################################################################
def getValStr(ind, inVal):
    if isinstance(inVal, int):
        val = '%d' % inVal
    elif inVal is None or len(inVal)==0:
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

