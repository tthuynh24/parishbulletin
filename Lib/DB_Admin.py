##############################################################################
# Module: DB_Admin                                                            #
##############################################################################
from __future__ import absolute_import

import re
import ast
import datetime
from decimal import Decimal

# Create your models here.
from www.settings import DATABASES, ParishCfg, FinAdmCfg
from Utils.PyUtils import setTimeStamp, joinPaths, copyFile, ExecExtCmd, isNull, isNotNull
from .AppShared import *

from django.contrib.sites.models import Site
from django.db import connection, transaction
from django.db.models import get_models
from django.db.models.loading import get_models, get_app
from django.core.management import call_command

from Bible.models import *
from Catechism.models import *
from Finance.models import *
from Members.models import dbID, Family, Member, AddrHistory, MemberAddrs
from Ministry.models import *
from Parish.models import *
from TNTT.models import *
from VietNamese.models import *

from Utils.AddrParser import normalize_addr1
from Utils.PyExcel import *
from Utils.AddrUtils import *
from .MinistrySched import *

PhoneFmt = re.compile('(\d{3} \d{3}-\d{4})')
DateFmt = re.compile('(\d{4}-\d{2}-\d{2})')

LitergyType_Dct = dict((v, k) for k, v in LITERGY_CHOICES)
LITERGY_CHOICES = ast.literal_eval(ParishCfg.get('Parish', 'LITERGY_CHOICES'))
GIAOKHU_RegNum = ast.literal_eval(ParishCfg.get('Parish', 'GIAOKHU_RegNum'))
Parish_Info = ast.literal_eval(ParishCfg.get('Parish', 'Parish_Info'))
PYMT_OPTS = ast.literal_eval(ParishCfg.get('Parish', 'PYMT_OPTS'))
PYMT_DUR = ast.literal_eval(ParishCfg.get('Parish', 'PYMT_DUR'))

ADMIN_CHOICES = ast.literal_eval(FinAdmCfg.get('FinAdmin', 'ADMIN_CHOICES'))


DateFld = re.compile('date', re.I)
NON_MEM, FAM_ID, MEM_ID = 0, 100, 101

DB_ID = {}
Member_ID =     {MEM_ID:    1000000,
                 NON_MEM:   2000000,
                 FAM_ID:    1,
                }
WEB, DB = 1, 2

##############################################################################
##############################################################################
def isValidData(v):
    return v is not None and (isinstance(v, (int, float, datetime.datetime)) or (isinstance(v, str) and len(v)>0))

##############################################################################
def addSiteId(newRec):
    rec = Site.objects.get(id=newRec.get('id'))
    rec.domain = newRec.get('domain')
    rec.name = newRec.get('name')
    rec.save()

##############################################################################
def addParish_Org():
    newRec = [100] + [info for k, info in Parish_Info]
    return Parish.objects.create(id             = newRec[0],
                                 name           = newRec[1],
                                 tx_name        = newRec[2],
                                 fidNumber      = newRec[3],
                                 siteUrl        = newRec[4],
                                 paster         = newRec[5],
                                 addr_id        = newRec[6],
                                 phone          = newRec[7],
                                 mobile         = newRec[8],
                                 email          = newRec[9],
                                 officDays      = newRec[10],
                                 morningHrs     = newRec[11],
                                 afternoonHrs   = newRec[12],
                                )

##############################################################################
def InitID():
    global DB_ID
    for grp, num in GIAOKHU_RegNum.items():
        grpId = num//100000
        if grpId in DB_ID: continue
        addDbIds((grpId, num), 2)

    for grp, num in Member_ID.items():
        if grp in DB_ID: continue
        addDbIds((grp, num), 2)



##############################################################################
##############################################################################
def addDbRec(dbObj, newRec):
    return dbObj.create(**newRec)



##############################################################################
def addParish(newRec, recLen):
    return Parish.objects.create(Parish.objects)


##############################################################################
def addDbIds(newRec, recLen):
    return dbID.objects.create(**newRec)
    
    
    #global DB_ID
    #if recLen < 2: return
    #DB_ID.update({newRec[0]: newRec[1]})
    #return dbID.objects.create(groupName = newRec[0],
    #                           regNum    = newRec[1],
    #                          )

##############################################################################
def add_Bibles(newRec, recLen):
    return BibleVerses.objects.create(**newRec)

##############################################################################
def add_BibleSel(newRec, recLen):
    return BibleSel.objects.create(**newRec)

##############################################################################
def addFamily(newRec, recLen):
    return Family.objects.create(**newRec)

#############################################################################
def addMember(newRec, recLen):
    return Member.objects.create(**newRec)

#############################################################################
def addStudMember(newRec):
    return Member.objects.create(**newRec)

##############################################################################
def addBaptism(newRec, recLen):
    return Baptism.objects.create(**newRec)

##############################################################################
def addFirstComn(newRec, recLen):
    return FirstCommunion.objects.create(**newRec)

##############################################################################
def addConfirmation(newRec, recLen):
    return Confirmation.objects.create(**newRec)

##############################################################################
def addHolyMatrimony(newRec, recLen):
    return HolyMatrimony.objects.create(**newRec)

##############################################################################
def addHolyOrderaddHolyOrder(newRec, recLen):
    return HolyOrders.objects.create(**newRec)

##############################################################################
def addDeceased(newRec, recLen):
    return Deceased.objects.create(**newRec)

##############################################################################
def addMistryMem(newRec, recLen):
    return MinistryMembers.objects.create(**newRec)

##############################################################################
def add_TNTT(newRec, recLen):
    return TNTTMembers.objects.create(**newRec)

##############################################################################
def add_CCD(newRec, recLen):
    return CCDMembers.objects.create(**newRec)

##############################################################################
def add_VN(newRec, recLen):
    return VNMembers.objects.create(**newRec)

##############################################################################
def addDonation(newRec, recLen):
    return Donation.objects.create(**newRec)

##############################################################################
def addWklyDonation(newRec):
    return WeeklyDonation.objects.create(**newRec)
    
##############################################################################
def addBldDonation(newRec, recLen):
    return BldgDonation.objects.create(**newRec)
        
##############################################################################
def addMiscDonation(newRec, recLen):
    return MiscDonation.objects.create(**newRec)

##############################################################################
def addAddrHistory(newRec, recLen):
    return AddrHistory.objects.create(**newRec)

##############################################################################
def addMemberAddrs(newRec, recLen):
    return MemberAddrs.objects.create(**newRec)
    
##############################################################################
def addLitergy(newRec, recLen):
    return Liturgy.objects.create(**newRec)



##############################################################################
##############################################################################
def valDateFmt(dateStr):
    if isinstance(dateStr, datetime.date):
        return dateStr.isoformat()
        #return dateStr.strftime('%Y-%m-%d')

    if isinstance(dateStr, str):
        return None if DateFmt.search(dateStr) is None or dateStr in ('1976-02-04', '1979-02-04', '2001-01-01') else dateStr

##############################################################################
def isValDateFmt(dateStr):
    return dateStr is not None and (isinstance(dateStr, datetime.date) or DateFmt.search(dateStr) is not None)

##############################################################################
def valPhoneNum(phoneNum):
    if PhoneFmt.search(phoneNum) is not None:
        flds, num = [], ''
        for c in phoneNum:
            if c.isdigit():
                num += c
            else:
                if len(num) in (3,4):
                    flds.append(num); num = ''
        else:
            if len(num) in (3,4):
                flds.append(num); num = ''
        if len(flds)>=3:
            return '(%s) %s-%s' % (flds[0], flds[1], flds[2])
        return ''
    return phoneNum


'''
Các nhóm nấu ăn em đặc theo các số
5001: Huệ Quốc
5001: Dung Bình
5003: Yến Mỹ

'''


##############################################################################
def AddFinGrps():
    for i, nm in FINGRPs_CHOICES:
        GrpFinancial.objects.create(deptNm = i)

def AddAdmOpts():
    def addOpt(id, memId, massTime):
        AdmOpts.objects.create(id           = id,
                               massTime     = massTime,
                               admin1       = memId,
                               admin2       = memId+1,
                               admin3       = memId+2,
                               admin4       = None,
                               admin5       = None,
                               )
    for id, memId, massTime in ((100, 1, 19), (200, 4, 35)):
        addOpt(id, memId, massTime)






##############################################################################
##############################################################################
def InitDB():
    for ind, rec in enumerate(Addresses):
        try:
            addFamily(rec)
        except Exception as err:
            print('ERROR: ', err)

    addParish([100]+ParishInfo)

    # Liturgy/Mass schedule
    recCnt = 0
    liturgyInfo = MassSchedule + OtherLiturgy
    for ind, rec in enumerate(liturgyInfo):
        rec[0] = LitergyType_Dct[rec[0]]
        rec.insert(0, ind+1)
        addLitergy(rec)
        recCnt += 1
    print('====>> InitDB Table: Litergy ', recCnt)

##############################################################################
def trtvDB_Models():
    table_info = []
    tables = connection.introspection.table_names()
    modelclasses = connection.introspection.installed_models(tables)
    for model in modelclasses:
        table = model._meta.db_table
        columns = [field.column for field in model._meta.fields]
        table_info.append((table, columns))

        print('trtvDB_Models: ', table, columns)

    return table_info[:]


##############################################################################
def trtvDB_Tables():
    table_info = []
    tables = connection.introspection.table_names()
    seen_models = connection.introspection.installed_models(tables)
    for model in seen_models:
        table = model._meta.db_table
        columns = [field.column for field in model._meta.fields]
        table_info.append((table, columns))

        print('table_info: ', table, columns)

    return table_info[:]

##############################################################################
def trtvDB_M2M_fields(modelclasses):
    table_info = []
    for model in modelclasses:
        for field in model._meta.local_many_to_many:
            if not field.creates_table:
                continue

            table = field.m2m_db_table()
            columns = ['id'] # They always have an id column
            columns.append(field.m2m_column_name())
            columns.append(field.m2m_reverse_name())
            table_info.append((table, columns))

            print('trtvDB_M2M_fields: ', table, columns)


##############################################################################
def trtv_Models():
    table_info = []
    for model in get_models():
        table = model._meta.db_table
        columns = [field.column for field in model._meta.fields]
        table_info.append((table, columns))

        print('trtvDB_Models: ', table, columns)

    return table_info[:]

##############################################################################
def db2xls(xlsHdl):
    for model in get_models():
        exportTable(xlsHdl, model)
    xlsHdl.saveXls()

##############################################################################
def exportTable(xlsHdl, model):
    tblNm = model._meta.db_table
    wkSheet = xlsHdl.addPage(wkSheetNm=tblNm)
    columns = [field.column for field in model._meta.fields]
    wkSheet.append(columns)
    try:
        for rowData in model.objects.all().values():
            wkSheet.append(fetchRec(rowData, columns))
        print('====>> Exported Table: ', tblNm, len(wkSheet.rows))
    except Exception as e:
        print(e)

##############################################################################
def fetchRec(rec, columns):
    values = []
    for fldNm in columns:
        if DateFld.search(fldNm):
            values.append(fmtDater(rec[fldNm]))
        else:
            values.append(rec[fldNm])
    return values[:]

def fmtDater(dateObj):
    "RFC 2822 formatted date; e.g. 'Thu, 21 Dec 2000 16:01:07 +0200'"

    return dateObj.isoformat() if isinstance(dateObj, datetime.date) else ''
    #return dateObj.strftime('%Y-%m-%d') if isinstance(dateObj, datetime.date) else ''

##############################################################################
##############################################################################
def fmtFileNm(selDir, filtType):
    fn = 'ParishDB_{0}_{1}.xlsx'.format(filtType, setTimeStamp())
    return joinPaths('..', 'DMLV_Data', fn), joinPaths(selDir, fn)


#from django.core.management import execute_from_command_line
#from django.core import management
##############################################################################
def BackupDB(selDir):
    cmdLst = ['FLUSH TABLES WITH READ LOCK',
              'UNLOCK TABLES',
              ]
    
    bkpFn = 'ParishDB_Backup_{0}.sql'.format(setTimeStamp())
    mainBk = joinPaths('../DMLV_Data', bkpFn) 
    
    cursor = connection.cursor()
    cursExec(cursor, cmdLst[0])
    cmdRsp = ExecExtCmd(['mysqldump', '-u', 'root', '-padmin', 'ParishDB', '>', mainBk])
    cursExec(cursor, cmdLst[1])
    
    if 'subprocess.CalledProcessError' in cmdRsp:
        return 'Unable to backup database...'
    
    bkpFn = joinPaths(selDir, bkpFn)
    if isFile(bkpFn):  return
    
    copyFile(mainBk, bkpFn)
    
    return bkpFn

##############################################################################
def ReBuildDB(forImport=True):
    dbName = DATABASES['default']['NAME']
    cmdLst = ['DROP DATABASE %s' % dbName,
              'CREATE DATABASE %s CHARACTER SET utf8 COLLATE utf8_unicode_ci' % dbName,
              'use %s' % dbName,
              'SET CHARACTER SET utf8',
              'SET NAMES UTF8 COLLATE utf8_unicode_ci',
              ]
    
    addMsg('Rebuilding MySQL database...')
    
    cursor = connection.cursor()
    for sqlCmd in cmdLst:
        cursor.execute(sqlCmd)
        transaction.commit_unless_managed()
        
    call_command('syncdb', interactive=False)
        
    return cursor


##############################################################################
def cursExec(cursor, sqlCmd):
    cursor.execute(sqlCmd)
    transaction.commit_unless_managed()
     
##############################################################################
def RestoreDB(selSqlBkp):
    '''
    cursor = ReBuildDB(forImport=False)
    
    cursor.execute('source %s' % selSqlBkp)
    
    transaction.commit_unless_managed()
    
    
    Get all tables ( select * from information_schema.tables ) and put in a Cursor... 
    '''
    
    print('====>> I am here RestoreDB...')
    
    cmdRsp = ExecExtCmd(['mysql', '-uroot', '-padmin', 'ParishDB < ', selSqlBkp])
    
    print('cmdRsp: ', cmdRsp)
        
    addMsg('RestoreDB completed: %s' % selSqlBkp)
    
    '''
    To dump all MySQL databases on the system, use the --all-databases shortcut:
    $ mysqldump -u root -p --all-databases > [backupfile.sql]
    Restoring a MySQL Database
    
    Use this method to rebuild a database from scratch:
    $ mysql -u [username] -p [password] [database_to_restore] < [backupfile]
    
    Use this method to import into an existing database (i.e. to restore a database that already exists):
    $ mysqlimport [options] database textfile1
    
    To restore your previously created custback.sql dump back to your 'Customers' MySQL database, you'd use:
    $ mysqlimport -u sadmin -p pass21 Customers custback.sql
    
    '''
    
    return True

##############################################################################
def ExportDB(selDir):
    parishFn, exportFn = fmtFileNm(selDir, 'Export')
    xlsHdl = PyExcel(parishFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to create: %s.' % exportFn)
    
    addMsg('====>> Start exporting %s' % exportFn)
    #xlsHdl.clearXls()
    db2xls(xlsHdl)
    copyFile(parishFn, exportFn)
    return exportFn

##############################################################################
##############################################################################
def ImportDB(selImportFn, impOpt=DB):
    xlsHdl = PyExcel(selImportFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to aceess: %s.' % selImportFn)

    tbllLst = xlsHdl.getAllWkSheets()[:]
    tbllLst.remove('Members_family')
    tbllLst.insert(0, 'Members_family')
    for wkSheetNm in tbllLst:
        dbObj = getDbObj(wkSheetNm, impOpt)
        if dbObj is None: continue
        
        row = 0
        hdrLst = []
        for rowData in xlsHdl.iterXlsRows(wkSheetNm): # it brings a new method: iter_rows()
            row += 1
            if row==1:
                hdrLst = rowData[:]
            else:
                try:
                    recDct = dict(parseRec(rowData, hdrLst))
                    if isinstance(recDct, dict):
                        a = addDbRec(dbObj, recDct) if wkSheetNm!='django_site' else addSiteId(recDct)
                        if impOpt==WEB and wkSheetNm=='Members_family' and recDct.get('regNum')==200:
                            break
                except Exception as err:
                    print('\tERROR: ', err)

        addMsg('====>> Imported Table: %s  %d' % (wkSheetNm, row))

    #addParish()
    #if impOpt==DB:
    #    AddFinGrps()s
    #    AddAdmOpts()
    return selImportFn




##############################################################################
##############################################################################
def ImportDB_Org(selImportFn, impOpt=DB):
    xlsHdl = PyExcel(selImportFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to aceess: %s.' % selImportFn)

    tbllLst = xlsHdl.getAllWkSheets()[:]
    tbllLst.remove('Members_family')
    tbllLst.insert(0, 'Members_family')
    for wkSheetNm in tbllLst:
        impFcn = getImpFnc(wkSheetNm, impOpt)
        if impFcn is None: continue
        recCnt = 0
        for rec in xlsHdl.iterXlsRows(wkSheetNm): # it brings a new method: iter_rows()
            newRec = ['' if v is None else v for v in rec]
            if isValidRec(newRec):
                recCnt += 1
                try:
                    impFcn(newRec, len(newRec))
                except Exception as err:
                    addMsg('\tERROR: %s' % err)

        addMsg('====>> Imported Table: %s  %d' % (wkSheetNm, recCnt))

    addParish()
    if impOpt==DB:
        AddFinGrps()
        AddAdmOpts()
    return selImportFn


##############################################################################
def getDbObj(wkSheetNm, impOpt):
    dbObj, selOpt = DBFuncDct.get(wkSheetNm, (None, None))
    return dbObj if impOpt in (DB, selOpt) else None

##############################################################################
def getImpFnc(wkSheetNm, impOpt):
    impFcn, selOpt = DBFuncDct.get(wkSheetNm, (None, None))
    return impFcn if impOpt in (DB, selOpt) else None

##############################################################################
def isValidRec(newRec):
    return len([v if v is None else v for v in newRec if isValidData(v)]) > 0

##############################################################################
def isValidFam(newRec):
    return len([v if v is None else v for v in newRec if isValidData(v)]) > 0


##############################################################################
# Bldg Don Rec variables
##############################################################################

#Title, Name, RegNum, Pledge, Paid, Ackmnt, PymtDur, PymtOpt = 1, 2, 18, 7, 8, 10, 11, 12
#Cash, Credit, Check, CheckNo, Other, PymtDate = range(24, 30)

(Num, Title, Name, OrgName, OrgAddress, OrgCity, OrgZip, Pledge, DownPymt, Remainder, Ackmnt, PymtDur,
PymtOpt, Phone, Email, InMemoryOf, PymtType, OrgOther, RegNum, Address, City, State,
ZipCode, DupPymtType, Cash,  Check, CheckNo, Credit, Other, PymtDate) = range(30)

PYMT_INT =    {'1': 1,
               'W': 2,
               'M': 3,
               'Q': 4,
               'S': 5,
               'Y': 6,
               1: 1,
               'w': 2,
               'm': 3,
               'q': 4,
               's': 5,
               'y': 6,
              }

##############################################################################
##############################################################################
def ImportBldgDon(selImportFn, initId=1, phase=0):
    xlsHdl = PyExcel(selImportFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to aceess: %s.' % selImportFn)

    wkSheetNm = 'Pledge'
    recCnt = initId        # Start ID for Promissed
    for rec in xlsHdl.iterXlsRows(wkSheetNm): # it brings a new method: iter_rows()
        newRec = ['' if v is None else v for v in rec]
        
        if isValidRec(newRec):
            try:
                if phase==0:
                    myRec = fmtPromisedRec(newRec, recCnt)
                else:
                    myRec = fmtBldgDonRec(newRec, recCnt, phase)
                if myRec is None:  continue
                
                if isinstance(myRec, list):
                    recCnt += 1
                    addMsg('====>> myRec: %s' % str(myRec).encode('utf-8'))
                    addBldDonation(myRec, len(myRec))
            except Exception as e:
                addMsg('ERROR: Exception: %s' % e)
                addMsg('ERROR: ImportBldgDon: %d  %s' % (recCnt, str(newRec).encode('utf-8')))
                
    addMsg('====>> Imported Table: %s  %d' % (wkSheetNm, recCnt))


##############################################################################
##############################################################################
def normAddrs(addrFn):
    appl = NormalizeAddr(addrFn)
    appl.procReq()


##############################################################################
##############################################################################
def SynNames(selImportFn, famNmDct):
    xlsHdl = PyExcel(selImportFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to aceess: %s.' % selImportFn)
    
    xlsWS = xlsHdl.getWkSheet()

    NameCol = Name + 1
    wkSheetNm = 'Pledge'
    recCnt, rId = 0, 1
    for rec in xlsHdl.iterXlsRows(wkSheetNm): # it brings a new method: iter_rows()
        newRec = ['' if v is None else v for v in rec]
        rId += 1
        regNum = getRegNum(newRec)
        if isNotNull(regNum):
            
            names = famNmDct.get(regNum)
            if isNotNull(names):
                recCnt += 1
                xlsWS.cell('%s%d' % (get_column_letter(NameCol), rId)).value = names

    newFileNm = selImportFn[:-5] + '_Updated.xlsx'
    xlsHdl.saveXls(newFileNm)                    
                
    addMsg('====>> Imported Table: %s  %d' % newFileNm, recCnt)

##############################################################################
PrtEnv, PrtLtr, TaxRept = range(1,4)
##############################################################################
def getNmsLst(selImportFn, reqData):
    xlsHdl = PyExcel(selImportFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to aceess: %s.' % selImportFn)
    
    fmtMethod = fmtRecNms if reqData==MemLst else getFamMailing
    print('====>> Processing getNmsLst: ', MemLst, reqData, fmtMethod==fmtRecNms)

    reqDataLst = []
    xlsWS = xlsHdl.getWkSheet()

    wkSheetNm = 'Pledge'
    recCnt = 0
    for rec in xlsHdl.iterXlsRows(wkSheetNm): # it brings a new method: iter_rows()
        newRec = ['' if v is None else v for v in rec]
        
        reqData = fmtMethod(newRec, reqData==PrtEnv)
        if isNotNull(reqData):
            reqDataLst.append(reqData)
            recCnt += 1
            
    addMsg('====>> Total number of donation recoreds:  %d' % recCnt)
    
    return reqDataLst


##############################################################################
def getMailInfo(selImportFn, prtEnvelope=True):
    xlsHdl = PyExcel(selImportFn)
    if xlsHdl.xlsBook is None:
        return addMsg('====>> ERRROR: unable to aceess: %s.' % selImportFn)
    
    reqDataLst = []
    xlsWS = xlsHdl.getWkSheet()

    wkSheetNm = 'Pledge'
    recCnt = 0
    for rec in xlsHdl.iterXlsRows(wkSheetNm): # it brings a new method: iter_rows()
        newRec = ['' if v is None else v for v in rec]
        
        reqData = getFamMailing(newRec, prtEnvelope)
        if isNotNull(reqData):
            reqDataLst.append(reqData)
            recCnt += 1
            
    addMsg('====>> Total number of donation recoreds: %d' % recCnt)
    
    return reqDataLst


##############################################################################
def fmtRecNms(rec, opt):

    title = getTitleAbbr(getTitle(rec))
    names = getName(rec)
    
    if not isinstance(rec[0], int) or isNull(title) or isNull(names):
        print('BAD REC fmtRecNms:', rec[0], isNull(title), isNull(names))
        return None
    
    return '%-8s %s' % (title, names)

    #return (rec[0], ' '.join([title, names]))
   
##############################################################################
def getFamMailing(rec, prtEnvelope):

    title = getTitle(rec)
    names = getName(rec)
    pledge = getPledge(rec)
    downPymt = getDownPymt(rec)
    addr = getAddr1(rec)
    city = getCity(rec)
    state = getState(rec)
    zipCode = getZipCode(rec)
    
    if isNull(title) or isNull(names) or isNull(addr) or isNull(state) or isNull(city) or isNull(pledge):
        print('BAD REC 1', rec[0], isNull(title), isNull(names), pledge, downPymt, addr, city, state, zipCode )
        return None
    
    title = title.capitalize() if prtEnvelope else title
    try:

        return {'Name': '%s %s' % (title, names),
                'Address': addr,
                'CityState': '%s, %s %d' % (city, state, zipCode),
                'Donation': fmtNameFld(pledge, downPymt),
                }
    
    except:
        print('BAD REC 2', rec[0], isNull(title), isNull(names), pledge, downPymt, city, state, zipCode )
         
   
            
    '''
    Addr_Maps = {'Name': '_NAME_',
     'Address': '_Address_',
     'CityState': '_City_State_ZipCode_',
     'Donation': '_DONATE_'
     }
    
    '''

##############################################################################
def fmtNameFld(pledge, downPymt):
    if isNull(downPymt):
        return 'đã nhận được thẻ anh chị hứa $%.2f (USD) sẽ đóng góp' % pledge
    else:
        return 'đã nhận được số tiền quý vị đóng góp là $%.2f (USD) trên tổng số $%.2f (USD) qúi vị hứa' % (downPymt, pledge)
    
##############################################################################
def fmtPromisedRec(rec, recId):
    DonRec = ['' for n in range(15)]
    DonRec[0]= recId
    DonRec[1]= getRegNum(rec)               # RegNum
    DonRec[2]= getPledge(rec)            # promissed
    
    if None in DonRec[:3]:
        addMsg('====>> Missing Family RegNum or Promissed Amount: %s' % str(rec).encode('utf-8'))
        return 
    
    DonRec[3]= getPymtOpt(rec)              # pymtOption
    DonRec[4]= getPymtDur(rec)              # pymtDuration
    DonRec[5]= getAck(rec)                  # acknowledge
    DonRec[7]= getPymntDate(rec)            # donateDate 1st Pymt
    
    DonRec[6]= 'L.M. Trần Văn Bản'          # adminPersonnel
    DonRec[14]= 'Hứa cho quỹ xây dựng'      # notes
    
    return DonRec[:]

##############################################################################
def fmtBldgDonRec(rec, recId, phase):
    offset = 6 * (phase-1)

    DonRec = ['' for n in range(15)]
    DonRec[0]= recId
    DonRec[1]= getRegNum(rec)             # RegNum
    DonRec[2]= getPledge(rec)             # promissed
    
    if None in DonRec[:3]:
        addMsg('====>> Missing Family RegNum or Promissed Amount: %s' % str(rec).encode('utf-8'))
        return 
    
    DonRec[7]= getPymntDate(rec, phase=phase, offset=offset)            # donateDate 1st Pymt
    
    DonRec[8]= getCashAmount(rec, offset=offset)           # cashAmount
    DonRec[9]= getCreditAmount(rec, offset=offset)         # creditAmount
    
    DonRec[10]= getCheckAmount(rec, offset=offset)          # checkAmount
    DonRec[11]= getCheckNum(rec, offset=offset)             # checkNumber

    DonRec[12]= getOtherAmount(rec, offset=offset)          # otherAmount
    
    DonRec[13]= calTotalDon([DonRec[8], DonRec[9], DonRec[10], DonRec[12]])          # total
    
    if DonRec[13]==0:  return
    
    DonRec[6]= 'L.M. Trần Văn Bản'          # adminPersonnel
    DonRec[14]= 'Góp quỹ xây dựng %2d' % phase          # notes
    
    return DonRec[:]
    
##############################################################################
##############################################################################
def verStrFld(fld):
    return fld if (isinstance(fld, str) and len(fld)>0) else None

def verIntFld(fld):
    return int(fld) if isinstance(fld, int) or (isinstance(fld, str) and fld.isdigit()) else None

def verDecFld(fld):
    if isinstance(fld, str) and fld.isdigit():
        return Decimal(fld)
    return fld if (isinstance(fld, (Decimal, int)) and fld>0) else None

def calTotalDon(donFlds):
    total = 0
    for v in donFlds:
        total += v if v is not None else 0
    return total

##############################################################################
##############################################################################
def getRegNum(rec, id=RegNum):
    return verIntFld(rec[id])

def getTitle(rec, id=Title):
    return verStrFld(rec[id])

def getName(rec, id=Name):
    return verStrFld(rec[id])

def getPledge(rec, id=Pledge):
    return verDecFld(rec[id])

def getDownPymt(rec, id=DownPymt):
    return verDecFld(rec[id])

def getAck(rec, id=Ackmnt):
    return rec[id] if rec[id] in ('Y', 'y', 'N', 'n') else None

def getPymtDur(rec, id=PymtDur):
    return rec[id] if isinstance(rec[id], int) else int(rec[id]) if (isinstance(rec[id], str) and rec[id].isdigit()) else 5

def getPymtOpt(rec, id=PymtOpt):
    return PYMT_INT[rec[id]] if rec[id] in PYMT_INT else 6

def getCashAmount(rec, id=Cash, offset=0):
    selId = id + offset
    return verDecFld(rec[selId])

def getCreditAmount(rec, id=Credit, offset=0):
    selId = id + offset
    return verDecFld(rec[selId])

def getCheckAmount(rec, id=Check, offset=0):
    selId = id + offset
    return verDecFld(rec[selId])

def getCheckNum(rec, id=CheckNo, offset=0):
    selId = id + offset
    return verIntFld(rec[selId])

def getOtherAmount(rec, id=Other, offset=0):
    selId = id + offset
    return verDecFld(rec[selId])

def getPymntDate(rec, id=PymtDate, phase=0, offset=0):
    if phase==0: return '2015-01-11'
    
    selId = id + offset
    if isinstance(rec[selId], datetime.date):
        return rec[selId].isoformat()
        #return rec[selId].strftime('%Y-%m-%d')

    if isinstance(rec[selId], str) and DateFmt.search(rec[selId]) is not None:
        return rec[selId]

    return '2015-%02d-15' % phase

def getAddr1(rec, id=Address):
    return verStrFld(rec[id])

def getCity(rec, id=City):
    return verStrFld(rec[id])

def getState(rec, id=State):
    return verStrFld(rec[id])

def getZipCode(rec, id=ZipCode):
    return verIntFld(rec[id])

    
def getTitleAbbr(title):
    if isNull(title):  return
    flds = [fld.capitalize()[0] for fld in [fld.strip() for fld in title.split(' ')] if len(fld)>0]
    if len(flds)>0:
        flds.append('')
        return '.'.join(flds)

    
    
'''    
    return BldgDonation.objects.create(id               = int(newRec[0]),
                                      familyDonation_id = int(newRec[1]),
                                      promissed         = newRec[2] if isValidData(newRec[2]) else None,
                                      pymtOption        = newRec[3] if isValidData(newRec[3]) else None,
                                      pymtDuration      = newRec[4] if isValidData(newRec[4]) else None,
                                      adminPersonnel    = newRec[5],
                                      donateDate        = valDateFmt(newRec[6]),
                                      cashAmount        = newRec[7] if isValidData(newRec[7]) else None,
                                      creditAmount      = newRec[8] if isValidData(newRec[8]) else None,
                                      checkAmount       = newRec[9] if isValidData(newRec[9]) else None,
                                      checkNumber       = newRec[10] if isValidData(newRec[10]) else None,
                                      otherAmount       = newRec[11] if isValidData(newRec[11]) else None,
                                      total             = newRec[12] if isValidData(newRec[12]) else None,
                                      notes             = newRec[13] if isValidData(newRec[13]) else None,
                                     )
'''

##############################################################################
##############################################################################
DBFuncDct = {'django_site': (Site.objects, WEB),
             'Members_dbid': (dbID.objects, DB),
             'Members_family': (Family.objects, WEB),
             'Members_member': (Member.objects, DB),

             'Members_addrhistory': (AddrHistory.objects, DB),
             'Members_memberaddrs': (MemberAddrs.objects, DB),
             
             'Parish_parish': (Parish.objects, WEB),
             'Ministry_baptism': (Baptism.objects, DB),
             'Ministry_firstcommunion': (FirstCommunion.objects, DB),
             'Ministry_confirmation': (Confirmation.objects, DB),
             'Ministry_holymatrimony': (HolyMatrimony.objects, DB),
             'Ministry_holyorders': (HolyOrders.objects, DB),
             'Ministry_deceased': (Deceased.objects, DB),
             'Ministry_ministrymembers': (MinistryMembers.objects, DB),

             'Ministry_liturgy': (Liturgy.objects, WEB),

             #'Ministry_association',
             #'Ministry_membership',

             'Finance_donation': (Donation.objects, DB),
             'Finance_weeklydonation': (WeeklyDonation.objects, DB),
             'Finance_bldgdonation': (BldgDonation.objects, DB),
             'Finance_miscdonation': (MiscDonation.objects, DB),

             'TNTT_tnttmembers': (TNTTMembers.objects, DB),
             'Catechism_ccdmembers': (CCDMembers.objects, DB),
             'VietNamese_vnmembers': (VNMembers.objects, DB),
             'Bible_bibleverses': (BibleVerses.objects, WEB),
             'Bible_biblesel': (BibleSel.objects, WEB),
             }



# End of views
