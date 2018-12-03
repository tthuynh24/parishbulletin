#!/usr/bin/env python
##############################################################################
# Module: PyUtils.py                                                           #
##############################################################################
from __future__ import absolute_import
import os
import sys
import platform
import re
import csv
import glob
import shutil
import pickle
import binascii
import string
import struct
import itertools
import traceback
import ipaddress
import unicodedata
from zipfile import ZipFile

from decimal import Decimal

from copy import deepcopy
import socket
from filecmp import cmp as cmpFile
from sys import _getframe as GetModInfo
import subprocess
from time import sleep
from datetime import date

try:
    from Lib.AppShared import *
except:
    pass

FMT = re.compile('\%\(\S+\)((\d*d)|(s)|(\d*x)|(\d*X))')
ParmNm = re.compile('\%\((\S+)\)\d{0,2}[d,s,x,X]')
PythonFmt = re.compile('%\((\w+(-\w+)*)\)(\d{0,2}[dsxX])')
SpecRange = re.compile('\[([R,L]:[a-f0-9A-F]+(,\d+)*)\]')
SeagullFmt = re.compile('(\$\(\w+([-_]\w+)*\))')
DollarChars = re.compile('(\$)|(\.00)')

BadFNChars = re.compile('[/\?<>\\\:\*(){}[\]\,\;\|\+\-&$%^#@!]', re.I)

sRange, eRange, Inc = range(3)
WhiteSpace = re.compile('\s+')
HEX_FMT = re.compile('\)\d*x', re.I)
Fedora = re.compile('(FC\d+)', re.I)

ABS_NAME, BASE_NAME, FILE_NAME = range(3)

Reserved_NW_Octets = tuple(range(11)) + (127,169,172,192)
Reserved_NW_Octets_Str = tuple([str(nw) for nw in Reserved_NW_Octets])

###########################################################################
def absPath(selPath):
    return os.path.abspath(selPath)

###########################################################################
def addFilePrefx(scrFileNm, suffix):
    nmFlds = scrFileNm.rsplit('_', 1)
    nmFlds.insert(1,suffix)
    return '_'.join(nmFlds)

###########################################################################
def addToDct(dct, key, val, uniq=True):
    '''
    Add the given dictionary entries (key = "key", value = "val") to the
    specified dictionary "dct".  If "uniq" is set to True, the entries is
    only added if it is not already exist in the "dct" else is will be added
    as an update.
    '''
    addItem = not (uniq and key in dct)
    if addItem:
        dct.update({key:val})
    return addItem

###########################################################################
def setFamBl(dct, key, val):
    addToDct(dct, key, val)
    return dct[key]

###########################################################################
def addToLst(lst, item, uniq=True, ignoreNone=False):
    '''
    Add the given "item" to specified list "lst".  If "uniq" is set to True,
    as Default, item is only added if it is not already exist in the "lst" else
    add to the list regardless.
    '''
    if ignoreNone and isNull(item): return
    addItem = not (uniq and item in lst)
    if addItem:
        lst.append(item)
    return addItem

###########################################################################
def asciiToHex(byteStr):
    return ''.join([ "%02X " % ord(x) for x in byteStr ]).strip()

###########################################################################
def backUpFile(fileNm):
    if isFile(fileNm):
        copyFile(os.path.realpath(fileNm), os.path.realpath(fileNm)+".bak")

###########################################################################
def compress(srcFn, zipFile, psswd):
    ZipFile.compress(srcFn, zipFile, psswd, 1)

###########################################################################
def unCompress(zipFile, psswd, dstPath):
    with ZipFile(zipFile) as zf:
        zf.extractall(dstPath, pwd=psswd)

###########################################################################
def chgMode(fileSpec, fMode=755):
    os.chmod(fileSpec, int('0o%3d' % fMode, 8))

###########################################################################
def closeFile(fhdl):
    '''
    Close the file handle "fhdl" in a grateful way.
    '''
    if fhdl:
        fhdl.flush()
        fhdl.close()
        fhdl = None

###########################################################################
def cmdTostr(cmd):
    return ' '.join(cmd) if isValid(cmd, (tuple, list)) else str(cmd)

###########################################################################
def cmpFiles(file1, file2):
    return cmpFile(file1, file2) == 1

###########################################################################
def compileRegex(specMasks):
    if specMasks:
        if isValid(specMasks, str):
            specMasks = str(specMasks)
        elif isValid(specMasks, (list, tuple)):
            masks = []
            for mask in specMasks:
                masks.append('(%s)' % mask)
            specMasks = '|'.join(masks)
        else:
            specMasks = None
    return re.compile(specMasks) if isNotNull(specMasks) else None

###########################################################################
def copyDir(frmDir, toDir):
    '''
    Copy the given input dir "frmDir" to the given output dir spec "toDir".
    '''
    try:
        shutil.copytree(frmDir, toDir)
        return True
    except:
        return False

###########################################################################
def copyFile(frmFile, toFile):
    '''
    Copy the given input file "frmFile" to the given output file spec "toFile".
    '''
    try:
        shutil.copy(frmFile, toFile)
        return True
    except:
        return False

###########################################################################
def copyObj(srcObj):
    ''' Return a duplicate copy of the given object "srcObj" '''
    return deepcopy(srcObj)

##########################################################################
def fmtVNDate(selDate=None):
    if isNull(selDate):  selDate = date.today()
    dateStr = 'Ngày {0} tháng {1} năm {2}'
    if isinstance(selDate, date):
        return dateStr.format(selDate.day, selDate.month, selDate.year)
    return dateStr.format(selDate.day(), selDate.month(), selDate.year())

##########################################################################
def fmtDate(selDate=None):
    if isNull(selDate):  selDate = date.today()
    if isinstance(selDate, date):
        return dateStr.format(selDate.day, selDate.month, selDate.year)
    return dateStr.format(selDate.day(), selDate.month(), selDate.year())

###################################################################################
def fmtFileNm(fileSpec, srcKey, dstKey):
    flds = fileSpec.split('_%s' % srcKey)
    return ('_%s' % dstKey).join(flds)

###########################################################################
def decodeUtf8(byteStr, asIs=False):
    if asIs:
        return bytes(byteStr).decode("utf-8", 'ignore')
    return bytes(byteStr).decode("utf-8", 'ignore').strip()

###########################################################################
def delDir(rmItem):
    '''
    delete the given dir "rmItem".
    '''
    try:
        if isDir(rmItem):
            shutil.rmtree(rmItem)
        return True
    except:
        return False

###########################################################################
def delFile(rmItem):
    '''
    delete the given file "rmItem".
    '''
    try:
        if isFile(rmItem):
            os.remove(rmItem)
        elif isDir(rmItem):
            shutil.rmtree(rmItem)
        wait4ReqCmplt(isGone, rmItem, 10000)
        return True
    except:
        return False

###########################################################################
def wait4ReqCmplt(reqFun, reqTarget, waitTime):
    sTime = time.time()
    while not reqFun(reqTarget):
        time.sleep(0.25)
        if (time.time()-sTime) >= waitTime: break

###########################################################################
def encodeFloat32(floatVal, endian="!"):
    return '0x%s' % decodeUtf8(binascii.hexlify(struct.pack('%sf' % endian, floatVal)), asIs=True)

###########################################################################
def encodeFloat64(floatVal, endian="!"):
    return '0x%s' % decodeUtf8(binascii.hexlify(struct.pack('%sd' % endian, floatVal)), asIs=True)

###########################################################################
def encodeUtf8(byteStr, asIs=False):
    if asIs:
        return byteStr.encode("utf-8")
    return byteStr.encode("utf-8").strip()

###########################################################################
def evenItemsFrLst(selLst):
    return selLst[::2]


def ExecExtCmd(extCmd):
    try:
        return decodeUtf8(subprocess.check_output(extCmd, stderr=subprocess.STDOUT, shell=True))
    except Exception as cmdError:
        print('Unable to execute your command: "%s" - "%s".' % (' '.join(extCmd), str(cmdError)))

###########################################################################
def execExtCmd_Org(extCmd, stderr2stdout=False, reportErr=True):
    if extCmd is None: return 'NONE'
    errMsg = None
    
    try:
        stderr = subprocess.STDOUT if stderr2stdout else None
        return decodeUtf8(subprocess.check_output(extCmd, stderr=stderr))
    except subprocess.CalledProcessError as cmdError:
        if stderr2stdout:
            return decodeUtf8(cmdError.output)
        errMsg = 'Unable to execute your command: "%s" - "%s".' % (cmdTostr(extCmd), str(cmdError))
    except OSError as cmdError:
        errMsg = 'The current OS does not support command: "%s" - "%s".' % (cmdTostr(extCmd), str(cmdError))
    except Exception as cmdError:
        errMsg = 'Unable to execute your command: "%s" - "%s".' % (cmdTostr(extCmd), str(cmdError))
    getTrace()

    if reportErr and isNotNull(errMsg):
        reportStat(errMsg)
    return 'NONE'

###########################################################################
def execSysCmd(sysCmd):
    try:
        return os.system(sysCmd)
    except:
        reportStat('Unable to execute your command: "%s".' % cmdTostr(sysCmd))
        getTrace()

###########################################################################
def fileToList(inFn):
    '''
    Read all lines in the given file spec "inFn" and return the list of
    stripped leading and trailing white-spaces.
    '''
    return [line.strip() for line in open(normpath(inFn), 'r').readlines()]

###########################################################################
def fmtIpAddr(selAddr):
    return '[%s]' % selAddr if ':' in selAddr else selAddr

###########################################################################
def getApplPath(xOS):
    if xOS:
        maxLinks = 255
        try:
            applPath = __file__
        except NameError:
            applPath = '/proc/%s/exe' % os.getpid()

        while os.path.islink(applPath):
            applPath = os.readlink(applPath)
            maxLinks -= 1
            if maxLinks==0:
                reportStat('ERROR: Incorrect synbolic link of %s' % applPath)
                sys.exit(1)
        return getDirname(os.path.abspath(applPath))
    return os.getcwd()

###########################################################################
def getCurrPath():
    return os.getcwd()

###########################################################################
def getBasename(fileSpec):
    return os.path.basename(fileSpec) if isValid(fileSpec) else None

###########################################################################
def getCommonItems(list1, list2, reqType=list):
    '''
    Return a list of all common items from both list1 and list2.
    '''
    return reqType(set(list1).intersection(set(list2)))

###########################################################################
def getDiffItems(list1, list2, reqType=list):
    '''
    Return a list of items that is only exist in list1 but not exist in list2.
    '''
    return reqType(set(list1).difference(set(list2)))

###########################################################################
def getDirname(fileSpec):
    return os.path.dirname(fileSpec) if isValid(fileSpec) else None

###########################################################################
def getFedoraRel():
    srchRslt = Fedora.search(platform.platform())
    return 'NT' if isNull(srchRslt) else srchRslt.group()

###########################################################################
def getFileName(fileSpec):
    baseNm = getBasename(fileSpec)
    fileNm = splitFileExt(baseNm)[0]
    return baseNm, fileNm

###########################################################################
def getFileSpec(myFn, dfltExt='All files (*.*)'):
    if isValid(myFn, str):
        baseFn, fileExt  = splitFileExt(myFn)
        if isValid(fileExt):
            return os.path.dirname(myFn), 'Select files (*%s)' % fileExt, fileExt
        return os.path.dirname(myFn), dfltExt, fileExt
    return getPath('selPath'), dfltExt, fileExt

###########################################################################
def getFilesPerTypes(filePath, fnMask, maxDepth=1, filePart=ABS_NAME):
    if filePart==ABS_NAME:
        return [normpath(p) for p in globFile(filePath, fnMask, maxDepth)]
    if filePart==BASE_NAME:
        return [getBasename(p) for p in globFile(filePath, fnMask, maxDepth)]
    if filePart==FILE_NAME:
        return [getFileName(p)[-1] for p in globFile(filePath, fnMask, maxDepth)]
    return []

###########################################################################
def getNewestFile(filePath, fnMask):
    try:
        return normpath(max([p for p in globFile(filePath, fnMask)], key=os.path.getctime))
    except:
        return None

###########################################################################
def getFirstDataLine(inFn, asIs=False):
    '''
    Find and return the first data line from the given file spec "inFn".
    The line is returned as read if "asIs" is set to True otherwise it is
    returned with stripped leading and trailing white-spaces.
    '''
    for i, line in enumerate(open(inFn, 'rb')):
        line = line.decode()
        data = line.strip()
        if data:
            return line if asIs else data

###########################################################################
def getFmtParmNms(srchStr):
    fmtVars = []
    for p in srchStr.split('%')[1:]:
        srchRslt = ParmNm.search('%' + p)
        if srchRslt:
            fmtVars.append( srchRslt.group(srchRslt.lastindex) )
    return fmtVars

###########################################################################
def getFmtVar(specFmt):
    return specFmt.split('%(', 1)[-1].split(')',1)

###########################################################################
def getIntDollar(moneyVal):
    if isinstance(moneyVal, (Decimal, int)):  return moneyVal
    return Decimal(DollarChars.sub('', moneyVal))

###########################################################################
def getLineCounts(selFn):
    cnt = 0
    try:
        with open(selFn, 'r') as fileHdl:
            for cnt, rec in enumerate(fileHdl): pass
            return cnt
    except:
        return -1
    return cnt

###########################################################################
def getPklData(pklFn):
    ''' UnPickle the given pickled file spec "pklFn" and return the pickled data structure '''
    if not isFile(pklFn):
        return None

    with open(pklFn, 'rb') as f:
        dataHdl = pickle.load(f)
        f.close()
        return dataHdl

###########################################################################
def getPythonPath():
    for p in sys.path:
        if p.endswith('site-packages'):
            return getDirname(p)
    return None

###########################################################################
def getRelPath(dirNn, dfltDir):
    for dotPath in ('./', '../'):
        selDir = os.path.join(dotPath, dirNn)
        if os.path.isdir(selDir):
            return selDir
    return dfltDir

###########################################################################
def getSystemIP():
    try:
        ipInfo = socket.gethostbyname_ex(socket.gethostname())
        hostNm = ipInfo[0]
        for ip in ipInfo[2]:
            ipFlds = list(map(int, splitNstrip(ip, sepChars='.')))
            if len(ipFlds) == 4 and ipFlds[0] != 127 and ipFlds[-1] != 0:
                return hostNm, ip
    except:
        pass
    return None, None

###########################################################################
def getTrace(retEMsg=False):
    try:
        traceRslt = traceback.format_exc()
        eMsg = traceRslt.strip().split('\n')[-1]
    except:
        traceRslt = eMsg = 'Sorry - Unable to get trace...'
    reptMsg(traceRslt, CRITICAL)

    if retEMsg is True: return eMsg
    if traceRslt!=eMsg:
        reptMsg(eMsg, ERROR)

###########################################################################
def getUnionItems(list1, list2, reqType=list):
    '''
    Return a list of all unique items of the combined of list1 and list2.
    '''
    return reqType(set(list1).union(set(list2)))

###########################################################################
def getUserHome():
    try:
        if isNixOs():
            return os.environ.get('HOME', '')
        else:
            return joinPaths(os.environ.get('HOMEDRIVE'), os.environ.get('HOMEPATH'))
    except:
        return os.path.expanduser('~')

###########################################################################
def globFile(filePath, fnMask, maxDepth=1):
    someFiles = []
    for d in range(maxDepth):
        maxGlob = "/".join( "*" * d )
        globDir = joinPaths(filePath, maxGlob, fnMask)
        someFiles.extend(glob.glob(globDir))
    return someFiles

###########################################################################
def hexEncodeIP(myIp, prefix=True):
    hexFmtIp, ipType = [], IPV4
    try:
        if isIpv6Addr(myIp):
            ipType = IPV6
            hexFmtIp = ['0x0002' if prefix else ''] + [octetVal.rjust(4, '0') for octetVal in parseIpv6Flds(myIp)]
        else:
            ipOctets = myIp if isValid(myIp, list, chkVal=False) else myIp.split('.')
            hexFmtIp = ['0x0001' if prefix else ''] + ['%02X' % int(octetVal) for octetVal in ipOctets]
    except:
        reportStat('Invalid IP specification: %s' % myIp)
    return ''.join(hexFmtIp), ipType

###########################################################################
def hexToAscii(hexStr):
    bytes = []
    try:
        for val in strToBytes(hexStr):
            bytes.append(decodeUtf8(binascii.a2b_hex(val), asIs=True))
    except:
        return 'Error: Unable to convert the input HEX string.'
    return ''.join(bytes)

###########################################################################
def ip4To6(ipv4Addr):
    if isValid(ipv4Addr, str):
        ipOctets = ['%02X' % int(octet) for octet in splitNstrip(ipv4Addr, sepChars='.')]
    else:
        ipOctets = ['%02X' % int(octet) for octet in ipv4Addr]

    if len(ipOctets)==4:
        hexStr = ''.join(ipOctets)
        return '2002:%s:%s::1' % (hexStr[:4], hexStr[4:])

###########################################################################
def isAbsPath(selPath):
    return os.path.isabs(selPath)

###########################################################################
def isAccessible(pathname):
    return os.access(pathname, os.R_OK)

###########################################################################
def isApplInstalled(applNm):
    cmdRsp = execExtCmd(['rpm', '-q', applNm])
    if cmdRsp.startswith(applNm):
        return cmdRsp
    return None


###########################################################################
def isData(selData):
    return selData is not None and len(selData)>0

###########################################################################
def isDecStr(myStr):
    return isValid(myStr, int, chkVal=False) or myStr.isdecimal()

###########################################################################
def isDir(selDir):
    return os.path.isdir(selDir) if isValid(selDir) else False

###########################################################################
def isExist(fileSpec):
    return os.path.exists(fileSpec) if isValid(fileSpec) else False

###########################################################################
def isGone(fileSpec):
    return not isExist(fileSpec)

###########################################################################
def isFile(selFn):
    return os.path.isfile(selFn) if isValid(selFn) else False

###########################################################################
def isFmtParms(specParm):
    return isNotNull(PythonFmt.search(specParm))

###########################################################################
def isHexFmt(fmtSpec):
    return isNotNull(HEX_FMT.search(fmtSpec))

###########################################################################
def isHexStr(myStr):
    return False not in map(ishexdigit, myStr)

###########################################################################
def isIpv4Addr(myIp):
    return isValid(ipaddress.ip_address(myIp), ipaddress.IPv4Address)

###########################################################################
def isIpv4_NW1(ipFirstOctet):
    if isValid(ipFirstOctet, str) and len(ipFirstOctet)<4 and ipFirstOctet.isdigit():
        return ipFirstOctet not in Reserved_NW_Octets_Str
    if isValid(ipFirstOctet, int) and ipFirstOctet<256:
        return ipFirstOctet not in Reserved_NW_Octets
    return not isValid(ipFirstOctet)

###########################################################################
def isIpv6Addr(myIp):
    return isValid(ipaddress.ip_address(myIp), ipaddress.IPv6Address)

###########################################################################
def isLink(selLnk):
    return os.path.islink(selLnk) if isValid(selLnk) else False

###########################################################################
def isNixOs():
    return platform.system() in ('Linux', 'Unix')

###########################################################################
def isNotNull(selData):
    return selData is not None

###########################################################################
def isNull(selData):
    return selData is None

###########################################################################
def isOsVersion(osVer):
    return osVer in platform.platform().split('.')

###########################################################################
def isReachableHost(myHost):
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    hostOk = True
    try:
        sock.connect((myHost, 80))
    except socket.error:
        hostOk = False
    sock.close()
    return hostOk

###########################################################################
def isSameLst(list1, list2):
    return set(list1)==set(list2)

###########################################################################
def isUnicode(dataStr):
    try:
        dataStr.encode('ascii')
    except:
        return True
    return False

###########################################################################
def isNotUnicode(dataStr):
    try:
        dataStr.encode('ascii')
    except:
        return False
    return True

###########################################################################
def isValCmdAbbr(cmdAbbr):
    return len(cmdAbbr.split('-')[0])==3

###########################################################################
def isValid(selData, targetType=None, chkVal=True):
    if isNull(targetType):
        return isValidData(selData)
    return isValidData(selData) and isinstance(selData, targetType) if chkVal else isinstance(selData, targetType)

###########################################################################
def isValidData(selData):
    dataType = type(selData)
    if dataType in (int, float): return selData>0
    if dataType in (str, list, tuple, dict):
        if isinstance(selData, str):
            selData = selData.strip()
        return len(selData)>0
    return isNotNull(selData)

###########################################################################
def isValidFmtSpec(specFmt):
    for grp in specFmt.split('%')[1:]:
        rslt = PythonFmt.search('%' + grp)
        if isNull(rslt):
            return False
        else:
            fmt = rslt.group(rslt.lastindex)
            reqLen = len(fmt)
            if not ((reqLen==1 and fmt in 'dsxX') or (reqLen==3 and fmt[-1] in 'dxX')):
                return False

    for grp in specFmt.split('$')[1:]:
        rslt = SeagullFmt.search('$' + grp)
        if isNull(rslt):
            return False

    return True

###########################################################################
def is_number(s):
    try:
        float(s)
        return True
    except:
        pass
    try:
        unicodedata.numeric(s)
        return True
    except:
        pass

    return False

###########################################################################
def ishexdigit( c ):
    return c in string.hexdigits

###########################################################################
def iterCsv(csvFn, **parms):
    for row in openCsv(csvFn, fileMode='r', **parms):
        yield row

###########################################################################
def iterFile(inFn, asIs=False):
    '''
    Iterate the given file spec "inFn" and return the one line at a time.
    each line is returned as read if "asIs" is set to True otherwise it is
    returned with stripped leading and trailing white-spaces.
    '''
    for line in open(inFn, 'rb'):
        line = line.decode('utf8', 'ignore')
        yield line if asIs else line.strip()

###########################################################################
def iterRange(rangeSpec):
    '''
    Iterate the given range "rangeSpec" and return one item at a time.
    range is specified as (start-value, end-value, increment)
    '''
    for val in range(rangeSpec[sRange], rangeSpec[eRange], rangeSpec[Inc]):
        yield val

###########################################################################
def iterSequence(seqSpec):
    '''
    Iterate the given sequence "seqSpec" and return the one item at a time.
    seqSpec is specified as list or tuple of items
    '''
    for val in seqSpec:
        yield val

###########################################################################
def iterWhile(startVal, inc=1):
    val = startVal
    while True:
        rslt = val; val += inc
        yield rslt

###########################################################################
def joinPaths(*pathFlds):
    '''
    Join all given parts of the given path using linux path separator "/"
    and return the complete jointed path.
    '''
    return normpath(os.path.join(*pathFlds))

###########################################################################
def listToFile(dataLst, toFile):
    ''' Save the list of data line in "dataLst" to the given file spec "toFile" '''
    if dataLst:
        fhdl = open(toFile, 'w')
        [fhdl.writelines(data.rstrip() + '\n') for data in dataLst]
        closeFile(fhdl)

###########################################################################
def mergeLst(mergeLst, srcLst):
    for item in srcLst:
        addToLst(mergeLst, item)
    return mergeLst[:]

###########################################################################
def moveFile(frmFile, toDir):
    try:
        shutil.move(frmFile, toDir)
        return toDir
    except Exception as Err:
        print('Unable to move Dir to: "%s" - "%s".' % (toDir, str(Err)))

###########################################################################
def myStrip(data, lowerCase=False):
    if lowerCase:
        return data.strip().lower()
    return data.strip()

###########################################################################
def nCycle(iterable, n=1):
    '''
    Cycle the given sequence "iterable" and return the one item at a time
    and repeated if needed.
    '''
    if len(iterable)==0:
        return None
    if n == 1:
        return itertools.cycle(iterable)
    return itertools.cycle(itertools.chain(*zip(*([iterable] * n))))

###########################################################################
def nCycleRange(rangeSpec, n=1):
    '''
    Cycle the given range "rangeSpec" and return the one item at a time
    and repeated if needed.
    range is specified as (start-value, end-value, increment)
    '''
    if n == 1:
        return itertools.cycle(range(rangeSpec[sRange], rangeSpec[eRange], rangeSpec[Inc]))
    return itertools.cycle(itertools.chain(*zip(*([range(rangeSpec[sRange], rangeSpec[eRange], rangeSpec[Inc])] * n))))

###########################################################################
def newObj():
    return lambda cls: cls.__name__

###########################################################################
def normpath(selPath):
    return os.path.normpath(selPath)

###########################################################################
def oddItemsFrLst(selLst):
    return selLst[1:][::2]

###########################################################################
def openCsv(csvFn, fileMode='r', **parms):
    try:
        csvHdl = csv.reader if 'r' in fileMode else csv.writer
        return csvHdl(open(csvFn, fileMode), dialect=csv.unix_dialect, **parms)
    except:
        return None

###########################################################################
def parseFmtParms(specFmt):
    fmtDct = {}
    while isNotNull(PythonFmt.search(specFmt)):
        fmtVar, specFmt = getFmtVar(specFmt)
        fmtDct[fmtVar] = fmtVar
    return fmtDct

###########################################################################
def parseIpv6Flds(ipv6Addr):
    if '::' in ipv6Addr:
        numFillOctets = 8 - ipv6Addr.count(':')
        ipv6Addr = ipv6Addr.replace('::', ':%s' % ('0:' * numFillOctets))
    return splitNstrip(ipv6Addr, sepChars=':')

###########################################################################
def parseList(dataStr, sepChr=',', func=tuple):
    return func([fld.strip() for fld in dataStr.split(sepChr)])

###########################################################################
def pause(pauseSecs=1, pauseInc=0.1):
    while pauseSecs > 0:
        sleep(pauseInc)
        pauseSecs -= pauseInc

###########################################################################
def pickleData(data2Pkl, pklFn):
    ''' Pickle data structure "data2Pkl" in the given file spec "pklFn" '''
    with open(pklFn, 'wb') as f:
        # Pickle the data dictionary using the highest protocol available.
        pickle.dump(data2Pkl, f, pickle.HIGHEST_PROTOCOL)
        f.close()

###########################################################################
def printDebug():
    '''
    Print the debug info when instance variable DebugOn is set.
    '''
    newLine = '\n'
    print(('%s\t====>> Process being executed: %-21sat %-6dcaller %s.' % \
          (newLine,
           GetModInfo(1).f_code.co_name,
           GetModInfo(1).f_lineno,
           GetModInfo(2).f_code.co_name)))

###########################################################################
def rReadLines(fileNm, numLines):
    def rReadBlock(file, blkSize):
        file.seek(0, 2)
        file.seek(file.tell()-blkSize)
        return file.readlines()

    fileSize = os.stat(fileNm).st_size
    file = open(fileNm, 'r')
    lines, blkSize = [], 4096
    while blkSize < fileSize:
        lines = rReadBlock(file, blkSize)
        if len(lines)>numLines: break
        blkSize *= 2
    return lines[:]

###########################################################################
def readCsv(csvFn):
    csvData = []
    with open(csvFn, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            csvData.append(stripDataLst(row))
    return csvData[:]

###########################################################################
def remove_viet_accents(str):
    ''' Helper function: Remove Vietnamese accent for string '''
    nkfd_form = unicodedata.normalize('NFKD', str)
    return u"".join([c for c in nkfd_form if not unicodedata.combining(c)]).replace('\u0111','d').replace('\u0110', 'D')

###########################################################################
def replaceSpace(data, newChr=''):
    return WhiteSpace.sub(newChr, data)

###########################################################################
def reptMsg(msg, statLvl):
    try:
        reportStat(msg, statLvl)
    except:
        try:
            logMsg(msg, statLvl)
        except: pass

###########################################################################
def rmFrmDct(dct, key):
    rmItem = key in dct
    if rmItem: del dct[key]
    return rmItem

###########################################################################
def rmFrmLst(lst, item):
    rmItem = item in lst
    if rmItem: lst.remove(item)
    return rmItem

###########################################################################
def saveToFile(data, toFile, bkp=True):
    ''' Save the input data to the given file spec "toFile" '''
    if data:
        try:
            if bkp:
                backUpFile(toFile)
            fhdl = open(toFile, 'wb' if isValid(data, bytes) else 'w')
            fhdl.write(data)
            closeFile(fhdl)
            return True
        except:
            getTrace()

    return False

###########################################################################
def setItemIndex(lst, item, dflt=0):
    return lst.index(item) if item in lst else dflt

###########################################################################
def setNewExtn(fileSpec, newExtn):
    return splitFileExt(fileSpec)[0] + newExtn

###########################################################################
def setSeaFood(rel):
    listToFile([rel], 'Parish.rel')

###########################################################################
def setTimeStamp():
    return time.strftime('%Y%m%d_%H%M%S', time.localtime(time.time()))

###########################################################################
def setValFn(fileNm):
    return WhiteSpace.sub('_', BadFNChars.sub('', fileNm))

###########################################################################
def splitFile(fileSpec):
    return os.path.split(fileSpec) if isValid(fileSpec) else [fileSpec]

###########################################################################
def splitFileExt(fileSpec):
    return os.path.splitext(fileSpec) if isValid(fileSpec) else [fileSpec]

###########################################################################
def splitLine(inLine, sep):
    '''
    Split the line by given separator "sep" character(s) and return
    the field list.
    '''
    return [fld.strip() for fld in inLine.split(sep)]

###########################################################################
def splitNstrip(myStr, sepChars=',', dataPri=True):
    flds = [fld.strip() for fld in myStr.split(sepChars)]
    return [fld for fld in flds if len(fld)>0] if dataPri else flds

###########################################################################
def splitStr(myStr, sepChars=','):
    try:
        return stripDataLst(myStr.split(sepChars))
    except:
        getTrace()
        return []

###########################################################################
def splitToInt(myStr, sepChr=',', intBase=10):
    return [int(val, intBase) for val in [fld.strip() for fld in myStr.split(sepChr)] if val.isdigit()]

###########################################################################
def srchRecordDct(modObjDct, srchVal):
    return [k for k, v in modObjDct.items() if re.search(srchVal, k, re.I)]

###########################################################################
def srch_VN_RecDct(modObjDct, srchVal):
    return [k for k, v in modObjDct.items() if re.search(srchVal, remove_viet_accents(k), re.I)]
    #return [k for k, v in modObjDct.items() if (re.search(srchVal, k, re.I) or re.search(srchVal, remove_viet_accents(k), re.I))]

###########################################################################
def strToBytes(hexStr):
    if re.search(' |:|\|', hexStr.strip()):
        return [v.encode() for v in re.split(' |:|\|', hexStr.strip())]
    return [hexStr[i:i + 2].encode() for i in range(0, len(hexStr), 2)]

###########################################################################
def stripDataLst(dataLst):
    return [myStrip(data) for data in dataLst]

###########################################################################
def toInt(valStr, intBase=10):
    if isValid(valStr, int, chkVal=False):
        return valStr
    return int(valStr, intBase)

###########################################################################
def unifyHexStr(hexStr):
    return hexStr.strip().replace(' ','').replace(':','')

###########################################################################
def uniqueLst(src):
    if not isValid(src, (list, tuple)):
        return src
    if isValid(src, tuple):
        src = list(src)
    return sorted(set(src), key=lambda x: src.index(x))

###########################################################################
def validateIps(lstIps):
    '''  IPV4 has 4 octets - IPV6 has 8 words  '''
    for addr in lstIps:
        if not isValid(ipaddress.ip_address(addr),
                       (ipaddress.IPv4Address, ipaddress.IPv6Address)):
            return False
    return True

###########################################################################
def verifyDir(verDir, createIfNotExist=True):
    if isDir(verDir):
        return True
    if createIfNotExist:
        try:
            os.mkdir(verDir)
        except:
            pass
    return isDir(verDir)

###########################################################################
def verifyFileNM(verFileNm, fileExt):
    if len(fileExt)>1 and not verFileNm.lower().endswith(fileExt.lower()):
        return splitFileExt(verFileNm)[0] + fileExt
    return verFileNm

###############################################################################
###############################################################################
class DiaObj:
    ###########################################################################
    def __init__(self, srcKey, recNm, name, abbr, flag, intf, recParms):
        self.srcKey = srcKey
        self.name = name
        self.abbr = abbr
        self.code = recParms[0]
        self.recId = recParms[1]
        self.flag = flag
        self.intf = intf
        self.pxy = 'P' if 'PXY' == recParms[3] else 'A'
        self.oFlag = 'A'
        self.avps = {}

        self.cmdKey = '%s(%s)' % (abbr, self.code)

        if self.isDiaCmd() and not isValCmdAbbr(abbr):
            print('WARNING: Diameter command abbreviation is too long: SrcKey=%s - ABBR=%s - CmdKey=%s' % (srcKey, abbr, self.cmdKey))

    ###########################################################################
    def addAvp(self, avpNm, prio):
        if avpNm != 'AVP':
            addToDct(self.avps, prio, [])
            addToLst(self.avps[prio], avpNm)


    ###########################################################################
    def isDiaCmd(self):
        return '-Request' in self.name or '-Answer' in self.name

###############################################################################
###############################################################################
class IterObj:
    offset = 0

    ###########################################################################
    def __init__(self, specRange):
        specRange, isSpecRange = specRange[0], specRange[-1]

        self.myIter = self.sVal = None
        self.num = None
        if isSpecRange:
            if len(specRange)==1:  specRange.append(1)
            if len(specRange)==2:                             # Unlimited range
                self.sVal, self.inc = specRange
            elif len(specRange)==3:                           # Limit range
                self.sVal, self.num, self.inc = specRange
            self.initIter([])
        else:
            self.sVal = specRange[0]
            self.num = len(specRange)
            self.initIter(specRange[:])

    ###########################################################################
    def getInitVal(self):
        return self.sVal

    ###########################################################################
    def incOffset(self, offsetInc):
        self.offset += offsetInc

    ###########################################################################
    def initIter(self, iterLst):
        if len(iterLst)>0:
            self.myIter = nCycle(iterLst)
        else:
            if self.num is None:                                    # Unlimited range
                self.myIter = iterWhile(self.sVal, self.inc)
            else:
                self.myIter = nCycleRange([self.sVal, self.sVal + self.num*self.inc, self.inc])

    ###########################################################################
    def next(self):
        return self.sVal if self.myIter is None else next(self.myIter)

    ###########################################################################
    def rewind(self):
        self.initIter()
        self.offset = 0

    ###########################################################################
    def setOffset(self, newOffset):
        self.offset = newOffset

###############################################################################
###############################################################################
class RangeObj:
    offset = 0

    ###########################################################################
    def __init__(self, fmtInfoDct):
        self.numItems = -1
        self.valDct = {}
        self.fmtDct = {}
        for var, rangeInfo in fmtInfoDct.items():
            self.fmtDct[var] = IterObj(rangeInfo)
            self.valDct[var] = 1

    ###########################################################################
    def genCompltLst(self):
        if self.numItems > 0:
            return [self.nextItem() for n in range(self.numItems)]
        return []

    ###########################################################################
    def getVal(self, nextVal=True):
        if nextVal:
            for var, fmtObj in self.fmtDct.items():
                self.valDct[var] = fmtObj.next()
        return copyObj(self.valDct)

    ###########################################################################
    def rewind(self):
        self.initIter()
        self.offset = 0

###############################################################################
###############################################################################
class Singleton(type):
    ###########################################################################
    def __call__(self, objId, *args, **kwargs):
        if isNull(self.__dict__.get(objId)):
            setattr(self, objId, super(self.__class__, self).__call__(*args, **kwargs))
        return self.__dict__.get(objId)

###############################################################################
#                                End of module                                #
###############################################################################
