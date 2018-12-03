"""
******************************************************************************
 NAME: lxmlUtils.py
******************************************************************************
"""
import re

from lxml import etree, html
from io import StringIO

from Lib.AppShared import *
from .PyUtils import (pause, copyObj, addToLst, saveToFile, verifyFileNM,
                      decodeUtf8, getTrace, isValid, isNull, isNotNull,
                      getDirname, getBasename, joinPaths, addToLst, isFile)

##############################################################################
XI_URL = 'http://www.w3.org/2003/XInclude'
XINCLUDE = '{%s}' % XI_URL
XINCLUDE_INCLUDE = XINCLUDE + 'include'
XINCLUDE_FALLBACK = XINCLUDE + 'fallback'
XI_NSMAP = {'xi' : XI_URL}

XML_Tail = '\n  '
EndTags = re.compile('">\s+\</\w+>')
XmlVer = re.compile('<\?.+\?>')

TagSpaces = re.compile('>\s+\<')

##########################################################################
def createXml(xmlName, xInclude=True, **xmlAttr):
    if xInclude:
        return etree.Element(xmlName, nsmap=XI_NSMAP, **xmlAttr)
    return etree.Element(xmlName, **xmlAttr)

##########################################################################
def createElement(parent, eName, addTail=False, **eAttr):
    if isNotNull(parent) and not addTail:
        return etree.SubElement(parent, eName, **eAttr)
    elm = etree.Element(eName, **eAttr)
    if addTail is True:
        elm.tail = XML_Tail
    if isNotNull(parent):
        parent.append(elm)
    return elm

##########################################################################
def addElement(addTo, elm, verChild=True):
    if isXmlElm(elm) and (not verChild or len(elm)>0):
        addTo.append(elm)

##########################################################################
def defXInclude(parent, addTail=False, **eAttr):
    elm = etree.Element(XINCLUDE + 'include', **eAttr)
    if addTail:
        elm.tail = XML_Tail
    if isNotNull(parent):
        parent.append(elm)
    return elm

##########################################################################
def addXmlComment(parent=None, comment=''):
    xmlComment = etree.Comment(comment) if isinstance(comment, str) else comment
    if isNotNull(parent):
        parent.append(xmlComment)
    return xmlComment

##########################################################################
def replaceElm(oldElm, newElm):
    oldElm.getparent().replace(oldElm, newElm)

##########################################################################
def removeElm(elm):
    elm.getparent().remove(elm)

##########################################################################
def duplicateElm(elm, occurence=0):
    parentElm = elm.getparent()
    if occurence>0:
        elm.attrib['occurence'] = '%d' % occurence
        occurence += 1
    myInd = parentElm.index(elm) + 1
    if occurence>0:
        cloneElm = copyObj(elm)
        cloneElm.attrib['occurence'] = '%d' % occurence
    parentElm.insert(myInd, cloneElm)

##########################################################################
def getElmIndex(elm):
    return elm.getparent().index(elm)

##########################################################################
def insertElm(elm, newElm, belowElm=True):
    parentElm = elm.getparent()
    myInd = parentElm.index(elm) + 1 if belowElm else 0
    parentElm.insert(myInd, newElm)

##########################################################################
def addOccurence(elm, occurence):
    elm.attrib['occurence'] = occurence

##########################################################################
def migrateXml(scrXml, xmlTag, **xmlAttr):
    dstXml = createXml(xmlTag, **xmlAttr)
    [dstXml.append(elm) for elm in scrXml]
    return dstXml

##########################################################################
def cloneTopElm(srcElm):
    dstElm = copyObj(srcElm)
    [dstElm.remove(elm) for elm in dstElm]
    return dstElm

##########################################################################
def getXiElms(selXml):
    xIncMsgs = []
    include_elements = list(selXml.iter('{%s}*' % XI_URL))
    [xIncMsgs.append((elm.attrib['href'], elm)) for elm in include_elements]
    return xIncMsgs[:]

##########################################################################
def getXiFiles(baseDir, selXml):
    xIncFiles = []
    include_elements = list(selXml.iter('{%s}*' % XI_URL))
    [addToLst(xIncFiles, joinPaths(baseDir, elm.attrib['href'])) for elm in include_elements]
    return xIncFiles[:]

##########################################################################
def readXml(xmlSrc):
    parser = etree.XMLParser(remove_blank_text=True)
    try:
        tree = etree.parse(xmlSrc, parser)
        return tree.getroot()
    except:
        pass

##########################################################################
def loadXml(xmlSrc, cmplt=False, frmFile=True, retEMsg=False):
    if not isValid(xmlSrc): return
    parser = etree.XMLParser(remove_blank_text=True)

    try:
        if frmFile:
            tree = etree.parse(xmlSrc, parser)
        else:
            tree = etree.parse(StringIO(XmlVer.sub('', xmlSrc)), parser)
        if cmplt:
            tree.xinclude()
        return tree.getroot()
    except Exception:
        statMsg = getTrace(retEMsg=True)

    if retEMsg is True: return statMsg
    reportStat(statMsg)

##########################################################################
def isXmlElm(xmlSrc):
    return isValid(xmlSrc, etree._Element)

##########################################################################
def isXmlFile(srcFile):
    parser = etree.HTMLParser()
    try:
        etree.parse(srcFile, parser)
        return True
    except:
        return False

##########################################################################
def loadHtml(htmlSrc, frmFile=True, retEMsg=False):
    parser = etree.HTMLParser()
    try:
        if frmFile:
            tree = etree.parse(htmlSrc, parser)
        else:
            tree = etree.parse(StringIO(XmlVer.sub('', htmlSrc)), parser)
        return tree.getroot()
    except Exception:
        statMsg = getTrace(retEMsg=True)

    if retEMsg: return statMsg
    reportStat(statMsg)

####################################################################################
def xmlSrch(xml, target):
    return xml.findall(target)

##########################################################################
def findTagret(xml, tag, fldNm, fldVal):
    for item in xmlSrch(xml, './/%s' % tag):
        if fldNm in item.attrib:
            if item.attrib[fldNm]==fldVal:
                return item
    return None

##########################################################################
##########################################################################
def findElm(xmlElm, srchXpath):
    '''
    srchXpath:
    .//tag
    .//tag/subTag1/subTag2
    .//tag[@attr]
    .//tag[@attr="attr_val"]
    '''
    return xmlElm.find(srchXpath)

##########################################################################
'''
srchXpath: formats
.//tag
.//tag/subTag1/subTag2
.//tag[@attr]
.//tag[@attr="attr_val"]
'''
##########################################################################
def findElm(xmlElm, srchXpath):
    return xmlElm.find(srchXpath)

##########################################################################
def findElms(xmlElm, srchXpath):
    return xmlElm.findall(srchXpath)

##########################################################################
def getElmAttr(xmlElm, attrNm, dflt=None):
    if isNotNull(xmlElm):
        return xmlElm.attrib.get(attrNm, dflt)

##########################################################################
def sortXmlElms(entriesLst, sortKey):
    data = {}
    for elm in entriesLst:
        key = elm.attrib[sortKey]
        data[key] = elm
    data = sorted(data.items())
    return [item[-1] for item in data]

##############################################################################
def getElmInfo(source, elmTree=True):
    if elmTree:
        item = source.currentItem()
    else:
        item = source.getparent()
    return item, item.statusTip(0)

##############################################################################
def getParentInfo(childElm):
    parentElm = childElm.getparent()
    return parentElm, parentElm.tag

##############################################################################
def hasChildren(elm):
    return len(elm.getchildren())>0

##############################################################################
def isComment(elm):
    return isValid(elm, etree._Comment)

##############################################################################
##############################################################################
def xmlToFile(xml, xmlFile, fileExtn='.xml', bkp=True, sig=None):
    addSignature(xml, sig)
    xmlFile = verifyFileNM(xmlFile, fileExtn)
    xmlFile = os.path.normpath(xmlFile)
    xmlStr = loadXml(printXml(xml), frmFile=False) if isXmlElm(xml) else xml
    if saveToFile(EndTags.sub('"/>', printXml(xmlStr) ), xmlFile, bkp):
        return xmlFile

##########################################################################
def iterChildren(xml, reversed=False):
    for elem in xml.iterchildren(reversed=reversed):
        yield elem

##########################################################################
def iterAllChildren(xml):
    for elem in xml.getiterator():
        yield elem

##############################################################################
def addSignature(xml, sig=None):
    xml.attrib.update({'signature':createSignature() if isNull(sig) else sig})

##############################################################################
def printXml(xml, prtPretty=True):
    if prtPretty:
        return decodeUtf8(etree.tostring(xml, pretty_print=True, encoding="iso-8859-1"), asIs=True)
    return decodeUtf8(etree.tostring(xml, encoding="iso-8859-1"), asIs=True)


##########################################################################
# HTML wrappers
##########################################################################
def loadHtml(htmlSrc, frmFile=True):
    parser = etree.HTMLParser(remove_blank_text=True)
    try:
        if frmFile:
            tree = etree.parse(htmlSrc, parser)
        else:
            tree = etree.parse(StringIO(XmlVer.sub('', htmlSrc)), parser)
        return tree.getroot()
    except Exception:
        statMsg = getTrace(retEMsg=True)
    reportStat(statMsg)

##########################################################################
def htmlElmFrStr(htmlStr):
    return html.fromstring(htmlStr)

##########################################################################
def htmlToStr(htmlStr):
    return decodeUtf8(html.tostring(htmlStr))

##########################################################################
def valXmlSyntax(xmlFileLst, callBack=None):
    if isNull(callBack):  callBack = valXmlSyntax
    for ind, xmlFn in enumerate(xmlFileLst):
        xml = isValidXml(xmlFn)
        srcDir = getDirname(xmlFn)
        if isXmlElm(xml):
            callBack(getXiFiles(srcDir, xml))
        else:
            reptBadXml(xmlFn, xml)
        setStat('\t\tTotal number of XML files from "%s" has been validated: %21d' % (getBasename(srcDir), ind+1))
        addMsg(((PROG_Stat, 1), PROG_BAR))

##########################################################################
def isValidXml(xmlFn):
    parser = etree.XMLParser(remove_blank_text=True)
    try:
        if isFile(xmlFn):
            return etree.parse(xmlFn, parser).getroot()
        return 'File DOES NOT exist: "%s".' % xmlFn
    except Exception:
        return getTrace(retEMsg=True)

##########################################################################
def reptBadXml(srcFn, eMsg):
    reportStat('Unable to process XML file: "%s".' % srcFn)
    reportStat('====>> Cause: "%s"' % eMsg)
    
    raise('Unable to process XML file: "%s".' % srcFn)

##########################################################################
##########################################################################
class XmlValidator:

    def __init__(self):
        self.fileLst = []

    def validate(self, fileLst):
        selLst = [fn for fn in fileLst if addToLst(self.fileLst, fn)]
        if any(selLst):
            valXmlSyntax(selLst, callBack=self.validate)

##############################################################################
# End of File
##############################################################################

