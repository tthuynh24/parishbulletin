"""
 NAME: QTreeUtils.py
"""
from .PyUtils import isValid, isNull, isNotNull

##########################################################################
def isTopLevel(elm):
    return isNull(elm.parent())

##########################################################################
def findItems(parent, peerSeq):
    try:
        parent.topLevelItem(0)
        return findTopElm(parent, peerSeq)
    except:
        return findChildElm(parent, peerSeq)

##########################################################################
def findTopElm(qTree, srchElm, peerSeq):
    numPeers = qTree.topLevelItemCount()
    myInd = peerInd = None
    for ind in range(numPeers):
        peerNm = qTree.topLevelItem(ind).text(0)
        if peerNm == srchElm:
            myInd = ind + 1
        elif peerNm in peerSeq:
            if isNull(peerInd):
                peerInd = ind
    return myInd if isNotNull(myInd) else peerInd if isNotNull(peerInd) else numPeers

##########################################################################
def insertTopElm(currElm, elm, dfltParent=None, atInd=-1):
    parent = currElm if atInd>=0 else currElm.parent()
    if atInd<0:
        atInd = dfltParent.indexOfTopLevelItem(currElm) + 1
    dfltParent.insertTopLevelItem(atInd, elm)

##########################################################################
def findChildElm(parent, srchElm, peerSeq=[]):
    numPeers = parent.childCount()
    myInd = peerInd = None
    for ind in range(numPeers):
        peerNm = parent.child(ind).text(0)
        if peerNm == srchElm:
            myInd = ind + 1
        elif peerNm in peerSeq:
            if isNull(peerInd):
                peerInd = ind
    return myInd if isNotNull(myInd) else peerInd if isNotNull(peerInd) else numPeers


##########################################################################
def cntTopElm(qTree, emlNm):
    numElms = qTree.topLevelItemCount()
    elmCnts = 0
    for ind in range(numElms):
        if qTree.topLevelItem(ind).text(0)==emlNm:
            elmCnts += 1
    return elmCnts

##########################################################################
def cntChildElm(parent, emlNm):
    numElms = parent.childCount()
    elmCnts = 0
    for ind in range(numElms):
        if parent.child(ind).text(0)==emlNm:
            elmCnts += 1
    return elmCnts

##########################################################################
def getChildElms(parent):
    childInfo = []
    for ind in range(parent.childCount()):
        childInfo.append(parent.child(ind).text(0))
    return childInfo


##########################################################################
##########################################################################
def getChild(parent, childInd):
    if childInd < parent.childCount():
        childElm = parent.child(childInd)
        return childElm, childElm.text(0)
    return None, None

##########################################################################
def getGrpElm(selItem, grpTags):
    pElm, pNm = getParent(selItem)
    while isValid(pElm) and pNm not in grpTags:
        pElm, pNm = getParent(pElm)
    return pElm

##########################################################################
def insertElm(currElm, elm, dfltParent=None, atInd=-1):
    parent = currElm if atInd>=0 else currElm.parent()
    if atInd<0:
        atInd = parent.indexOfChild(currElm)+1
    parent.insertChild(atInd, elm)

##########################################################################
def insertElmAt(parent, elm, atInd=-1, dfltParent=None):
    if atInd>=0:
        parent.insertChild(atInd, elm)

##########################################################################
def getSelItems(tWidget):
    return tWidget.selectedItems()

##########################################################################
def getCurrItem(tWidget):
    currItem = tWidget.currentItem()
    if isNull(currItem):
        return None, 'Unknown'
    return currItem, currItem.text(0)

##########################################################################
def getParent(child):
    if isNull(child):
        return None, 'UnKnown'
    parent = child.parent()
    if isNull(parent):
        return None, 'TopLevel'
    return parent, parent.text(0)

##########################################################################
def dltElm(elm2Dlt):
    parent, parentNm = getParent(elm2Dlt)
    parent.takeChild(parent.indexOfChild(elm2Dlt))
    del elm2Dlt

##############################################################################
# End of File
##############################################################################
