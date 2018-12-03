##############################################################################
# Module: PaperAttr                                                            #
##############################################################################
from __future__ import absolute_import

import os
import sys
import win32con
import win32api
import win32print
import win32com.client
import pywintypes

import win32gui, win32ui
  
from Lib.AppShared import *
from Utils.PyUtils import joinPaths, ExecExtCmd, isNull
from PyPDF2 import pdf, PdfFileWriter, PdfFileReader

X, Y = 0,1
TabloidSize = 1224, 792
LetterSize = 792, 612

Quality = 1200
TrayID = 7154
##########################################################################
# Units (Tenths of a millimeter)
# Length of the paper. To determine the size of the paper in inches, divide this value by 254.
def prtSettingMMTT():
    return {'BitsPerPel' : 0,
            'Collate' : win32con.DMCOLLATE_TRUE,        # DMCOLLATE_FALSE
            'Color' : win32con.DMCOLOR_MONOCHROME,      # DMCOLOR_COLOR
            'Copies' : 1,
            'DefaultSource' : win32con.DMBIN_AUTO,      # DMBIN_MANUAL
            'DeviceName' : '5325 PCL6',
            'DisplayFixedOutput' : win32con.DMDFO_CENTER, # DMDFO_DEFAULT, DMDFO_STRETCH
            'DisplayFlags' : win32con.DM_GRAYSCALE,       # DM_INTERLACED
            'DisplayFrequency' : 0,
            'DisplayOrientation' : win32con.DMDO_DEFAULT, # DMDO_90, DMDO_180, DMDO_270
            'DitherType' : win32con.DMDITHER_GRAYSCALE,   # DMDITHER_NONE, DMDITHER_COARSE, DMDITHER_FINE, ...
            'DriverVersion' : 1539,
            'Duplex' :  win32con.DMDUP_HORIZONTAL,      # DMDUP_SIMPLEX, DMDUP_VERTICAL
            'Fields' : 33666883,
            'FormName' : 'Tabloid',
            'ICMIntent' : win32con.DMICM_CONTRAST,      # DMICM_SATURATE, DMICM_COLORIMETRIC, DMICM_ABS_COLORIMETRIC, DMICM_USER
            'ICMMethod' : win32con.DMICMMETHOD_DEVICE,  # DMICMMETHOD_NONE, DMICMMETHOD_SYSTEM, DMICMMETHOD_DRIVER, DMICMMETHOD_USER
            'LogPixels' : 0,
            'MediaType' : win32con.DMMEDIA_STANDARD,    # DMMEDIA_TRANSPARENCY, DMMEDIA_GLOSSY, DMMEDIA_USER
            'Nup' : win32con.DMNUP_SYSTEM,              # DMNUP_ONEUP
            'Orientation' : win32con.DMORIENT_PORTRAIT,    # DMORIENT_LANDSCAPE, 1-Portrait, 2-Landscape
            'PanningHeight' : 0,
            'PanningWidth' : 0,
            'PaperLength' : 4318,                       # Letter(2794 = 11 X 254), Tabloid(4318 = 11 X 254)
            'PaperSize' : win32con.DMPAPER_TABLOID,     # DMPAPER_LETTER, DMPAPER_LETTERSMALL, DMPAPER_LEGAL, DMPAPER_A3, DMPAPER_A4 ...
            'PaperWidth' : 2794,                        # Letter(2159 = 8.5 X 254), Tabloid(2794 = 11 X 254)
            'PelsHeight' : 0,
            'PelsWidth' : 0,
            'Position_x' : 65538,
            'Position_y' : 141495018,
            'PrintQuality' : Quality,
            'Reserved1' : 0,
            'Reserved2' : 0,
            'Scale' : 100,
            'SpecVersion' : 1025,
            'TTOption' : win32con.DMTT_SUBDEV,          # DMTT_BITMAP, DMTT_DOWNLOAD, DMTT_DOWNLOAD_OUTLINE
            'YResolution' : Quality}

##########################################################################
def rtrvPrinters():
    strComputer = "."  
    objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")  
    objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")  
    colItems = objSWbemServices.ExecQuery("Select * from Win32_PrinterConfiguration")  
    return [objItem.DeviceName for objItem in colItems]

##########################################################################
def setPaperScale(paperForm, docSize):
    if paperForm == 'Tabloid':
        xPrime, yPrime = TabloidSize
        scale = scale2 = 1
    else:
        xPrime, yPrime = LetterSize
        scale = 0.5 * xPrime / docSize[X]
        scale2 = yPrime / docSize[Y]

    if (scale > scale2): scale = scale2
    yOffset = (yPrime - scale * docSize[Y]) / 2
    yOffsetLeft = yOffset
    xOffsetLeft = (xPrime - 2 * scale * docSize[X]) / 4
    xOffset = xOffsetLeft + xPrime / 2
    
    # detecting if pages are already landscape
    # This should not be the case - For booklet, the original pages should be in portrait format
    landscape = False
    if docSize[X] > docSize[Y]:
        landscape = True
        xPrime = 612
        yPrime = 792
        scale = 0.5 * yPrime / docSize[Y]
        scale2 = xPrime / docSize[X]
        if (scale > scale2): scale = scale2
        xOffset = (xPrime - scale * docSize[X]) / 2
        xOffsetLeft = xOffset
        yOffset = (yPrime - 2 * scale * docSize[Y]) / 4
        yOffsetLeft = yOffset + yPrime / 2

    return scale, xPrime, yPrime, xOffset, xOffsetLeft, yOffset, yOffsetLeft
    
##########################################################################
def ordBookletPages(numPages, firstPage, selLastPage):
    lastPage = numPages if selLastPage <= 0 or selLastPage>numPages else selLastPage

    numPages = 1 + lastPage - firstPage
    numSheets = numPages // 4
    if ((numPages > 4) & (numPages % 4 > 0)) | (numSheets == 0):
        numSheets += 1

    pagesToPrint = numSheets * 4
    numBlankPages = pagesToPrint - numPages
    frontOffset, backOffset = 0, 1

    pagesInOrder = []
    pagesInOrder.append(firstPage + pagesToPrint - 1)
    while len(pagesInOrder) < pagesToPrint:
        pagesInOrder.append(firstPage + frontOffset)
        pagesInOrder.append(firstPage + 1 + frontOffset)
        frontOffset += 2
        pagesInOrder.append(firstPage + pagesToPrint - 1 - backOffset)
        pagesInOrder.append(firstPage + pagesToPrint - 1 - backOffset - 1)
        backOffset += 2
    del(pagesInOrder[len(pagesInOrder) - 1])

    return pagesInOrder[:], firstPage, lastPage

##############################################################################
# Printting utilitis 
##############################################################################
#To show Advance Document property setting
def showPrtSetting(selPrt):
    from ctypes import windll

    PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}  
    pHandle = win32print.OpenPrinter(selPrt, PRINTER_DEFAULTS)  

    #To show Document property setting
    win32print.DocumentProperties(0, pHandle, selPrt, None, None, 5)

    #To show Advance Document property setting
    #windll['winspool.drv'].AdvancedDocumentPropertiesA(None, pHandle.handle, selPrt, None, None)

##############################################################################
#To print pDevMode setting
def showPtrAttr(prtProperties):
    print('\n**** Current Printer Setting... ****')
    
    for n in dir(prtProperties):  
        print( "%s\t%s" % (n,getattr(prtProperties,n)) )
    
    print('\n**** Current Printer pDevMode Setting... ****')
    
    # Print current pDevMode setting
    devmode=prtProperties['pDevMode']  
    for n in dir(devmode):  
        if n=="DriverData": continue
        print( "%s\t%s" % (n,getattr(devmode,n)) )

##############################################################################
#To print pDevMode setting
def showpDevMode(prtProperties):
    print('\n**** Current Printer pDevMode Setting... ****')
    
    # Print current pDevMode setting
    devmode=prtProperties['pDevMode']  
    
    print('showpDevMode: ', type(prtProperties['pDevMode']))
    
    for n in dir(devmode):  
        if n=="DriverData" or n[:2]=='__': continue
        print( "%s\t%s" % (n,getattr(prtProperties['pDevMode'] ,n)) )

##############################################################################
def rtrvPrtInfo(handle, level=2):
    return win32print.GetPrinter(handle, level), level

##############################################################################
def setXeroxAttr__(prtProperties, paperForm, numCopies):
    #prtProperties['pDevMode'].Duplex = 1    # no flip

    #prtProperties['pDevMode'].Orientation = 1    # Portrait
    prtProperties['pDevMode'].Orientation = 2     # Landscape

    prtProperties['pDevMode'].FormName = paperForm
    prtProperties['pDevMode'].Copies = numCopies
    
    if paperForm == 'Tabloid':
        prtProperties['pDevMode'].Duplex = 3    # flip over (Short edge)
        prtProperties['pDevMode'].DefaultSource = 7154
    else:
        prtProperties['pDevMode'].Duplex = 2    # flip up (long edge)
        prtProperties['pDevMode'].DefaultSource = 15

##############################################################################
def setXeroxAttr(prtProperties, paperForm, numCopies):
    prtAttrDct = prtSettingMMTT()
    for attrNm in prtAttrDct:
        setattr(prtProperties['pDevMode'], attrNm, prtAttrDct[attrNm])
        
    prtProperties['pDevMode'].Copies = numCopies
    #prtProperties['pDevMode'].PaperSize = win32con.DMPAPER_TABLOID
    prtProperties['pDevMode'].PaperSize = win32con.DMORIENT_PORTRAIT    # DMORIENT_LANDSCAPE
    prtProperties['pDevMode'].PaperLength = prtAttrDct['PaperLength']
    prtProperties['pDevMode'].PaperWidth = prtAttrDct['PaperWidth']
    
    showpDevMode(prtProperties)
    
    return prtProperties

##############################################################################
def setXeroxAttr___(prtProperties, paperForm, numCopies):
    prtProperties['pDevMode'].BitsPerPel = 0
    prtProperties['pDevMode'].Collate = 1
    prtProperties['pDevMode'].Color = 1
    prtProperties['pDevMode'].Copies = numCopies
    prtProperties['pDevMode'].DefaultSource = 7154
    prtProperties['pDevMode'].DeviceName = '5325 PCL6'
    prtProperties['pDevMode'].DisplayFixedOutput = 78650354
    prtProperties['pDevMode'].DisplayFlags = 1
    prtProperties['pDevMode'].DisplayFrequency = 0
    prtProperties['pDevMode'].DisplayOrientation = 65636
    prtProperties['pDevMode'].DitherType = 5
    #prtProperties['pDevMode'].DriverExtra = 7200                    # ReadOnly
    prtProperties['pDevMode'].DriverVersion = 1539
    prtProperties['pDevMode'].Duplex = 3
    prtProperties['pDevMode'].Fields = 33666883
    prtProperties['pDevMode'].FormName = paperForm
    prtProperties['pDevMode'].ICMIntent = 2
    prtProperties['pDevMode'].ICMMethod = 4
    prtProperties['pDevMode'].LogPixels = 0
    prtProperties['pDevMode'].MediaType = 1
    prtProperties['pDevMode'].Nup = 1
    prtProperties['pDevMode'].Orientation = 2
    prtProperties['pDevMode'].PanningHeight = 0
    prtProperties['pDevMode'].PanningWidth = 0
    prtProperties['pDevMode'].PaperLength = 4318
    prtProperties['pDevMode'].PaperSize = win32con.DMPAPER_TABLOID
    prtProperties['pDevMode'].PaperWidth = 2794
    prtProperties['pDevMode'].PelsHeight = 0
    prtProperties['pDevMode'].PelsWidth = 0
    prtProperties['pDevMode'].Position_x = 131074
    prtProperties['pDevMode'].Position_y = 183111902
    prtProperties['pDevMode'].PrintQuality = 1200
    prtProperties['pDevMode'].Reserved1 = 0
    prtProperties['pDevMode'].Reserved2 = 0
    prtProperties['pDevMode'].Scale = 100
    #prtProperties['pDevMode'].Size = 220                                # ReadOnly
    prtProperties['pDevMode'].SpecVersion = 1025
    prtProperties['pDevMode'].TTOption = 3
    prtProperties['pDevMode'].YResolution = 1200

    showpDevMode(prtProperties)
    
    return prtProperties


    
##############################################################################
def createDC(printer, devmode):
    
    # create dc using new settings.
    # first get the integer hDC value.  note that we need the name.
    hdc = win32gui.CreateDC("WINSPOOL", printer, devmode)
    
    print
    # next create a PyCDC from the hDC.
    dc = win32ui.CreateDCFromHandle(hdc)
    
    # now you can set the map mode, etc. and actually print.
              
##############################################################################
def prepPrinter(selPrt):
    #printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_USE}
    #printdefaults = {"DesiredAccess": win32print.PRINTER_ACCESS_ADMINISTER}
    printdefaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    handle = win32print.OpenPrinter(selPrt, printdefaults)

    prtProperties, level = rtrvPrtInfo(handle)

    return handle, level, prtProperties

##############################################################################
def printPdf(selPrt, pdfFN, paperForm, numCopies):
    #Save the current default Printer
    dfltPrt = win32print.GetDefaultPrinter()
    win32print.SetDefaultPrinter(selPrt)
    #reportStat('====>> DefaultPrinter: {}'.format( win32print.GetDefaultPrinter()), None)

    handle, level, prtProperties = prepPrinter(selPrt)
    setXeroxAttr(prtProperties, paperForm, numCopies)
    
    
    
    
    
    #createDC(dfltPrt, prtProperties['pDevMode'])
    

    ## 'SetPrinter' fails because of 'Access is denied.'
    ## But the attribute 'Duplex' is set correctly
    #try:
    win32print.SetPrinter(handle, level, prtProperties, 0)
    #except:
    #    #reportStat("win32print.SetPrinter: set 'Duplex'", None)
    #    pass
    
    #cfgPrt(selPrt)

    res = win32api.ShellExecute(0, 'print', pdfFN, None, '.', 0)
    win32print.ClosePrinter(handle)
    
    
    handle, level, prtProperties = prepPrinter(selPrt)
    showpDevMode(prtProperties)
    

    #Restore the current default Printer
    win32print.SetDefaultPrinter(dfltPrt) 



##############################################################################
def cfgPrt(selPrt):


    p=win32print.OpenPrinter(selPrt)
    printprocessor=win32print.GetPrinter(p,2)['pPrintProcessor']
    dmsize=win32print.DocumentProperties(0,p,selPrt,None,None,0)
    driverextra=dmsize - pywintypes.DEVMODEType().Size
    dm=pywintypes.DEVMODEType(driverextra)
    dm.Fields=dm.Fields|win32con.DM_ORIENTATION|win32con.DM_COPIES#|win32con.DM_PAPERSIZE
    dm.Orientation=win32con.DMORIENT_LANDSCAPE
    dm.Copies=1
    #dm.Papersize=win32con.DM_PAPERLENGTH(790)
    #dm.Paperwidth=530
    win32print.DocumentProperties(0,p,selPrt,dm,dm,win32con.DM_IN_BUFFER|win32con.DM_OUT_BUFFER)
    hDC=win32gui.CreateDC(printprocessor,selPrt,dm)
    printerwidth=win32ui.GetDeviceCaps(hDC,110)
    printerheight=win32ui.GetDeviceCaps(hDC,111)


##############################################################################
def showPrtQue(printerNM=None):
    # rundll32.exe printui.dll,PrintUIEntry /o /n "Printer name"
    if isNull(printerNM):
        printerNM = win32print.GetDefaultPrinter()
    ExecExtCmd(["rundll32.exe", "printui.dll", "PrintUIEntry", "/o", "/n", printerNM])
    
##############################################################################
def printAttr(selPrt):
    #Save the current defauly Printer
    dfltPrt = win32print.GetDefaultPrinter()
    win32print.SetDefaultPrinter(selPrt)
    print('====>> My Current Printer: ', win32print.GetDefaultPrinter())

    handle, level, prtProperties = prepPrinter(selPrt)
    showPtrAttr(prtProperties)
    
##############################################################################
def pdfToBooklet(appPath, pdfFN, paperForm):
    pdfOut = joinPaths(appPath, 'TTMV', 'TTMV_Booklet.pdf')
    
    firstPage = 1
    # Get numPagesInFile using PyPDF
    pdfReader = PdfFileReader(open(pdfFN, 'rb'))
    lastPage = numPagesInFile = pdfReader.getNumPages()
    pageSize = pdfReader.getPage(firstPage).mediaBox.upperRight

    pagesInOrder, firstPage, lastPage = ordBookletPages(numPagesInFile, firstPage, lastPage)

    # Create 2-per-page PDF, ready for printing
    # This is done by creating a side-ways page, and overlaying two scaled pages
    # on top.  US Letter is 612 x 792 pts (portrait)
    x = float(pageSize[0])
    y = float(pageSize[1])
    scale, xPrime, yPrime, xOffset, xOffsetLeft, yOffset, yOffsetLeft = setPaperScale(paperForm, (x, y))

    writer = PdfFileWriter()
    leftPage = True
    for pageNum in pagesInOrder:
        if pageNum > lastPage:
            page = pdf.PageObject.createBlankPage(width=x, height=y)
        else:
            page = pdfReader.getPage(pageNum - 1)

        if leftPage == True:
            combinedPage = pdf.PageObject.createBlankPage(width=xPrime, height=yPrime)
            combinedPage.mergeScaledTranslatedPage(page, scale, xOffsetLeft, yOffsetLeft)
        else:
            combinedPage.mergeScaledTranslatedPage(page, scale, xOffset, yOffset)
            writer.addPage(combinedPage)
        leftPage = not leftPage

    writer.write(open(pdfOut, 'wb'))
    #reportStat("Completed!  Booklet is in {}".format(pdfOut), None)
    
    return pdfOut

##############################################################################
def printTTMV(appPath, selPrt, pdfFN, paperForm, numCopies):
    pdfOut = pdfToBooklet(appPath, pdfFN, paperForm)
    printPdf(selPrt, pdfOut, paperForm, numCopies)
    return pdfOut

##############################################################################
##############################################################################


##############################################################################
# End of File
##############################################################################

