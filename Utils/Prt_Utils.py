"""

 NAME: Prt_Utils.py
 
        # http://timgolden.me.uk/pywin32-docs/win32print__DeviceCapabilities_meth.html
        # http://www.blog.pythonlibrary.org/2010/02/14/python-windows-and-printers/
        # http://timgolden.me.uk/python/win32_how_do_i/print.html


 REVISION HISTORY:
  dd-mmm-yyyy BCR/PR  Created      Comments
  ----------- ------- ------------ -------------------------------------------
******************************************************************************
"""
import win32print
import win32gui
import win32con
import wmi

##############################################################################




##############################################################################
def getInstalledPrts():
    c = wmi.WMI ()
    
    installed_printers = []
    for p in c.Win32_Printer():
        if p.WorkOffline:
            print (p.caption + " is offline")
        else:
            print (p.caption + " is online")
            addToLst(installed_printers, p.caption)
    
    return installed_printers[:]

##############################################################################
def isPrtAvailable():
    c = wmi.WMI()
    for p in c.Win32_Printer():
        if p.Default:
            print ('====>> Printer Status: "%s" - Available: %s' % (p.caption, not p.WorkOffline))
    
##############################################################################
def showNewPrtJob():
    c = wmi.WMI ()
    
    print_job_watcher = c.Win32_PrintJob.watch_for (
        notification_type="Creation",
        delay_secs=1
    )
    
    while 1:
        pj = print_job_watcher ()
        print ("User %s has submitted %d pages to printer %s" % (pj.Owner, pj.TotalPages, pj.Name))
    

##########################################################################
def prepPrt(): 
    # Setting paper tray
    PRINTER_DEFAULTS = {"DesiredAccess":win32print.PRINTER_ALL_ACCESS}
    pHandle = win32print.OpenPrinter('RICOH-LOCAL', PRINTER_DEFAULTS)
    properties = win32print.GetPrinter(pHandle, 2)  #get the properties
    pDevModeObj = properties["pDevMode"]            #get the devmode
    tray_one = 1
    tray_two = 3
    tray_three = 2
    printer_tray = []
    automaticTray = 7
    pDevModeObj.DefaultSource = tray_two            #set the tray
    properties["pDevMode"]=pDevModeObj              #write the devmode back to properties
    win32print.SetPrinter(pHandle,2,properties,0)   #save the properties to the printer

   

##########################################################################
def prepPrtTray():
    
    # Constants from wingdi.h
    DM_OUT_BUFFER = 0x02
    DM_IN_BUFFER = 0x08
    DM_IN_PROMPT = 0x04
    DM_DEFAULT_SOURCE = 0x200
    
    # Get a handle for the default printer
    device_name = win32print.GetDefaultPrinter()
    handle = win32print.OpenPrinter(device_name)
    
    # Get the default properties for the printer
    properties = win32print.GetPrinter(handle, 2)
    devmode = properties['pDevMode']
    
    # Print the default paper source (prints '7' for 'Automatically select')
    print(devmode.DefaultSource)
    
    # Change the default paper source to '4' for 'Manual feed'
    devmode.DefaultSource = 4
    devmode.Fields = devmode.Fields | DM_DEFAULT_SOURCE
    
    # Write these changes back to the printer
    win32print.DocumentProperties(None, handle, device_name, devmode, devmode, DM_IN_BUFFER | DM_OUT_BUFFER)
    
    # Confirm the changes were updated
    print(devmode.DefaultSource)  # Aaargh! Prints '7' again!
    
    # Start printing with the device
    hdc = win32gui.CreateDC('', device_name, devmode)
    win32print.StartDoc(hdc, ('Test', None, None, 0))
    win32print.StartPage(hdc)
    
    # ... GDI drawing commands ...
    
    win32print.EndPage(hdc)
    win32print.EndDoc(hdc)
        
      

##########################################################################
def prepPrt1():
    #http://newcenturycomputers.net/projects/pythonicwindowsprinting.html
    # create a dc (Device Context) object (actually a PyCDC)
    dc = win32ui.CreateDC()
    
    # convert the dc into a "printer dc"
    
    # leave out the printername to get the default printer automatically
    dc.CreatePrinterDC(printername)
    
    # you need to set the map mode mainly so you know how
    # to scale your output.  I do everything in points, so setting 
    # the map mode as "twips" works for me.
    dc.SetMapMode(win32con.MM_TWIPS) # 1440 per inch
    
    # here's that scaling I mentioned:
    scale_factor = 20 # i.e. 20 twips to the point
    
    # start the document.  the description variable is a string
    # which will appear in the print queue to identify the job.
    dc.StartDoc(description)
    
    # to draw anything (other than text) you need a pen.
    # the variables are pen style, pen width and pen color.
    pen = win32ui.CreatePen(0, int(scale_factor), 0L)
    
    # SelectObject is used to apply a pen or font object to a dc.
    dc.SelectObject(pen)
    
    # how about a font?  Lucida Console 10 point.
    # I'm unsure how to tell if this failed.
    font = win32ui.CreateFont({
        "name": "Lucida Console",
        "height": int(scale_factor * 10),
        "weight": 400,
    })
    
    # again with the SelectObject call.
    dc.SelectObject(font)
    
    # okay, now let's print something.
    # TextOut takes x, y, and text values.
    # the map mode determines whether y increases in an
    # upward or downward direction; in MM_TWIPS mode, it
    # advances up, so negative numbers are required to 
    # go down the page.  If anyone knows why this is a
    # "good idea" please email me; as far as I'm concerned
    # it's garbage.
    dc.TextOut(scale_factor * 72,
        -1 * scale_factor * 72,
        "Testing...")
    
    # for completeness, I'll draw a line.
    # from x = 1", y = 1"
    dc.MoveTo((scale_factor * 72, scale_factor * -72))
    # to x = 6", y = 3"
    dc.LineTo((scale_factor * 6 * 72, scale_factor * 3 * -72))
    
    # must not forget to tell Windows we're done.
    dc.EndDoc()

##########################################################################
def prepPrt2():
    #http://newcenturycomputers.net/projects/pythonicwindowsprinting.html
        
    # if you just want to use the default printer, you need
    # to retrieve its name.
    printer = win32print.GetDefaultPrinter()
    
    # open the printer.
    hprinter = win32print.OpenPrinter(printer)
    
    # retrieve default settings.  this code does not work on
    # win95/98, as GetPrinter does not accept two 
    devmode = win32print.GetPrinter(hprinter, 2)["pDevMode"]
    
    # change paper size and orientation
    # constants are available here:
    # http://msdn.microsoft.com/library/default.asp?
    #      url=/library/en-us/intl/nls_Paper_Sizes.asp
    # number 10 envelope is 20
    devmode.PaperSize = 20
    # 1 = portrait, 2 = landscape
    devmode.Orientation = 2
    
    # create dc using new settings.
    # first get the integer hDC value.  note that we need the name.
    hdc = win32gui.CreateDC("WINSPOOL", printer, devmode)
    # next create a PyCDC from the hDC.
    dc = win32ui.CreateDCFromHandle(hdc)
    
    # now you can set the map mode, etc. and actually print.
              
              
                
##############################################################################
# End of File
##############################################################################
