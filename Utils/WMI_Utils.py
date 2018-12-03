"""

 NAME: WMI_Utils.py
 
 From http://timgolden.me.uk/python/wmi/cookbook.html
 # http://www.activexperts.com/admin/scripts/wmi/python/0353/


 REVISION HISTORY:
  dd-mmm-yyyy BCR/PR  Created      Comments
  ----------- ------- ------------ -------------------------------------------
******************************************************************************
"""
import wmi
#import _winreg

from .PyUtils import *

##############################################################################
##############################################################################
def getRunningProc():
    c = wmi.WMI ()
    actProcLst = []
    for process in c.Win32_Process():
        addToLst(actProcLst, process.Name)
    return actProcLst[:]

##############################################################################
def getActiveProc(proName='python.exe'):
    c = wmi.WMI ()
    actProcLst = []
    for process in c.Win32_Process(name=proName):
        addToLst(actProcLst, process.Name)
    return actProcLst[:]

##############################################################################
def pKill(proName='python.exe'):
    c = wmi.WMI ()
    for process in c.Win32_Process(name=proName):
        print('====>> terminating process: ', process.Name)
        process.Terminate()

##############################################################################
def getWinServicesStat():
    c = wmi.WMI ()
    winSrvcLst = []

    for srvc in c.Win32_Service (StartMode="Auto"):
        addToLst(winSrvcLst, (srvc.Name, srvc.State))
        print('====>> terminating process: ', srvc.Name, srvc.State)
    return winSrvcLst[:]

##############################################################################
def getActvWinServices():
    c = wmi.WMI ()
    winSrvcLst = []

    for srvc in c.Win32_Service (StartMode="Auto", State='Running'):
        addToLst(winSrvcLst, (srvc.Name, srvc.State))
        print('====>> terminating process: ', srvc.Name, srvc.State)
    return winSrvcLst[:]

##############################################################################
def getStoppedWinServices():
    c = wmi.WMI ()
    winSrvcLst = []

    for srvc in c.Win32_Service (StartMode="Auto", State='Stopped'):
        addToLst(winSrvcLst, (srvc.Name, srvc.State))
        print('====>> terminating process: ', srvc.Name, srvc.State)
    return winSrvcLst[:]

##############################################################################
def reptDiskSpace():
    c = wmi.WMI ()
    
    for disk in c.Win32_LogicalDisk(DriveType=3):
        usage = "%0.2f%% free" % (100 * int(disk.FreeSpace) / int(disk.Size))
        print(disk.Caption, 'Available: %s' % disk.FreeSpace, 'Cap: %s' % disk.Size, usage)
        



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
    
##############################################################################
def rebootRemote(remoteSyst):
    c = wmi.WMI ()
    # remoteSyst = "machine name of your choice"
    c = wmi.WMI (computer=remoteSyst, privileges=["RemoteShutdown"])
    
    os = c.Win32_OperatingSystem (Primary=1)[0]
    os.Reboot ()
    
##############################################################################
def showIP_MAC():
    c = wmi.WMI ()
    for interface in c.Win32_NetworkAdapterConfiguration (IPEnabled=1):
        print('\n', interface.Description, interface.MACAddress)
        for ip_address in interface.IPAddress:
            print(ip_address)
        
##############################################################################
def showStartupProc():
    c = wmi.WMI ()
    for s in c.Win32_StartupCommand ():
        print("[%s] %s <%s>" % (s.Location, s.Caption, s.Command))
        
##############################################################################
def showEventLog():
    c = wmi.WMI (privileges=["Security"])

    watcher = c.watch_for (
        notification_type="Creation",
        wmi_class="Win32_NTLogEvent",
        Type="error"
        )
    while 1:
        error = watcher ()
        print("Error in %s log: %s" %  (error.Logfile, error.Message))
        # send mail to sysadmin etc.

##############################################################################
def lstRegistryKeys():
    regKeys = wmi.WMI(namespace="DEFAULT").StdRegProv
    print(regKeys)
          
##############################################################################
def addRegistryKey(keyName):
    r = wmi.Registry()
    result, = r.CreateKey(
        hDefKey=_winreg.HKEY_LOCAL_MACHINE,
        sSubKeyName=keyName
    )         
          
##############################################################################
def addRegistryValue(keyName, value):
    r = wmi.Registry ()
    result, = r.SetStringValue (
        hDefKey=_winreg.HKEY_LOCAL_MACHINE,
        sSubKeyName=keyName,                        # r"Software\TJG"
        sValueName=value,                           # "ApplicationName",
        sValue="TJG App"
    )          
          
##############################################################################
def addIISsite():
    c = wmi.WMI (namespace="MicrosoftIISv2")
    
    #
    # Could as well be achieved by doing:
    #  web_server = c.IISWebService (Name="W3SVC")[0]
    #
    for web_server in c.IIsWebService (Name="W3SVC"):
        break
    
    binding = c.new ("ServerBinding")
    binding.IP = ""
    binding.Port = "8383"
    binding.Hostname = ""
    result, = web_server.CreateNewSite (
        PathOfRootVirtualDir=r"c:\inetpub\wwwroot",
        ServerComment="My Web Site",
        ServerBindings= [binding.ole_object]
    )    
    
    
          
                
##############################################################################
# End of File
##############################################################################
