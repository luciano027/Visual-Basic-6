Thank you for purchasing an Epson printer. This document contains notes regarding your new printer.

---------------------------------------------------------------------

The following contains the latest information about your EPSON LX-300+.
                                                          26/11/99
---------------------------------------------------------------------

Using device fonts on Microsoft Office 97
=========================================

When you set to print with device fonts on Microsoft Office 97, 
you may not get printout printed in device fonts.
In such a case, set the following modes to make device fonts available.

*Change compatibility mode from Office 97 to Office 95.


---------------------------------------------------------------------

The following contains the latest information about your EPSON Status Monitor 3.
                                                          12/27/99
---------------------------------------------------------------------
1. General troubleshooting solutions

1-1. Notes for using EPSON Status Monitor 2 and EPSON Status Monitor 3
     together

     EPSON Status Monitor3 turns on monitoring of the following types of
     connections, even if these settings are disabled by EPSON Status 
     Monitor 2.
      * local printers
      * LPR printers
      * NetWare printers
      * Windows shared printers
     These settings should work most in cases.

      If you install EPSON Status Monitor 2 and change the monitoring settings,
      it will also change EPSON Status Monitor 3's settings. This
      can cause problems printing and monitoring.
     (Note)   *If a particular communication protocol is not installed, the
               monitor settings will not be enabled.
              *To setup EPSON Status Monitor 3, please re-install your printer
               driver. 

1-2. If a Windows shared printer cannot be monitored, check that the following 
     settings have been made:

      * Make sure that the printer driver is installed on the computer that is
        hosting the Windows shared printer, that it is set to be shared, and 
        "Allow monitoring of shared printers" is selected in the 
        "Monitoring Preferences" of the printer driver.
     
      * (Windows98/95 only) Open the "Network" icon in the control panel on the
        computer that is sharing the printer, and check that File and Printer 
        Sharing for Microsoft Networks component is installed.

      * (Windows98/95 only) As above, check that IPX/SPX-compatible Protocol is
        NOT installed on either the client or server computers.


2. Solutions for Windows NT/2000 environment

2-1. When un-installing EPSON Status Monitor 3 from a multi-user Windows NT/2000 environment,
     turn off the Shortcut Icon on all clients before un-installing. You can turn off
     the Shortcut Icon in Monitoring Preferences.

2-2. When using a Windows NT 4.0 shared printer in a non-updated Windows NT environment,
     if you check "Allow monitoring of shared printer", and then start EPSON Status Monitor
     3, the Application Error Notice Window appears and you cannot monitor NetWare shared
     printers. To avoid this problem, make sure you have installed Windows NT 4.0 Service
     Pack 2 or later.

2-3. When using a Windows NT/2000 shared printer, if you check "Allow monitoring of shared printer",
     and then start EPSON Status Monitor 3, you cannot monitor a NetWare shared printer . To avoid
     this problem open "Services" from the control panel, stop the EPSON Printer Status Agent
     service, and start EPSON Status Monitor 3. After EPSON Status Monitor 3 has started, restart
     the printer status service. You must do this each time you log into Windows NT/2000.

2-4. When uninstalling the printer driver from the "Add/Remove Programs" Control Panel in
     Windows 2000, you may be unable to close the "Add/Remove Programs" window by clicking "OK".
     To close the window, click the Close [X] button at the upper right corner of the window.
     This problem does not effect the printer driver uninstallation. 


3. Solutions for network printers

3-1. To monitor a NetWare printer, connect only one print server to
     each queue. See your User's Guide for directions.

3-2. You cannot monitor a NetWare queue printer beyond the IPX router
     because the setting for IPX hops is 0.

3-3. When monitoring NDS printers, the client supplied by Novell must
     be used:

     When using Windows NT: Novell IntranetWare Client for Windows NT
     When using Windows 95: Novell IntranetWare Client 32 for Windows 95

3-4. Monitoring NDPS printers

    This product supports Netware printers configured by a queue-based
    printing system such as Bindery or NDS server.  You cannot monitor
    NDPS (Novell Distributed Printing Services) printers.


4. Solutions for using EPSON Status Monitor 3 for impact printers.

4-1. Notes for error and warning notification when printing

    *If connecting to a NetWare printer and monitoring it while printing, make
     sure EPSON Status Monitor 3 is running.

    *If connecting to a LPR printer and monitoring it while printing, make
     sure EPSON Status Monitor 3 is running.

    *When connecting to a Windows NT/2000 shared printer from Windows 98/95, error and
     warning notification are not possible during printing. Make sure EPSON Status
     Monitor 3 is running, and use the Printer Status Window to find out about errors.

    *When connecting to a shared printer over router, in some cases, error
     and warning notification for shared printers during printing is not available
     because of the router settings.


---------------------------------------------------------------------

General notice: Product names used herein are for identification 
purposes only and may be trademarks of their respective owners. 
EPSON disclaims any and all rights in those marks.

(C)SEIKO EPSON Corp.1999

END
