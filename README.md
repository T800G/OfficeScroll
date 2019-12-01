# Office Scroll Add-In
Enhanced mouse wheel scroll add-in for Microsoft Excel and Word.<br/>
Horizonal scroll while Shift key is pressed and scroll by page when Alt key is pressed.
![alt text](https://github.com/T800G/OfficeScroll/blob/master/xlhscroll.gif "Excel horizontal scroll")
![alt text](https://github.com/T800G/OfficeScroll/blob/master/wdvscroll.gif "Word vertical scroll")<br/>

## Minimum requirements:
  * Windows 2000
  * Office 2000


Supports both 32-bit and 64-bit Office.

### IMPORTANT
This add-in is not compatible with KatMouse and similar software. Excel.exe must be added to application ignore list in KatMouse options.<br>
This add-in is not compatible with Excel Horizontal Scroll COM add-in, you need to uninstall it first.<br>
If Windows AppLocker is active on your system, depending on the configuration, you may not be able to use this add-in.

### KNOWN PROBLEMS
The Office version installed from Microsoft Store has problems loading external dll files.
For solution, refer to https://stackoverflow.com/questions/50683727/loadlibrary-in-vba-returns-0-when-trying-to-load-r-dll-in-office-from-microsoft

## Install/Remove
You can install this add-in as a standard user (non-administrator).<br/>
1. Unzip all files.<br/>
2. Double click INSTALL.bat to copy files to their folders (double click REMOVE.bat to delete all files).<br/>
3. Restart Word and Excel.
