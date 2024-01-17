# ExportVBACode
This repository provide Excel VBA code to automatically export VBA modules when a workbook is saved so that they may be recovered if the pseudocode becomes corrupted. It als provides a module with a macro to reimport multiple files at one time.

I’ve been burned by the poor implementation of VBA into Excel countless times. It is absolutely shameful that Microsoft has let this problem go on for 10+ years with no real attempt to fix it. It rears its ugly head in many different ways but there is one overriding theme: hundreds of lines of VBA code getting crossed up between multiple bidirectional compilation paths and corrupting itself to the point it prevents saving of the file. I’ve seen at least 100 threads of users reporting erratic behavior of VBA-enabled workbooks resulting in loss of tens of thousands of hours of work which I’m highly confident originated from this problem. (Admittedly, I cannot prove that)

Microsoft has fixed this overriding problem in VB.NET. But Excel will not be put over a .NET core anytime soon, if ever, because to do so basically requires a complete rewrite and shareholders will not tolerate the cash flow impact that many resources committed to old code will have. 

OK, I’m off the soapbox and will get down to solutions. While not a fix by any means, I’ve avoided the worst of this problem by adding this code to the MyWorbook object of the VBA project and use the following workflow when saving the project:

1.	When saving the workbook, a message box will ask if you want to export all VBA modules just prior to saving. Check that the code in the modules of the VBA project is visible before saying yes to the dialog. (Hence the cancel button on the dialog in case you forget.) If it is, the it is probably going to export OK.
2.	A file dialog will appear, allowing you to select a folder into which the macro will place a folder called “Source” and into that folder, export the code.
3.	After saving, check that exported files are not empty and are readable with a text editor prior to ditching any previously saved module code. If empty text files are exported, corruption of the intermediate pseudocode is present and code loss is imminent. The previously saved version of the VBA code and Excel file should be retained in case the current one cannot complete the save operation.
4.	If the workbook save operation completes without error, count yourself lucky and keep working.
If the file does become corrupt, here is the recovery procedure:
1.	Browse to the last successfully saved .xlsm file, change the extension to .zip, and open it in a Windows Explorer file view.
2.	In the “xl” folder of the opened zip file is a file called “VBAproject.bin”. Delete it.
3.	Exit the zip file and change the extension back to .xlsm. This will recover the file but at the expense of the VBA code. It must now be reimported.
4.	Locate the set of VBA source files corresponding to the .xlsm file recovered in the previous step.
5.	Open the .xlsm in Excel and then access the VBA IDE. (Alt-F11)
6.	All your VBA code will be gone, but so will the offending pseudocode. You will need to restore your VBA project name manually and reapply any references, as those are lost and not retained in the exported files.
7.	Use the macro in the Import module to reimport all the VBA code files at once. Be sure and use the set that corresponds to the last successful save operation.
8.	Save the rebuilt .xlsm file and continue working.

Note: Both of these modules use a subroutine called GetLocalPath. This subroutine will convert OneDrive and Sharepoint web URLs to traditional Windows file system paths for processing.
