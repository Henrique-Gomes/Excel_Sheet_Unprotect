### Excel_Sheet_Unprotect is a Delphi 2006 program that discovers the password to unprotect Excel sheets


* Open or create an Excel file (.xls, .xlsx, ...)
* Click Review -> Protect Sheet
* Insert some password, edit the permissions on the check-list, click OK, confirm password
* Open .dpr file on Delphi as a Project
* Build and run Excel_Sheet_Unprotect.exe and select your excel file on the open dialog (maybe you will need close Excel first)
* The program will generate a password for unprotecting your sheet

The generated password is NOT equals to the password you created, but will equally unprotect the sheet.

Nothing will be modified, not even the Excel file or the Excel program. Generated password will works, even being different from the right password, due to an Excel security problem.

**This program requires installed Excel**

**Sheet's protection password is NOT the same as document's encryption password (found in File->Info->Protect Document->Enctrypt with password)**
