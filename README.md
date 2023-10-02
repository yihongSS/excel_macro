# excel_macro

As suggested by Microsoft documentation, "If you want to share your Personal.xlsb file with others, you can copy it to the XLSTART folder on other computers. In Windows 10, Windows 7, and Windows Vista, this workbook is saved in the C:\Users\user name\AppData\Local\Microsoft\Excel\XLStart folder. In Microsoft Windows XP, this workbook is saved in the C:\Documents and Settings\user name\Application Data\Microsoft\Excel\XLStart folder. Workbooks in the XLStart folder are opened automatically whenever Excel starts, and any code you have stored in the personal macro workbook will be listed in the Macro dialog,

"

Source: https://support.microsoft.com/en-us/office/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566



Notes: 
* if you could not find the XLStart folder, try to check if the Macro is enabled; normally after macros being enabled, the XLStart would be created automatically by MS excel.
* sometimes it's in the Roaming folder instead of Local folder, for instance: "C:\Users\user name\AppData\Roaming\Microsoft\Excel\XLSTART"
