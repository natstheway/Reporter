'Reporting Tool

' INPUT
'Asset inventory
'Excel report of vulnerabilities and missing patches
'CSV output of the info level findings  
'CSV file containing the results for Unauthorized Software enumeration

' OUTPUT
'Total Number of assets scanned with status
'Assets with auth failure
'Vulnerability summary
'Missing Patches summary
'Vulnerability Aging Summary
'Missing Patches Aging Summary
'USB enumeration
'Unsupported Software enumeration
'Outdated Antivirus enumeration
'DLP enumeration
'Share enumeration


' Globals
MsgBox("Welcome to Reporter v1.0")
ret = MsgBox("Select the Vulnerability and Missing Patches file",vbYesNo,"Reporter v1.0")
if ret = 6 Then
	filepath1 = BrowseForFile
	If filePath1 = "" Then
		MsgBox "Operation canceled", vbcritical
	Else
		MsgBox filePath1, vbinformation
	End If
Else
	MsgBox "Exiting...."
	WScript.quit 1
End If
	
'Vulnerability and Missing Patches file
filePath1 = "E:\Reporter Project\test.xlsx"
Set objExcel1 = CreateObject("Excel.Application")
Set objWorkbook1 = objExcel1.Workbooks.Open(filePath1)
Set vulnsheet = objWorkbook1.Sheets(1)
Set msngPatches = objWorkbook1.Sheets(2)
'Info file
filePath2 = "E:\Reporter Project\info.xlsx"
Set objExcel2 = CreateObject("Excel.Application")
Set objWorkbook2 = objExcel2.Workbooks.Open(filePath2)

MsgBox GetMaxColumn(vulnsheet)
MsgBox GetMaxColumn(msngPatches)
MsgBox GetMaxRow(vulnsheet)
MsgBox GetMaxRow(msngPatches)


'Assets scanned with status

'Vulnerability Summary



'Missing Patches Summary

'Vulnerability Aging Summary

'Missing Patches Aging Summary


objExcel1.Quit





'----------------------------------------------------------
'------------------HELPER FUNCTIONS -----------------------
'----------------------------------------------------------

' Returns the maximum number of rows in a excel file
'https://stackoverflow.com/questions/29017663/vbscript-to-read-excel-1-how-to-get-the-row-count-of-specific-column-2-to
Function GetMaxRow(sheet)
	GetMaxRow = sheet.Range("A65536").End(-4162).Row
End Function

' Returns the maximum number of columns in a excel file
Function GetMaxColumn(sheet)
	GetMaxColumn = sheet.Range("XFD4").End(-4159).column
End Function

' Returns the file that is selected
Function BrowseForFile()
'@description: Browse for file dialog.
'@author: Jeremy England (SimplyCoded)
  BrowseForFile = CreateObject("WScript.Shell").Exec( _
    "mshta.exe ""about:<input type=file id=f>" & _
    "<script>resizeTo(0,0);f.click();new ActiveXObject('Scripting.FileSystemObject')" & _
    ".GetStandardStream(1).WriteLine(f.value);close();</script>""" _
  ).StdOut.ReadLine()
End Function
