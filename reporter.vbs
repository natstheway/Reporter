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
MsgBox("Welcome to Reporting Tool v1.0")

Set objXLApp = CreateObject("Excel.Application")

'Vulnerability and Missing Patches file
filePath1 = "E:\Reporter Project\test.xlsx"
Set objExcel1 = CreateObject("Excel.Application")
Set objWorkbook1 = objExcel1.Workbooks.Open(filePath1)

'Info file
filePath2 = "E:\Reporter Project\info.xlsx"
Set objExcel2 = CreateObject("Excel.Application")
Set objWorkbook2 = objExcel2.Workbooks.Open(filePath2)


'Assets scanned with status

'Vulnerability Summary
intRow = 2

Do Until objExcel1.Cells(intRow,1).Value = ""

intRow = intRow + 1
Loop
objExcel1.Quit

'Missing Patches Summary

'Vulnerability Aging Summary

'Missing Patches Aging Summary

