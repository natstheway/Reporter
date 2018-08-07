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

on error resume next

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

ret = MsgBox("Select the file with Informational findings",vbYesNo,"Reporter v1.0")
if ret = 6 Then
	filepath2 = BrowseForFile
	If filePath2 = "" Then
		MsgBox "Operation canceled", vbcritical
	Else
		MsgBox filePath2, vbinformation
	End If
Else
	MsgBox "Exiting...."
	WScript.quit 1
End If

ret = MsgBox("Select the file with Asset inventory",vbYesNo,"Reporter v1.0")
if ret = 6 Then
	filepath3 = BrowseForFile
	If filePath3 = "" Then
		MsgBox "Operation canceled", vbcritical
	Else
		MsgBox filePath3, vbinformation
	End If
Else
'	MsgBox "Exiting...."
'	WScript.quit 1
End If

	
'Vulnerability and Missing Patches file
vulagecatheader=1
vulagecat1=2
vulagecat2=3
vulagecat3=4
vulagecat4=5
vulagecat5=6
vulagecat6=7
severityhdr = 8
critical = 9
high = 10
medium = 11
low = 12


'Open the Missing Patches and Vulnerability file
Set objExcel1 = CreateObject("Excel.Application")
Set objWorkbook1 = objExcel1.Workbooks.Open(filePath1)
Set vulnsheet = objWorkbook1.Sheets(1)
Set msngPatches = objWorkbook1.Sheets(2)
ObjWorkbook1.Sheets.Add.Name = "Summary"
Set SummarySheet = objWorkbook1.Sheets("Summary")

'Open the Info file
Set objExcel2 = CreateObject("Excel.Application")
Set objWorkbook2 = objExcel2.Workbooks.Open(filePath2)
Set infosheet = objWorkbook2.Sheets(1)

'Open the Asset inventory file
Set objExcel3 = CreateObject("Excel.Application")
Set objWorkbook3 = objExcel3.Workbooks.Open(filePath3)
Set assetsheet = objWorkbook3.Sheets(1)

'Open the Target Word File
Set objWord = CreateObject("Word.Application")
objWord.Caption = "Security Services Report"
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection


'Few basic checks
'MsgBox wordOutputDir
MsgBox GetMaxColumn(vulnsheet)
MsgBox GetMaxColumn(msngPatches)
MsgBox GetMaxRow(vulnsheet)
MsgBox GetMaxRow(msngPatches)

MsgBox "Added new columns for vulnerability aging calculation..."
'Create a new column - last observed date with a formula
LastObsDate = GetMaxColumn(vulnsheet) + 1
vulnsheet.Cells(1,LastObsDate).value = "Last Observed Date"

PluginPubDate = GetMaxColumn(vulnsheet) + 1
MsgBox PluginPubDate
'Create a new column - plugin pub date with a formula
vulnsheet.Cells(1,PluginPubDate).value = "Plugin Published Date"

VulnAge = GetMaxColumn(vulnsheet) + 1
MsgBox VulnAge
'Create a new column - plugin pub date with a formula
vulnsheet.Cells(1,VulnAge).value = "Vulnerability Age"

VulnAgeCategry = GetMaxColumn(vulnsheet) + 1
MsgBox VulnAgeCategry
'Create a new column - plugin pub date with a formula
vulnsheet.Cells(1,VulnAgeCategry).value = "Vulnerability Age Category"

'severity is in 5th column
' TODO : automatically find iter
severity = 5

'Initializing Summary Sheet
SummarySheet.Cells(vulagecatheader,1) = "Vulnerability Aging Summary"
SummarySheet.Cells(vulagecat1,1) = "0-3 months"
SummarySheet.Cells(vulagecat2,1) = "3-6 months"
SummarySheet.Cells(vulagecat3,1) = "6-12 Months"
SummarySheet.Cells(vulagecat4,1) = "1-2 Years"
SummarySheet.Cells(vulagecat5,1) = "2-5 Years"
SummarySheet.Cells(vulagecat6,1) = ">5 Years"
SummarySheet.Cells(vulagecat1,2) = 0
SummarySheet.Cells(vulagecat2,2) = 0
SummarySheet.Cells(vulagecat3,2) = 0
SummarySheet.Cells(vulagecat4,2) = 0
SummarySheet.Cells(vulagecat5,2) = 0
SummarySheet.Cells(vulagecat6,2) = 0

SummarySheet.Cells(severityhdr,1) = "Severity Summary"
SummarySheet.Cells(Critical,1) = "Critical"
SummarySheet.Cells(High,1) = "High"
SummarySheet.Cells(Medium,1) = "Medium"
SummarySheet.Cells(Low,1) = "Low"
SummarySheet.Cells(Critical,2) = 0
SummarySheet.Cells(High,2) = 0
SummarySheet.Cells(Medium,2) = 0
SummarySheet.Cells(Low,2) = 0

SummarySheet.Cells(vulagecatheader,3) = "Missing Patches Aging Summary"
SummarySheet.Cells(vulagecat1,3) = "0-3 months"
SummarySheet.Cells(vulagecat2,3) = "3-6 months"
SummarySheet.Cells(vulagecat3,3) = "6-12 Months"
SummarySheet.Cells(vulagecat4,3) = "1-2 Years"
SummarySheet.Cells(vulagecat5,3) = "2-5 Years"
SummarySheet.Cells(vulagecat6,3) = ">5 Years"
SummarySheet.Cells(vulagecat1,4) = 0
SummarySheet.Cells(vulagecat2,4) = 0
SummarySheet.Cells(vulagecat3,4) = 0
SummarySheet.Cells(vulagecat4,4) = 0
SummarySheet.Cells(vulagecat5,4) = 0
SummarySheet.Cells(vulagecat6,4) = 0

SummarySheet.Cells(severityhdr,3) = "Severity Summary"
SummarySheet.Cells(Critical,3) = "Critical"
SummarySheet.Cells(High,3) = "High"
SummarySheet.Cells(Medium,3) = "Medium"
SummarySheet.Cells(Low,3) = "Low"
SummarySheet.Cells(Critical,4) = 0
SummarySheet.Cells(High,4) = 0
SummarySheet.Cells(Medium,4) = 0
SummarySheet.Cells(Low,4) = 0


'Vulnerability age and count of severity calculation
MsgBox "Computing the aging and severity summary "

iter = 2
Do Until iter = GetMaxRow(vulnsheet)
vulnsheet.Cells(iter,LastObsDate).Formula = "=LEFT(M" & iter & ",(12))"
vulnsheet.Cells(iter,PluginPubDate).Formula = "=LEFT(O" & iter & ",(12))"
vulnsheet.Cells(iter,VulnAge).Formula = "=R" & iter & "-S" & iter
vulnsheet.Cells(iter,VulnAgeCategry).Formula = "=IF(T" & iter & "<=90," & chr(34) & "0-3 Months" & chr(34) & ",IF(AND(T" & iter & ">90,T" & iter & "<=180)," & chr(34) & "3-6 Months" & chr(34) & ",IF(AND(T" & iter & ">180,T" & iter & "<=365)," & chr(34) & "6-12 Months" & chr(34) & ",IF(AND(T" & iter & ">365,T" & iter & "<=730)," & chr(34) & "1-2 Years" & chr(34) & ",IF(AND(T" & iter & ">730,T" & iter & "<=1824)," & chr(34) & ">2 Years" & chr(34) & ",IF(T" & iter & ">1825," & chr(34) & ">5 Years" & chr(34) & "))))))"


if vulnsheet.Cells(iter,VulnAgeCategry).value = "0-3 Months" then 
	SummarySheet.Cells(vulagecat1,2) = SummarySheet.Cells(vulagecat1,2) + 1
ElseIf vulnsheet.Cells(iter,VulnAgeCategry).value = "3-6 Months" then 
	SummarySheet.Cells(vulagecat2,2) = SummarySheet.Cells(vulagecat2,2) + 1
ElseIf vulnsheet.Cells(iter,VulnAgeCategry).value = "6-12 Months" then 
	SummarySheet.Cells(vulagecat3,2) = SummarySheet.Cells(vulagecat3,2) + 1
ElseIf vulnsheet.Cells(iter,VulnAgeCategry).value = "1-2 Years" then 
	SummarySheet.Cells(vulagecat4,2) = SummarySheet.Cells(vulagecat4,2) + 1
ElseIf vulnsheet.Cells(iter,VulnAgeCategry).value = "2-5 Years" then 
	SummarySheet.Cells(vulagecat5,2) = SummarySheet.Cells(vulagecat5,2) + 1
ElseIf vulnsheet.Cells(iter,VulnAgeCategry).value = ">5 Years" then 
	SummarySheet.Cells(vulagecat6,2) = SummarySheet.Cells(vulagecat6,2) + 1
End If

if vulnsheet.Cells(iter,severity).value = "Critical" then 
	SummarySheet.Cells(Critical,2) = SummarySheet.Cells(Critical,2) + 1
Elseif vulnsheet.Cells(iter,severity).value = "High" then 
	SummarySheet.Cells(High,2) = SummarySheet.Cells(High,2) + 1
Elseif vulnsheet.Cells(iter,severity).value = "Medium" then 
	SummarySheet.Cells(Medium,2) = SummarySheet.Cells(Medium,2) + 1
Elseif vulnsheet.Cells(iter,severity).value = "Low" then 
	SummarySheet.Cells(Low,2) = SummarySheet.Cells(Low,2) + 1
End If
iter = iter + 1
Loop
'
'
''Create a new column - last observed date with a formula
'LastObsDate = GetMaxColumn(msngPatches) + 1
'msngPatches.Cells(1,LastObsDate).value = "Last Observed Date"
'
'PluginPubDate = GetMaxColumn(msngPatches) + 1
'MsgBox PluginPubDate
''Create a new column - plugin pub date with a formula
'msngPatches.Cells(1,PluginPubDate).value = "Plugin Published Date"
'
'VulnAge = GetMaxColumn(msngPatches) + 1
'MsgBox VulnAge
''Create a new column - plugin pub date with a formula
'msngPatches.Cells(1,VulnAge).value = "Vulnerability Age"
'
'VulnAgeCategry = GetMaxColumn(msngPatches) + 1
'MsgBox VulnAgeCategry
''Create a new column - plugin pub date with a formula
'msngPatches.Cells(1,VulnAgeCategry).value = "Vulnerability Age Category"
'
'iter = 2
'Do Until iter = GetMaxRow(msngPatches)
'msngPatches.Cells(iter,LastObsDate).Formula = "=LEFT(M" & iter & ",(12))"
'msngPatches.Cells(iter,PluginPubDate).Formula = "=LEFT(O" & iter & ",(12))"
'msngPatches.Cells(iter,VulnAge).Formula = "=R" & iter & "-S" & iter
'msngPatches.Cells(iter,VulnAgeCategry).Formula = "=IF(T" & iter & "<=90," & chr(34) & "0-3 Months" & chr(34) & ",IF(AND(T" & iter & ">90,T" & iter & "<=180)," & chr(34) & "3-6 Months" & chr(34) & ",IF(AND(T" & iter & ">180,T" & iter & "<=365)," & chr(34) & "6-12 Months" & chr(34) & ",IF(AND(T" & iter & ">365,T" & iter & "<=730)," & chr(34) & "1-2 Years" & chr(34) & ",IF(AND(T" & iter & ">730,T" & iter & "<=1824)," & chr(34) & ">2 Years" & chr(34) & ",IF(T" & iter & ">1825," & chr(34) & ">5 Years" & chr(34) & "))))))"
'
'
'if msngPatches.Cells(iter,VulnAgeCategry).value = "0-3 Months" then 
'	SummarySheet.Cells(vulagecat1,4) = SummarySheet.Cells(vulagecat1,4) + 1
'ElseIf msngPatches.Cells(iter,VulnAgeCategry).value = "3-6 Months" then 
'	SummarySheet.Cells(vulagecat2,4) = SummarySheet.Cells(vulagecat2,4) + 1
'ElseIf msngPatches.Cells(iter,VulnAgeCategry).value = "6-12 Months" then 
'	SummarySheet.Cells(vulagecat3,4) = SummarySheet.Cells(vulagecat3,4) + 1
'ElseIf msngPatches.Cells(iter,VulnAgeCategry).value = "1-2 Years" then 
'	SummarySheet.Cells(vulagecat4,4) = SummarySheet.Cells(vulagecat4,4) + 1
'ElseIf msngPatches.Cells(iter,VulnAgeCategry).value = "2-5 Years" then 
'	SummarySheet.Cells(vulagecat5,4) = SummarySheet.Cells(vulagecat5,4) + 1
'ElseIf msngPatches.Cells(iter,VulnAgeCategry).value = ">5 Years" then 
'	SummarySheet.Cells(vulagecat6,4) = SummarySheet.Cells(vulagecat6,4) + 1
'End If
'
'if msngPatches.Cells(iter,severity).value = "Critical" then 
'	SummarySheet.Cells(Critical,4) = SummarySheet.Cells(Critical,4) + 1
'Elseif msngPatches.Cells(iter,severity).value = "High" then 
'	SummarySheet.Cells(High,4) = SummarySheet.Cells(High,4) + 1
'Elseif msngPatches.Cells(iter,severity).value = "Medium" then 
'	SummarySheet.Cells(Medium,4) = SummarySheet.Cells(Medium,4) + 1
'Elseif msngPatches.Cells(iter,severity).value = "Low" then 
'	SummarySheet.Cells(Low,4) = SummarySheet.Cells(Low,4) + 1
'End If
'iter = iter + 1
'Loop

'---------------------------------------------------------------
'-------------------------REPORTING-----------------------------
'---------------------------------------------------------------

' PLEASE ENSURE ALL CALCULATIONS ARE COMPLETED BEFORE THIS
' THIS IS ONLY FOR CREATING CHARTS IN EXCEL AND COPYING IN TO WORD
														
objWorkbook1.Save

MsgBox "Reporting in progress"
'Add basic contents and headers to Word File
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "18"
objSelection.TypeText "Vulnerability Assessment"
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "13"
objSelection.TypeText "Description"
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "A breakdown of the vulnerabilities based on the severity is highlighted for the London devices in the below graph.  " 

'TODO : INSERT CHART
Set oRange=SummarySheet.Range("A9:B12")
oRange.select
Set oChart=objExcel1.charts
oChart.Add()
Set oMychart=oChart(1)
oMychart.SetSourceData oRange
oMychart.Activate
oMychart.ChartTitle.Text = "Vulnerability Summary"
ActiveChart.ChartStyle = 205
oMychart.ApplyDataLabels 5
oMychart.PlotArea.Fill.Visible=False
oMychart.PlotArea.Border.LineStyle=-4142
oMychart.ChartArea.Fill.Forecolor.SchemeColor=49
oMychart.ChartArea.Fill.Backcolor.SchemeColor=14
oMychart.ChartArea.Fill.TwoColorGradient 1,1
oMychart.ChartTitle.Font.Size=13
oMychart.ChartTitle.Font.ColorIndex=4
'oMychart.Shapes.AddChart2(201, xlColumnClustered).Select

oMychart.FullSeriesCollection(1).Points(1).Format.Fill.Visible = msoTrue
oMychart.FullSeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
oMychart.FullSeriesCollection(1).Points(1).Format.Line.ForeColor.RGB = RGB(192, 0, 0)
oMychart.FullSeriesCollection(1).Points(1).Format.Fill.Transparency = 0
oMychart.FullSeriesCollection(1).Points(1).Format.Fill.Solid
oMychart.FullSeriesCollection(1).Points(2).Format.Fill.Visible = msoTrue
oMychart.FullSeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
oMychart.FullSeriesCollection(1).Points(2).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
oMychart.FullSeriesCollection(1).Points(2).Format.Fill.Transparency = 0
oMychart.FullSeriesCollection(1).Points(2).Format.Fill.Solid
oMychart.FullSeriesCollection(1).Points(3).Format.Fill.Visible = msoTrue
oMychart.FullSeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
oMychart.FullSeriesCollection(1).Points(3).Format.Line.ForeColor.RGB = RGB(255, 192, 0)
oMychart.FullSeriesCollection(1).Points(3).Format.Fill.Transparency = 0
oMychart.FullSeriesCollection(1).Points(3).Format.Fill.Solid
oMychart.FullSeriesCollection(1).Points(4).Format.Fill.Visible = msoTrue
oMychart.FullSeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = RGB(255, 255, 0)
oMychart.FullSeriesCollection(1).Points(4).Format.Line.ForeColor.RGB = RGB(255, 255, 0)
oMychart.FullSeriesCollection(1).Points(4).Format.Fill.Transparency = 0
oMychart.FullSeriesCollection(1).Points(4).Format.Fill.Solid

oMychart.ChartTitle.Text = "Vulnerability Summary"
oMychart.SeriesCollection(1).DataLabels.Font.Size=15
oMychart.SeriesCollection(1).DataLabels.Font.ColorIndex=2
oMychart.ApplyDataLabels 5

oMychart.Select
oMychart.Copy
objWord.Selection.PasteSpecial

objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "TODO : INSERT CHART"
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "A detailed summary of vulnerability data in London Branch is provided in the below attachment."
'TODO : INSERT ATTACHMENT
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "TODO : INSERT ATTACHMENT "
objSelection.TypeParagraph()
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "18"
objSelection.Font.Bold = False
objSelection.TypeText "Missing Operating System Patches"
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "13"
objSelection.TypeText "Description"
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "Patches are fixes provided by the vendor to close the security gaps found on the operating system. Unpatched devices may be easily exploited to gain unauthorized privileges on the system which may lead to further damage being caused by the attacker." 
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "A summary of the missing patches since their release year in London devices is provided in the below depiction:"
'TODO : INSERT CHART
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "TODO : INSERT AGING Report"
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "A detailed summary of missing patches data in London devices can be found in the previously attached file."
objSelection.TypeParagraph()
objSelection.Font.Bold = False
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "A vulnerability ageing summary for London devices is as provided in below graph:"
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = "11"
objSelection.TypeText "TODO : INSERT AGING Report"


'Assets scanned with status

'Vulnerability Summary

'Missing Patches Summary

'Vulnerability Aging Summary

'Missing Patches Aging Summary



'Save and close all opened files
objWorkbook1.Save
objWorkbook1.Close
objExcel1.Quit

objWorkbook2.Save
objWorkbook2.Close
objExcel2.Quit

objWorkbook3.Save
objWorkbook3.Close
objExcel3.Quit

objDoc.SaveAs(GetCurrentDir & "\testdoc.doc")
objWord.Save
objWord.Close
objWord.Quit

MsgBox "Completed"

'----------------------------------------------------------
'------------------HELPER FUNCTIONS -----------------------
'----------------------------------------------------------

' Returns the maximum number of rows in a excel file
'https://stackoverflow.com/questions/29017663/vbscript-to-read-excel-1-how-to-get-the-row-count-of-specific-column-2-to
Function GetMaxColumn(sheet)
intColumn = 2
	Do Until sheet.Cells(1,intColumn).Value = ""
		intColumn = intColumn + 1
	Loop
	GetMaxColumn = intColumn - 1
End Function

' Returns the maximum number of columns in a excel file
Function GetMaxRow(sheet)
intRow = 2
	Do Until sheet.Cells(intRow,1).Value = ""
		intRow = intRow + 1
	Loop
	GetMaxRow = intRow - 1
End Function

Function BrowseForFile()
'@description: Browse for file dialog.
'@author: Jeremy England (SimplyCoded)
  BrowseForFile = CreateObject("WScript.Shell").Exec( _
    "mshta.exe ""about:<input type=file id=f>" & _
    "<script>resizeTo(0,0);f.click();new ActiveXObject('Scripting.FileSystemObject')" & _
    ".GetStandardStream(1).WriteLine(f.value);close();</script>""" _
  ).StdOut.ReadLine()
End Function

Function GetCurrentDir
	GetCurrentDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
End Function
														
														
'**********************************************
'Function: CreateChart()
'Description:Creates a graph in an excel sheet.
'Author:QTP Lab:--A touch of madness
'Website:https://automationlab09.wordpress.com
'Last modified:29/04/2009

'****************************************************
Function CreateChart()

On Error Resume next
Dim oExl,oWrkbk,oWrkst,oMychart

Set oExl=CreateObject("Excel.Application")

With oExl

        .Visible=True

End With

Set oWrkbk=oExl.Workbooks.Add()
Set oWrkst=oWrkbk.Worksheets(1)

With oWrkst
.Cells(1,1)="Critical"
.Cells(2,1)="Very Serious"
.Cells(3,1)="Serious"
.Cells(4,1)="Moderate"
.Cells(5,1)="Mild"

.Cells(1,2)="Bugs Severity"

For i=2 to 5

.Cells(i,2)=i+21

 If i>4 Then
 .Cells(5,2)=9
 End If

Next

End With
Set oRange=oWrkst.UsedRange
oRange.Select
Set oChart=oExl.charts
oChart.Add()
Set oMychart=oChart(1)
oMychart.Activate
oMychart.ChartType=5
oMychart.ApplyDataLabels 5

oMychart.PlotArea.Fill.Visible=False
oMychart.PlotArea.Border.LineStyle=-4142
oMychart.SeriesCollection(1).DataLabels.Font.Size=15
oMychart.SeriesCollection(1).DataLabels.Font.ColorIndex=2

oMychart.ChartArea.Fill.Forecolor.SchemeColor=49
oMychart.ChartArea.Fill.Backcolor.SchemeColor=14
oMychart.ChartArea.Fill.TwoColorGradient 1,1

oMychart.ChartTitle.Font.Size=20
oMychart.ChartTitle.Font.ColorIndex=4
'oWrkbk.Close
Set oExl=Nothing

If err.number<>0 then

 Msgbox "error occurred while drawing..."
 Msgbox err.Description

Else

 Msgbox  "Successfully drawn"

End If

End Function

'Copy pasting from excel to word																
'oMychart.ChartArea.Select
'oMychart.ChartArea.Copy
'objWord.Selection.PasteSpecial

																
