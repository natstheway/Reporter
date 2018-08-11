Set objExcel = CreateObject("Excel.Application")
Set objReadWorkbook = objExcel.Workbooks.Open("E:\Reporter Project\test.xlsx")
Set oExcelReadWorkSheet = objReadWorkbook.Worksheets(1)
objExcel.Visible = True

Sub GraphCreate ()
   objExcel.ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
   set xlRange = objExcel.ActiveSheet.Range("D1:D70")
   objExcel.ActiveChart.SetSourceData(xlRange)
End Sub

GraphCreate
objReadWorkbook.SaveAs("E:\Reporter Project\test2.xlsx"),-4143
objExcel.Quit
