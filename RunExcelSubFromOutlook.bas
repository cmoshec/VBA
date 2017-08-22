Attribute VB_Name = "Module1"
Sub runExcelMacro()
 
 Dim ExApp As Excel.Application
 Dim ExWbk As Workbook
 
 Set ExApp = New Excel.Application
 Set ExWbk = ExApp.Workbooks.Open("D:\VBSamples\VBAfind.xlsm")
 'ExApp.Visible = True

 ExWbk.Application.Run "ThisWorkbook.UseLookIn"

 ExApp.Visible = True
 'ExWbk.Close 'SaveChanges:=True
End Sub
