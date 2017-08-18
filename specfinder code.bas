Attribute VB_Name = "Module1"
Sub SpecFinder()
Attribute SpecFinder.VB_ProcData.VB_Invoke_Func = " \n14"

Dim rng As Range
 Dim chtChart As Chart

'ask for template
'myValue = InputBox("Give me some input")

'*******copy table*******
getbook = ActiveWorkbook.Name
 
    Windows("specFinder.xlsm").Activate
    Range("E7:M12").Select
    Selection.Copy
    Windows(getbook).Activate
    Range("E14").Select
    ActiveSheet.Paste

'********finding OD values and wavelength picks*******

Set rng = Range("b501:b505")
ph = Application.WorksheetFunction.Max(rng)
Cells(15, 10) = ph
'find ph wavelength pick
For i = 501 To 505
If Cells(i, 2) = ph Then
ph_nm = Cells(i, 1)
max_loc = i
End If
Next
Cells(15, 9) = ph_nm
'Cells(15, 11) = Cells(max_loc, 3)

Set rng = Range("b549:b549")
pf = Application.WorksheetFunction.Max(rng)
Cells(16, 10) = pf
'find pf wavelength pick
For i = 549 To 549
If Cells(i, 2) = pf Then
pf_nm = Cells(i, 1)
max_loc = i
End If
Next
Cells(16, 9) = pf_nm
'Cells(16, 11) = Cells(max_loc, 3)

Set rng = Range("b664:b664")
bc = Application.WorksheetFunction.Max(rng)
Cells(17, 10) = bc
'find bc wavelength pick
For i = 664 To 664
If Cells(i, 2) = bc Then
bc_nm = Cells(i, 1)
max_loc = i
End If
Next
Cells(17, 9) = bc_nm
'Cells(17, 11) = Cells(max_loc, 3)

' *****chart*****

sheetname = ActiveSheet.Name
  
    
   y1 = "='" + sheetname + "'!$B$417:$B$817"
   y2 = "='" + sheetname + "'!$C$417:$C$817"
   x = "='" + sheetname + "'!$A$417:$A$817"
   
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlLine
    
     With ActiveChart
            .ChartType = xlLine

            '' Remove any series created with the chart
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete
            Loop
            End With
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).Values = y1
    ActiveChart.SeriesCollection(1).Name = "OD 1"
    ActiveChart.SeriesCollection(1).XValues = x
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Values = y2
    ActiveChart.SeriesCollection(2).Name = "OD 2"
  

'********graph range selection for spectrum generation********

'Range("B417", "B817").Select
'Selection.Copy
'Windows("Spectrum.xlsx").Activate
'ActiveSheet.Paste



'****** copy data to TempResults file ***********
'Range("L15", "L17").Select
'Selection.Copy
'Windows("TempResults.xlsx").Activate
'col = ActiveCell.Column
'Cells(1, col + 1) = sheetname
'Cells(2, col + 1).Select


'Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
'        xlNone, SkipBlanks:=False, Transpose:=False




End Sub
