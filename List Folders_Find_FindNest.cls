VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub listFolders()
Dim FileSystem As Object
Dim HostFolder As String

Application.ScreenUpdating = False

HostFolder = "D:\Adobe projects"

Set FileSystem = CreateObject("Scripting.FileSystemObject")
DoFolder FileSystem.GetFolder(HostFolder)

   



End Sub

Sub DoFolder(Folder)
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
             ActiveCell.Value = SubFolder
      ActiveCell.Offset(1, 0).Select
 
    Next
   ' Dim File
   ' For Each File In Folder.Files
   '   ActiveCell.Value = File
   '   ActiveCell.Offset(1, 0).Select
  
   ' Next
End Sub


Sub UseLookIn()


    Dim res As Range
    Dim firstAdress As String
    Dim Ftext As String
    
   Application.ScreenUpdating = False
    
    ActiveSheet.Columns(2).ClearContents
    ActiveSheet.Columns(4).ClearContents
    Ftext = Cells(1, 3)
    Cells(1, 4).Select
    
With ActiveSheet.Range("A1:A5000")

   Set res = .Find(Ftext)
   If Not res Is Nothing Then
   firstAdress = res.Address
   Do
           Cells(res.Row, 2) = "x"
           ActiveCell.Value = res.Row
           ActiveCell.Offset(1, 0).Select
           Set res = .FindNext(res)
    If res Is Nothing Then
    GoTo DoneFinding
    End If
    Loop While res.Address <> firstAdress
  End If
DoneFinding:
End With
    
       Application.ScreenUpdating = True

End Sub
