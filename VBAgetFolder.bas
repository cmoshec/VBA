Attribute VB_Name = "Module1"

Public Sub GetfolderNamePath()
Dim pathName As String
Dim folderName As String

pathName = Getfolder("")
folderName = pathOfFile(pathName)

Cells(1, 1) = pathName
Cells(2, 1) = folderName

End Sub

Function Getfolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
    
End With
NextCode:
Getfolder = sItem
Set fldr = Nothing
End Function

Function pathOfFile(fileName As String) As String
    Dim posn As Integer
    Dim strLen As Integer
    
    'strlen = the length of the srtring
    strLen = Len(fileName)
    'posn=position of the first apearence of "\" from the end of the string
    posn = InStrRev(fileName, "\")
    If posn > 0 Then
        pathOfFile = Right$(fileName, strLen - posn)  'cuting the string at position (strlen-posn) from the right
    Else
        pathOfFile = ""
    End If
End Function
