VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Save Mails"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "Marketing170930.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public PathName As String       'public variables
Const ExcelFile = "C:\Users\Cohen\Documents\Flist.xlsx"  '"C:\Users\User\Documents\Flist.xlsx"
Const MAIN_FOLDER = "C:\Drive D"   '"\\server2k11\IBRNet\Marketing Contact Data"


Private Sub BtnSave_Click()     '******Save button
Dim selectedComb As String
Dim selectedOption As String
Dim i As Integer
Dim oCtrl As Control
'Dim lRow As Integer
'Dim ExWbk As Workbook
'Dim xlSheet As Worksheet

 Me.CommandButton1.Visible = False
selectedComb = ComboBox1.Value


 '***** Checking which option button was selected in Frame1
    For Each oCtrl In Frame1.Controls
        '***** Try only option buttons
        If TypeName(oCtrl) = "OptionButton" Then
            '***** Which one is checked?
            If oCtrl.Value = True Then
                '***** What's the caption?
                selectedOption = oCtrl.Caption
                Exit For
            End If
        End If
    Next
    
      '*****in case there was no selection, exit sub
 
 If selectedComb = "" Or selectedOption = "" Then
     MsgBox "Please Choose costumer folder and sub folder"
     Exit Sub
 End If
 
   

                    '******* setting the path to save
                    
PathName = MAIN_FOLDER & "\" & selectedComb & "\" & selectedOption & "\"


'***** in case folder doesn't exist
  
  If Dir(PathName, vbDirectory) = "" Then
    MsgBox "The folder does not exist"
    Exit Sub
  End If

'
'                '*****setting the excel objects
'Set ExWbk = GetObject(excelFile)
'Set xlSheet = ExWbk.Sheets("Sheet1")
'lRow = xlSheet.Range("C1")
'
'
'         '*******find the value of the combobox in the excel folder list
'For i = 1 To lRow
'If xlSheet.Range("A" & i) = selectedComb Then
'selectedComb = xlSheet.Range("B" & i)
'Exit For
'End If
'Next



 If selectedOption = "Submittals" Then
 Call SaveAttachments(PathName)                 '*****calling the submittals sub
  Else
   Call SaveMessageAsMsg(PathName)             '***** calling the save mail sub
 End If


Me.CommandButton1.Visible = True


End Sub

Public Sub SaveMessageAsMsg(PathName As String)     '***** Saving the mails to the path
  
  On Error GoTo ErrorHandler
  
  Dim oMail As Outlook.MailItem
  Dim objItem As Object
  Dim sPath As String
  Dim dtDate As Date
  Dim sName As String
  Dim sBody As String
  Dim Scount As Integer
  Dim i As Integer
  
  
   i = 1
    
   Label2.Visible = True
   
   Scount = ActiveExplorer.Selection.Count

   For Each objItem In ActiveExplorer.Selection
    If objItem.MessageClass = "IPM.Note" Then
     Set oMail = objItem
   
     sName = oMail.Subject
     sBody = oMail.Body
  
     ReplaceCharsForFileName sName, "-"
 
     dtDate = oMail.ReceivedTime
     sName = Format(dtDate, "yy mm dd", vbUseSystemDayOfWeek, _
     vbUseSystem) & "-" & Format(dtDate, "hhnnss", vbUseSystemDayOfWeek, vbUseSystem) & "-" & sName & ".msg"

     oMail.SaveAs PathName & sName, olMSG
   
    Label2.Caption = i & "/" & Scount & "   Mails were saved"
    DoEvents
   
   End If
   i = i + 1
  Next

 Exit Sub 'skip error handler
  
ErrorHandler:
  MsgBox "An Error was Occur"
  
End Sub

Private Sub ReplaceCharsForFileName(sName As String, _
  sChr As String _
)                                                           '******Replacing charaters in mail filename
  sName = Replace(sName, "'", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
End Sub


Private Sub BtnUpdateFolder_Click()  '***** Updating folder list in excel file

Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim i As Integer
Dim x As Integer
Dim ExApp As Excel.Application
Dim ExWbk As Workbook
Dim xlSheet As Worksheet


Set ExWbk = GetObject(ExcelFile)
Set xlSheet = ExWbk.Sheets("Sheet1")
 
 
 'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
'Set objFolder = objFSO.GetFolder() '***** location  of main folder
Set objFolder = objFSO.GetFolder(MAIN_FOLDER) '***** location  of main folder

x = objFolder.subfolders.Count  'count how many subfolders is the main folder
i = 1
'loops through each folder in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
    xlSheet.Range("A" & i) = objSubFolder.Name
    'print folder path
    'xlSheet.Range("B" & i) = objSubFolder.Path  'omited if the main folder path is constant
    Label3.Caption = i & "  /  " & x    'counter for the folders
    i = i + 1
Next objSubFolder

xlSheet.Range("C1") = i - 1

ExWbk.Save

MsgBox "Folder list was updated, total of " & i - 1 & " folders "

Unload UserForm2
UserForm2.show


End Sub


Private Sub CommandButton1_Click()                          '********  Show folder button
Call Shell("explorer.exe" & " " & PathName, vbNormalFocus)
Me.CommandButton1.Visible = False
End Sub


Private Sub UserForm_Initialize()           '*****Form initialization
Dim ExWbk As Workbook
Dim xlSheet As Worksheet
Dim i As Integer
Dim lastRow


Me.CommandButton1.Visible = False

Set ExWbk = GetObject(ExcelFile)
Set xlSheet = ExWbk.Sheets("Sheet1")

lastRow = xlSheet.Range("C1")

For i = 1 To lastRow                    '***** populating combobox from excel file
 ComboBox1.AddItem xlSheet.Range("A" & i)
Next

End Sub


Public Sub SaveAttachments(PathName)                        '**********Save attachements
Dim objOL As Outlook.Application                    'creates a folder with the name of the mails's subject
Dim objMsg As Outlook.MailItem                      'copy all the atachments of the mail into that folder
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderPath As String
Dim strDeletedFiles As String
Dim strSubject As String
Dim dtDate As Date
Dim Scount As Integer
Dim j As Integer



   j = 1
   Label2.Visible = True
   Scount = ActiveExplorer.Selection.Count


    ' Get the path to your My Documents folder
   ' strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
   ' On Error Resume Next

    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")

    ' Get the collection of selected objects.
    Set objSelection = objOL.ActiveExplorer.Selection

    
    ' Check each selected item for attachments.
    For Each objMsg In objSelection

    strFolderPath = PathName      'the main folder's destination where the folder will be created
    
    
    Set objAttachments = objMsg.Attachments
    lngCount = objAttachments.Count         'number of attachements
    strSubject = objMsg.Subject             'getting the name of the mail's subject
    ReplaceCharsForFileName strSubject, "-"      'calling the sub to replce characters in the name of the folder
    dtDate = objMsg.ReceivedTime            ' getting the date of the mail sent/received
    strFolderPath = strFolderPath & "\" & Format(dtDate, "yy mm dd_", vbUseSystemDayOfWeek, _
    vbUseSystem) & Format(dtDate, "hhnnss_", vbUseSystemDayOfWeek, vbUseSystem) & strSubject ' the final folder name: Date_time_subject
    
    
    If Dir(strFolderPath, vbDirectory) = "" Then
    MkDir strFolderPath     'creating the folder
    Else
    MsgBox "Folder is already exists"
    Exit Sub
    End If
    
    Label2.Caption = j & "/" & Scount & "   Mails were saved"
    DoEvents

    j = j + 1
        
    If lngCount > 0 Then
         
     For i = lngCount To 1 Step -1  'loops through the attachments
    
     ' Get the file name.
      strFile = objAttachments.Item(i).filename
    
     ' Combine with the path
      strFile = strFolderPath & "\" & strFile
    
     ' Save the attachment as a file.
      objAttachments.Item(i).SaveAsFile strFile
    
    Next i
    End If
    
    Next
    
ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

End Sub



