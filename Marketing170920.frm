VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Save Mails"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "Marketing170920.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public excelFile As String

Private Sub BtnSave_Click()     '******Save button
Dim selectedComb As String
Dim selectedOption As String
Dim lRow As Integer

Dim ExWbk As Workbook
Dim xlSheet As Worksheet
Dim i As Integer
Dim PathName As String
Dim oCtrl As Control



                '*****setting the excel objects
Set ExWbk = GetObject(excelFile)
Set xlSheet = ExWbk.Sheets("Sheet1")
lRow = xlSheet.Range("C1")

selectedComb = ComboBox1.Value

         '*******find the value of the combobox in the excel folder list
For i = 1 To lRow
If xlSheet.Range("A" & i) = selectedComb Then
selectedComb = xlSheet.Range("B" & i)
Exit For
End If
Next

 

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

                    '******* setting the path to save
                    
PathName = selectedComb & "\" & selectedOption & "\"

        
If selectedComb = "" Or selectedOption = "" Then        '*****in case there was no selection
MsgBox "Please Choose costumer folder and sub folder"
Else
Call SaveMessageAsMsg(PathName)             '***** calling the save mail sub
End If


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
Dim ExApp As Excel.Application
Dim ExWbk As Workbook
Dim xlSheet As Worksheet

 

Set ExWbk = GetObject(excelFile)
Set xlSheet = ExWbk.Sheets("Sheet1")

 
 
 'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder("\\server2k11\IBRNet\Marketing Contact Data")                  '************* location  of main folder
i = 1
'loops through each folder in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
    'print folder name
     
   
    xlSheet.Range("A" & i) = objSubFolder.Name
       
    'print folder path
    xlSheet.Range("B" & i) = objSubFolder.Path
    i = i + 1
Next objSubFolder

xlSheet.Range("C1") = i - 1

ExWbk.Save

MsgBox "Folder list was updated, total of " & i - 1 & " folders "

Unload UserForm1
UserForm1.Show


End Sub

Private Sub UserForm_Initialize()           '*****Form initialization
Dim ExWbk As Workbook
Dim xlSheet As Worksheet
Dim i As Integer
Dim lastRow

 excelFile = "C:\Users\User\Documents\Flist.xlsx"     '******* Location of excel file

Set ExWbk = GetObject(excelFile)
Set xlSheet = ExWbk.Sheets("Sheet1")

lastRow = xlSheet.Range("C1")

For i = 1 To lastRow                    '***** populating combobox from excel file

ComboBox1.AddItem xlSheet.Range("A" & i)

Next

End Sub


Private Sub UserForm_Terminate()            '****** sub for Closing the form

'Dim ExWbk As Workbook
'
'Set ExWbk = GetObject(excelFile)
'ExWbk.Save
'ExWbk.Close

End Sub
