VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Sub SaveMessageAsMsg()
  Dim oMail As Outlook.MailItem
  Dim objItem As Object
  Dim sPath As String
  Dim dtDate As Date
  Dim sName As String
  Dim enviro As String
 
    enviro = CStr(Environ("USERPROFILE"))
   For Each objItem In ActiveExplorer.Selection
   If objItem.MessageClass = "IPM.Note" Then
    Set oMail = objItem
    
   
  sName = oMail.Subject
  sName = sName + "_"
  ReplaceCharsForFileName sName, "-"
 
  dtDate = oMail.ReceivedTime
  sName = Format(dtDate, "yy-mm-dd", vbUseSystemDayOfWeek, _
    vbUseSystem) & Format(dtDate, "-hhnnss", _
    vbUseSystemDayOfWeek, vbUseSystem) & "-" & sName & ".msg"
     
    sPath = enviro & "\Documents\"
  Debug.Print sPath & sName
  oMail.SaveAs sPath & sName, olMSG
  
  End If
  Next
  
End Sub
 
Private Sub ReplaceCharsForFileName(sName As String, _
  sChr As String _
)
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
