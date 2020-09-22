Attribute VB_Name = "mdlMain"
Option Explicit
'FUNCTION TO WRITE TO INI
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Global DB As Database
Global RS As Recordset
Global WS As Workspace
Global Msg As String
Global dbPath As String
Global dbFile As String
Global IniFile As String
Global Source As String
Global Dest As String
Global ReportShow As Boolean
Global ReportFile As String
Global Printing As String
Global FileName1
Global FileName2

Dim FileName3
Dim FileName4
Dim FileNameOpen As String

Sub Main()
On Error GoTo ErrorMain
 
 dbPath = App.Path
 If Right$(dbPath, 1) <> "\" Then dbPath = dbPath & "\"
 IniFile = dbPath & "data\dab.ini"
 AddLastFile2Menu

If Command = "/new" Then
MakeNewFile
Exit Sub
End If

If Command = "" Then
    Load mdiMain
    mdiMain.Show

ElseIf InStr(1, Command$, "/open") Then
 FileName3 = Split(Command$, "/open")
 FileName4 = FileName3(UBound(FileName3))
 FileNameOpen = Trim$(FileName4)
 OpenFile
 Exit Sub
ElseIf InStr(1, Command$, "print") Then
 FileName3 = Split(Command$, "/print")
 FileName4 = FileName3(UBound(FileName3))
 FileNameOpen = Trim$(FileName4)
 PrintReport
Else
MsgBox "No falide command" & vbCrLf & Command$
End If
Exit Sub

ErrorMain:
  Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error MainLoad!!")
End Sub

'FUNCTION TO ADD THE LAST FILE TO THE MENU
Function AddLastFile2Menu()
If ReadINI("LastFile", "1", IniFile) <> "" Then
 mdiMain.mnuLastLine.Visible = True
 mdiMain.mnuLast1.Visible = True
 mdiMain.mnuLast1.Caption = "&1 " & ReadINI("LastFile", "1", IniFile)
If ReadINI("LastFile", "2", IniFile) <> "" Then
 mdiMain.mnuLastLine.Visible = True
 mdiMain.mnuLast2.Visible = True
 mdiMain.mnuLast2.Caption = "&2 " & ReadINI("LastFile", "2", IniFile)
If ReadINI("LastFile", "3", IniFile) <> "" Then
 mdiMain.mnuLastLine.Visible = True
 mdiMain.mnuLast3.Visible = True
 mdiMain.mnuLast3.Caption = "&3 " & ReadINI("LastFile", "3", IniFile)
If ReadINI("LastFile", "4", IniFile) <> "" Then
 mdiMain.mnuLastLine.Visible = True
 mdiMain.mnuLast4.Visible = True
 mdiMain.mnuLast4.Caption = "&4 " & ReadINI("LastFile", "4", IniFile)
   End If
  End If
 End If
End If
End Function

'FUNCTION TO CHECK WHAT TO WRITE TO THE INI FOR LAST FILE ORDERING
Function AddLastFile2Ini()
Dim Last0 As String
Dim Last1 As String
Dim Last2 As String
Dim Last3 As String
Dim Last4 As String

If dbFile = "" Then
 Exit Function
Else
Last0 = dbFile
Last1 = ReadINI("LastFile", "1", IniFile)
Last2 = ReadINI("LastFile", "2", IniFile)
Last3 = ReadINI("LastFile", "3", IniFile)
Last4 = ReadINI("LastFile", "4", IniFile)
End If

If Last0 Like Last1 Then
 Exit Function
ElseIf Last0 Like Last2 Then
 WriteINI "lastFile", "1", Last2, IniFile
 WriteINI "lastFile", "2", Last1, IniFile
 WriteINI "lastFile", "3", Last3, IniFile
 WriteINI "lastFile", "4", Last4, IniFile
ElseIf Last0 Like Last3 Then
 WriteINI "lastFile", "1", Last3, IniFile
 WriteINI "lastFile", "2", Last1, IniFile
 WriteINI "lastFile", "3", Last2, IniFile
 WriteINI "lastFile", "4", Last4, IniFile
ElseIf Last0 Like Last4 Then
 WriteINI "lastFile", "1", Last4, IniFile
 WriteINI "lastFile", "2", Last1, IniFile
 WriteINI "lastFile", "3", Last2, IniFile
 WriteINI "lastFile", "4", Last3, IniFile
ElseIf Not (Last0 Like Last1) Or Not (Last0 Like Last2) Or Not (Last0 Like Last3) Or Not (Last0 Like Last4) Then
 WriteINI "lastFile", "1", Last0, IniFile
 WriteINI "lastFile", "2", Last1, IniFile
 WriteINI "lastFile", "3", Last2, IniFile
 WriteINI "lastFile", "4", Last3, IniFile
End If

End Function

'FUNCTION TO READ FROM AN INI FILE
Function ReadINI(RMain, RKey, File As String) As String
 Dim RTemp As String
 RTemp = String(255, Chr(0))
 ReadINI = Left(RTemp, GetPrivateProfileString(RMain, ByVal RKey, "", RTemp, Len(RTemp), File))
End Function

'FUNCTION TO WRITE TO AN INI FILE
Function WriteINI(WMain As String, WKey As String, WString As String, WFile As String)
 Call WritePrivateProfileString(WMain, WKey, WString, WFile)
End Function

'FUNCTION TO MAKE AN NEW FILE IF THE /NEW OPTION IS USED
Function MakeNewFile()
On Error GoTo ErrorMakeNewFile

 With mdiMain.OpenSave
  .FileName = "New.dab"
  .Filter = "Address Book (*.dab)|*.dab"
  .InitDir = dbPath & "data"
  .DialogTitle = "New Dextra's Address Book file"
  .ShowSave
  dbFile = .FileName
 End With
 
 If Dir(dbFile) <> "" Then
  Msg = MsgBox("The file:" & vbCrLf & dbFile & vbCrLf & " already existe" & vbCrLf & "Do you want to Overwrite?", vbExclamation + vbYesNo, "File existe!")
   Select Case Msg
    Case vbNo
     MakeNewFile
     Exit Function
    Case vbYes
     GoTo Copy
   End Select
 End If
 
Copy:
 Source = dbPath & "data" & "\new.dak"
 MsgBox Source
 Dest = dbFile
 FileCopy Source, Dest
 mdiMain.StatusBar.Panels(1).Text = dbFile
 FileName1 = Split(dbFile, "\")
 FileName2 = FileName1(UBound(FileName1))
 mdiMain.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
 mdiMain.mnuClose.Enabled = True
 mdiMain.mnuReport.Enabled = False
 frmMain.Show
Exit Function

ErrorMakeNewFile:
 If Err.Number = 32755 Then
 Unload mdiMain
 Exit Function
 End If
  Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error MakeNewFile!!")
  Unload mdiMain
End Function
'FUNCTION TO OPEN AN FILE WHEN THE /OPEN COMMAND LINE IS USED
Function OpenFile()
On Error GoTo ErrorOpenFile

   dbFile = FileNameOpen
   mdiMain.StatusBar.Panels(1).Text = dbFile
   FileName1 = Split(dbFile, "\")
   FileName2 = FileName1(UBound(FileName1))
   mdiMain.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
   mdiMain.mnuClose.Enabled = True
   mdiMain.mnuReport.Enabled = True
   frmMain.Show
Exit Function

ErrorOpenFile:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error OpenFile!!")
 Unload mdiMain
End Function

'FUNCTION TO PRINT THE DATAREPORT WHEN THE "/PRINT" COMMAND LINE IS USED
Function PrintReport()
On Error GoTo ErrorPrintReport

Printing = "Yes"
ReportFile = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & FileNameOpen & Chr(34) & ";Persist Security Info=False"
MainEnv.DbCon.ConnectionString = ReportFile
Report.Title = FileNameOpen
Report.PrintReport True
Unload mdiMain
Exit Function

ErrorPrintReport:
If Err.Number = 0 Then Exit Function
Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error '/Print' Command")
End Function
