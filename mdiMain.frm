VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Dextra's Address Book"
   ClientHeight    =   7590
   ClientLeft      =   2400
   ClientTop       =   2385
   ClientWidth     =   10335
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog OpenSave 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7275
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   556
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13829
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "21-1-2001"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuLastLine 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLast1 
         Caption         =   "&1 "
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLast2 
         Caption         =   "&2"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLast3 
         Caption         =   "&3"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLast4 
         Caption         =   "&4"
         Visible         =   0   'False
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Extra"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Report"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo ErrorMDIFormUnload

If Printing = "Yes" Then
Cancel = False
Unload Me
Exit Sub
End If

Msg = MsgBox("Ar you sure you want to exit Dextra's Address Book?", vbInformation + vbYesNo, "Exit?")
Select Case Msg
Case vbNo
 Cancel = True
 If dbFile <> "" Then
 frmMain.Show
 mdiMain.mnuClose.Enabled = True
 mdiMain.mnuReport.Enabled = True
 End If
 Exit Sub
Case vbYes
 AddLastFile2Ini
 Unload Me
End Select
Exit Sub

ErrorMDIFormUnload:
Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Unloading MDIForm")
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

'MENU ACTIONS
Private Sub mnuClose_Click()
On Error GoTo ErrorMnuClose

If ReportShow Then
Unload Report
ReportShow = False
Exit Sub
End If

 Unload frmMain
 StatusBar.Panels(1).Text = ""
 Me.Caption = "Dextra's Address Book"
 DB.Close
 MainEnv.DbCon.Cancel
 mnuClose.Enabled = False
 mnuReport.Enabled = False
 AddLastFile2Ini
 AddLastFile2Menu
Exit Sub

ErrorMnuClose:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error MnuClose!")
End Sub

Private Sub mnuReport_Click()

ReportFile = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & dbFile & Chr(34) & ";Persist Security Info=False"
Unload MainEnv
MainEnv.DbCon.ConnectionString = ReportFile
Report.Title = dbFile
ReportShow = True
Report.Show
End Sub

Private Sub mnuExit_Click()
On Error GoTo ErrorMnuExit
Unload Me
Exit Sub
ErrorMnuExit:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error MnuExit!")
End Sub

Private Sub mnuLast1_Click()

If mnuClose.Enabled = True Then
 Msg = MsgBox("There already is an Address Book opend" & vbCrLf & "Do you want the curent Address book to close?", vbQuestion + vbYesNo, "Close curent Address Book!")
Select Case Msg
 Case vbYes
  Unload frmMain
  StatusBar.Panels(1).Text = ""
  Me.Caption = "Dextra's Address Book"
  DB.Close
  mnuClose.Enabled = False
  mnuReport.Enabled = False
  GoTo OpenLast1
 Case vbNo
  Exit Sub
 End Select
End If

OpenLast1:
   dbFile = ReadINI("lastfile", "1", IniFile)
   StatusBar.Panels(1).Text = dbFile
   FileName1 = Split(dbFile, "\")
   FileName2 = FileName1(UBound(FileName1))
   Me.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
   mnuClose.Enabled = True
   mnuReport.Enabled = True
   frmMain.Show
End Sub

Private Sub mnuLast2_Click()
If mnuClose.Enabled = True Then
 Msg = MsgBox("There already is an Address Book opend" & vbCrLf & "Do you want the curent Address book to close?", vbQuestion + vbYesNo, "Close curent Address Book!")
Select Case Msg
 Case vbYes
  Unload frmMain
  StatusBar.Panels(1).Text = ""
  Me.Caption = "Dextra's Address Book"
  DB.Close
  mnuClose.Enabled = False
  mnuReport.Enabled = False
  GoTo OpenLast2
 Case vbNo
  Exit Sub
 End Select
End If

OpenLast2:
   dbFile = ReadINI("lastfile", "2", IniFile)
   StatusBar.Panels(1).Text = dbFile
   FileName1 = Split(dbFile, "\")
   FileName2 = FileName1(UBound(FileName1))
   Me.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
   mnuClose.Enabled = True
   mnuReport.Enabled = True
   frmMain.Show

End Sub

Private Sub mnuLast3_Click()
If mnuClose.Enabled = True Then
 Msg = MsgBox("There already is an Address Book opend" & vbCrLf & "Do you want the curent Address book to close?", vbQuestion + vbYesNo, "Close curent Address Book!")
Select Case Msg
 Case vbYes
  Unload frmMain
  StatusBar.Panels(1).Text = ""
  Me.Caption = "Dextra's Address Book"
  DB.Close
  mnuClose.Enabled = False
  mnuReport.Enabled = False
  GoTo OpenLast3
 Case vbNo
  Exit Sub
 End Select
End If

OpenLast3:
   dbFile = ReadINI("lastfile", "3", IniFile)
   StatusBar.Panels(1).Text = dbFile
   FileName1 = Split(dbFile, "\")
   FileName2 = FileName1(UBound(FileName1))
   Me.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
   mnuClose.Enabled = True
   mnuReport.Enabled = True
   frmMain.Show

End Sub

Private Sub mnuLast4_Click()
If mnuClose.Enabled = True Then
 Msg = MsgBox("There already is an Address Book opend" & vbCrLf & "Do you want the curent Address book to close?", vbQuestion + vbYesNo, "Close curent Address Book!")
Select Case Msg
 Case vbYes
  Unload frmMain
  StatusBar.Panels(1).Text = ""
  Me.Caption = "Dextra's Address Book"
  DB.Close
  mnuClose.Enabled = False
  mnuReport.Enabled = False
  GoTo OpenLast4
 Case vbNo
  Exit Sub
 End Select
End If

OpenLast4:
   dbFile = ReadINI("lastfile", "4", IniFile)
   StatusBar.Panels(1).Text = dbFile
   FileName1 = Split(dbFile, "\")
   FileName2 = FileName1(UBound(FileName1))
   Me.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
   mnuClose.Enabled = True
   mnuReport.Enabled = True
   frmMain.Show

End Sub

Private Sub mnuNew_Click()
On Error GoTo ErrorNew

If mnuClose.Enabled = True Then
Msg = MsgBox("There already is an Address Book opend" & vbCrLf & "Do you want the curent Address book to close?", vbQuestion + vbYesNo, "Close curent Address Book!")
Select Case Msg
 Case vbYes
  Unload frmMain
  StatusBar.Panels(1).Text = ""
  Me.Caption = "Dextra's Address Book"
  AddLastFile2Ini
  AddLastFile2Menu
  DB.Close
  mnuClose.Enabled = False
  mnuReport.Enabled = False
  GoTo NewOpen
 Case vbNo
  Exit Sub
 End Select
End If

NewOpen:
 With OpenSave
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
     Call mnuNew_Click
     Exit Sub
    Case vbYes
     GoTo Copy
   End Select
 End If
 
Copy:
 Source = dbPath & "data" & "\new.dak"
 Dest = dbFile
 FileCopy Source, Dest
 StatusBar.Panels(1).Text = dbFile
 FileName1 = Split(dbFile, "\")
 FileName2 = FileName1(UBound(FileName1))
 Me.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
 mnuClose.Enabled = True
 mnuReport.Enabled = True
 frmMain.Show
Exit Sub

ErrorNew:
 If Err.Number = 32755 Then
  mnuClose.Enabled = False
  mnuReport.Enabled = False
 Exit Sub
 End If
  Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error New!!")
End Sub

Private Sub mnuOpen_Click()
On Error GoTo ErrorOpen

If mnuClose.Enabled = True Then
 Msg = MsgBox("There already is an Address Book opend" & vbCrLf & "Do you want the curent Address book to close?", vbQuestion + vbYesNo, "Close curent Address Book!")
Select Case Msg
 Case vbYes
  Unload frmMain
  StatusBar.Panels(1).Text = ""
  Me.Caption = "Dextra's Address Book"
  AddLastFile2Ini
  AddLastFile2Menu
  DB.Close
  mnuClose.Enabled = False
  mnuReport.Enabled = False
  GoTo SaveOpen
 Case vbNo
  Exit Sub
 End Select
End If

SaveOpen:
  With OpenSave
   .FileName = ""
   .Filter = "Address Book (*.dab)|*.dab"
   .InitDir = dbPath & "data"
   .DialogTitle = "Open an Dextra's Address Book file"
   .ShowOpen
   dbFile = .FileName
   StatusBar.Panels(1).Text = dbFile
   FileName1 = Split(dbFile, "\")
   FileName2 = FileName1(UBound(FileName1))
   Me.Caption = "Dextra's Address Book " & "[" & FileName2 & "]"
  End With
  mnuClose.Enabled = True
  mnuReport.Enabled = True
  frmMain.Show
Exit Sub

ErrorOpen:
 If Err.Number = 32755 Then
  mnuClose.Enabled = False
  mnuReport.Enabled = False
 Exit Sub
 End If
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Opening!!")
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show
End Sub
