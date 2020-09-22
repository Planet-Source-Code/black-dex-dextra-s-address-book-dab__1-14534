VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Address Book"
   ClientHeight    =   5490
   ClientLeft      =   285
   ClientTop       =   570
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9570
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   255
      Left            =   5400
      TabIndex        =   40
      Top             =   1320
      Width           =   1575
      Begin VB.OptionButton dbFam 
         Caption         =   "Yes"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   42
         Top             =   0
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton dbFam 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   255
      Left            =   1560
      TabIndex        =   36
      Top             =   1320
      Width           =   2535
      Begin VB.OptionButton dbSex 
         Caption         =   "Male"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton dbSex 
         Caption         =   "Female"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   38
         Top             =   0
         Width           =   855
      End
      Begin VB.OptionButton dbSex 
         Caption         =   "Unknown"
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   37
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.TextBox dbState 
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Text            =   "dbState"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   3720
   End
   Begin VB.TextBox dbEMail 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Text            =   "dbEMail"
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox dbZipCode 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Text            =   "dbZipCode"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox dbAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "dbAddress"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox dbCountry 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Text            =   "dbCountry"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox dbCPhone 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "dbCPhone"
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox dbWPhone 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Text            =   "dbWPhone"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox dbMName 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "dbMName"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton UpdateBtn 
      Caption         =   "Update"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton NewBtn 
      Caption         =   "New"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton DeleteBtn 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton AddBtn 
      Caption         =   "Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton NextBtn 
      Caption         =   "Next"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton PrevBtn 
      Caption         =   "Prev"
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   4920
      Width           =   855
   End
   Begin VB.ListBox dbNameList 
      Height          =   5130
      ItemData        =   "frmMain.frx":0000
      Left            =   7200
      List            =   "frmMain.frx":0007
      TabIndex        =   23
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox dbHPhone 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Text            =   "dbHPhone"
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox dbCity 
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Text            =   "dbCity"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox dbLName 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "dbLName"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox dbFName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "dbFName"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox dbBirthDay 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM-dd-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1043
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Text            =   "dbBirthDay"
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   6960
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label15 
      Caption         =   "State:"
      Height          =   255
      Left            =   4200
      TabIndex        =   35
      Top             =   2040
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label14 
      Caption         =   "Birthday:"
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "E-Mail Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Family:"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Zip Code:"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Home Address:"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Country:"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Cellular Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Work Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Middle Name:"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Aantal 
      Alignment       =   2  'Center
      Caption         =   "Aantal"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "Phone Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "City:"
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecNum As Long
Dim CurRec As Long
Dim I As Long
Dim Index As Long
Dim LoopIndex As Long
Dim dbSexField As String
Dim dbFamField As String

'THE MAIN PART OF THIS APP
Private Sub Form_Load()
LoopIndex = 0
On Error GoTo Error_Load

Me.Top = 0
Me.Left = 0
Set WS = DBEngine.Workspaces(0)
Set DB = WS.OpenDatabase(dbFile)
Set RS = DB.OpenRecordset("main", dbOpenTable)

RecNum = RS.RecordCount
 If RecNum = 0 Then
     CurRec = 0
      dbSex(3).Value = True
      dbFam(1).Value = True
      dbFName.Text = ""
      dbLName.Text = ""
      dbMName.Text = ""
      dbBirthDay.Text = ""
      dbAddress.Text = ""
      dbZipCode.Text = ""
      dbState.Text = ""
      dbCity.Text = ""
      dbCountry.Text = ""
      dbHPhone.Text = ""
      dbWPhone.Text = ""
      dbCPhone.Text = ""
      dbEMail.Text = ""
      dbNameList.Clear
      PrevBtn.Enabled = False
      NextBtn.Enabled = False
      NewBtn.Enabled = False
      UpdateBtn.Enabled = False
      DeleteBtn.Enabled = False
      Aantal.Caption = "Record " & CurRec & " of " & "0"
     Exit Sub
 End If
 CurRec = 1
 Reload
 RS.MoveFirst
Exit Sub

Error_Load:
 If Err.Number = 3044 Then
  Msg = MsgBox("This path: " & vbCrLf & dbFile & vbCrLf & "Does not existe please check if path existe!", vbCritical, "No Path")
  Timer1.Enabled = True
  Exit Sub
 End If
 If Err.Number = 3024 Then
   Msg = MsgBox("Could not find file:" & vbCrLf & dbFile & vbCrLf & "Please check if file existe!", vbCritical, "No File")
   Timer1.Enabled = True
   Exit Sub
 End If
 If Err.Number = 3343 Then
   Msg = MsgBox("This is not an right Address Book File." & vbCrLf & "Please pik an other file.", vbInformation, "Not an right file")
   Timer1.Enabled = True
   Exit Sub
 End If
 If Err.Number = 3261 Then
   Msg = MsgBox("The File:" & vbCrLf & dbFile & vbCrLf & "Is already active.", vbInformation, "Error opening Address Book")
   Timer1.Enabled = True
   Exit Sub
 End If
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Loading!!")
  Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorUnload

 mdiMain.mnuClose.Enabled = False
 mdiMain.mnuReport.Enabled = False
Exit Sub

ErrorUnload:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Unload!!")
End Sub

'HERE IS AN FUNCTION TO CHECK IF ALL FIELDS AR FILLED
Function CheckFields()
On Error GoTo ErrorCheckFields

If Len(dbFName.Text) = 0 Then
 Msg = MsgBox("The field [First Name] can't be empty", vbExclamation, "Error Field")
 dbFName.SetFocus
ElseIf Len(dbLName.Text) = 0 Then
 Msg = MsgBox("The field [Last Name] can't be empty", vbExclamation, "Error Field")
 dbLName.SetFocus
ElseIf Len(dbAddress.Text) = 0 Then
 Msg = MsgBox("The field [Home Address] can't be empty", vbExclamation, "Error Field")
 dbAddress.SetFocus
ElseIf Len(dbZipCode.Text) = 0 Then
 Msg = MsgBox("The field [Zip Code] can't be empty", vbExclamation, "Error Field")
 dbZipCode.SetFocus
ElseIf Len(dbCity.Text) = 0 Then
 Msg = MsgBox("The field [City] can't be empty", vbExclamation, "Error Field")
 dbCity.SetFocus
Exit Function
End If

ErrorCheckFields:
End Function


'HERE START THE FUNCTIONS OF THIS APP. LIKE CLEARING AND ADDING NEW DATA TO THE DATABASE
Function dbFields()
On Error GoTo ErrordbFields

 If RecNum = 0 Then
  Exit Function
 End If
    With RS
     .Edit
      dbFName.Text = !fname
      dbLName.Text = !lname
      dbMName.Text = !mname
      dbBirthDay.Text = !birthday
      dbAddress.Text = !address
      dbZipCode.Text = !zipcode
      dbState.Text = !State
      dbCity.Text = !city
      dbCountry.Text = !country
      dbHPhone.Text = !hphone
      dbWPhone.Text = !wphone
      dbCPhone.Text = !cphone
      dbEMail.Text = !email
      If !sex = "Male" Then dbSex(1).Value = True
      If !sex = "Female" Then dbSex(2).Value = True
      If !sex = "Unknown" Then dbSex(3).Value = True
      If !fam = "No" Then dbFam(1).Value = True
      If !fam = "Yes" Then dbFam(2).Value = True
        
    End With
Exit Function

ErrordbFields:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error dbFields!")
End Function

Function dbClsFields()
On Error GoTo ErrordbClsFields

    With RS
     .Edit
      dbFName.Text = ""
      dbLName.Text = ""
      dbMName.Text = ""
      dbSex(3).Value = True
      dbBirthDay.Text = ""
      dbFam(1).Value = True
      dbAddress.Text = ""
      dbZipCode.Text = ""
      dbState.Text = ""
      dbCity.Text = ""
      dbCountry.Text = ""
      dbHPhone.Text = ""
      dbWPhone.Text = ""
      dbCPhone.Text = ""
      dbEMail.Text = ""
    End With
Exit Function

ErrordbClsFields:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error dbClsFields!")
End Function

Function dbUpFields()
On Error GoTo ErrordbUpFields

CheckFields

 If RecNum = 0 Then
  Exit Function
 End If

    With RS
     .Edit
      !fname = dbFName.Text
      !lname = dbLName.Text
      !mname = dbMName.Text
      !sex = dbSexField
      !birthday = dbBirthDay.Text
      !fam = dbFamField
      !address = dbAddress.Text
      !zipcode = dbZipCode.Text
      !State = dbState.Text
      !city = dbCity.Text
      !country = dbCountry.Text
      !hphone = dbHPhone.Text
      !wphone = dbWPhone.Text
      !cphone = dbCPhone.Text
      !email = dbEMail.Text
     .Update
    End With
    NameList
Exit Function

ErrordbUpFields:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error dbUpFields!")
End Function

Function Reload()
On Error GoTo ErrorReload

 dbFields
 NameList
 Aantal.Caption = "Record " & CurRec & " of " & RS.RecordCount
Exit Function

ErrorReload:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Reload!!")
End Function
'HERE IS THE FUNCTION
Function NameList()
On Error GoTo Error_NameList
 
 RS.MoveFirst
 Index = 0
 dbNameList.Clear
 For I = 1 To RecNum
  dbNameList.AddItem RS!fname
  dbNameList.ItemData(Index) = RS!id
  Index = Index + 1
  RS.MoveNext
 Next I
Exit Function

Error_NameList:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error NameList!!")
End Function

Private Sub Timer1_Timer()
On Error GoTo ErrorTimer

 Unload Me
Exit Sub

ErrorTimer:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Timer!!")
End Sub

Private Sub dbNameList_Click()
On Error GoTo ErrordbNameList

RS.MoveFirst
For LoopIndex = 0 To dbNameList.ListCount - 1
  If dbNameList.Selected(LoopIndex) Then
    Do Until RS.EOF
     If RS!id Like dbNameList.ItemData(dbNameList.ListIndex) Then
      Aantal.Caption = "Record " & LoopIndex + 1 & " of " & RS.RecordCount
      CurRec = LoopIndex + 1
      dbFields
      Exit Sub
     Else
      RS.MoveNext
     End If
     Loop
  End If
 Next LoopIndex
Exit Sub

ErrordbNameList:
 Msg = MsgBox(Err.Description & vbCrLf, vbCritical, "Error NameListClick")
End Sub

Private Sub AddBtn_Click()
On Error GoTo Error_AddBtn

CheckFields
 With RS
  .AddNew
     !fname = dbFName.Text
     !lname = dbLName.Text
     !mname = dbMName.Text
     !sex = dbSexField
     !birthday = dbBirthDay.Text
     !fam = dbFamField
     !address = dbAddress.Text
     !zipcode = dbZipCode.Text
     !State = dbState.Text
     !city = dbCity.Text
     !country = dbCountry.Text
     !hphone = dbHPhone.Text
     !wphone = dbWPhone.Text
     !cphone = dbCPhone.Text
     !email = dbEMail.Text
  .Update
 End With
 RecNum = RS.RecordCount
 Aantal.Caption = RecNum & " Records"
 If (RecNum = 1) Or (NewBtn.Caption = "Clear") Then
  PrevBtn.Enabled = True
  NextBtn.Enabled = True
  NewBtn.Enabled = True
  UpdateBtn.Enabled = True
  DeleteBtn.Enabled = True
  UpdateBtn.Caption = "Update"
  NewBtn.Caption = "New"
 End If
Call Form_Load
RS.MoveLast
Exit Sub

Error_AddBtn:
If CurRec = 0 Then
 PrevBtn.Enabled = False
 NextBtn.Enabled = False
 NewBtn.Enabled = False
 UpdateBtn.Enabled = False
 DeleteBtn.Enabled = False
End If
If Err.Number = 3315 Then
 Exit Sub
End If
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error AddBtn!!")
End Sub

Private Sub NewBtn_Click()
On Error GoTo Error_NewBtn
 
 RS.Edit
 dbClsFields
 PrevBtn.Enabled = False
 NextBtn.Enabled = False
 DeleteBtn.Enabled = False
 NewBtn.Enabled = True
 UpdateBtn.Caption = "Cancel"
 NewBtn.Caption = "Clear"
 dbFName.SetFocus
 Aantal.Caption = "Record " & RS.RecordCount + 1 & " of " & RS.RecordCount + 1
Exit Sub

Error_NewBtn:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error NewBtn!!")
End Sub

Private Sub NextBtn_Click()
On Error GoTo Error_NextBtn

RS.MoveNext
 If RS.EOF Then
  Msg = MsgBox("There are no more Records", vbInformation, "Last Record!")
  RS.MoveLast
 Else
  RS.Edit
  dbFields
  CurRec = CurRec + 1
  Aantal.Caption = "Record " & CurRec & " of " & RS.RecordCount
 End If
Exit Sub

Error_NextBtn:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error NextBtn!!")
End Sub

Private Sub PrevBtn_Click()
On Error GoTo Error_PrevBtn

RS.MovePrevious
 If RS.BOF Then
  Msg = MsgBox("There are no more Records", vbInformation, "Info!")
  RS.MoveFirst
 Else
  RS.Edit
   dbFields
  CurRec = CurRec - 1
  Aantal.Caption = "Record " & CurRec & " of " & RS.RecordCount
 End If
Exit Sub

Error_PrevBtn:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error PrevBtn!!")
End Sub

Private Sub DeleteBtn_Click()
On Error GoTo Error_DeleteBtn

Msg = MsgBox("You ar about to Delete record: " & CurRec & vbCrLf & "Ar you shure to delete?", vbQuestion + vbYesNo, "Delete Record!")
Select Case Msg
 Case vbYes
  GoTo Delete
 Case vbNo
  Exit Sub
End Select

Delete:
 RS.Delete
 Call Form_Load
Exit Sub

Error_DeleteBtn:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Deleting!!")
End Sub

Private Sub UpdateBtn_Click()
On Error GoTo ErrorUpdateBtn

If UpdateBtn.Caption = "Update" Then
 dbUpFields
 Call Form_Load
ElseIf UpdateBtn.Caption = "Cancel" Then
 CurRec = RS.RecordCount
 Aantal.Caption = "Record " & CurRec & " of " & RS.RecordCount
 UpdateBtn.Caption = "Update"
 NewBtn.Caption = "New"
 PrevBtn.Enabled = True
 NextBtn.Enabled = True
 DeleteBtn.Enabled = True
Call Form_Load
End If
Exit Sub

ErrorUpdateBtn:
 Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error UpdateBtn!")
End Sub

Private Sub dbSex_Click(Index As Integer)
On Error GoTo ErrordbSexClick

If dbSex(1).Value = True Then dbSexField = "Male"
If dbSex(2).Value = True Then dbSexField = "Female"
If dbSex(3).Value = True Then dbSexField = "Unknown"
Exit Sub

ErrordbSexClick:
Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Gender click")
End Sub

Private Sub dbFam_Click(Index As Integer)
On Error GoTo ErrordbSexClick

If dbFam(1).Value = True Then dbFamField = "No"
If dbFam(2).Value = True Then dbFamField = "Yes"
Exit Sub

ErrordbSexClick:
Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error Family click")
End Sub

