VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAB Options"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3180
   Begin VB.Frame Frame1 
      Caption         =   "File Association"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.CommandButton RemFileAss 
         Caption         =   "Remove Files Association"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton MakeFileAss 
         Caption         =   "Make File Association"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "If you want the .dab files to open with this app when you dubbel click them press ""Make File Association""."
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub MakeFileAss_Click()
On Error GoTo ErrorMakeFileAss

'CREAT THE FILE ASSOSIAT
RGCreateKey HKEY_CLASSES_ROOT, ".dab"
RGSetKeyValue HKEY_CLASSES_ROOT, ".dab", "", "DABFile"
RGSetKeyValue HKEY_CLASSES_ROOT, ".dab", "Content Type", "DBase/Address"
'SET .DAB FILE INFO
RGCreateKey HKEY_CLASSES_ROOT, "DABFile"
RGSetKeyValue HKEY_CLASSES_ROOT, "DABFile", "", "Dextra's Address Book Files"
'SET .DAB DEFAULTICON
RGCreateKey HKEY_CLASSES_ROOT, "DABFile\DefaultIcon"
RGSetKeyValue HKEY_CLASSES_ROOT, "DABFile\DefaultIcon", "", dbPath & "dabf.ico"
'SET THE RIGHT MOUSE OPEN NAME AND OPEN COMMAND
RGCreateKey HKEY_CLASSES_ROOT, "DABFile\Shell\Open\command"
RGSetKeyValue HKEY_CLASSES_ROOT, "DABFile\Shell\Open\command", "", Chr(34) & dbPath & "dab.exe" & Chr(34) & " /open %1"
'SET THE RIGHT MOUSE PRINT NAME AND PRINT COMMAND
RGCreateKey HKEY_CLASSES_ROOT, "DABFile\Shell\Print Report\command"
RGSetKeyValue HKEY_CLASSES_ROOT, "DABFile\Shell\Print Report\command", "", Chr(34) & dbPath & "dab.exe" & Chr(34) & " /print %1"
'NOW TO TELL THE PERSON THAT IT IS DONNE WE MAKEN AN MSGBOX APPERE
Msg = MsgBox("File Association has been done." & vbCrLf & "You now can dubbel click on a .dab to open the file.", vbInformation, "File Association")
Exit Sub

ErrorMakeFileAss:
Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error MakeFileAss")
End Sub

Private Sub RemFileAss_Click()
On Error GoTo ErrorRemFileAss

'DELETES THE RGCreateKey COMMAND VALUES
RGDeleteKey HKEY_CLASSES_ROOT, ".dab"
RGDeleteKey HKEY_CLASSES_ROOT, "DABFile\DefaultIcon"
RGDeleteKey HKEY_CLASSES_ROOT, "DABFile\Shell\Open\command"
RGDeleteKey HKEY_CLASSES_ROOT, "DABFile\Shell\Print Report\command"
RGDeleteKey HKEY_CLASSES_ROOT, "DABFile"
'NOW TO TELL THE PERSON THAT IT IS DONNE WE LET AN MSGBOX APPERE
Msg = MsgBox("File Association has been Removed." & vbCrLf & "The .dab file is now an Unknown File Type to Windows.", vbInformation, "File Association")

Exit Sub

ErrorRemFileAss:
Msg = MsgBox(Err.Description & vbCrLf & Err.Number, vbCritical, "Error RemFileAss")
End Sub
