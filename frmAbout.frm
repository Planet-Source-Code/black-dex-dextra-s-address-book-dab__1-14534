VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Dextra's Address Book"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close About"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4440
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmAbout.frx":058A
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "dab@dexunderground.com"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "www.dexunderground.com"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   240
      Picture         =   "frmAbout.frx":06DB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2820
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub
