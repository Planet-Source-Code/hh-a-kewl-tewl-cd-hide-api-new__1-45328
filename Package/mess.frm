VERSION 5.00
Begin VB.Form mess 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HH Messenger"
   ClientHeight    =   735
   ClientLeft      =   3285
   ClientTop       =   4440
   ClientWidth     =   5985
   Icon            =   "mess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   5985
   Begin VB.CommandButton Command8 
      Caption         =   "Out To Lunch"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "On The Phone"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Be Right Back"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Online"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Away"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Invisible"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Text            =   "-enter nick here-"
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Busy"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "mess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public msn As MsgrObject

Private Sub Command1_Click()
msn.LocalState = MSTATE_BUSY
End Sub

Private Sub Command2_Click()
If Text1.Text = vbNullString Then
MsgBox ("Sorry, You must have a name."), vbOKOnly, "Invalid Name"

Exit Sub
Else
msn.Services.PrimaryService.FriendlyName = Text1.Text
End If
End Sub

Private Sub Command3_Click()
msn.LocalState = MSTATE_INVISIBLE
End Sub

Private Sub Command4_Click()
msn.LocalState = MSTATE_AWAY
End Sub

Private Sub Command5_Click()
msn.LocalState = MSTATE_ONLINE
End Sub

Private Sub Command6_Click()
msn.LocalState = MSTATE_BE_RIGHT_BACK
End Sub

Private Sub Command7_Click()
msn.LocalState = MSTATE_ON_THE_PHONE
End Sub

Private Sub Command8_Click()
msn.LocalState = MSTATE_OUT_TO_LUNCH
End Sub

Private Sub Command9_Click()



End Sub

Private Sub Form_Load()
Set msn = New MsgrObject
Dim X As String
X = msn.LocalLogonName
mess.Caption = "User: " & X
End Sub
