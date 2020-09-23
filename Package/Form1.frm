VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HH TooLz"
   ClientHeight    =   1440
   ClientLeft      =   3660
   ClientTop       =   2310
   ClientWidth     =   6495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6495
   Begin VB.CommandButton Command18 
      Caption         =   "Restart pc"
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Shut down pc"
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Empty trash"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Show taskbar"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Hide taskbar"
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Computer name"
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Hide icons"
      Height          =   495
      Left            =   2160
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Show icons"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Msn Messenger"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Show mouse"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Hide mouse"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Swap mouse normal"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Swap mouse buttons"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Scroll on/off"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Caps on/off"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Num on/off"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CD - Close"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CD - Open"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type INITCOMMONCONTROLSEX_TYPE
dwSize As Long
dwICC As Long
End Type
Private Const ICC_INTERNET_CLASSES = &H800
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long


Private Sub Command1_Click()
On Error Resume Next
Dim tReturn As Long
mciSendString "set CDAudio door open", tReturn, 127, 0
'this below didnt seem to work with 95/98
'mciSendString "set cdaudio door open", 0, 0, 0

End Sub



Private Sub Command10_Click()
mess.Show
End Sub

Private Sub Command11_Click()
    Dim hWnd As Long

    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 5
End Sub


Private Sub Command12_Click()
Dim hWnd As Long

    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 0
End Sub
Function GetNameofComputer() As String
    Dim Buff As String
    Dim Sze As Long
    
    Buff = Space$(255)
    Sze = 255
    
    GetComputerName Buff, Sze
    GetNameofComputer = Left$(Buff, Sze)
End Function
Private Sub Command13_Click()
'function is above ^^
MsgBox GetNameofComputer, vbOKOnly, "Computers Name"
End Sub

Private Sub Command14_Click()
Dim tReturn
    
    tReturn = FindWindow("Shell_traywnd", "")
    SetWindowPos tReturn, 0, 0, 0, 0, 0, &H80
End Sub

Private Sub Command15_Click()
Dim tReturn As Long

    tReturn = FindWindow("Shell_traywnd", "")
    SetWindowPos tReturn, 0, 0, 0, 0, 0, &H40
End Sub


Private Sub Command16_Click()
SHEmptyRecycleBin hWnd, "", &H2
End Sub

Private Sub Command17_Click()
ExitWindowsEx 1, 0&
End Sub

Private Sub Command18_Click()
ExitWindowsEx 2, 0&
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim tReturn As Long
mciSendString "set CDAudio door closed", tReturn, 127, 0
'this didnt seem to work with 95/98
'mciSendString "set cdaudio door closed", 0, 0, 0

End Sub



Private Sub Command3_Click()
Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
Numlockstate = keys(VK_NUMLOCK)
If Numlockstate <> True Then
'Simulate Key Press
          keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY _
            Or KEYEVENTF_KEYUP, 0




End If

End Sub



Private Sub Command4_Click()
Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
Numlockstate = keys(VK_NUMLOCK)
If Numlockstate <> True Then
'Simulate Key Press
          keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY _
            Or KEYEVENTF_KEYUP, 0
End If
End Sub

Private Sub Command5_Click()
Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
Numlockstate = keys(VK_NUMLOCK)
If Numlockstate <> True Then
'Simulate Key Press
          keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY _
            Or KEYEVENTF_KEYUP, 0
End If
End Sub

Private Sub Command6_Click()
 SwapMouseButton 1
End Sub

Private Sub Command7_Click()
SwapMouseButton 0
End Sub



Private Sub Command8_Click()
ShowCursor 0
End Sub

Private Sub Command9_Click()
ShowCursor 1
End Sub




