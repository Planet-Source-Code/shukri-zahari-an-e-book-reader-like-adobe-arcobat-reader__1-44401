VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Win32 e-Book Reader"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   4050
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   1005
      Left            =   1110
      Picture         =   "About.frx":0000
      Top             =   150
      Width           =   4590
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "This software is released for FREE of charge. But if you wish to use it for your own program, please include me in your credit."
      Height          =   675
      Index           =   4
      Left            =   1110
      TabIndex        =   5
      Top             =   2970
      Width           =   4275
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":F10A
      Height          =   915
      Index           =   3
      Left            =   1110
      TabIndex        =   4
      Top             =   2040
      Width           =   4275
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003, Shukri Zahari"
      Height          =   255
      Index           =   2
      Left            =   1110
      TabIndex        =   3
      Top             =   1680
      Width           =   3105
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0 Build (0001)"
      Height          =   255
      Index           =   1
      Left            =   1110
      TabIndex        =   2
      Top             =   1470
      Width           =   3105
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Win32 e-Book Reader for Windows"
      Height          =   255
      Index           =   0
      Left            =   1110
      TabIndex        =   1
      Top             =   1260
      Width           =   3105
   End
   Begin VB.Image imgLogo 
      Height          =   720
      Left            =   120
      Picture         =   "About.frx":F1DA
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = Nothing
End Sub
