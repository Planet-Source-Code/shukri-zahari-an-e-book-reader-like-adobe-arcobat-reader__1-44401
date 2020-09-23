VERSION 5.00
Begin VB.Form Skip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Skip to Page..."
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3060
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
   ScaleHeight     =   1485
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtPage 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1170
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Skip to page:"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "Skip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
If KeyAscii = 13 Then
If CInt(txtPage.Text - 1) > Main.lstPage.ListCount - 1 Then: MsgBox "Invalid page", vbCritical, "Invalid Page": txtPage.SelStart = 0: txtPage.SelLength = Len(txtPage.Text): Exit Sub
Main.lstPage.ListIndex = CInt(txtPage.Text - 1)
OpenPage Main.lstPage
Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Icon = Nothing
txtPage.SelStart = 0
txtPage.SelLength = Len(txtPage.Text)
End Sub
