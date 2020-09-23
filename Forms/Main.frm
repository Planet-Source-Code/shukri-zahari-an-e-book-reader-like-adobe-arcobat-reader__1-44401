VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Caption         =   "Win32 e-Book Reader v1.0"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8925
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
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   8925
   Begin ComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imMenu"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   " Open an e-Book "
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "close"
            Object.ToolTipText     =   " Close current e-Book "
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "print"
            Object.ToolTipText     =   " Print current page "
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "prop"
            Object.ToolTipText     =   " View e-Book Properties "
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "copy"
            Object.ToolTipText     =   " Copy Selection to Clipboard "
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "all"
            Object.ToolTipText     =   " Select all text in current page "
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "back"
            Object.ToolTipText     =   " Back one page "
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "forward"
            Object.ToolTipText     =   " Forward one page "
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "skip"
            Object.ToolTipText     =   " Skip to page... "
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "config"
            Object.ToolTipText     =   " Setting up e-Book Reader "
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "help"
            Object.ToolTipText     =   " View e-Book Reader Help "
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPage 
      Height          =   5475
      Left            =   30
      Locked          =   -1  'True
      MouseIcon       =   "Main.frx":2CFA
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   420
      Width           =   8805
   End
   Begin VB.Timer tmrBackForward 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3690
      Top             =   5220
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5070
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstPage 
      Height          =   1425
      ItemData        =   "Main.frx":39C4
      Left            =   5430
      List            =   "Main.frx":39C6
      TabIndex        =   3
      Top             =   3300
      Visible         =   0   'False
      Width           =   2115
   End
   Begin ComctlLib.StatusBar sbStat 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7185
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "e-Book Path:"
            TextSave        =   "e-Book Path:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Page 0 of 0"
            TextSave        =   "Page 0 of 0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imMenu 
      Left            =   330
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":39C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":3D1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":406C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":43BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":4710
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":4A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":4DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":5106
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":5458
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":57AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":5AFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open e-Book"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close e-Book"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprop 
         Caption         =   "e-Book P&roperties"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuprint 
         Caption         =   "&Print current page"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy Selected"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuall 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnutbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuviewpage 
         Caption         =   "&Page"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnusbar 
         Caption         =   "&StatusBar"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuconfig 
         Caption         =   "&Config"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnupage 
      Caption         =   "&Page"
      Begin VB.Menu mnuback 
         Caption         =   "&Back one Page"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuforward 
         Caption         =   "&Forward one Page"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuskip 
         Caption         =   "&Skip to Page..."
         Enabled         =   0   'False
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuebrhelp 
         Caption         =   "e-Book Reader &Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEULA 
         Caption         =   "&License Agreement"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About e-Book Reader"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
MsgBox lstPage.ListCount
MsgBox lstPage.ListIndex
End Sub

Private Sub Form_Load()

SetPos
LoadReader

End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMinimized Then: Exit Sub
Call ResizeReader

End Sub

Private Sub mnuabout_Click()
About.Show vbModal, Me
End Sub

Private Sub mnuall_Click()
txtPage.SelStart = 0
txtPage.SelLength = Len(txtPage.Text)
End Sub

Private Sub mnuback_Click()

If lstPage.ListIndex < 0 Then: lstPage.ListIndex = 0: OpenPage lstPage: Exit Sub
lstPage.ListIndex = lstPage.ListIndex - 1
OpenPage lstPage

End Sub

Private Sub mnuclose_Click()
tmrBackForward.Enabled = False
Me.Caption = "Win32 e-Book Reader v1.0"
mnuopen.Enabled = True
mnuclose.Enabled = False
mnuskip.Enabled = False
mnuprint.Enabled = False
mnuprop.Enabled = False
tbMenu.Buttons(2).Enabled = False
lstPage.Clear
txtPage.Text = ""
mnuback.Enabled = False
mnuforward.Enabled = False
tbMenu.Buttons(10).Enabled = mnuback.Enabled
tbMenu.Buttons(11).Enabled = mnuforward.Enabled
sbStat.Panels(1).Text = "e-Book Path: "
sbStat.Panels(2).Text = "Page 0 of 0"
End Sub

Private Sub mnuconfig_Click()
Config.Show vbModal, Me
End Sub

Private Sub mnucopy_Click()
If txtPage.SelText = "" Then: Exit Sub
Clipboard.SetText txtPage.SelText
End Sub

Private Sub mnuebrhelp_Click()

Call ShowHelp

End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnufind_Click()
MsgBox "This function is not yet available.", vbCritical, "Function Unavailable"
End Sub

Private Sub mnuforward_Click()

If lstPage.ListIndex >= lstPage.ListCount Then: lstPage.ListIndex = lstPage.ListCount - 1: Exit Sub
lstPage.ListIndex = lstPage.ListIndex + 1
OpenPage lstPage

End Sub

Private Sub mnuopen_Click()

Call OpenFile(CD)

End Sub

Private Sub mnuprint_Click()
On Error GoTo PrintErr:
txtPage.Print Printer.hDC
PrintErr:
If Err.Number <> 0 Then: MsgBox "A printer error occured (" & Err.Number & "):" & vbCrLf & Err.Description, vbCritical, "Printer Error"
End Sub

Private Sub mnuprop_Click()
If CD.Filename = "" Then: Exit Sub
Dim FSO As FileSystemObject
Set FSO = New FileSystemObject
Dim BookInfo As String
BookInfo = BookInfo & "Book's Title: " & FSO.GetBaseName(CD.Filename) & vbCrLf
BookInfo = BookInfo & "Book's Path: " & FSO.GetParentFolderName(CD.Filename) & vbCrLf
BookInfo = BookInfo & "Book Size: " & Replace(Format(FileLen(CD.Filename), "##.#0KB"), ",", ".") & vbCrLf
BookInfo = BookInfo & "Book's Attribute: " & GetAttr(CD.Filename)
MsgBox BookInfo, vbInformation, "Book's Info"
End Sub

Private Sub mnuskip_Click()
Skip.Show vbModal, Me
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key
Case "open"
mnuopen_Click
Case "close"
mnuclose_Click
Case "print"
mnuprint_Click
Case "prop"
mnuprop_Click
Case "copy"
mnucopy_Click
Case "all"
mnuall_Click
Case "back"
mnuback_Click
Case "forward"
mnuforward_Click
Case "skip"
mnuskip_Click
Case "config"
mnuconfig_Click
Case "help"
mnuebrhelp_Click
End Select

End Sub

Private Sub tmrBackForward_Timer()
If lstPage.ListCount = 0 Then
mnuback.Enabled = False
mnuforward.Enabled = False
tbMenu.Buttons(10).Enabled = mnuback.Enabled
tbMenu.Buttons(11).Enabled = mnuforward.Enabled
End If
If lstPage.ListIndex = 0 Then
mnuback.Enabled = False
mnuforward.Enabled = True
tbMenu.Buttons(10).Enabled = mnuback.Enabled
tbMenu.Buttons(11).Enabled = mnuforward.Enabled
End If
If lstPage.ListIndex = lstPage.ListCount - 1 Then
mnuback.Enabled = True
mnuforward.Enabled = False
tbMenu.Buttons(10).Enabled = mnuback.Enabled
tbMenu.Buttons(11).Enabled = mnuforward.Enabled
End If
If lstPage.ListIndex <> 0 And lstPage.ListIndex <> lstPage.ListCount - 1 Then
mnuback.Enabled = True
mnuforward.Enabled = True
tbMenu.Buttons(10).Enabled = mnuback.Enabled
tbMenu.Buttons(11).Enabled = mnuforward.Enabled
End If
sbStat.Panels(2).Text = "Page " & lstPage.ListIndex + 1 & " of " & lstPage.ListCount
End Sub
