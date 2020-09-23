Attribute VB_Name = "modBook"
Private Type STRUCT_FILE
 Version As String
 Filename() As String
 FileContent() As String
End Type

Dim sFile As STRUCT_FILE

Public Sub LoadReader()

sFile.Version = Hex(CInt(EBOOKREADER1))
ReDim sFile.FileContent(0)
ReDim sFile.Filename(0)
With Main
.tbMenu.Buttons(10).Enabled = .mnuback.Enabled
.tbMenu.Buttons(11).Enabled = .mnuforward.Enabled
End With

End Sub

Private Function ListFile()

With Main
.lstPage.Clear

For i = 1 To UBound(sFile.Filename)
.lstPage.AddItem "-" + sFile.Filename(i)
Next i
.lstPage.ListIndex = 0
OpenPage .lstPage
End With

End Function

Public Function OpenFile(CD As CommonDialog) As Boolean

Dim Free As Long
Free = FreeFile
On Error GoTo OpenErr:
CD.DialogTitle = "Open e-Book"
CD.Filter = "e-Book Files (*.book)|*.book|All Files (*.*)|*.*"
CD.ShowOpen

If CD.Filename <> "" Then
    Screen.MousePointer = vbHourglass
        Open CD.Filename For Binary As #Free
        Get #Free, , sFile
            If sFile.Version = Hex(CInt(EBOOKREADER1)) Then
            Else
                MsgBox "Unrecognized format of e-Book. This e-book may be compiled" & vbCrLf & "using older version of e-Maker", vbCritical, "Error"
                ReDim sFile.FileContent(0)
                ReDim sFile.Filename(0)
                sFile.Version = ""
                Main.mnuback.Enabled = False
                Main.mnuforward.Enabled = False
                Main.mnuclose.Enabled = False
                Main.mnuskip.Enabled = False
                Main.mnuprint.Enabled = False
                Main.mnuprop.Enabled = False
                Main.lstPage.Clear
                Main.tbMenu.Buttons(10).Enabled = Main.mnuback.Enabled
                Main.tbMenu.Buttons(11).Enabled = Main.mnuforward.Enabled
                Main.tbMenu.Buttons(2).Enabled = Main.mnuclose.Enabled
                Main.tbMenu.Buttons(12).Enabled = Main.mnuskip.Enabled
                Main.Caption = "Win32 e-Book Reader v1.0"
                Main.tbMenu.Buttons(4).Enabled = Main.mnuprint.Enabled
                Main.tbMenu.Buttons(5).Enabled = Main.mnuprop.Enabled
                Close #1
                Screen.MousePointer = vbDefault
                Exit Function
            End If
            Main.sbStat.Panels(1).Text = "e-Book Path: " & Left(CD.Filename, Len(CD.Filename) - Len(CD.FileTitle))
            Main.Caption = Left(CD.FileTitle, Len(CD.FileTitle) - 5) & " - Win32 e-Book Reader v1.0"
            Main.mnuclose.Enabled = True
            Main.mnuskip.Enabled = True
            Main.mnuprint.Enabled = True
            Main.mnuprop.Enabled = True
            Main.tbMenu.Buttons(2).Enabled = True
            Main.tbMenu.Buttons(12).Enabled = True
            Main.tbMenu.Buttons(4).Enabled = Main.mnuprint.Enabled
            Main.tbMenu.Buttons(5).Enabled = Main.mnuprop.Enabled
        ListFile
    Screen.MousePointer = vbDefault
    Main.tmrBackForward.Enabled = True
End If

OpenErr:
If Err.Number <> 0 Then: MsgBox "An error #" & Err.Number & " occured." & vbCrLf & Err.Description, vbCritical, "Error": Screen.MousePointer = vbDefault: Exit Function

End Function

Public Function OpenPage(LST As ListBox)

With Main
.txtPage.Text = sFile.FileContent(LST.ListIndex + 1)
End With

End Function
