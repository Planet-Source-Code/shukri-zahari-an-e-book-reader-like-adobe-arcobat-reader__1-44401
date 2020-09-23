Attribute VB_Name = "modMain"
'ATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTION
'ATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTION
'
'    After this submission, I will no longer comment my submission coz I don't have time 2 do so
'    If you want to add comments to it, please do so & posted it to me so I can re-upload it..
'    OK, dude?
'
'ATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTION
'ATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTIONATTENTION



Public Sub ResizeReader()

'#####################################################
' # Sub/Function Name : ResizeReader
'# Argument(s)            : N/A
'# What For?               : Resize the reader when user presses the Max button...
'# Copyright                : Me lah...
'#####################################################

With Main
.txtPage.Top = .tbMenu.Top + .tbMenu.Height + 30
.txtPage.Left = 30
.txtPage.Width = .ScaleWidth - 60
.txtPage.Height = .ScaleHeight - (60 + .tbMenu.Height + .sbStat.Height)
End With

End Sub


Public Sub ShowHelp()

'####################################################
'# Sub/Function Name   :   ShowHelp
'# Argument(s)             :   N/A
'# What For?                :   Determine the existance of the Help file & show it to
'#                                    the user....
'# Copyright                 :   Me lah...
'####################################################

Screen.MousePointer = vbHourglass 'Change app's mousepointer to hourglass...
Dim HelpLoc As String
HelpLoc = Dir$(IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "\Help\Help.chm")
If HelpLoc = "" Then: MsgBox "e-Book Documentation is not available. Please reinstall e-Book Documentation.", vbCritical, "Help System"
'JumpTo(IIf(Right(App.Path,1) = "\", App.Path, App.Path & "\") & "\Help\Help.chm")
Screen.MousePointer = vbDefault 'Revert the app's mousepointer back to it default...

End Sub



Public Sub ShowEULA()

'####################################################
'# Sub/Function Name   :   ShowEULA
'# Argument(s)             :   N/A
'# What For?                :   Determine the existance of the EULA file & show it to
'#                                    the user....
'# Copyright                 :   Me lah...
'####################################################
Screen.MousePointer = vbHourglass 'Change app's mousepointer to hourglass...
Dim EULALoc As String
EULALoc = Dir$(IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "\Help\EULA.txt")
If EULALoc = "" Then: MsgBox "e-Book License Agreement is not available. Please reinstall e-Book Documentation.", vbCritical, "Help System"
'JumpTo(IIf(Right(App.Path,1) = "\", App.Path, App.Path & "\") & "\Help\EULA.txt")
Screen.MousePointer = vbDefault 'Revert the app's mousepointer back to it default...

End Sub

Public Function SetPos()

' I think this shouldn't be commented...
With Main
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With

End Function
