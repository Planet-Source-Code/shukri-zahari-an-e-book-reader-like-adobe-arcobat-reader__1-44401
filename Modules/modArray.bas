Attribute VB_Name = "modArray"
'©Backwoods Interactive 2002. All rights reserved.
'Programmed by James J. Kelly Jr.
'
'www.Backwoods-Interactive.com

Public Function DeleteValue(ByVal Index As Long, ByRef UserArray As Variant)

On Error Resume Next

Dim ValBuffer() As Variant

ReDim ValBuffer(UBound(UserArray)) As Variant

'collect array information. This simply jumps ahead and gets all the
'values ahead of the one to be deleted

For i = Index To UBound(ValBuffer)
On Error Resume Next
ValBuffer(i) = UserArray(i + 1)
Next i

'This here copy's the values ahead of the deleted values from the array buffer into
'the main array.
For h = Index To UBound(UserArray)
On Error Resume Next
UserArray(h) = ValBuffer(h)
Next h

'This will finish the job
ReDim Preserve UserArray(UBound(UserArray) - 1) As Variant

End Function

Public Function InsertValue(ByVal Index As Long, ByVal Value As Variant, ByRef UserArray As Variant)

Dim ValBuffer() As Variant
Dim CurrentVal As Variant

CurrentVal = Index + 1

ReDim ValBuffer(UBound(UserArray)) As Variant

'Collect array information and copy it into the Array Buffer
For i = CurrentVal To UBound(ValBuffer)
On Error Resume Next
ValBuffer(i) = UserArray(i)
Next i

'Resize array but keep the values in it
ReDim Preserve UserArray(UBound(UserArray) + 1) As Variant
UserArray(CurrentVal) = Value

For h = CurrentVal To UBound(UserArray)
On Error Resume Next
UserArray(h + 1) = ValBuffer(h)
Next h

End Function

Public Function AddValue(ByVal Value As Variant, ByRef UserArray As Variant)

On Error Resume Next
ReDim Preserve UserArray(UBound(UserArray) + 1) As Variant
UserArray(UBound(UserArray)) = Value

End Function

Public Function ModifyValue(ByVal Index As Long, ByVal Value As Variant, ByRef UserArray As Variant)

UserArray(Index) = Value

End Function

Public Function SaveArray(ByVal Filename As String, ByRef UserArray As Variant)

Dim Free As Long

Free = FreeFile

Open Filename For Binary As #Free
Put #Free, , UserArray
Close #Free

End Function

Public Function LoadArray(ByVal Filename As String, ByRef UserArray As Variant)

Dim Free As Long

Free = FreeFile

Open Filename For Binary As #Free
Get #Free, , UserArray
Close #Free

End Function

Public Function CleanResize(ByVal Size As Long, ByRef UserArray As Variant)
ReDim UserArray(Size) As Variant
End Function

Public Function Resize(ByVal Size As Long, ByRef UserArray As Variant)
ReDim Preserve UserArray(Size) As Variant
End Function

Public Function MoveValue(ByVal Index As Long, ByVal Index2 As Long, ByRef UserArray As Variant)

Dim ValBuffer() As Variant
Dim CurrentVal As Variant

CurrentVal = Index2

ReDim ValBuffer(UBound(UserArray)) As Variant

'Collect array information and copy it into the Array Buffer
For i = CurrentVal To UBound(ValBuffer)
On Error Resume Next
ValBuffer(i) = UserArray(i)
Next i

'Resize array but keep the values in it
ReDim Preserve UserArray(UBound(UserArray) + 1) As Variant
UserArray(CurrentVal) = UserArray(Index)

For h = CurrentVal To UBound(UserArray)
On Error Resume Next
UserArray(h + 1) = ValBuffer(h)
Next h

Call DeleteValue(Index, UserArray)

End Function

