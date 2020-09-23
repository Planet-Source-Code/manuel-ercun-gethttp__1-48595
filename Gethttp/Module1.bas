Attribute VB_Name = "Module1"
Option Explicit

Public i As Long
Public z As Long
Public v As Variant
Dim y  As Long
Dim ser As String
Dim res(), res1() As String
Public salir As Boolean
Public cas As Boolean
Dim Result As Long

Private Function Jump(str As String) As String

v = Split(str, ".")
For i = LBound(v) To UBound(v)
ReDim Preserve res(i)
res(i) = v(i)
Next i

End Function


Private Sub Range()

res(3) = res(3) + 1
If res(3) = 255 Then
res(3) = 0
res(2) = res(2) + 1
ElseIf res(2) = 255 Then
res(3) = 0
res(2) = 0
res(1) = res(1) + 1
ElseIf res(1) = 255 Then
res(3) = 0
res(2) = 0
res(1) = 0
res(0) = res(0) + 1
ElseIf res(0) = 255 Then
res(3) = 0
res(1) = 0
res(2) = 0
res(0) = 0
End If

ser = res(0) & "." & res(1) & "." & res(2) & "." & res(3)


End Sub

Public Sub IPs(str As String, str1 As String, port As String)
Dim ip() As String

 ser = str
 Jump (ser)


Do
For i = 1 To 50
y = y + 1
 
 Range
 ReDim Preserve ip(i)
 ip(i) = ser
Form1.UserControl11.Value = Form1.UserControl11.Value + 1
Form1.Label10.Caption = y & "/" & Result
Form1.Winsock1(i).Close
Form1.Winsock1(i).Connect ip(i), port
If CStr(ser) = Trim(str1) Or salir = True Then
Form1.Command4.Enabled = True
salir = False
Exit Sub
End If
Next i
Form1.Timer1.Enabled = True

cas = False
Do
DoEvents
Loop Until cas = True
Form1.Timer1.Enabled = False

Loop





End Sub



Public Sub Calc()

Jump1 Trim(Form1.ListView2.ListItems.Item(1).ListSubItems.Item(1).Text)
Jump Trim(Form1.ListView2.ListItems.Item(1).Text)

Result = ((res1(0) - res(0)) * 255) + ((res1(1) - res(1)) * 255) + ((res1(2) - res(2)) * 255) + (res1(3) - res(3))
If Result < 0 Then MsgBox "ranges is misplaced", vbCritical, "GetHttp": Exit Sub
Form1.UserControl11.Min = 0
Form1.UserControl11.Max = Result

End Sub




Private Sub Jump1(str As String)
v = Split(str, ".")
For i = LBound(v) To UBound(v)
ReDim Preserve res1(i)
res1(i) = v(i)
Next i
End Sub









Public Sub Connection0(str As String, port As Long)
Form1.Winsock1(0).Close
Form1.Winsock1(0).Connect str, port
End Sub


Public Function Icons(ByRef str As Variant) As Integer
Dim icon As Integer
If InStr(str, "Microsoft-IIS/4.0") <> 0 Or InStr(str, "Microsoft-IIS/3.0") <> 0 Then
icon = 14
ElseIf InStr(str, "Microsoft-IIS/5") <> 0 Or InStr(str, "Microsoft-IIS/6") <> 0 Then
icon = 8
ElseIf InStr(str, "(Win32)") <> 0 Then
icon = 14
ElseIf InStr(str, "Apache") <> 0 Then
icon = 9
ElseIf InStr(str, "Netscape") <> 0 Then
icon = 4
ElseIf InStr(str, "Zeus") <> 0 Then
icon = 5
ElseIf InStr(str, "Lotus") <> 0 Then
icon = 2
ElseIf InStr(str, "Allegro-Software-RomPager") <> 0 Then
icon = 1
ElseIf InStr(str, "IBM_HTTP_SERVER") <> 0 Then
icon = 3
ElseIf InStr(str, "Oracle") <> 0 Then
icon = 6
ElseIf InStr(str, "Stronghold") <> 0 Then
icon = 13
ElseIf InStr(str, "Rapidsite") <> 0 Then
icon = 7
ElseIf InStr(str, "Rapidsite") <> 0 Then
icon = 7
Else
icon = 12
End If
Icons = icon
End Function
