VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GetHttp"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3960
      Top             =   3120
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":030A
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   28
      Top             =   7080
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Save"
      Height          =   375
      Left            =   7920
      TabIndex        =   27
      Top             =   6000
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   9000
      Top             =   6480
   End
   Begin GetHttp.UserControl1 UserControl11 
      Height          =   225
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   397
      Value           =   0
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   7920
      Picture         =   "Form1.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   7920
      Picture         =   "Form1.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scan Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   3495
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "3000"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   18
         Text            =   "80"
         Top             =   360
         Width           =   615
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5040
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0C28
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1504
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1DE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":26BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2F98
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3874
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3B90
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":446C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":47BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4B10
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4E62
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":51B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5506
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5858
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(ms)"
         Height          =   195
         Left            =   2880
         TabIndex        =   24
         Top             =   390
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout"
         Height          =   195
         Left            =   1440
         TabIndex        =   22
         Top             =   390
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   390
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IPs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Command6 
         Caption         =   "Browser..."
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "->"
         Height          =   735
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear "
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear Selected"
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "->"
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   290
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1270
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   890
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   290
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Read IPs From file"
         Height          =   195
         Left            =   1200
         TabIndex        =   15
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End IP"
         Height          =   195
         Left            =   550
         TabIndex        =   7
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start IP"
         Height          =   195
         Left            =   525
         TabIndex        =   5
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hostname/IP"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   945
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3360
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4260
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GETHTTP"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   5040
      TabIndex        =   29
      Top             =   2450
      Width           =   1995
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/0"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scanned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   25
      Top             =   5940
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   2040
      TabIndex        =   2
      Top             =   4320
      Width           =   15
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuip 
         Caption         =   "Copy IP clipboard"
      End
      Begin VB.Menu mnuweb 
         Caption         =   "Webbrowser"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnushell 
      Caption         =   "options"
      Visible         =   0   'False
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuUni 
         Caption         =   "Unistall"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim u, h As Integer
Dim change As Boolean
Dim r As Object
Dim Shell_lnk As New Class1


Private Sub Command1_Click()
If IsNumeric(Text1) = False Then
change = True
Winsock1(0).Close
Beep
Winsock1(0).Connect Trim(Text1), Trim(Text4)
Else
Text2 = Text1
Text3 = Text2
ListView2.ListItems.Add 1, , Trim(Text1)
End If
End Sub

Private Sub Command2_Click()
ListView2.ListItems.Remove ListView2.SelectedItem.Index
End Sub

Private Sub Command3_Click()
ListView2.ListItems.Clear
End Sub

Private Sub Command4_Click()
If ListView2.ListItems.Count = 0 Then Exit Sub
If ListView2.ListItems.Item(1) <> 0 And ListView2.ListItems.Item(1).ListSubItems.Count = 0 Then change = False: Connection0 Trim(ListView2.ListItems.Item(1).Text), Trim(Text4)
If ListView2.ListItems.Item(1) <> 0 And ListView2.ListItems.Item(1).ListSubItems.Count <> 0 Then
Command4.Enabled = False
Command7.Enabled = True
Timer1.Interval = Val(Text5)
salir = False
Calc


IPs Trim(ListView2.ListItems.Item(1).Text), Trim(ListView2.ListItems.Item(1).ListSubItems.Item(1).Text), Trim(Text4)

End If
End Sub

Private Sub Command5_Click()
If InStr(Text2, "255") <> 0 Then Text2 = Replace(Text2, "255", "254")
If InStr(Text3, "255") <> 0 Then Text3 = Replace(Text3, "255", "254")
ListView2.ListItems.Add 1, , Trim(Text2)
ListView2.ListItems.Item(1).SubItems(1) = Trim(Text3)
End Sub

Private Sub Command6_Click()
Dim g As String
On Error GoTo ema
With CommonDialog1
  .CancelError = True
  .Filter = "txt(*.txt)|*.txt"
  .DialogTitle = "Open file...."
  .ShowOpen
  If Len(.FileName) = 0 Then Exit Sub
  Command4.Enabled = False
  Command7.Enabled = False
  Timer1.Interval = Val(Text5)
  Open .FileName For Input As #1
  Do
  Line Input #1, g
  ListView2.ListItems.Add 1, , g
  Connection0 Trim(g), 80
  Timer1.Enabled = True
  cas = False
  Do
  DoEvents
  Loop Until cas = True
  Timer1.Enabled = False
  Loop Until EOF(1)
  Close #1
  cas = False
  Close #1
  Command4.Enabled = True
  
End With
ema:
End Sub

Private Sub Command7_Click()
salir = True

Command7.Enabled = False
End Sub

Private Sub Command8_Click()

With CommonDialog1
 .CancelError = True
 .Filter = "All files(*.*)|*.*"
 .DefaultExt = "TXT"
 .DialogTitle = "Save File"
 .ShowSave
 If Len(.FileName) = 0 Then Exit Sub
Open .FileName For Output As #1
Print #1, ListView1.ColumnHeaders.Item(1).Text & "     " & ListView1.ColumnHeaders.Item(2).Text & "     " & ListView1.ColumnHeaders.Item(3).Text & "     " & ListView1.ColumnHeaders.Item(4).Text & _
 "     " & ListView1.ColumnHeaders.Item(5).Text & vbCrLf & vbCrLf

For i = 1 To ListView1.ListItems.Count
Print #1, ListView1.ListItems.Item(i).Text & "     " & ListView1.ListItems.Item(i).ListSubItems.Item(1) & "     " & ListView1.ListItems.Item(i).ListSubItems.Item(2) & "     " & ListView1.ListItems.Item(i).ListSubItems.Item(3) & _
"     " & ListView1.ListItems.Item(i).ListSubItems.Item(4) & "     "


Next i
 
 Close #1

End With

End Sub

Private Sub Form_Activate()
For z = 1 To 100
 Load Winsock1(z)
Next z
End Sub

Private Sub Form_Load()
Me.Move (Screen.Width / 2) - (Me.Width / 2), (Screen.Height / 2) - (Me.Height / 2)
Set ListView1.SmallIcons = ImageList1
ListView2.View = lvwReport
ListView1.View = lvwReport
ListView2.ColumnHeaders.Add 1, , "Start IP", 1500
ListView2.ColumnHeaders.Add 2, , "End IP", 1500
ListView1.ColumnHeaders.Add 1, , "IPs", 1800
ListView1.ColumnHeaders.Add 2, , "Port", 700
ListView1.ColumnHeaders.Add 3, , "Code", 700
ListView1.ColumnHeaders.Add 4, , "Date", 2200
ListView1.ColumnHeaders.Add 5, , "Server", 3400
Text1 = Winsock1(0).LocalIP
Command7.Enabled = False
Shell_lnk.Lnk
End Sub






Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Shell_lnk.ShellAdd
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnufile
End Sub

Private Sub mnuabout_Click()
Form2.Show vbModal
End Sub

Private Sub mnuclear_Click()
ListView1.ListItems.Clear
End Sub

Private Sub mnudelete_Click()
ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuip_Click()
Clipboard.Clear
Clipboard.SetText ListView1.SelectedItem
End Sub

Private Sub mnuUni_Click()
Shell_lnk.Dulnk
End Sub

Private Sub mnuweb_Click()
Shell ("explorer http://" & ListView1.SelectedItem)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
u = u + 1
If u = 2 Then
Shell_lnk.ShellDel
u = 0
End If
ElseIf Button = 2 Then
PopupMenu mnushell
End If
End Sub

Private Sub Timer1_Timer()

For i = 1 To 100
 Winsock1(i).Close
Next i
cas = True
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()

h = (h + 1) Mod 16
Label11.ForeColor = QBColor(h)
End Sub

Private Sub Winsock1_Connect(Index As Integer)
If Winsock1(Index).State = sckConnected Then

If Index = 0 Then
 If change = False Then

  Winsock1(0).SendData "HEAD http://" & Winsock1(0).RemoteHostIP & " HTTP/1.0" & vbCrLf & vbCrLf

 End If
 
 If change = True Then
  ListView2.ListItems.Add 1, , Winsock1(0).RemoteHostIP
  Text2 = Winsock1(0).RemoteHostIP
  Text3 = Text2
  change = False
  Winsock1(0).Close
  Exit Sub
 End If
 

 
End If


If Index > 0 Then

Winsock1(Index).SendData "HEAD http://" & Winsock1(Index).RemoteHostIP & " HTTP/1.0" & vbCrLf & vbCrLf
End If
End If
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim g As String
Winsock1(Index).GetData g

v = Split(g, vbCrLf)
Set r = ListView1.ListItems.Add(, , Winsock1(Index).RemoteHostIP)
If InStr(g, "Server:") = 0 Then r.SmallIcon = 12

r.SubItems(1) = Trim(Text4)
For i = LBound(v) To UBound(v) - 1
If Left(v(i), 4) = "HTTP" Then r.SubItems(2) = Mid(v(i), 9, InStr(v(i), Chr(32)) - 5)
If Left(v(i), 5) = "Date:" Then r.SubItems(3) = Mid(v(i), 6, Len(v(i)))
If Left(v(i), 7) = "Server:" Then r.SmallIcon = CInt(Icons(CStr(Mid(v(i), 9)))): r.SubItems(4) = Mid(v(i), 8, Len(v(i)))
Next i

Winsock1(Index).Close
DoEvents

cas = True
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1(Index).Close
End Sub

Private Sub descargar()
For i = 1 To 100
 Unload Winsock1(i)
Next i
End Sub
