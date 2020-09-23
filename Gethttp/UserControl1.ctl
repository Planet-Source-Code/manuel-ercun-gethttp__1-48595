VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox Picture1 
      Height          =   225
      Left            =   360
      ScaleHeight     =   165
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   1320
      Width           =   3810
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim imax, imin, ivalue As Long
Private Sub UserControl_Initialize()
Picture1.Move 0, 0
imin = 0
imax = 100
ivalue = 1
End Sub

Public Property Get Max() As Long
Max = imax
End Property
Public Property Let Max(ByVal new_max As Long)
imax = new_max
If imax < ivalue Then imax = ivalue
If imax < imin Then imax = imin
PropertyChanged "Max"
End Property

Public Property Get Min() As Long
Min = imin
End Property

Public Property Let Min(ByVal new_min As Long)
imin = new_min
If imin > imax Then imin = imax
If imin > ivalue Then imin = ivalue
PropertyChanged "Min"
End Property

Public Property Get Value() As Long
Value = ivalue
End Property
Public Property Let Value(ByVal new_value As Long)
ivalue = new_value
If ivalue > imax Then ivalue = imax
If ivalue < imin Then ivalue = imin
Label1.Width = Int(ivalue - imin) / Int(imax - imin) * Picture1.ScaleWidth
PropertyChanged "Value"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Max = PropBag.ReadProperty("Max", 100)
Min = PropBag.ReadProperty("Min", 0)
Value = PropBag.ReadProperty("Value", 1)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Picture1.Height
Picture1.Width = UserControl.Width
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Max", imax, 100)
Call PropBag.WriteProperty("Min", imin, 0)
Call PropBag.WriteProperty("Value", ivalue, 1)
End Sub



