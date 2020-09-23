VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GetBkColor API Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'In general section
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long

Private Sub Timer1_Timer()

    'Set the Form's backcolor
    Me.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    'Get the backcolor
    MsgBox "My backcolor is:" + Str$(GetBkColor(Me.hDC))
End Sub

