VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GetROP2 API Demo"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const R2_BLACK = 1    '  0
Private Const R2_COPYPEN = 13  ' P
Private Const R2_LAST = 16
Private Const R2_MASKNOTPEN = 3 ' DPna
Private Const R2_MASKPEN = 9   ' DPa
Private Const R2_MASKPENNOT = 5 ' PDna
Private Const R2_MERGENOTPEN = 12    ' DPno
Private Const R2_MERGEPEN = 15  ' DPo
Private Const R2_MERGEPENNOT = 14    ' PDno
Private Const R2_NOP = 11    ' D
Private Const R2_NOT = 6 ' Dn
Private Const R2_NOTCOPYPEN = 4 ' PN
Private Const R2_NOTMASKPEN = 8 ' DPan
Private Const R2_NOTMERGEPEN = 2 ' DPon
Private Const R2_WHITE = 16   '  1
Private Const R2_XORPEN = 7   ' DPx
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Sub Form_Load()

    'set the graphics mode to persistent
    Me.AutoRedraw = True
    'check the current mix mode
    If GetROP2(Me.hdc) <> R2_WHITE Then
        'set the current foreground mix mode to R2_WHITE (Pixel is always 1)
        SetROP2 Me.hdc, R2_WHITE
    End If
    'Draw a line from (0,0)-(200,200)
    LineTo Me.hdc, 200, 200
End Sub

