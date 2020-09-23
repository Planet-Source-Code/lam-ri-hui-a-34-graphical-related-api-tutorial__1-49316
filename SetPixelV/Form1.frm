VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SetPixelV API Demo"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Sub Form_Load()

    Dim mRGN As Long, R As RECT, x As Long, y As Long
    'Set the graphical mode to persistent
    Me.AutoRedraw = True
    'Set the rectangle's values
    SetRect R, 100, 100, 350, 350
    'Create an elliptical region
    mRGN = CreateEllipticRgnIndirect(R)
    For x = R.Left To R.Right
        For y = R.Top To R.Bottom
            'If the point is in the region, draw a green pixel
            If PtInRegion(mRGN, x, y) <> 0 Then
                'Draw a green pixel
                SetPixelV Me.hdc, x, y, vbGreen
            ElseIf PtInRect(R, x, y) <> 0 Then
                'Draw a red pixel
                SetPixelV Me.hdc, x, y, vbRed
            End If
        Next y
    Next x
    'delete our region
    DeleteObject mRGN
End Sub

