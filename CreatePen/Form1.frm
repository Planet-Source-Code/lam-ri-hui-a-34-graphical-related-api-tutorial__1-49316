VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CreatePen, CreatePenIndirect, FloodFill, FrameRect, FrameRgn and InvertRgn API Demo"
   ClientHeight    =   5970
   ClientLeft      =   1620
   ClientTop       =   2040
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   12765
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PS_DOT = 2
Const PS_SOLID = 0
Const RGN_AND = 1
Const RGN_COPY = 5
Const RGN_OR = 2
Const RGN_XOR = 3
Const RGN_DIFF = 4
Const HS_DIAGCROSS = 5
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Sub Form_Load()
    Me.ScaleMode = vbPixels
End Sub
Private Sub Form_Paint()
    Dim hHBr As Long, R As RECT, hFRgn As Long, hRRgn As Long, hRPen As Long, LP As LOGPEN
    Dim hFFBrush As Long, mIcon As Long, Cnt As Long
    'Clear the form
    Me.Cls
    'Set the rectangle's values
    SetRect R, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    'Create a new brush
    hHBr = CreateHatchBrush(HS_DIAGCROSS, vbRed)
    'Draw a frame
    FrameRect Me.hdc, R, hHBr
    'Draw a rounded rectangle
    hFRgn = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, (Me.ScaleWidth / 3) * 2, (Me.ScaleHeight / 3) * 5)
    'Draw a frame
    FrameRgn Me.hdc, hFRgn, hHBr, Me.ScaleWidth, Me.ScaleHeight
    'Invert a region
    InvertRgn Me.hdc, hFRgn
    'Move our region
    OffsetRgn hFRgn, 10, 10
    'Create a new region
    hRRgn = CreateRectRgnIndirect(R)
    'Combine our two regions
    CombineRgn hRRgn, hFRgn, hRRgn, RGN_XOR
    'Draw a frame
    FrameRgn Me.hdc, hRRgn, hHBr, Me.ScaleWidth, Me.ScaleHeight
    'Crete a new pen
    hRPen = CreatePen(PS_SOLID, 5, vbBlue)
    'Select our pen into the form's device context and delete the old pen
    DeleteObject SelectObject(Me.hdc, hRPen)
    'Draw a rectangle
    Rectangle Me.hdc, Me.ScaleWidth / 2 - 25, Me.ScaleHeight / 2 - 25, Me.ScaleWidth / 2 + 25, Me.ScaleHeight / 2 + 25
    'Delete our pen
    DeleteObject hRPen
    LP.lopnStyle = PS_DOT
    LP.lopnColor = vbGreen
    'Create a new pen
    hRPen = CreatePenIndirect(LP)
    'Select our pen into the form's device context
    SelectObject Me.hdc, hRPen
    'Draw a rounded rectangle
    RoundRect Me.hdc, Me.ScaleWidth / 2 - 25, Me.ScaleHeight / 2 - 25, Me.ScaleWidth / 2 + 25, Me.ScaleHeight / 2 + 25, 50, 50
    'Create a new solid brush
    hFFBrush = CreateSolidBrush(vbYellow)
    'Select this brush into our form's device context
    SelectObject Me.hdc, hFFBrush
    'Floodfill our form
    FloodFill Me.hdc, Me.ScaleWidth / 2, Me.ScaleHeight / 2, vbBlue
    'Delete our brush
    DeleteObject hFFBrush
    'Create a new solid brush
    hFFBrush = CreateSolidBrush(vbMagenta)
    'Select our solid brush into our form's device context
    SelectObject Me.hdc, hFFBrush
    'Draw a Pie
    Pie Me.hdc, Me.ScaleWidth / 2 - 15, Me.ScaleHeight / 2 - 15, Me.ScaleWidth / 2 + 15, Me.ScaleHeight / 2 + 15, 20, 20, 20, 20
    'Extract icons from 'shell32.dll' and draw them on the form
    For Cnt = 0 To Me.ScaleWidth / 32
        ExtractIconEx "shell32.dll", Cnt, mIcon, ByVal 0&, 1
        DrawIcon Me.hdc, 32 * Cnt, 0, mIcon
        DestroyIcon mIcon
    Next Cnt
    'Clean up
    DeleteObject hFFBrush
    DeleteObject hRPen
    DeleteObject hRRgn
    DeleteObject hFRgn
    DeleteObject hHBr
End Sub
Private Sub Form_Resize()
    Form_Paint
End Sub

