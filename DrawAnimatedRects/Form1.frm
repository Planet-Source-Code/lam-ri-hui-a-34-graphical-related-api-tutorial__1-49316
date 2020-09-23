VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DrawAnimatedRects API Demo"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const IDANI_OPEN = &H1
Const IDANI_CLOSE = &H2
Const IDANI_CAPTION = &H3
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawAnimatedRects Lib "User32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Sub Form_Load()

    Dim rSource As RECT, rDest As RECT, ScreenWidth As Long, ScreenHeight As Long
    'retrieve the screen width and height
    ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
    'set the source and destination rects
    SetRect rSource, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight
    SetRect rDest, 0, 0, 200, 200
    'animate
    DrawAnimatedRects Me.hWnd, IDANI_CLOSE Or IDANI_CAPTION, rSource, rDest
    'set the form's position
    Me.Move 0, 0, 200 * Screen.TwipsPerPixelX, 200 * Screen.TwipsPerPixelY
End Sub

