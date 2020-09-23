VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DrawCaption, DrawEdge, DrawFocusRect and DrawFrameControl API Demo"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DC_ACTIVE = &H1
Const DC_NOTACTIVE = &H2
Const DC_ICON = &H4
Const DC_TEXT = &H8
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const DFC_BUTTON = 4
Const DFC_POPUPMENU = 5            'Only Win98/2000 !!
Const DFCS_BUTTON3STATE = &H10
Const DT_CENTER = &H1
Const DC_GRADIENT = &H20          'Only Win98/2000 !!
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Sub Form_Paint()

    Dim R As RECT
    'Clear the form
    Me.Cls
    'API uses pixels
    Me.ScaleMode = vbPixels
    'Set the rectangle's values
    SetRect R, 0, 0, Me.ScaleWidth, 20
    'Draw a caption on the form
    DrawCaption Me.hWnd, Me.hdc, R, DC_ACTIVE Or DC_ICON Or DC_TEXT Or DC_GRADIENT
    'Move the recatangle
    OffsetRect R, 0, 22
    'Draw an edge on our window
    DrawEdge Me.hdc, R, EDGE_ETCHED, BF_RECT
    OffsetRect R, 0, 22
    'Draw a focus rectangle on our window
    DrawFocusRect Me.hdc, R
    OffsetRect R, 0, 22
    'Draw a frame control on our window
    DrawFrameControl Me.hdc, R, DFC_BUTTON, DFCS_BUTTON3STATE
    OffsetRect R, 0, 22
    'draw some text on our form
    DrawText Me.hdc, "Hello World !", Len("Hello World !"), R, DT_CENTER
End Sub

