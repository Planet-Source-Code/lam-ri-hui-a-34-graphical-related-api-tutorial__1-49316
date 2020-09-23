VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ExtFloodFill API Demo"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Const FLOODFILLBORDER = 0 ' Fill until crColor& color encountered.
Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not encountered.
Const crNewColor = &HFFFF80
Dim mBrush As Long
Private Sub Form_Load()

    'Create a solid brush
    mBrush = CreateSolidBrush(crNewColor)
    'Select the brush into the PictureBox' device context
    SelectObject Picture1.hdc, mBrush
    'API uses pixels
    Picture1.ScaleMode = vbPixels
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Delete our new brush
    DeleteObject mBrush
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Floodfill...
    ExtFloodFill Picture1.hdc, x, y, GetPixel(Picture1.hdc, x, y), FLOODFILLSURFACE
End Sub

