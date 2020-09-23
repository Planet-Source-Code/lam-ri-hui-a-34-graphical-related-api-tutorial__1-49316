VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type
Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Private Declare Function GradientFillTriangle Lib "msimg32" _
Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, _
ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, _
ByVal dwMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Sub Form_Load()
    Dim vert(4) As TRIVERTEX
    Dim gTRi(1) As GRADIENT_TRIANGLE
    ScaleMode = vbPixels
    AutoRedraw = True
    Move Left, Top, 3945, 4230
    vert(0).X = 0
    vert(0).Y = 0
    vert(0).Red = -256
    vert(0).Green = 0&
    vert(0).Blue = 0&
    vert(0).Alpha = 0&
    
    vert(1).X = 255
    vert(1).Y = 0
    vert(1).Red = 0&
    vert(1).Green = -256
    vert(1).Blue = 0&
    vert(1).Alpha = 0&
    
    vert(2).X = 256
    vert(2).Y = 256
    vert(2).Red = 0&
    vert(2).Green = 0&
    vert(2).Blue = -256
    vert(2).Alpha = 0&
    
    vert(3).X = 0
    vert(3).Y = 256
    vert(3).Red = -256
    vert(3).Green = -256
    vert(3).Blue = -256
    vert(3).Alpha = 0&
    
    gTRi(0).Vertex1 = 0
    gTRi(0).Vertex2 = 1
    gTRi(0).Vertex3 = 2
    
    gTRi(1).Vertex1 = 0
    gTRi(1).Vertex2 = 2
    gTRi(1).Vertex3 = 3
    GradientFillTriangle hDC, vert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
    Form1.Show
End Sub
Private Function RgbParse(hDC As Long, X As Single, Y As Single) As String
    Dim ColorMe As Long
    ColorMe = GetPixel(hDC, X, Y)
    Dim rgbRed, rgbGreen, rgbBlue As Long
    rgbRed = Abs(ColorMe Mod &H100)
    ColorMe = Abs(ColorMe \ &H100)
    rgbGreen = Abs(ColorMe Mod &H100)
    ColorMe = Abs(ColorMe \ &H100)
    rgbBlue = Abs(ColorMe Mod &H100)
    ColorMe = RGB(rgbRed, rgbGreen, rgbBlue)
    RgbParse = "RGB(" & rgbRed & ", " & rgbGreen & ", " & rgbBlue & ")"
End Function
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Caption = RgbParse(hDC, X, Y)
End Sub



