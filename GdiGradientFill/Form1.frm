VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GdiGradientFillRect API Demo"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TRIVERTEX
    x As Long
    y As Long
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
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Function LongToUShort(Unsigned As Long) As Integer
    'A small function to convert from long to unsigned short
    LongToUShort = CInt(Unsigned - &H10000)
End Function
Private Sub Form_Load()

    'API uses pixels
    Me.ScaleMode = vbPixels
End Sub
Private Sub Form_Paint()
    Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT

    'from black
    With vert(0)
        .x = 0
        .y = 0
        .Red = 0&
        .Green = 0& '&HFF&   '0&
        .Blue = 0&
        .Alpha = 0&
    End With

    'to blue
    With vert(1)
        .x = Me.ScaleWidth
        .y = Me.ScaleHeight
        .Red = 0&
        .Green = 0&
        .Blue = LongToUShort(&HFF00&)
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    GradientFillRect Me.hdc, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H
End Sub

