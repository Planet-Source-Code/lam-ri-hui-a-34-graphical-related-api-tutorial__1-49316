VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FillRgn API Demo"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type COORD
    x As Long
    y As Long
End Type
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Const WINDING = 2 ' constants for FillMode.
Const BLACKBRUSH = 4 ' Constant for brush type.
Private Sub Form_Paint()

    Dim poly(1 To 3) As COORD, NumCoords As Long, hBrush As Long, hRgn As Long
    Me.Cls
    ' Number of vertices in polygon.
    NumCoords = 3
    ' Set scalemode to pixels to set up points of triangle.
    Me.ScaleMode = vbPixels
    ' Assign values to points.
    poly(1).x = Form1.ScaleWidth / 2
    poly(1).y = Form1.ScaleHeight / 2
    poly(2).x = Form1.ScaleWidth / 4
    poly(2).y = 3 * Form1.ScaleHeight / 4
    poly(3).x = 3 * Form1.ScaleWidth / 4
    poly(3).y = 3 * Form1.ScaleHeight / 4
    ' Polygon function creates unfilled polygon on screen.
    ' Remark FillRgn statement to see results.
    Polygon Me.hdc, poly(1), NumCoords
    ' Gets stock black brush.
    hBrush = GetStockObject(BLACKBRUSH)
    ' Creates region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    ' If the creation of the region was successful then color.
    If hRgn Then FillRgn Me.hdc, hRgn, hBrush
    DeleteObject hRgn
End Sub
Private Sub Form_Resize()
    Form_Paint
End Sub

