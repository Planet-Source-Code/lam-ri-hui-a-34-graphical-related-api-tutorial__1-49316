VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OleLoadPicturePath API Demo"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Public Function LoadPicture(ByVal strFileName As String) As Picture
Dim IID  As TGUID
    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
On Error GoTo ERR_LINE
    OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
    Exit Function
ERR_LINE:
    Set LoadPicture = VB.LoadPicture(strFileName)
End Function
Private Sub Form_Load()
    Set Me.Picture = LoadPicture(App.Path & "\" & "bird.jpg")
End Sub

