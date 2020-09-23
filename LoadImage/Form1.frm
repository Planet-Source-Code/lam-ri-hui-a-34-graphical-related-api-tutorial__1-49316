VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LoadImage Api Demo"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LR_LOADFROMFILE = &H10
Const IMAGE_BITMAP = 0
Const IMAGE_ICON = 1
Const IMAGE_CURSOR = 2
Const IMAGE_ENHMETAFILE = 3
Const CF_BITMAP = 2
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Sub Form_Load()

    Dim hDC As Long, hBitmap As Long
    'Load the bitmap into the memory
    hBitmap = LoadImage(App.hInstance, "c:\windows\logow.sys", IMAGE_BITMAP, 320, 200, LR_LOADFROMFILE)
    If hBitmap = 0 Then
        MsgBox "There was an error while loading the bitmap"
        Exit Sub
    End If
    'open the clipboard
    OpenClipboard Me.hwnd
    'Clear the clipboard
    EmptyClipboard
    'Put our bitmap onto the clipboard
    SetClipboardData CF_BITMAP, hBitmap
    'Check if there's a bitmap on the clipboard
    If IsClipboardFormatAvailable(CF_BITMAP) = 0 Then
        MsgBox "There was an error while pasting the bitmap to the clipboard!"
    End If
    'Close the clipboard
    CloseClipboard
    'Get the picture from the clipboard
    Me.Picture = Clipboard.GetData(vbCFBitmap)
End Sub

