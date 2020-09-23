VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "PrintWindow API Demo"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Sub Form_Load()

    Dim mWnd As Long
    'launch notepad
    Shell "notepad.exe", vbNormalNoFocus
    DoEvents
    'set the graphics mode to persistent
    Me.AutoRedraw = True
    'search the handle of the notepad window
    mWnd = FindWindow("Notepad", vbNullString)
    If mWnd = 0 Then
        Me.Print "NotePad window not found!"
    Else
        'draw the image of the notepad window on our form
        PrintWindow mWnd, Me.hDC, 0
    End If
End Sub

