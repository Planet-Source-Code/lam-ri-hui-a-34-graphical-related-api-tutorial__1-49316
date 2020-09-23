VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DrawState API Demo"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   7050
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here"
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   4800
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project needs a picturebox, Picture1, with a picture
'and a button
Const DST_COMPLEX = &H0
Const DST_TEXT = &H1
Const DST_PREFIXTEXT = &H2
Const DST_ICON = &H3
Const DST_BITMAP = &H4
Const DSS_NORMAL = &H0
Const DSS_UNION = &H10 '/* Gray string appearance */
Const DSS_DISABLED = &H20
Const DSS_MONO = &H80
Const DSS_RIGHT = &H8000
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
Private Sub Command1_Click()

    'API uses pixels
    Picture1.ScaleMode = vbPixels
    Picture1.AutoSize = True
    'Dither the image
    DrawState Picture1.hDC, 0, 0, Picture1.Picture, 0, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, DST_BITMAP Or DSS_UNION
End Sub

