VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GetDeviceGammaRamp and SetDeviceGammaRamp API Demo"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You better close this via the form's X button instead of the stop button of MS VB
Option Explicit
Private Ramp1(0 To 255, 0 To 2) As Integer
Private Ramp2(0 To 255, 0 To 2) As Integer
Private Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Private Declare Function SetDeviceGammaRamp Lib "gdi32" (ByVal hdc As Long, lpv As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Form_Load()
   '----------------------------------------------------------------
   Dim iCtr       As Integer
   Dim lVal       As Long
   '----------------------------------------------------------------
   GetDeviceGammaRamp Me.hdc, Ramp1(0, 0)
      For iCtr = 0 To 255
         lVal = Int2Lng(Ramp1(iCtr, 0))
         Ramp2(iCtr, 0) = Lng2Int(Int2Lng(Ramp1(iCtr, 0)) / 2)
         
         Ramp2(iCtr, 1) = Lng2Int(Int2Lng(Ramp1(iCtr, 1)) / 2)
         Ramp2(iCtr, 2) = Lng2Int(Int2Lng(Ramp1(iCtr, 2)) / 2)
      Next iCtr
   SetDeviceGammaRamp Me.hdc, Ramp2(0, 0)
   '----------------------------------------------------------------
End Sub
Private Sub Form_Unload(Cancel As Integer)
   '----------------------------------------------------------------
   SetDeviceGammaRamp Me.hdc, Ramp1(0, 0)
   '----------------------------------------------------------------
End Sub
Public Function Int2Lng(IntVal As Integer) As Long
   '----------------------------------------------------------------
   CopyMemory Int2Lng, IntVal, 2
   '----------------------------------------------------------------
End Function
Public Function Lng2Int(Value As Long) As Integer
   '----------------------------------------------------------------
   CopyMemory Lng2Int, Value, 2
   '----------------------------------------------------------------
End Function

