Attribute VB_Name = "transp_module"
Option Explicit
Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Sub set_transparency(handle As Long, val As Integer)

    If val < 0 Then val = 0
    If val > 255 Then val = 255
    
    SetWindowLong handle, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes handle, RGB(255, 255, 0), val, LWA_ALPHA

End Sub

