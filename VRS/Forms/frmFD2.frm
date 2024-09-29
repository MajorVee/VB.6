VERSION 5.00
Begin VB.Form frmFD2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFD2.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmFD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim trans As Integer

Private Sub Form_Load()
    trans = 255
End Sub

Private Sub Timer1_Timer()
    trans = trans - 5
    set_transparency Me.hwnd, trans
    
    If (trans <= 0) Then
        frmFD1.Show
    End If
End Sub

