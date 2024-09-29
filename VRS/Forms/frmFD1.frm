VERSION 5.00
Begin VB.Form frmFD1 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "frmFD1.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmFD1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim trans As Integer

Private Sub Timer1_Timer()
    trans = trans - 5
    set_transparency Me.hwnd, trans
    
    If (trans <= 0) Then
        frmFD2.Show
    End If
End Sub
Private Sub Form_Load()
frmLock.Show 1
trans = 255
End Sub
