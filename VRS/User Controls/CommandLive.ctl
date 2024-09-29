VERSION 5.00
Begin VB.UserControl CommandLive 
   BackColor       =   &H00400040&
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5040
   ScaleWidth      =   7890
   ToolboxBitmap   =   "CommandLive.ctx":0000
   Begin VB.Frame Line1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2295
      Index           =   3
      Left            =   3690
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Frame Line1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2295
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   -120
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Frame Line1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1770
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Frame Line1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   30
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5640
      Top             =   1920
   End
   Begin VB.Timer tmrAnimationEnd 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   1920
   End
   Begin VB.Timer tmrControl 
      Enabled         =   0   'False
      Left            =   5160
      Top             =   1440
   End
   Begin VB.Timer tmrAnimationStart 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   1920
   End
   Begin VB.Timer tmrIF 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   960
   End
   Begin VB.PictureBox Image1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   120
      Picture         =   "CommandLive.ctx":0314
      ScaleHeight     =   1800
      ScaleWidth      =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   2520
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   540
      End
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   7440
      Picture         =   "CommandLive.ctx":4CAD
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label lblDetail 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   1185
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3540
   End
   Begin VB.Label lblCaption2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   540
   End
End
Attribute VB_Name = "CommandLive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Dim i As Integer
Dim m_BackColor As OLE_COLOR
Event Click()
Event DblClick()
Event KeyPress(KeyAscii As Integer)

Private Sub Image1_Click()
RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Image1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CtlMouseOver
End Sub

Private Sub lblCaption_Click()
RaiseEvent Click
End Sub

Private Sub lblCaption_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lblCaption2_Click()
RaiseEvent Click
End Sub

Private Sub lblCaption2_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lblDetail_Click()
RaiseEvent Click
End Sub

Private Sub lblDetail_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lblDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CtlMouseOver
End Sub

Private Sub tmrIF_Timer()
On Error Resume Next
    Dim p As POINTAPI
    Dim R As RECT
    GetWindowRect UserControl.hWnd, R
    GetCursorPos p
    If p.X < R.Left Or p.X > R.Right Or p.Y < R.Top Or p.Y > R.Bottom Then
    For i = 1 To 4
    
    Line1(i).Visible = False
    Next i
    tmrIF.Enabled = False
    End If
End Sub
Private Sub CtlMouseOver()
    For i = 1 To 4
    Line1(i).Visible = True
    Next i
    tmrIF.Enabled = True
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    For i = 1 To 4
    Line1(i).Visible = True
    Next i
End Sub
Public Property Get Interval() As String
    Interval = tmrControl.Interval
End Property

Public Property Let Interval(ByVal New_Interval As String)
    tmrControl.Interval() = New_Interval
    UserControl_Resize
    PropertyChanged "Interval"
End Property
Public Property Get Animate() As Boolean
    Animate = tmrControl.Enabled
End Property

Public Property Let Animate(ByVal New_Animate As Boolean)
    tmrControl.Enabled() = New_Animate
    PropertyChanged "Animate"
    
    If New_Animate = False Then
    tmrAnimationStart.Enabled = False
    tmrAnimationEnd.Enabled = False
    tmrControl.Enabled = False
    tmrWait.Enabled = False
    UserControl_Resize
    Else
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get BorderColor() As OLE_COLOR
For i = 1 To 4
    BorderColor = Line1(i).BackColor
Next i
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
For i = 1 To 4
    Line1(i).BackColor = New_BorderColor
Next i
    PropertyChanged "BorderColor"
End Property
Public Property Get Picture() As Picture
    Set Picture = Image1.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
    Set Image1.Picture = New_Picture
    UserControl_Show
    PropertyChanged "Picture"
End Property
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property
Public Property Get Detail() As String
    Detail = lblDetail.Caption
End Property

Public Property Let Detail(ByVal New_Detail As String)
    lblDetail.Caption() = New_Detail
    UserControl_Resize
    PropertyChanged "Detail"
End Property

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CtlMouseOver
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    tmrControl.Interval = PropBag.ReadProperty("Interval", "0")
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Title")
    lblDetail.Caption = PropBag.ReadProperty("Detail", "Detail")
    tmrControl.Enabled = PropBag.ReadProperty("Animate", False)
    
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H40C0&)
    For i = 1 To 4
    Line1(i).BackColor = PropBag.ReadProperty("BorderColor", &H8000000D)
    Next i
    Set Image1.Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Picture", Image1.Picture, Nothing)
    Call PropBag.WriteProperty("Interval", tmrControl.Interval, "0")
    Call PropBag.WriteProperty("Animate", tmrControl.Enabled, False)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Caption")
    Call PropBag.WriteProperty("Detail", lblDetail.Caption, "Detail")
    For i = 1 To 4
    Call PropBag.WriteProperty("BorderColor", Line1(i).BackColor)
    Next i
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor)
End Sub


Private Sub UserControl_Show()
If Image1.Picture = 0 Then
Image1.Picture = Image2.Picture
Else
End If
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
Image1.Top = 0
Image1.Left = 0
UserControl.Width = Image1.Width
UserControl.Height = Image1.Height
Line1(3).Left = Image1.Width - Line1(3).Width
Line1(1).Top = Image1.Height - Line1(1).Height
lblDetail.Width = Image1.Width - (lblDetail.Left * 2)

End Sub





Private Sub tmrAnimationEnd_Timer()
If Image1.Top >= -600 Then
Image1.Top = Image1.Top + 11.5
''PC(1).Top = image1.Top + image1.Height
Else
Image1.Top = Image1.Top + 41.5
''PC(1).Top = image1.Top + image1.Height
End If

If Image1.Top >= -1 Then
Image1.Top = 0
''PC(1).Top = image1.Height
tmrAnimationStart.Enabled = False
tmrAnimationEnd.Enabled = False
tmrControl.Enabled = True
End If
End Sub
Private Sub tmrAnimationStart_Timer()
If Image1.Top <= -1200 Then
Image1.Top = Image1.Top - 11.5
'PC(1).Top = Image1.Top + Image1.Height
Else
Image1.Top = Image1.Top - 41.5
'PC(1).Top = Image1.Top + Image1.Height
End If

If Image1.Top <= -Image1.Height Then
Image1.Top = -Image1.Height
'PC(1).Top = 0
tmrAnimationStart.Enabled = False
Image1.Refresh
'PC(1).Refresh
tmrWait.Enabled = True
End If
End Sub

Private Sub tmrControl_Timer()
tmrAnimationStart.Enabled = True
lblCaption2.Caption = lblCaption.Caption
tmrControl.Enabled = False
End Sub

Private Sub tmrWait_Timer()
tmrAnimationStart.Enabled = False
tmrAnimationEnd.Enabled = True
tmrWait.Enabled = False
End Sub










