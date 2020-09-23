VERSION 5.00
Begin VB.Form Mp3Pipes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   -135
   ClientTop       =   -270
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll3 
      Height          =   135
      Left            =   1440
      Max             =   7
      Min             =   1
      TabIndex        =   2
      Top             =   8760
      Value           =   1
      Width           =   615
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   10
      Left            =   720
      Max             =   600
      Min             =   50
      TabIndex        =   1
      Top             =   8760
      Value           =   50
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   10
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   8760
      Value           =   1
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Clear"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   8520
      Width           =   615
   End
End
Attribute VB_Name = "Mp3Pipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CtX, CtY, Lef As Boolean, Righ As Boolean, Up As Boolean, down As Boolean, OldX, OldY, OldR, Clr As Boolean, B, C, TY



Private Sub Form_Click()
Dim Rtn As Long
MP3PlayerFrm1.Timer3.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
HScroll1.Value = 20
HScroll2.Value = 550
HScroll3.Value = 3
Dim Rtn As Long
MP3PlayerFrm1.Timer3.Enabled = False

Me.WindowState = 2
'Hide the taskbar
Me.Move 0, 0, Screen.Width, Screen.Height 'hide the Taskbar
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
Clr = True
Lef = False
Up = False
down = False
Righ = False
CtX = 0
CtY = 0
C = 0
End Sub
Private Sub DrawBall()
Dim R
R = Me.HScroll2
If Clr = False Then B = B - 10
If Clr = True Then B = B + 10
If B > 255 Then Clr = False
If B <= 0 Then Clr = True
If B <= 0 Then C = C + 1
If C = Me.HScroll3 Then C = 0
If Lef = True Then CtX = CtX - 30
If Righ = True Then CtX = CtX + 30
If Up = True Then CtY = CtY - 30
If down = True Then CtY = CtY + 30
Dim Q
For Q = 0 To R Step Me.HScroll1
If C = 0 Then Me.Circle (CtX, CtY), Q, RGB(B, 0, 0)
If C = 1 Then Me.Circle (CtX, CtY), Q, RGB(B, B, 0)
If C = 2 Then Me.Circle (CtX, CtY), Q, RGB(B, B, B)
If C = 3 Then Me.Circle (CtX, CtY), Q, RGB(B, 0, B)
If C = 4 Then Me.Circle (CtX, CtY), Q, RGB(0, B, 0)
If C = 5 Then Me.Circle (CtX, CtY), Q, RGB(0, 0, B)
If C = 6 Then Me.Circle (CtX, CtY), Q, RGB(0, B, B)
Next Q
OldX = CtX
OldY = CtY
OldR = R
If CtX < R + 50 Then Righ = True
If CtX < R + 50 Then Lef = False
If CtX > Me.Width - R + 50 Then Righ = False
If CtX > Me.Width - R + 50 Then Lef = True
If CtY < R + 50 Then down = True
If CtY < R + 50 Then Up = False
If CtY > Me.Height - R + 50 Then down = False
If CtY > Me.Height - R + 50 Then Up = True
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Label4_Click()
Mp3Pipes.Cls
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Do
DrawBall
DoEvents
Label1.Refresh
Label2.Refresh
Label3.Refresh
Label4.Refresh
Loop Until Label1.ForeColor = vbBlack
End Sub


