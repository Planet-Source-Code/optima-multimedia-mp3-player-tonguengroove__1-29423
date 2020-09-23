VERSION 5.00
Begin VB.Form Mp3Circles 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   ControlBox      =   0   'False
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      Left            =   2160
      Max             =   10
      TabIndex        =   3
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1080
      Top             =   3480
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   0
      Max             =   10
      TabIndex        =   1
      Top             =   8760
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Big Circles"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1030
      TabIndex        =   0
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   480
      Top             =   3480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Circle Speed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Clear Speed"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   8520
      Width           =   975
   End
End
Attribute VB_Name = "Mp3Circles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer

Dim R As Integer
Dim G As Integer
Dim B As Integer

Dim s As Integer

Dim RA As Integer

Private Sub Command2_Click()

End Sub



Private Sub Form_Click()
Dim Rtn As Long

MP3PlayerFrm1.Timer3.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False

'show the taskbar

Unload Me
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
MP3PlayerFrm1.Timer3.Enabled = False
HScroll1.Value = 8
HScroll2.Value = 8

'Hide the taskbar
Me.Move 0, 0, Screen.Width, Screen.Height 'hide the Taskbar 'hide the Taskbar
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
Me.Width = Screen.Width

End Sub

Private Sub Picture1_Click()
Dim Rtn As Long

MP3PlayerFrm1.Timer3.Enabled = True

Unload Me
End Sub


Private Sub Timer1_Timer()
On Error GoTo err
Label1.Refresh
Label2.Refresh
If HScroll2.Value = 0 Then
Pause (1)

End If
If HScroll2.Value = 1 Then
Pause (0.9)

End If
If HScroll2.Value = 2 Then
Pause (0.8)

End If
If HScroll2.Value = 3 Then
Pause (0.7)

End If
If HScroll2.Value = 4 Then
Pause (0.6)

End If
If HScroll2.Value = 5 Then
Pause (0.5)

End If
If HScroll2.Value = 6 Then
Pause (0.4)

End If
If HScroll2.Value = 7 Then
Pause (0.3)

End If
If HScroll2.Value = 8 Then
Pause (0.2)

End If
If HScroll2.Value = 9 Then
Pause (0.1)

End If
If HScroll2.Value = 10 Then
Pause (0)

End If
x = Rnd * Mp3Circles.Width
y = Rnd * Mp3Circles.Height
s = 1

R = Rnd * 255
G = Rnd * 255
B = Rnd * 255


Do

RA = RA + 1

If RA >= 20 And Check1.Value = 0 Then
R = R - 1
G = G - 1
B = B - 1
RA = 0
End If

If RA >= 20 And Check1.Value = 1 Then
R = R + 1
G = G + 1
B = B + 1
RA = 0

If Mp3Circles.ForeColor = RGB(255, 255, 255) Then GoTo err

End If

Mp3Circles.ForeColor = RGB(R, G, B)
s = s + 1
Circle (x, y), s
Loop While s < 10000

Exit Sub
err:
Exit Sub
End Sub

Private Sub Timer2_Timer()
Label1.Refresh
Label2.Refresh
If HScroll1.Value = 0 Then
Pause (60)

End If
If HScroll1.Value = 1 Then
Pause (50)

End If
If HScroll1.Value = 2 Then
Pause (40)

End If
If HScroll1.Value = 3 Then
Pause (30)

End If
If HScroll1.Value = 4 Then
Pause (20)

End If
If HScroll1.Value = 5 Then
Pause (15)

End If
If HScroll1.Value = 6 Then
Pause (10)

End If
If HScroll1.Value = 7 Then
Pause (5)

End If
If HScroll1.Value = 8 Then
Pause (4)

End If
If HScroll1.Value = 9 Then
Pause (3)

End If
If HScroll1.Value = 10 Then
Pause (0)

End If
Mp3Circles.Cls
End Sub
