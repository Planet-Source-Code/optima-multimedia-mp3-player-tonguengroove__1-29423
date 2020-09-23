VERSION 5.00
Begin VB.Form Mp3Twist 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1200
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   8640
      Width           =   615
   End
End
Attribute VB_Name = "Mp3Twist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JB1 As Integer
Dim JB2 As Integer
Dim JB3 As Integer
Dim JB4 As Integer
Dim nisse As Integer

Dim F1 As Integer
Dim F2 As Integer
Dim F3 As Integer

Dim F11 As Integer
Dim F22 As Integer
Dim F33 As Integer

Dim JB11 As Integer
Dim JB22 As Integer
Dim JB33 As Integer
Dim JB44 As Integer
Private Sub Form_Click()
Dim Rtn As Long
MP3PlayerFrm1.Timer3.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Dim Rtn As Long
MP3PlayerFrm1.Timer3.Enabled = False
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
Me.Width = Screen.Width


Timer1.interval = 1
JB1 = 1
JB2 = 1
JB3 = 1
JB4 = 1
nisse = 1

F11 = 0
F22 = 0
F33 = 0

JB11 = 0
JB22 = 0
JB33 = 0
JB44 = 0
End Sub

Private Sub Label1_Click()
Mp3Twist.Cls
End Sub

Private Sub Timer1_Timer()
nisse = 1
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Label1.Refresh
Do
nisse = nisse + 1
   
        If JB1 > Mp3Twist.Width Then
JB11 = 1
End If
If JB1 < 0 Then
JB11 = 0
End If

If JB2 > Mp3Twist.Height Then
JB22 = 1
End If
If JB2 < 0 Then
JB22 = 0
End If

If JB3 > Mp3Twist.Width Then
JB33 = 1
End If
If JB3 < 0 Then
JB33 = 0
End If

If JB4 > Mp3Twist.Height Then
JB44 = 1
End If
If JB4 < 0 Then
JB44 = 0
End If



If JB11 = 0 Then
JB1 = JB1 + 1
Else
JB1 = JB1 - 4
End If

If JB22 = 0 Then
JB2 = JB2 + 2
Else
JB2 = JB2 - 3
End If

If JB33 = 0 Then
JB3 = JB3 + 3
Else
JB3 = JB3 - 2
End If

If JB44 = 0 Then
JB4 = JB4 + 4
Else
JB4 = JB4 - 1
End If

If F1 >= 255 Then F11 = 1
If F2 >= 255 Then F22 = 1
If F3 >= 255 Then F33 = 1

If F1 <= 0 Then F11 = 0
If F2 <= 0 Then F22 = 0
If F3 <= 0 Then F33 = 0


Mp3Twist.ForeColor = RGB(F1, F2, F3)

Line (JB1, JB2)-(JB3, JB4)

Loop Until nisse > 100
If F11 = 0 Then
F1 = F1 + 1
Else
F1 = F1 - 3
End If

If F22 = 0 Then
F2 = F2 + 2
Else
F2 = F2 - 2
End If

If F33 = 0 Then
F3 = F3 + 3
Else
F3 = F3 - 1
End If
End Sub
