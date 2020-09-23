VERSION 5.00
Begin VB.Form MP3CDFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1185
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   6000
   ControlBox      =   0   'False
   ForeColor       =   &H00B31AB3&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MP3CDFrm.frx":0000
   ScaleHeight     =   1185
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Random"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   820
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   820
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   1200
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Random"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1480
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percent Played/0%"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CDLength/"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   255
      Left            =   4130
      TabIndex        =   4
      Top             =   825
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   120
      Picture         =   "MP3CDFrm.frx":530E
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Track No:"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   5160
      Picture         =   "MP3CDFrm.frx":66CB
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CD Player"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   5295
   End
   Begin VB.Menu mnuCD 
      Caption         =   "CDMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open CD Tray"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close CD Tray"
      End
   End
End
Attribute VB_Name = "MP3CDFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FFSpeed As Long    'Seconds to seek for ff/rew
Dim CDPlaying As Boolean        'true if CD is currently playing
Dim CDLoaded As Boolean         'true if CD is the the player
Dim NumTracks As Integer        'number of Tracks on audio CD
Dim TrackLength() As String     'array containing length of each Track
Dim Track As Integer            'current Track
Dim Min As Integer              'current Minute on Track
Dim Sec As Integer              'current Second on Track
Dim Cmd As String               'string to hold mci command strings
Dim TotalTrackTime As String    'For Display.
Dim TotalTrackPlay As String    'For Display.
Private Sub Combo1_Click()
Static s As String * 30
mciSendString "status cd length wait", s, Len(s), 0
Sup = Mid(s, 1, 5)
Label4 = "CDLength/" & Sup
Timer1.Enabled = True
Send "play cd from " & Val(Combo1.Text)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
MP3PlayerFrm1.mnuID3.Enabled = False
MP3PlayerFrm1.Timer4.Enabled = False
Me.Combo1.ForeColor = MP3playerShowFav.Option1.BackColor
Me.Label4.ForeColor = MP3playerShowFav.Option1.BackColor
Me.Label5.ForeColor = MP3playerShowFav.Option1.BackColor
Me.Option1.BackColor = MP3playerShowFav.Option1.BackColor
Me.Option2.BackColor = MP3playerShowFav.Option1.BackColor
Static s As String * 30
mciSendString "status cd length wait", s, Len(s), 0
Sup = Mid(s, 1, 5)
Label4 = "CDLength/" & Sup
If Combo1.ListCount = 0 Then
MsgBox "Could not detect a CD in the drive"
MP3PlayerFrm1.Label2 = "00:00"
MP3PlayerFrm1.Label3 = "00:00"
Unload Me
End If
MP3PlayerFrm1.Label1.Caption = "             ^v^  Tongue 'N Groove MP3/CD Player  ^v^  "
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Picture = MP3PlayerFrm1.Picture
Send "Open All"
Combo1.ForeColor = Label4.ForeColor
Label5.ForeColor = Label4.ForeColor
If (Send("open cdaudio alias cd wait shareable") = False) Then
Send "Close all"
Send "Open All"
End If
Send "set cd time format tmsf wait"
Static s As String * 30
mciSendString "status cd number of Tracks wait", s, Len(s), 0
NumTracks = CInt(Mid(s, 1, 2))
For i = 1 To NumTracks
Combo1.AddItem i
Next i


MP3CDFrm.Width = MP3PlayerFrm1.Width + 25
MP3CDFrm.Height = 1275
MP3CDFrm.Top = MP3PlayerFrm1.Top + 1600
MP3CDFrm.Left = MP3PlayerFrm1.Left

End Sub

Private Sub Image1_Click()
Send "stop cd wait"
Send "Close all"
MP3PlayerFrm1.Label3 = "00:00"
MP3PlayerFrm1.Label2 = "00:00"
MP3PlayerFrm1.mnuID3.Enabled = True
Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Image4_Click()
Call PopupMenu(mnuCD)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = vbBlack
Call MoveForm(Me)
Label1.ForeColor = vbYellow
End Sub

Private Sub mnuClose_Click()
Send "set cd door closed"
End Sub

Private Sub mnuOpen_Click()
Send "set cd door open"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim Os3 As Integer
Static s As String * 30
Cmd = "status cd length Track " & Val(MP3CDFrm.Combo1.Text)
mciSendString Cmd, s, Len(s), 0
Leng = Mid(s, 1, 5)
MP3PlayerFrm1.Label3 = Leng
If MP3PlayerFrm1.Label3 = "" Then MP3PlayerFrm1.Label3 = "00:00"
mciSendString "status cd position", s, Len(s), 0
Leng = Mid(s, 4, 5)
MP3PlayerFrm1.Label2 = Leng
Track = Mid(s, 1, 2)
Combo1.Text = Track


Op = Mid(MP3PlayerFrm1.Label3, 1, 2)
Op = Val(Op) * 60
Op2 = Mid(MP3PlayerFrm1.Label3, 4, 2)
op3 = Op + Val(Op2)
Os = Mid(MP3PlayerFrm1.Label2, 1, 2)
Os = Val(Os) * 60
Os2 = Mid(MP3PlayerFrm1.Label2, 4, 2)
Os3 = Os + Val(Os2)
MP3PlayerFrm1.Picture1.ScaleWidth = 143
Cent = Percent(Os3, op3, MP3PlayerFrm1.Picture1.Width / 100 * 5.3)
MP3PlayerFrm1.Picture2.Left = Cent

If Option1.Value = True Then
If Combo1.Text = Combo1.ListCount Then
If MP3PlayerFrm1.Picture2.Left > 113.9 Then
Send "stop cd wait"
Send "seek cd to" & "1"
Combo1.Text = 1
Send "play cd from " & Val(MP3CDFrm.Combo1.Text)
MP3PlayerFrm1.Picture2.Left = 0
End If
End If
End If
If Option2.Value = True Then
If MP3PlayerFrm1.Picture2.Left > 113.99 Then
Combo1.ListIndex = Int(Combo1.ListCount * Rnd)
Send "play cd from " & Val(Combo1.Text)
End If
End If
End Sub
