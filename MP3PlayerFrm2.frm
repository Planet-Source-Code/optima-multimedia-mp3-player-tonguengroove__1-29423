VERSION 5.00
Begin VB.Form MP3PlayerFrm2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   8655
      Left            =   0
      ScaleHeight     =   8595
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.HScrollBar HScroll3 
         Height          =   135
         LargeChange     =   5
         Left            =   1920
         Max             =   70
         Min             =   10
         TabIndex        =   3
         Top             =   8400
         Value           =   10
         Width           =   735
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         LargeChange     =   50
         Left            =   1080
         Max             =   800
         Min             =   10
         SmallChange     =   25
         TabIndex        =   2
         Top             =   8400
         Value           =   10
         Width           =   735
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         LargeChange     =   3
         Left            =   240
         Max             =   15
         Min             =   1
         TabIndex        =   1
         Top             =   8400
         Value           =   1
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   8160
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   8160
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scatter"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   8160
         Width           =   735
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   1080
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   600
   End
End
Attribute VB_Name = "MP3PlayerFrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NowBeingUsed(100) As Boolean    'Do not touch it!
Dim Direction(100) As Integer       'the direction of the each flame
Dim Speed(100) As Integer           'the speed of the each flame
Dim y(100) As Integer               'the location of the each flame
Dim x(100) As Integer               '               "
Dim COLORRGB(100, 2) As Long        'the color of the each flame(RGB)
                                        'ColorRGB(n, 0) -> Red value
                                'ColorRGB(n, 1) -> Green value
                                        'ColorRGB(n, 2) -> Blue value

Dim Scatter As Integer              'It decides how far the flames are
                                    'scattered when they launched
Dim ScatterFalling As Integer       'It decides how far the flames are
                                    'scattered when they start falling.
Dim SpeedDif As Integer             'The larger it is, the distance between
                                    'the higher flame and the lower flame
                                    'become bigger.
Dim Size As Integer

Private Sub Form_Load()
Dim Rtn As Long

HScroll2.Value = 600
HScroll1.Value = 9
ScatterFalling = 40
SpeedDif = 20
HScroll3.Value = 60
MP3PlayerFrm1.Timer3.Enabled = False


'Hide the taskbar
Me.Move 0, 0, Screen.Width, Screen.Height 'hide the Taskbar
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
Picture1.Height = Me.Height
Picture1.Width = Me.Width
End Sub

Private Sub Picture1_Click()
Dim Rtn As Long

MP3PlayerFrm1.Timer3.Enabled = True


Unload Me
End Sub

Private Sub Timer1_Timer()
        'Whenever this timer work, all of the flames move as their
        'direction & speed
        Label1.Refresh
        Label2.Refresh
        Label3.Refresh
       
        Timer2.interval = HScroll2.Value + 200
Scatter = HScroll1.Value
ScatterFalling = 40
SpeedDif = 20
Size = HScroll3.Value
For i = 1 To 100
If NowBeingUsed(i) = False Then GoTo 20
        'If the flame isn't on the screen, we don't need to process.
Picture1.Line (x(i) - Size, y(i) - Size)-(x(i) + Size, y(i) + Size), 0, BF
        'erase the original flame.
x(i) = x(i) - Cos(Direction(i) * 3.141592654 / 180) * Speed(i)
y(i) = y(i) - Sin(Direction(i) * 3.141592654 / 180) * Speed(i)
        'move the location as its direction & speed
If Direction(i) >= 0 And Direction(i) <= 180 Then
        Speed(i) = Speed(i) - 7
        'The higher it fly, the slower its speed become.
        '(because of the gravity) so I make it slower as it goes higher.
        If Speed(i) < 0 Then
            Speed(i) = 0
            Direction(i) = Int(Rnd(1) * ScatterFalling * 2) + 270 - ScatterFalling
            'if its speed became zero, it should be fell down.
        End If
Else
        Speed(i) = Speed(i) + 7
            'I make the speed fast as it fall down
            'because of the gravity
        If Speed(i) > 80 Then
            'if the flame should be removed,
            NowBeingUsed(i) = False: GoTo 20
        Else
            GoTo 25
        End If
End If
'Don't touch - Picture1.Line (X(i) - Size, Y(i) - Size)-(X(i) + Size, Y(i) + Size), RGB(ColorRGB(i, 0), ColorRGB(i, 1), ColorRGB(i, 2)), BF: GoTo 20
A1 = Picture1.ScaleHeight / 40 + 150 + SpeedDif
A2 = ((A1 - Speed(i)) / 2 + A1 / 2) / A1
Picture1.Line (x(i) - Size * A2, y(i) - Size * A2)-(x(i) + Size * A2, y(i) + Size * A2), RGB(COLORRGB(i, 0), COLORRGB(i, 1), COLORRGB(i, 2)), BF: GoTo 20
'I draw a new flame for it (it is going higher)
25
R = COLORRGB(i, 0) * (80 - Speed(i)) / 80
G = COLORRGB(i, 1) * (80 - Speed(i)) / 80
B = COLORRGB(i, 2) * (80 - Speed(i)) / 80
Picture1.Line (x(i) - Size, y(i) - Size)-(x(i) + Size, y(i) + Size), RGB(R, G, B), BF
'I draw a new flame for it (it is falling down)
'As it fall down, its color become dark.
'Because it will be removed soon, so it should be ready for being removed.
20 Next i
End Sub

Private Sub Timer2_Timer()
Randomize Timer
10 R = Int(Rnd(1) * 128) + 128
   G = Int(Rnd(1) * 128) + 128
   B = Int(Rnd(1) * 128) + 128  'Decide the color of the fire
If Abs(R - G) < 30 And Abs(G - B) < 30 And Abs(R - B) < 30 Then GoTo 10
    'if the color is like gray, it makes a new color
    '(because I don't like gray color).
LocationX = Int(Rnd(1) * (Picture1.ScaleWidth - 2000)) + 1000
MainDirection = Int(Rnd(1) * 40) + 70   'Main Direction
MainSpeed = Picture1.ScaleHeight / 40 + 50 + Int(Rnd(1) * 101) 'Main Speed
For i = 1 To 100
If NowBeingUsed(i) = True Then GoTo 30 'If it is being used, I can't use it
NowBeingUsed(i) = True

x(i) = LocationX
y(i) = Picture1.ScaleHeight
COLORRGB(i, 0) = R
COLORRGB(i, 1) = G
COLORRGB(i, 2) = B      'Save the values.

Direction(i) = MainDirection + Int(Rnd(1) * Scatter * 2) - Scatter
Speed(i) = MainSpeed + Int(Rnd(1) * SpeedDif * 2) - SpeedDif
30 Next i
End Sub


