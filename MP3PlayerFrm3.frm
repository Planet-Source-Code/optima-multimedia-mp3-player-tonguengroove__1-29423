VERSION 5.00
Begin VB.Form MP3PlayerFrm3 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   2280
   End
   Begin VB.Shape iCirc 
      BorderWidth     =   3
      Height          =   495
      Index           =   0
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
End
Attribute VB_Name = "MP3PlayerFrm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Awsome form FX
'Author: Dustin Davis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'Edited  By (Arc) Software Design
Public Num As Long
Public ResizeOk As Boolean
Public Col As Boolean

Public Sub Resize()
Dim i As Long
For i = 1 To Num
        Unload iCirc(i) 'must unload them from memory first, This
        'is not nessesary, but way more easily done than not
Next i
Create_Circles
End Sub

Private Sub Form_Click()
Dim Rtn As Long
MP3PlayerFrm1.Timer1.Enabled = True
MP3PlayerFrm1.Timer2.Enabled = True
MP3PlayerFrm1.Timer3.Enabled = True
'show the taskbar
Rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(Rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar

Timer1.Enabled = False
Pause (2)
Unload Me
End Sub

Private Sub Form_Load()
Dim Rtn As Long
MP3PlayerFrm1.Timer1.Enabled = False
MP3PlayerFrm1.Timer2.Enabled = False
MP3PlayerFrm1.Timer3.Enabled = False
Me.WindowState = 2
'Hide the taskbar
Rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(Rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Taskbar
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
Me.Width = Screen.Width
Num = 512
Create_Circles
Timer1.Enabled = True
End Sub
Public Sub Create_Circles()
'This loads the circles into memory, and places them in the
'right spot and then resizes them.
Dim i As Long
Dim X As Long
Dim Y As Long
Dim x1 As Long
Dim y1 As Long
Dim cx As Long
Dim cy As Long
Dim c As Long
'Get the absolute center
X = MP3PlayerFrm3.Width
Y = MP3PlayerFrm3.Height
cx = iCirc(0).Width
cy = iCirc(0).Height
iCirc(0).Visible = False
Col = False 'This is the variable that i use to figure out if I should go forward
            'or backwards with the colors
c = 0 'The color variable
ResizeOk = False 'Make sure nothing happens
For i = 1 To Num
   Load iCirc(i) 'Load a new circle into memory
        'Set the circles height and width
    iCirc(i).Height = cy + (i * 10)
    iCirc(i).Width = cx + (i * 10)
    'Set the circle in the absolute center
    x1 = iCirc(i).Width / 2
    y1 = iCirc(i).Height / 2
    iCirc(i).Top = (Y / 2) - y1
    iCirc(i).Left = (X / 2) - x1
    DoEvents
    Next i
Show_all 'Show the circles
End Sub

Public Sub Show_all()
'This is where the circles become visible. This is faster than doing it in the
'Create circles function
Dim i As Long

'Start showing the circles
For i = 1 To Num
    iCirc(i).Visible = True
Next i

End Sub

Public Sub Color_Red()
'This colors the circle red
Dim r As Long
Dim c As Long
c = 0
Col = False

For r = 1 To Num
    If Col = False Then
        iCirc(r).BorderColor = RGB(c, 0, 0)
        c = c + 1
    ElseIf Col = True Then
        iCirc(r).BorderColor = RGB(c, 0, 0)
        c = c - 1
    End If
    If c >= 256 Then
        Col = True
    ElseIf c <= 1 Then
        Col = False
    End If
Next r
ResizeOk = True
End Sub

Public Sub Color_Blue()
'This colors the circle blue
Dim b As Long
Dim c As Long
c = 0
Col = False

i = 1
For b = 1 To Num
    If Col = False Then
        iCirc(b).BorderColor = RGB(0, 0, c)
        c = c + 1
    ElseIf Col = True Then
        iCirc(b).BorderColor = RGB(0, 0, c)
        c = c - 1
    End If
    If c >= 256 Then
        Col = True
    ElseIf c <= 1 Then
        Col = False
    End If
    Next b

End Sub

Public Sub Color_green()
'This colors the circle green
Dim g As Long
Dim c As Long
c = 0
Col = False

For g = 1 To Num
    If Col = False Then
        iCirc(g).BorderColor = RGB(0, c, 0)
        c = c + 1
    ElseIf Col = True Then
        iCirc(g).BorderColor = RGB(0, c, 0)
        c = c - 1
    End If
    If c >= 256 Then
        Col = True
    ElseIf c <= 1 Then
        Col = False
    End If
Next g

End Sub

Public Sub Color_Purple()
'This colors the circle purple
Dim p As Long
Dim c As Long
c = 0
Col = False

For p = 1 To Num
    If Col = False Then
        iCirc(p).BorderColor = RGB(c, 0, c)
        c = c + 1
    ElseIf Col = True Then
        iCirc(p).BorderColor = RGB(c, 0, c)
        c = c - 1
    End If
    If c >= 256 Then
        Col = True
    ElseIf c <= 1 Then
        Col = False
    End If
Next p

End Sub

Public Sub Color_yellow()
'This colors the circle yellow
Dim o As Long
Dim c As Long
c = 0
Col = False

For o = 1 To Num
    If Col = False Then
        iCirc(o).BorderColor = RGB(c, c, 0)
        c = c + 1
    ElseIf Col = True Then
        iCirc(o).BorderColor = RGB(c, c, 0)
        c = c - 1
    End If
    If c >= 256 Then
        Col = True
    ElseIf c <= 1 Then
        Col = False
    End If
Next o

End Sub

Public Sub Color_seafoam()
'This colors the circle a sea foam type color
Dim Y As Long
Dim c As Long
c = 0
Col = False

For Y = 1 To Num
    If Col = False Then
        iCirc(Y).BorderColor = RGB(0, c, c)
        c = c + 1
    ElseIf Col = True Then
        iCirc(Y).BorderColor = RGB(0, c, c)
        c = c - 1
    End If
    If c >= 256 Then
        Col = True
    ElseIf c <= 1 Then
        Col = False
    End If
Next Y

End Sub


Private Sub Timer1_Timer()
Do
DoEvents
Color_Blue
Pause (0.5)

iCirc(0).Shape = 1
Resize
Pause (0.5)
Color_Purple
Pause (0.5)
iCirc(0).Shape = 3
Resize
Pause (0.5)
Loop
End Sub
