VERSION 5.00
Begin VB.Form MP3playerShowFav 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5505
   ControlBox      =   0   'False
   Icon            =   "MP3playerShowFav.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MP3playerShowFav.frx":08CA
   ScaleHeight     =   3330
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   920
      Left            =   5140
      ScaleHeight     =   915
      ScaleWidth      =   210
      TabIndex        =   12
      Top             =   2060
      Width           =   215
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Height          =   585
         Left            =   0
         TabIndex        =   15
         Top             =   165
         Width           =   210
      End
      Begin VB.Image Image5 
         Height          =   165
         Left            =   0
         Picture         =   "MP3playerShowFav.frx":93E1
         Top             =   755
         Width           =   225
      End
      Begin VB.Image Image4 
         Height          =   165
         Left            =   0
         Picture         =   "MP3playerShowFav.frx":A858
         Top             =   5
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   920
      Left            =   5160
      ScaleHeight     =   915
      ScaleWidth      =   210
      TabIndex        =   11
      Top             =   850
      Width           =   210
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Height          =   585
         Left            =   0
         TabIndex        =   14
         Top             =   165
         Width           =   225
      End
      Begin VB.Image Image3 
         Height          =   165
         Left            =   0
         Picture         =   "MP3playerShowFav.frx":BCB2
         Top             =   750
         Width           =   225
      End
      Begin VB.Image Image2 
         Height          =   165
         Left            =   0
         Picture         =   "MP3playerShowFav.frx":D129
         Top             =   5
         Width           =   225
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
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
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   330
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00B31AB3&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   245
      Left            =   3720
      TabIndex        =   7
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00B31AB3&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   245
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   255
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   960
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   5270
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5265
   End
   Begin VB.Image Image7 
      Height          =   180
      Left            =   3840
      Picture         =   "MP3playerShowFav.frx":E583
      Top             =   600
      Width           =   675
   End
   Begin VB.Image Image6 
      Height          =   180
      Left            =   1800
      Picture         =   "MP3playerShowFav.frx":FBF4
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label12 
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
      Left            =   3880
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label11 
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
      Left            =   1800
      TabIndex        =   16
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   5160
      Picture         =   "MP3playerShowFav.frx":11243
      ToolTipText     =   "Hide Play List"
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear List"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-Play The List-"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Click here or double click the play list"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play List"
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
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play List"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Song List"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add to Play List"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddAll 
         Caption         =   "Add All to Play List"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFavs 
         Caption         =   "Add to Favorites List"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete From Favorites List"
      End
      Begin VB.Menu line55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteFromDisk 
         Caption         =   "Delete From Disk"
      End
   End
   Begin VB.Menu Mpopupmenu2 
      Caption         =   "popupmenu2"
      Visible         =   0   'False
      Begin VB.Menu mClearList 
         Caption         =   "Clear List"
      End
      Begin VB.Menu line74 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClearSong 
         Caption         =   "Clear Song"
      End
   End
End
Attribute VB_Name = "MP3playerShowFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim down As Boolean




Private Sub Form_Load()

MP3playerShowFav.Width = MP3PlayerFrm1.Width + 25
MP3playerShowFav.Height = 3390
MP3playerShowFav.Top = MP3PlayerFrm1.Top + 1600
MP3playerShowFav.Left = MP3PlayerFrm1.Left
Option1.Value = True
List1.ForeColor = Option1.BackColor
List2.ForeColor = Option1.BackColor
Option3.BackColor = Option1.BackColor
Text1.ForeColor = Option1.BackColor

End Sub

Private Sub Image1_Click()
MP3playerShowFav.Hide
MP3PlayerFrm1.MediaPlayer1.Stop
MP3PlayerFrm1.Label3 = "00:00"
MP3PlayerFrm1.MediaPlayer1.CurrentPosition = 0
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Picture1.BackColor = vbYellow
Image2.BorderStyle = 1
high = 585 / List1.ListCount
Scroll = 0
Dim lR As Long
Dim lLineCount As Long
down = True
lLineCount = 1
Do
DoEvents
Label9.Height = Label9.Height + high
If Label9.Height > 585 Then Label9.Height = 585
lR = SendMessage(List1.hwnd, WM_VSCROLL, SB_LINEUP, lLineCount)
Loop Until down = False
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image2.BorderStyle = 0
Picture1.BackColor = vbBlack
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Picture1.BackColor = vbYellow
high = 585 / List1.ListCount
Image3.BorderStyle = 1
Dim lR As Long
Dim lLineCount As Long
down = True
lLineCount = 1
Do
DoEvents
Label9.Height = Label9.Height - high
If Label9.Height < 15 Then Label9.Height = 15
lR = SendMessage(List1.hwnd, WM_VSCROLL, SB_LINEDOWN, lLineCount)
Loop Until down = False
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image3.BorderStyle = 0
Picture1.BackColor = vbBlack
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Picture2.BackColor = vbYellow
Image4.BorderStyle = 1
high = 585 / List2.ListCount
Dim lR As Long
Dim lLineCount As Long
lLineCount = 1
down = True
Do
DoEvents
Label10.Height = Label10.Height + high
If Label10.Height > 585 Then Label10.Height = 585
lR = SendMessage(List2.hwnd, WM_VSCROLL, SB_LINEUP, lLineCount)
Loop Until down = False
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image4.BorderStyle = 0
Picture2.BackColor = vbBlack
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Picture2.BackColor = vbYellow
high = 585 / List2.ListCount
Image5.BorderStyle = 1
Dim lR As Long
Dim lLineCount As Long
down = True
lLineCount = 1
Do
DoEvents
Label10.Height = Label10.Height - high
If Label10.Height < 15 Then Label10.Height = 15
lR = SendMessage(List2.hwnd, WM_VSCROLL, SB_LINEDOWN, lLineCount)
Loop Until down = False
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image5.BorderStyle = 0
Picture2.BackColor = vbBlack
End Sub

Private Sub Image6_Click()
Label9.Height = 585

mnuAddFavs.Visible = True
mnuDelete.Visible = False
Mp3AddDirFrm.Show vbModal
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.BorderStyle = 1
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.BorderStyle = 0
End Sub

Private Sub Image7_Click()
Label9.Height = 585

mnuAddFavs.Visible = False
mnuDelete.Visible = True
List1.Clear
a = App.Path & "\mp3favorites.dat"
Open a For Binary As #1
Dim sInp As String
sInp = String(LOF(1), 0)
Get #1, 1, sInp
Close #1
arry = Split(sInp, vbCrLf)
For i = LBound(arry) To UBound(arry) - 1
List1.AddItem (arry(i))
Next
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image7.BorderStyle = 1
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image7.BorderStyle = 0
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.ForeColor = vbBlack
Call MoveForm(Me)
Label3.ForeColor = vbYellow
End Sub





Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.ForeColor = vbYellow
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.ForeColor = vbBlack
End Sub





Private Sub Label6_Click()
On Error Resume Next
If List2.ListCount = 0 Then
MsgBox "You haven't selected any songs to play"
Exit Sub
End If
List2.ListIndex = 0
If Option3.Value = True Then
MP3playerShowFav.List2.ListIndex = Int(MP3playerShowFav.List2.ListCount * Rnd)
End If

MP3PlayerFrm1.Text1 = List2.List(List2.ListIndex)
MP3PlayerFrm1.Label1 = List2.List(List2.ListIndex)
 FindStr = InStrRev(MP3PlayerFrm1.Label1, "\")
  getstr = Mid(MP3PlayerFrm1.Label1, FindStr + 1, Len(MP3PlayerFrm1.Label1))
 If InStr(getstr, ".MP3") Then
RepStr = Replace(getstr, ".MP3", "")
GoTo Line1
End If
If InStr(getstr, ".Mp3") Then
RepStr = Replace(getstr, ".Mp3", "")
GoTo Line1
End If
If InStr(getstr, ".mP3") Then
RepStr = Replace(getstr, ".mP3", "")
GoTo Line1
End If
RepStr = Replace(getstr, ".mp3", "")
Line1:
MP3PlayerFrm1.Label1 = RepStr

MP3PlayerFrm1.MediaPlayer1.Filename = List2.List(List2.ListIndex)
 iTheTime = CInt(MP3PlayerFrm1.MediaPlayer1.Duration)
  iTheSeconds = iTheTime Mod 60
   iTheMinutes = iTheTime \ 60
  MP3PlayerFrm1.Label3.Caption = Format(iTheMinutes, "00") & ":" & Format(iTheSeconds, "00")
 MP3PlayerFrm1.Slider1.Max = MP3PlayerFrm1.MediaPlayer1.Duration
MP3PlayerFrm1.Timer4.Enabled = True

End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.ForeColor = vbBlack
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.ForeColor = vbYellow
End Sub

Private Sub Label7_Click()
Call PopupMenu(Mpopupmenu2)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.ForeColor = vbBlack
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.ForeColor = vbYellow
End Sub



Private Sub Label8_Click()
If Text1.Visible = True Then Text1.Visible = False: Exit Sub
If Text1.Visible = False Then Text1.Visible = True: Exit Sub

End Sub

Private Sub mDelete_Click()
List2.RemoveItem (List2.ListIndex)
End Sub



Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label8.ForeColor = vbBlack
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label8.ForeColor = vbYellow
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then Exit Sub
Call PopupMenu(mnuSelect)
End Sub

Private Sub List2_DblClick()
On Error Resume Next
Text1.Visible = False
MP3PlayerFrm1.Text1 = List2.List(List2.ListIndex)
MP3PlayerFrm1.MediaPlayer1.Filename = List2.List(List2.ListIndex)
MP3PlayerFrm1.Label1 = List2.List(List2.ListIndex)
 FindStr = InStrRev(MP3PlayerFrm1.Label1, "\")
  getstr = Mid(MP3PlayerFrm1.Label1, FindStr + 1, Len(MP3PlayerFrm1.Label1))
 If InStr(getstr, ".MP3") Then
RepStr = Replace(getstr, ".MP3", "")
GoTo Line1
End If
If InStr(getstr, ".Mp3") Then
RepStr = Replace(getstr, ".Mp3", "")
GoTo Line2
End If
If InStr(getstr, ".mP3") Then
RepStr = Replace(getstr, ".mP3", "")
GoTo line3
End If
RepStr = Replace(getstr, ".mp3", "")
Line1:
Line2:
line3:
MP3PlayerFrm1.Label1 = RepStr
iTheTime = CInt(MP3PlayerFrm1.MediaPlayer1.Duration)
  iTheSeconds = iTheTime Mod 60
   iTheMinutes = iTheTime \ 60
  MP3PlayerFrm1.Label3.Caption = Format(iTheMinutes, "00") & ":" & Format(iTheSeconds, "00")
 MP3PlayerFrm1.Slider1.Max = MP3PlayerFrm1.MediaPlayer1.Duration
MP3PlayerFrm1.Timer4.Enabled = True
End Sub

Private Sub mClearList_Click()
List2.Clear
End Sub

Private Sub mnuAdd_Click()
For i = 0 To List1.ListCount - 1
If List1.List(List1.ListIndex) = "" Then MsgBox ("Please select a song to add"): Exit Sub
If List1.Selected(i) = True Then Text2 = List1.List(i)
Next i
List2.AddItem (Text2)
List1.ListIndex = -1
End Sub

Private Sub mnuAddAll_Click()
List2.Clear
For i = 0 To List1.ListCount - 1
List2.AddItem List1.List(i)
Next i
List1.ListIndex = -1
End Sub

Private Sub mnuAddFavs_Click()
If List1.ListIndex = -1 Then
MsgBox "Please Select a Song to Add"
Exit Sub
End If
a = App.Path + "\Mp3Favorites.dat"
Open a For Append As 1
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
Print #1, List1.List(i)
End If
Next i
Close
List1.ListIndex = -1
End Sub

Private Sub MnuClearSong_Click()
If List2.ListIndex = -1 Then
MsgBox "Please Select a Song to Remove"
Else
List2.RemoveItem (List2.ListIndex)
End If
End Sub

Private Sub mnuDelete_Click()
a = MsgBox("Are you Sure you want to delete this file from your Favorites? This will not Delete the file form your hard drive.", 36, "     Tongue 'N Groove")
If a = 7 Then Exit Sub
If a = 6 Then GoTo line7
line7:
List1.RemoveItem (List1.ListIndex)
List1.Refresh
a = App.Path + "\mp3favorites.dat"
Open a For Output As 1
For i = 0 To List1.ListCount - 1
Print #1, List1.List(i)
Next i
Close
List1.ListIndex = -1
End Sub

Private Sub mnuDeleteFromDisk_Click()
a = MsgBox("Are you Sure you want to delete this file from your hard drive? This will not remove the refrence to the file from your Favorites list.", 36, "     Tongue 'N Groove")
If a = 7 Then Exit Sub
If a = 6 Then GoTo line7
line7:
Call DeleteFile(List1.List(List1.ListIndex))
List1.RemoveItem (List1.ListIndex)
List1.Refresh
List2.Clear
For i = 0 To List1.ListCount - 1
List2.AddItem List1.List(i)
Next i
End Sub

Private Sub Text1_Change()
For i = 0 To MP3playerShowFav.List1.ListCount - 1
If InStr(LCase(MP3playerShowFav.List1.List(i)), LCase(Text1.Text)) Then MP3playerShowFav.List1.Selected(i) = True: Exit For
Next i
For i = 0 To MP3playerShowFav.List2.ListCount - 1
If InStr(LCase(MP3playerShowFav.List2.List(i)), LCase(Text1.Text)) Then MP3playerShowFav.List2.Selected(i) = True: Exit For
Next i
End Sub

Private Sub Text1_DblClick()
Text1.Visible = False
End Sub
