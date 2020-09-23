VERSION 5.00
Begin VB.Form Mp3AddDirFrm 
   BackColor       =   &H00B31AB3&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mp3AddDirFrm.frx":0000
   ScaleHeight     =   3255
   ScaleWidth      =   1860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   240
      Pattern         =   "*.mp3*"
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
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
      Height          =   2130
      Left            =   100
      TabIndex        =   1
      Top             =   720
      Width           =   1655
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00808080&
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
      Height          =   345
      Left            =   100
      TabIndex        =   0
      Top             =   360
      Width           =   1655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Directory"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
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
      Left            =   950
      TabIndex        =   3
      Top             =   2880
      Width           =   828
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add It"
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
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   828
   End
End
Attribute VB_Name = "Mp3AddDirFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
Dim a As Long
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
For GetIt = 1 To File1.ListCount
File1.ListIndex = GetIt - 1
If FileLen(Dir1.Path & "\" & File1.Filename) < Val(Text1) Then
Call DeleteFile(Dir1.Path & "\" & File1.Filename)
a = a + 1
End If
Next GetIt
End If
MsgBox a & " Files deleted"
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
Me.Dir1.ForeColor = MP3playerShowFav.Option1.BackColor
Me.Drive1.ForeColor = MP3playerShowFav.Option1.BackColor
End Sub

Private Sub Form_Load()
Mp3AddDirFrm.Top = MP3PlayerFrm1.Top
Mp3AddDirFrm.Left = MP3PlayerFrm1.Left - Mp3AddDirFrm.Width
Me.Dir1.ForeColor = MP3playerShowFav.Option1.BackColor
Me.Drive1.ForeColor = MP3playerShowFav.Option1.BackColor
If Me.Dir1.ForeColor = &HB31AB3 Then
Me.Picture = LoadPicture(App.Path & "\mp3bg3.jpg")
End If
If Me.Dir1.ForeColor = &HA51B30 Then
Me.Picture = LoadPicture(App.Path & "\mp3bg3blue.jpg")
End If
If Me.Dir1.ForeColor = &H6A6A6A Then
Me.Picture = LoadPicture(App.Path & "\mp3bg3Silver.jpg")
End If
If Me.Dir1.ForeColor = &H63A817 Then
Me.Picture = LoadPicture(App.Path & "\mp3bg3Green.jpg")
End If
End Sub

Private Sub Label1_Click()
MP3playerShowFav.List1.Clear
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
For GetIt = 1 To File1.ListCount
File1.ListIndex = GetIt - 1
If Len(Dir1.Path) > 3 Then
MP3playerShowFav.List1.AddItem Dir1.Path & "\" & File1.Filename
Else
MP3playerShowFav.List1.AddItem Dir1.Path & File1.Filename
End If

Next GetIt
Unload Me
Else
MsgBox "No MP3 files were found in that folder folder"

End If
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call MoveForm(Me)
End Sub
