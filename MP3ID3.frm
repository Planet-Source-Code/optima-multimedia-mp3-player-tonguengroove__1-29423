VERSION 5.00
Begin VB.Form Mp3ID3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MP3ID3.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Text            =   "Text7"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00B31AB3&
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B31AB3&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   3360
      Picture         =   "MP3ID3.frx":6DBA
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   960
      TabIndex        =   12
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
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
      Left            =   960
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
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
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ID3 Information"
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   40
      Width           =   3855
   End
End
Attribute VB_Name = "Mp3ID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
On Error Resume Next

If MP3playerShowFav.Option1.BackColor = &HB31AB3 Then
Me.Picture = LoadPicture(App.Path & "\images\mp3id3.jpg")
Text1.ForeColor = &HB31AB3
Text2.ForeColor = &HB31AB3
Text3.ForeColor = &HB31AB3
Text4.ForeColor = &HB31AB3
Text5.ForeColor = &HB31AB3
Text7.ForeColor = &HB31AB3
End If
If MP3playerShowFav.Option1.BackColor = &HA51B30 Then
Me.Picture = LoadPicture(App.Path & "\images\mp3id3blue.jpg")
Text1.ForeColor = &HA51B30
Text2.ForeColor = &HA51B30
Text3.ForeColor = &HA51B30
Text4.ForeColor = &HA51B30
Text5.ForeColor = &HA51B30
Text7.ForeColor = &HA51B30
End If
If MP3playerShowFav.Option1.BackColor = &H6A6A6A Then
Me.Picture = LoadPicture(App.Path & "\images\mp3id3Silver.jpg")
Text1.ForeColor = &H6A6A6A
Text2.ForeColor = &H6A6A6A
Text3.ForeColor = &H6A6A6A
Text4.ForeColor = &H6A6A6A
Text5.ForeColor = &H6A6A6A
Text7.ForeColor = &H6A6A6A
End If
If MP3playerShowFav.Option1.BackColor = &H63A817 Then
Me.Picture = LoadPicture(App.Path & "\images\mp3id3Green.jpg")
Text1.ForeColor = &H63A817
Text2.ForeColor = &H63A817
Text3.ForeColor = &H63A817
Text4.ForeColor = &H63A817
Text5.ForeColor = &H63A817
Text7.ForeColor = &H63A817
End If

GenreArray = Split(sGenreMatrix, "|")   ' we fill the array with the Genre's
For i = LBound(GenreArray) To UBound(GenreArray)
Combo1.AddItem GenreArray(i)        ' now fill the Combobox with the array, and voila, the code you
                                    ' you recieve form the Genre part of the Type, represents the combobox Listindex =)
Next
Dim Position(0 To 147) As Long
Dim Start As Long
Start = 1
Position(0) = 1
On Error Resume Next        ' it creates an error once sGenreMatrix runs out of "|"
For i = 1 To 147             ' number of Genre's in sGenreMatrix
pt = InStr(Start, sGenreMatrix, "|", Position(i) = pt + 1)
Start = pt + 1
Next
For i = 0 To 147
X = (Mid$(sGenreMatrix, Position(i), Position(i + 1) - Position(i) - 1))
Combo1.AddItem X

Next

  GetId3 MP3PlayerFrm1.MediaPlayer1.Filename       ' Get the filename
Text1 = RTrim(id3Info.Title)            ' since the fields in the type are
Text2 = RTrim(id3Info.Artist)                  ' fixed lenght, we use Rtrim to cut the
Text3 = RTrim(id3Info.Album)                   ' trailing bytes
Text4 = RTrim(id3Info.sYear)
Text5 = RTrim(id3Info.Comments)
Text6 = RTrim(id3Info.Genre)
Combo1.ListIndex = id3Info.Genre
Text7 = Combo1.List(Combo1.ListIndex)
Text7.ForeColor = Text1.ForeColor
End Sub

Private Sub Form_Load()
Me.Top = MP3PlayerFrm1.Top + 1200
Me.Left = MP3PlayerFrm1.Left + 900
End Sub

Private Sub Image1_Click()

Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlack
Call MoveForm(Me)
Label1.ForeColor = vbYellow
End Sub


