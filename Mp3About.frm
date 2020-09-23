VERSION 5.00
Begin VB.Form Mp3About 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   3225
   ClientTop       =   2700
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mp3About.frx":0000
   ScaleHeight     =   3150
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks To:"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LLLArc@Home.Com"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Email:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tongue 'N Groove Copyrited 2000-2002 All Rights Reserved ©"
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
      TabIndex        =   4
      Top             =   2760
      Width           =   5295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fly.To/ArcVBPalace"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   5160
      Picture         =   "Mp3About.frx":236B
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jongmin Baek ,Andreas Åhlfeldt, Johannes.B, Stuart Pennington, For the great Visual effects."
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Made By: (Arc) Software Design"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tongue 'N Groove"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Mp3About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
MP3PlayerFrm1.Hide
MP3playerShowFav.Hide
MP3CDFrm.Hide
Timer1.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = vbBlue
Label4.ForeColor = vbBlue
End Sub

Private Sub Image1_Click()
Unload Me
MP3PlayerFrm1.Show

End Sub

Private Sub Label4_Click()
ShellExecute 0, "open", "http://fly.to/arcvbpalace", 0, 0, SW_NORMAL

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
End Sub
Sub SendMail(ByVal strAddress As String, _
Optional ByVal strCC As String, _
Optional ByVal strBCC As String, _
Optional ByVal strSubject As String, _
Optional ByVal strBodyText As String)


Dim strTemp As String

If Trim(Len(strCC)) Then
strTemp = "&CC=" & strCC
End If

If Trim(Len(strBCC)) Then
strTemp = strTemp & "&BCC=" & strBCC
End If

If Trim(Len(strSubject)) Then
strTemp = strTemp & "&Subject=" & strSubject
End If

If Trim(Len(strBodyText)) Then
strTemp = strTemp & "&Body=" & strBodyText
End If

If Len(strTemp) Then
Mid(strTemp, 1, 1) = "?"
End If

strTemp = "mailto:" & strAddress & strTemp

ShellExecute 0, "open", strTemp, 0, 0, SW_NORMAL


End Sub

Private Sub Label8_Click()
 SendMail "lllArc@Home.com", , , "Tongue 'N Groove"
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = vbRed

End Sub

Private Sub Timer1_Timer()
Label1.FontSize = Label1.FontSize + 1
If Label1.FontSize >= 22 Then
Label1.ForeColor = vbYellow
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Label1.FontSize = Label1.FontSize - 1
If Label1.FontSize <= 4 Then
Label1.ForeColor = vbBlack
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub




