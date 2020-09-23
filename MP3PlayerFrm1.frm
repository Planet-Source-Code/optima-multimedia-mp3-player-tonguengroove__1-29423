VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MP3PlayerFrm1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tongue 'N Groove"
   ClientHeight    =   4035
   ClientLeft      =   5565
   ClientTop       =   2475
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Timmons"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MP3PlayerFrm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MP3PlayerFrm1.frx":08CA
   ScaleHeight     =   4035
   ScaleWidth      =   5475
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   200
      Left            =   3960
      ScaleHeight     =   135
      ScaleWidth      =   1275
      TabIndex        =   15
      Top             =   760
      Width           =   1335
      Begin VB.Label Label9 
         BackColor       =   &H0000FFFF&
         Height          =   135
         Left            =   650
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Image Image12 
         Height          =   135
         Left            =   1120
         Picture         =   "MP3PlayerFrm1.frx":5251
         Top             =   0
         Width           =   135
      End
      Begin VB.Image Image11 
         Height          =   135
         Left            =   0
         Picture         =   "MP3PlayerFrm1.frx":661A
         Top             =   0
         Width           =   135
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   135
      Left            =   1680
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   238
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   200
      Left            =   1560
      ScaleHeight     =   135
      ScaleWidth      =   2115
      TabIndex        =   12
      Top             =   390
      Width           =   2175
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   0
         Picture         =   "MP3PlayerFrm1.frx":79D0
         ScaleHeight     =   135
         ScaleWidth      =   420
         TabIndex        =   13
         Top             =   0
         Width           =   425
      End
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   120
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Balance"
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   212
      _Version        =   393216
      BorderStyle     =   1
      Min             =   -5000
      Max             =   5000
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   120
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Volume"
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   212
      _Version        =   393216
      BorderStyle     =   1
      TickStyle       =   3
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2040
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   5525
      _ExtentX        =   9737
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      RightMargin     =   5
      TextRTF         =   $"MP3PlayerFrm1.frx":8E68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Timmons"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   120
      Top             =   2040
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Tongue 'N Groove MP3 Player"
      Top             =   3000
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog cdbopenfolder 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   150
      Left            =   4125
      TabIndex        =   18
      Top             =   415
      Width           =   1005
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image10 
      Height          =   135
      Left            =   5125
      Picture         =   "MP3PlayerFrm1.frx":8F37
      Top             =   420
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Left            =   3990
      Picture         =   "MP3PlayerFrm1.frx":A300
      Top             =   420
      Width           =   135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   3960
      TabIndex        =   16
      Top             =   390
      Width           =   1335
   End
   Begin VB.Image Image8 
      Height          =   180
      Left            =   120
      Picture         =   "MP3PlayerFrm1.frx":B6B6
      ToolTipText     =   "Minimize Tongue 'N Groove"
      Top             =   900
      Width           =   180
   End
   Begin VB.Image Image7 
      Height          =   180
      Left            =   5160
      Picture         =   "MP3PlayerFrm1.frx":CB08
      ToolTipText     =   "Exit Tongue 'N Groove"
      Top             =   120
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   135
      Left            =   120
      Picture         =   "MP3PlayerFrm1.frx":DFB5
      Top             =   120
      Width           =   420
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   3180
      Picture         =   "MP3PlayerFrm1.frx":F372
      ToolTipText     =   "FastFoward"
      Top             =   600
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   1800
      Picture         =   "MP3PlayerFrm1.frx":10A4F
      ToolTipText     =   "Rewind"
      Top             =   600
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   2830
      Picture         =   "MP3PlayerFrm1.frx":1214A
      ToolTipText     =   "Pause"
      Top             =   600
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   2470
      Picture         =   "MP3PlayerFrm1.frx":13836
      ToolTipText     =   "Stop"
      Top             =   600
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   2130
      Picture         =   "MP3PlayerFrm1.frx":14EDA
      ToolTipText     =   "Play"
      Top             =   600
      Width           =   345
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   3960
      TabIndex        =   11
      Top             =   580
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Timmons"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   3960
      TabIndex        =   10
      Top             =   210
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tongue 'N Groove"
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
      Left            =   20
      TabIndex        =   7
      Top             =   45
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Length of Song"
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Time Played"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   4095
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   2280
      Width           =   615
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   1
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnu1 
      Caption         =   "Stuff"
      Visible         =   0   'False
      Begin VB.Menu mnuSongs 
         Caption         =   "Songs"
         Begin VB.Menu mnuFileSelect 
            Caption         =   "Select a Song"
         End
         Begin VB.Menu mnuline 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileSaveSongs 
            Caption         =   "Add Current Song To Favorites"
         End
         Begin VB.Menu Mnuline2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileShowFav 
            Caption         =   "Show Song  List"
         End
         Begin VB.Menu line99 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCD 
            Caption         =   "Audio CD "
         End
      End
      Begin VB.Menu mnuVideo 
         Caption         =   "Video"
         Begin VB.Menu mnuFireWorks 
            Caption         =   "Fire Works"
         End
         Begin VB.Menu mnuCircles 
            Caption         =   "Circles"
         End
         Begin VB.Menu mnuPipes 
            Caption         =   "Pipes"
         End
         Begin VB.Menu mnuTwister 
            Caption         =   "Twister"
         End
         Begin VB.Menu mnuMatrix 
            Caption         =   "Matrix"
         End
      End
      Begin VB.Menu mnuSkins 
         Caption         =   "Skins"
         Begin VB.Menu mnuPurple 
            Caption         =   "Purple"
         End
         Begin VB.Menu MNuBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuGrey 
            Caption         =   "Grey"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "Green"
         End
      End
      Begin VB.Menu mnuline34 
         Caption         =   "-"
      End
      Begin VB.Menu mnuID3 
         Caption         =   "ID3 Info"
      End
      Begin VB.Menu Line456 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuline56 
         Caption         =   "-"
      End
      Begin VB.Menu mnOnTop 
         Caption         =   "OnTop"
      End
   End
End
Attribute VB_Name = "MP3PlayerFrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim hmixer As Long             ' mixer handle
Dim WavCtrl As MIXERCONTROL    ' wave output volume control
Dim CDVol As MIXERCONTROL      ' CD Volume
Dim LineVol As MIXERCONTROL    ' Line/In Volume
Dim MBOOST As MIXERCONTROL     ' Microphone Volume
Dim PSPKVol As MIXERCONTROL    ' PcSpeaker Volume
Dim CurVol As Long
Dim AUXVol As MIXERCONTROL     ' Auxillary Volume
Dim TADVol As MIXERCONTROL     ' TAD-In Volume

Dim MIDIVol As MIXERCONTROL    ' Midi Volume

Dim I25InVol As MIXERCONTROL   ' I25In Volume
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL

Dim rc As Long                 ' return code
Dim ok As Boolean
Dim NowBeingUsed(100) As Boolean    'Do not touch it!
Dim Direction(100) As Integer       'the direction of the each flame
Dim Speed(100) As Integer           'the speed of the each flame
Dim y(100) As Integer               'the location of the each flame
Dim x(100) As Integer               '               "
Dim COLORRGB(100, 2) As Long        'the color of the each flame(RGB)
Dim VolCtrl As MIXERCONTROL                                         'ColorRGB(n, 0) -> Red value
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
Dim TotalTrackPlay As String    'For Display.                      'ColorRGB(n, 2) -> Blue value
Public SkinLoc As String
Public SkinLocRev As String
Public SkinLocPlay As String
Public SkinLocStop As String
Public SkinLocPause As String
Public SkinLocFFW As String
Public SkinLocShow As String
Public SkinLocDir As String
Public SkinLocID3 As String
Dim Scatter As Integer              'It decides how far the flames are
                                    'scattered when they launched
Dim ScatterFalling As Integer       'It decides how far the flames are
                                    'scattered when they start falling.
Dim SpeedDif As Integer             'The larger it is, the distance between
                                    'the higher flame and the lower flame
Dim Per As Integer                                 'become bigger.
Dim Size As Integer
Dim down As Boolean







Private Sub Command2_Click()
  rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        
        Exit Sub
    End If

    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)
    If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, VolCtrl)
      
    End If
   
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, WavCtrl)

    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MBOOST, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, MBOOST)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MBOOST)
  
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, CDVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, CDVol)
   
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, AUXVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, AUXVol)
   
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TADVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, TADVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, TADVol)

    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, MIDIVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MIDIVol)
   
    End If

        ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PSPKVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, PSPKVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, PSPKVol)
     
    End If

    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, I25InVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, I25InVol)
     
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, _
                                  MIXERCONTROL_CONTROLTYPE_VOLUME, LineVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, LineVol)
      
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_BASS, Bass)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Bass)
       
    End If
    
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                                  MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Treble)
       
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
MP3playerShowFav.Show
Send "close all"
Send "open all"
Call TaskBarIcon(Me)
ShowTitleBar False
CurVol = 65500
' Loading Pictures~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Me.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "FrmPicture"))
Mp3ID3.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "ID3Picture"))
Image4.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "RevPicture"))
Image1.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "PlayPicture"))
Image2.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "StopPicture"))
Image3.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "PausePicture"))
Image5.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "FFWPicture"))
MP3playerShowFav.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "ShowPicture"))
MP3playerShowFav.Option1.BackColor = GetSetting("ToungeNGroove", "Exit", "ShowFrm")
MP3playerShowFav.Option3.BackColor = GetSetting("ToungeNGroove", "Exit", "ShowFrmOption3")
MP3playerShowFav.List1.ForeColor = GetSetting("ToungeNGroove", "Exit", "ShowFrmList1")
MP3playerShowFav.List2.ForeColor = GetSetting("ToungeNGroove", "Exit", "ShowFrmList2")
MP3playerShowFav.Text1.ForeColor = GetSetting("ToungeNGroove", "Exit", "ShowFrmText1")
MP3CDFrm.Label4.ForeColor = GetSetting("ToungeNGroove", "Exit", "CdFrm")
MP3CDFrm.Combo1.ForeColor = GetSetting("ToungeNGroove", "Exit", "CdFrmCombo1")
MP3CDFrm.Label5.ForeColor = GetSetting("ToungeNGroove", "Exit", "CdFrmLabel5")
Mp3AddDirFrm.Picture = LoadPicture(GetSetting("ToungeNGroove", "Exit", "Dirfrm"))
' End Load~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If (App.PrevInstance = True) Then
End
End If
Label6.FontSize = 6
Slider2.Value = 10
MP3PlayerFrm1.Height = 1620
MP3PlayerFrm1.Width = 5595
Scatter = 5
ScatterFalling = 60
SpeedDif = 30
Size = 30
Label1.Caption = "             ^v^  Tongue 'N Groove MP3/CD Player  ^v^  "
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
Dim TotalTrackPlay As String
Command2_Click
volume = 65500
SetVolumeControl hmixer, VolCtrl, volume
End Sub


Private Sub Image1_Click()
On Error Resume Next
If MP3CDFrm.Visible = True Then
Send "play cd from " & Val(MP3CDFrm.Combo1.Text)
Else
MediaPlayer1.Play
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.BorderStyle = 0
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image10.BorderStyle = 1
down = True
Do
DoEvents
Label7.Width = Label7.Width + 100
If Label7.Width > 1000 Then Label7.Width = 1000
Pause (0.3)
Slider2.Value = Slider2.Value + 1
Loop Until down = False

End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image10.BorderStyle = 0
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image11.BorderStyle = 1
Label9.Visible = True
down = True
Do
DoEvents
Label9.Left = Label9.Left - 8
If Label9.Left < 180 Then Label9.Left = 180
Pause (0.00001)
Slider3.Value = Slider3.Value - 80
Loop Until down = False
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image11.BorderStyle = 0
Label9.Visible = False
End Sub

Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image12.BorderStyle = 1
Label9.Visible = True
down = True
Do
DoEvents
Label9.Left = Label9.Left + 8
If Label9.Left > 1100 Then Label9.Left = 1100
Pause (0.00001)
Slider3.Value = Slider3.Value + 80
Loop Until down = False
End Sub

Private Sub Image12_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image12.BorderStyle = 0
Label9.Visible = False
End Sub

Private Sub Image2_Click()

If MP3CDFrm.Visible = True Then
Send "stop cd wait"
Send "seek cd to " & MP3CDFrm.Combo1.Text
Else
MediaPlayer1.Stop
MediaPlayer1.CurrentPosition = 0
End If

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image2.BorderStyle = 0
End Sub

Private Sub Image3_Click()
On Error Resume Next
If MP3CDFrm.Visible = True Then
Send "stop cd wait"
Else
MediaPlayer1.Pause
End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image3.BorderStyle = 0
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
volume = 0
SetVolumeControl hmixer, VolCtrl, volume
Image4.BorderStyle = 1
If MP3CDFrm.Visible = True Then
down = True
Do
DoEvents

Send "stop cd wait"
  FFSpeed = -1
  Dim s As String * 40
            Send "set cd time format milliSeconds"
            mciSendString "status cd position wait", s, Len(s), 0
            Cmd = "play cd from " & CStr(CLng(s) + FFSpeed * 3000)
            mciSendString Cmd, 0, 0, 0
            Send "set cd time format tmsf"

            
Loop Until down = False
Else
down = True
Do
DoEvents
Slider1.Value = Slider1.Value - 1
MediaPlayer1.CurrentPosition = Slider1.Value
Loop Until down = False
End If

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image4.BorderStyle = 0
SetVolumeControl hmixer, VolCtrl, CurVol
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
volume = 0
SetVolumeControl hmixer, VolCtrl, volume
Dim e As String * 40
Image5.BorderStyle = 1
If MP3CDFrm.Visible = True Then
down = True
Send "stop cd wait"
Do
DoEvents
Send "stop cd wait"
  FFSpeed = 1
  Dim s As String * 40
            Send "set cd time format milliSeconds"
            mciSendString "status cd position wait", s, Len(s), 0
            Cmd = "play cd from " & CStr(CLng(s) + FFSpeed * 3000)
            mciSendString Cmd, 0, 0, 0
            Send "set cd time format tmsf"
Loop Until down = False
Else
down = True
Do
DoEvents
Slider1.Value = Slider1.Value + 1
MediaPlayer1.CurrentPosition = Slider1.Value
Loop Until down = False
End If

End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image5.BorderStyle = 0
SetVolumeControl hmixer, VolCtrl, CurVol
End Sub

Private Sub Label10_Click()
Call PopupMenu(mnuVideo)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.ForeColor = vbYellow
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.ForeColor = vbBlack
End Sub

Private Sub Image6_Click()
Call PopupMenu(mnu1)
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.BorderStyle = 1
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image6.BorderStyle = 0
End Sub

Private Sub Image7_Click()
volume = 65500
SetVolumeControl hmixer, VolCtrl, volume
Send "stop cd wait"
Send "Close all"
If SkinLoc = "" Then GoTo En
SaveSetting "ToungeNGroove", "Exit", "ID3Picture", SkinLocID3
SaveSetting "ToungeNGroove", "Exit", "FrmPicture", SkinLoc
SaveSetting "ToungeNGroove", "Exit", "RevPicture", SkinLocRev
SaveSetting "ToungeNGroove", "Exit", "PlayPicture", SkinLocPlay
SaveSetting "ToungeNGroove", "Exit", "StopPicture", SkinLocStop
SaveSetting "ToungeNGroove", "Exit", "PausePicture", SkinLocPause
SaveSetting "ToungeNGroove", "Exit", "FFWPicture", SkinLocFFW
SaveSetting "ToungeNGroove", "Exit", "CdFrm", MP3CDFrm.Label4.ForeColor
SaveSetting "ToungeNGroove", "Exit", "CdFrmCombo1", MP3CDFrm.Combo1.ForeColor
SaveSetting "ToungeNGroove", "Exit", "CdFrmLabel5", MP3CDFrm.Label5.ForeColor
SaveSetting "ToungeNGroove", "Exit", "ShowPicture", SkinLocShow
SaveSetting "ToungeNGroove", "Exit", "ShowFrm", MP3playerShowFav.Option1.BackColor
SaveSetting "ToungeNGroove", "Exit", "ShowFrmOption3", MP3playerShowFav.Option3.BackColor
SaveSetting "ToungeNGroove", "Exit", "ShowFrmList1", MP3playerShowFav.List1.ForeColor
SaveSetting "ToungeNGroove", "Exit", "ShowFrmList2", MP3playerShowFav.List2.ForeColor
SaveSetting "ToungeNGroove", "Exit", "ShowFrmText1", MP3playerShowFav.Text1.ForeColor
SaveSetting "ToungeNGroove", "Exit", "DirFrm", SkinLocDir
En:
End
End Sub



Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image7.BorderStyle = 1
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Image7.BorderStyle = 0
End Sub

Private Sub Image8_Click()
MP3PlayerFrm1.WindowState = vbMinimized
MP3playerShowFav.Hide
MP3CDFrm.Hide
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Image9.BorderStyle = 1
down = True
Do
DoEvents
Label7.Width = Label7.Width - 100
Pause (0.3)
Slider2.Value = Slider2.Value - 1
Loop Until down = False

End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
Image9.BorderStyle = 0

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.ForeColor = vbBlack
Call MoveForm(Me)
Label4.ForeColor = vbYellow
MP3playerShowFav.Top = MP3PlayerFrm1.Top + 1600
MP3playerShowFav.Left = MP3PlayerFrm1.Left
Mp3AddDirFrm.Top = MP3PlayerFrm1.Top
Mp3AddDirFrm.Left = MP3PlayerFrm1.Left - Mp3AddDirFrm.Width
MP3CDFrm.Top = MP3PlayerFrm1.Top + 1600
MP3CDFrm.Left = MP3PlayerFrm1.Left
End Sub




Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
down = True
Do
DoEvents
Slider1.Value = Slider1.Value - 1
MediaPlayer1.CurrentPosition = Slider1.Value
Loop Until down = False
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
down = True
Do
DoEvents
Slider1.Value = Slider1.Value + 1
MediaPlayer1.CurrentPosition = Slider1.Value
Loop Until down = False
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
down = False
End Sub
Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label9.ForeColor = vbBlack
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label9.ForeColor = vbYellow
End Sub
Private Sub Label8_Click()
Slider3.Value = 0
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)

On Error GoTo err
If MP3playerShowFav.List2.ListIndex + 1 = MP3playerShowFav.List2.ListCount Then
MP3playerShowFav.List2.ListIndex = -1
MediaPlayer1.Filename = MP3playerShowFav.List2.List(MP3playerShowFav.List2.ListIndex)
End If
If MP3playerShowFav.Option1.Value = True Then
MP3playerShowFav.List2.ListIndex = MP3playerShowFav.List2.ListIndex + 1
MP3PlayerFrm1.Label1 = MP3playerShowFav.List2.List(MP3playerShowFav.List2.ListIndex)
End If

If MP3playerShowFav.Option3.Value = True Then
MP3playerShowFav.List2.ListIndex = Int(MP3playerShowFav.List2.ListCount * Rnd)
MP3PlayerFrm1.MediaPlayer1.Filename = MP3playerShowFav.List2.List(MP3playerShowFav.List2.ListIndex)
MP3PlayerFrm1.Label1 = MP3playerShowFav.List2.List(MP3playerShowFav.List2.ListIndex)
End If
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
MediaPlayer1.Filename = MP3playerShowFav.List2.List(MP3playerShowFav.List2.ListIndex)
 iTheTime = CInt(MP3PlayerFrm1.MediaPlayer1.Duration)
  iTheSeconds = iTheTime Mod 60
   iTheMinutes = iTheTime \ 60
  MP3PlayerFrm1.Label3.Caption = Format(iTheMinutes, "00") & ":" & Format(iTheSeconds, "00")
 MP3PlayerFrm1.Slider1.Max = MP3PlayerFrm1.MediaPlayer1.Duration
MP3PlayerFrm1.Timer4.Enabled = True
err:
End Sub

Private Sub mnOnTop_Click()
If mnOnTop.Checked = False Then
Call Stayontop(Me)
Call Stayontop(MP3playerShowFav)

Call Stayontop(Mp3AddDirFrm)
mnOnTop.Checked = True
Else
If mnOnTop.Checked = True Then
Call StayOnBottom(Me)
Call StayOnBottom(MP3playerShowFav)

Call StayOnBottom(Mp3AddDirFrm)
mnOnTop.Checked = False
End If
End If
End Sub

Private Sub mnuAbout_Click()
Mp3About.Show
End Sub

Private Sub MNuBlue_Click()
Mp3ID3.Picture = LoadPicture(App.Path & "\mp3id3blue.jpg")
SkinLocDir = App.Path & "\mp3bg3Blue.jpg"
SkinLoc = App.Path & "\mp3bgBlue.jpg"
SkinLocRev = App.Path & "\mp3Revblue.jpg"
SkinLocPlay = App.Path & "\mp3Playblue.jpg"
SkinLocStop = App.Path & "\mp3Stopblue.jpg"
SkinLocPause = App.Path & "\mp3Pauseblue.jpg"
SkinLocFFW = App.Path & "\mp3FFWblue.jpg"
SkinLocShow = App.Path & "\mp3BG2Blue.jpg"
SkinLocID3 = App.Path & "\mp3ID3Blue.jpg"
MP3PlayerFrm1.Picture = LoadPicture(App.Path & "\mp3bgblue.jpg")
MP3playerShowFav.Picture = LoadPicture(App.Path & "\mp3bg2blue.jpg")
MP3CDFrm.Picture = LoadPicture(App.Path & "\mp3bgblue.jpg")
Mp3AddDirFrm.Picture = LoadPicture(App.Path & "\mp3bg3blue.jpg")
Image4.Picture = LoadPicture(App.Path & "\mp3Revblue.jpg")
Image1.Picture = LoadPicture(App.Path & "\mp3Playblue.jpg")
Image2.Picture = LoadPicture(App.Path & "\mp3Stopblue.jpg")
Image3.Picture = LoadPicture(App.Path & "\mp3Pauseblue.jpg")
Image5.Picture = LoadPicture(App.Path & "\mp3FFWblue.jpg")
MP3playerShowFav.Option1.BackColor = &HA51B30
MP3playerShowFav.Option3.BackColor = &HA51B30
MP3playerShowFav.List1.ForeColor = &HA51B30
MP3playerShowFav.List2.ForeColor = &HA51B30
MP3playerShowFav.Text1.ForeColor = &HA51B30
Mp3AddDirFrm.Dir1.ForeColor = &HA51B30
Mp3AddDirFrm.Drive1.ForeColor = &HA51B30
MP3CDFrm.Combo1.ForeColor = &HA51B30
MP3CDFrm.Label4.ForeColor = &HA51B30
MP3CDFrm.Label5.ForeColor = &HA51B30
MP3CDFrm.Option1.BackColor = &HA51B30
MP3CDFrm.Option2.BackColor = &HA51B30
End Sub



Private Sub mnuCD_Click()
On Error Resume Next
MP3PlayerFrm1.Label3 = "00:00"
MP3PlayerFrm1.Label2 = "00:00"
MP3CDFrm.Show
MP3playerShowFav.Hide
MP3PlayerFrm1.MediaPlayer1.Stop
MP3PlayerFrm1.MediaPlayer1.CurrentPosition = 0
End Sub

Private Sub mnuCircles_Click()
Mp3Circles.Show
End Sub

Private Sub mnuFileSaveSongs_Click()
On Error Resume Next
If Text1 = "Tongue 'N Groove MP3 Player" Then
a = MsgBox("Please select a song to save"): Exit Sub
End If
a = App.Path + "\Mp3Favorites.dat"
Open a For Append As 1
Print #1, Text1
Close
End Sub

Private Sub mnuFileSelect_Click()
Dim iTheTime As Integer, iTheMinutes As Integer, iTheSeconds As Integer
On Error Resume Next

MP3CDFrm.Visible = False

cdbopenfolder.Filter = "Music(*.mp3;*.wav)|*.mp3;*.wav"
cdbopenfolder.Action = 1
If cdbopenfolder.Filename = "" Then Exit Sub
Text1 = cdbopenfolder.Filename
Label1 = cdbopenfolder.Filename
MediaPlayer1.Filename = Text1
'gets rid of the file extension and.mp3 in the file name
FindStr = InStrRev(Label1, "\")
getstr = Mid(Label1, FindStr + 1, Len(Text1))
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
Label1 = RepStr
'converts label3's caption from seconds to minutes
iTheTime = CInt(MediaPlayer1.Duration)
iTheSeconds = iTheTime Mod 60
iTheMinutes = iTheTime \ 60
Label3.Caption = Format(iTheMinutes, "00") & ":" & Format(iTheSeconds, "00")
Slider1.Max = MediaPlayer1.Duration
Timer4.Enabled = True
End Sub

Private Sub mnuFileShowFav_Click()
Send "stop cd wait"
Timer4.Enabled = False
MP3CDFrm.Timer1.Enabled = False
MP3CDFrm.Hide
MP3PlayerFrm1.Label2 = "00:00"
MP3PlayerFrm1.Label3 = "00:00"
Picture2.Left = 0
MediaPlayer1.Stop
Label1.Caption = "             ^v^  Tongue 'N Groove MP3/CD Player  ^v^  "
MP3playerShowFav.Show
End Sub

Private Sub mnuFullScreenVideo_Click()
MP3PlayerFrm2.Show
MP3PlayerFrm2.WindowState = 2
End Sub

Private Sub mnuFireWorks_Click()
MP3PlayerFrm2.Show
End Sub

Private Sub mnuGreen_Click()
Mp3ID3.Picture = LoadPicture(App.Path & "\mp3id3green.jpg")
SkinLocDir = App.Path & "\mp3bg3Green.jpg"
SkinLoc = App.Path & "\mp3bgGreen.jpg"
SkinLocRev = App.Path & "\mp3RevGreen.jpg"
SkinLocPlay = App.Path & "\mp3PlayGreen.jpg"
SkinLocStop = App.Path & "\mp3StopGreen.jpg"
SkinLocPause = App.Path & "\mp3PauseGreen.jpg"
SkinLocFFW = App.Path & "\mp3FFWGreen.jpg"
SkinLocShow = App.Path & "\mp3BG2Green.jpg"
SkinLocID3 = App.Path & "\mp3ID3Green.jpg"
MP3PlayerFrm1.Picture = LoadPicture(App.Path & "\mp3bggreen.jpg")
MP3playerShowFav.Picture = LoadPicture(App.Path & "\mp3bg2green.jpg")
MP3CDFrm.Picture = LoadPicture(App.Path & "\mp3bggreen.jpg")
Mp3AddDirFrm.Picture = LoadPicture(App.Path & "\mp3bg3green.jpg")
Image4.Picture = LoadPicture(App.Path & "\mp3Revgreen.jpg")
Image1.Picture = LoadPicture(App.Path & "\mp3Playgreen.jpg")
Image2.Picture = LoadPicture(App.Path & "\mp3Stopgreen.jpg")
Image3.Picture = LoadPicture(App.Path & "\mp3Pausegreen.jpg")
Image5.Picture = LoadPicture(App.Path & "\mp3FFWgreen.jpg")
MP3playerShowFav.Option1.BackColor = &H63A817
MP3playerShowFav.Option3.BackColor = &H63A817
MP3playerShowFav.Option1.BackColor = &H63A817
MP3playerShowFav.List1.ForeColor = &H63A817
MP3playerShowFav.List2.ForeColor = &H63A817
MP3playerShowFav.Text1.ForeColor = &H63A817
Mp3AddDirFrm.Dir1.ForeColor = &H63A817
Mp3AddDirFrm.Drive1.ForeColor = &H63A817
MP3CDFrm.Combo1.ForeColor = &H63A817
MP3CDFrm.Label4.ForeColor = &H63A817
MP3CDFrm.Label5.ForeColor = &H63A817
MP3CDFrm.Option1.BackColor = &H63A817
MP3CDFrm.Option2.BackColor = &H63A817
End Sub

Private Sub mnuGrey_Click()
Mp3ID3.Picture = LoadPicture(App.Path & "\mp3id3silver.jpg")
SkinLocDir = App.Path & "\mp3bg3Silver.jpg"
SkinLoc = App.Path & "\mp3bgSilver.jpg"
SkinLocRev = App.Path & "\mp3Revsilver.jpg"
SkinLocPlay = App.Path & "\mp3Playsilver.jpg"
SkinLocStop = App.Path & "\mp3Stopsilver.jpg"
SkinLocPause = App.Path & "\mp3Pausesilver.jpg"
SkinLocFFW = App.Path & "\mp3FFWsilver.jpg"
SkinLocShow = App.Path & "\mp3BG2Silver.jpg"
SkinLocID3 = App.Path & "\mp3ID3Silver.jpg"
MP3PlayerFrm1.Picture = LoadPicture(App.Path & "\mp3bgsilver.jpg")
MP3playerShowFav.Picture = LoadPicture(App.Path & "\mp3bg2silver.jpg")
MP3CDFrm.Picture = LoadPicture(App.Path & "\mp3bgsilver.jpg")
Mp3AddDirFrm.Picture = LoadPicture(App.Path & "\mp3bg3silver.jpg")
Image4.Picture = LoadPicture(App.Path & "\mp3Revsilver.jpg")
Image1.Picture = LoadPicture(App.Path & "\mp3Playsilver.jpg")
Image2.Picture = LoadPicture(App.Path & "\mp3Stopsilver.jpg")
Image3.Picture = LoadPicture(App.Path & "\mp3Pausesilver.jpg")
Image5.Picture = LoadPicture(App.Path & "\mp3FFWsilver.jpg")
MP3playerShowFav.Option1.BackColor = &H6A6A6A
MP3playerShowFav.Option3.BackColor = &H6A6A6A
MP3playerShowFav.List1.ForeColor = &H6A6A6A
MP3playerShowFav.List2.ForeColor = &H6A6A6A
MP3playerShowFav.Text1.ForeColor = &H6A6A6A
Mp3AddDirFrm.Dir1.ForeColor = &H6A6A6A
Mp3AddDirFrm.Drive1.ForeColor = &H6A6A6A
MP3CDFrm.Combo1.ForeColor = &H6A6A6A
MP3CDFrm.Label4.ForeColor = &H6A6A6A
MP3CDFrm.Label5.ForeColor = &H6A6A6A
MP3CDFrm.Option1.BackColor = &H6A6A6A
MP3CDFrm.Option2.BackColor = &H6A6A6A
End Sub

Private Sub mnuID3_Click()
If MP3PlayerFrm1.MediaPlayer1.Filename = "" Then
MsgBox "There is no MP3 playing."
Exit Sub
End If
GetId3 MediaPlayer1.Filename
End Sub

Private Sub mnuMatrix_Click()
Mp3Matrix.Show
End Sub

Private Sub mnuPipes_Click()
Mp3Pipes.Show
End Sub

Private Sub mnuPurple_Click()
Mp3ID3.Picture = LoadPicture(App.Path & "\mp3id3.jpg")
SkinLocDir = App.Path & "\mp3bg3.jpg"
SkinLoc = App.Path & "\mp3bg.jpg"
SkinLocRev = App.Path & "\mp3Rev.jpg"
SkinLocPlay = App.Path & "\mp3Play.jpg"
SkinLocStop = App.Path & "\mp3Stop.jpg"
SkinLocPause = App.Path & "\mp3Pause.jpg"
SkinLocFFW = App.Path & "\mp3FFW.jpg"
SkinLocShow = App.Path & "\mp3BG2.jpg"
SkinLocID3 = App.Path & "\mp3ID3.jpg"
MP3PlayerFrm1.Picture = LoadPicture(App.Path & "\mp3bg.jpg")
MP3playerShowFav.Picture = LoadPicture(App.Path & "\mp3bg2.jpg")
MP3CDFrm.Picture = LoadPicture(App.Path & "\mp3bg.jpg")
Mp3AddDirFrm.Picture = LoadPicture(App.Path & "\mp3bg3.jpg")
Image4.Picture = LoadPicture(App.Path & "\mp3Rev.jpg")
Image1.Picture = LoadPicture(App.Path & "\mp3Play.jpg")
Image2.Picture = LoadPicture(App.Path & "\mp3Stop.jpg")
Image3.Picture = LoadPicture(App.Path & "\mp3Pause.jpg")
Image5.Picture = LoadPicture(App.Path & "\mp3FFW.jpg")
MP3playerShowFav.Option1.BackColor = &HB31AB3
MP3playerShowFav.Option3.BackColor = &HB31AB3
MP3playerShowFav.List1.ForeColor = &HB31AB3
MP3playerShowFav.List2.ForeColor = &HB31AB3
MP3playerShowFav.Text1.ForeColor = &HB31AB3
Mp3AddDirFrm.Dir1.ForeColor = &HB31AB3
Mp3AddDirFrm.Drive1.ForeColor = &HB31AB3
MP3CDFrm.Combo1.ForeColor = &HB31AB3
MP3CDFrm.Label4.ForeColor = &HB31AB3
MP3CDFrm.Label5.ForeColor = &HB31AB3
MP3CDFrm.Option1.BackColor = &HB31AB3
MP3CDFrm.Option2.BackColor = &HB31AB3
End Sub

Private Sub mnuShapes_Click()
Mp3Shapes.Show
End Sub



Private Sub mnuTwister_Click()
Mp3Twist.Show
End Sub

Private Sub Slider1_Scroll()
Timer4.Enabled = False
MediaPlayer1.CurrentPosition = Slider1.Value
Timer4.Enabled = True
End Sub

Private Sub Slider2_Change()
If MP3CDFrm.Visible = True Then
If Slider2.Value = 0 Then
volume = 0
CurVol = 0
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 1 Then
volume = 10000
CurVol = 10000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 2 Then
volume = 15000
CurVol = 15000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 3 Then
volume = 20000
CurVol = 20000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 4 Then
volume = 25000
CurVol = 25000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 5 Then
volume = 32000
CurVol = 32000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 6 Then
volume = 39000
CurVol = 39000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 7 Then
volume = 47000
CurVol = 47000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 8 Then
volume = 56000
CurVol = 56000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 9 Then
volume = 60000
CurVol = 60000
SetVolumeControl hmixer, VolCtrl, volume
ElseIf Slider2.Value = 10 Then
volume = 65500
CurVol = 65500
SetVolumeControl hmixer, VolCtrl, volume
End If

Else
If Slider2.Value = 0 Then
MediaPlayer1.volume = -10000

ElseIf Slider2.Value = 1 Then
MediaPlayer1.volume = -3000

ElseIf Slider2.Value = 2 Then
MediaPlayer1.volume = -2000

ElseIf Slider2.Value = 3 Then
MediaPlayer1.volume = -1500

ElseIf Slider2.Value = 4 Then
MediaPlayer1.volume = -1000

ElseIf Slider2.Value = 5 Then
MediaPlayer1.volume = -900

ElseIf Slider2.Value = 6 Then
MediaPlayer1.volume = -700

ElseIf Slider2.Value = 7 Then
MediaPlayer1.volume = -500

ElseIf Slider2.Value = 8 Then
MediaPlayer1.volume = -200

ElseIf Slider2.Value = 9 Then
MediaPlayer1.volume = -50

ElseIf Slider2.Value = 10 Then
MediaPlayer1.volume = 0
End If

End If
End Sub



Private Sub Slider3_Change()
On Error Resume Next
If Slider3.Value > -500 And Slider2.Value < 500 Then

End If
If Slider3.Value < -500 Then

End If
If Slider3.Value > 500 Then

End If
MediaPlayer1.Balance = Slider3.Value

End Sub







Private Sub Timer3_Timer()

RTrim (Label1)
LTrim (Label1)
B = Len(Label1)
a = 1
Do: DoEvents
Pause (0.2)
a = a + 1
C = Mid(Label1, 1, a)
Text2 = C
Call FadePreview2(RichTextBox1, FadeFourColor(244, 45, 99, 99, 45, 233, 245, 78, 99, 99, 45, 233, Text2, False))
Loop Until Len(Text2) = Len(Label1)
Pause (2)
Do: DoEvents
Pause (0.2)
a = a - 1
C = Left(Label1, a)
Text2 = C
Call FadePreview2(RichTextBox1, FadeFourColor(244, 45, 99, 99, 45, 233, 245, 78, 99, 99, 45, 233, Text2, False))
Loop Until Len(Text2) = 0

End Sub

Private Sub Timer4_Timer()
On Error Resume Next
Dim Cent
Dim iTheTime As Integer
Dim iTheSeconds As Integer
Dim iTheMinutes As Integer
iTheTime = CInt(MediaPlayer1.CurrentPosition)
iTheSeconds = iTheTime Mod 60
iTheMinutes = iTheTime \ 60
Label2.Caption = Format(iTheMinutes, "00") & ":" & Format(iTheSeconds, "00")
Slider1.Max = MediaPlayer1.Duration
Slider1.Value = MediaPlayer1.CurrentPosition
If Label2 = "0:-01" Then
Label2 = "00:00"
Timer4.Enabled = False
End If
Per = MediaPlayer1.CurrentPosition
Picture1.ScaleWidth = 143
Cent = Percent(Per, MediaPlayer1.Duration, Picture1.Width / 100 * 5.3)
Picture2.Left = Cent
End Sub



Private Sub VScroll1_Change()
If VScroll1.Value = 10 Then
MediaPlayer1.volume = -10000
VScroll1.TaG = -10000
ElseIf VScroll1.Value = 9 Then
MediaPlayer1.volume = -3000
VScroll1.TaG = -3000
ElseIf VScroll1.Value = 8 Then
MediaPlayer1.volume = -2000
VScroll1.TaG = -2000
ElseIf VScroll1.Value = 7 Then
MediaPlayer1.volume = -1500
VScroll1.TaG = -1500
ElseIf VScroll1.Value = 6 Then
MediaPlayer1.volume = -1000
VScroll1.TaG = -1000
ElseIf VScroll1.Value = 5 Then
MediaPlayer1.volume = -900
VScroll1.TaG = -900
ElseIf VScroll1.Value = 4 Then
MediaPlayer1.volume = -700
VScroll1.TaG = -700
ElseIf VScroll1.Value = 3 Then
MediaPlayer1.volume = -500
VScroll1.TaG = -500
ElseIf VScroll1.Value = 2 Then
MediaPlayer1.volume = -200
VScroll1.TaG = -200
ElseIf VScroll1.Value = 1 Then
MediaPlayer1.volume = -50
VScroll1.TaG = -50
ElseIf VScroll1.Value = 0 Then
MediaPlayer1.volume = 0
VScroll1.TaG = 0
End If
End Sub


