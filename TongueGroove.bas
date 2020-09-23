Attribute VB_Name = "TongueGroove"
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1
Public Const MMSYSERR_NOERROR = 0
Public Const MAXPNAMELEN = 32
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Public Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2& ' separate left-right volume control
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Public Const MIXERLINE_COMPONENTTYPE_SRC_MIDIVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Public Const MIXERLINE_COMPONENTTYPE_SRC_I25InVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 1)
Public Const MIXERLINE_COMPONENTTYPE_SRC_TADVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 6)
Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
                             (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERLINE_COMPONENTTYPE_src_AUXVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 9)
Public Const MIXERLINE_COMPONENTTYPE_SRC_PSPKVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 7)
Public Const MIXERLINE_COMPONENTTYPE_SRC_MBOOST = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Public Const MIXERLINE_COMPONENTTYPE_SRC_LINEVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Public Const MIXERLINE_COMPONENTTYPE_SRC_CDVol = _
                             (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20
' Mixer control types
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Declare Function RegisterDLL Lib "Regist10.dll" Alias "REGISTERDLL" _
(ByVal DllPath As String, bRegister As Boolean) As Boolean

Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadFormat Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As waveFormat, ByVal cch As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
Declare Function mixerGetControlDetails Lib "winmm.dll" _
               Alias "mixerGetControlDetailsA" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
            
Declare Function mixerGetLineControls Lib "winmm.dll" _
               Alias "mixerGetLineControlsA" _
               (ByVal hmxobj As Long, _
               pmxlc As MIXERLINECONTROLS, _
               ByVal fdwControls As Long) As Long
               
Declare Function mixerGetLineInfo Lib "winmm.dll" _
               Alias "mixerGetLineInfoA" _
               (ByVal hmxobj As Long, _
               pmxl As MIXERLINE, _
               ByVal fdwInfo As Long) As Long
               
Declare Function mixerOpen Lib "winmm.dll" _
               (phmx As Long, _
               ByVal uMxId As Long, _
               ByVal dwCallback As Long, _
               ByVal dwInstance As Long, _
               ByVal fdwOpen As Long) As Long
               
Declare Function mixerSetControlDetails Lib "winmm.dll" _
               (ByVal hmxobj As Long, _
               pmxcd As MIXERCONTROLDETAILS, _
               ByVal fdwDetails As Long) As Long
               
Declare Sub CopyStructFromPtr Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (struct As Any, _
               ByVal ptr As Long, ByVal cb As Long)
               
Declare Sub CopyPtrFromStruct Lib "kernel32" _
               Alias "RtlMoveMemory" _
               (ByVal ptr As Long, _
               struct As Any, _
               ByVal cb As Long)
               
Declare Function GlobalAlloc Lib "kernel32" _
               (ByVal wFlags As Long, _
               ByVal dwBytes As Long) As Long
               
Declare Function GlobalLock Lib "kernel32" _
               (ByVal hmem As Long) As Long
               
Declare Function GlobalFree Lib "kernel32" _
               (ByVal hmem As Long) As Long

Dim rc As Long


' variables for managing wave file
Public formatA As waveFormat
Dim hmmioOut As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim bufferIn As Long
Dim hmem As Long
Public numSamples As Long
Public drawFrom As Long
Public drawTo As Long
Public fFileLoaded As Boolean

Type waveFormat
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type

Type MIXERCONTROL
   cbStruct As Long           '  size in Byte of MIXERCONTROL
   dwControlID As Long        '  unique control id for mixer device
   dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
   fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
   cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
   szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
   szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
   lMinimum As Long           '  Minimum value
   lMaximum As Long           '  Maximum value
   Reserved(10) As Long       '  reserved structure space
   End Type

Type MIXERCONTROLDETAILS
   cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
   dwControlID As Long    '  control id to get/set details on
   cChannels As Long      '  number of channels in paDetails array
   Item As Long           '  hwndOwner or cMultipleItems
   cbDetails As Long      '  size of _one_ details_XX struct
   paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
   dwValue As Long        '  value of the control (volume level)
End Type

Type MIXERLINE
   cbStruct As Long               '  size of MIXERLINE structure
   dwDestination As Long          '  zero based destination index
   dwSource As Long               '  zero based source index (if source)
   dwLineID As Long               '  unique line id for mixer device
   fdwLine As Long                '  state/information about line
   dwUser As Long                 '  driver specific information
   dwComponentType As Long        '  component type line connects to
   cChannels As Long              '  number of channels line supports
   cConnections As Long           '  number of connections (possible)
   cControls As Long              '  number of controls at this line
   szShortName As String * MIXER_SHORT_NAME_CHARS
   szName As String * MIXER_LONG_NAME_CHARS
   dwType As Long
   dwDeviceID As Long
   wMid  As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * MAXPNAMELEN
End Type

Type MIXERLINECONTROLS
   cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
   dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                          '  MIXER_GETLINECONTROLSF_ONEBYID or
   dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
   cControls As Long      '  count of controls pmxctrl points to
   cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
   pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type

Type mmioinfo
   dwFlags As Long
   fccIOProc As Long
   pIOProc As Long
   wErrorRet As Long
   htask As Long
   cchBuffer As Long
   pchBuffer As String
   pchNext As String
   pchEndRead As String
   pchEndWrite As String
   lBufOffset As Long
   lDiskOffset As Long
   adwInfo(4) As Long
   dwReserved1 As Long
   dwReserved2 As Long
   hmmio As Long
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F For version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
    End Type



Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)
    Public Const WS_CAPTION = &HC00000            ' WS_BORDER Or WS_DLGFRAME
Dim hmen As Long

Public Const HWND_TOPMOST = -1

Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000

Public Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum




Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Red As Long, Green As Long, Blue As Long, Color As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Const WM_VSCROLL = &H115
Public Const SB_LINEUP = 0
Public Const SB_LINEDOWN = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Type POINT_TYPE
  X As Long
  Y As Long
End Type
Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type
Public Type Id3                 'This type is standard for
Title As String * 30            ' Id3 Tags
Artist As String * 30           ' Although later versions
Album As String * 30            ' use comments for 28 bytes
sYear  As String * 4            ' and they use the 2 remaining  bytes for "TrackNumber"!
Comments As String * 30
Genre As Byte
End Type
Public ScrW%, ScrH%
Public TxtHght%, TxtWdth%
Public cy%, Dy%, K%, N%, Mx%
Public hMemDc&, hBmp&, hBmpOld&
Public hBrush&, hFont&, hFontOld&



Public Rct As RECT

Type StringData
     CurX As Integer
     CurY As Integer
     Dy As Integer
     NumChars As Integer
End Type

Public Mtrx(1 To 100) As StringData   ' One Hundred Output Strings.

Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x1&, ByVal y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hdc&, ByVal nWidth&, ByVal nHeight&)
Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hdc&)
Declare Function DeleteDC& Lib "gdi32" (ByVal hdc&)
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Declare Function FillRect& Lib "user32" (ByVal hdc&, lpRect As RECT, ByVal hBrush&)
Declare Function GetStockObject& Lib "gdi32" (ByVal nIndex&)
Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)
Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)
Declare Function SelectObject& Lib "gdi32" (ByVal hdc&, ByVal hObject&)
Declare Function SetBkMode& Lib "gdi32" (ByVal hdc&, ByVal nBkMode&)
Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal x1&, ByVal y1&, ByVal x2&, ByVal y2&)
Declare Function SetTextColor& Lib "gdi32" (ByVal hdc&, ByVal crColor&)
Declare Function SetTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Declare Function ShowCursor& Lib "user32" (ByVal bShow&)
Declare Function TextOut& Lib "gdi32" Alias "TextOutA" (ByVal hdc&, ByVal x1&, ByVal y1&, ByVal lpString$, ByVal nCount&)

Public Const BLACK_BRUSH = 4
Public Const TRANSPARENT = 1
Public Const WM_GETFONT = &H31
Public id3Info As Id3           ' Declare a variable as the id3 type
Public GenreArray() As String         ' we use this array to fill all the Genre's ( look in form load)

Public Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
' We can use the split function to fill this into an array
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2


Sub KillListDupes(Lst As Control)
On Error Resume Next
For i = 0 To Lst.ListCount - 1
For E = 0 To Lst.ListCount - 1
If LCase(Lst.List(i)) Like LCase(Lst.List(E)) And i <> E Then
Lst.RemoveItem (E)
End If
Next E
Next i
End Sub

Public Sub Stayontop(frm As Form)
Dim OnTop%
OnTop% = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub StayOnBottom(frm As Form)
Dim OnTop%
OnTop% = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Function RGBtoHEX(RGB)

    a$ = Hex(RGB)
    B% = Len(a$)
    If B% = 5 Then a$ = "0" & a$
    If B% = 4 Then a$ = "00" & a$
    If B% = 3 Then a$ = "000" & a$
    If B% = 2 Then a$ = "0000" & a$
    If B% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function
Public Sub MoveForm(frm As Form)
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
'MoveForm(frm)

End Sub
Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
 DoEvents
Loop
End Sub
Sub FadePreview2(RichTB As Control, ByVal FadedText As String)
'Modified by monk-e-god for use in a RichTextBox
On Error Resume Next
'NOTE: RichTB must be a RichTextBox.
'NOTE: You cannot preview wavy fades with this sub.
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%

RichTB.SelBold = True: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.Text = ""
For X = 1 To Len(FadedText$)

  C$ = Mid$(FadedText$, X, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If C$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case T$
      Case "u"
        RichTB.SelUnderline = True
      Case "/u"
        RichTB.SelUnderline = False
      Case "s"
        RichTB.SelStrikeThru = True
      Case "/s"
        RichTB.SelStrikeThru = False
      Case "b"    'start bold
        RichTB.SelBold = True
      Case "/b"   'stop bold
        RichTB.SelBold = False
      Case "i"    'start italic
        RichTB.SelItalic = True
      Case "/i"   'stop italic
        RichTB.SelItalic = False
      
      Case Else
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
            RichTB.SelStart = StartPlace%
            RichTB.SelFontName = dafont$
        End If
    End Select
  Else
    RichTB.SelText = RichTB.SelText + C$
    StartPlace% = StartPlace% + 1
    RichTB.SelStart = StartPlace%
  End If
  
Next X
End Sub
Function GETVAL%(ByVal strLetter$)
'by aDRaMoLEk
  Select Case strLetter$
    Case "0"
      GETVAL = 0
    Case "1"
      GETVAL = 1
    Case "2"
      GETVAL = 2
    Case "3"
      GETVAL = 3
    Case "4"
      GETVAL = 4
    Case "5"
      GETVAL = 5
    Case "6"
      GETVAL = 6
    Case "7"
      GETVAL = 7
    Case "8"
      GETVAL = 8
    Case "9"
      GETVAL = 9
    Case "A"
      GETVAL = 10
    Case "B"
      GETVAL = 11
    Case "C"
      GETVAL = 12
    Case "D"
      GETVAL = 13
    Case "E"
      GETVAL = 14
    Case "F"
      GETVAL = 15
  End Select
End Function
Function Hex2Dec!(ByVal strHex$)
'by aDRaMoLEk
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function
Public Function GetId3(Filename As String)
On Error Resume Next
Dim TaG As String * 3               ' We use this variable to make sure the file has an ID3TAG
Open Filename For Binary As #1      ' we open the file as binary for total control (we need it for the Genre part)
Get #1, FileLen(Filename) - 127, TaG    ' Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
If TaG = "TAG" Then                     ' "TAG" is put at position filesize-127 to show that this file indeed contains an Id3
Get #1, FileLen(Filename) - 124, id3Info
Mp3ID3.Show ' if the file has a tag, we put it into our earlier declared variable id3info
Else
MsgBox "No ID3 Information Is Available For That MP3"  ' if the "TAG" wasnt at position filesize-127
End If
Close #1                                            ' close the file
End Function
Public Function Send(Cmd As String)
    Static rc As Long
    Static errStr As String * 200

    rc = mciSendString(Cmd, 0, 0, hwnd)
    
    Send = (rc = 0)
End Function

Public Sub InitSaver()

    Dim MaxHeight%, MinHeight%

    ' Hide The Cursor.
    ShowCursor 1    ' (The Universe Crumbles Into A Heap Of Helplessness As The Mouse Cursor Is Removed).
    
    ' Cover The Screen With Our Form.
    Mp3Matrix.Move 0, 0, Screen.Width, Screen.Height

    ' Aquire The Screen Width And Height In Pixels.
    ScrW = GetSystemMetrics(0)
    ScrH = GetSystemMetrics(1)

    ' Setup A RECT Structure The Size Of The Screen.
    ' This Will Be Used Later With The API Function "FillRect"
    ' To Clear The Back Buffer.
    SetRect Rct, 0, 0, ScrW, ScrH
    ' Create A Brush To Fill The Rectangle With.
    hBrush = GetStockObject(BLACK_BRUSH)
    
    ' Create An Off Screen Drawing Area In Memory (Back Buffer)... (Backbuffer,.. That Picture NoOne Can See).
    hMemDc = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(Mp3Matrix.hdc, ScrW, ScrH)
    hBmpOld = SelectObject(hMemDc, hBmp)
    SetBkMode hMemDc, TRANSPARENT

    ' Get The Form's Font (Courier, Regular, 15)... (Just Call Me Spock!).
    hFont = SendMessage(Mp3Matrix.hwnd, WM_GETFONT, 0, 0&)
    ' Select It Into Our Back Buffer So We Can Output Text.
    hFontOld = SelectObject(hMemDc, hFont)

    TxtWdth = Mp3Matrix.TextWidth("A")
    TxtHght = Mp3Matrix.TextHeight("A")
    MaxHeight = ScrH - TxtHght

    ' Seed Random Number Generator.
    Randomize

    For K = 1 To 100
        Mtrx(K).CurX = Rnd * (ScrW - TxtWdth)
        Mtrx(K).NumChars = Int((20 - 5 + 1) * Rnd + 5)
        Mtrx(K).Dy = TxtHght + Rnd * TxtHght
        MinHeight = -2 * Mtrx(K).Dy * Mtrx(K).NumChars
        Mtrx(K).CurY = Int((MaxHeight - MinHeight + 1) * Rnd + MinHeight)
    Next

    ' Create An API Timer With An ID Of "1" And A Firing Interval
    ' Of 75 Milliseconds.
    SetTimer Mp3Matrix.hwnd, 1, 75, AddressOf TimerProc

End Sub
Public Sub TimerProc(ByVal hw&, ByVal msg&, ByVal id&, ByVal ntime&)

    ' ========================================
    ' API Timer Message Processing Department.
    ' ========================================

    ' Clear The BackBuffer.
    FillRect hMemDc, Rct, hBrush

    ' Output Our Strings.
    For K = 1 To 100
        cy = Mtrx(K).CurY
        Mx = Mtrx(K).NumChars
        For N = 1 To Mx
            If N = Mx Then ' Last Char In String.
               SetTextColor hMemDc, &H80FF80  ' The Brightest Letter.
            Else
               SetTextColor hMemDc, &H8000&   ' The Darker Letters.
            End If
            ' OutPut The Character On The Back Buffer.
            TextOut hMemDc, Mtrx(K).CurX, cy, Chr(Int((255 - 33 + 1) * Rnd + 33)), 1
            cy = cy + Mtrx(K).Dy
        Next
        Mtrx(K).CurY = Mtrx(K).CurY + Mtrx(K).Dy
        If Mtrx(K).CurY > ScrH Then UpdateMatrix
    Next

    ' Now That The Off Screen Drawing Is Complete,
    ' Blit The Finished Frame Onto The Screen.
    BitBlt Mp3Matrix.hdc, 0, 0, ScrW, ScrH, hMemDc, 0, 0, vbSrcCopy

End Sub
Public Sub UpdateMatrix()

    ' A String Has Now Left The Screen So
    ' Need To Initialize Another One.
    Mtrx(K).CurX = Rnd * (ScrW - TxtWdth)
    Mtrx(K).NumChars = Int((20 - 5 + 1) * Rnd + 5)
    Mtrx(K).Dy = TxtHght + Rnd * (TxtHght \ 2)
    Mtrx(K).CurY = -2 * Mtrx(K).Dy * Mtrx(K).NumChars

End Sub
Public Sub DeleteObjects()

    ' Delete The Font We Created.
    DeleteObject SelectObject(hMemDc, hFontOld)

    ' Delete The Back Buffer.
    DeleteObject SelectObject(hMemDc, hBmpOld)
    DeleteDC hMemDc

End Sub
Public Function ShowTitleBar(ByVal bState As Boolean)
Dim lStyle As Long
Dim tR As RECT

    ' Get the window's position:
    GetWindowRect MP3PlayerFrm1.hwnd, tR

    ' Modify whether title bar will be visible:
    lStyle = GetWindowLong(MP3PlayerFrm1.hwnd, GWL_STYLE)
    If (bState) Then
 
    If MP3PlayerFrm1.ControlBox Then
        lStyle = lStyle Or WS_SYSMENU
    End If
    If MP3PlayerFrm1.MaxButton Then
        lStyle = lStyle Or WS_MAXIMIZEBOX
    End If
    If MP3PlayerFrm1.MinButton Then
        lStyle = lStyle Or WS_MINIMIZEBOX
    End If
    If MP3PlayerFrm1.Caption <> "" Then
        lStyle = lStyle Or WS_CAPTION
    End If
    Else
    MP3PlayerFrm1.TaG = MP3PlayerFrm1.Caption
    'Me.Caption = "Blah1"
    lStyle = lStyle And Not WS_SYSMENU
    lStyle = lStyle And Not WS_MAXIMIZEBOX
    lStyle = lStyle And Not WS_MINIMIZEBOX
    lStyle = lStyle And Not WS_CAPTION
End If
SetWindowLong MP3PlayerFrm1.hwnd, GWL_STYLE, lStyle

' Ensure the style takes and make the window the
' same size, regardless that the title bar etc
' is now a different size:
SetWindowPos MP3PlayerFrm1.hwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
MP3PlayerFrm1.Refresh

' Ensure that your resize code is fired, as the client area
' has changed:
'Form_Resize

End Function

Function TaskBarIcon(frm As Form)
Dim lStyle As Long
lStyle = GetWindowLong(frm.hwnd, GWL_STYLE) Or WS_SYSMENU
SetWindowLong frm.hwnd, GWL_STYLE, lStyle
End Function
Function SetVolumeControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
  Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.Item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = volume
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetVolumeControl = True
   Else
       SetVolumeControl = False
   End If
End Function
Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
'This function Gets the value for a volume control. Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.Item = 0
    mxcd.cbDetails = Len(vol)
    mxcd.paDetails = 0
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function
Function GetMixerControl(ByVal hmixer As Long, _
                        ByVal componentType As Long, _
                        ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                        
' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = 1
       mxlc.cbmxctrl = Len(mxc)
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       If (MMSYSERR_NOERROR = rc) Then
           GetMixerControl = True
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetMixerControl = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If
   GetMixerControl = False
End Function
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
