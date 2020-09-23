VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETLINECONTROLSF_ALL = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Private Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
      
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000

Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
    (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
                     
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
      
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
      
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
    
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)

Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = _
    (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)

Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = _
    (MIXERCONTROL_CT_CLASS_METER Or _
    MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)

Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = _
    (MIXERCONTROL_CT_CLASS_SWITCH Or _
    MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = _
    (MIXERCONTROL_CT_CLASS_LIST Or _
    MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_FADER = _
    (MIXERCONTROL_CT_CLASS_FADER Or _
    MIXERCONTROL_CT_UNITS_UNSIGNED)
      
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Private Const MIXERCONTROL_CONTROLTYPE_BASS = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 2)
    
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 4)
    
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = _
    (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_MONO = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
    
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
    
Private Const MIXERCONTROL_CONTROLTYPE_ONOFF = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = _
    (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 3)

Private Const MIXERCONTROL_CONTROLTYPE_SLIDER = _
    (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)

Private Const MIXERCONTROL_CONTROLTYPE_PAN = _
    (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
    
Private Declare Function mixerClose Lib "winmm.dll" _
    (ByVal hmx As Long) As Long
         
Private Declare Function mixerGetControlDetails Lib "winmm.dll" _
    Alias "mixerGetControlDetailsA" _
    (ByVal hmxobj As Long, _
    pMxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
         
Private Declare Function mixerGetDevCaps Lib "winmm.dll" _
    Alias "mixerGetDevCapsA" _
    (ByVal uMxId As Long, _
    ByVal pmxcaps As MIXERCAPS, _
    ByVal cbmxcaps As Long) As Long
         
Private Declare Function mixerGetID Lib "winmm.dll" _
    (ByVal hmxobj As Long, _
    pumxID As Long, _
    ByVal fdwId As Long) As Long
                     
Private Declare Function mixerGetLineControls Lib "winmm.dll" _
    Alias "mixerGetLineControlsA" _
    (ByVal hmxobj As Long, _
    pmxlc As MIXERLINECONTROLS, _
    ByVal fdwControls As Long) As Long
                     
Private Declare Function mixerGetLineInfo Lib "winmm.dll" _
    Alias "mixerGetLineInfoA" _
    (ByVal hmxobj As Long, _
    pmxl As MIXERLINE, _
    ByVal fdwInfo As Long) As Long
                     
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
      
Private Declare Function mixerMessage Lib "winmm.dll" _
    (ByVal hmx As Long, _
    ByVal uMsg As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long) As Long
                     
Private Declare Function mixerOpen Lib "winmm.dll" _
    (phmx As Long, _
    ByVal uMxId As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long) As Long
                     
Private Declare Function mixerSetControlDetails Lib "winmm.dll" _
    (ByVal hmxobj As Long, _
    pMxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
              
Private Declare Sub CopyStructFromPtr Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (struct As Any, _
    ByVal ptr As Long, ByVal cb As Long)
                     
Private Declare Sub CopyPtrFromStruct Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (ByVal ptr As Long, _
    struct As Any, _
    ByVal cb As Long)
                     
Private Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
                     
Private Declare Function GlobalLock Lib "kernel32" _
    (ByVal hmem As Long) As Long
                     
Private Declare Function GlobalFree Lib "kernel32" _
    (ByVal hmem As Long) As Long
      
Private Type MIXERCAPS
    wMid As Integer                   '  manufacturer id
    wPid As Integer                   '  product id
    vDriverVersion As Long            '  version of the driver
    szPname As String * MAXPNAMELEN   '  product name
    fdwSupport As Long                '  misc. support bits
    cDestinations As Long             '  count of destinations
End Type
      
Private Type MIXERCONTROL
    cbStruct As Long                  '  size in Byte of MIXERCONTROL
    dwControlID As Long               '  unique control id for mixer device
    dwControlType As Long             '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long                '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long            '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
    szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
    lMinimum As Long                  '  Minimum value
    lMaximum As Long                  '  Maximum value
    Reserved(10) As Long              '  reserved structure space
End Type
      
Private Type MIXERCONTROLDETAILS
    cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
    dwControlID As Long    '  control id to get/set details on
    cChannels As Long      '  number of channels in paDetails array
    item As Long           '  hwndOwner or cMultipleItems
    cbDetails As Long      '  size of _one_ details_XX struct
    paDetails As Long      '  pointer to array of details_XX structs
End Type
      
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long        '  value of the control
End Type

Private Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type

Private Type MIXERCONTROLDETAILS_SIGNED
' The MIXERCONTROLDETAILS_SIGNED user-defined type retrieves and
' sets signed type control
' properties for an audio mixer control.
   lValue As Long
End Type

Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1 As Long
    dwParam2 As Long
    szName As String * MIXER_LONG_NAME_CHARS
End Type

Private Type MIXERLINE
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
      
Private Type MIXERLINECONTROLS
    cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
    dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                                '  MIXER_GETLINECONTROLSF_ONEBYID or
    dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long      '  count of controls pmxctrl points to
    cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
    pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type

Private hmixer As Long          ' mixer handle

'-----------------------------------------------------------------------------
'Volume
Private volCtrl As MIXERCONTROL ' waveout volume control
'Mute
Private volMute As MIXERCONTROL ' waveout volume control
'Fader
Private volFad As MIXERCONTROL
'Bass
Private volBass As MIXERCONTROL
'Treble
Private volTreb As MIXERCONTROL
'Pan
Private volPan As MIXERCONTROL
'-----------------------------------------------------------------------------
Private rc As Long              ' return code
Private ok As Boolean           ' boolean return code

' Audiolevel------------------------------------------------------------------
Private Const CALLBACK_FUNCTION = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1         '  done bit
Private Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin

Private Type WAVEHDR
' The WAVEHDR user-defined type defines the header used to identify a waveform-audio buffer.
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   Reserved As Long        ' Not used
End Type

Private Type WAVEINCAPS
   
   wMid As Integer

   wPid As Integer         ' Product identifier for the waveform-audio input device.
   
   vDriverVersion As Long  ' Version number of the device driver for the
   
   szPname As String * 32  ' Product name in a null-terminated string.
   dwFormats As Long       ' Standard formats that are supported. See the Platform
                           ' SDK product documentation for more information.
   wChannels As Integer    ' Number specifying whether the device supports
                           ' mono (1) or stereo (2) input.
End Type

Private Type WAVEFORMAT
   
   wFormatTag As Integer      ' Format type. Use the constant WAVE_FORMAT_PCM Waveform-audio data
                              ' to define the data as PCM.
   nChannels As Integer       ' Number of channels in the waveform-audio data. Mono data uses one
                              ' channel and stereo data uses two channels.
   nSamplesPerSec As Long     ' Sample rate, in samples per second.
   nAvgBytesPerSec As Long    ' Required average data transfer rate, in bytes per second. For
                              ' example, 16-bit stereo at 44.1 kHz has an average data rate of
                              ' 176,400 bytes per second (2 channels — 2 bytes per sample per
                              ' channel — 44,100 samples per second).
   nBlockAlign As Integer     ' Block alignment, in bytes. The block alignment is the minimum atomic unit of data. For PCM data, the block alignment is the number of bytes used by a single sample, including data for both channels if the data is stereo. For example, the block alignment for 16-bit stereo PCM is 4 bytes (2 channels — 2 bytes per sample).
   wBitsPerSample As Integer  ' For buffer estimation
   cbSize As Integer          ' Block size of the data.
End Type

Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
                                             ByVal uDeviceID As Long, _
                                             lpFormat As WAVEFORMAT, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal dwFlags As Long) As Long

Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                      lpWaveInHdr As WAVEHDR, _
                                                      ByVal uSize As Long) As Long

Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Private Declare Function waveInUnprepareHeader Lib "winmm.dll" _
                                          (ByVal hWaveIn As Long, _
                                          lpWaveInHdr As WAVEHDR, _
                                          ByVal uSize As Long) As Long

Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long

Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                  (ByVal uDeviceID As Long, _
                  lpCaps As WAVEINCAPS, _
                  ByVal uSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long

Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
                     (ByVal err As Long, _
                     ByVal lpText As String, _
                     ByVal uSize As Long) As Long

Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                   lpWaveInHdr As WAVEHDR, _
                                                   ByVal uSize As Long) As Long
Private i As Integer
Private j As Integer
Private msg As String * 200
Private hWaveIn As Long
Private format As WAVEFORMAT

Private Const NUM_BUFFERS = 2
Private Const BUFFER_SIZE = 8192
Private Const DEVICEID = 0
Private hmem(NUM_BUFFERS) As Long
Private inHdr(NUM_BUFFERS) As WAVEHDR

Private fRecording As Boolean

Private inputVolCtrl As MIXERCONTROL    ' waveout volume control
Private outputVolCtrl As MIXERCONTROL   ' microphone volume control

Private vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Private volume As Long                      ' volume value
Private volHmem As Long                     ' handle to volume memory
Private mxcd As MIXERCONTROLDETAILS         ' control info
Private Sub Class_Initialize()
    Debug.Print "îôòéì àú äà÷åìééæø"
    
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Der Mixer kann nicht geöffnet werden.", vbCritical + vbOKOnly, App.Title & _
        " - Mixercontrol Fehler"
        Exit Sub
    End If

' Volume----------------------------------------------------------------------
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        volCtrl)
' Mute------------------------------------------------------------------------
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        volMute)
'=============================================================================
' Fader-----------------------------------------------------------------------
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_FADER, _
        volFad)
' Bass------------------------------------------------------------------------
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_BASS, _
        volBass)
' Treble----------------------------------------------------------------------
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_TREBLE, _
        volTreb)
' Pan-------------------------------------------------------------------------
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_PAN, _
        volPan)
' Audiolevel------------------------------------------------------------------
         
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1

   ' Get the input volume meter
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, _
   MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   
   If (ok <> True) Then
       ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   End If
   If (ok <> True) Then
   End If

   ' Get the output volume meter
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
   MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   
   If (ok <> True) Then
      
   End If

End Sub
Private Sub Class_Terminate()
    Debug.Print "ñåâø àú äà÷åìééæø"
    mixerClose hmixer
    If (fRecording = True) Then
        StopInput
    End If
    GlobalFree volHmem

End Sub
Private Function SetVolume(mxc As MIXERCONTROL, _
                              ByVal volume As Long) As Boolean
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    vol.dwValue = volume
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
        SetVolume = True
    Else
        SetVolume = False
    End If
End Function

Private Function GetVolume(mxc As MIXERCONTROL) As Long
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
    
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    
    GetVolume = vol.dwValue
    
    GlobalFree (hmem)
    
End Function

Private Function GetTreble(mxc As MIXERCONTROL) As Long
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Treb As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Treb)
    
    hmem = GlobalAlloc(&H40, Len(Treb))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr Treb, mxcd.paDetails, Len(Treb)
    
    GetTreble = Treb.dwValue
    
    GlobalFree (hmem)
    
End Function


Private Function SetTreble(mxc As MIXERCONTROL, _
                              ByVal Treble As Long)
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Treb As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Treb)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(Treb))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Treb.dwValue = Treble
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Treb, Len(Treb)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
End Function

Private Function GetBass(mxc As MIXERCONTROL) As Long
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Bas As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Bas)
    
    hmem = GlobalAlloc(&H40, Len(Bas))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr Bas, mxcd.paDetails, Len(Bas)
    
    GetBass = Bas.dwValue
    
    GlobalFree (hmem)
    
End Function


Private Function SetBass(mxc As MIXERCONTROL, _
                              ByVal Bass As Long)
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Bas As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Bas)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(Bas))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Bas.dwValue = Bass
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Bas, Len(Bas)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
End Function


Private Function GetPan(mxc As MIXERCONTROL) As Boolean
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Pan As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Pan)
    
    hmem = GlobalAlloc(&H40, Len(Pan))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr Pan, mxcd.paDetails, Len(Pan)
    
    GetPan = Pan.dwValue
    
    GlobalFree (hmem)
    
End Function


Private Function SetPan(mxc As MIXERCONTROL, _
                              ByVal Pano As Long)
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Pan As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Pan)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(Pan))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Pan.dwValue = Pano
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Pan, Len(Pan)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
End Function

Private Function GetMute(mxc As MIXERCONTROL) As Boolean
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
    
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr Mut, mxcd.paDetails, Len(Mut)
    
    GetMute = IIf((Mut.fValue = 1), True, False)
    
    GlobalFree (hmem)
    
End Function

Private Function SetMute(mxc As MIXERCONTROL, _
                              ByVal Mute As Boolean)
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Mut.fValue = IIf(Mute, 1, 0)
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Mut, Len(Mut)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
End Function


Private Function GetFader(mxc As MIXERCONTROL) As Long
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Fad As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Fad)
    
    hmem = GlobalAlloc(&H40, Len(Fad))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr Fad, mxcd.paDetails, Len(Fad)
    
    GetFader = Fad.dwValue
    
    GlobalFree (hmem)
    
End Function

Private Function SetFader(mxc As MIXERCONTROL, _
                              ByVal Fader As Long)
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Fad As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Fad)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(Fad))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Fad.dwValue = Fader
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Fad, Len(Fad)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
End Function
' Get Min---------------------------------------------------------------------
Public Property Get VolumeMin() As Long
'Minimaler Wert am Ausgang
    ' Sollte normalerweise 0 sein
    VolumeMin = volCtrl.lMinimum
End Property
Public Property Get VolFadMin() As Long
    VolFadMin = volFad.lMinimum
End Property
Public Property Get VolTrebleMin() As Long
    VolTrebleMin = volTreb.lMinimum
End Property
Public Property Get VolBassMin() As Long
    VolBassMin = volBass.lMinimum
End Property
Public Property Get VolPanMin() As Long
    VolPanMin = volPan.lMinimum
End Property
' Get Max---------------------------------------------------------------------
Public Property Get VolumeMax() As Long
    VolumeMax = volCtrl.lMaximum
End Property
Public Property Get VolFaderMax() As Long
    VolFaderMax = volFad.lMaximum
End Property
Public Property Get VolTrebleMax() As Long
    VolTrebleMax = volTreb.lMaximum
End Property
Public Property Get VolBassMax() As Long
    VolBassMax = volBass.lMaximum
End Property
Public Property Get VolPanMax() As Long
    VolPanMax = volPan.lMaximum
End Property
' Get Level-------------------------------------------------------------------
Public Property Get VolumeLevel() As Long
'Aktueller Wert am Ausgang
    VolumeLevel = GetVolume(volCtrl)
End Property
' Get Bass--------------------------------------------------------------------
Public Property Get VolumeLevelBass() As Long
    VolumeLevelBass = GetBass(volBass)
End Property
' Get Treble------------------------------------------------------------------
Public Property Get VolumeLevelTreble() As Long
    VolumeLevelTreble = GetTreble(volTreb)
End Property
' Get Pan---------------------------------------------------------------------
Public Property Get VolumeLevelPan() As Long
    VolumeLevelPan = GetPan(volPan)
End Property
' Get Fader-------------------------------------------------------------------
Public Property Get VolumeLevelFader() As Long
    VolumeLevelFader = GetFader(volFad)
End Property
' Let Level aktuellen Wert Setzen---------------------------------------------
Public Property Let VolumeLevel(ByVal New_Volume As Long)
    SetVolume volCtrl, New_Volume
End Property
' Let Bass--------------------------------------------------------------------
Public Property Let VolumeLevelBass(ByVal New_Bass As Long)
    SetBass volBass, New_Bass
End Property
' Let Treble------------------------------------------------------------------
Public Property Let VolumeLevelTreble(ByVal New_Treble As Long)
    SetTreble volTreb, New_Treble
End Property
' Let Pan---------------------------------------------------------------------
Public Property Let VolumeLevelPan(ByVal New_Pan As Long)
    SetPan volPan, New_Pan
End Property
' Let Fader-------------------------------------------------------------------
Public Property Let VolumeLevelFader(ByVal New_Fader As Long)
    SetFader volFad, New_Fader
End Property
' Get Mute gibt True zurück wenn ein Aus-, Eingang stummgeschaltet ist--------
Public Property Get VolumeMute() As Boolean
    VolumeMute = GetMute(volMute)
End Property
' Let Mute True für Mute-------------------------------------------------------
Public Property Let VolumeMute(New_Mute As Boolean)
    SetMute volMute, New_Mute
End Property
' This function starts recording from the soundcard. The soundcard
' must be recording in order to
' monitor the input level. Without starting the recording from
' this application, input level
' can still be monitored if another application is recording audio
Private Function StartInput() As Boolean

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg, vbCritical + vbOKOnly, _
        App.Title & " - Mixercontrol Fehler"
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, _
            App.Title & " - Mixercontrol Fehler"
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, _
            App.Title & " - Mixercontrol Fehler"
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

' Stop receiving audio input on the soundcard
Private Sub StopInput()

    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub

' Function to process the wave recording notifications.
Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

Private Function GetVolumeControl(ByVal hmixer As Long, _
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
    rc = mixerGetLineInfo(hmixer, mxl, _
        MIXER_GETLINEINFOF_COMPONENTTYPE)
         
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
        rc = mixerGetLineControls(hmixer, _
            mxlc, _
            MIXER_GETLINECONTROLSF_ONEBYTYPE)
                  
        If (MMSYSERR_NOERROR = rc) Then
            GetVolumeControl = True
                 
            ' Copy the control into the destination structure
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
        Else
            GetVolumeControl = False
        End If
        GlobalFree (hmem)
        Exit Function
    End If
      
    GetVolumeControl = False
End Function


Public Property Get MaxVolumeMeterInput() As Long
    MaxVolumeMeterInput = inputVolCtrl.lMaximum
End Property

Public Property Get MinVolumeMeterInput() As Long
    ' Sollte normalerweise 0 sein
    MinVolumeMeterInput = inputVolCtrl.lMinimum
End Property

Public Property Get MaxVolumeMeterOutput() As Long
    MaxVolumeMeterOutput = outputVolCtrl.lMaximum
End Property

Public Property Get MinVolumeMeterOutput() As Long
    ' Sollte normalerweise 0 sein
    MinVolumeMeterOutput = outputVolCtrl.lMinimum
End Property

Public Property Get CurrentVolumeMeterOutput() As Long
    ' Process sound buffer if recording
    If (fRecording) Then
       For i = 0 To (NUM_BUFFERS - 1)
          If inHdr(i).dwFlags And WHDR_DONE Then
             rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
          End If
       Next
    End If

    ' Get the current output level
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, _
    MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
      
    If (volume < 0) Then volume = -volume
   
    CurrentVolumeMeterOutput = volume

End Property

Public Property Get CurrentVolumeMeterInput() As Long
' This also don´t really work fine :-(

    ' Process sound buffer if recording
    If Not StartInput Then StartInput


    If (fRecording) Then
        For i = 0 To (NUM_BUFFERS - 1)
            If inHdr(i).dwFlags And WHDR_DONE Then
                rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
            End If
        Next
    End If

   ' Get the current input level
    mxcd.dwControlID = inputVolCtrl.dwControlID
    mxcd.item = inputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, _
    MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then
        volume = -volume
    End If
    CurrentVolumeMeterInput = volume
End Property

