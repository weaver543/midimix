Attribute VB_Name = "midi"

Option Explicit
Option Compare Text
Public Const BUILD_DATE = "1/2/03" '"10/25/03w" '"8/6/2003 h"

Public prognum(16) As Integer

Public Declare Function GetDoubleClickTime Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const notebase = 60
Public Const PATCH_SELECTED = &H80C0FF  'yellow
Public Const PATCH_NOT_SELECTED = &HFFFFFF
Public Const CHANNEL_SELECTED = &H80C0FF  'yellow

' DYNAMIC CONTROL DEFINITIONS
Public Const KEYBOARD = 0
Public Const CHANNEL = 1
Public Const PATCH = 2
Public Const METER = 3
Public Const PARAMETER = 4
Public Const ANIMATION = 5
Public Const RECEIVE_SWITCH = 6
Public Const KEYBOARD_SELECT = 7
Public Const KEYBOARD_DESELECT = 8
' the following MUST be the last control in the array???
Public Const STATLABEL = 9
Public Const OLDKEYBOARD = 10
Public Const MAX_MODULES = 11 ' must be 1 more than last module index
Public curDevice As Long

Public Const FADERCHAN = 22 ' chan to send CC's on for keyboard fader
Public Const MAX_CC_TYPES = 50
Public Const MAX_KEYMAPS = 20
Public Const INCREMENT_AMT = 45
Public Const MAX_PRESETS = 5

Public Const WINDEFAULT = 0
Public Const UPDOWN = 7
Public Const UPARROW = 10
Public Const HOURGLASS = 11
Public Const CIRCLESLASH = 12

Public Const CHROMATIC = 0
Public Const DIATONIC = 1
Public Const PENTATONIC = 2
Public Const CHORDAL = 3

Public Const MAX_APPS = 3
Public Const MAX_IGNORE = 5
Public Const MAX_BANKS = 20
Public Const MAX_SONGSECTIONS = 30
Public Const LOWEST_NOTE = 48
Public Const HIGHEST_NOTE = 84
Public sClassName As String
Public Declare Function FindWindow Lib "user 32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public patchbuffer(16) As String
Public alt_patch(16) As Integer
Public oldchan As Integer, keypress_mode As Boolean
Public ringing_note As Integer, drumkeymode As Boolean, fadermode As Boolean
Public shifted As Boolean, lastplaykey As String, sentshiftedkey As Boolean
Public par
Public remap_mode As Integer, midimon_dec_mode As Boolean
Public remap_data_slider As Boolean, cc_interval As Integer
Public remap_pedal As Boolean
Public thru_input As Long, thru_output As Long, doubleclick_time As Long
Public mousedown_time As Long
Public sSysEx As String * 256
Public Const APP_NAME = "Midimix"
Global Const LENMIDIHDR = 28

Public lo_note As Integer, hi_note As Integer, delayed_note As Integer, num_apps As Integer
Public scale_mode As Integer, num_ccs As Integer, num_visible_ccs As Integer
Public idx As Integer, num_presets As Integer, increment As Integer

Public songpath As String, passed_filename As String, goto_key As String, followup_key As String
Public mci_active As Boolean, mci_playing As Boolean, thruports_open As Boolean
Public update_with_midi As Boolean, update_patchnames As Boolean
Public midi_monitoring As Boolean, settings_changed As Boolean
Public invisible_controllers As Boolean

Public bypass_pg(16) As Boolean
Public num_modules As Integer, i As Integer, j As Integer, k As Integer, num_passkeys As Integer, show_keycodes As Integer, switch_key As Integer
Public ignore_channel(16) As Integer
Public bank_descrip(MAX_BANKS) As String, bank_filename(MAX_BANKS) As String
Public bank_msb(MAX_BANKS) As String, bank_lsb(MAX_BANKS) As String
Public bank(16) As Integer

Public patch_name(16) As String
Public presets(MAX_PRESETS) As String
Public patch_msb(16) As Integer, patch_lsb(16) As Integer, patch_pg(16) As Integer
Public cctypes(MAX_CC_TYPES) As Integer
Public ccval(MAX_CC_TYPES, 16) As Integer
Public active_cc_idx As Integer

Public num_maps As Integer
Public map_statusbyte(MAX_KEYMAPS) As Integer
Public map_data1(MAX_KEYMAPS) As Integer
Public map_data2(MAX_KEYMAPS) As Integer
Public map_key(MAX_KEYMAPS) As String
Public passthru_in(MAX_KEYMAPS) As Integer
Public passthru_out(MAX_KEYMAPS) As String
Public appTitle(MAX_APPS) As String
Public active_appnum As Integer


Global num_banks As Integer

Public songsection(MAX_SONGSECTIONS) As String
Public num_songsections As Integer
Public getting_dump As Boolean, dump_enabled As Boolean, transposing As Boolean
Public transpose_amt As Integer
Const OFF_DELAY As Integer = 755
Public record_mode As Boolean, show_inactive As Boolean, patchform_loaded As Boolean, multibank_mode As Boolean
Dim stage(16) As Integer
Public meter_on As Boolean, showkeyboard As Boolean, animating As Boolean
Public on_notes(16) As Integer
Const DX = 90
Const DY = 70
Const BAR_INC = 20
Public Const SOLO_ON = &HFF  'red
Public Const SOLO_OFF = &HC0C0C0  'grey
Public Const CH_ON = &H8080FF       '&HC0E0FF     '
Public Const CH_OFF = &HE0E0E0  '&H808080
Public chan As Integer
Public solo As Integer
Public tmp As Integer
Public tmp_str As String '* 255
Public tmp_str2 As String '* 255
Public status_str As String * 255
Public numDevices As Long
Public user_chan As Integer
Public songfile As String, bankdir As String
Public mciout As Integer
Public playing(16) As Boolean
Public visible_modules(MAX_MODULES) As Integer
Public row_ht(MAX_MODULES) As Integer
Public x_offset(MAX_MODULES) As Integer
Public y_offset(MAX_MODULES) As Integer
Public descrip(MAX_MODULES) As String
Global vColor

' this is the main startup form

Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
' setmessage 3rd arg may need to be defined as string also!

Public hMidiIn As Long, hMidiIn2 As Long, hMidiOut As Long, hMidiOut2 As Long
Public midiMessageOut As Long       ' short message status byte
Public midiData1 As Long            ' short message data byte
Public midiData2 As Long            ' short message data byte
Public CurChannel As Integer        ' short msg  channel/part sequence 0-15

' *** API ***
' general
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020     ' (DWORD) dest = source

' midi
Public Const MAXPNAMELEN = 32       '  max product name length (including NULL)


Public Const MCI_SEQ_SET_PORT = &H20000
Public Const MCI_SET = &H80D
Public Const NO_ERROR = &H0
'Public Const MCI_OPEN = &H803
'public const MCI_CLOSE = &H804

Public Type MIDIHDR
  lpData As String            'Address of MIDI data
  dwBufferLength As Long      'Size of the buffer
  dwBytesRecorded As Long     'Actual amount of data in the buffer. This value should be less than or equal to the value given in the dwBufferLength member
  dwUser As Long              'Custom user data
  dwFlags As Long
    lpNext As Long              'Reserved - do not use
    Reserved As Long            'Reserved - do not use
End Type
'Flags giving information about the buffer
'MHDR_DONE : Set by the device driver to indicate that it is finished with the buffer and is returning it to the application
'MHDR_INQUEUE : Set by Windows to indicate that the buffer is queued for playback
'MHDR_ISSTRM : Set to indicate that the buffer is a stream buffer.
'MHDR_PREPARED : Set by Windows to indicate that the buffer has been prepared by using the midiInPrepareHeader or midiOutPrepareHeader function

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function midiOutPrepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiOutUnprepareHeader Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiOutLongMsg Lib "winmm.dll" (ByVal hMidiOut As Long, lpMidiOutHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiInPrepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiInAddBuffer Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long
Public Declare Function midiInUnprepareHeader Lib "winmm.dll" (ByVal hMidiIn As Long, lpMidiInHdr As MIDIHDR, ByVal uSize As Long) As Long

Type MIDIINCAPS
  wMid As Integer
  wPid As Integer
  vDriverVersion As Long
  szPname As String * MAXPNAMELEN
End Type

Type MIDIOUTCAPS
  wMid As Integer
  wPid As Integer
  vDriverVersion As Long
  szPname As String * MAXPNAMELEN
  wTechnology As Integer
  wVoices As Integer
  wNotes As Integer
  wChannelMask As Integer
  dwSupport As Long
End Type

Type MCI_SEQ_SET_PARMS
  dwCallback As Long
  dwTimeFormat As Long
  dwAudio As Long
  dwTempo As Long
  dwPort As Long
  dwSlave As Long
  dwMaster As Long
  dwOffset As Long
End Type


' MIDI API Functions
Declare Function midiConnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiDisconnect Lib "winmm.dll" (ByVal hmi As Long, ByVal hmo As Long, pReserved As Any) As Long
Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Declare Function midiInGetErrorText Lib "winmm.dll" Alias "midiInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiInReset Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIn As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutMessage Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal msg As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" _
    Alias "mciGetErrorStringA" _
    (ByVal dwError As Long, _
    ByVal lpstrBuffer As String, _
    ByVal uLength As Long) As Long

Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" _
   (ByVal wDeviceID As Long, _
   ByVal uMessage As Long, _
   ByVal dwParam1 As Long, _
   ByRef dwParam2 As Any) As Long
    

Declare Function GetShortPathName Lib "kernel32" _
    Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
    ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long

Public Const MMSYSERR_NOERROR = 0               '  no error

Public Const CALLBACK_NULL = &H0                '  no callback
Public Const CALLBACK_FUNCTION = &H30000        '  dwCallback is a FARPROC

Public Const MM_MIM_OPEN = &H3C1                '  MIDI input
Public Const MM_MIM_CLOSE = &H3C2
Public Const MM_MIM_DATA = &H3C3
Public Const MM_MIM_LONGDATA = &H3C4
Public Const MM_MIM_ERROR = &H3C5
Public Const MM_MIM_LONGERROR = &H3C6

' MIDI status messages
Public Const NOTE_OFF = &H80
Public Const NOTE_ON = &H90
Public Const POLY_KEY_PRESS = &HA0
Public Const CONTROLLER_CHANGE = &HB0
Public Const PROGRAM_CHANGE = &HC0
Public Const CHANNEL_PRESSURE = &HD0
Public Const PITCH_BEND = &HE0

'MIDI Mapper
Public Const MIDI_MAPPER = -1

'  flags for wTechnology field of MIDIOUTCAPS structure
Public Const MOD_MIDIPORT = 1    '  output port
Public Const MOD_SYNTH = 2       '  generic internal synth
Public Const MOD_SQSYNTH = 3     '  square wave internal synth
Public Const MOD_FMSYNTH = 4     '  FM internal synth
Public Const MOD_MAPPER = 5      '  MIDI mapper

'  flags for dwSupport field of MIDIOUTCAPS
Public Const MIDICAPS_VOLUME = &H1           '  supports volume control
Public Const MIDICAPS_LRVOLUME = &H2         '  separate left-right volume control
Public Const MIDICAPS_CACHE = &H4




' used to find partial window string
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias _
 "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private psAppNameContains As String
Private pbFound As Boolean


'Public Function isNote(ByVal Nr As Long) As String
'  Dim octave As Long
'  Dim note As String
'  octave = (Nr \ 12)
'  note = Nr Mod 12
'  isNote = Choose(note + 1, "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B") & Format(octave - 1)
'End Function

Public Sub ShowMMErr(InFunct As String, MMErr)
  Dim msg As String
  
  msg = String(255, " ")
  If InStr(1, InFunct, "out", vbTextCompare) = 0 Then
     midiInGetErrorText MMErr, msg, 255
  Else
     midiOutGetErrorText MMErr, msg, 255
  End If
  msg = InFunct & vbCrLf & msg & vbCrLf
  MsgBox msg
End Sub

Public Sub SendMidiShortOut()
  Dim midiMessage As Long
  Dim lowint As Long, highint As Long
  
  lowint = (midiData1 * 256) + midiMessageOut
  highint = (midiData2 * 256) * 256
  midiMessage = lowint + highint
  midiOutShortMsg hMidiOut, midiMessage
End Sub

Public Sub animate()

on_notes(chan) = on_notes(chan) + 1

With zmain.blob(chan)

Select Case stage(chan)
 Case 0:
  .Left = .Left - DX
  .Width = .Width + DX
 Case 1:
  .Top = .Top + DY
  .Height = .Height - DY
 Case 2:
  .Top = .Top - DY
  .Height = .Height + DY
 Case 3:
  .Width = .Width - DX
 Case 4:
  .Width = .Width + DX
 Case 5:
  .Height = .Height - DY
 Case 6:
  .Height = .Height + DY
 Case 7:
  .Left = .Left + DX
  .Width = .Width - DX
 End Select

.Visible = True
 End With
 
stage(chan) = stage(chan) + 1
If stage(chan) = 8 Then stage(chan) = 0

If on_notes(chan) = 1 Then zmain.blobtimer(chan).Interval = 0
End Sub
  
Public Function Replace5(sIn As String, sFind As String, _
  sReplace As String, Optional nStart As Long = 1, _
  Optional nCount As Long = -1, Optional bCompare As _
  VbCompareMethod = vbBinaryCompare) As String

  Dim nC As Long, nPos As Integer, sOut As String
  sOut = sIn
  nPos = InStr(nStart, sOut, sFind, bCompare)
  If nPos = 0 Then GoTo EndFn:
  Do
    nC = nC + 1
    sOut = Left(sOut, nPos - 1) & sReplace & _
       Mid(sOut, nPos + Len(sFind))
    If nCount <> -1 And nC >= nCount Then Exit Do
    nPos = InStr(nStart, sOut, sFind, bCompare)
  Loop While nPos > 0
EndFn:
  Replace5 = sOut
End Function

'vb5 implementation of split() in vb6
Public Function Split5(ByVal sIn As String, _
  Optional sDelim As String, Optional nLimit As Long = -1, _
  Optional bCompare As VbCompareMethod = vbBinaryCompare) As Variant
  Dim sOut() As String
  Dim sRead As String, nC As Integer

  If sDelim = "" Then sDelim = " "
   
  If InStr(sIn, sDelim) = 0 Then
    ReDim sOut(0) As String
    sOut(0) = sIn
    Split5 = sOut
    Exit Function
  End If

  sRead = ReadUntil(sIn, sDelim, bCompare)

  Do
    ReDim Preserve sOut(nC)
    sOut(nC) = sRead
    nC = nC + 1
    If nLimit <> -1 And nC >= nLimit Then Exit Do
    sRead = ReadUntil(sIn, sDelim)
  Loop While sRead <> "~TWA"
  
  ReDim Preserve sOut(nC)
  sOut(nC) = sIn
  Split5 = sOut
End Function
' used by split5()
Private Function ReadUntil(ByRef sIn As String, sDelim As String, Optional bCompare As VbCompareMethod = vbBinaryCompare) As String
  Dim nPos As Long
  nPos = InStr(1, sIn, sDelim, bCompare)
  If nPos > 0 Then
     ReadUntil = Left(sIn, nPos - 1)
     sIn = Mid(sIn, nPos + Len(sDelim))
  Else
     ReadUntil = "~TWA"
  End If
End Function

Public Function Inc(ByRef i As Integer) As Integer
  Inc = i
  i = i + 1
End Function

Public Function midisend(ByVal statusmsg As Long, ByVal data1msg As Long, Optional ByVal data2msg As Long)
  Dim midiMessage As Long, lowint As Long, highint As Long
  
Debug.Print "sending " + Hex(statusmsg) + "," + Hex(data1msg) + "," + Hex(data2msg)
  lowint = (data1msg * 256) + statusmsg
  highint = (data2msg * 256) * 256
  midiMessage = lowint + highint
  midiOutShortMsg hMidiOut, midiMessage
End Function


Public Sub say(txt As String)
  midimon.Text1.SelStart = Len(midimon.Text1.Text)
  midimon.Text1.SelText = txt + vbCrLf
End Sub

' next 3 functions are related
Public Function AppActivateByStringPart(StringPart As String) As Boolean

'PURPOSE: Activates the first window that contains any part of
'of StringPart
'PARAMETERS:
   'AppNamePart = Any Part of the WindowTitle for the App

'RETURNS: True if successful (i.e., a running app was found)
'false otherwise (e.g., no running app was found with StringPart
'as part of a title, or an error occurred

'EXAMPLE:
' ActivateAppByStringPart "Microsoft Internet Explorer"
'Will Activate the first running instance of IE it finds,
'even though the window title in most cases does
'not begin with with "Microsoft Internet Explorer"

Dim lRet As Long

psAppNameContains = StringPart
lRet = EnumWindows(AddressOf CheckForInstance, 0)

AppActivateByStringPart = pbFound
'reset
pbFound = False

End Function

Private Function CheckForInstance(ByVal lhWnd As Long, ByVal _
lParam As Long) As Long

Dim sTitle As String
Dim lRet As Long
Dim iNew As Integer

If Trim(psAppNameContains = "") Then
    CheckForInstance = False
    Exit Function
End If

sTitle = Space(255)
lRet = GetWindowText(lhWnd, sTitle, 255)

sTitle = StripNull(sTitle)
If InStr(sTitle, psAppNameContains) > 0 Then
 'we're done, stop looking
    CheckForInstance = False
    pbFound = True
    AppActivate sTitle
    
Else

    CheckForInstance = True
End If
End Function

Private Function StripNull(ByVal InString As String) As String

'Input: String containing null terminator (Chr(0))
'Returns: all character before the null terminator

Dim iNull As Integer
If Len(InString) > 0 Then
    iNull = InStr(InString, vbNullChar)
    Select Case iNull
    Case 0
        StripNull = InString
    Case 1
        StripNull = ""
    Case Else
       StripNull = Left$(InString, iNull - 1)
   End Select
End If

End Function


Public Function JUNKthiswontwork_getparthandle(StringPart As String)
  Dim lRet As Long
  
  psAppNameContains = StringPart
  lRet = EnumWindows(AddressOf CheckForInstance, 0)
  
  'AppActivateByStringPart = pbFound
  'reset
  pbFound = False
End Function



Public Sub MidiIN_Proc(ByVal hmIN As Long, ByVal wMsg As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long)
   Dim txt As String
   Dim status As Long, OnOff As Long
'   Dim data1 As Integer, data2 As Integer, etype As Integer
   Dim data1, data2, etype As Integer
 
  If curDevice = -2 Then Exit Sub
   
   With zmain
   'On Error Resume Next
   Select Case wMsg
     ' Case MM_MIM_OPEN: MM_MIM_CLOSE: MM_MIM_ERROR: MM_MIM_LONGERROR:
    Case MM_MIM_LONGDATA:
        If dump_enabled Then
          MsgBox "recieved longdata..."
          If status = &HF0 Then
            MsgBox "setting getting_dump"
            getting_dump = True
          End If
        End If

    Case MM_MIM_DATA:
        status = (dwParam1 Mod 256)
        
        If status = &HFE Then GoTo SKIPNOTE ' active sensing
        
        If getting_dump Then
          MsgBox "in middle of geting dump"
          'midishow_frm.Text1 = "!!" + midishow_frm.Text1 + status + vbCrLf
          GoTo SKIPNOTE
        End If
        
        If status < &HF0 Then ' if its not a SYSTEM message

        etype = Val(status And &HF0)
        data1 = ((dwParam1 \ 256) Mod 256)
        data2 = ((dwParam1 \ (256 ^ 2)) Mod 256)

        If midi_monitoring Then
          midimon.Text1.SelStart = Len(midimon.Text1.Text)
          If midimon_dec_mode Then
            midimon.Text1.SelText = CInt(status) + "," + CInt(data1) + "," + CInt(data2) + vbCrLf
          Else
            midimon.Text1.SelText = Hex(status) + "," + Hex(data1) + "," + Hex(data2) + vbCrLf
          End If
        End If
        
        ' RECORD mode
        
        If record_mode Then
'Left =B1,7,7F
'      B1 , 7, 0
'Right = B1, 6, 0
'        B1,26, A
'        B1,6,7F
'        B1,26,76
          If remap_mode = 2 Then ' 2-button sequencer remote
            If etype = &HB0 Then
              If data1 = 6 Then ' right button
                If data2 = &H7F Then ' button went up
                  If Not sentshiftedkey Then
                    SendKeys "^{PGUP}", True
                    lastplaykey = " "
                  End If
                ElseIf data2 = 0 Then
                  sentshiftedkey = False
                End If
                shifted = IIf(data2 = 0, True, False)
                GoTo SKIPNOTE
              ElseIf data1 = 7 Then  ' left button
                If data2 = 0 Then
                  If shifted Then
                    lastplaykey = "^z"
                    sentshiftedkey = True
                  Else
                    lastplaykey = IIf(lastplaykey = "R", " ", "R")
                  End If
                  SendKeys lastplaykey, True
                End If
                GoTo SKIPNOTE
              ElseIf data2 = &H26 Then
                GoTo SKIPNOTE
              End If
            End If
          ElseIf remap_mode = 3 Then ' step time
            If etype = &HB0 Then
              If data1 = 7 And data2 = 0 Then SendKeys "{TAB}{TAB}{TAB}{TAB}{TAB} ", True ' LEFT button went up
              GoTo SKIPNOTE
            End If
          ElseIf remap_mode = 1 Then ' standard remapping per ini file
            For i = 0 To num_maps - 1
              If etype = map_statusbyte(i) Then
                If data1 = map_data1(i) Then
                  If (map_data2(i) = &HFF) Or (data2 = map_data2(i)) Then
                    If map_key(i) <> "" Then
                      'If FindWindow(sClassName, appTitle) = 0 Then AppActivate appTitle
                      SendKeys map_key(i), True
                    End If
                    GoTo SKIPNOTE
                  End If
                End If
              End If
            Next
          End If

          If transposing Then
            If (etype = NOTE_ON) Or (etype = NOTE_OFF) Then
              tmp = (data1 + transpose_amt) * &H100
              dwParam1 = dwParam1 And &HFF00FF
              dwParam1 = dwParam1 Or tmp
            End If
          End If

          dwParam1 = dwParam1 And &HFFFFF0
          dwParam1 = dwParam1 Or user_chan
          midiOutShortMsg hMidiOut, dwParam1
          Exit Sub
        End If
          
        ' PLAYBACK MODE
                  
        chan = status And &HF 'Status \ 16
        
        If (etype = NOTE_ON) Then
        If data2 = 0 Then GoTo OFFNOTE
          If playing(chan) Then
              If meter_on Then _
                  If .bar(chan).Value < 100 Then .bar(chan).Value = .bar(chan).Value + BAR_INC
              If showkeyboard Then
                If ignore_channel(chan) = 0 Then
                 .ShowNote data1, 1, chan
                End If
              End If
              If animating Then Call animate
          Else: GoTo SKIPNOTE
          End If

        ElseIf etype = NOTE_OFF Then ' note off
OFFNOTE:
          If meter_on Then _
            If .bar(chan).Value > 0 Then .bar(chan).Value = .bar(chan).Value - BAR_INC
          If showkeyboard Then
            If ignore_channel(chan) < MAX_IGNORE Then ' allow 1st 5 notes to be turned off in case
              .ShowNote data1, 0, chan
              If ignore_channel(chan) > 0 Then ignore_channel(chan) = ignore_channel(chan) + 1
            End If
          End If
          If animating Then
             on_notes(chan) = on_notes(chan) - 1
             If on_notes(chan) = 0 Then
               .blobtimer(chan).Interval = OFF_DELAY
             End If
          End If
        ElseIf etype = CONTROLLER_CHANGE Then
          If remap_data_slider Then
            If data1 = 38 Then
              If data2 > 76 Then
                data2 = (data2 - 77) * cc_interval
              Else
                data2 = (data2 * cc_interval) + (9 * cc_interval)
              End If
            End If
          End If
          
          If remap_pedal Then
             If data1 = 67 Then
               data1 = 64
               dwParam1 = dwParam1 And &HFF00FF
               dwParam1 = dwParam1 Or &H4000
             End If
          End If
          
          If update_with_midi And multibank_mode Then
          
            If data1 = 0 Then ' cc#0
              patch_msb(chan) = data2
            ElseIf data1 = 32 Then 'cc #32
              patch_lsb(chan) = data2
            Else

              For i = 0 To num_ccs - 1
                If cctypes(i) = data1 Then
                  ccval(i, chan) = data2
                  If active_cc_idx <> -1 And i = active_cc_idx Then .param(chan).Caption = str(ccval(active_cc_idx, chan))
                  i = num_ccs
                Else
                End If
              Next
              
            End If
          End If
        ElseIf etype = PROGRAM_CHANGE Then
          If update_with_midi And data1 < 128 Then
           patch_pg(chan) = data1
            If update_patchnames Then
              If multibank_mode And patchform_loaded Then
' not implemented yet
'                For I = 0 To num_banks - 1
'                  If bank_msb(I) <> -1 And patch_msb(chan) = bank_msb(I) And patch_lsb(chan) = bank_lsb(I) Then
'                    Open bank_filename(I) For Input As #1
'                    For j = 0 To data1
'                      If Not EOF(1) Then Line Input #1, tmp_str
'                    Next
'                    If j = data1 Then .patch_label(chan).Caption = tmp_str
'                    I = num_banks
'                  End If
'                Next
              Else
'Debug.Print ("else part")
'Debug.Print ("its " + patchfrm.box(data1).Caption)
               .patch_label(chan).Caption = patchfrm.box(data1).Caption
'Debug.Print "done..."
              End If
            End If
          End If
        Else
        End If
        
        midiOutShortMsg hMidiOut, dwParam1
SKIPNOTE:
            
      End If
 
      'Case Else:
    End Select
    End With
End Sub
