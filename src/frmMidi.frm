VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form zmain 
   BackColor       =   &H80000000&
   Caption         =   "mm"
   ClientHeight    =   4485
   ClientLeft      =   255
   ClientTop       =   2970
   ClientWidth     =   15000
   Icon            =   "frmMidi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   15000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   2520
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   29
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMidi.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMidi.frx":12BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMidi.frx":1746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMidi.frx":1C0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer transposenotetimer 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5400
      Top             =   2520
   End
   Begin VB.CommandButton hidden_btn 
      Caption         =   "Dont Press This!"
      Height          =   375
      Left            =   23440
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton kbutton 
      Height          =   255
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer firstnote_timer 
      Enabled         =   0   'False
      Left            =   6120
      Top             =   2520
   End
   Begin VB.PictureBox picKlav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   350
      Left            =   240
      MousePointer    =   10  'Up Arrow
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   525
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   7875
   End
   Begin VB.Timer blobtimer 
      Index           =   0
      Left            =   5760
      Top             =   2520
   End
   Begin MSComctlLib.ProgressBar bar 
      Height          =   615
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
      Orientation     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "."
   End
   Begin VB.CommandButton ch 
      BackColor       =   &H008080FF&
      Height          =   195
      Index           =   0
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label msg 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label klabel 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ptop 
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape pkey 
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label stat 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   12015
   End
   Begin VB.Label param 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape blob 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label patch_label 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "patch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label channel_label 
      Caption         =   "16"
      Height          =   225
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Menu file_menu 
      Caption         =   "File"
      Begin VB.Menu filenew_menu 
         Caption         =   "&New"
      End
      Begin VB.Menu loadsongsettings 
         Caption         =   "&Load Song Settings"
         Shortcut        =   ^Q
      End
      Begin VB.Menu savesongsettings 
         Caption         =   "&Save Song Settings"
         Shortcut        =   ^S
      End
      Begin VB.Menu savesongsettingsas 
         Caption         =   "Save Song Settings &As..."
      End
      Begin VB.Menu saveini_menu 
         Caption         =   "Save Initialization Settings"
      End
      Begin VB.Menu spacer1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mru_menu 
         Caption         =   ""
      End
      Begin VB.Menu spacer2 
         Caption         =   "-"
      End
      Begin VB.Menu about_menu 
         Caption         =   "About..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu file_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu midi_input 
      Caption         =   "I&n"
      Begin VB.Menu minput 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu midi_device 
      Caption         =   "Ou&t"
      Begin VB.Menu device 
         Caption         =   "MIDI Mapper"
         Index           =   0
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu mnu_keyfader 
         Caption         =   "&Keyboard Fader Control"
      End
      Begin VB.Menu mnuTransmitSettings 
         Caption         =   "&Transmit Songsettings"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu_remap 
         Caption         =   "Remap Data Slider"
      End
      Begin VB.Menu mnu_senddummy 
         Caption         =   "Send dummy event"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu update_with_midi_menu 
         Caption         =   "Update with midi input"
         Visible         =   0   'False
      End
      Begin VB.Menu update_patchnames_menu 
         Caption         =   "Update patchnames"
         Visible         =   0   'False
      End
      Begin VB.Menu multibank_mode_menu 
         Caption         =   "Search multiple bank files"
         Visible         =   0   'False
      End
      Begin VB.Menu fd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu preset_menu 
         Caption         =   "P&resets"
         Begin VB.Menu preset_menu_item 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu swap_ports_menu 
         Caption         =   "Swap Ports"
      End
      Begin VB.Menu show_ccs_menu 
         Caption         =   "Show Controller Values"
      End
      Begin VB.Menu ontop 
         Caption         =   "Always on top"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuTempDisconnect 
         Caption         =   "Temporarily disconnect ports"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnu_sendalt 
         Caption         =   "Send Alternative Patches"
      End
      Begin VB.Menu allnotesoff 
         Caption         =   "Send All Notes Off"
      End
   End
   Begin VB.Menu popups 
      Caption         =   "popups"
      Begin VB.Menu keyboard_menu 
         Caption         =   "keyboard menu"
         Begin VB.Menu setcolor_menu 
            Caption         =   "Set Color"
         End
         Begin VB.Menu all_on_menu 
            Caption         =   "All on"
         End
         Begin VB.Menu keysolo_menu 
            Caption         =   "Solo "
         End
      End
      Begin VB.Menu patch_menu 
         Caption         =   "patch_menu"
         Begin VB.Menu patch_audition_item 
            Caption         =   "Audition Patches"
         End
         Begin VB.Menu dblclick_audition_item 
            Caption         =   "Double Click to Audition Patches"
         End
         Begin VB.Menu mouse_notes_menu 
            Caption         =   "Send Mouse Notes"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu_sendnow 
            Caption         =   "Send program change now"
         End
         Begin VB.Menu bypass_pg_menu 
            Caption         =   "Don't Send Program Change"
         End
         Begin VB.Menu modify_mnu 
            Caption         =   "Modify"
            Begin VB.Menu modify_cc_mnu 
               Caption         =   "Controller Value"
            End
            Begin VB.Menu modify_patch_mnu 
               Caption         =   "Patch Number"
            End
            Begin VB.Menu renamepatch_menu 
               Caption         =   "Patch Name"
            End
            Begin VB.Menu mnu_clearpatch 
               Caption         =   "Clear all Patch data for this channel"
            End
         End
         Begin VB.Menu mnu_enteralt 
            Caption         =   "Edit Alternative Patch"
         End
         Begin VB.Menu mnualternative 
            Caption         =   "Send Alternative Patch"
         End
      End
      Begin VB.Menu ccmenu 
         Caption         =   "ccmenu"
      End
      Begin VB.Menu chanmenu 
         Caption         =   "chanmenu"
         Begin VB.Menu selectpatch 
            Caption         =   "Select Patch"
         End
         Begin VB.Menu Notecolor 
            Caption         =   "Note Color"
         End
         Begin VB.Menu clear_channelsettings_menu 
            Caption         =   "Clear Channel Settings"
         End
      End
      Begin VB.Menu solomenu 
         Caption         =   "solomenu"
         Begin VB.Menu solo_channel 
            Caption         =   "Solo"
            Index           =   0
         End
         Begin VB.Menu un_solo 
            Caption         =   "Unsolo"
            Index           =   0
         End
         Begin VB.Menu all_on 
            Caption         =   "All on"
         End
         Begin VB.Menu all_off 
            Caption         =   "All off"
         End
      End
   End
   Begin VB.Menu window 
      Caption         =   "Window"
      Begin VB.Menu midifileplayer 
         Caption         =   "Midi File Player"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuChangepatch 
         Caption         =   "Change Patch"
         Shortcut        =   ^C
      End
      Begin VB.Menu inputsoutputs 
         Caption         =   "Input/Output Po&rts"
         Shortcut        =   ^R
      End
      Begin VB.Menu module_menu 
         Caption         =   "Add/Delete Controls"
      End
      Begin VB.Menu midimonitor_menu 
         Caption         =   "Midi Monitor"
      End
      Begin VB.Menu songchart_menu 
         Caption         =   "Song Chart"
      End
      Begin VB.Menu mcicommand_menu 
         Caption         =   "MCI Scratchpad"
      End
      Begin VB.Menu preferences_menu 
         Caption         =   "Options..."
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu controllers_menu 
      Caption         =   "&Controllers"
      Begin VB.Menu ccitem 
         Caption         =   "no contollers configured"
         Index           =   0
      End
   End
End
Attribute VB_Name = "zmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private key_offset As Integer
Private keyisdown As Boolean
Const KNOB_HEIGHT = 420
Const KNOB_X_OFFSET = 200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const MAX_NOTE = 96
' piano constants
Const BLACKKEY_COLOR = 0
Const WHITEKEY_COLOR = &HFFFFFF
Const WHITEKEY_DOWN = &HC0C0C0
Const BLACKKEY_DOWN = &H808080
Const wwidth = 317
Const bwidth = 175
Const wheight = 1935
Const bheight = 1100
Const bottomkey = 36
Const topkey = 96

Dim htarget As Long
Private num_drumkeys As Integer
Private drumkeycode(MAX_NOTE), drumnote(MAX_NOTE)
Private notedown(MAX_NOTE) As Boolean

Private num_passthrus As Integer
Dim save_input As Integer, save_device As Integer, note As Integer
Dim show_extended_functions As Boolean
Dim mdown As Boolean
Dim b4solo(16) As Integer

Dim control_ref As Variant
Dim ctl As Object
Public curInput As Long       ' current midi device
Dim curDevice As Long       ' current midi device
Dim CurKeyID As Long          ' remember note on for note off
Dim numInputs As Long      ' number of midi output devices
Dim rc As Long              ' return code
Dim midimsg As Long         ' midi output message buffer
Dim baseNote As Integer     ' the first note on our "piano"
Dim prev_y As Long
Dim caps As MIDIOUTCAPS
Dim incaps As MIDIINCAPS
'Dim song_path, song_filename As String
Dim full_filename As String, short_filename As String
Private mixdir As String
Dim chancolor(16) As Long

' see if this transmits extended bit!
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const VK_SNAPSHOT As Byte = &H2C
Const VK_NEXT As Byte = &H22
Const WM_CHAR As Long = &H102
Const WM_KEYDOWN As Long = &H100
Const WM_KEYUP = &H101

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Type POINTAPI
      X As Long
      Y As Long
End Type
   
Private stayy As Integer
Dim n As POINTAPI

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long
Dim lHandle As Long

Private Sub about_menu_Click()
  MsgBox APP_NAME + " " + BUILD_DATE + vbCrLf + vbCrLf + _
  "shortcuts: transpose down/up: ^1/^2 " + vbCrLf + _
  "transpose down/up octave: ^3/^4 " + vbCrLf + _
  "increase/decrease channel: alt-1/2" + vbCrLf + _
  "disconnect ports temporarily: ^W"
End Sub

Private Sub bypass_pg_menu_Click()
  bypass_pg(user_chan) = Not bypass_pg(user_chan)
End Sub

Private Sub Form_Resize()
  If optionsform.chkRearrange Then
'MsgBox "wrr 2 before hide"
    Call hide_modules
'MsgBox "wrr 2 before reveal"
    Call reveal_modules
'MsgBox "wrr 4 after reveal"
  End If
End Sub

Private Sub mnu_clearpatch_Click()
  patch_pg(user_chan) = -1
  patch_label(user_chan).Caption = ""
  patch_msb(user_chan) = -1
  patch_lsb(user_chan) = -1
  For i = 0 To num_ccs - 1
    ccval(i, user_chan) = -1
  Next
End Sub

Private Sub filenew_menu_Click()
  Call clearsettings
End Sub

Private Sub kbutton_Click(Index As Integer)
  If ignore_channel(Index) > 0 Then
    ignore_channel(Index) = 0
  Else
    ignore_channel(Index) = 1
  End If
  
  If ignore_channel(Index) Then
    kbutton(Index).Visible = False
    klabel(Index).Visible = True
  Else
    kbutton(Index).Visible = True
    klabel(Index).Visible = False
  End If
  hidden_btn.SetFocus
End Sub

Private Sub kbutton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  user_chan = Index
  If Shift > 0 Then
    Call setcolor_menu_Click
  Else
    If Button = 2 Then PopupMenu keyboard_menu
  End If
End Sub

Private Sub all_on_menu_Click()
  For i = 0 To 15
    ignore_channel(i) = 0
  Next
End Sub

Private Sub keysolo_menu_Click()
  Dim black As Boolean
  For i = bottomkey To topkey
    idx = i - bottomkey + 1
    note = (idx - 1) Mod 12
    black = CBool(Choose(note + 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 0))
    If black Then
      pkey(idx).BackColor = BLACKKEY_COLOR
    Else
      pkey(idx).BackColor = WHITEKEY_COLOR
    End If
  Next
  
  For i = 0 To 15
    ignore_channel(i) = MAX_IGNORE
  Next
  ignore_channel(user_chan) = 0
  
End Sub

Private Sub klabel_Click(Index As Integer)
  kbutton_Click (Index)
End Sub

Private Sub klabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  user_chan = Index
  If Button = 2 Then PopupMenu keyboard_menu
End Sub

Private Sub midimonitor_menu_Click()
  midimon.Show
End Sub


Private Sub mnu_enteralt_Click()
  tmp_str = InputBox("Enter Alternative Patch (-1 to turn off)", "Alternative Patch")
  If tmp_str = "" Then Exit Sub
  alt_patch(user_chan) = CInt(tmp_str)
End Sub

Private Sub mnu_keyfader_Click()
  mnu_keyfader.Checked = Not mnu_keyfader.Checked
  fadermode = mnu_keyfader.Checked
End Sub
  
Private Sub mnu_remap_Click()
  mnu_remap.Checked = False
  remap_data_slider = False
  tmp_str = InputBox("enter Min value", "remap")
  If tmp_str = "" Then Exit Sub
  tmp_str2 = InputBox("enter Max value (and ensure PLAYBACK mode)", "remap")
  If tmp_str2 = "" Then Exit Sub
  cc_interval = (CInt(tmp_str2) - CInt(tmp_str)) \ 20
  mnu_remap.Checked = True
  remap_data_slider = True
  showmsg "interval: " + CStr(cc_interval)
End Sub

Private Sub mnu_sendalt_Click()
  For i = 0 To 15
    If alt_patch(i) <> -1 Then
      midisend PROGRAM_CHANGE + i, alt_patch(i), -1
    End If
  Next
End Sub

Private Sub mnu_senddummy_Click()
    midisend CONTROLLER_CHANGE, 60, 0
End Sub

Private Sub mnu_sendnow_Click()
  If patch_pg(user_chan) <> -1 Then
    If patch_msb(user_chan) <> -1 Then midisend CONTROLLER_CHANGE + user_chan, 0, patch_msb(user_chan)
    If patch_lsb(user_chan) <> -1 Then midisend CONTROLLER_CHANGE + user_chan, 32, patch_lsb(user_chan)
    If patch_pg(user_chan) <> -1 Then midisend PROGRAM_CHANGE + user_chan, patch_pg(user_chan), -1
  End If
End Sub


Private Sub mnualternative_Click()
  If alt_patch(user_chan) <> -1 Then
    midisend PROGRAM_CHANGE + user_chan, alt_patch(user_chan), -1
  End If
End Sub


Private Sub mnuChangepatch_Click()
  patchfrm.Show
End Sub

Private Sub mnuTempDisconnect_Click()
  ' temporarily disconnect ports if needed by another app
  Dim savIn As Integer, savOut As Integer
  savIn = curInput
  savOut = curDevice
  
  minput_Click (0)
  device_Click (0)
  MsgBox "Click to reconnect"
  minput_Click (savIn + 1)
  device_Click (savOut + 2)
End Sub

Private Sub mnuTransmitSettings_Click()
  Call sendsettings
  showmsg "settings transmission complete"
End Sub

Private Sub modify_cc_mnu_Click()
  If active_cc_idx = -1 Then
    MsgBox "no active controller selected"
  Else
    tmp_str = InputBox("Enter Controller Value (Cancel to clear value)", "Modify Controller Value")
    If tmp_str = "" Then tmp_str = "-1"
    ccval(active_cc_idx, user_chan) = CInt(tmp_str)
  End If
End Sub

Private Sub modify_patch_mnu_Click()
  tmp_str = InputBox("Enter patch Number (Cancel to clear value)", "Modify Patch Number")
  If tmp_str = "" Then tmp_str = "-1"
  patch_pg(user_chan) = CInt(tmp_str)
End Sub

Private Sub mru_menu_Click()
  passed_filename = mru_menu.Caption
  Call loadsongsettings_Click
End Sub

' kwrr deleted the menu option - so delete this function next
Private Sub multibank_mode_menu_Click()
  multibank_mode = Not multibank_mode
  multibank_mode_menu.Checked = Not multibank_mode_menu.Checked
End Sub

Private Sub preferences_menu_Click()
  optionsform.Show
End Sub

Private Sub renamepatch_menu_Click()
  tmp_str = InputBox("Enter patch name", "Rename Patch", patch_label(user_chan).Caption)
  If Len(tmp_str) > 0 Then patch_label(user_chan).Caption = tmp_str
End Sub

Private Sub setcolor_menu_Click()
  CommonDialog1.Flags = &H1& Or &H4&
  CommonDialog1.ShowColor
  chancolor(user_chan) = CommonDialog1.color
  kbutton(user_chan).BackColor = CommonDialog1.color
  klabel(user_chan).BackColor = CommonDialog1.color
  blob(user_chan).FillColor = CommonDialog1.color
End Sub
  
Private Sub songchart_menu_Click()
  sectionform.Show
End Sub


Private Sub channel_label_DblClick(Index As Integer)
  change_user_chan Index
  patchfrm.Show
End Sub

Private Sub dblclick_audition_item_Click()
dblclick_audition_item.Checked = Not dblclick_audition_item.Checked
End Sub

Private Sub Form_KeyUp(keycode As Integer, Shift As Integer)
  Dim note As Integer
  On Error Resume Next
  
  note = tonote(keycode, False, Shift)
  If note <> -100 Then
    If Shift = 4 Then note = note + 12
    If Shift = 2 Then note = note - 12
    If notedown(note) Then
      ShowNote note + IIf(drumkeymode, 0, notebase), 0, user_chan
      If drumkeymode Then
        midisend NOTE_OFF + user_chan, note, 80
      Else
        midisend NOTE_OFF + user_chan, notebase + note + transpose_amt, 80
      End If
    Else
      For note = 1 To MAX_NOTE
        If notedown(note) Then
          ShowNote note + IIf(drumkeymode, 0, notebase), 0, user_chan
          midisend NOTE_OFF + user_chan, notebase + note + transpose_amt, 80
          Exit For
        End If
      Next
      
    End If
'    showmsg "note: " + str(note)
    notedown(note) = False
  End If
End Sub

Private Function tonote(keycode As Integer, down As Boolean, Shift As Integer)
  tonote = -100
  If Not keypress_mode Then Exit Function
  
  If drumkeymode Then
    For i = 0 To num_drumkeys - 1
      If keycode = drumkeycode(i) Then tonote = drumnote(i)
    Next
  Else
    Select Case keycode
      Case vbKeyF1:
        If down Then
          If Shift = 1 Then
            transpose_amt = transpose_amt - 12
          Else
            transpose_amt = transpose_amt - 1
          End If
'          transposeform.Label2.Caption = CStr(transpose_amt)
          zmain.Show
          ringing_note = notebase + transpose_amt
          midisend NOTE_ON + user_chan, ringing_note, 80
          transposenotetimer.Enabled = True
          showmsg "transposing " + CStr(transpose_amt)
        End If
      Case vbKeyF2:
        If down Then
          If Shift = 1 Then
            transpose_amt = transpose_amt + 12
          Else
            transpose_amt = transpose_amt + 1
          End If
'          transposeform.Label2.Caption = CStr(transpose_amt)
          zmain.Show
          ringing_note = notebase + transpose_amt
          midisend NOTE_ON + user_chan, ringing_note, 80
          transposenotetimer.Enabled = True
          showmsg "transposing " + CStr(transpose_amt)
        End If
  '    Case vbKey1: If Shift = 1 Then Call change_user_chan(user_chan - 1) ' 1 to decrease user_chan
  '    Case vbKey2: If Shift = 1 Then Call change_user_chan(user_chan + 1) ' 2 to increase user_chan
        Case 192:
    '  Case vbKeyTab: tonote = 0  why doesnt tab keyup trigger keyup event?
        Case vbKey1: tonote = 1
      Case vbKeyQ: tonote = 2
        Case vbKey2: tonote = 3
      Case vbKeyW: tonote = 4
        Case vbKey3:
      Case vbKeyE: tonote = 5
        Case vbKey4: tonote = 6
      Case vbKeyR: tonote = 7
        Case vbKey5: tonote = 8
      Case vbKeyT: tonote = 9
        Case vbKey6: tonote = 10
      Case vbKeyY: tonote = 11
        Case vbKey7:
      Case vbKeyU: tonote = 12
        Case vbKey8: tonote = 13
      Case vbKeyI: tonote = 14
        Case vbKey9: tonote = 15
      Case vbKeyO: tonote = 16
        Case vbKey0:
      Case vbKeyP: tonote = 17
        Case 189: tonote = 18
      Case 219: tonote = 19
        Case 187: tonote = 20
      Case 221: tonote = 21
        Case vbKeyBack: tonote = 22
      Case vbKeyReturn: tonote = 23
         
      Case vbKeyDelete: tonote = 11
      Case vbKeyHome: tonote = 13
      Case vbKeyEnd: tonote = 12
      Case vbKeyPageDown:
      Case vbKeyNumpad7:
      Case vbKeyNumpad8:
      Case vbKeyNumpad9:
      Case vbKeyAdd:
    End Select
  End If
End Function

Private Function scalenote(keycode As Integer, down As Boolean, Shift As Integer)
  Dim scalekey As Integer

  
  If Shift = 2 Then key_offset = key_offset - 1
  If Shift = 4 Then key_offset = key_offset + 1
  
  Select Case keycode
    Case vbKeyQ:
    Case vbKeyA: scalekey = 0
    Case vbKeyW:
    Case vbKeyS: scalekey = 1
    Case vbKeyE:
    Case vbKeyD: scalekey = 2
    Case vbKeyR:
    Case vbKeyF: scalekey = 3
    Case vbKeyT:
    Case vbKeyG: scalekey = 4
    Case vbKeyY:
    Case vbKeyH: scalekey = 5
    Case vbKeyU:
    Case vbKeyJ: scalekey = 6
    Case vbKeyI:
    Case vbKeyK: scalekey = 7
    Case vbKeyO:
    Case vbKeyL: scalekey = 8
    Case vbKeyP:
    Case 186:    scalekey = 9 ' ;
    Case 219: ' [
    Case 222:    scalekey = 10 ' '
    Case 221: ' ]
  End Select
  
  scalekey = scalekey + key_offset
  
  'scalekey = scalekey Mod 12
  Select Case scalekey
    Case 0: scalenote = 60
    Case 1: scalenote = 62
    Case 2: scalenote = 64
    Case 3: scalenote = 65
    Case 4: scalenote = 67
    Case 5: scalenote = 69
    Case 6: scalenote = 71
    Case 7: scalenote = 72
  End Select

End Function

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
  Dim note As Integer, velocity As Integer
  On Error Resume Next
  zmain.stat.Caption = ""

  ' transpose for ctrl-1/2  ctrl-alt-1/2
  If ((keycode = vbKey1) Or (keycode = vbKey2)) And ((Shift = 2) Or (Shift = 3)) Then
    If keycode = vbKey1 And Shift = 2 Then transpose_amt = transpose_amt - 1
    If keycode = vbKey2 And Shift = 2 Then transpose_amt = transpose_amt + 1
    If keycode = vbKey1 And Shift = 3 Then transpose_amt = transpose_amt - 12
    If keycode = vbKey2 And Shift = 3 Then transpose_amt = transpose_amt + 12
    transposing = True ' maybe eliminate this boolean
    
    ' maybe have option to enable/disable sending tonic note
    ringing_note = notebase + transpose_amt
    midisend NOTE_ON + user_chan, ringing_note, 80
    transposenotetimer.Enabled = True

    showmsg "transpose: " + CStr(transpose_amt)
  End If
  
  If Shift = 4 Then
    If keycode = vbKey1 Then
      change_user_chan user_chan - 1
    ElseIf keycode = vbKey2 Then
      change_user_chan user_chan + 1
    End If
  End If

  
  If fadermode Then
    If keycode = vbKeyL Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 127
    ElseIf keycode = vbKeyK Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 120
    ElseIf keycode = vbKeyJ Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 110
    ElseIf keycode = vbKeyH Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 100
    ElseIf keycode = vbKeyG Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 80
    ElseIf keycode = vbKeyF Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 60
    ElseIf keycode = vbKeyD Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 40
    ElseIf keycode = vbKeyS Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 20
    ElseIf keycode = vbKeyA Then
      midisend CONTROLLER_CHANGE + user_chan, FADERCHAN, 0
    End If
    
    Exit Sub
  End If
  
  note = tonote(keycode, True, Shift)
  GetCursorPos n ' use vert mouse position to calculate velocity
  velocity = (800 - n.Y) \ 6
  If velocity > 127 Then velocity = 127
  
  If note <> -100 Then
    If Not notedown(note) Then
   Debug.Print "before its " & str(note)
   
      If Shift = 4 Then note = note + 12
      If Shift = 2 Then note = note - 12
      If drumkeymode Then
        midisend NOTE_ON + user_chan, note, velocity
      Else
  Debug.Print "its " & str(note) + " notebase: " + str(notebase)
        midisend NOTE_ON + user_chan, notebase + note + transpose_amt, velocity
      End If
      notedown(note) = True
    End If
    
    ShowNote note + IIf(drumkeymode, 0, notebase), 1, user_chan

  End If
  
  If optionsform.passthru_keys.Value = 1 Then
      If keycode = switch_key Then ' switch focus to sequencer
        Debug.Print "switchkey to switch to: " + appTitle(active_appnum)
        keycode = 0
      Else
        For i = 0 To num_passthrus - 1
          If passthru_in(i) <> 0 And keycode = passthru_in(i) Then
            AppActivateByStringPart appTitle(active_appnum)
            SendKeys passthru_out(i), True
            SetForegroundWindow Me.hwnd
            keycode = 0
          End If
        Next
      End If
  End If
End Sub

Private Sub Form_Load()
  Dim numInputs As Long
  Dim numrows As Integer
  Dim chan As Integer
  
  If Command$ <> "" Then passed_filename = Replace5(Command$, """", "")

  popups.Visible = False
  Me.Width = 11700
  'Me.Width = 4000 ' smaller for debug mode

  shifted = False
  
  sentshiftedkey = False
  lastplaykey = ""
  transpose_amt = 0
  
  scale_mode = 1
  active_cc_idx = -1
  lo_note = LOWEST_NOTE
  hi_note = HIGHEST_NOTE
  curDevice = -2
  curInput = -1
  'mciout = -2
  pkey(0).Visible = False

  Load Devicefrm

  row_ht(CHANNEL) = 240
  x_offset(CHANNEL) = 240
  descrip(CHANNEL) = "Channel label"
  
  row_ht(PATCH) = 540
  x_offset(PATCH) = 0
  descrip(PATCH) = "Patch title"
  
  row_ht(METER) = 660
  x_offset(METER) = 240
  descrip(METER) = "Polyphony Meter"
  
  row_ht(KEYBOARD_SELECT) = 0
  x_offset(KEYBOARD_SELECT) = 250
  descrip(KEYBOARD_SELECT) = "Keyboard Selection Buttons"
  
  row_ht(KEYBOARD_DESELECT) = 280
  x_offset(KEYBOARD_DESELECT) = 250
  descrip(KEYBOARD_DESELECT) = "Keyboard DeSelection Button"
  
  row_ht(PARAMETER) = 300
  x_offset(PARAMETER) = 164
  y_offset(PARAMETER) = 5
  descrip(PARAMETER) = "Parameter Window"
  
  row_ht(ANIMATION) = 470
  x_offset(ANIMATION) = 257
  y_offset(ANIMATION) = 30
  descrip(ANIMATION) = "Note Activity Display"
  
  row_ht(RECEIVE_SWITCH) = 220
  x_offset(RECEIVE_SWITCH) = 175
  y_offset(RECEIVE_SWITCH) = 0
  descrip(RECEIVE_SWITCH) = "Receive switch"
  
  '***** THE CONTROLS BELOW do not have 16 elements
  ' But when adding elements above, the below need to be renumbered in the BAS module
  
  row_ht(KEYBOARD) = 2580
  x_offset(KEYBOARD) = 0
  descrip(KEYBOARD) = "Keyboard"
  
  row_ht(OLDKEYBOARD) = 580
  x_offset(OLDKEYBOARD) = 1800
  descrip(OLDKEYBOARD) = "OldKeyboard"
  
  row_ht(STATLABEL) = 250
  x_offset(STATLABEL) = 50
  descrip(STATLABEL) = "Status Bar"

  control_ref = Array(blobtimer, channel_label, patch_label, bar, param, blob, ch, kbutton, klabel)
 
  For i = 0 To MAX_MODULES - 1
    y_offset(i) = 0
  Next i

  'create all controls
  For i = 1 To UBound(control_ref)
    If i <> 99 Then
      Set ctl = control_ref(i)
      For chan = 1 To 15
        alt_patch(chan) = -1
        patchbuffer(chan) = "    "
        Load ctl(chan)
      Next chan
    End If
  Next i
  
  device(0).Caption = "None"
  
  Load device(1) ' add to menu
  device(1).Caption = "Midi Mapper"
  Devicefrm.mixer_out.AddItem "None"
  Devicefrm.mixer_out.AddItem "Midi Mapper"
  Devicefrm.mci_out.AddItem "None"
  Devicefrm.mci_out.AddItem "Midi Mapper"
  
  ' Get the rest of the midi devices
  numDevices = midiOutGetNumDevs()
  For i = 0 To (numDevices - 1)
    Load device(i + 2) ' add to menu
    midiOutGetDevCaps i, caps, Len(caps)
    device(i + 2).Caption = caps.szPname
    
    Devicefrm.mixer_out.AddItem caps.szPname
    Devicefrm.mci_out.AddItem caps.szPname
  Next
  
  minput(0).Caption = "None"
  numInputs = midiInGetNumDevs()
  Devicefrm.mixer_in.AddItem "None"
  Devicefrm.in2.AddItem "None"
   
  For i = 0 To (numInputs - 1)
    Load minput(i + 1)
    midiInGetDevCaps i, incaps, Len(incaps)
    minput(i + 1).Caption = incaps.szPname
    minput(i + 1).Visible = True
    minput(i + 1).Enabled = True
    Devicefrm.mixer_in.AddItem incaps.szPname
    Devicefrm.in2.AddItem incaps.szPname
  Next

  device_Click (0) ' none
  minput_Click (0) ' none
  
  zmain.Caption = APP_NAME
  CurChannel = 0
  record_mode = True
  mixdir = "."
  bankdir = App.Path
  mru_menu.Visible = False
  active_appnum = -1
  
  Call read_ini
  
  For i = 0 To num_presets - 1
    par = Split5(presets(i), ",")
    If i > 0 Then Load preset_menu_item(i)
    preset_menu_item(i).Caption = par(0)
  Next

  Call reveal_modules
  If multibank_mode Then
    Load patchfrm
  End If

  Call clearsettings
  optionsform.chkRearrange = 1
  If passed_filename <> "" Then loadsongsettings_Click
  patch_label(0).BackColor = CHANNEL_SELECTED
End Sub



Private Sub clearsettings()
  For i = 0 To num_ccs - 1
    For j = 0 To 15
      ccval(i, j) = -1
    Next
  Next
  
  For i = 0 To 15
    patch_label(i).Caption = ""
    param(i).Caption = ""
    patch_msb(i) = -1
    patch_lsb(i) = -1
    patch_pg(i) = -1
    bypass_pg(i) = False
    ignore_channel(i) = 0
    chancolor(i) = WHITEKEY_DOWN
    alt_patch(i) = -1
  Next
  
  full_filename = ""
  short_filename = ""
  zmain.Caption = APP_NAME
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim mHdr As MIDIHDR
  
  showmsg "Exiting..."
  
  DoEvents
  If dump_enabled Then Call midiInUnprepareHeader(hMidiIn, mHdr, Len(mHdr))
  If optionsform.save_settings.Value = 1 Then Call write_ini
  showmsg "closing output port"
  device_Click (0)
  showmsg "closing input port"
  minput_Click (0)
  If thruports_open Then Close_thruports
  DoEvents
  DoEvents
  For i = Forms.Count - 1 To 0 Step -1
    showmsg "unloading " + Forms(i).Name
    Unload Forms(i)
  Next i
  showmsg "Program Closed"
  End

End Sub

 Sub MakePiano(pic As PictureBox)
  Dim wX1 As Long, wY1 As Long
  Dim wdX As Long, wdY As Long
  Dim zX1 As Long, zY1 As Long
  Dim zdX As Long, zdY As Long
  Dim AaWTs As Long                   ' count white keys
  Dim i As Long                       ' counter
  
  wX1 = 0: wY1 = 0: wdX = 7: wdY = 22 ' witte toets
  zX1 = 5: zY1 = 0: zdX = 4: zdY = 16 ' zwarte

  AaWTs = (128 / 12) * 7

  pic.Width = AaWTs * wdX * 15
  pic.AutoRedraw = True
  
  ' make 1st white key & copy other white keys
  pic.Line (wX1, wY1)-Step(wdX, wdY), QBColor(15), BF
  pic.Line (wX1, wY1)-Step(wdX, wdY), QBColor(0), B
  For i = 0 To AaWTs - 1
     BitBlt pic.hDC, wX1 + i * wdX, wY1, wdX, wdY + 1, pic.hDC, wX1, wY1, SRCCOPY
  Next i
     
  ' 1st black & copy other
  pic.Line (zX1, zY1)-Step(zdX, zdY), QBColor(0), BF
  For i = 1 To AaWTs - 1
     If Mid("110111", (i Mod 7) + 1, 1) = "1" Then
        BitBlt pic.hDC, zX1 + i * wdX, zY1, zdX + 1, zdY, pic.hDC, zX1, zY1, SRCCOPY
        End If
  Next i

  pic.Line (pic.ScaleWidth - 1, wY1)-Step(0, wdY), QBColor(0)
  pic.Picture = pic.Image
  pic.AutoRedraw = False
End Sub

Public Sub OldShowNote(ByVal Nr As Long, OnOff As Long)
  Dim octave As Long, note As Long, bw As Long
  Dim X As Long, Y As Long, s As Long
  Dim color As Long
  
  octave = (Nr \ 12)
  note = Nr Mod 12
  bw = Choose(note + 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 0) ' black or white
  X = octave * 49 + Choose(note + 1, 0, 3, 7, 10, 14, 21, 24, 28, 31, 35, 38, 42, 49)
  If bw = 1 Then
     Y = 11: X = X + 3: s = 2 ' black key
     color = IIf(OnOff = 1, QBColor(15), 0)
     'color = IIf(OnOff = 1, blob(chan).FillColor, 0)
  Else
     Y = 17: X = X + 2: s = 3 ' white key
     'color = IIf(OnOff = 1, 0, blob(chan).FillColor)
     color = IIf(OnOff = 1, 0, QBColor(15))
  End If
  picKlav.ForeColor = color
  picKlav.FillColor = color
  picKlav.Line (X, Y)-Step(s, s), color, BF
End Sub

Private Sub allnotesoff_Click()

  midiOutReset (hMidiOut)
  
  For i = 0 To 15
    midisend CONTROLLER_CHANGE + i, 121, CByte(0)
    midisend CONTROLLER_CHANGE + i, 123, CByte(0)
  Next i
End Sub

Sub write_ini()
  showmsg "writing ini..."
  Open App.Path + "\midimix.ini" For Output As #1
  
  Print #1, "[general]"
  Print #1, "record_mode=" + CStr(Abs(CInt(record_mode)))
  Print #1, "remap_mode=" + CStr(Abs(CInt(remap_mode)))
  Print #1, "update_with_midi=" + CStr(Abs(CInt(update_with_midi)))
  Print #1, "update_patchnames=" + CStr(Abs(CInt(update_patchnames)))
  'Print #1, "save_settings=" + CStr(optionsform.save_settings.Value)
  Print #1, "show_inactive=" + CStr(Abs(CInt(show_inactive)))
  Print #1, "passthru_keys=" + CStr(Abs(CInt(optionsform.passthru_keys.Value)))
  Print #1, "dblclick_edit=" + CStr(CInt(dblclick_audition_item.Checked))
  Print #1, "song=" + songfile
  If Len(full_filename) Then
    Print #1, "mru=" + full_filename
  ElseIf Len(mru_menu.Caption) Then
    Print #1, "mru=" + mru_menu.Caption
  End If
  Print #1, ""
  Print #1, "[ports]"
  Print #1, "midi_in=" + CStr(curInput + 1)
  Print #1, "midi_out=" + CStr(curDevice + 2)
  'Print #1, "mci_out=" + CStr(mciout)
  
  Print #1, ""
  Print #1, "[modules]"
  For i = 0 To num_modules - 1
    Print #1, "module=" + CStr(visible_modules(i))
  Next i
  
  Print #1, ""
  Close #1
  showmsg "initialization written"
End Sub

Sub read_ini()
  Dim param
  Dim initfile As Integer
  Dim initfilename(2) As String
  
  initfilename(0) = App.Path + "\midimix.ini"
  initfilename(1) = App.Path + "\mmdata.ini"
  
  For initfile = 0 To 1
  showmsg "reading " & initfilename(initfile)
  
  Open initfilename(initfile) For Input As #1

  Do While Not EOF(1)
    Line Input #1, tmp_str
    
    If Mid(tmp_str, 1, 1) = " " Or Len(tmp_str) = 0 Or Mid(tmp_str, 1, 1) = "#" Then
    ElseIf Mid(tmp_str, 1, 11) = "[invisible]" Then  'Mid(tmp_str, 1, 1) = "[" Or
      invisible_controllers = True
    Else
      param = Split5(tmp_str, "#") ' filter out end-of-line comments
      param = Split5(param(0), "=")
      
      If param(0) = "record_mode" Then
        record_mode = CBool(param(1))
        If record_mode Then
          Call optionsform.optRecord_Click
        Else
          Call optionsform.optPlayback_Click
        End If

      ElseIf param(0) = "dblclick_edit" Then
        dblclick_audition_item.Checked = CBool(param(1))
      ElseIf param(0) = "drumkey" Then
        param = Split5(param(1), ",")
        drumkeycode(num_drumkeys) = CInt(param(0))
        drumnote(Inc(num_drumkeys)) = CInt(param(1))
      ElseIf param(0) = "update_with_midi" Then
        update_with_midi = CBool(param(1))
        If update_with_midi Then update_with_midi_menu_Click
      ElseIf param(0) = "show_extended_functions" Then
        show_extended_functions = CBool(param(1))
      ElseIf param(0) = "update_patchnames" Then
        update_patchnames = CBool(param(1))
        update_patchnames_menu.Checked = update_patchnames
      ElseIf param(0) = "switch_key" Then
        switch_key = CInt(param(1))
      ElseIf param(0) = "remap_mode" Then
        remap_mode = CInt(param(1))
      ElseIf param(0) = "remap" Then
        param = Split5(param(1), ",")
        map_statusbyte(num_maps) = Val("&h" + param(0))
        map_statusbyte(num_maps) = map_statusbyte(num_maps) And &HF0 ' disregard the channel
        map_data1(num_maps) = Val("&h" + param(1))
        map_data2(num_maps) = Val("&h" + param(2))
        map_key(num_maps) = param(3)
        num_maps = num_maps + 1

      ElseIf param(0) = "passthru_key" Then
        param = Split5(param(1), ",")
        passthru_in(num_passthrus) = Asc(UCase(param(0)))
        If UBound(param) > 0 Then param = Split5(param(1), ",")
        passthru_out(num_passthrus) = param(0)
        num_passthrus = num_passthrus + 1
      ElseIf param(0) = "passthru_app" Then
        active_appnum = active_appnum + 1
        appTitle(active_appnum) = param(1)
        optionsform.appnames.AddItem appTitle(active_appnum)

'      ElseIf param(0) = "goto_key" Then
'        goto_key = param(1)
'      ElseIf param(0) = "followup_key" Then
'        followup_key = param(1)
'      ElseIf param(0) = "passthru_keys" Then
'        optionsform.passthru_keys.Value = CInt(param(1))
      ElseIf param(0) = "song" Then
        songfile = param(1)
      ElseIf param(0) = "mixdir" Then
        mixdir = param(1)
      ElseIf param(0) = "bankdir" Then
        bankdir = param(1)
      ElseIf param(0) = "mru" Then
        spacer1.Visible = True
        mru_menu.Visible = True
        mru_menu.Caption = param(1)
      ElseIf param(0) = "save_settings" Then
        optionsform.save_settings.Value = CInt(param(1))
      ElseIf param(0) = "show_inactive" Then
        show_inactive = CBool(param(1))
      ElseIf param(0) = "midi_out" Then
        device_Click (Val(param(1)))
      ElseIf param(0) = "midi_in" Then
        minput_Click (Val(param(1)))
      ElseIf param(0) = "preset" Then
        If num_presets < MAX_PRESETS Then
          presets(Inc(num_presets)) = param(1)
        Else
          MsgBox "Max presets exceeded - " & str(MAX_PRESETS)
        End If
      
      ElseIf param(0) = "doubleclick_time" Then
        doubleclick_time = (Val(param(1)))
        firstnote_timer.Interval = doubleclick_time
      ElseIf param(0) = "module" Then
        If Val(param(1)) < MAX_MODULES Then
          visible_modules(num_modules) = Val(param(1))
          num_modules = num_modules + 1
        Else
          MsgBox "invalid module in ini file"
        End If
      ElseIf param(0) = "cc" Then
        param = Split5(param(1), ",")
        If Not invisible_controllers Then
          If num_ccs > 0 Then Load ccitem(num_ccs)
          ccitem(num_ccs).Caption = param(1)
          ccitem(num_ccs).Visible = True
          ccitem(num_ccs).Enabled = True
          num_visible_ccs = num_visible_ccs + 1
        End If
        cctypes(num_ccs) = Val(param(0))
        num_ccs = num_ccs + 1
        
      ElseIf param(0) = "favorites" Then
        param = Split5(param(1), ",")
        bank_filename(num_banks) = param(0)
        bank_descrip(num_banks) = param(1)
        bank_msb(num_banks) = -2 ' cc0/32 is -2 for favorites banks
        bank_lsb(num_banks) = -2
        num_banks = num_banks + 1
      ElseIf param(0) = "bank" Then
        If num_banks < MAX_BANKS Then
          param = Split5(param(1), ",")
          bank_descrip(num_banks) = param(0)
          bank_filename(num_banks) = param(1)
          ' -1 means dont send msb (or lsb)
          If UBound(param) > 1 Then
            bank_msb(num_banks) = Val(param(2))
          Else
            bank_msb(num_banks) = -1
          End If
          If UBound(param) > 2 Then
            bank_lsb(num_banks) = Val(param(3))
          Else
            bank_lsb(num_banks) = -1
          End If
          num_banks = num_banks + 1
        Else
          MsgBox "max # of patch banks exceeded (" + CStr(MAX_BANKS) + ")"
        End If
      Else
      End If
    End If
  Loop
  Close #1

  Next initfile

' delete this
'  If record_mode Then
'    recordmode_menu_Click
'  Else
'    playmode_menu_Click
'  End If

  If Not show_extended_functions Then
    sep1.Visible = False
    mnu_senddummy.Visible = False
    mnu_remap.Visible = False
    mnu_keyfader.Visible = False
  End If
  
  If active_appnum > -1 Then optionsform.appnames.ListIndex = 0
  showmsg "initialized..."
End Sub


Public Sub reveal_modules()
  Dim y_position As Integer, rownum As Integer
  Dim CHANWIDTH As Integer
  Dim xoffset As Integer
  'CHANWIDTH = 720
  CHANWIDTH = (Me.Width - 120) \ 16 '180

  y_position = 20
  For rownum = 0 To num_modules - 1

    If visible_modules(rownum) = ANIMATION Then animating = True
    If visible_modules(rownum) = RECEIVE_SWITCH Then Call all_on_Click
    If visible_modules(rownum) = METER Then meter_on = True
    
    If visible_modules(rownum) = OLDKEYBOARD Then
      picKlav.Left = 1440
      picKlav.Visible = True
      picKlav.Top = y_position
      showkeyboard = True
          
    ElseIf visible_modules(rownum) = KEYBOARD Then
      drawPiano
      showkeyboard = True
    
    ElseIf visible_modules(rownum) = STATLABEL Then
      stat.Visible = True
      stat.Top = y_position
        
    Else:
      Set ctl = control_ref(visible_modules(rownum))
      For chan = 0 To 15
        xoffset = (CHANWIDTH - ctl(chan).Width) \ 2
        If (xoffset < 0) Or (visible_modules(rownum) = PATCH) Then xoffset = 0
        If visible_modules(rownum) = PATCH Then
          xoffset = 0
          If optionsform.chkSinglelinename Then
            row_ht(visible_modules(rownum)) = 240
          Else
            row_ht(visible_modules(rownum)) = 540
          End If
        End If
        
        ctl(chan).Left = (CHANWIDTH * chan) + xoffset
        ctl(chan).Top = y_position + y_offset(visible_modules(rownum))
        If visible_modules(rownum) = ANIMATION Then
          If show_inactive Then
            ctl(chan).Visible = True
          Else
            ctl(chan).Visible = False
          End If
        ElseIf visible_modules(rownum) = KEYBOARD_DESELECT Then
            ctl(chan).Visible = False
        Else
          ctl(chan).Visible = True
        End If
      Next chan

    End If
    y_position = y_position + row_ht(visible_modules(rownum))
      
  Next rownum
  
  'change defaults
  For chan = 0 To 15
    playing(chan) = True
    
    Set ctl = control_ref(CHANNEL)
    ctl(chan).Caption = str(chan + 1)
    ctl(chan).Alignment = 2
    ctl(chan).Width = 255
   
    Set ctl = control_ref(PATCH)
'    ctl(chan).Width = 685
    ctl(chan).Width = CHANWIDTH '- 0 '35
    If optionsform.chkSinglelinename Then
      ctl(chan).Height = 240
    Else
      ctl(chan).Height = 505
    End If
    'label height for 1 line: 225, 2 lines:385, 3 lines: 585
    ctl(chan).Caption = patchbuffer(chan)
  Next chan
  If zmain.WindowState = 0 Then zmain.Height = y_position + 655  '705
End Sub

Public Sub device_Click(Index As Integer)
  Dim newdevice As Long
   
  MousePointer = HOURGLASS
  If Index <> curDevice + 2 Then
    device(curDevice + 2).Checked = False
     
    If curDevice <> -2 Then
      midiOutReset (hMidiOut)
      rc = midiOutClose(hMidiOut)
      If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiOUT_Close", rc
      hMidiOut = 0
    End If
  
    newdevice = Index - 2
    If newdevice <> -2 Then
      rc = midiOutOpen(hMidiOut, newdevice, 0, 0, 0)
      If rc <> MMSYSERR_NOERROR Then
        ShowMMErr "midiOUT_Open", rc
        newdevice = -2
      End If
    End If
    
    curDevice = newdevice
  End If
  
  device(curDevice + 2).Checked = True
  MousePointer = WINDEFAULT
End Sub

Private Sub mcicommand_menu_Click()
  mci_scratchpad.Show
End Sub

Public Sub minput_Click(Index As Integer)
  Dim midiError As Long, newinput As Long
  Dim mHdr As MIDIHDR

  MousePointer = HOURGLASS
  showmsg "initializing input port..."
  
  If Index <> curInput + 1 Then
    minput(curInput + 1).Checked = False
    
    If curInput <> -1 Then
      rc = midiInStop(hMidiIn)
      If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Open-Stop", rc
      
      rc = midiInClose(hMidiIn)
      If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Close", rc
    End If
    
    newinput = Index - 1
    If newinput <> -1 Then
        rc = midiInOpen(hMidiIn, newinput, AddressOf MidiIN_Proc, 0, CALLBACK_FUNCTION)
        If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiInOpen", rc
        
        If rc <> MMSYSERR_NOERROR Then
           ShowMMErr "midiIN_Open", rc
           newinput = -1
        Else
          ' NEW code to enable sysex hopefully
          If dump_enabled Then
Debug.Print "dump enabled"
            mHdr.lpData = sSysEx
            mHdr.dwBufferLength = LENMIDIHDR 'or Len(sSysEx)   or  sizeof(SysXBuffer);
            mHdr.dwFlags = 0
            
            rc = midiInPrepareHeader(hMidiIn, mHdr, Len(mHdr))
            If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiInPrepareHeader", rc
            
            rc = midiInAddBuffer(hMidiIn, mHdr, Len(mHdr))
            If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiInaddBuffer", rc
            MsgBox "midiInPrepareHeader and midiInAddBuffer were executed"
          Else
Debug.Print "dump not enabled"
            showmsg "NOT enabling sysex"
          End If
        
           midiError = midiInStart(hMidiIn)
           If midiError <> MMSYSERR_NOERROR Then
              ShowMMErr "midiIN_Open-Start", midiError
              newinput = -1
           End If
        End If
    End If
      
    curInput = newinput
  End If
  
  minput(curInput + 1).Checked = True
  showmsg "input port initialized..."
  MousePointer = WINDEFAULT
End Sub


Private Sub inputsoutputs_Click()
    Devicefrm.Show
End Sub

Private Sub midifileplayer_Click()
  mcifrm.Show
End Sub

Private Sub module_menu_Click()
  modulefrm.Show 1
End Sub

Private Sub mouse_notes_menu_Click()
  mouse_notes_menu.Checked = Not mouse_notes_menu.Checked
End Sub

Private Sub Notecolor_Click()
  CommonDialog1.Flags = &H1& Or &H4&
  CommonDialog1.ShowColor
  blob(user_chan).FillColor = CommonDialog1.color
End Sub

Private Sub ontop_Click()
  ontop.Checked = Not ontop.Checked
  If ontop.Checked Then
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Else
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub

Private Sub param_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  zmain.stat.Caption = ""
 
  If Button = 1 And active_cc_idx <> -1 Then
    If Not settings_changed Then indicated_changed_settings
    MousePointer = UPDOWN
    mdown = True
    GetCursorPos n
    stayy = n.Y
    user_chan = Index
  End If
End Sub

Private Sub param_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim amt As Integer, changed As Boolean
  amt = 1
  changed = True
  
  If Shift And &H2 Then amt = 10 ' ctrl key
  
  If mdown Then
    If Not settings_changed Then indicated_changed_settings
    If ccval(active_cc_idx, Index) = -1 Then
      ccval(active_cc_idx, Index) = 64
    Else
      
      GetCursorPos n
      If n.Y > stayy And ccval(active_cc_idx, Index) > 0 Then
        ccval(active_cc_idx, Index) = ccval(active_cc_idx, Index) - amt
      ElseIf n.Y < stayy And ccval(active_cc_idx, Index) < 127 Then
        ccval(active_cc_idx, Index) = ccval(active_cc_idx, Index) + amt
      Else
        changed = False
      End If
    End If
    If ccval(active_cc_idx, Index) > 127 Then ccval(active_cc_idx, Index) = 127
    If ccval(active_cc_idx, Index) < 0 Then ccval(active_cc_idx, Index) = 0
    
    If changed Then midisend CONTROLLER_CHANGE + user_chan, cctypes(active_cc_idx), ccval(active_cc_idx, Index)

    param(Index).Caption = str(ccval(active_cc_idx, user_chan))
    DoEvents
    SetCursorPos n.X, stayy
  End If

End Sub

Private Sub param_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If active_cc_idx = -1 Or Button = 2 Then
    PopupMenu controllers_menu, vbPopupMenuCenterAlign
  Else
    MousePointer = WINDEFAULT
    mdown = False
  End If
End Sub

Private Sub channel_label_Click(Index As Integer)
  change_user_chan Index
End Sub

Private Sub patch_audition_item_Click()
  patchfrm.Show
End Sub

Private Function change_user_chan(Index As Integer)

  If Index = -1 Then Index = 15 ' when changing channel from computer keyboard
  If Index = 16 Then Index = 0
  
  If user_chan = Index Then
    change_user_chan = False
  Else
    For i = 0 To 15
      patch_label(i).BackColor = &H80000004
    Next
    
    user_chan = Index
    patch_label(Index).BackColor = CHANNEL_SELECTED
    change_user_chan = True
  End If
End Function

Private Sub patch_label_DblClick(Index As Integer)
  If dblclick_audition_item.Checked Then patchfrm.Show
End Sub

Private Sub patch_label_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim col As Integer
  Dim chan_changed As Boolean
  
  oldchan = user_chan
  chan_changed = change_user_chan(Index)
  
  patchfrm.chanbox.ListIndex = user_chan
  
  If Button = 2 Then
    bypass_pg_menu.Checked = bypass_pg(user_chan)
    PopupMenu patch_menu, Index
  Else
    If chan_changed And dblclick_audition_item.Checked Then
      firstnote_timer.Enabled = True
      delayed_note = lo_note + CInt((X * (hi_note - lo_note) / 685))
      Exit Sub
    End If
    
    If mouse_notes_menu.Checked Then
      MousePointer = UPARROW
      note = lo_note + CInt((X * (hi_note - lo_note) / 685))
      midisend NOTE_ON + user_chan, note, 80
    End If
  End If
  zmain.stat.Caption = ""
End Sub
Private Sub firstnote_timer_Timer()
  If delayed_note <> 0 Then
    firstnote_timer.Enabled = False
    MousePointer = UPARROW
    note = delayed_note
    midisend NOTE_ON + user_chan, note, 80
  End If
End Sub


Private Sub patch_label_Mouseup(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  firstnote_timer.Enabled = False
  
  If Button = 1 And mouse_notes_menu.Checked Then
    delayed_note = 0
    MousePointer = WINDEFAULT
    If note <> 0 Then
      midisend NOTE_OFF + user_chan, note, 64
      note = 0
    End If
  End If
End Sub

Private Sub picKlav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim oct As Long
  Dim No As Long
  Dim mX As Single
  
  oct = X \ 49
  If picKlav.Point(X, Y) = 0 And Y < 17 Then
    mX = X - 4
    No = oct * 12 + Choose(((mX \ 7) Mod 7) + 1, 1, 3, 5, 6, 8, 10, 11)
  Else
    mX = X
    No = oct * 12 + Choose(((mX \ 7) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
  End If
  
  CurKeyID = No
  OldShowNote No, 1
    
  midisend NOTE_ON + user_chan, No, 120 'CurChannel
End Sub

Private Sub picKlav_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  midisend NOTE_OFF + user_chan, CurKeyID, 120 'CurChannel
  OldShowNote CurKeyID, 0
End Sub

'Private Sub playmode_menu_Click()
'  record_mode = False
'  playmode_menu.Checked = True
'  recordmode_menu.Checked = False
'  optionsform.optRemap(0).Enabled = False
'  optionsform.optRemap(1).Enabled = False
'  optionsform.optRemap(2).Enabled = False
'  optionsform.optRemap(3).Enabled = False
'End Sub

'Private Sub recordmode_menu_Click()
'  record_mode = True
'  playmode_menu.Checked = False
'  recordmode_menu.Checked = True
'  optionsform.optRemap(0).Enabled = True
'  optionsform.optRemap(1).Enabled = True
'  optionsform.optRemap(2).Enabled = True
'  optionsform.optRemap(3).Enabled = True
'End Sub

Private Sub saveini_menu_Click()
  Call write_ini
End Sub

Private Sub loadsongsettings_Click()
  Dim param
  Dim realcc As Integer
  
  Call clearsettings
  
  If passed_filename <> "" Then
    full_filename = passed_filename
    passed_filename = ""
  
  Else
    CommonDialog1.Filter = "Songsetting Files|*.mix"
    CommonDialog1.InitDir = mixdir
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.FileName) = 0 Then Exit Sub
    full_filename = CommonDialog1.FileName
  End If

  MousePointer = HOURGLASS
  
  param = Split5(full_filename, "\")
  short_filename = param(UBound(param))
  zmain.Caption = APP_NAME & " - " & short_filename
  
  Open full_filename For Input As #1
  Do While Not EOF(1)
    Line Input #1, tmp_str
  
    If Mid(tmp_str, 1, 1) = " " Or Len(tmp_str) = 0 Then
    Else
      If Mid(tmp_str, 1, 8) = "[channel" Then
        param = Split5(tmp_str, " ")
        param = Split5(param(1), "]")
        chan = CInt(param(0)) - 1
      Else
        param = Split5(tmp_str, "=")
        If param(0) = "patch" Then
          param = Split5(param(1), ",")
          If Len(param(0)) Then patch_msb(chan) = CInt(param(0))
          If Len(param(1)) Then patch_lsb(chan) = CInt(param(1))
          If Len(param(2)) Then patch_pg(chan) = CInt(param(2))
          patch_label(chan).Caption = param(3)
        ElseIf param(0) = "alternative" Then
          alt_patch(chan) = CInt(param(1))
        ElseIf param(0) = "cc" Then
          param = Split5(param(1), ",")
          realcc = CInt(param(0))
          For i = 0 To num_ccs - 1
            If cctypes(i) = realcc Then
              ccval(i, chan) = CInt(param(1))
              i = num_ccs + 1
            End If
          Next
          If i = num_ccs Then MsgBox "Controller #" + CStr(realcc) + " not configured in ini file"
        ElseIf param(0) = "bypass" Then
          If param(1) = "1" Then bypass_pg(chan) = True
          
          
        ElseIf param(0) = "enabled" Then
          If param(1) = "1" Then
            If ch(chan).BackColor = CH_OFF Then
              ch_Click (chan)
              Debug.Print "enabling " + str(chan)
            End If
          Else
            If ch(chan).BackColor = CH_ON Then
              ch_Click (chan)
            End If
          End If

        ElseIf param(0) = "alternative" Then
          alt_patch(chan) = CInt(param(1))
        
        ElseIf param(0) = "section" Then
          If num_songsections < MAX_SONGSECTIONS Then
            songsection(Inc(num_songsections)) = param(1)
          Else
            MsgBox "max songsections exceeded"
          End If
        End If
      End If
    End If
  Loop
  
  Close #1
  
'  Call sendsettings
  savesongsettings.Enabled = True
  MousePointer = WINDEFAULT
  showmsg "songsettings loaded"
End Sub

Private Sub sendsettings()
  Dim param
  For chan = 0 To 15
    If Not bypass_pg(chan) Then
      If patch_pg(chan) <> -1 Then
        If patch_msb(chan) <> -1 Then midisend CONTROLLER_CHANGE + chan, 0, patch_msb(chan)
        If patch_lsb(chan) <> -1 Then midisend CONTROLLER_CHANGE + chan, 32, patch_lsb(chan)
        If patch_pg(chan) <> -1 Then midisend PROGRAM_CHANGE + chan, patch_pg(chan), -1
      End If
      For i = 0 To num_ccs - 1
        If ccval(i, chan) <> -1 Then
          midisend CONTROLLER_CHANGE + chan, cctypes(i), ccval(i, chan)
        End If
      Next
    End If
  Next
  showmsg "transmission complete"
End Sub

Private Sub savesongsettings_Click()
  If full_filename = "" Then
    Call savesongsettingsas_Click
  Else
    Call save_songsettings
  End If
End Sub

Private Sub savesongsettingsas_Click()
  CommonDialog1.Filter = "Songsetting Files|*.mix"
  CommonDialog1.InitDir = mixdir
  CommonDialog1.ShowSave
  If Len(CommonDialog1.FileName) = 0 Then Exit Sub
  full_filename = CommonDialog1.FileName
  
  par = Split5(full_filename, "\")
  short_filename = par(UBound(par))
  Call save_songsettings
End Sub


Private Sub save_songsettings()
  MousePointer = HOURGLASS
  Dim channel_header As Boolean
  
  Open full_filename For Output As #1
  For chan = 0 To 15
    channel_header = False
    
    If patch_pg(chan) <> -1 Or Len(patch_label(chan).Caption) Then
      If Not channel_header Then
        Print #1, "[channel " + CStr(chan + 1) + "]"
        channel_header = True
      End If
      Print #1, "patch=" + CStr(patch_msb(chan)) + "," + CStr(patch_lsb(chan)) + _
        "," + CStr(patch_pg(chan)) + "," + patch_label(chan).Caption
      If ch(chan).BackColor = CH_OFF Then
        Print #1, "enabled=0"
      Else
        Print #1, "enabled=1"
      End If
      If alt_patch(chan) <> -1 Then Print #1, "alternative=" + CStr(alt_patch(chan))
    
    End If
    
    For i = 0 To num_ccs - 1

      If ccval(i, chan) <> -1 Then
        If Not channel_header Then
          Print #1, "[channel " + CStr(chan + 1) + "]"
          channel_header = True
        End If

        Print #1, "cc=" + CStr(cctypes(i)) + "," + CStr(ccval(i, chan))
      End If
    Next
    If bypass_pg(chan) Then Print #1, "bypass=1"
    If channel_header Then Print #1, ""
  Next
  
  Print #1, ""
  If num_songsections > 0 Then
    Print #1, "[songchart]"
    For i = 0 To num_songsections - 1
      Print #1, "section=" & songsection(i)
    Next
  End If

  Close #1
  MousePointer = WINDEFAULT
  showmsg "songsettings saved"
  
  settings_changed = False
  zmain.Caption = APP_NAME & " - " & short_filename
End Sub

Private Sub show_ccs_menu_Click()
  tmp_str = ""
  For i = 0 To num_visible_ccs - 1
    tmp_str = tmp_str + ccitem(i).Caption + " (" + CStr(cctypes(i)) + ") "
    For chan = 0 To 15
      tmp_str = tmp_str + CStr(ccval(i, chan))
    Next
    tmp_str = tmp_str + vbCrLf
  Next
  
  MsgBox tmp_str
End Sub

Private Sub solo_channel_Click(Index As Integer)
  For chan = 0 To 15
    b4solo(chan) = IIf(ch(chan).BackColor = CH_ON, 1, 0)
    playing(chan) = False
    ch(chan).BackColor = CH_OFF
    set_channel chan, 0
  Next

  ch(i).BackColor = CH_ON ' i was set in mouseup event
  set_channel i, 1
  playing(i) = True
End Sub

Private Sub channel_label_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    user_chan = Index
    PopupMenu chanmenu
  End If
End Sub

Private Sub selectpatch_Click()
  patchfrm.Show
End Sub

Private Sub blobtimer_Timer(Index As Integer)
  blobtimer(Index).Interval = 0
  blob(Index).Visible = False
End Sub

Public Sub hide_modules()
  animating = 0
  meter_on = 0

  For i = 0 To num_modules - 1
    If visible_modules(i) = OLDKEYBOARD Then
      picKlav.Visible = False
      showkeyboard = False
    ElseIf visible_modules(i) = KEYBOARD Then
      For j = bottomkey To topkey
        idx = j - bottomkey + 1
        Unload pkey(idx)
      Next
      showkeyboard = False
      ptop.Visible = False

    ElseIf visible_modules(i) = STATLABEL Then
      stat.Visible = False
    Else
      Set ctl = control_ref(visible_modules(i))
      For chan = 0 To 15
          ctl(chan).Visible = False
          If visible_modules(i) = PATCH Then patchbuffer(chan) = ctl(chan).Caption
      Next chan
    End If
  Next i
End Sub

Function enable_controls(ByVal setting As Boolean)
  For i = 0 To num_modules - 1
    Set ctl = control_ref(visible_modules(i))
    For chan = 0 To 15
        ctl(chan).Visible = setting
    Next chan
  Next i
End Function


Private Sub ccitem_Click(Index As Integer)
  For i = 0 To num_visible_ccs - 1
    ccitem(i).Checked = False
  Next
  
  active_cc_idx = Index
  ccitem(Index).Checked = True
  For i = 0 To 15
    param(i).ToolTipText = ccitem(Index).Caption
    
    If ccval(active_cc_idx, i) = -1 Then
      param(i).Caption = ""
    Else
      param(i).Caption = str(ccval(active_cc_idx, i))
    End If
  Next
  controllers_menu.Caption = "&Change " + ccitem(Index).Caption
    
End Sub

Private Sub file_exit_Click()
  Dim ret As Integer
  
  If settings_changed Then
    showmsg "songsettings have changed!"

    ret = MsgBox("Save songsettings?", vbYesNoCancel + vbQuestion + vbDefaultButton1)
    If ret = 2 Then Exit Sub
    If ret = 6 Then Call savesongsettings_Click
  End If
  
  Unload zmain
End Sub

Private Function send_syx(sSysEx As String)
  Dim mHdr As MIDIHDR
  Dim LenSysEx As Long

  LenSysEx = Len(sSysEx)
  With mHdr
    .lpData = sSysEx
    .dwBufferLength = LenSysEx
    .dwBytesRecorded = LenSysEx
    .dwUser = 0
    .dwFlags = 0
  End With
  midiOutPrepareHeader hMidiOut, mHdr, Len(mHdr)
  midiOutLongMsg hMidiOut, mHdr, Len(mHdr)
  mHdr.dwFlags = 0
  midiOutUnprepareHeader hMidiOut, mHdr, Len(mHdr)
End Function

Private Sub set_channel(ByVal chan As Integer, ByVal ch_status As Integer)
  Dim offset As Integer
  offset = chan + IIf(ch_status, 1, 0)
  tmp_str = Chr(&HF0) & Chr(&H41) & Chr(&H10) & Chr(&H6A) & Chr(&H12) & Chr(&H1) & Chr(&H0) & Chr(&H10 + chan) & Chr(&H0) & Chr(ch_status) & Chr(&H6F - offset) & Chr(&HF7)
  send_syx tmp_str
End Sub

Private Sub set_channel2(ByVal chan As Integer, ByVal ch_status As Integer)
  Dim offset As Integer
  Dim tmp_str2 As String
  tmp_str = "f0 41 10 6A 12 01 00 10+ 00 status 6F- F7"
  par = Split5(tmp_str, " ")
  tmp_str2 = ""
  Dim byteval As Integer
  
  For i = 0 To UBound(par)
    byteval = Val("&h" + Mid(par(i), 1, 2))
    
    If Len(par(i)) = 3 Then
      If Mid(par(i), 3, 1) = "+" Then
        byteval = byteval + chan
      ElseIf Mid(par(i), 3, 1) = "-" Then
        byteval = byteval - chan
      End If
    End If
              
    tmp_str2 = tmp_str2 & Chr(byteval)
  Next
  
  'tmp_str2 = Chr(&HF0) & Chr(&H41) & Chr(&H10) & Chr(&H6A) & Chr(&H12) & Chr(&H1) & Chr(&H0) & Chr(&H10 + chan) & Chr(&H0) & Chr(ch_status) & Chr (&H6F - offset) & Chr(&HF7)
  
  send_syx tmp_str2
  minput_Click (0)
End Sub


Private Sub ch_Mouseup(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  i = Index ' use to get index when selecting solo channel
  If Button = 2 Then PopupMenu solomenu
End Sub

Private Sub transposenotetimer_Timer()
  midisend NOTE_OFF + user_chan, ringing_note, 80
  transposenotetimer.Enabled = False
End Sub

Private Sub un_solo_Click(Index As Integer)
  For i = 0 To 15
    ch(i).BackColor = IIf(b4solo(i), CH_ON, CH_OFF)
    playing(i) = IIf(b4solo(i), True, False)
    set_channel i, IIf(b4solo(i), 1, 0)
  Next
End Sub
Private Sub ch_Click(Index As Integer)
  If Len(Dir("c:\jkj")) > 0 Then MsgBox "channel status " + IIf(ch(Index).BackColor = CH_ON, "on", "off")

  ch(Index).BackColor = IIf(ch(Index).BackColor = CH_ON, CH_OFF, CH_ON)
  playing(Index) = Not playing(Index)
  set_channel Index, IIf(ch(Index).BackColor = CH_ON, 1, 0)
  hidden_btn.SetFocus
End Sub
  
Private Sub all_off_Click()
  For chan = 0 To 15
    ch(chan).BackColor = CH_OFF
    playing(chan) = False
    b4solo(chan) = False
    set_channel chan, 0
  Next
End Sub

Private Sub all_on_Click()
  For chan = 0 To 15
    ch(chan).BackColor = CH_ON
    playing(chan) = True
    b4solo(chan) = True
    set_channel chan, 1
  Next
End Sub

Public Sub connect_thruports(in_index As Integer, out_index As Integer)
  Dim mHdr As MIDIHDR
  MousePointer = HOURGLASS
 
  thru_input = in_index - 1
  thru_output = out_index - 2
  
  midiOutGetDevCaps thru_output, caps, Len(caps)
  
  midiInGetDevCaps thru_input, incaps, Len(incaps)
  
  Debug.Print "thru_connect actual input #" + CStr(thru_input) + " (" + incaps.szPname
  Debug.Print " to output #" + CStr(thru_output) + " (" + caps.szPname
  
  rc = midiOutOpen(hMidiOut2, thru_output, 0, 0, 0)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiOutOpen", rc
  
  rc = midiInOpen(hMidiIn2, thru_input, 0, 0, 0)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiIn_Open", rc
  

  If dump_enabled Then
    MsgBox "preparing for sysex!"
    mHdr.lpData = sSysEx
    
    '/* Store its size in the MIDIHDR */
    mHdr.dwBufferLength = LENMIDIHDR
    
    '/* Flags must be set to 0 */
    mHdr.dwFlags = 0
    
    'err = midiInPrepareHeader(handle, &midiHdr, sizeof(MIDIHDR));
    rc = midiInPrepareHeader(hMidiIn2, mHdr, Len(mHdr))  ' or LENMIDIHDR - is it a  long?
    If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiInPrepareHeader", rc
    
    rc = midiInAddBuffer(hMidiIn2, mHdr, Len(mHdr))
    If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiInaddBuffer", rc
  Else
    MsgBox "NOT enabling sysex"
  End If
  

  rc = midiInStart(hMidiIn2)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Start", rc
  
  rc = midiConnect(hMidiIn2, hMidiOut2, 0)
  If rc <> MMSYSERR_NOERROR Then
    ShowMMErr "midiConnect", rc
  Else
    thruports_open = True
    Devicefrm.connect_btn.Enabled = False
    Devicefrm.disconnect_btn.Enabled = True
  End If
  MousePointer = WINDEFAULT
End Sub

Public Sub Close_thruports()
  Devicefrm.in2.ListIndex = 0
  Devicefrm.mci_out.ListIndex = 0
  stat.Caption = "Closing Thruports"
 
  ' close input
  rc = midiDisconnect(hMidiIn2, hMidiOut2, 0)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiDisconnect", rc
  
  ' close output
  rc = midiOutReset(hMidiOut2)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiOutReset", rc
  
  rc = midiOutClose(hMidiOut2)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiOutClose", rc

  rc = midiInStop(hMidiIn2)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Open-Stop", rc
  
  rc = midiInClose(hMidiIn2)
  If rc <> MMSYSERR_NOERROR Then ShowMMErr "midiIN_Close", rc
  stat.Caption = "thruports are Closed"
  Devicefrm.connect_btn.Enabled = True
  Devicefrm.disconnect_btn.Enabled = False

  thruports_open = False

End Sub

Private Sub swap_ports_menu_Click()
  Dim newCurInput As Long, newCurDevice As Long, newThruInput As Long, newThruOutput As Long
  
  MousePointer = HOURGLASS
  newCurInput = thru_input
  newCurDevice = thru_output
  newThruInput = curInput
  newThruOutput = curDevice
  
  minput_Click (0) ' none
  device_Click (0) ' none
  Call Close_thruports

  minput_Click (newCurInput + 1)
  device_Click (newCurDevice + 2)
  
  Call connect_thruports(newThruInput + 1, newThruOutput + 2)
'  recordmode_menu.Checked = Not recordmode_menu.Checked
'  playmode_menu.Checked = Not playmode_menu.Checked
  record_mode = Not record_mode
  If record_mode Then
    Call optionsform.optRecord_Click
  Else
    Call optionsform.optPlayback_Click
  End If
  MousePointer = WINDEFAULT
End Sub

Public Function mci_send_command(ByVal cmd As String)
  tmp = mciSendString(cmd, tmp_str, 255, 0)

  If tmp Then ' error
    Call mciGetErrorString(tmp, tmp_str, 255)
    MsgBox tmp_str
  End If
End Function

' kwrr deleted the menu option - so delete this function next
Private Sub update_patchnames_menu_Click()
  update_patchnames_menu.Checked = Not update_patchnames_menu.Checked
  update_patchnames = update_patchnames_menu.Checked
  If update_patchnames Then Load patchfrm ' SHOULD BE IF FORM NOT LOADED YET!
End Sub


Private Sub drawPiano()
Dim octave As Integer, note As Integer, black As Integer, offset As Integer
For i = bottomkey To topkey
  idx = i - bottomkey + 1

  octave = ((idx - 1) \ 12)
  note = (idx - 1) Mod 12
  black = CBool(Choose(note + 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 0))
  offset = Choose(note + 1, 0, 0, 1, 1, 2, 3, 3, 4, 4, 5, 5, 6)

  Load pkey(idx)
  pkey(idx).Top = 0
  pkey(idx).Visible = True
  If Not black Then
    pkey(idx).Left = (octave * 7 * wwidth) + (offset * wwidth)
    pkey(idx).BackColor = WHITEKEY_COLOR ' &HFFFFFF
    pkey(idx).Height = wheight
    pkey(idx).Width = wwidth
  Else
    pkey(idx).BackColor = BLACKKEY_COLOR '&H80000008
    pkey(idx).Height = bheight
    pkey(idx).Left = wwidth - (bwidth / 2) + (octave * 7 * wwidth) + (offset * wwidth)
    pkey(idx).ZOrder 0
    pkey(idx).Width = bwidth
  End If

Next

  ptop.Width = wwidth * 7 * ((topkey - bottomkey) \ 12) + (wwidth) - 80
  ptop.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not showkeyboard Then Exit Sub
  If Button <> 1 Then Exit Sub
  If Y > wheight Then Exit Sub
  Dim oct, No, offset As Integer
  Dim mX As Single
  
  oct = X \ (7 * wwidth) '49
  'No = bottomkey + (oct * 12) + Choose(((X \ wwidth) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
  
  If Y < bheight Then ' if its black, and above bottom boundary of black keys
'    If zmain.Point(X, Y) <> &HFFFFFF Then
    If zmain.Point(X, Y) = BLACKKEY_COLOR Then
      mX = X - (bwidth \ 2) '4
      offset = Choose(((mX \ wwidth) Mod 7) + 1, 1, 3, -1, 6, 8, 10, -1)
      If offset <> -1 Then No = bottomkey + oct * 12 + offset 'Choose(((mX \ wwidth) Mod 7) + 1, 1, 3, 5, 6, 8, 10, 11)
    End If
  Else
    mX = X
    No = bottomkey + (oct * 12) + Choose(((mX \ wwidth) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
  End If
  
  If No = 0 Then Exit Sub
  
  If No <> CurKeyID Then
    ShowNote CurKeyID, 0, user_chan
    midisend NOTE_OFF + user_chan, CurKeyID, 120 'CurChannel
    CurKeyID = No
    ShowNote No, 1, user_chan
    midisend NOTE_ON + user_chan, No, 120 'CurChannel
  End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim oct As Integer, No As Integer
  Dim mX As Single
  If Not showkeyboard Then Exit Sub
  If Y > wheight Then Exit Sub
  oct = X \ (7 * wwidth) '49
  
  If zmain.Point(X, Y) = BLACKKEY_COLOR And Y < bheight Then ' if its black, and above bottom boundary of black keys
    mX = X - (bwidth \ 2) '4
    No = bottomkey + oct * 12 + Choose(((mX \ wwidth) Mod 7) + 1, 1, 3, 5, 6, 8, 10, 11)
  Else
    mX = X
    No = bottomkey + (oct * 12) + Choose(((mX \ wwidth) Mod 7) + 1, 0, 2, 4, 5, 7, 9, 11)
  End If

  CurKeyID = No
  ShowNote No, 1, user_chan
  midisend NOTE_ON + user_chan, No, 120 'CurChannel

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not showkeyboard Then Exit Sub
  If Y > wheight Then Exit Sub
  midisend NOTE_OFF + user_chan, CurKeyID, 120
  ShowNote CurKeyID, 0, user_chan
  CurKeyID = 0
End Sub

Public Sub ShowNote(ByVal note_num As Long, key_down As Long, chan As Integer)
  idx = note_num - bottomkey + 1
  Dim black As Boolean
  On Error GoTo NONOTE
  black = CBool(Choose((note_num - bottomkey) Mod 12 + 1, 0, 1, 0, 1, 0, 0, 1, 0, 1, 0, 1, 0))

If key_down Then
  If black Then
    pkey(idx).Height = bheight + 50
    If chancolor(chan) = WHITEKEY_DOWN Then ' if default then use darker color for black down
      pkey(idx).BackColor = BLACKKEY_DOWN
    Else
    pkey(idx).BackColor = chancolor(chan) '&HE0E0E0
    End If
  Else
    pkey(idx).BackColor = chancolor(chan)    '&HE0E0E0
    ' for some reason next statement results in exception
    pkey(idx).Height = wheight + 50
  End If
Else
  If black Then
    pkey(idx).BackColor = BLACKKEY_COLOR
    pkey(idx).Height = bheight
  Else
    pkey(idx).BackColor = WHITEKEY_COLOR
    pkey(idx).Height = wheight
  End If
End If
Exit Sub

NONOTE:
End Sub
  
Public Sub indicated_changed_settings()
  settings_changed = True
  If full_filename = "" Then
    zmain.Caption = APP_NAME & " *"
  Else
    zmain.Caption = APP_NAME & " - " & short_filename & " *"
  End If
End Sub


Private Sub preset_menu_item_Click(Index As Integer)
  Call hide_modules
  par = Split5(presets(Index), ",")
  
  For i = 1 To UBound(par)
    visible_modules(i - 1) = CInt(par(i))
  Next
  
  num_modules = UBound(par)
  
  Call reveal_modules
End Sub

Public Sub showmsg(str As String)
  msg.Caption = str
  msg.Width = Len(str) * 134
  msg.Visible = True
  Timer1.Enabled = False
  Timer1.Interval = 1000
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  msg.Visible = False
  Timer1.Enabled = False
End Sub

Private Sub Command2_Click()

'  If htarget = 0 Then
    htarget = FindWindow(vbNullString, "SONAR - [SONAR1* - Track]")
    MsgBox "set htarget to " + str(htarget)
'  End If

  If htarget <> 0 Then
      PostMessage htarget, WM_KEYDOWN, Asc(" "), &H1&
      PostMessage htarget, WM_KEYUP, Asc(" "), &HC0010001
  End If
End Sub

' kwrr deleted the menu option - so delete this function next
Private Sub update_with_midi_menu_Click()
  patchfrm.Hide
  update_with_midi_menu.Checked = Not update_with_midi_menu.Checked
  update_with_midi = Not update_with_midi
  update_patchnames_menu.Checked = update_with_midi
  update_patchnames = update_with_midi
End Sub
