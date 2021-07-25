VERSION 5.00
Begin VB.Form optionsform 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   4380
   ClientLeft      =   5355
   ClientTop       =   3600
   ClientWidth     =   4980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Playback Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   2415
      Begin VB.CheckBox chkUpdatePatchnames 
         Caption         =   "Update Patchnames"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkUpdateWithMidi 
         Caption         =   "Update with midi input"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkSearchMultiple 
         Caption         =   "Search Multiple Bank Files"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Midi to Keypress Mapping"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
      Begin VB.OptionButton optRemap 
         Caption         =   "Step Time"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optRemap 
         Caption         =   "Keyboard Remote"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optRemap 
         Caption         =   "Remap"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optRemap 
         Caption         =   "Disabled"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Mode 
      Caption         =   "Mode"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton optRecord 
         Caption         =   "Record"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPlayback 
         Caption         =   "Playback"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkSinglelinename 
      Caption         =   "Single Line Patchnames"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CheckBox chkRearrange 
      Caption         =   "Rearrange upon resize"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox keypress 
      Caption         =   "&Keypresses play notes"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox chkDrumkeys 
      Caption         =   "Drumkeys"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox save_settings 
      Caption         =   "Save Settings Upon Exit"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.CheckBox showinactive_check 
      Caption         =   "Show Inactive Notes"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.CheckBox sysex_check 
      Caption         =   "Enable Sysex"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox appnames 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CheckBox passthru_keys 
      Caption         =   "pass keystrokes thru"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label stat 
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
End
Attribute VB_Name = "optionsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub appnames_Click()
  active_appnum = appnames.ListIndex
  'MsgBox "active appnum: " + str(active_appnum)
End Sub


Private Sub chkDrumkeys_Click()
  drumkeymode = CBool(Check1.Value)
End Sub

Private Sub chkSearchMultiple_Click()
  multibank_mode = chkSearchMultiple.Value
End Sub

Private Sub chkSinglelinename_Click()
    Call zmain.hide_modules
    Call zmain.reveal_modules
End Sub

Private Sub chkUpdatePatchnames_Click()
  update_patchnames = chkUpdatePatchnames.Value
End Sub

Private Sub chkUpdateWithMidi_Click()
  'patchfrm.Hide
  update_with_midi = chkUpdateWithMidi.Value
  If chkUpdateWithMidi Then
    chkUpdatePatchnames.Enabled = True
    chkSearchMultiple.Enabled = True
    update_patchnames = chkUpdateWithMidi.Value
  Else
    chkUpdatePatchnames.Enabled = False
    chkSearchMultiple.Enabled = False
  End If

End Sub

Private Sub keypress_Click()
  keypress_mode = CBool(keypress.Value)
End Sub

Private Sub Form_Paint()
  optRemap(remap_mode) = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
  Call set_remap
  Me.Hide
End Sub

Public Sub optPlayback_Click()
  record_mode = False
  optRemap(0).Enabled = False
  optRemap(1).Enabled = False
  optRemap(2).Enabled = False
  optRemap(3).Enabled = False
End Sub

Public Sub optRecord_Click()
  record_mode = True
  optRemap(0).Enabled = True
  optRemap(1).Enabled = True
  optRemap(2).Enabled = True
  optRemap(3).Enabled = True
End Sub

Private Sub showinactive_check_Click()
  If showinactive_check.Value Then
    show_inactive = True
  Else
    show_inactive = False
  End If

  For i = 0 To 15
    zmain.blob(i).Visible = show_inactive
  Next i
End Sub

Private Sub sysex_check_Click()
  If sysex_check.Value Then
    dump_enabled = True
    midiInReset (hMidiIn)
  Else
    dump_enabled = False
  End If
End Sub

Private Sub set_remap()
  If optRemap(0) Then
    remap_mode = 0
  ElseIf optRemap(1) Then
    remap_mode = 1
  ElseIf optRemap(2) Then
    remap_mode = 2
  ElseIf optRemap(3) Then
    remap_mode = 3
  End If
End Sub

Private Sub btnClose_Click()
  Call set_remap
  Me.Hide
End Sub

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
  If keycode = 27 Or keycode = vbKeyReturn Then Unload Me
  stat.Caption = "keycode: " + str(keycode)
End Sub

Private Sub Form_Load()
'  Load patchfrm
  If show_inactive Then showinactive_check = 1
End Sub



