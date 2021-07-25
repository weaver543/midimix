VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Begin VB.Form mcifrm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Midi File Player"
   ClientHeight    =   5100
   ClientLeft      =   6165
   ClientTop       =   4785
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin MCI.MMControl MMControl1 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   327680
      DeviceType      =   "sequencer"
   End
   Begin VB.CommandButton loadsong 
      BackColor       =   &H00C6C3C6&
      Height          =   495
      Left            =   240
      Picture         =   "mcifrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox mciport 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton rewind_btn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      Picture         =   "mcifrm.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton pause_btn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      Picture         =   "mcifrm.frx":0734
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton stop_btn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      Picture         =   "mcifrm.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      TickStyle       =   3
      TickFrequency   =   0
   End
   Begin VB.CommandButton fwd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      Picture         =   "mcifrm.frx":0AEC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton back 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      Picture         =   "mcifrm.frx":0BB6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton playbtn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Picture         =   "mcifrm.frx":0C80
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2400
      Top             =   2280
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1560
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "."
   End
   Begin VB.Label songlength 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label current_pos 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "mcifrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SKIP_AMOUNT = 100
Private update_slider As Boolean

Private Sub back_Click()
  Dim p
  'to seek backward in a file
  p = MMControl.position - SKIP_AMOUNT
  current_pos.Caption = str(p)
  
  If p < 0 Then p = 0
  MMControl.From = p
  MMControl.Command = "play"
End Sub


Private Sub fwd_Click()
  'to seek forward in a file
  p = MMControl.position + SKIP_AMOUNT
  current_pos.Caption = str(p)
  If p > MMControl.length Then p = MMControl.length - 1
  MMControl.Command = "close"
  MMControl.Command = "open"
  MMControl.From = p
  MMControl_PlayClick (0)
  MMControl.Command = "play"
End Sub


Private Sub Form_Load()
  For i = 0 To Devicefrm.mixer_out.ListCount - 1
    mciport.AddItem Devicefrm.mixer_out.List(i)
  Next
  
  mciport.ListIndex = zmain.curInput + 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MMControl.Command = "close"
  Unload Me
End Sub


Private Sub loadsong_Click()
  On Error GoTo ENDIT
   
  CommonDialog1.Filter = "Midi Files|*.mid"
  CommonDialog1.ShowOpen
  If Len(CommonDialog1.FileName) = 0 Then Exit Sub
  
  If songfile <> "" Then MMControl.Command = "close"

  songfile = CommonDialog1.FileName
  songfile = RTrim(songfile)
  
  MMControl.FileName = songfile
  MMControl.Command = "open"
  songlength.Caption = str(MMControl.length)
  Slider1.Max = MMControl.length
  
  update_slider = True
  rewind_btn.Enabled = True
  back.Enabled = True
  playbtn.Enabled = True
  stop_btn.Enabled = True
  pause_btn.Enabled = True
  fwd.Enabled = True
  Slider1.Enabled = True
  
  
  Exit Sub
  
ENDIT:
  MsgBox "problem with MCI multimedia control"
  Unload Me

End Sub

Private Sub mciport_Click()
'  mciout = mciport.ListIndex
End Sub

Private Sub pause_btn_Click()
  MMControl.Command = "pause"
End Sub

Private Sub playbtn_Click()
' if stopped at end, reopen & start over
MsgBox "kwrr 3"
  If MMControl.position = MMControl.length Then
   MMControl.Command = "close"
   MMControl.Command = "open"
  End If

  MMControl_PlayClick (0)
  MMControl.Command = "play"
  zmain.stat.Caption = MMControl.Error
End Sub

Private Sub rewind_btn_Click()
' rewind doesnt work!  MMControl.Command = "rewind"
  MMControl.Command = "close"
  MMControl.Command = "open"
  zmain.stat.Caption = MMControl.Error
End Sub

Private Sub Slider1_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MMControl.Command = "stop"
  MMControl.Command = "close"
  MMControl.Command = "open"
  MMControl.From = Slider1.Value
  MMControl_PlayClick (0)
  MMControl.Command = "play"
  zmain.stat.Caption = MMControl.Error
  
  update_slider = True
End Sub
Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  update_slider = False
End Sub

Private Sub stop_btn_Click()
  MMControl.Command = "stop"
  zmain.stat.Caption = MMControl.Error
End Sub

Private Sub Timer1_Timer()
  current_pos.Caption = str(MMControl.position)
  If update_slider Then Slider1.Value = MMControl.position
End Sub

Private Sub MMControl_PlayClick(Cancel As Integer)
  Dim parms As MCI_SEQ_SET_PARMS
  Dim rc As Long
  Dim msg As String * 150

  parms.dwPort = mciport.ListIndex - 2
MsgBox "kwrr using port " + str(mciport.ListIndex)
  rc = mciSendCommand(MMControl.DeviceID, MCI_SET, MCI_SEQ_SET_PORT, parms)
MsgBox "kwrr opened sucessfully with rc: " + str(rc)
  If (rc <> NO_ERROR) Then
    mciGetErrorString rc, msg, Len(msg)
    MsgBox "Errcode " + str(rc) + ": " + msg
  End If
MsgBox "kwrr exiting"
End Sub

