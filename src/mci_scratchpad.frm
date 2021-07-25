VERSION 5.00
Begin VB.Form mci_scratchpad 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "MCI Scratchpad"
   ClientHeight    =   3900
   ClientLeft      =   4590
   ClientTop       =   3420
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox errbox 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   6255
   End
   Begin VB.TextBox statbox 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.TextBox cmdbox 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.ListBox history 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label4 
      Caption         =   "History"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Error Messages"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Status Messages"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Type MCI Commands here and press Enter"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "mci_scratchpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mci_playing As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
  Dim cmd As String
 
  If KeyAscii = 32 Then
    If mci_playing Then
      cmd = "pause m"
    Else
      cmd = "play m"
    End If
    mci_playing = Not mci_playing
  ElseIf KeyAscii = 119 Then
    cmd = "rewind m"
  End If
  
  zmain.mci_send_command (cmd)
End Sub

Private Sub Form_Load()
  history.AddItem "set port 1"
  history.AddItem "open c:\song.mid type sequencer alias m"
  history.AddItem "set m port 1"
  history.AddItem "play m"
  history.AddItem "stop m"
  history.AddItem "status m port"
  history.AddItem "status m tempo"
  history.AddItem "status m position"
  history.AddItem "pause m"
  history.AddItem "resume m"
  history.AddItem "close m"
  history.AddItem ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
  tmp = mciSendString("close m", 0&, 50, 0)
End Sub

Private Sub history_Click()
  cmdbox.Text = history.Text
End Sub


Private Sub history_DblClick()
  cmdbox_KeyPress (42)
End Sub

Private Sub cmdbox_KeyPress(KeyAscii As Integer)
  If KeyAscii <> vbKeyReturn And KeyAscii <> 42 Then Exit Sub
  'KeyAscii = 0 'Eat the Key to prevent a Beep
  
   If cmdbox.Text = "x" Then
     history.Height = 7000
     Exit Sub
   End If
   
   errbox.Text = ""
   statbox.Text = ""
   tmp = mciSendString(cmdbox.Text, tmp_str, 255, 0)
 
   statbox.Text = tmp_str
 
   If tmp Then ' error
       Call mciGetErrorString(tmp, tmp_str, 255)
       errbox.Text = tmp_str
       Debug.Print tmp_str
   End If

  If KeyAscii = vbKeyReturn Then history.AddItem cmdbox.Text

  cmdbox.Text = ""
  cmdbox.SetFocus
End Sub


