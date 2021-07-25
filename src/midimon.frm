VERSION 5.00
Begin VB.Form midimon 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Midi Monitor"
   ClientHeight    =   1590
   ClientLeft      =   7575
   ClientTop       =   4140
   ClientWidth     =   1335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   1335
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu midimon_menu 
      Caption         =   "midimon"
      Visible         =   0   'False
      Begin VB.Menu clear_menu 
         Caption         =   "Clear"
      End
      Begin VB.Menu clipboard_menu 
         Caption         =   "Enable Clipboard Controls"
      End
      Begin VB.Menu mode_mnu 
         Caption         =   "Hex Mode"
      End
      Begin VB.Menu ontop 
         Caption         =   "Always On Top"
      End
   End
End
Attribute VB_Name = "midimon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
  Text1.Width = midimon.Width - 130
  Text1.Height = midimon.Height - 380
End Sub

Private Sub ontop_Click()
  ontop.Checked = Not ontop.Checked
  If ontop.Checked Then
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Else
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub

Private Sub clipboard_menu_Click()
  Text1.Enabled = Not Text1.Enabled
  clipboard_menu.Checked = Not clipboard_menu.Checked
End Sub

Private Sub clear_menu_Click()
  Text1.Text = ""
End Sub

Private Sub Form_Load()
  midi_monitoring = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu midimon_menu
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Text1.Text = ""
  midi_monitoring = False
End Sub


Private Sub mode_mnu_click()
  midimon_dec_mode = Not midimon_dec_mode
  mode_mnu.Checked = midimon_dec_mode
End Sub
