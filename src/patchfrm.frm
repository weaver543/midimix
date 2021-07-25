VERSION 5.00
Begin VB.Form patchfrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patch Selection"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5760
      TabIndex        =   21
      Top             =   5280
      Width           =   375
   End
   Begin VB.ListBox bankname 
      Height          =   1425
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton sendchange 
      Caption         =   "Send PG"
      Height          =   495
      Left            =   9840
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton hidden_btn 
      Caption         =   "Command1"
      Height          =   495
      Left            =   13680
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox colornames 
      Caption         =   "Color Names"
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Top             =   4665
      Width           =   1695
   End
   Begin VB.CheckBox send_upon_change 
      Caption         =   "Send note upon patch change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.ComboBox scale_type 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox dblclickexit 
      Caption         =   "Double Click to exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6600
      TabIndex        =   14
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox cclsb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   12
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox ccmsb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   10
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox hinote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   9
      Top             =   4560
      Width           =   375
   End
   Begin VB.TextBox lonote 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   8
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox velocity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   5
      Text            =   "95"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton closeme 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox chanbox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5040
      Width           =   540
   End
   Begin VB.Label Label7 
      Caption         =   "delay"
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Low Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "LSB (cc32)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "High Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Channel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   3
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "MSB (cc0)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label box 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Virtual Piano"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "patchfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Const LEFT_OFFSET = 50
Const ROW_HEIGHT = 240
Const COL_WIDTH = 1400

Dim keyisdown As Boolean
Dim tmpbank As Integer
Dim note, prev_x, scalestage, scale_mode As Integer
Dim favorites As Boolean
Dim favemsb(128) As Integer
Dim favelsb(128) As Integer
Dim favepg(128) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Colornames_Click()
  colors = Array(&HC0C0FF, &HFFFFC0, &HFF80FF, &H8080FF, &HFFFF00, &H80, &H800080, &HC0C0FF, &HFFFFC0, &HFF80FF, &H8080FF, &HFFFF00, &H80)
  j = 0
  If colornames.Value = 1 Then
    For i = 0 To 127
      box(i).ForeColor = colors(j)
      j = j + 1
      If j = UBound(colors) Then j = 0
    Next
  Else
    For i = 0 To 127
      box(i).ForeColor = &H0
    Next
  End If
End Sub

Private Sub closeme_Click()
  Me.Hide
  'Unload Me
End Sub

Private Sub Form_KeyUp(keycode As Integer, Shift As Integer)
  keyisdown = False
  Call box_Mouseup(0, 0, 0, 0, 0)
End Sub

Private Sub Form_Load()
Debug.Print "wrr loading patchfrm"
' make textboxes 285 * 1335
  Dim idx, col, row As Integer
  
  increment = INCREMENT_AMT
  patchform_loaded = True
  lonote.Text = CStr(lo_note)
  hinote.Text = CStr(hi_note)

  For i = 0 To num_banks - 1
    bankname.AddItem bank_descrip(i)
  Next i
  
  For i = 0 To 15
    chanbox.AddItem CStr(i + 1)
    prognum(i) = -1
  Next i
  
  i = 0
  For col = 0 To 7
    Left = LEFT_OFFSET + (1395 * col)
    For row = 0 To 15
        If i > 0 Then Load box(i)
        box(i).Width = COL_WIDTH
        Top = LEFT_OFFSET + (row * ROW_HEIGHT)
        box(i).Top = Top
        box(i).Left = Left
        box(i).Visible = True
        box(i).ToolTipText = CStr(i)
        i = i + 1
     Next row
   Next col
   
  If num_banks > 0 Then bankname.ListIndex = bank(user_chan)
  
  scale_type.AddItem "chromatic"
  scale_type.AddItem "diatonic"
  scale_type.AddItem "pentatonic"
  scale_type.ListIndex = scale_mode
End Sub

Function sendlsb(ByVal Index As Integer)
  If favorites Then
    cclsb.Text = favelsb(Index)
  End If
  If Not CBool(Len(cclsb.Text)) Or Mid(cclsb.Text, 1, 1) = " " Then Exit Function

  midisend CONTROLLER_CHANGE + user_chan, 32, Val(cclsb.Text)
End Function

Function sendmsb(ByVal Index As Integer)
  If favorites Then ccmsb.Text = favemsb(Index)
  If Not CBool(Len(ccmsb.Text)) Or Mid(ccmsb.Text, 1, 1) = " " Then Exit Function
  
  midisend CONTROLLER_CHANGE + user_chan, 0, Val(ccmsb.Text)
End Function

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
  Dim Index As Integer
  If keycode = 27 Then Unload Me

  If keyisdown Then
    keycode = 0
  End If

  If Shift = 2 And keycode = vbKeyE Then
    If bankname.ListIndex > 0 Then bankname.ListIndex = bankname.ListIndex - 1
    keycode = 0
  ElseIf Shift = 2 And keycode = vbKeyD Then

  If bankname.ListIndex < num_banks - 1 Then bankname.ListIndex = bankname.ListIndex + 1
    keycode = 0
  ElseIf keycode = vbKeyDown Then
    Index = prognum(user_chan) + 1
    If Index > 127 Then
      Index = 0
      If bankname.ListIndex < num_banks - 1 Then bankname.ListIndex = bankname.ListIndex + 1
      bankname_Click
    End If
    keyisdown = True

  ElseIf keycode = vbKeyUp Then
    Index = prognum(user_chan) - 1
    If Index < 0 Then
      Index = 127
      If bankname.ListIndex > 0 Then bankname.ListIndex = bankname.ListIndex - 1
      bankname_Click
    End If
    keyisdown = True
  ElseIf keycode = vbKeyRight Then
    Index = prognum(user_chan) + 16
    If Index > 127 Then
      Index = 0
      If bankname.ListIndex < num_banks - 1 Then bankname.ListIndex = bankname.ListIndex + 1
      bankname_Click
    End If
    keyisdown = True
  ElseIf keycode = vbKeyLeft Then
    Index = prognum(user_chan) - 16
    If Index < 0 Then
      Index = Index + 128
      If bankname.ListIndex > 0 Then bankname.ListIndex = bankname.ListIndex - 1
      bankname_Click
    End If
    keyisdown = True
  End If
  
  If keycode > 36 And keycode < 41 Then
    Call box_MouseDown(Index, 1, 0, 810, 120)
    keycode = 0
  ElseIf keycode = 32 And prognum(user_chan) <> -1 Then ' Spacebar to play note
    Call box_MouseDown(prognum(user_chan), 1, 0, 810, 120)
    keycode = 0
    keyisdown = True
  End If

End Sub

Sub load_patchnames()
  Dim Text As String
  On Error GoTo WOOPS
  
  Open bankdir + "\" + bank_filename(tmpbank) For Input As #1
  
  For i = 0 To 127
    If Not EOF(1) Then
      If favorites Then
        Do
          Line Input #1, tmp_str
        Loop While Mid(tmp_str, 1, 1) = "#"
        param = Split5(tmp_str, ",")
        favemsb(i) = param(0)
        favelsb(i) = param(1)
        favepg(i) = param(2)
        
        box(i).Caption = param(3)
      Else
        Do
          Line Input #1, Text
        Loop While Mid(Text, 1, 1) = "#"
        box(i).Caption = Text
      End If
    Else
      box(i).Caption = ""
    End If
  Next i

  Close #1
  
  ccmsb.Text = ""
  cclsb.Text = ""
  If Not favorites Then
    ' -1 means dont send msb (or lsb)
    If bank_msb(tmpbank) <> "-1" Then ccmsb.Text = bank_msb(tmpbank)
    If bank_lsb(tmpbank) <> "-1" Then cclsb.Text = bank_lsb(tmpbank)
    If ccmsb.Text <> "" Then sendmsb (0)
    If cclsb.Text <> "" Then sendlsb (0)
  End If
  
  Exit Sub
  
WOOPS:
  MsgBox "ERROR: Missing bank filename specified: " + bank_filename(tmpbank)
End Sub

Private Sub bankname_Click()
  If prognum(user_chan) <> -1 And prognum(user_chan) <> -1 Then box(prognum(user_chan)).BackColor = PATCH_NOT_SELECTED
  tmpbank = bankname.ListIndex
  If tmpbank = bank(user_chan) And prognum(user_chan) <> -1 Then box(prognum(user_chan)).BackColor = PATCH_SELECTED
  If bank_msb(tmpbank) = "-2" Then
    favorites = True
  Else
    favorites = False
  End If
  
  Call load_patchnames
End Sub

Public Sub chanbox_Click()
  Dim i As Integer
  If oldchan <> -1 Then ' changed from zmain.patchbox_click()
    If prognum(oldchan) <> -1 Then box(prognum(oldchan)).BackColor = PATCH_NOT_SELECTED
    oldchan = -1
  Else
    For i = 0 To 15
      zmain.patch_label(i).BackColor = &H80000004
    Next
    If chanbox.ListIndex <> -1 Then user_chan = chanbox.ListIndex
    zmain.patch_label(user_chan).BackColor = CHANNEL_SELECTED
    If prognum(user_chan) <> -1 Then box(prognum(user_chan)).BackColor = PATCH_NOT_SELECTED
    
    For i = 0 To 127
      box(i).BackColor = PATCH_NOT_SELECTED
    Next
  End If
  
  bankname.ListIndex = bank(user_chan)
  If prognum(user_chan) <> -1 Then
    box(prognum(user_chan)).BackColor = PATCH_SELECTED
  Else
    For i = 0 To num_banks - 1
      If bank_msb(i) = patch_msb(user_chan) And bank_lsb(i) = patch_lsb(user_chan) _
      And bank_msb(i) <> -1 And bank_lsb(i) <> -1 Then
        prognum(user_chan) = patch_pg(user_chan)
        bank(user_chan) = i
        bankname.ListIndex = bank(user_chan)
        If prognum(user_chan) <> -1 Then box(prognum(user_chan)).BackColor = PATCH_SELECTED
      Else
      End If
      Debug.Print patch_msb(user_chan), bank_msb(i), patch_lsb(user_chan), bank_lsb(i)
    Next
  End If
End Sub

Private Sub box_DblClick(Index As Integer)
  If dblclickexit.Value = 1 Then Unload Me ' exit when checkbox is checked
End Sub

Private Sub box_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not settings_changed Then zmain.indicated_changed_settings
  scalestage = 1
  If Index <> prognum(user_chan) Then
    bank(user_chan) = tmpbank
    box(Index).BackColor = PATCH_SELECTED
    If prognum(user_chan) <> -1 Then box(prognum(user_chan)).BackColor = PATCH_NOT_SELECTED
    prognum(user_chan) = Index
    If favorites Then
      sendlsb (Index)
      sendmsb (Index)
      midiData1 = favepg(Index)

      patch_pg(user_chan) = favepg(Index)
    Else
      midiData1 = Index
      patch_pg(user_chan) = Index
    End If
    
    midisend PROGRAM_CHANGE + user_chan, midiData1
    
    If ccmsb.Text = "" Then
      patch_msb(user_chan) = -1
    Else
       patch_msb(user_chan) = CInt(ccmsb.Text)
    End If
    
    If cclsb.Text = "" Then
      patch_lsb(user_chan) = -1
    Else
       patch_lsb(user_chan) = CInt(cclsb.Text)
    End If
    
    patch_name(user_chan) = box(Index).Caption
    zmain.patch_label(user_chan).Caption = box(Index).Caption
    box(Index).BackColor = PATCH_SELECTED
    
    If Not send_upon_change.Value = 1 Then Exit Sub
    If Text1.Text <> "" Then Sleep CInt(Text1.Text)
  End If
   
  MousePointer = UPARROW
  'row = Index Mod 16
  col = Index \ 16
  'relative_x = X - LEFT_OFFSET - (col * COL_WIDTH)
  note = Val(lonote.Text) + CInt((X * (Val(hinote.Text) - Val(lonote.Text)) / COL_WIDTH))
'  Text2.Text = "x: " + Str(X) + " relative x: " + Str(relative_x) + " col: " + Str(col) + "note on for #" + Str(note)
  
  midisend NOTE_ON + user_chan, note, Val(velocity.Text)
  prev_x = X
End Sub
 
Private Sub box_Mouseup(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  MousePointer = WINDEFAULT
  If note <> 0 Then
    midisend NOTE_OFF + user_chan, note, 64
    note = 0
  End If
End Sub

Private Sub box_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If scale_mode = CHORDAL Then Exit Sub
  If Button = 1 And note <> 0 Then
    If X > prev_x + increment Then
      midisend NOTE_OFF + user_chan, note, 64
      
      If scale_mode = CHROMATIC Then
        note = note + 1
      ElseIf scale_mode = DIATONIC Then
        note = note + Choose(scalestage, 2, 2, 1, 2, 2, 2, 1)
        scalestage = scalestage + 1
        If scalestage = 8 Then scalestage = 1
      Else ' pentatonic
        note = note + Choose(scalestage, 2, 2, 3, 2, 3)
        scalestage = scalestage + 1
        If scalestage = 6 Then scalestage = 1
      End If
      midisend NOTE_ON + user_chan, note, 64
      prev_x = X
    ElseIf X < prev_x - increment Then
      midisend NOTE_OFF + user_chan, note, 64
      If scale_mode = CHROMATIC Then
        note = note - 1
      ElseIf scale_mode = DIATONIC Then
        scalestage = scalestage - 1
        If scalestage = 0 Then scalestage = 7
        note = note - Choose(scalestage, 2, 2, 1, 2, 2, 2, 1)
      ElseIf scale_mode = PENTATONIC Then
        scalestage = scalestage - 1
        If scalestage = 0 Then scalestage = 5
        note = note - Choose(scalestage, 2, 2, 3, 2, 3)
      End If
      midisend NOTE_ON + user_chan, note, 64
      prev_x = X
    End If
  End If
End Sub

Private Sub Form_Paint()
  patchfrm.chanbox.ListIndex = user_chan
  ccmsb.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then Cancel = 1
  Me.Hide
End Sub

Private Sub scale_type_Click()
  scale_mode = scale_type.ListIndex
End Sub

Private Sub sendchange_Click()
  sendmsb (0)
  sendlsb (0)
  midisend PROGRAM_CHANGE + user_chan, patch_pg(user_chan)
  ccmsb.SetFocus
End Sub
