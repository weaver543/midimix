VERSION 5.00
Begin VB.Form sectionform 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Sections"
   ClientHeight    =   7275
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   2805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "Seq"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Len"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Msr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label length_lbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.Label name_lbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Intro"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Label letter_lbl 
      BackColor       =   &H80000009&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Label measure_lbl 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   285
   End
   Begin VB.Menu add_menu 
      Caption         =   "add_menu"
      Begin VB.Menu append_menu 
         Caption         =   "Append"
      End
   End
   Begin VB.Menu section_menu 
      Caption         =   "menu"
      Begin VB.Menu goto_menu 
         Caption         =   "&Goto"
      End
      Begin VB.Menu delete_menu 
         Caption         =   "Delete"
      End
      Begin VB.Menu insert_menu 
         Caption         =   "Insert"
      End
   End
End
Attribute VB_Name = "sectionform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim measure, visible_sections, active_section, orig_caption, prev_y As Integer
Const LINE_HEIGHT = 240

Private Sub append_menu_Click()
  strr = InputBox("enter section in format: b,bridge,length[,repeats]")
  If strr = "" Then Exit Sub
  songsection(Inc(num_songsections)) = strr
  add_songsection strr
End Sub

Private Sub add_songsection(ByVal strr As String)
  Dim param
  param = Split5(strr, ",")
  If visible_sections > 0 Then
    Load measure_lbl(visible_sections)
    Load letter_lbl(visible_sections)
    Load name_lbl(visible_sections)
    Load length_lbl(visible_sections)
    measure = measure + CInt(length_lbl(visible_sections - 1).Caption)
    
    measure_lbl(visible_sections).Visible = True
    measure_lbl(visible_sections).Left = measure_lbl(0).Left
    measure_lbl(visible_sections).Top = 360 + (visible_sections * LINE_HEIGHT)
    letter_lbl(visible_sections).Visible = True
    letter_lbl(visible_sections).Left = letter_lbl(0).Left '480
    letter_lbl(visible_sections).Top = 360 + (visible_sections * LINE_HEIGHT)
    name_lbl(visible_sections).Visible = True
    name_lbl(visible_sections).Left = name_lbl(0).Left '720
    name_lbl(visible_sections).Top = 360 + (visible_sections * LINE_HEIGHT)
    length_lbl(visible_sections).Visible = True
    length_lbl(visible_sections).Left = length_lbl(0).Left '2400
    length_lbl(visible_sections).Top = 360 + (visible_sections * LINE_HEIGHT)
  End If

  measure_lbl(visible_sections).Caption = CStr(measure)
  measure_lbl(visible_sections).ToolTipText = CStr(measure)
  letter_lbl(visible_sections).Caption = param(0)
  name_lbl(visible_sections).Caption = param(1)
  length_lbl(visible_sections).Caption = param(2)
  
  visible_sections = visible_sections + 1
End Sub

Private Sub insert_menu_Click()
  strr = InputBox("enter section in format: b,bridge,length[,repeats]")
  If strr = "" Then Exit Sub

  For i = num_songsections - 1 To 1 Step -1
    Unload measure_lbl(i)
    Unload letter_lbl(i)
    Unload name_lbl(i)
    Unload length_lbl(i)
  Next

  For i = num_songsections To active_section + 1 Step -1
    songsection(i) = songsection(i - 1)
  Next
  
  songsection(active_section) = strr
  num_songsections = num_songsections + 1

  build_sectionlist
End Sub

Private Sub delete_menu_Click()
  For i = active_section To num_songsections - 1
    songsection(i) = songsection(i + 1)
  Next
    
  For i = num_songsections - 1 To 1 Step -1
    Unload measure_lbl(i)
    Unload letter_lbl(i)
    Unload name_lbl(i)
    Unload length_lbl(i)
  Next
  
  num_songsections = num_songsections - 1
  
  build_sectionlist
End Sub

Private Sub build_sectionlist()
  visible_sections = 0
  measure = 1
  For i = 0 To num_songsections - 1
    add_songsection songsection(i)
  Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 1 Then append_menu_Click
End Sub

Private Sub Form_Load()
  prev_y = 0
  measure = 1
  measure_lbl(0).Caption = ""
  letter_lbl(0).Caption = ""
  name_lbl(0).Caption = ""
  length_lbl(0).Caption = ""
  
  build_sectionlist
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu add_menu
End Sub

Private Sub goto_menu_Click()
  AppActivate appTitle
  SendKeys goto_key, True '"^g", True
  strr = Mid("000", 1, 3 - Len(measure_lbl(active_section).Caption)) & measure_lbl(active_section).Caption
  'SendKeys strr, True
  SendKeys Mid(strr, 1, 1), True
  SendKeys Mid(strr, 2, 1), True
  SendKeys Mid(strr, 3, 1), True
  SendKeys "~", True
  SendKeys followup_key, True ' spacebar to start playback
Debug.Print strr
End Sub

Private Sub measure_lbl_DblClick(Index As Integer)
Debug.Print "**** DOUBLECLICK"
  active_section = Index
  measure_lbl(Index).Caption = orig_caption ' cancel mousemove changes for dblClick
'MsgBox "reset to " '+ orig_caption
  goto_menu_Click
End Sub

Private Sub measure_lbl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  mousedown_time = GetTickCount
  orig_caption = measure_lbl(Index).Caption
End Sub

Private Sub measure_lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If Y > prev_y + 2 Then
      measure_lbl(Index).Caption = CStr(CInt(measure_lbl(Index).Caption) + 1)
      
    ElseIf Y < prev_y - 2 Then
      measure_lbl(Index).Caption = CStr(CInt(measure_lbl(Index).Caption) - 1)
    End If
    prev_y = Y 'CInt(measure_lbl(Index).Caption)
  End If
End Sub

Private Sub measure_lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    active_section = Index
    If Button = 2 Then
      measure_lbl(Index).Caption = orig_caption
      PopupMenu section_menu
    Else
      If GetTickCount - mousedown_time > doubleclick_time Then
Debug.Print "****  SINGLE CLICK; " + Str(GetTickCount - mousedown_time)
        goto_menu_Click
      End If
      
      measure_lbl(Index).Caption = orig_caption ' cancel mousemove changes
    End If
End Sub


