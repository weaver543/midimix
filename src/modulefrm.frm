VERSION 5.00
Begin VB.Form modulefrm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Add/Delete Controls"
   ClientHeight    =   2745
   ClientLeft      =   3600
   ClientTop       =   4200
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton helpbtn 
      Caption         =   "help"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox presetbox 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cancelbutton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton okbutton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton add2avail 
      Caption         =   "<<"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton add2used 
      Caption         =   ">>"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.ListBox used 
      Height          =   2205
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox avail 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presets"
      Height          =   735
      Left            =   6000
      TabIndex        =   10
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Displayed"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Available"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "modulefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DragIndex As Long
Dim leaveOpen As Boolean

Private Sub add2used_Click()
  If avail.ListIndex = -1 Then Exit Sub
  If avail.ItemData(avail.ListIndex) = KEYBOARD Then MsgBox "Now move KEYBOARD to top of list"
  used.AddItem avail.Text
  used.ItemData(used.NewIndex) = avail.ItemData(avail.ListIndex)
  avail.RemoveItem avail.ListIndex
End Sub

Private Sub add2avail_Click()
  If used.ListIndex = -1 Then Exit Sub
  avail.AddItem used.Text
  avail.ItemData(avail.NewIndex) = used.ItemData(used.ListIndex)
  used.RemoveItem used.ListIndex
End Sub

Private Sub avail_DblClick()
  Call add2used_Click
End Sub

Private Sub cancelbutton_Click()
  Unload Me
End Sub

Private Sub helpbtn_Click()
MsgBox "You can change the order of the controls in the used box by drag and drop. " + vbCrLf _
 + "If adding the 'keyboard selection' and 'keyboard deselection' buttons, they must be added as a pair, with the selection button on top." + vbCrLf + _
 "If selecting KEYBOARD, it must be the top item in list"
End Sub

Private Sub okbutton_Click()
  Call zmain.hide_modules
  
  For i = 0 To used.ListCount - 1
    visible_modules(i) = used.ItemData(i)
  Next i
  num_modules = used.ListCount
  
  Call zmain.reveal_modules
  Me.Hide
End Sub

Private Sub presetbox_Click()
Dim par
  If presetbox.ListIndex = 0 Then Exit Sub
  
  used.Clear
  par = Split5(presets(presetbox.ListIndex - 1), ",")
  
  For i = 1 To UBound(par)
    used.AddItem descrip(CInt(par(i)))
    used.ItemData(used.NewIndex) = CInt(par(i))
    
  Next
  leaveOpen = True
  Me.Hide
  Call okbutton_Click

End Sub

Private Sub used_DblClick()
  Call add2avail_Click
End Sub

Private Sub Form_Load()

For i = 0 To num_modules - 1
  used.AddItem descrip(visible_modules(i))
  
  used.ItemData(used.NewIndex) = visible_modules(i)
Next i

For i = 0 To MAX_MODULES - 1
  found = False
  
  For vis = 0 To num_modules - 1
    If visible_modules(vis) = i Then
      found = True
      Exit For
    End If
  Next vis
    
  If Not found Then
    avail.AddItem descrip(i)
    avail.ItemData(avail.NewIndex) = i
  End If
Next i

presetbox.AddItem "Select Preset"
For i = 0 To num_presets - 1
  par = Split5(presets(i), ",")
  presetbox.AddItem par(0)
Next

presetbox.ListIndex = 0
  
End Sub

Private Sub used_MouseDown(Button As Integer, _
        shift As Integer, X As Single, Y As Single)
    DragIndex = used.ListIndex
End Sub

Private Sub used_MouseUp(Button As Integer, _
            shift As Integer, X As Single, Y As Single)
  If (DragIndex <> used.ListIndex) And used.ListIndex <> -1 Then
    tmptxt = used.List(DragIndex)
    tmpdata = used.ItemData(DragIndex)
    used.RemoveItem DragIndex
    used.AddItem tmptxt, used.ListIndex + Abs(shift = vbShiftMask)
    used.ItemData(used.NewIndex) = tmpdata
    used.ListIndex = used.NewIndex
  End If
End Sub

Private Sub Form_KeyDown(keycode As Integer, shift As Integer)
  If keycode = 27 Then Unload Me
End Sub

