VERSION 5.00
Begin VB.Form Devicefrm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Devices"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton connect_btn 
      Caption         =   "Connect"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton disconnect_btn 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.ListBox in2 
      Height          =   2400
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton closebox 
      Caption         =   "Close"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin VB.ListBox mci_out 
      Height          =   2400
      Left            =   2640
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ListBox mixer_out 
      Height          =   2400
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.ListBox mixer_in 
      Height          =   2400
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Processed Connection"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4935
      Begin VB.Label Label2 
         Caption         =   "Outputs"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Inputs"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Thru Connection"
      Height          =   3615
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   4935
      Begin VB.Label Label4 
         Caption         =   "Outputs"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Inputs"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label stat 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7080
      Width           =   3495
   End
End
Attribute VB_Name = "Devicefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub closebox_Click()
  Me.Hide
End Sub

Private Sub disconnect_btn_Click()
  Call zmain.Close_thruports
  stat.Caption = "Thruports disconnected"
End Sub

Private Sub connect_btn_Click()
  MousePointer = HOURGLASS
  If in2.ListIndex <> 0 Then
    stat.Caption = "connecting..."
    Call zmain.connect_thruports(in2.ListIndex, mci_out.ListIndex)
    stat.Caption = "Thruports connected"
  End If
  MousePointer = WINDEFAULT

End Sub

Private Sub mixer_in_Click()
  MousePointer = HOURGLASS
  Call zmain.minput_Click(mixer_in.ListIndex)
  MousePointer = WINDEFAULT
End Sub

Private Sub mixer_out_Click()
  MousePointer = HOURGLASS
  Call zmain.device_Click(mixer_out.ListIndex)
  MousePointer = WINDEFAULT
End Sub

Private Sub Form_KeyDown(keycode As Integer, shift As Integer)
  If keycode = 27 Then Unload Me
End Sub
