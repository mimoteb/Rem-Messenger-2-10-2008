VERSION 5.00
Begin VB.Form FrmPopup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Memo 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer TmrWait 
      Left            =   2280
      Top             =   720
   End
   Begin VB.TextBox TxtHeight 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer TmrPopup 
      Interval        =   10
      Left            =   720
      Top             =   840
   End
   Begin VB.Label Lbljoined 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "FrmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
TxtHeight.Text = Screen.Height
Me.Top = TxtHeight.Text
Memo.Text = TxtHeight.Text - 2200
End Sub

Private Sub Lbljoined_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub LblTitle_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub TmrPopup_Timer()
On Error Resume Next
 If Me.Top = Memo.Text Then
  TmrWait.Interval = 3000
  TmrPopup.Interval = 0
 Else
  Me.Top = TxtHeight.Text
  TxtHeight.Text = TxtHeight.Text - 20
 End If
End Sub

Private Sub TmrWait_Timer()
On Error Resume Next
Unload Me
TmrWait.Interval = 0
End Sub
