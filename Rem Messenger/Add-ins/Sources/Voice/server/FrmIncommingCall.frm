VERSION 5.00
Begin VB.Form FrmIncommingCall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incomming Call"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Refuse Call"
      Height          =   555
      Left            =   2400
      Picture         =   "FrmIncommingCall.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Accept Call"
      Default         =   -1  'True
      Height          =   555
      Left            =   240
      Picture         =   "FrmIncommingCall.frx":022E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "FrmIncommingCall.frx":0488
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "FrmIncommingCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
FrmMainServer.StartVoiceServer
FrmMainServer.SendData "OkCall"
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
FrmMainServer.SendData "NoCall"
FrmMainServer.SendData "Discon"
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
FrmMainServer.palySound "Call"
Me.Left = Screen.Width - Me.Width
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Static i
i = i + 30
Me.Top = Screen.Height - i
If Me.Top <= Screen.Height - Me.Height - 360 Then
Timer1.Enabled = False
i = 0
End If
End Sub
