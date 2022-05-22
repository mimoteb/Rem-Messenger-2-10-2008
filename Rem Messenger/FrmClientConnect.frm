VERSION 5.00
Begin VB.Form FrmClientConnect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrCo 
      Left            =   1230
      Top             =   2100
   End
   Begin VB.Timer TmrConnecting 
      Interval        =   1000
      Left            =   2790
      Top             =   3240
   End
   Begin VB.TextBox TxtMyIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   30
      TabIndex        =   5
      Text            =   "127.0.0.1"
      ToolTipText     =   "Type the Server's IP."
      Top             =   5250
      Width           =   4380
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   915
      TabIndex        =   3
      Text            =   "Client"
      ToolTipText     =   "Type the Server's IP."
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox TxtIP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   900
      TabIndex        =   1
      Text            =   "127.0.0.1"
      ToolTipText     =   "Type the Server's IP."
      Top             =   345
      Width           =   1200
   End
   Begin VB.Label LblConnect 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   705
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label CmdConnect 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   1230
      TabIndex        =   6
      Top             =   3900
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Nickname :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   780
   End
   Begin VB.Label LblIP 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Server IP :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "Connect to the Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "FrmClientConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If
End Sub

Private Sub CmdConnect_Click()
On Error Resume Next
If CmdConnect.Caption = "Connect" Then
  CmdConnect.Caption = "Cancel"
  If FrmClient.wsk.State <> 7 Then
  LblConnect.Visible = True
  FrmClient.wsk.Close
  FrmClient.wsk.Connect TxtIP.Text, 9000
  End If
  Else
  LblConnect.Visible = False
  CmdConnect.Caption = "Connect"
  FrmClient.wsk.Close
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Load FrmClient
 TxtMyIP.Text = "Your IP is : " & FrmClient.wsk.LocalIP
End Sub

Private Sub TmrCo_Timer()
On Error Resume Next
If LblConnect.Visible = True Then
LblConnect.Visible = False
Else
LblConnect.Visible = True
End If
End Sub

Private Sub TmrConnecting_Timer()
On Error Resume Next
If FrmClient.wsk.State = 7 Then
LblConnect.Visible = False
Unload Me
FrmClient.Show
Else
End If

If CmdConnect.Caption = "Connect" Then
  TmrCo.Interval = 0
  LblConnect.Visible = False
Else
  TmrCo.Interval = 400
End If
End Sub

Private Sub TxtIP_Change()
On Error Resume Next
FrmClient.TxtIP.Text = Me.TxtIP.Text
End Sub

Private Sub TxtName_Change()
On Error Resume Next
FrmClient.TxtName.Text = Me.TxtName.Text
End Sub
