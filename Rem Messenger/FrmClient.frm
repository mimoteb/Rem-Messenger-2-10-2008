VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "System32COMDLG32.OCX"
Begin VB.Form FrmClient 
   BackColor       =   &H00000000&
   Caption         =   "Client"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   9885
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
   Icon            =   "FrmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   8790
      Top             =   5820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save conversation"
      Filter          =   ".txt"
   End
   Begin VB.TextBox TxtConnectedS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9270
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "00"
      Top             =   6735
      Width           =   285
   End
   Begin VB.TextBox TxtConnectedM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9015
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "00"
      Top             =   6735
      Width           =   285
   End
   Begin VB.TextBox TxtConnectedH 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "00"
      Top             =   6735
      Width           =   285
   End
   Begin VB.Timer TmrConnectedTime 
      Left            =   8205
      Top             =   6165
   End
   Begin VB.TextBox TxtTrayIconToolTip 
      Height          =   285
      Left            =   1620
      TabIndex        =   23
      Top             =   1410
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.PictureBox PicHook 
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   3720
      Picture         =   "FrmClient.frx":08CA
      ScaleHeight     =   960
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   3750
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CheckBox ChkConnect 
      BackColor       =   &H00000000&
      Caption         =   "Auto Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4260
      TabIndex        =   20
      Top             =   5835
      Width           =   1365
   End
   Begin VB.CommandButton CmdNudge 
      Caption         =   "N"
      Height          =   330
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5760
      Width           =   390
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "127.0.0.1"
      Top             =   7005
      Width           =   1170
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   975
      TabIndex        =   13
      Text            =   "Client"
      Top             =   7020
      Width           =   1140
   End
   Begin VB.TextBox num 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   7005
      Width           =   435
   End
   Begin VB.ListBox lstusers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5520
      ItemData        =   "FrmClient.frx":1194
      Left            =   7860
      List            =   "FrmClient.frx":1196
      TabIndex        =   11
      Top             =   60
      Width           =   1980
   End
   Begin VB.CommandButton CmdItalic 
      Caption         =   "I "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   510
      TabIndex        =   10
      Top             =   5760
      Width           =   390
   End
   Begin VB.CommandButton CmdUnderLined 
      Caption         =   "U"
      Height          =   330
      Left            =   930
      TabIndex        =   9
      Top             =   5760
      Width           =   390
   End
   Begin VB.CommandButton CmdBold 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   5760
      Width           =   390
   End
   Begin VB.TextBox txts 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   840
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6900
      TabIndex        =   5
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox TxtServer 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   1920
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Connect"
      Height          =   330
      Left            =   1770
      TabIndex        =   3
      Top             =   5760
      Width           =   1200
   End
   Begin VB.TextBox txtsend 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   6120
      Width           =   6795
   End
   Begin VB.Timer ReTime 
      Interval        =   500
      Left            =   2400
      Top             =   540
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   6900
      TabIndex        =   1
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   3000
      TabIndex        =   21
      Top             =   5760
      Width           =   1200
   End
   Begin VB.TextBox TxtList 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5520
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   60
      Width           =   7755
   End
   Begin VB.Label LblConnectedTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Connected :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7815
      TabIndex        =   24
      Top             =   6735
      Width           =   885
   End
   Begin VB.Image ImgIcon 
      Height          =   960
      Left            =   5325
      Picture         =   "FrmClient.frx":1198
      Top             =   3015
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LblYourIP 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   5130
      TabIndex        =   18
      Top             =   7020
      Width           =   630
   End
   Begin VB.Label LblServerName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   75
      TabIndex        =   17
      Top             =   6990
      Width           =   885
   End
   Begin VB.Label LblIp 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   7005
      Width           =   780
   End
   Begin VB.Label LblUsers 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Users :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4140
      TabIndex        =   15
      Top             =   7005
      Width           =   510
   End
   Begin VB.Label lblstat 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "lblstat"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6900
      TabIndex        =   7
      Top             =   7020
      Width           =   2955
   End
   Begin VB.Menu mnuMessenger 
      Caption         =   "Messenger"
      Begin VB.Menu mnuMessengerConnect 
         Caption         =   "Connection"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuspt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMessengerQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuConversation 
      Caption         =   "Conversation"
      Begin VB.Menu mnuMessengerSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuMessengerSaveAs 
         Caption         =   "Save As"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsGenerals 
         Caption         =   "Generals"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuPlus 
      Caption         =   "Plus!"
      Begin VB.Menu mnuPlusDraw 
         Caption         =   "Draw"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenMessenger 
         Caption         =   "Open Messenger"
      End
      Begin VB.Menu mnuQuitMessenger 
         Caption         =   "Quit Messenger"
      End
   End
End
Attribute VB_Name = "FrmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Tray Icon
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA

Private Sub Wait20()
On Error Resume Next
Dim T As Single
T = Timer
Do Until (Timer - T >= 4)
DoEvents
Loop
End Sub

Private Sub wait1()
On Error Resume Next
Dim T As Single
T = Timer
Do Until (Timer - T >= 2)
DoEvents
Loop
End Sub

Private Sub CmdBold_Click()
On Error Resume Next
If TxtList.FontBold = True Then
TxtList.FontBold = False
txtsend.FontBold = False
Else
TxtList.FontBold = True
txtsend.FontBold = True
End If
txtsend.SetFocus
End Sub

Private Sub CmdClear_Click()
On Error Resume Next
txtsend.Text = ""
txtsend.SetFocus
End Sub

Private Sub CmdItalic_Click()
On Error Resume Next
If TxtList.FontItalic = True Then
TxtList.FontItalic = False
txtsend.FontItalic = False
Else
TxtList.FontItalic = True
txtsend.FontItalic = True
End If
txtsend.SetFocus
End Sub

Private Sub CmdNudge_Click()
On Error Resume Next
If wsk.State = 7 Then
NudgeSound = sndPlaySound(App.Path & "\Core Files\Nudge.wav", 3)
wsk.SendData "nudgeme" & TxtName.Text
End If
End Sub

Private Sub CmdSend_Click()
On Error Resume Next
If wsk.State = 7 Then
wsk.SendData "newmssg" & TxtName.Text + " : " & txtsend.Text
txtsend.Text = ""
End If
txtsend.SetFocus
End Sub

Private Sub CmdUnderLined_Click()
On Error Resume Next
If txtsend.FontUnderline = True Then
txtsend.FontUnderline = False
TxtList.FontUnderline = False
Else
txtsend.FontUnderline = True
TxtList.FontUnderline = True
End If
txtsend.SetFocus
End Sub

Private Sub CmdClose_Click()
On Error Resume Next
ChkConnect.Value = 0
wsk.SendData "deluser" & TxtName.Text
wsk.Close
End Sub

Private Sub Form_Load()
Call Loading
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ChkConnect.Value = 0
Call QuitServr
End Sub

Private Sub mnuMessengerConnect_Click()
On Error Resume Next
FrmClientConnect.Show
End Sub

Private Sub mnuMessengerQuit_Click()
On Error Resume Next
ChkConnect.Value = 0
Call QuitServr
End Sub

Private Sub mnuMessengerSaveAs_Click()
On Error Resume Next
If TxtList.Text <> "" Then
CD.FileName = Format(Now, Time)
CD.Filter = ".txt"
CD.ShowSave
Open CD.FileName For Output As #1
Print , TxtList.Text
Close #1
End If
End Sub

Private Sub mnuOpenMessenger_Click()
On Error Resume Next
FrmClient.Show
End Sub

Private Sub mnuPlusDraw_Click()
On Error Resume Next
FrmClientDrawBoard.Show
End Sub

Private Sub mnuQuitMessenger_Click()
On Error Resume Next
Call QuitServr
End Sub

Private Sub ReTime_Timer()
On Error Resume Next
If ChkConnect.Value = 1 Then
 If wsk.State = 0 Or wsk.State = 8 Then
  wsk.Close
  wsk.Connect TxtIP.Text, 9000
 End If
End If

Select Case wsk.State
 Case 0
  FrmClientConnect.Show 1
 Case 7
  lblstat.Caption = "Connected" & " ; Server is : " & TxtServer.Text
End Select
num.Text = lstusers.ListCount
End Sub

Private Sub TmrConnectedTime_Timer()
On Error Resume Next
If wsk.State = 7 Then
If TxtConnectedS.Text > 59 Then
 TxtConnectedM.Text = Val(TxtConnectedM.Text) + 1
 TxtConnectedS.Text = "00"
End If

If TxtConnectedM.Text > 59 Then
   TxtConnectedH.Text = Val(TxtConnectedH.Text) + 1
   TxtConnectedM.Text = "00"
End If

TxtConnectedS.Text = Val(TxtConnectedS.Text) + 1
End If
End Sub

Private Sub TxtName_Change()
On Error Resume Next
Me.Caption = TxtName.Text
End Sub

Private Sub Wsk_Close()
On Error Resume Next
lstusers.Clear
TxtName.Locked = False
CmdConnect.Default = True
TmrConnectedTime.Interval = 0
lblstat.Caption = ""
FrmClient.Show 1
End Sub

Private Sub Wsk_Connect()
On Error Resume Next
TxtName.Locked = True
wsk.SendData "newuser" & TxtName.Text
CmdSend.Default = True
ChkConnect.Value = 1
WaitConnect
TmrConnectedTime.Interval = 1000
End Sub

Private Sub Wsk_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim S As String
Dim Cmd As String
Dim Msg As String

wsk.GetData S

Cmd = Left(S, 7)
Msg = Mid(S, 8)
txts.Text = S
    inX = Val(Mid(S, 8, 4))
    inY = Val(Mid(S, 12, 4))
    Clr = Val(Mid(S, 16))

'--------------------------------------------
Select Case Cmd

Case "newmssg"
 TxtList.Text = TxtList.Text + Msg + vbNewLine + vbNewLine
 TxtList.SelStart = Len(TxtList)
 TypeSound = sndPlaySound(App.Path & "\Core Files\Type.wav", 3)

Case "kilsrvr"
 ' The server logged out
 TxtList.Text = TxtList.Text + "The Server logged out!" + vbNewLine + vbNewLine
 
 Call Wsk_Close

Case "message"
 TxtList.Text = TxtList.Text + Msg + vbNewLine + vbNewLine
 TxtList.SelStart = Len(TxtList)
 TypeSound = sndPlaySound(App.Path & "\Core Files\Type.wav", 3)

Case "mssgbox"
' Server name
 TxtServer.Text = Msg
 
'Pop up
Case "newuser"
  If Msg = TxtName.Text Then
  Else
  TxtList.Text = TxtList.Text + Msg + " has just signed in." + vbNewLine + vbNewLine
  TxtList.SelStart = Len(TxtList)
  FrmPopup.Show
  FrmPopup.LblTitle.Caption = "Hey " & TxtName.Text
  FrmPopup.Lbljoined.Caption = Msg & " has just signed in!"
  NewUserSound = sndPlaySound(App.Path & "\Core Files\OnLine.wav", 3)
  End If
 
' Get a new list
Case "newlist"
 Dim a() As String
 Dim I As Integer
 a = Split(Msg, "#")
 lstusers.Clear
 For I = LBound(a) To UBound(a)
 lstusers.AddItem a(I)
 Next

Case "nudgeme"
 If Msg = TxtName.Text Then
   TxtList.Text = TxtList.Text + "You sent just a Nudge!" + vbNewLine + vbNewLine
   TxtList.SelStart = Len(TxtList)
   Exit Sub
  Else
   NudgeSound = sndPlaySound(App.Path & "\Core Files\Nudge.wav", 3)
   TxtList.Text = TxtList.Text + vbNewLine + vbNewLine + Msg + " Sent you a Nudge! repeat it!"
   TxtList.SelStart = Len(TxtList)
 End If
' Draw Borad
     Case "PSetxxx": FrmClientDrawBoard.Pic.PSet (inX, inY), Clr
     Case "Linexxx": FrmClientDrawBoard.Pic.Line -(inX, inY), Clr
     Case "Undoxxx": FrmClientDrawBoard.Pic.Cls
' Close
Case "clzclnt"
 If Msg = TxtName.Text Then
 MsgBox "Hey " & TxtName.Text & " The Server closed you", vbCritical, "Sorry! we must disconnect you!"
 wsk.SendData "deluser" & TxtName.Text
 ChkConnect.Value = 0
 Call Wsk_Close
 End If

End Select
End Sub

Private Sub wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
TmrConnectedTime.Interval = 0
lstusers.Clear
TxtName.Locked = False
CmdConnect.Caption = "Connect"
End Sub
' Here is Some Properties
Private Sub Loading()
On Error Resume Next
Dim ConnectedH As Integer, ConnectedM As Integer, ConnectedS As Integer
TxtList.Top = 4
TxtList.Left = 4
b$ = App.Path
StartUpSound = sndPlaySound(b$ + "\" + "Core Files\Start up.WAV", 3)
LblYourIP.Caption = "Your IP :" & wsk.LocalIP
TxtTrayIconToolTip.Text = "Hi " & TxtName.Text & " , this messenger's verision is :" & " " & App.Major & "." & App.Minor & "." & App.Revision
'Tray Icon
    TrayI.cbSize = Len(TrayI)
    'Set the window's handle (this will be used to hook the specified window)
    TrayI.hWnd = PicHook.hWnd
    'Application-defined identifier of the taskbar icon
    TrayI.uId = 1&
    'Set the flags
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    TrayI.hIcon = ImgIcon.Picture
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    'Set the tooltiptext
    
    TrayI.szTip = TxtTrayIconToolTip.Text
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI

End Sub
'Quit the messenger
Private Sub QuitServr()
On Error Resume Next
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = PicHook.hWnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
wsk.SendData "deluser" & TxtName.Text
Me.Hide
DoEvents
End
End Sub

Private Sub mnuPop_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmClient.Show
    End Select
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Msgo = X / Screen.TwipsPerPixelX
    If Msgo = WM_LBUTTONDBLCLK Then
        'Left button double click
        mnuPop_Click 0
    ElseIf Msgo = WM_RBUTTONUP Then
        'Right button click
        Me.PopupMenu mnuPopup
    End If
End Sub

Private Sub WaitConnect()
On Error Resume Next
Dim T As Single
T = Timer
Do Until (Timer - T >= 1)
DoEvents
Loop
End Sub

