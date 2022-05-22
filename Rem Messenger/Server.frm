VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "System32COMDLG32.OCX"
Begin VB.Form FrmServer 
   BackColor       =   &H00000000&
   Caption         =   "Server"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   14130
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
   Icon            =   "Server.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   942
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClearPic 
      Caption         =   "Cls"
      Height          =   315
      Left            =   9900
      TabIndex        =   28
      Top             =   4830
      Width           =   645
   End
   Begin VB.OptionButton optColor1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   13290
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   13710
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4470
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Height          =   315
      Index           =   7
      Left            =   12240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   12630
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Height          =   315
      Index           =   5
      Left            =   11460
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Height          =   315
      Index           =   4
      Left            =   11850
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Height          =   315
      Index           =   3
      Left            =   11070
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      Height          =   315
      Index           =   2
      Left            =   10680
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF00FF&
      Height          =   315
      Index           =   1
      Left            =   10290
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H006B7184&
      Height          =   315
      Index           =   0
      Left            =   9900
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4470
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   7860
      RightToLeft     =   -1  'True
      ScaleHeight     =   4335
      ScaleWidth      =   6195
      TabIndex        =   17
      Top             =   60
      Width           =   6225
   End
   Begin VB.PictureBox PicHook 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4920
      Picture         =   "Server.frx":0CCA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox TxtTrayIconToolTip 
      Height          =   285
      Left            =   2820
      TabIndex        =   15
      Top             =   1380
      Visible         =   0   'False
      Width           =   1890
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
      Left            =   30
      TabIndex        =   14
      Top             =   5580
      Width           =   390
   End
   Begin VB.CommandButton CmdUnderLined 
      Caption         =   "U"
      Height          =   330
      Left            =   870
      TabIndex        =   13
      Top             =   5580
      Width           =   390
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
      Left            =   450
      TabIndex        =   12
      Top             =   5580
      Width           =   390
   End
   Begin VB.CommandButton CmdNudge 
      Caption         =   "N"
      Height          =   330
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5580
      Width           =   390
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   390
      Left            =   6840
      TabIndex        =   7
      Top             =   6405
      Width           =   855
   End
   Begin VB.TextBox num 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   5670
      Width           =   435
   End
   Begin VB.ListBox lstusers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   7860
      TabIndex        =   5
      Top             =   4440
      Width           =   1980
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2655
      TabIndex        =   4
      Text            =   "Server"
      Top             =   5655
      Width           =   1140
   End
   Begin VB.TextBox TxtIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4485
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   5670
      Width           =   1170
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6840
      TabIndex        =   2
      Top             =   5940
      Width           =   855
   End
   Begin VB.Timer ReTime 
      Interval        =   1000
      Left            =   7080
      Top             =   5040
   End
   Begin VB.TextBox txtsend 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   5940
      Width           =   6795
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   0
      Left            =   7050
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   1
      Left            =   7050
      Top             =   570
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   2
      Left            =   7050
      Top             =   990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   3
      Left            =   7050
      Top             =   1395
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   4
      Left            =   7050
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   5
      Left            =   7050
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   6
      Left            =   7050
      Top             =   2610
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   7
      Left            =   7050
      Top             =   3015
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   8
      Left            =   7050
      Top             =   3405
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   9
      Left            =   7050
      Top             =   3795
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   10
      Left            =   7050
      Top             =   4185
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Wsk 
      Index           =   11
      Left            =   7050
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TxtList 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5520
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   7755
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   13080
      Top             =   4890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please choose your favorite color"
   End
   Begin VB.Image ImgIcon 
      Height          =   480
      Left            =   5550
      Picture         =   "Server.frx":1994
      Top             =   3720
      Width           =   480
   End
   Begin VB.Label LblUsers 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Users :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   5730
      TabIndex        =   10
      Top             =   5670
      Width           =   510
   End
   Begin VB.Label LblIp 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3825
      TabIndex        =   9
      Top             =   5670
      Width           =   630
   End
   Begin VB.Label LblServerName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name :"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1740
      TabIndex        =   8
      Top             =   5655
      Width           =   885
   End
   Begin VB.Menu mnuMessenger 
      Caption         =   "messenger"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuspt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMessengerQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuConversation 
      Caption         =   "Conversation"
      Begin VB.Menu mnuMessengerSave 
         Caption         =   "Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuMessengerSaveAs 
         Caption         =   "Save As"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPlusDraw 
         Caption         =   "Draw"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "ServerCommads"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Option Explicit
Dim Color As Long
Dim strX As String * 4
Dim strY As String * 4
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

Private Sub CmdClearPic_Click()
On Error Resume Next
   Dim I As Integer
    Pic.Cls
    For I = 0 To 11
     If FrmServer.Wsk(I).State = 7 Then
        FrmServer.Wsk(I).SendData "Undoxxx"
        DoEvents
     End If
    Next
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

Private Sub CmdClear_Click()
On Error Resume Next
txtsend.Text = ""
End Sub

Private Sub CmdNudge_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To 11
 If Wsk(I).State = 7 Then
  Wsk(I).SendData "nudgeme" & txtname.Text
  TxtList.Text = TxtList.Text + "You sent just a Nudge!" + vbNewLine + vbNewLine
  TxtList.SelStart = Len(TxtList)
 End If
Next
End Sub

Private Sub CmdSend_Click()
On Error Resume Next
Dim f As Integer
For f = 0 To 11
If Wsk(f).State = 7 Then
Wsk(f).SendData "newmssg" & txtname.Text + " : " & txtsend.Text
DoEvents
End If
Next
TxtList.Text = TxtList.Text + txtname.Text + " : " + txtsend.Text + vbNewLine + vbNewLine
txtsend.Text = ""
TxtList.SelStart = Len(TxtList)
End Sub


Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then
frmServerPrevInstance.Show 1, Me
Else
Call Loading
End If
End Sub


Private Sub lstusers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 2 Or lstusers.ListCount < 1 Then Exit Sub
mnuClose.Caption = "Close " & lstusers.List(lstusers.ListIndex)
PopupMenu mnuServer
End Sub

Private Sub mnuClose_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To 11
Wsk(I).SendData "clzclnt" & Mid(mnuClose.Caption, 7)
Next
End Sub

Private Sub mnuMessengerSaveAs_Click()
On Error Resume Next
CD.DialogTitle = "Save conversaton"
CD.FileName = Format(Now, Time)
CD.Filter = ".txt"
If TxtList.Text <> "" Then
CD.ShowSave
Open CD.FileName For Output As #1
Print , TxtList.Text
Close #1
End If
End Sub

Private Sub optColor_Click(Index As Integer)
    Color = optColor(Index).BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call QuitServer
End Sub

Private Sub mnuMessengerQuit_Click()
Me.Hide
Call QuitServer
End Sub

Private Sub optColor1_Click()
On Error Resume Next
CD.DialogTitle = "Choose a color"
CD.ShowColor
Color = CD.Color
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim I As Integer
If Button <> 1 Then Exit Sub
For I = 0 To 11
Pic.PSet (X, Y), Color
strX = X
strY = Y
If FrmServer.Wsk(I).State = 7 Then
Wsk(I).SendData "PSetxxx" + strX + strY + CStr(Color)
DoEvents
End If
Next
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim I As Integer
If Button <> 1 Then Exit Sub
For I = 0 To 11
  Pic.Line -(X, Y), Color
  strX = X
  strY = Y
 If FrmServer.Wsk(I).State = 7 Then
  Wsk(I).SendData "Linexxx" + strX + strY + CStr(Color)
  DoEvents
 End If
Next
End Sub

Private Sub ReTime_Timer()
On Error Resume Next
num.Text = lstusers.ListCount
End Sub

Private Sub txtname_Change()
On Error Resume Next
Me.Caption = txtname.Text
End Sub

Private Sub Wsk_Close(Index As Integer)
On Error Resume Next
Dim I As Integer
lstusers.Clear
For I = 0 To 11
Wsk(I).Close
Wsk(I).LocalPort = 9000
Wsk(I).Listen
txtname.Locked = False
Next
End Sub

Private Sub Wsk_Connect(Index As Integer)
txtname.Locked = True
End Sub

Private Sub Wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim I As Integer
For I = 0 To 11
If Wsk(I).State = 0 Then
Wsk(I).Accept requestID
Exit For
End If
Next
End Sub

Private Sub Wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'--------------------------------------------------
On Error Resume Next
Dim Cmd As String, TypeSound, I As Integer, Msg As String, S As String, NewUserSound, inX, inY, Clr

Wsk(Index).GetData S
Cmd = Left(S, 7)
Msg = Mid(S, 8)
    inX = Val(Mid(S, 8, 4))
    inY = Val(Mid(S, 12, 4))
    Clr = Val(Mid(S, 16))

'--------------------------------------------------
Select Case Cmd
 
 Case "newmssg"
  TxtList.Text = TxtList.Text + Msg + vbNewLine + vbNewLine
  TxtList.SelStart = Len(TxtList)
  TypeSound = sndPlaySound(App.Path & "\Core Files\Type.wav", 3)
  For I = 0 To 11
   If Wsk(I).State = 7 Then
    Wsk(I).SendData "message" + Msg
    DoEvents
   End If
  Next

 Case "newuser"
  NewUserSound = sndPlaySound(App.Path & "\Core Files\OnLine.wav", 3)
  TxtList.Text = TxtList.Text + Msg + " has just signed in." + vbNewLine + vbNewLine
  TxtList.SelStart = Len(TxtList)
  lstusers.AddItem Msg
  If Msg <> txtname.Text Then
  FrmPopup.Show
  FrmPopup.LblTitle.Caption = "Hey " & txtname.Text
  FrmPopup.Lbljoined.Caption = Msg & " has just signed in!"
  End If
  
  For I = 0 To 11
  If Wsk(I).State = 7 Then
  Wsk(I).SendData "newuser" & Msg
  DoEvents
  Wait1
  Wsk(I).SendData "mssgbox" + txtname.Text
  DoEvents
  Wait1
  Wsk(I).SendData "newlist" & Getlist
  DoEvents
  Wait1
  End If
  Next
 
 Case "deluser"
  lstusers.Clear
  Dim Y As Integer
  For Y = 0 To 11
  If Wsk(Y).State = 7 Then
  Wsk(Y).SendData "newlist" & Getlist
  DoEvents
  End If
  Next
     
 Case "nudgeme"
 For I = 0 To 11
  If Wsk(I).State = 7 Then
   TxtList.Text = TxtList.Text + Msg + " Sent you a Nudge! repeat it!" + vbNewLine + vbNewLine
   TxtList.SelStart = Len(TxtList)
   Wsk(I).SendData "nudgeme" & Msg
  End If
 Next
'----------------------------------------------------
' Draw Board
     Case "PSetxxx": Pic.PSet (inX, inY), Clr
     Case "Linexxx": Pic.Line -(inX, inY), Clr
     Case "Undoxxx": Pic.Cls

End Select
End Sub

Private Function Getlist() As String
Dim S As String, I As Integer
S = ""
For I = 0 To lstusers.ListCount - 1
S = S + "#" + lstusers.List(I)
Next
Getlist = Mid(S, 2)
End Function

Private Sub Wsk_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Dim I As Integer
lstusers.Clear
txtname.Locked = False
For I = 0 To 10
Wsk(I).Close
Wsk(I).LocalPort = 9000
Wsk(I).Listen
Next
End Sub

Private Sub Loading()
On Error Resume Next
Wsk(0).Close
Wsk(0).LocalPort = 9000
Wsk(0).Listen
Dim ax, I As Integer
    Pic.AutoRedraw = True
    Pic.ScaleMode = vbPixels
    ax = sndPlaySound(App.Path & "\Core Files\Start up.WAV", 3)
    TxtTrayIconToolTip.Text = "Hi " & txtname.Text & " , this messenger's verision is :" & " " & App.Major & "." & App.Minor & "." & App.Revision
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

Private Sub QuitServer()
On Error Resume Next
Dim I As Integer
TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = PicHook.hWnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI

For I = 0 To 11
 Wsk(I).SendData "kilsrvr"
 DoEvents
Next
End
End Sub

Private Sub Wait1()
On Error Resume Next
Dim T As Single
T = Timer
Do Until (Timer - T >= 1)
DoEvents
Loop
End Sub
Private Sub mnuPop_Click(Index As Integer)
    Select Case Index
        Case 0
            FrmServer.Show
    End Select
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MsGo
    MsGo = X / Screen.TwipsPerPixelX
    If MsGo = WM_LBUTTONDBLCLK Then
        'Left button double click
        mnuPop_Click 0
    ElseIf MsGo = WM_RBUTTONUP Then
        'Right button click
        Me.PopupMenu mnuPopup
    End If
End Sub

