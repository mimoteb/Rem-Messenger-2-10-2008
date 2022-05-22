VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Capture Webcam Client By Alch3mist"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Pgbar 
      Height          =   135
      Left            =   1560
      TabIndex        =   8
      Top             =   3360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   120
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start Capture"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Driver"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Quality"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtQuality 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Text            =   "50"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox TxtIp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Quality:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label LbStatus 
      Caption         =   "Status:- Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image ImCam 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Capture Webcam code By Alch3mist http://c-i-a.2ya.com
'forums http://darksideofkalez.com
'use this code as you wish just give me some credits in the about please

Private Sub Command1_Click()

Select Case Command1.Caption
    Case "Connect"
        Sock.Close 'close socket
        Sock.Connect TxtIp, 5500 ' make connection with server
        LbStatus = "Status: - Connecting to " & TxtIp
        Command1.Caption = "Disconnect"
    Case "Disconnect"
        Sock.Close ' close socket
        LbStatus = "Status: - Disconnected from " & TxtIp
        Command1.Caption = "Connect"
End Select

End Sub

Private Sub Command2_Click()

If Sock.State = 7 Then
    Sock.SendData "3@" & TxtQuality ' send the quality value
End If

End Sub

Private Sub Command3_Click()
If Sock.State = 7 Then
    Sock.SendData "2@" ' request driver infomation
End If
End Sub

Private Sub Command4_Click()

Select Case Command4.Caption

Case "Start Capture"
If Sock.State = 7 Then
    Sock.SendData "1@" ' request imagees
    LbStatus = "Status: - Requesting Capture"
    Command4.Caption = "Stop Capture"
End If
        
Case "Stop Capture"
If Sock.State = 7 Then
    LbStatus = "Status: - Requesting Capture"
    Sock.Close ' close connecton then reconnect
    DoEvents: DoEvents
    Sleep (200)
    Command4.Caption = "Start Capture"
    Close #2
    Sock.Connect
    
End If
    
End Select

End Sub

Private Sub Sock_Close()
    LbStatus = "Status: - Connection Closed"
    Command1.Caption = "Connect"
End Sub

Private Sub Sock_Connect()
    LbStatus = "Status: - Connected to " & TxtIp ' where connected
    Command1.Caption = "Disconnect"
End Sub

Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    LbStatus = "Status: - Connection Error" ' the was a socket error
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Rdata As String

    Rbytes = Rbytes + bytesTotal ' counts the bytes recived
    Sock.GetData Rdata ' get the data
    lstring = Mid(Rdata, 3, Len(Rdata)) ' strip the remaining string we use

Select Case Left(Rdata, 2)
    Case "1@"
    SfileLen = lstring ' get file length
    Sock.SendData "5@"
    Open "C:\camcap.jpg" For Binary Access Write As #2 ' open the file to write into
    Rbytes = 0
    Pgbar.Max = SfileLen ' set the progress bar max value
    LbStatus = "Status: - Recieving Images"
    
    Case "2@"
    LbStatus = "Status: - Drivers: " & lstring ' we have the drives
    
    Case "3@"
    LbStatus = "Status: - Quality Set to " & lstring ' quality was set
    
    Case "4@"
    LbStatus = "Status: - Webcam Error" ' there was an error
    Case Else
    
Put #2, , Rdata
Pgbar.Value = Rbytes
If Rbytes = SfileLen Then ' we recived the whole file
    Close #2 ' close it
    ImCam.Picture = LoadPicture("C:\camcap.jpg") ' display it
    Sock.SendData "1@" ' request the next image
    LbStatus = "Status: - Requesting Capture"
End If

End Select

End Sub
