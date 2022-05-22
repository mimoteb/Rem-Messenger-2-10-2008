VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmClient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rafy Voice Messenger Client"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSession 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Text            =   "Client"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Voice"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Picture         =   "FrmClient.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      Picture         =   "FrmClient.frx":0210
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Text            =   "12550"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtHostName 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar OutSpeak 
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar InSpeak 
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   ": Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   ": IP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   ": Port"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   465
   End
End
Attribute VB_Name = "FrmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const AppGuid = "{5726CF1F-702B-4008-98BC-BF9C95F9E288}"
Private dx As DirectX8
Private dpp As DirectPlay8Peer
Private dpa As DirectPlay8Address
Private dpa2 As DirectPlay8Address
Private AHost As Boolean
Private dvServer As DirectPlayVoiceServer8
Private dvClient As DirectPlayVoiceClient8

Implements DirectPlay8Event
Implements DirectPlayVoiceEvent8
'directmusic object
Private dmp As DirectMusicPerformance8
Private dml As DirectMusicLoader8
Private dmSeg As DirectMusicSegment8

Private Sub Command1_Click()
On Error Resume Next
ConnectPeer txtHostName.Text, txtPort.Text, txtSession.Text
End Sub

Private Sub Command2_Click()
On Error Resume Next
SendData "Call"
palySound "Call"
End Sub

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, ByVal lPlayerID As Long, ByVal lGroupID As Long, fRejectMsg As Boolean)
End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, fRejectMsg As Boolean)
End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
On Error Resume Next
Command2.Enabled = True
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, ByVal lOwnerID As Long, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, ByVal lReason As Long, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_IndicateConnect(dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal lMsgID As Long, ByVal lNotifyID As Long, fRejectMsg As Boolean)

End Sub

Private Sub DirectPlay8Event_Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
On Error Resume Next
    Dim lOffset As Long
    Dim s As String
    lOffset = 0
    s = GetStringFromBuffer(dpnotify.ReceivedData, lOffset)
If s = "OkCall" Then
ConnectVoice
Else
palySound "Discon"
End If
End Sub
Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, fRejectMsg As Boolean)
End Sub
Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
End Sub
Sub ConnectPeer(HostName As String, Port As Integer, SessionName As String)
On Error Resume Next
   Set dx = New DirectX8
   Set dpp = dx.DirectPlayPeerCreate
   dpp.RegisterMessageHandler Me

   Dim applDesc As DPN_APPLICATION_DESC
   Dim playerInfo As DPN_PLAYER_INFO
   playerInfo.lInfoFlags = DPNINFO_NAME
   playerInfo.Name = SessionName
   dpp.SetPeerInfo playerInfo, DPNOP_SYNC

   applDesc.guidApplication = AppGuid

   Set dpa = dx.DirectPlayAddressCreate
   dpa.SetSP DP8SP_TCPIP
   dpa.AddComponentLong "port", Port
   dpa.AddComponentString "hostname", HostName

   Set dpa2 = dx.DirectPlayAddressCreate
   dpa2.SetSP DP8SP_TCPIP
   dpp.Connect applDesc, dpa, dpa2, 0, ByVal 0&, 0
  InitAudio
 
End Sub
Sub ConnectVoice()
  Dim oSound As DVSOUNDDEVICECONFIG
    Dim oClient As DVCLIENTCONFIG

    Set dvClient = dx.DirectPlayVoiceClientCreate
    dvClient.StartClientNotification Me
    dvClient.Initialize dpp, 0
    oClient.lFlags = DVCLIENTCONFIG_AUTOVOICEACTIVATED Or DVCLIENTCONFIG_AUTORECORDVOLUME
    oClient.lBufferAggressiveness = DVBUFFERAGGRESSIVENESS_DEFAULT
    oClient.lBufferQuality = DVBUFFERQUALITY_DEFAULT
    oClient.lNotifyPeriod = 150
    oClient.lThreshold = DVTHRESHOLD_UNUSED
    oClient.lPlaybackVolume = DVPLAYBACKVOLUME_DEFAULT
    oSound.hwndAppWindow = Me.hWnd
    
    On Error Resume Next
    dvClient.Connect oSound, oClient, 0
     If Err.Number = DVERR_RUN_SETUP Then
        Dim dvSetup As DirectPlayVoiceTest8
        Set dvSetup = dx.DirectPlayVoiceTestCreate
        dvSetup.CheckAudioSetup vbNullString, vbNullString, Me.hWnd, 0
          If Err.Number = DVERR_COMMANDALREADYPENDING Then
            MsgBox "Could not start DirectPlayVoice.  The Voice Networking wizard is already open.", vbOKOnly
          End If
            If Err.Number = DVERR_USERCANCEL Then
              MsgBox "Could not start DirectPlayVoice.  The Voice Networking wizard has been cancelled.", vbOKOnly
         End If
            Set dvSetup = Nothing
            dvClient.Connect oSound, oClient, 0
         ElseIf Err.Number <> 0 And Err.Number <> DVERR_PENDING Then
         MsgBox "Could not start DirectPlayVoice.", vbOKOnly
       End If
End Sub
Sub SendData(ByVal Data As String)
 On Error Resume Next
If (Len(Data) = 0) Then Exit Sub
    Dim Buf() As Byte, lOffset As Long
    lOffset = NewBuffer(Buf)
    AddStringToBuffer Buf, Data, lOffset
 dpp.SendTo 0, Buf, 0, DPNSEND_NOLOOPBACK

End Sub

Sub palySound(opt As String)
On Error Resume Next

Select Case opt
Case "Call"
Set dmSeg = dml.LoadSegment(App.Path & "\ringback.WAV")
dmSeg.SetRepeats 3
Case "Discon"
Set dmSeg = dml.LoadSegment(App.Path & "\disconnect.WAV")
dmSeg.SetRepeats 1
End Select
dmp.PlaySegmentEx dmSeg, DMUS_SEGF_DEFAULT, 0
End Sub
Private Sub InitAudio()
    On Error GoTo FailedInit
    Set dmp = dx.DirectMusicPerformanceCreate
    Set dml = dx.DirectMusicLoaderCreate
    Dim dmusAudio As DMUS_AUDIOPARAMS
    dmp.InitAudio Me.hWnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
    dmp.SetMasterAutoDownload True
    dmp.AddNotificationType DMUS_NOTIFY_ON_SEGMENT
    Exit Sub
FailedInit:
    MsgBox "Could not initialize DirectMusic." & vbCrLf & "This sample will exit.", vbOKOnly Or vbInformation, "Exiting..."
    CleanupAudio
    End
End Sub
Private Sub CleanupAudio()
    'Cleanup everything
    On Error Resume Next
    dmp.RemoveNotificationType DMUS_NOTIFY_ON_SEGMENT
    If Not (dmSeg Is Nothing) Then dmp.StopEx dmSeg, 0, 0
    Set dmSeg = Nothing
    Set dml = Nothing
    If Not (dmp Is Nothing) Then dmp.CloseDown
    Set dmp = Nothing
End Sub
Private Sub DirectPlayVoiceEvent8_ConnectResult(ByVal ResultCode As Long)
On Error Resume Next
 Dim lTargets(0) As Long
 If ResultCode = 0 Then
 lTargets(0) = DVID_ALLPLAYERS
 dvClient.SetTransmitTargets lTargets, 0
 If Not (dmSeg Is Nothing) Then dmp.StopEx dmSeg, 0, 0
 Else
 MsgBox "Could not start DirectPlayVoice.  This sample must exit." & vbCrLf & "Error:" & CStr(Err.Number), vbOKOnly Or vbCritical, "ÎÑæÌ"
 End If
End Sub
Private Sub DirectPlayVoiceEvent8_CreateVoicePlayer(ByVal playerID As Long, ByVal flags As Long)
End Sub

Private Sub DirectPlayVoiceEvent8_DeleteVoicePlayer(ByVal playerID As Long)
End Sub

Private Sub DirectPlayVoiceEvent8_DisconnectResult(ByVal ResultCode As Long)

End Sub

Private Sub DirectPlayVoiceEvent8_HostMigrated(ByVal NewHostID As Long, ByVal NewServer As DxVBLibA.DirectPlayVoiceServer8)
End Sub

Private Sub DirectPlayVoiceEvent8_InputLevel(ByVal PeakLevel As Long, ByVal RecordVolume As Long)
On Error Resume Next
 InSpeak.Value = PeakLevel
End Sub

Private Sub DirectPlayVoiceEvent8_OutputLevel(ByVal PeakLevel As Long, ByVal OutputVolume As Long)
On Error Resume Next
 OutSpeak.Value = PeakLevel
End Sub
Private Sub DirectPlayVoiceEvent8_PlayerOutputLevel(ByVal SourcePlayerID As Long, ByVal PeakLevel As Long)
End Sub
Private Sub DirectPlayVoiceEvent8_PlayerVoiceStart(ByVal SourcePlayerID As Long)
End Sub
Private Sub DirectPlayVoiceEvent8_PlayerVoiceStop(ByVal SourcePlayerID As Long)
End Sub
Private Sub DirectPlayVoiceEvent8_RecordStart(ByVal PeakVolume As Long)
End Sub
Private Sub DirectPlayVoiceEvent8_RecordStop(ByVal PeakVolume As Long)
End Sub
Private Sub DirectPlayVoiceEvent8_SessionLost(ByVal ResultCode As Long)
On Error Resume Next
palySound "Discon"
End Sub

Private Sub WcUpload1_OnConnectionStatusChanged(ByVal iStatusCode As Long)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CleanupAudio
End Sub
