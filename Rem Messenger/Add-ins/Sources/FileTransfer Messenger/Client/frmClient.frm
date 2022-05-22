VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client - Secure File Transfer"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1080
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar ProgressBarClnt 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBarClnt 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1860
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   450
      SimpleText      =   "s"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Speed KB/s"
            TextSave        =   "Speed KB/s"
            Key             =   "Speed"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
            Text            =   "Info"
            TextSave        =   "Info"
            Key             =   "Info"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse ..."
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   4560
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblFileName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblKey 
      Caption         =   "Please enter a passphrase:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblFileToSend 
      Caption         =   "File to send:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       Secure File Transfer v0.1
' FILENAME:     frmClient.frm
' AUTHOR:       Tom Adelaar
' CREATED:      12-Dec-2003
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' E-mail:    TomAdelaar@hotmail.com
'
' MODIFICATION HISTORY:
' 12-Dec-2003   Tom Adelaar     Initial Version
'******************************************************************

Option Explicit
Option Base 0

Private m_FileCompletePath As String
Private m_FileName As String
Private m_FileLength As Long
Private m_BlockSize As Long 'Blocksize to encrypt and to send TCP data, should be a multiple of 16 bytes (128-bits)
Private m_BytesSend As Boolean 'For TCP flow control (Very Important!)

Private m_hKey As String
Private m_hIV As String

Private m_IPAddress As String

Private m_ConnectionStatus As Byte

Private Const DELIMITER As String = "ÿ"

Private Const m_STATUS_INIT As Byte = 0
Private Const m_STATUS_TRANSFER As Byte = 10
Private Const m_STATUS_WAIT_ACK As Byte = 20

'Needed for Pause Sub
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Very fast function, when dealing with byte Arrays and other variables (except Strings).
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' In milliseconds
Private Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

Private Sub cmdBrowse_Click()
   Load frmBrowse
   frmBrowse.Visible = True
End Sub

Public Sub EnterFileData(ByVal FilePath As String, ByVal FileName As String)
   'Called from frmBrowse
   Dim CompleteFilePath As String
      
   'For loading file
   CompleteFilePath = FilePath & "\" & FileName
   
   'Load file info to globlal variables
   m_FileCompletePath = CompleteFilePath
   m_FileName = FileName
      
   'Update screen
   lblFileName.Caption = FileName
   lblFileName.ToolTipText = CompleteFilePath
End Sub


Private Sub cmdSend_Click()
      
   'Check the button
   If cmdSend.Caption = "&Cancel" Then
      
      'Close and reset socket
      ResetSck "File transfer cancelled."
      'Change button command
      cmdSend.Caption = "&Send"
      'Exit the sub
      Exit Sub
      
   Else
      'Change button command
      cmdSend.Caption = "&Cancel"
   End If
   
   
   'Check if a file is selected
   If m_FileName = vbNullString Then
      
      'Change name button back to 'Send'
      cmdSend.Caption = "&Send"
      'Pop-up a messagebox
      MsgBox "Please select file.", vbInformation, "Info"
      Exit Sub
   End If
   
   'Get the key (or passphrase ...)
   If txtKey.Text = vbNullString Then
      m_hKey = hexSHA256(StrToHex("()^% 5t4nRd K3y ?!@*"))
      m_hIV = hexGetHASHIV(m_hKey)
   Else
      m_hKey = hexSHA256(StrToHex("()^% 5t4nRd K3y ?!@*" & txtKey.Text))
      m_hIV = hexGetHASHIV(m_hKey)
   End If
   
   'Get info for global variables
   'Blocksize:
   '8192 for LAN
   '4096 for ADSL / Internet
   '1024 for Modem
   m_BlockSize = 4096 '(Bytes)
   m_FileLength = FileLen(m_FileCompletePath)
      
   'Get IP Address
   m_IPAddress = InputBox("Please enter IP address: ", "Enter IP Address", m_IPAddress)
   
   If m_IPAddress <> vbNullString Then
      'Setup the TCP Link
      
      Status "Connecting..."
      
      With sckClient
      
         .Close 'Just to make sure ...
         DoEvents
         .Connect m_IPAddress, 10102
         
      End With
      
   Else
      cmdSend.Caption = "&Send"
   End If
   
End Sub

Private Sub Form_Load()
   
   'Get focus to browse button
   frmClient.Visible = True
   DoEvents
   cmdBrowse.SetFocus
   
   'Init padding mask (Not really necessary ... you may omit this)
   UpdatePaddingMask
   
   'Load default status
   Status "Client Idle."
   Speed 0
   Info "Welcome to Secure File Transfer"
   m_IPAddress = "Localhost"
   
End Sub


Private Sub sckClient_Close()
   ResetSck "Connection closed"
End Sub

Private Sub sckClient_Connect()
   Dim sSend As String
   
   'Report and init connection status
   Status "Connected"
   Info "New connection established"
   m_ConnectionStatus = m_STATUS_INIT
   
   'Do some action, send file info and blocksize
   sSend = m_FileName & DELIMITER & CStr(m_FileLength) & DELIMITER & _
           CStr(m_BlockSize) & DELIMITER & "-CRC-"
   
   'Encrypt the string
   sSend = EncryptStr(sSend, m_hKey, m_hIV, m_hKey, m_hIV)
   
   'Send File info
   SendStrData sSend
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
   'Makes use of connection status
   Dim sData As String
   
   sckClient.GetData sData, vbString
   
   'Connection status, when server is ready to receive
   If m_ConnectionStatus = m_STATUS_INIT Then
      
      If DecryptStr(sData, m_hKey, m_hIV, m_hKey, m_hIV) = "OK" Then
         
         m_ConnectionStatus = m_STATUS_WAIT_ACK
         
         'Send the File!
         Info "Sending file."
         SendFile
         
      Else
         
         ResetSck "Error at DataArrival (INIT)"
      
      End If
   'Waiting for acknownledgement from server, when complete file has transferred
   ElseIf m_ConnectionStatus = m_STATUS_WAIT_ACK Then
   
      If DecryptStr(sData, m_hKey, m_hIV, m_hKey, m_hIV) = "OK" Then
         
         m_ConnectionStatus = m_STATUS_INIT
         
         ResetSck "File transfer successful ..."
                  
         
      Else
         
        ResetSck "Error at DataArrival (WAIT_ACK) " & sData
         
      End If
   'Error report ... but not very likely to happen
   Else
      ResetSck "Error at DataArrival"
   End If
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   ResetSck Description
End Sub

Private Sub sckClient_SendComplete()
   'This is needed for TCP flow control, without this the Server
   '(or even the Client) can be flooded.
   'm_ByteSend is set to true when a TCP packet has been send.
   'Works in combination with SendStrData and SendByteData
   m_BytesSend = True
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Then
      
      KeyAscii = 0
      cmdSend.SetFocus
   
   End If
   
End Sub

Private Sub Status(ASCII As String)
   StatusBarClnt.Panels("Status").Text = ASCII
End Sub

Private Sub Speed(Value As Single)
   StatusBarClnt.Panels("Speed").Text = Format(Value, "#####0.00") & " KB/s"
End Sub

Private Sub Info(ASCII As String)
   StatusBarClnt.Panels("Info").Text = ASCII
End Sub

Public Sub SendStrData(ByVal ASCII As String)
   
   'This Sub includes TCP flow control
   m_BytesSend = False ' See SendComplete event
   
   With sckClient
      If .State = sckConnected Then
         .SendData ASCII
      End If
   End With
   
   'Needed for TCP flow control
   Do While m_BytesSend = False
      DoEvents
   Loop
   m_BytesSend = False
   
End Sub

Public Sub SendByteData(ByRef Data() As Byte)
   'This Sub includes TCP flow control
   m_BytesSend = False ' See SendComplete event
   
   With sckClient
      If .State = sckConnected Then
         .SendData Data
      End If
   End With
   
   'Needed for TCP flow control
   Do While m_BytesSend = False
      DoEvents
   Loop
   
   m_BytesSend = False
   
End Sub

Private Sub SendFile()
   'Main function to send files
   Dim Cipher As New eb_c_IncrementalCipher
   Dim Data() As Byte
   Dim DataEnc() As Byte
   Dim FileNumber As Integer
   Dim PaddingLength As Long
   Dim PaddingByte As Byte
   
   Dim TimeRef As Single
   Dim TotalBytes As Long
   Dim TotalFileLength As Long
   Dim NrOfBlocks As Long
      
   'Get filenumber
   FileNumber = FreeFile
   
   'Open the file
   Open m_FileCompletePath For Binary Access Read As #FileNumber
       
   'Load filelength for progressbar only
   TotalFileLength = LOF(FileNumber)
   
   'Init flowControl
   m_BytesSend = False
   
   ' Init data array
   ReDim Data(m_BlockSize - 1) As Byte
      
   'Init Cipher
   Cipher.StartEncryptRaw EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, m_hKey, m_hIV
      
   'Init timer
   TimeRef = Timer
   
   'Send the data
   'This also avoids the inclusion of the EOF byte?
   Do While (m_FileLength > 0)
                  
      'Check if it is the last block to send
      If m_FileLength < m_BlockSize And m_FileLength > 0 Then
         ReDim Data(m_FileLength - 1) As Byte
      End If
      
      'Read data from file
      Get #FileNumber, , Data
      
      'Check if extra padding bytes are needed, when sending the last packet
      If m_FileLength < m_BlockSize And m_FileLength > 0 Then
         
         'Make array suitable for encryption, multiple of 16 bytes(128-bit blocks)
         PaddingLength = 16 - (m_FileLength Mod 16)
         PaddingByte = CByte(PaddingLength)
         
         If PaddingLength = 16 Then PaddingByte = 0
                           
         'Redim to multiple of 16 bytes
         ReDim Preserve Data(m_FileLength + PaddingLength - 1) As Byte
                                                
         'Adding paddingbyte
         Data(m_FileLength + PaddingLength - 1) = PaddingByte
         
      End If
                 
      'Encrypt array
      DataEnc = Cipher.EncryptBLOB(Data)
            
      'Check for blocksize
      Info "Sending file, TCP Blocksize: " & (UBound(DataEnc) + 1)
      
      'SEND THE DATA!
      SendByteData DataEnc
      
      'For status reporting
      NrOfBlocks = NrOfBlocks + 1
                                   
      'Give the Client and other applications time to process other data
      Pause 1 'milliseconds
                  
      'Update status and progressbar
      TotalBytes = TotalBytes + CLng(UBound(Data) + 1)
      If NrOfBlocks Mod 10 = 0 And ((Timer - TimeRef) > 0.5) Then
         Speed (CSng(TotalBytes) / (CSng(Timer - TimeRef) * 1024!))
      End If
      
      If NrOfBlocks Mod 2 = 0 Then
         ProgressBarClnt.Value = (CInt(CSng(TotalBytes) * 100! / CSng(TotalFileLength)))
      End If
      
      'Update remaining filelength
      m_FileLength = m_FileLength - m_BlockSize
   
   Loop
      
   'Change connection status
   If sckClient.State = sckConnected Then
      m_ConnectionStatus = m_STATUS_WAIT_ACK
      Info "Waiting for acknowledgement."
   Else
      ResetSck "Error during file transfer."
   End If
         
   'Close the file
   Close #FileNumber
      
   'End cipher
   Cipher.FinishEncrypt
   Set Cipher = Nothing
      
End Sub

Private Sub ResetSck(ByVal txtInfo As String)
   'Very usefull Sub!!!
   
   'Reset all
   
   'General info
   sckClient.Close
   Status "Client Idle."
   
   'Report why the socket has closed and everything is reset
   Info txtInfo
   
   'Change 'Send'-button
   cmdSend.Caption = "&Send"
   
   'Re-init paddingmask
   UpdatePaddingMask
   
   'Reset global variables
   m_BlockSize = 0
   m_ConnectionStatus = m_STATUS_INIT
   'm_FileCompletePath = vbNullString
   m_FileLength = 0
   'm_FileName = vbNullString
   
   'Extra reset check for sub's SendStrData and SendByteData
   m_BytesSend = True
   DoEvents
   
End Sub

