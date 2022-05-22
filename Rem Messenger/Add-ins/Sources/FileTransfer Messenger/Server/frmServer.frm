VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server - Secure File Transfer"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBarArrival 
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer ServerTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4920
      Top             =   1080
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CommandButton cmdBrowsePath 
      Caption         =   "&Browse ..."
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "&Start Server"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBarSrv 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3015
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   450
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
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   5640
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBarSrv 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblReceivingFileName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label lblKey 
      Caption         =   "Please enter a passphrase:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblDestinationPathTxt 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblReceivingFile 
      Caption         =   "Receiving filename:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblDestinationPath 
      Caption         =   "Destination path:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       Secure File Transfer v0.1
' FILENAME:     frmServer.frm
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

Private m_CompleteFilePath As String
Private m_FileName As String
Private m_FileLength As Long
Private m_FileNrPadding As Long  'Number of needed padding bytes
Private m_BlockSize As Long      'Blocksize packets when encrypted and transmitted, should be a multiple of 16 bytes (128-bits)
Private m_FileNumber As Long     'Needed for reading and writing file
Private m_NrOfBlocks As Long     'Temp variable, mainly to keep track of progress

Private m_FileProcessLen As Long     'Amount of processed data
Private m_FileDecryptLen As Long
Private m_TotalDataArrival As Long  'Amount of arrived data and stored in Queue
Private m_DataQueue As New Collection 'The Queue to store all arriving TCP data in

Private m_DestinationPath As String  'Where to save the file

Private m_TimerRef As Single         'Time reference for download speed

Private m_Cipher As New eb_c_IncrementalCipher 'ebCrypt.dll, in this application: Rijndael, 128-bit blocks, 256-bit key

Private m_hKey As String   'hex key for rijndael CBC
Private m_hIV As String    'init vector for CBC mode

Private m_ConnectionStatus As Byte

Private Const DELIMITER As String = "ÿ"

Private Const m_STATUS_INIT As Byte = 0
Private Const m_STATUS_TRANSFER As Byte = 10
Private Const m_STATUS_WAIT_ACK As Byte = 20

'For sub Pause
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Very fast function, needed when dealing with byte-arrays and other variables (except Strings)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

Private Sub cmdBrowsePath_Click()

   Load frmBrowsePath
   frmBrowsePath.Visible = True
   
End Sub

Private Sub cmdStartServer_Click()
   Dim strTemp As String
   
   If cmdStartServer.Caption = "&Start Server" Then
      strTemp = "&Stop Server"
      
      'Reset some things
      ProgressBarSrv.Value = 0
      ProgressBarArrival.Value = 0
      Speed 0
      
      'Get the key (or passphrase ...)
      If txtKey.Text = vbNullString Then
         m_hKey = hexSHA256(StrToHex("()^% 5t4nRd K3y ?!@*"))
         m_hIV = hexGetHASHIV(m_hKey)
      Else
         m_hKey = hexSHA256(StrToHex("()^% 5t4nRd K3y ?!@*" & txtKey.Text))
         m_hIV = hexGetHASHIV(m_hKey)
      End If
         
      'Start Server Timer
      ServerTimer.Enabled = True
                           
      'Init connection status
      m_ConnectionStatus = m_STATUS_INIT
      
      'Start to listen for connections
      With sckServer
         .Close 'Just in case
         DoEvents
         .Bind 10102
         .Listen
      End With
      
      'Report
      Status "Listening"
      Info "Ready to receive file"
      
   Else
      
      strTemp = "&Start Server"
      
      'Close and reset socket
      ResetSck "Stopped listening."
      
   End If
   
   cmdStartServer.Caption = strTemp
End Sub

Private Sub Form_Load()
   Dim tmpPath As String
   
   'Get application path and analyse
   tmpPath = App.Path
   If Left$(tmpPath, 1) = "\" Then tmpPath = "C:" 'Network drives start with "\"
   If Right$(tmpPath, 1) = "\" Then tmpPath = Left$(tmpPath, Len(tmpPath) - 1)
   tmpPath = tmpPath & "\"
   
   'Load default destination directory
   lblDestinationPathTxt.Caption = tmpPath
   lblDestinationPathTxt.ToolTipText = lblDestinationPathTxt.Caption
   m_DestinationPath = lblDestinationPathTxt.Caption
    
   'Set focus to start server button
   frmServer.Visible = True
   DoEvents
   cmdStartServer.SetFocus
   
   'Init Padding Mask (Not really necessary ... you can omit this if you want)
   'It's just a trick to increase the security (not much) when adding padding bytes
   UpdatePaddingMask
   
   'Load default status
   Status "Server Idle"
   Speed 0
   Info "Welcome to Secure File Transfer"
   
End Sub

Public Sub EnterDestinationPath(DestinationPath As String)
   'This sub is called from frmBrowsePath
   
   lblDestinationPathTxt.Caption = DestinationPath & "\"
   lblDestinationPathTxt.ToolTipText = lblDestinationPathTxt.Caption
   m_DestinationPath = lblDestinationPathTxt.Caption
End Sub

Private Sub sckServer_Close()

   If m_ConnectionStatus = m_STATUS_WAIT_ACK Then
   
      ResetSck "File transfer successful ..."
   
   ElseIf m_ConnectionStatus = m_STATUS_TRANSFER Then
      
      ResetSck "Error During file transfer"
   
   Else
      
      ResetSck "Connection closed."
   
   End If
End Sub

Private Sub sckServer_ConnectionRequest(ByVal requestID As Long)
   'Close socket and accept request
   With sckServer
      .Close
      .Accept requestID
   End With
   
   Status "Connected"
   Info "Connection Request"
End Sub

Private Sub sckServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   
   ResetSck Description
   
End Sub

Private Sub sckServer_DataArrival(ByVal bytesTotal As Long)
   Dim sData As String
   Dim byteData() As Byte
   
   If m_ConnectionStatus = m_STATUS_TRANSFER Then
   
      sckServer.GetData byteData, vbByte + vbArray
                  
      'Add to buffer
      m_DataQueue.Add byteData
      
      'Prevents locking
       DoEvents
            
      'Update for progressbar
      m_TotalDataArrival = m_TotalDataArrival + bytesTotal
      
            
   ElseIf m_ConnectionStatus = m_STATUS_INIT Then
   
      sckServer.GetData sData, vbString
      
      'Process the data
      ProcessStrData sData
   Else
      'Error
      ResetSck "Error at Data Arrival"
   End If
   
   
End Sub

Private Sub ProcessStrData(ByVal ASCII As String)
   'Called during connection init, see Data_arrival, to get file info
   Dim sArray() As String
   Dim sTemp As String
      
   'Decrypt the string
   sTemp = DecryptStr(ASCII, m_hKey, m_hIV, m_hKey, m_hIV)
   
   If sTemp = vbNullString Then
      ResetSck "Error during Decryption (ProcessStrData)"
      Exit Sub
   End If
   
   Info sTemp
   
   'First a simple CRC check
   If Right$(sTemp, 5) = "-CRC-" Then
      
      'Load global variables
      sArray = Split(sTemp, DELIMITER)
      m_FileName = sArray(0)
      m_FileLength = CLng(sArray(1))   'Length receiving file, without padding bytes
      m_BlockSize = CLng(sArray(2))    'TCP Blocksize (8192 = LAN, 4096 = Internet, 1024 = Modem)
      
      'Calculate nr of padding bytes
      m_FileNrPadding = NrOfPaddingBytes(m_FileLength, m_BlockSize)
                              
      'Total filelength including padding
      m_FileLength = m_FileLength + m_FileNrPadding
      
      'Get complete file path
      m_CompleteFilePath = m_DestinationPath & m_FileName
            
      'Init the Incremental cipher
      m_Cipher.StartDecryptRaw EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, m_hKey, m_hIV
                        
      'Open file
      m_FileNumber = FreeFile
      Open m_CompleteFilePath For Binary Access Write As #m_FileNumber
      
      'Publish file name
      lblReceivingFileName.Caption = m_FileName
      
      'Change connection status
      m_ConnectionStatus = m_STATUS_TRANSFER
      
      'Send OK back for starting the file transfer
      SendStrData EncryptStr("OK", m_hKey, m_hIV, m_hKey, m_hIV)
      
      'Load timer reference
      m_TimerRef = Timer
      
      'Info
      Info "Receiving file."
      
   Else
   
      ResetSck "CRC check failed."
      
   End If
End Sub

Private Sub ServerTimer_Timer()
   'This Sub is needed to avoid flooding the server with
   'TCP data packets
   Dim byteData() As Byte
   Dim ShiftData() As Byte
   Dim byteLenData As Long
   Dim RemainLen As Long
   Dim PaddingLength As Long
            
   'Stop the timer
   ServerTimer.Interval = 0
            
   'Start the packet processing
   Do While m_DataQueue.Count > 0
      
      'Avoids locking the system
      DoEvents
      
      'Prevents an error when server is suddenly stopped
      If sckServer.State <> sckConnected Then Exit Do
      
      'Load data
      byteData = m_DataQueue.Item(1)
      
      'Remove directly!
      m_DataQueue.Remove 1
      
      'Calculate length packet from Queue
      byteLenData = UBound(byteData) + 1
      
      'Update total processed length
      m_FileProcessLen = m_FileProcessLen + byteLenData
                  
      'Start analysing PacketBuffer
      If m_FileProcessLen < m_FileLength Then
                                     
         'Decrypt data block
         byteData = m_Cipher.DecryptBLOB(byteData)
         
         'Update amount of decrypted material
         m_FileDecryptLen = m_FileDecryptLen + (UBound(byteData) + 1)
         
         'Write to disk
         Put #m_FileNumber, , byteData
                           
         'Status reporting
         m_NrOfBlocks = m_NrOfBlocks + 1
                               
      ElseIf m_FileProcessLen = m_FileLength Then 'm_Filelength includes padding bytes!
         
         'Check if padding removal is needed
         If m_FileNrPadding = 0 Then
            
            'Why not byteData(UBound(byteData)) for last packet?
            'Need to add one element extra, due to a strange 'padding bug' in Ebcrypt
            'Else the last 16 bytes are kept in memory by m_Cipher and are waiting
            'for the next decryption call.
            'Due to this extra element all the elements (last 16 bytes) still stored
            'in the cipher are released!
                        
            ReDim Preserve byteData(UBound(byteData) + 1) As Byte
                                                            
            'Decrypt
            byteData = m_Cipher.DecryptBLOB(byteData)
            
            'Update amount of decrypted material
            m_FileDecryptLen = m_FileDecryptLen + (UBound(byteData) + 1)
                                                            
            'Write Data
            Put #m_FileNumber, , byteData
                     
         Else
            
            'Get data from packetbuffer
            'Why not byteData(UBound(byteData))?
            'Need to add one element extra, due to a strange 'padding bug' in Ebcrypt
            'Else the last 16 bytes are kept in memory by m_Cipher and are waiting
            'for the next decryption call.
            'Due to this extra element all the elements (last 16 bytes) still stored
            'in the cipher are released!
                        
            ReDim Preserve byteData(UBound(byteData) + 1) As Byte
                        
            'Decrypt
            byteData = m_Cipher.DecryptBLOB(byteData)
                                    
            'Update amount of decrypted material
            m_FileDecryptLen = m_FileDecryptLen + (UBound(byteData) + 1)
         
            'Get Paddinglength and change the array
            PaddingLength = byteData(UBound(byteData))
            
            If PaddingLength = 0 Then PaddingLength = 16
            
            ReDim ShiftData(UBound(byteData) - PaddingLength) As Byte
            
            CopyMemory ShiftData(0), byteData(0), UBound(byteData) - PaddingLength + 1
                  
            'Write Data
            Put #m_FileNumber, , ShiftData
            
         
         End If
                           
         'Close file
         Close #m_FileNumber
         
         'Status reporting
         m_NrOfBlocks = m_NrOfBlocks + 1
                                    
         'Change connection status (only needed when disconnecting :)
          m_ConnectionStatus = m_STATUS_WAIT_ACK
          
         'Send ACK back, that everything went okay
         SendStrData EncryptStr("OK", m_hKey, m_hIV, m_hKey, m_hIV)
         
         'Quick check
         If m_FileLength <> m_FileDecryptLen Then
            MsgBox "Error ?: " & m_FileLength & vbCrLf & m_FileDecryptLen
         End If
                                    
         'Full bar looks nice ;)
         ProgressBarSrv.Value = 100
         ProgressBarArrival.Value = 100
      
      ElseIf m_FileProcessLen > m_FileLength Then 'Error!!!
      
         ResetSck "Error: Received too much data." ' Should not happen
         Exit Sub
      End If
      
      'Status reporting
      If m_NrOfBlocks Mod 11 = 0 Then
              
         If m_FileLength > 0 Then
            'Progressbar processed data
            ProgressBarSrv.Value = CInt(m_FileProcessLen * 100! / m_FileLength)
            
            'Progressbar received data
            ProgressBarArrival.Value = CInt(CSng(m_TotalDataArrival) * 100! / CSng(m_FileLength))
         End If
         
         'Processing speed in KByte/sec
         'Check if no division by zero is done
         If (Timer - m_TimerRef) > 0.2 Then
            Speed CSng(m_FileProcessLen) / ((Timer - m_TimerRef) * 1024!)
         End If
         
      End If
      
   Loop
   
   'Start the timer
   ServerTimer.Interval = 1
End Sub


Private Sub txtKey_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      cmdStartServer.SetFocus
   End If
End Sub

Public Sub SendStrData(ByVal ASCII As String)
   With sckServer
      If .State = sckConnected Then
         .SendData ASCII
      End If
   End With
End Sub

Private Sub Status(ASCII As String)
   StatusBarSrv.Panels("Status").Text = ASCII
End Sub

Private Sub Speed(Value As Single)
   StatusBarSrv.Panels("Speed").Text = Format(Value, "#####0.00") & " KB/s"
End Sub

Private Sub Info(ASCII As String)
   StatusBarSrv.Panels("Info").Text = ASCII
End Sub

Private Sub ResetSck(ByVal txtInfo As String)
   'Very usefull sub!
   
   'Reset all
   
   'General info
   sckServer.Close
   Status "Server Idle"
   Info txtInfo 'Report why the socket has been closed
         
   'Reset global variables
   m_BlockSize = 0
   m_ConnectionStatus = m_STATUS_INIT
   m_CompleteFilePath = vbNullString
   m_FileLength = 0
   m_FileName = vbNullString
   m_TimerRef = 0
   m_FileNrPadding = 0
   
   m_NrOfBlocks = 0
   
   m_TotalDataArrival = 0
   m_FileProcessLen = 0
   m_FileDecryptLen = 0
      
   'Close file ... just to make sure
   Close #m_FileNumber
   m_FileNumber = 0
      
   'Close Cipher and re-init
   Set m_Cipher = Nothing
   Set m_Cipher = New eb_c_IncrementalCipher
   
   'Remove displayed name
   lblReceivingFileName.Caption = vbNullString
   
   'Re-init Paddingmask
   UpdatePaddingMask
      
   'Disable Server Timer
   ServerTimer.Enabled = False
    
   'Reset progressbar -> Done when server starts to listen
      
   'Empty Queue
   Do While m_DataQueue.Count > 0
      m_DataQueue.Remove 1
   Loop
      
   'Re-Init the start server button
   cmdStartServer.Caption = "&Start Server"
End Sub
