VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capture Webcam Server By Alch3mist"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Sock 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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

Private Sub Form_Load()
    Sock.Close 'Close sockets first
    Sock.LocalPort = 5500 ' select local port
    Sock.Listen 'listen for incoming connections
    CamHwnd = capCreateCaptureWindow("Webcam", 0, 0, 0, 320, 240, Me.hWND, 0) '  create capture window
    SendMessage CamHwnd, 1034, 0, 0 'connect to capture source
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SendMessage CamHwnd, 1035, 0, 0 ' disconnect from capture source
End Sub

Private Sub Sock_Close()
    Sock.Close ' if socket closes re-listen
    Sock.Listen
End Sub

Private Sub Sock_ConnectionRequest(ByVal requestID As Long)
    Sock.Close 'close
    Sock.Accept requestID ' then accept the connection
    Me.Caption = "Connection Accepted From " & Sock.RemoteHostIP
End Sub


Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Rdata As String

    Sock.GetData Rdata
    Lstring = Mid(Rdata, 3, Len(Rdata))

Select Case Left(Rdata, 2)

    Case "1@"
    Call CapCam ' requested capture
    
    Case "2@"
    Sock.SendData "2@" & GetCamDrvs ' requested drivers
    
    Case "3@"
    Quality Lstring ' sets jpg quality
    Sock.SendData "3@" & Lstring
    
    Case "5@"

Open "C:\cam.jpg" For Binary As #1 ' open file to send

    TmpFsize = LOF(1) ' get files size
    
If TmpFsize = 0 Then
    Close #1
    Sock.SendData "4@" ' no file exists close send error
    Exit Sub
End If

ReDim TmpByte(TmpFsize - 1)
Get #1, 1, TmpByte ' get the entire byte array
    Sock.SendData TmpByte ' send the bytes
    TmpFsize = 0
Close #1 ' close file
    

End Select
End Sub

Function GetCamDrvs() As String ' holds the drivers string
On Error GoTo ErrDrvs
lResult = capGetDriverDescriptionA(0, DrvName, 100, DrvVer, 100)
If lResult Then
GetCamDrvs = DrvName & " Version " & DrvVer
Else
GetCamDrvs = "No Drivers" ' no drivers exist
End If
Exit Function
ErrDrvs:
GetCamDrvs = "No Drivers"
End Function

Sub CapCam()
On Error GoTo Errcam

    SendMessage CamHwnd, 1084, 0, 0 'Get Current Frame
    SendMessage CamHwnd, 1054, 0, 0 'Copy Current Frame to ClipBoard
    SavePicture Clipboard.GetData, "C:\cam.tmp" ' save picture
    DoEvents: DoEvents
    LoadBitmapFromFile "C:\cam.tmp" ' load bitmap
    DoEvents: DoEvents
    SaveJpegToFile "C:\cam.jpg" ' save the jpg
    TmpFsize = FileLen("C:\cam.jpg") ' get file size
    
    Sock.SendData "1@" & TmpFsize ' send size
    

    Exit Sub
Errcam:
   Sock.SendData "4@" ' there was an error
End Sub
