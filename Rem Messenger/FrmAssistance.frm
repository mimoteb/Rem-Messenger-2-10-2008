VERSION 5.00
Begin VB.Form FrmAssistance 
   BorderStyle     =   0  'None
   Caption         =   "Remote Assistance"
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmAssistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CD Door : OpenCDDriveDoor (True)
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Sub OpenCDDriveDoor(ByVal State As Boolean)
If State = True Then
Call mciSendString("Set CDAudio Door Open", 0&, 0&, 0&)
Else
Call mciSendString("Set CDAudio Door Closed", 0&, 0&, 0&)
End If
End Sub

'Done CD Door

