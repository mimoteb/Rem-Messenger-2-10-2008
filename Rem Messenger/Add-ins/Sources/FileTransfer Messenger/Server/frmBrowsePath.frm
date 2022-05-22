VERSION 5.00
Begin VB.Form frmBrowsePath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Path"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   3480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.DirListBox fileDirectory 
      Height          =   4140
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.DriveListBox fileDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmBrowsePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       Secure File Transfer v0.1
' FILENAME:     frmBrowsePath.frm
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

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOk_Click()
   Dim tmpDir As String
   'Process here!
   tmpDir = fileDirectory.List(fileDirectory.ListIndex)
   
   If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
   
   frmServer.EnterDestinationPath tmpDir
      
   Unload Me
End Sub

Private Sub fileDirectory_KeyPress(KeyAscii As Integer)
   Dim tmpDir As String
   If KeyAscii = vbKeyReturn Then
   
      KeyAscii = 0
      
      'Process here!
      tmpDir = fileDirectory.List(fileDirectory.ListIndex)
      
      If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
      
      frmServer.EnterDestinationPath tmpDir
      Unload Me
   
   ElseIf KeyAscii = vbKeyEscape Then
      Unload Me
   End If
   
   Unload Me
End Sub

Private Sub fileDrive_Change()
   On Error GoTo ErrorHandle
   
   fileDirectory.Path = fileDrive.Drive
   fileDirectory.SetFocus
   
   Exit Sub
ErrorHandle:
   MsgBox "Drive not available.", vbCritical, "Error"
   fileDrive.Drive = "c:\"
   fileDirectory.Path = "c:\"
End Sub

Private Sub Form_Load()
   Dim TempPath As String
      
   'To avoid problems with network shares
   TempPath = App.Path
   If Left$(TempPath, 1) = "\" Then TempPath = "C:\"
   
   DoEvents
   'Default directory
   fileDrive.Drive = Left$(TempPath, 3)
   DoEvents
   fileDirectory.Path = Left$(TempPath, 3)
   
   'Set focus to directory list
   frmBrowsePath.Visible = True
   DoEvents
   fileDirectory.SetFocus
End Sub
