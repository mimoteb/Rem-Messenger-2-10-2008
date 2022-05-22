VERSION 5.00
Begin VB.Form frmBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.FileListBox fileFile 
      Height          =   4380
      Left            =   2640
      System          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.DirListBox fileDirectory 
      Height          =   4365
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox fileDrive 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       Secure File Transfer v0.1
' FILENAME:     frmBrowse.frm
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
      
   If fileFile.FileName <> vbNullString Then
      'Process here!
      tmpDir = fileFile.Path
      If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
      frmClient.EnterFileData tmpDir, fileFile.FileName
   End If
   
   Unload Me
End Sub

Private Sub fileDirectory_Change()
   fileFile.Path = fileDirectory.Path
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

Private Sub fileFile_DblClick()
   Dim tmpDir As String
   
   tmpDir = fileFile.Path
   If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
         
   frmClient.EnterFileData tmpDir, fileFile.FileName
   
   Unload Me
End Sub

Private Sub fileFile_KeyPress(KeyAscii As Integer)
   Dim tmpDir As String
   
   If KeyAscii = vbKeyReturn Then
      
      KeyAscii = 0
      
      If fileFile.FileName <> vbNullString Then
         tmpDir = fileFile.Path
         If Right$(tmpDir, 1) = "\" Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
         frmClient.EnterFileData tmpDir, fileFile.FileName
         Unload Me
         
      End If
   
   
   End If
   
End Sub

Private Sub Form_Load()
   Dim TempPath As String
   TempPath = App.Path
   
   If Left$(TempPath, 1) = "\" Then TempPath = "C:\"
      
   DoEvents
      
   'Default directory
   fileDrive.Drive = Left$(TempPath, 3)
   fileDirectory.Path = Left$(TempPath, 3)
   
   'Set focus to directory list
   frmBrowse.Visible = True
   DoEvents
   fileDirectory.SetFocus
   
End Sub
