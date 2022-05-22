VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "System32COMDLG32.OCX"
Begin VB.Form FrmServerDrawBoard 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClearPic 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4500
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4860
      Top             =   4530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please choose your favorite color"
   End
End
Attribute VB_Name = "FrmServerDrawBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClear_Click()
End Sub



Private Sub Sending()
 On Error Resume Next

End Sub

