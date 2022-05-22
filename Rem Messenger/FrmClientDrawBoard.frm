VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "System32COMDLG32.OCX"
Begin VB.Form FrmClientDrawBoard 
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optColor1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   5490
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4470
      Width           =   375
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   330
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   4470
      Width           =   960
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   5910
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4470
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   2400
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Height          =   285
      Index           =   6
      Left            =   2790
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   1620
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Height          =   285
      Index           =   4
      Left            =   2010
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   3
      Left            =   1230
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      Height          =   285
      Index           =   2
      Left            =   840
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF00FF&
      Height          =   285
      Index           =   1
      Left            =   450
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4470
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H006B7184&
      Height          =   285
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4470
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   60
      RightToLeft     =   -1  'True
      ScaleHeight     =   4335
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   60
      Width           =   6225
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5340
      Top             =   4380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please choose your favorite color"
   End
End
Attribute VB_Name = "FrmClientDrawBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Color As Long
Dim strX As String * 4
Dim strY As String * 4

Private Sub CmdClear_Click()
On Error Resume Next
    Pic.Cls
    If FrmClient.wsk.State = 7 Then
       FrmClient.wsk.SendData "Undoxxx"
       DoEvents
    End If
End Sub

Private Sub Form_Load()
    Pic.AutoRedraw = True
    Pic.ScaleMode = vbPixels
End Sub

Private Sub optColor_Click(Index As Integer)
On Error Resume Next
    Color = optColor(Index).BackColor
End Sub

Private Sub optColor1_Click()
On Error Resume Next
CD.ShowColor
Color = CD.Color
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
    If Button <> 1 Or FrmClient.wsk.State <> 7 Then Exit Sub
    Pic.PSet (X, Y), Color
    strX = X
    strY = Y
    FrmClient.wsk.SendData "PSetxxx" + strX + strY + CStr(Color)
    DoEvents
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
     If Button <> 1 Or FrmClient.wsk.State <> 7 Then Exit Sub
      Pic.Line -(X, Y), Color
      strX = X
      strY = Y
      FrmClient.wsk.SendData "Linexxx" + strX + strY + CStr(Color)
      DoEvents
End Sub
