VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "File Drag and Drop"
   ClientHeight    =   3450
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7725
   Icon            =   "Drag and Drop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   7725
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   150
      TabIndex        =   2
      Top             =   480
      Width           =   7395
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "E&xit"
      Height          =   315
      Index           =   1
      Left            =   6480
      TabIndex        =   1
      Top             =   90
      Width           =   1065
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Clear List"
      Height          =   315
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   90
      Width           =   1065
   End
   Begin VB.Label lblCaption 
      Caption         =   "List of dropped file names:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   2565
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)

Select Case Index
    Case 0      ' Clear List
        lstFiles.Clear
            
    Case 1      ' Exit
        Unload frmMain
            
End Select

End Sub

Private Sub Form_Load()

' Establish the callback windows message hook
Call SetHook((Me.hWnd))

' Tell the OS this form accepts dropped files
Call DragAcceptFiles((Me.hWnd), True)
    
End Sub

Public Sub MessageProc(lMsg As Long, _
                wParam As Long, _
                lParam As Long)
                
Dim nDropCount          As Integer
Dim nLoopCtr            As Integer
Dim lReturn             As Long
Dim hDrop               As Long
Dim sFileName           As String

Select Case lMsg
    Case WM_DROPFILES
        ' Save the drop structure handle
        hDrop = wParam

        ' Allocate space for the return value
        sFileName = Space$(255)

        ' Get the number of file names dropped
        nDropCount = DragQueryFile(hDrop, -1, sFileName, 254)

        ' Loop to get each dropped file name and
        ' add it to the list box

        For nLoopCtr = 0 To nDropCount - 1
            ' Allocate space for the return value
            sFileName = Space$(255)

            ' Get a dropped file name
            lReturn = DragQueryFile(hDrop, nLoopCtr, sFileName, 254)
            lstFiles.AddItem Left$(sFileName, lReturn)

        Next nLoopCtr

        ' Release the drop structure from memory
        Call DragFinish(hDrop)
    
End Select

End Sub

Private Sub Form_Terminate()

Call Clearhook

End Sub
