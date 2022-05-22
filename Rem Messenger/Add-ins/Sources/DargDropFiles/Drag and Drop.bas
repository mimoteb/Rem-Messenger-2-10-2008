Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, _
         ByVal hWnd As Long, _
         ByVal Msg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
        (ByVal hWnd As Long, _
         ByVal nIndex As Long, _
         ByVal dwNewLong As Long) As Long

Private lPrevWndProc    As Long
Private hHookWindow     As Long

Public Declare Sub DragAcceptFiles Lib "shell32.dll" _
        (ByVal hWnd As Long, _
         ByVal fAccept As Long)
         
Public Declare Sub DragFinish Lib "shell32.dll" _
        (ByVal hDrop As Long)
        
Public Declare Function DragQueryFile Lib "shell32.dll" _
    Alias "DragQueryFileA" _
        (ByVal hDrop As Long, _
         ByVal UINT As Long, _
         ByVal lpStr As String, _
         ByVal ch As Long) As Long

Public Const GWL_WNDPROC = -4
Public Const WM_DROPFILES = &H233

Public Sub SetHook(lHwnd As Long)

' Be sure the code is not already subclassing
If hHookWindow <> 0 Then Call Clearhook

' Save the handle of the calling window
hHookWindow = lHwnd

' Set the subclassing window message hook
lPrevWndProc = SetWindowLong(hHookWindow, _
                GWL_WNDPROC, _
                AddressOf HookCallback)

End Sub

Public Sub Clearhook()

Dim lReturn         As Long

' Check to be sure that there is a hook active
If hHookWindow = 0 Then Exit Sub
If IsEmpty(hHookWindow) = True Then Exit Sub
If IsNull(hHookWindow) = True Then Exit Sub

' Remove the hook from the system
lReturn = SetWindowLong(hHookWindow, _
                GWL_WNDPROC, _
                lPrevWndProc)

End Sub

Function HookCallback(ByVal hWnd As Long, _
            ByVal lMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
            
Select Case hWnd
    Case hHookWindow
        ' The message is for the window that is subclassed so ...
        ' Pass the message to the form for it to handle
        frmMain.MessageProc lMsg, wParam, lParam

    Case Else
        ' The message is for some other window so ...

End Select

' Pass the message through to the next message processor
HookCallback = CallWindowProc(lPrevWndProc, hWnd, lMsg, wParam, lParam)

End Function
