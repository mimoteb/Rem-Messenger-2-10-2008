Attribute VB_Name = "IconTray"
'Declare function
Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'
'Define icon structure
Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
'
'Used to pass message param to Shell_NotifyIcon
Global Const ADD_ICON = 0
Global Const MODIFY_ICON = 1
Global Const DELETE_ICON = 2
Global Const ICON_MESSAGE = 1
Global Const ICON_ICON = 2
Global Const ICON_TIP = 4
Public Sub InitIconStruct(hWnd As Long, TheIcon As Long, sTip As String, IconData As NOTIFYICONDATA)
    '
    IconData.cbSize = Len(IconData)
    IconData.hWnd = hWnd
    IconData.uID = vbNull
    IconData.uFlags = ICON_MESSAGE Or ICON_ICON Or ICON_TIP
    IconData.uCallbackMessage = vbNull
    IconData.hIcon = TheIcon
    IconData.szTip = sTip
    '
End Sub




