Attribute VB_Name = "ModApi"
Public Declare Sub SaveJpegToFile Lib "cjpg.dll" (ByVal FileName As String)
Public Declare Sub Quality Lib "cjpg.dll" (ByVal Percent As Integer)
Public Declare Function LoadBitmapFromFile Lib "cjpg.dll" (ByVal FileName As String) As Boolean

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal nID As Long) As Long
Public Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Integer, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean

Public CamHwnd As Long
Public Lstring As String
Public DrvName As String * 100
Public DrvVer As String * 100
Public lResult As Long
Public TmpByte() As Byte ' Transfer strings
Public TmpFsize As Long
