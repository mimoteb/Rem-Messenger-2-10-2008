Attribute VB_Name = "General"
'*******************************************************************************
' MODULE:       Secure File Transfer v0.1
' FILENAME:     General.bas
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

Public Function NrOfBlocks(ByVal FileLength As Long, ByVal BlockSize As Long) As Long
   Dim lngOutput As Long
   
   lngOutput = Int(FileLength / BlockSize)
   
   If FileLength Mod BlockSize <> 0 Then
      lngOutput = lngOutput + 1
   End If
   
   NrOfBlocks = lngOutput
End Function

Public Function NrOfPaddingBytes(ByVal FileLength As Long, ByVal BlockSize As Long) As Long
   Dim lngOutput As Long
         
   lngOutput = 16 - (FileLength Mod 16) 'Range 1..16
      
   If FileLength Mod BlockSize = 0 Then
      lngOutput = 0
   End If
      
   NrOfPaddingBytes = lngOutput
End Function



