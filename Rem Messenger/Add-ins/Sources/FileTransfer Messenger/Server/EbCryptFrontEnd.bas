Attribute VB_Name = "EbCryptFrontEnd"
'*******************************************************************************
' MODULE:       Secure File Transfer v0.1
' FILENAME:     EbCryptFrontEnd.bas
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

' ------------------------------------------------------------------------------
'
' Crypto procedures - not all ebCrypt.dll functions are used
'
' ------------------------------------------------------------------------------
'
' This section is (for me) an "easier" front-end for EbCrypt.dll
'
' ------------------------------------------------------------------------------

Option Explicit

'Very fast function, needed when dealing with byte-arrays and other variables (except Strings)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Makes VB Random function (rnd) more secure, see AddPadding function,
'and also UpdatePaddingMask sub.
Public PaddingMask() As Byte

Public Type RSAKeys
   hPublicKey As String
   hPrivateKey As String
End Type

Public Function StrToHex(ByVal txtASCII As String) As String
  Dim Lib As New eb_c_Library
  Dim Data() As Byte
  Dim StrLength As Long
  Dim strTemp As String
  
  Data = StrConv(txtASCII, vbFromUnicode)
  strTemp = Lib.BLOBToHex(Data)
  
  'Remove from memory
  Set Lib = Nothing
  Erase Data
    
  'Publish the result
  StrToHex = strTemp
End Function

Public Function HexToStr(ByVal hexInput As String) As String
  Dim Lib As New eb_c_Library
  Dim strTemp As String
  Dim Data() As Byte
    
  Data = Lib.HexToBLOB(hexInput)
    
  strTemp = StrConv(Data, vbUnicode)
  
  'Remove from memory
  Set Lib = Nothing
  Erase Data
      
  'Publish the result
  HexToStr = strTemp

End Function

Public Function StrToByte(ByVal txtASCII As String) As Byte()
  Dim ByteArray() As Byte
  
  'Fast native function
  ByteArray = StrConv(txtASCII, vbFromUnicode)
  
  StrToByte = ByteArray
End Function

Public Function ByteToStr(ByRef ByteArray() As Byte) As String
  Dim strOutput As String
  
  On Error GoTo ErrorHandle
    
  'Fast native function
  strOutput = StrConv(ByteArray, vbUnicode)
  
  ByteToStr = strOutput
  
  Exit Function
  
ErrorHandle:

  ByteToStr = vbNullString
  
End Function

Public Function CopyByteArray(ByRef ByteArray() As Byte) As Byte()
  'Function to copy arrays
  'Better is to use CopyMemory Sub!
  Dim OutputArray() As Byte
  
  ReDim OutputArray(UBound(ByteArray)) As Byte
  
  OutputArray = ByteArray
  
  CopyByteArray = OutputArray
  
End Function

Public Function IntToStr(ByVal intValue As Integer) As String
  Dim ByteArray(1) As Byte
  Dim strOutput As String
  
  CopyMemory ByteArray(0), intValue, 2
    
  strOutput = StrConv(ByteArray, vbUnicode)
  
  IntToStr = strOutput
  
End Function

Public Function StrToInt(ByVal strValue As String) As Integer
  Dim ByteArray() As Byte
  Dim intOutput As Integer
  
  ByteArray = StrConv(strValue, vbFromUnicode)
  
  CopyMemory intOutput, ByteArray(0), 2
  
  StrToInt = intOutput
  
End Function

Public Function LongToStr(ByVal lngValue As Long) As String
  Dim ByteArray(3) As Byte
  Dim strOutput As String
  
  CopyMemory ByteArray(0), lngValue, 4
    
  strOutput = StrConv(ByteArray, vbUnicode)
  
  LongToStr = strOutput
  
End Function

Public Function StrToLong(ByVal strValue As String) As Long
  Dim ByteArray() As Byte
  Dim lngOutput As Long
  
  ByteArray = StrConv(strValue, vbFromUnicode)
  
  CopyMemory lngOutput, ByteArray(0), 4
  
  StrToLong = lngOutput
  
End Function


Public Function hexGetRandomData(ByVal numBytes As Long) As String
  'Relative slow function
  Dim Generator As New eb_c_PRNGenerator
  Dim Lib As New eb_c_Library
  Dim Data() As Byte
  Dim strOutput As String
  
  Generator.SeedUpdateFrequency = 100
  Generator.SeedWithSystemInfo
  
  Data = Generator.GetRandomBytes(numBytes)
  strOutput = Lib.BLOBToHex(Data)
  
  'Remove from memory
  Set Generator = Nothing
  Set Lib = Nothing
  Erase Data
  
  'Publish the result
  hexGetRandomData = strOutput
  
End Function

Public Function hexGenerateNewKey(ByVal hexPrefixStr As String) As String
  
  Dim Lib As New eb_c_Library
  Dim DataArray1() As Byte
  Dim DataArray2() As Byte
  Dim i As Long
  Dim strTemp As String
  Dim strOutput As String
    
  'Obtain first key based on SHA256, PRNG and your Hex Prefix String
  strTemp = hexPrefixStr
  strTemp = strTemp & hexGetRandomData(1024)
  strTemp = hexSHA256(strTemp)
  DataArray1 = Lib.HexToBLOB(strTemp)
  
  'Get new 256-bit randomdata
  strTemp = hexGetRandomData(32)
  DataArray2 = Lib.HexToBLOB(strTemp)
  
  'Get new key using XOR on both arrays
  For i = 0 To UBound(DataArray1)
    DataArray1(i) = DataArray1(i) Xor DataArray2(i)
  Next i
  
  'Transfer key to Hex
  strOutput = Lib.BLOBToHex(DataArray1)
  
  'Remove from memory
  Set Lib = Nothing
  Erase DataArray1
  Erase DataArray2
  
  'Publish the result
  hexGenerateNewKey = strOutput
  
End Function


Public Function hexSHA256(ByVal hexInput As String) As String
  Dim SHA256 As New eb_c_Hash
  Dim Lib As New eb_c_Library
  Dim ByteArray() As Byte
  Dim strOutput As String
    
  ByteArray = Lib.HexToBLOB(hexInput)
  strOutput = SHA256.HashBLOB(EB_CRYPT_HASH_ALGORITHM_SHA256, ByteArray)
  
  'Remove from memory
  Set SHA256 = Nothing
  Set Lib = Nothing
  Erase ByteArray
    
  'Publish result
  strOutput = Right$("00000000" & strOutput, 64) 'just in case
  hexSHA256 = strOutput
  
End Function

Public Function hexMD5(ByVal hexText As String) As String
  Dim MD5 As New eb_c_Hash
  Dim Lib As New eb_c_Library
  Dim ByteArray() As Byte
  Dim strOutput As String
    
  ByteArray = Lib.HexToBLOB(hexText)
  strOutput = MD5.HashBLOB(EB_CRYPT_HASH_ALGORITHM_MD5, ByteArray)
  
  'Remove from memory
  Set MD5 = Nothing
  Set Lib = Nothing
  Erase ByteArray
    
  'Publish result
  strOutput = Right$("00000000" & strOutput, 32) 'Just in case
  hexMD5 = strOutput
  
End Function

Public Function hexHMAC_SHA256(ByVal hex_Input As String, ByVal hex_Key As String) As String
  'For SHA256 only,
  '(This function is verified by replacing SHA256 with MD5)
  ' This function is too slow to handle very large bitrates
  
  Dim Lib As New eb_c_Library
  Dim SHA256 As New eb_c_Hash
  Dim IPad As Byte
  Dim OPad As Byte
  Dim B As Integer
  Dim ByteArray1() As Byte
  Dim ByteArray2() As Byte
  Dim hexInput As String
  Dim hexKey As String
  Dim strTemp As String
  Dim strOutput As String
  Dim i As Long
  Dim LengthKey As Long
  
  'Else problems with passing function values
  hexInput = hex_Input
  hexKey = hex_Key
  
  IPad = &H36
  OPad = &H5C
  B = 128 'hex block length is 2*Byte(=2*64)
  
  ' Append key with trailing zero's
  LengthKey = Len(hexKey)
  
  If LengthKey > B Then
    hexKey = SHA256.HashBLOB(EB_CRYPT_HASH_ALGORITHM_SHA256, Lib.HexToBLOB(hexKey))
    LengthKey = Len(hexKey)
  End If
  
  For i = LengthKey To (B - 2) Step 2
    hexKey = hexKey & "00"
  Next i
    
  'Xor new key with IPad
  ByteArray1 = Lib.HexToBLOB(hexKey)
  
  For i = 0 To UBound(ByteArray1)
    ByteArray1(i) = ByteArray1(i) Xor IPad
  Next i
  
  ' Append Xored key with text and hash it
  strTemp = Lib.BLOBToHex(ByteArray1)
  strTemp = strTemp & hexInput
  ByteArray1 = Lib.HexToBLOB(strTemp)
  strTemp = SHA256.HashBLOB(EB_CRYPT_HASH_ALGORITHM_SHA256, ByteArray1)
  
  'Xor new key with OPad
  ByteArray2 = Lib.HexToBLOB(hexKey)
  
  For i = 0 To UBound(ByteArray2)
    ByteArray2(i) = ByteArray2(i) Xor OPad
  Next i
  
  ' Append new Xored key with the hashed-text and hash it again
  strOutput = Lib.BLOBToHex(ByteArray2)
  strOutput = strOutput & strTemp
  ByteArray2 = Lib.HexToBLOB(strOutput)
  strOutput = SHA256.HashBLOB(EB_CRYPT_HASH_ALGORITHM_SHA256, ByteArray2)
  
  'Remove from memory
  Set Lib = Nothing
  Set SHA256 = Nothing
  Erase ByteArray1
  Erase ByteArray2
  
  'Publish the result
  hexHMAC_SHA256 = strOutput
End Function

Public Function hexQuickMAC(ByVal hexText As String, ByVal hexKey As String) As String
  'Proprietary MAC ... Two times faster than strHMAC_SHA256
  Dim hexOutput As String
  Dim tmpKey As String
       
  tmpKey = hexSHA256(tmpKey)
  hexOutput = hexSHA256(hexText)
  
  ' Append both results and hash it again!
  hexOutput = hexOutput & tmpKey
  hexOutput = hexSHA256(hexOutput)
  
  'Publish result
  hexQuickMAC = hexOutput
  
End Function

Public Function hexGetHASHIV(ByVal hexPassword As String) As String
  Dim strOutput As String
  
  strOutput = hexMD5(hexSHA256(hexPassword))
  
  'Just in case
  strOutput = Right$("0000000000" & strOutput, 32)
    
  'Publish the result
  hexGetHASHIV = strOutput
  
End Function

Public Sub UpdatePaddingMask()
   
  Dim Generator As New eb_c_PRNGenerator
  Dim Lib As New eb_c_Library
      
  Generator.SeedUpdateFrequency = 100
  Generator.SeedWithSystemInfo
  
  PaddingMask = Generator.GetRandomBytes(17) '16 is actually enough
    
  'Remove from memory
  Set Generator = Nothing
  Set Lib = Nothing
         
End Sub

Public Function AddPadding(ByRef DataArray() As Byte) As Byte()
   Dim ArrayLength As Long
   Dim i As Long, j As Long
   Dim PaddingByte As Byte
   Dim PaddingLength As Long
   Dim outArray() As Byte
   
   ArrayLength = UBound(DataArray) + 1
   
   'Calculate the paddingbytes
   If ArrayLength <> 0 Then
      
      PaddingLength = 16 - (ArrayLength Mod 16)
      PaddingByte = CByte(PaddingLength)
      If PaddingLength = 16 Then PaddingByte = 0
   
   Else
      
      ArrayLength = 0
      PaddingLength = 16
      PaddingByte = 0
   
   End If
   
   'Mask the paddingbyte
   Randomize
   PaddingByte = (Int(Rnd * 256) And &HF0) Xor PaddingByte
   
   'Redim the new array
   ReDim outArray(ArrayLength + PaddingLength - 1) As Byte
   CopyMemory outArray(PaddingLength), DataArray(0), ArrayLength
      
   'Fill padding with random data and paddingmask
   ' Add padding bytes in front of string
   For i = 0 To (PaddingLength - 1)
      outArray(i) = Int(Rnd * 256) Xor PaddingMask(j)
      j = j + 1
   Next i
   
   'Add the first byte with info about padding size
   outArray(0) = PaddingByte
   
   'Publish the result
   AddPadding = outArray
   
End Function

Public Function RemovePadding(ByRef DataArray() As Byte) As Byte()
   Dim ArrayLength As Long
   Dim PaddingLength As Byte
   Dim outArray() As Byte
   
   ArrayLength = UBound(DataArray) + 1
   
   PaddingLength = DataArray(0) And &HF
   
   If ArrayLength <> 0 Then
   
      If PaddingLength <> 0 Then
         
         ArrayLength = ArrayLength - PaddingLength
         ReDim outArray(ArrayLength - 1) As Byte
         CopyMemory outArray(0), DataArray(PaddingLength), ArrayLength
         
      Else
      
         If ArrayLength > 16 Then
         
            ArrayLength = ArrayLength - 16
            ReDim outArray(ArrayLength - 1) As Byte
            CopyMemory outArray(0), DataArray(16), ArrayLength
            
         End If
                  
      End If
   End If
      
   'Publish the result
   RemovePadding = outArray
End Function

Public Function EncryptByte( _
      ByRef Data() As Byte, _
      ByVal hexSKey As String, _
      ByVal hexSIV As String, _
      ByVal hexRKey As String, _
      ByVal hexRIV As String) As Byte()
      
  'Double encryption ... first Serpent and than Rijndael, both using 256-bit keys
  'Double encryption an overkill ? -> YES! (Who cares!)
        
  Dim Cipher1 As New eb_c_IncrementalCipher
  Dim Cipher2 As New eb_c_IncrementalCipher
  Dim InputData() As Byte
  Dim LengthArray As Long
  
  LengthArray = UBound(Data) + 1
  
  ReDim InputData(LengthArray - 1) As Byte
  CopyMemory InputData(0), Data(0), LengthArray
  
  'Add padding
  InputData = AddPadding(InputData)
      
  'Start the encryption
  Call Cipher1.StartEncryptRaw(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, hexSKey, hexSIV)
    
  InputData = Cipher1.EncryptBLOB(InputData)
      
  ' End the encryption
  Cipher1.FinishEncrypt
      
  Call Cipher2.StartEncryptRaw(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, hexRKey, hexRIV)
      
  InputData = Cipher2.EncryptBLOB(InputData)
        
  ' End the encryption
  Cipher2.FinishEncrypt
        
  'Remove from memory
  Set Cipher1 = Nothing
  Set Cipher2 = Nothing
    
  'Publish the result
  EncryptByte = InputData
  Exit Function
  
ErrorHandle:
  MsgBox "Error during byte array encryption"
  
End Function

Public Function DecryptByte( _
        ByRef Data() As Byte, _
        ByVal hexSKey As String, _
        ByVal hexSIV As String, _
        ByVal hexRKey As String, _
        ByVal hexRIV As String) As Byte()
    
  Dim Cipher2 As New eb_c_IncrementalCipher
  Dim Cipher1 As New eb_c_IncrementalCipher
  Dim InputData() As Byte
  Dim LengthArray As Long
  
  On Error GoTo ErrorHandle
  
  'For decryption padding bytes are needed, this is an error in the
  'ebCrypt.dll or maybe in the OpenSSL package, on which it is based.
  '16 bytes = 128 bit
      
  LengthArray = UBound(Data) + 1
  ReDim InputData(LengthArray - 1) As Byte
  CopyMemory InputData(0), Data(0), LengthArray
    
  'Add padding
  ReDim Preserve InputData(LengthArray + 15) As Byte '16 - 1 = 15
    
  'Decrypt with Rijndael
  Call Cipher2.StartDecryptRaw(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, hexRKey, hexRIV)
  InputData = Cipher2.DecryptBLOB(InputData)
      
  'Add padding again
  ReDim Preserve InputData(LengthArray + 15) As Byte '16 - 1 = 15
          
  'Remove from memory
  Set Cipher2 = Nothing 'This is one way to finish decryption :)
  
  'Decrypt with Serpent
  Call Cipher1.StartDecryptRaw(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, hexSKey, hexSIV)
  InputData = Cipher1.DecryptBLOB(InputData)
         
  'Remove from memory
  Set Cipher1 = Nothing
   
  ' End the Decryption, unfortunately this is doomed to fail,
  ' because of the missing padding bytes
  'Cipher2.FinishDecryptBLOB
  'Cipher1.FinishDecryptBLOB
   
  'Remove padding
  InputData = RemovePadding(InputData)
  
  'Publish the result
  DecryptByte = InputData
  Exit Function
  
ErrorHandle:
  MsgBox "Error during byte array decryption"
End Function


Public Function EncryptStr( _
      ByVal ASCIIText As String, _
      ByVal hexSKey As String, _
      ByVal hexSIV As String, _
      ByVal hexRKey As String, _
      ByVal hexRIV As String) As String
      
  'Double encryption ... first Serpent and than Rijndael, both using 256-bit keys
  'Double encryption an overkill ? -> YES! (Who cares!)
        
  Dim Cipher1 As New eb_c_IncrementalCipher
  Dim Cipher2 As New eb_c_IncrementalCipher
  Dim InputData() As Byte
  Dim strOutput As String
  
  InputData = StrConv(ASCIIText, vbFromUnicode)
      
  'Add padding
  InputData = AddPadding(InputData)
      
  'Start the encryption
  Call Cipher1.StartEncryptRaw(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, hexSKey, hexSIV)
    
  InputData = Cipher1.EncryptBLOB(InputData)
      
  ' End the encryption
  Cipher1.FinishEncrypt
      
  Call Cipher2.StartEncryptRaw(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, hexRKey, hexRIV)
      
  InputData = Cipher2.EncryptBLOB(InputData)
        
  ' End the encryption
  Cipher2.FinishEncrypt
      
  'Convert to string
  strOutput = StrConv(InputData, vbUnicode)
      
  'Remove from memory
  Set Cipher1 = Nothing
  Set Cipher2 = Nothing
  Erase InputData
  
  'Publish the result
  EncryptStr = strOutput
  Exit Function
  
ErrorHandle:
  EncryptStr = "Error with Encryption!"
  
End Function

Public Function DecryptStr( _
        ByVal ASCIIText As String, _
        ByVal hexSKey As String, _
        ByVal hexSIV As String, _
        ByVal hexRKey As String, _
        ByVal hexRIV As String) As String
    
  Dim Cipher2 As New eb_c_IncrementalCipher
  Dim Cipher1 As New eb_c_IncrementalCipher
  Dim InputData() As Byte
  Dim strOutput As String
  
  On Error GoTo ErrorHandle
  
  'For decryption padding bytes are needed, this is an error in the
  'ebCrypt.dll or maybe in the OpenSSL package, on which it is based
  '16 bytes = 128 bit
      
  InputData = StrConv(ASCIIText, vbFromUnicode)
  ReDim Preserve InputData(UBound(InputData) + 16) As Byte
    
  Call Cipher2.StartDecryptRaw(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, hexRKey, hexRIV)
  InputData = Cipher2.DecryptBLOB(InputData)
      
  'Add padding again
  ReDim Preserve InputData(UBound(InputData) + 16) As Byte
          
  'Remove from memory
  Set Cipher2 = Nothing 'This is one way to finish decryption :)
  
  Call Cipher1.StartDecryptRaw(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, hexSKey, hexSIV)
  InputData = Cipher1.DecryptBLOB(InputData)
         
  'Remove from memory
  Set Cipher1 = Nothing
   
  ' End the Decryption, unfortunately this is doomed to fail,
  ' because of the missing padding bytes
  'Cipher2.FinishDecryptBLOB
  'Cipher1.FinishDecryptBLOB
   
  'Remove padding
  InputData = RemovePadding(InputData)
  
  'Convert bytearray to string
  strOutput = StrConv(InputData, vbUnicode)
  
  'Remove from memory
  Erase InputData
  
  'Publish the result
  DecryptStr = strOutput
  Exit Function
  
ErrorHandle:
  DecryptStr = "Error with decryption!"
End Function

Public Function EncryptStrRAW( _
      ByVal ASCIIText As String, _
      ByVal hexSKey As String, _
      ByVal hexSIV As String, _
      ByVal hexRKey As String, _
      ByVal hexRIV As String) As String
      
  'Double encryption ... first Serpent and than Rijndael, both using 256-bit keys
  'Double encryption an overkill ? -> YES! (Who cares!)
        
  Dim Cipher1 As New eb_c_IncrementalCipher
  Dim Cipher2 As New eb_c_IncrementalCipher
  Dim InputData() As Byte
  Dim strOutput As String
  
  InputData = StrConv(ASCIIText, vbFromUnicode)
        
  'Start the encryption
  Call Cipher1.StartEncryptRaw(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, hexSKey, hexSIV)
    
  InputData = Cipher1.EncryptBLOB(InputData)
      
  ' End the encryption
  Cipher1.FinishEncrypt
      
  Call Cipher2.StartEncryptRaw(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, hexRKey, hexRIV)
      
  InputData = Cipher2.EncryptBLOB(InputData)
        
  ' End the encryption
  Cipher2.FinishEncrypt
      
  'Convert to string
  strOutput = StrConv(InputData, vbUnicode)
      
  'Remove from memory
  Set Cipher1 = Nothing
  Set Cipher2 = Nothing
  Erase InputData
  
  'Publish the result
  EncryptStrRAW = strOutput
  Exit Function
  
ErrorHandle:
  EncryptStrRAW = "Error with Encryption!"
  
End Function

Public Function DecryptStrRAW( _
        ByVal ASCIIText As String, _
        ByVal hexSKey As String, _
        ByVal hexSIV As String, _
        ByVal hexRKey As String, _
        ByVal hexRIV As String) As String
    
  Dim Cipher2 As New eb_c_IncrementalCipher
  Dim Cipher1 As New eb_c_IncrementalCipher
  Dim InputData() As Byte
  Dim strOutput As String
  
  On Error GoTo ErrorHandle
  
  'For decryption padding bytes are needed, this is an error in the
  'ebCrypt.dll or maybe in the OpenSSL package, on which it is based
  '16 bytes = 128 bit
      
  InputData = StrConv(ASCIIText, vbFromUnicode)
  ReDim Preserve InputData(UBound(InputData) + 16) As Byte
    
  Call Cipher2.StartDecryptRaw(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, hexRKey, hexRIV)
  InputData = Cipher2.DecryptBLOB(InputData)
      
  'Add padding again
  ReDim Preserve InputData(UBound(InputData) + 16) As Byte
          
  'Remove from memory
  Set Cipher2 = Nothing 'This is one way to finish decryption :)
  
  Call Cipher1.StartDecryptRaw(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, hexSKey, hexSIV)
  InputData = Cipher1.DecryptBLOB(InputData)
         
  'Remove from memory
  Set Cipher1 = Nothing
   
  ' End the Decryption, unfortunately this is doomed to fail,
  ' because of the missing padding bytes
  'Cipher2.FinishDecryptBLOB
  'Cipher1.FinishDecryptBLOB
   
   'Convert bytearray to string
  strOutput = StrConv(InputData, vbUnicode)
  
  'Remove from memory
  Erase InputData
  
  'Publish the result
  DecryptStrRAW = strOutput
  Exit Function
  
ErrorHandle:
  DecryptStrRAW = "Error with decryption!"
End Function



Public Function GetRSAKeys() As RSAKeys
  Dim RSA As New eb_c_RSAKey
  Dim Keys As RSAKeys
    
  'Don't use public exponent 3!
  'RSA Modulus 2304-bit, is 18 x 128-bit
  'Good size when using 128-bit symmetric block ciphers!
  
  RSA.GenerateKey EB_CRYPT_RSA_KEY_EXPONENT_10001, 2304
 
  Keys.hPrivateKey = RSA.ExportPrivateKey(EB_CRYPT_EXPORT_FORMAT_DER, _
    EB_CRYPT_CIPHER_ALGORITHM_NONE, "")
    
  Keys.hPublicKey = RSA.ExportPublicKey(EB_CRYPT_EXPORT_FORMAT_DER, _
    EB_CRYPT_CIPHER_ALGORITHM_NONE, "")
    
  ' Remove from memory
  Set RSA = Nothing
  
  'Publish result
  GetRSAKeys = Keys
  
End Function

Public Function FixPublicKeyForEncryption(ByVal hPublicKey As String) As String
   ' ONLY FOR 2304-bit RSA KEYS !!!
   ' The exporting DER-format for publickeys has a fixed header and trailer.
   ' To avoid known-plaintext-attacks, the header and trailer are removed
   ' when encrypting the publickey.
   
   ' HEADER = 3082012A0282012100 (only for 2304-bit modulus, 1024-bit has a different and smaller header)
   ' TRAILER = 0203010001

   Dim sOutput As String
   
   'Remove trailer
   sOutput = Left$(hPublicKey, Len(hPublicKey) - 10)
   'Remove Header
   sOutput = Right$(sOutput, Len(sOutput) - 18)
         
   'Public the result
   FixPublicKeyForEncryption = sOutput
End Function

Public Function FixPublicKeyAfterDecryption(ByVal hPublicKey As String) As String
   ' This is the reverse function of "FixPublicKeyForEncryption"
   
   Dim sOutput As String
   Dim Header As String
   Dim Trailer As String
   
   Header = "3082012A0282012100"
   Trailer = "0203010001"
   
   ' Add Header and Trailer back again!
   sOutput = Header & hPublicKey & Trailer
   
   'Publish the resuls
   FixPublicKeyAfterDecryption = sOutput
End Function

Public Function RSAPublicEncryptToHex(ByVal hexInput As String, ByVal hexPublicKey As String) As String
  Dim RSA As New eb_c_RSAKey
  Dim strOutput As String
  
  Call RSA.ImportPublicKey(EB_CRYPT_EXPORT_FORMAT_DER, "", hexPublicKey)
  strOutput = RSA.PublicEncryptEx(hexInput, EB_CRYPT_RSA_PAD_PKCS1)
  
  'Remove from memory
  Set RSA = Nothing
  
  'Publish the result
  RSAPublicEncryptToHex = strOutput

End Function

Public Function RSAPublicDecryptToHex(ByVal hexInput As String, ByVal hexPublicKey As String) As String
  Dim RSA As New eb_c_RSAKey
  Dim strOutput As String
  
  Call RSA.ImportPublicKey(EB_CRYPT_EXPORT_FORMAT_DER, "", hexPublicKey)
  strOutput = RSA.PublicDecryptEx(hexInput, EB_CRYPT_RSA_PAD_PKCS1)
  
  'Remove from memory
  Set RSA = Nothing
  
  'Publish the result
  RSAPublicDecryptToHex = strOutput

End Function

Public Function RSAPrivateEncryptToHex(ByVal hexInput As String, ByVal hexPrivateKey As String) As String
  Dim RSA As New eb_c_RSAKey
  Dim strOutput As String
  
  Call RSA.ImportPrivateKey(EB_CRYPT_EXPORT_FORMAT_DER, "", hexPrivateKey)
  strOutput = RSA.PrivateEncryptEx(hexInput, EB_CRYPT_RSA_PAD_PKCS1)
  
  'Remove from memory
  Set RSA = Nothing
  
  'Publish the result
  RSAPrivateEncryptToHex = strOutput

End Function

Public Function RSAPrivateDecryptToHex(ByVal hexInput As String, ByVal hexPrivateKey As String) As String
  Dim RSA As New eb_c_RSAKey
  Dim hexOutput As String
  
  Call RSA.ImportPrivateKey(EB_CRYPT_EXPORT_FORMAT_DER, "", hexPrivateKey)
  hexOutput = RSA.PrivateDecryptEx(hexInput, EB_CRYPT_RSA_PAD_PKCS1)
  
  'Remove from memory
  Set RSA = Nothing
  
  'Publish the result
  RSAPrivateDecryptToHex = hexOutput
  
End Function

'
' ------------------------------------------------------------------------------
' EOF.
' ------------------------------------------------------------------------------
'


