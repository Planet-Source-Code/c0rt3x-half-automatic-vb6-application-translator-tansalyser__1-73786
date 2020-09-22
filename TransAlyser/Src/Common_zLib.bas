Attribute VB_Name = "Common_ZLibModule"
Private Declare Function zLib_Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function zLib_Decompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private Const ZLIB_NOERROR = 0

Public Function Compress(TheString As String) As String
On Error GoTo Error
    
    Dim lResult As Long
    Dim lCmpSize As Long
    Dim sTBuff As String
    Dim lOriginalSize As Long
    
    lOriginalSize = Len(TheString)
    lCmpSize = lOriginalSize
    lCmpSize = lCmpSize + (lCmpSize * 0.01) + 12
    sTBuff = String$(lCmpSize, 0)
    lResult = zLib_Compress(ByVal sTBuff, lCmpSize, ByVal TheString, lOriginalSize)
    If lResult = ZLIB_NOERROR Then
        Compress = NumberToString(lOriginalSize, 255) & Chr(255) & Left$(sTBuff, lCmpSize)
        sTBuff = ""
    Else
        GoTo Error
    End If
    
    Exit Function
Error:
    Assert , "ZLibModule.Compress", Err.Number, Err.Description
    Resume Next
End Function

Public Function Decompress(TheString As String) As String
On Error GoTo Error
    
    Dim lResult As Long
    Dim lCmpSize As Long
    Dim sTBuff As String
    Dim lCompressedSize
    Dim OrigSize As Long
    Dim lEnd As Long
    Dim sCompressedString As String
    
    lEnd = InStr(1, TheString, Chr(255))
    OrigSize = StringToNumber((Left(TheString, lEnd - 1)), 255)
    sCompressedString = Right(TheString, Len(TheString) - lEnd)
    lCompressedSize = Len(sCompressedString)
    sTBuff = String$(OrigSize + 1, 0)
    lCmpSize = Len(sTBuff)
    lResult = zLib_Decompress(ByVal sTBuff, lCmpSize, ByVal sCompressedString, lCompressedSize)
    If lResult = ZLIB_NOERROR Then
        sCompressedString = Left$(sTBuff, lCmpSize)
    Else
        sCompressedString = Left$(sTBuff, lCmpSize)
        GoTo Error
    End If
    Decompress = Left$(sTBuff, lCmpSize)
    
    Exit Function
Error:
    Assert , "ZLibModule.Decompress", Err.Number, Err.Description
    If InIDE Then Resume Next
End Function
