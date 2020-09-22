Attribute VB_Name = "mConversion"
Option Explicit

Public Function ConvertBase10ToX(lNumber&, lBaseOut&, lResult&())
'On Error GoTo Error

    Dim lNum&
    Dim lCount&
    Dim lRest&()
    Dim i&


    lNum = lNumber
    
    Do
        lCount = lCount + 1
        ReDim Preserve lRest(1 To lCount)
        lRest(lCount) = lNum Mod lBaseOut
        lNum = lNum \ lBaseOut
    Loop While lNum > 0
    
    ReDim lResult(1 To lCount)
    
    For i = 0 To lCount - 1
        lResult(lCount - i) = lRest(i + 1)
    Next i
    
'Exit Function
'Error:
'    Assert , "ConversionModule.ConvertBase10ToX", Err.Number, Err.Description, "lNumber: '" & CStr(lNumber) & "', lBaseOut: '" & CStr(lBaseOut) & "'"
'    Resume Next
End Function

Public Function ConvertBaseXTo10(lNumber&(), lBaseIn&) As Long
'On Error GoTo Error
    Dim i&
    Dim lFactor&
    For i = LBound(lNumber) To UBound(lNumber)
        lFactor = lNumber(i) + lFactor * lBaseIn
    Next i
    ConvertBaseXTo10 = lFactor
    
'Exit Function
'Error:
'    lFactor = 0
'    Assert , "ConversionModule.ConvertBaseXTo10", Err.Number, Err.Description, "lBaseIn: '" & CStr(lBaseIn) & "'"
'   Resume Next
End Function

Public Function NumberToString(lNumber As Long, Optional lBase As Long = 254) As String
'On Error GoTo Error
    Dim lResult&()
    Dim i&
    ConvertBase10ToX lNumber, lBase, lResult()
    For i = 1 To UBound(lResult)
        NumberToString = NumberToString & Chr(lResult(i))
    Next i
    
'Exit Function
'Error:
'    Assert , "ConversionModule.NumberToString", Err.Number, Err.Description, "lNumber: '" & CStr(lNumber) & "', cstr(lBase): '" & lBase & "'"
'    Resume Next
End Function

Public Function StringToNumber(sNumber As String, Optional lBase As Long = 254)
'On Error GoTo Error
    Dim lNumber&()
    Dim i&
    ReDim lNumber(Len(sNumber))
    For i = 1 To Len(sNumber)
        lNumber(i) = Asc(Mid(sNumber, i, 1))
    Next i
    StringToNumber = ConvertBaseXTo10(lNumber(), lBase)
    
'Exit Function
'Error:
'    Assert , "ConversionModule.StringToNumber", Err.Number, Err.Description, "sNumber: '" & sNumber & "', lBase: '" & CStr(lBase) & "'"
'    Resume Next
End Function

Function b2v(bValue As Boolean) As Variant
On Error GoTo Error
    b2v = Abs(CInt(bValue))
Exit Function
Error:
    Assert , "ConversionModule.b2v", Err.Number, Err.Description, "bValue: '" & CStr(bValue) & "'"
    Resume Next
End Function

Function v2b(vValue As Variant) As Boolean
On Error GoTo Error
    v2b = CBool(vValue)
Exit Function
Error:
    Assert , "ConversionModule.v2b", Err.Number, Err.Description, "vValue: '" & CStr(vValue) & "'"
    Resume Next
End Function

Public Function String2Bytes(sString As String, Bytes() As Byte)
On Error GoTo Error
    Dim i As Double
    ReDim Bytes(1 To Len(sString))
    For i = 1 To Len(sString)
        Bytes(i) = AscB(Mid(sString, i, 1))
    Next i
Exit Function
Error:
    Assert , "ConversionModule.String2Bytes", Err.Number, Err.Description, "sString: '" & sString & "'"
    Resume Next
End Function

Public Function Bytes2String(Bytes() As Byte) As String
On Error GoTo Error
    Dim i As Double
    For i = LBound(Bytes) To UBound(Bytes)
        Bytes2String = Bytes2String & Chr(Bytes(i))
    Next i
Exit Function
Error:
    Assert , "ConversionModule.Bytes2String", Err.Number, Err.Description
    Resume Next
End Function

Function IsDate(Str As String) As Boolean
On Error GoTo Error
    Dim d As Date
    d = CDate(Str)
    IsDate = True
Exit Function
Error:
    
End Function

Function Seconds2Time(Seconds As Long) As String
On Error GoTo Error
    
    Dim Value() As Long
    
    ConvertBase10ToX Seconds, 60, Value()
    Select Case UBound(Value)
        Case 1
            Seconds2Time = "0:00:" & Format(Value(1), "00")
        Case 2
            Seconds2Time = "0:" & Format(Value(1), "00") & ":" & Format(Value(2), "00")
        Case 3
            Seconds2Time = CStr(Value(1)) & ":" & Format(Value(2), "00") & ":" & Format(Value(3), "00")
    End Select

Exit Function
Error:
    Assert , "TimeModule.Seconds2Time", Err.Number, Err.Description, "Seconds: '" & Seconds & "'"
    Resume Next
End Function

'Function StringToFont(FontString As String, Font As StdFont)
'Dim sFont() As String
'Dim fFont As New StdFont
'On Error GoTo Error
'    sFont() = Split(FontString, "|")
'    fFont.Name = sFont(0)
'    If UBound(sFont()) > 0 Then fFont.Size = CInt(sFont(1))
'    If UBound(sFont()) > 1 Then fFont.Bold = CBool(sFont(2))
'    If UBound(sFont()) > 2 Then fFont.Italic = CBool(sFont(3))
'    If UBound(sFont()) > 3 Then fFont.Underline = CBool(sFont(4))
'    If UBound(sFont()) > 4 Then fFont.Strikethrough = CBool(sFont(5))
'    If UBound(sFont()) > 5 Then fFont.Weight = sFont(6)
'    If UBound(sFont()) > 6 Then fFont.Charset = sFont(7)
'    Set Font = fFont
'Exit Function
'Error:
'   Assert , "ConversionModule.StringToFont", Err.Number, Err.Description, "FontString: '" & FontString & "'"
'End Function

'Function FontToString(Font As StdFont) As String
'On Error GoTo Error
'    FontToString = Font.Name & "|"
'    FontToString = FontToString & CInt(Font.Size) & "|"
'    FontToString = FontToString & Abs(CInt(Font.Bold)) & "|"
'    FontToString = FontToString & Abs(CInt(Font.Italic)) & "|"
'    FontToString = FontToString & Abs(CInt(Font.Underline)) & "|"
'    FontToString = FontToString & Abs(CInt(Font.Strikethrough)) & "|"
'    FontToString = FontToString & Font.Weight & "|"
'    FontToString = FontToString & Font.Charset
'Exit Function
'Error:
'   Assert , "ConversionModule.FontToString", Err.Number, Err.Description
'End Function
