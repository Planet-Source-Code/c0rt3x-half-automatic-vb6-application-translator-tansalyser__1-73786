Attribute VB_Name = "mString"
Option Explicit

Function ContainsKey(Text As String, Key As String) As Long
On Error GoTo Error
    
    Dim i As Long
    
    For i = 1 To Len(Text)
        If Mid(Text, i, Len(Key)) = Key Then
            ContainsKey = ContainsKey + 1
        End If
    Next i
    
Exit Function
Error:
    Assert , "StringModule.ContainsKey", Err.Number, Err.Description, "Text: '" & Text & "', Key: '" & Key & "'"
    Resume Next
End Function

Function SplitText(Text As String, sWord() As String) As Long
On Error GoTo Error

    Dim sDelimiters As String
    Dim i As Long, d As Long
    Dim lPoint As Long
    Dim lWords
    Dim sText As String
    
    sText = Text & Chr(0)
    For i = 0 To 255
        If InStr(1, "abcdefghijklmnopqrstuvwxyzäöüß", LCase(Chr(i))) = 0 Then
            sDelimiters = sDelimiters & Chr(i)
        End If
    Next i
    ReDim sWord(0)
    For i = 1 To Len(sText)
        If i > lPoint Then
            If InStr(1, sDelimiters, Mid(sText, i, 1)) = 0 Then
                For d = i To Len(sText)
                    If InStr(1, sDelimiters, Mid(sText, d, 1)) <> 0 Then
                        lWords = lWords + 1
                        ReDim Preserve sWord(lWords + 1)
                        sWord(lWords) = Mid(sText, i, d - i)
                        lPoint = d
                        Exit For
                    End If
                Next d
            End If
        End If
    Next i
    SplitText = lWords

Exit Function
Error:
    Assert , "StringModule.SplitText", Err.Number, Err.Description, "Text: '" & Text & "'"
    Resume Next
End Function

Public Function MultiStr(sString As String, lMultiplier As Long)
On Error GoTo Error

    Dim i As Long
    
    For i = 1 To lMultiplier
        MultiStr = MultiStr & sString
    Next i

Exit Function
Error:
    Assert , "StringModule.MultiStr", Err.Number, Err.Description, "sString: '" & sString & "', lMultiplier: '" & lMultiplier & "'"
    Resume Next
End Function

Public Function NewLine(lMultiplier As Long) As String
On Error GoTo Error
    
    NewLine = MultiStr(vbNewLine, lMultiplier)
    
Exit Function
Error:
    Assert , "StringModule.NewLine", Err.Number, Err.Description, "lMultiplier: '" & lMultiplier & "'"
    Resume Next
End Function

Function FilterString(Str As String, Junk As String) As String
On Error GoTo Error

    Dim i As Long
    
    For i = 1 To Len(Str)
        If InStr(1, Junk, Mid(Str, i, 1)) = 0 Then
            FilterString = FilterString & Mid(Str, i, 1)
        End If
    Next i
    
Exit Function
Error:
    Assert , "StringModule.FilterString", Err.Number, Err.Description, "Str: '" & Str & "', Junk:'" & Junk & "'"
    Resume Next
End Function

Function CompareString(String1 As String, String2 As String) As Double
On Error GoTo Error

    Dim i&
    Dim s1$, s2$
    Dim CharValue As Double
    
    s1 = String1
    s2 = String2
    If Len(s1) > Len(s2) Then
        s2 = s2 & MultiStr(Chr(255), Len(s1) - Len(s2))
    ElseIf Len(s1) < Len(s2) Then
        s1 = s1 & MultiStr(Chr(255), Len(s2) - Len(s1))
    End If
    CharValue = 100 / Len(s1)
    For i = 1 To Len(s1)
        If Mid(s1, i, 1) = Mid(s2, i, 1) Then
            CompareString = CompareString + CharValue
        ElseIf LCase(Mid(s1, i, 1)) = LCase(Mid(s2, i, 1)) Then
            CompareString = CompareString + (CharValue / 2)
        End If
    Next i
    
Exit Function
Error:
    Assert , "StringModule.CompareString", Err.Number, Err.Description, "String1: '" & String1 & "', String2:'" & String2 & "'"
    Resume Next
End Function


Function StrBetween(Str As String, Char1 As String, Char2 As String) As String
    
    Dim a As Long
    Dim z As Long
    
    a = InStr(1, Str, Char1)
    z = InStr(a + 1, Str, Char2)
    
    StrBetween = Mid(Str, a + 1, (z - 1) - a)
    
End Function

'Public Function NewLines(Count As Long) As String
'
'    Dim i As Long
'    Dim s As String
'
'    For i = 1 To Count
'
'        s = s & vbNewLine
'
'    Next i
'
'    NewLines = s
'
'End Function
