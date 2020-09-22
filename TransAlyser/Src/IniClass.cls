VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSys_Ini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function GetValue(sINI As String, Section As Variant, Key As Variant) As String
On Error GoTo Error

Dim sSection$, sKey$
Dim lSectionStart&, lSectionEnd&, lKeyStart&, lValueEnd&

    sSection = "[" & CStr(Section) & "]"
    sKey = CStr(Key) & "="
    lSectionStart = InStr(1, sINI, sSection)
    If lSectionStart < 1 Then GoTo Error
    lSectionStart = lSectionStart + Len(sSection)
    lKeyStart = InStr(lSectionStart, sINI, vbNewLine & sKey) + 2
    lSectionEnd = InStr(lSectionStart + 1, sINI, "[")
    If lSectionEnd < 1 Then lSectionEnd = Len(sINI)
    'If (lKeyStart < 3) Or (lKeyStart > lSectionEnd) Then GoTo Error
    lValueEnd = InStr(lKeyStart, sINI & vbNewLine, vbNewLine)
    If lValueEnd < 1 Then GoTo Error
    GetValue = Mid(sINI, lKeyStart + Len(sKey), lValueEnd - (lKeyStart + Len(sKey)))

Exit Function
Error:
    GetValue = ""
    'Assert "cSys_Ini.GetValue", Err.Number, Err.Description, "INI: '" & sINI & "', Section: '" & Section & "', " & "Key: '" & Key & "'"
    'Resume Next
End Function

Function SetValue(sINI As String, Section As Variant, Key As Variant, Value As Variant)
On Error GoTo Error

    Dim pINI$, sSection$, sKey$, sValue$, Line$(), sSectionChunk$, sLChunk$, sRChunk$
    Dim lSectionStart&, lSectionEnd&, lKeyStart&, lValueStart&, lValueEnd&, i&

    sSection = "[" & CStr(Section) & "]"
    sKey = vbNewLine & CStr(Key) & "="
    sValue = CStr(Value)
    pINI = sINI
    If InStr(1, pINI, sSection) < 1 Then
        If pINI <> "" Then
            Do While Right(pINI, 4) <> vbNewLine & vbNewLine
                pINI = pINI & vbNewLine
            Loop
        End If
        pINI = pINI & sSection
    End If
    lSectionStart = InStr(1, pINI, sSection) + Len(sSection)
    lSectionEnd = lSectionStart
    sSectionChunk = Mid(pINI, lSectionStart)
    Line = Split(sSectionChunk, vbNewLine)
    For i = LBound(Line()) To UBound(Line())
        If Left(Line(i), 1) = "[" And Right(Line(i), 1) = "]" Then Exit For
        If Len(Line(i)) > 0 Then
            lSectionEnd = lSectionEnd + Len(Line(i)) + 2
        End If
    Next i
    lKeyStart = InStr(lSectionStart, pINI, sKey)
    If lKeyStart < 1 Or lKeyStart > lSectionEnd Then
        sLChunk = Left(pINI, lSectionEnd - 1)
        sRChunk = Mid(pINI, lSectionEnd)
        If Right(sLChunk, 2) = vbNewLine Then sLChunk = Left(sLChunk, Len(sLChunk) - 2)
        If Left(sRChunk, 2) <> vbNewLine Then sRChunk = vbNewLine & sRChunk
        If Left(sRChunk, 3) = vbNewLine & "[" Then sRChunk = vbNewLine & sRChunk
        pINI = sLChunk & sKey & sRChunk
        lKeyStart = InStr(lSectionStart, pINI, sKey)
    End If
    lValueStart = lKeyStart + Len(sKey)
    lValueEnd = InStr(lValueStart, pINI, vbNewLine)
    If lValueEnd < 1 Then lValueEnd = lValueStart + 1 '
    sLChunk = Left(pINI, lValueStart - 1) '
    sRChunk = Mid(pINI, lValueEnd)
    pINI = sLChunk & sValue & sRChunk
    Do While Right(pINI, 2) = vbNewLine
        pINI = Left(pINI, Len(pINI) - 2)
    Loop
    pINI = pINI & vbNewLine
    sINI = pINI
    
Exit Function
Error:
    Debug.Print "cSys_Ini.SetValue", Err.Number, Err.Description, "INI: '" & sINI & "', Section: '" & Section & "'; " & "Key: '" & Key & "'; " & "Value: '" & Value & "'"
    Resume Next
End Function
