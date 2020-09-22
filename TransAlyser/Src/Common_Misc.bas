Attribute VB_Name = "mMisc"
Option Explicit

Function Pause(Delay As Double)
On Error GoTo Error
    Dim Start As Double
    If Delay < 0 Then Exit Function
    Start = Timer
    Do While Timer < Start + Delay
        DoEvents
    Loop
Exit Function
Error:
    Assert , "MiscModule.Pause", Err.Number, Err.Description, "Delay: '" & CStr(Delay) & "'"
    Resume Next
End Function

Function IsNumber(Number) As Boolean
Dim Test As Double
On Error GoTo Error
    Test = CDbl(Number)
    IsNumber = True
Exit Function
Error:

End Function

Function CodeLineCount(Str) As Long
    Dim l() As String
    Dim i As Long
    l = Split(Str, vbNewLine)
    For i = 0 To UBound(l)
        If Trim(l(i)) <> "" Then CodeLineCount = CodeLineCount + 1
    Next i
End Function

Function GetDirCodeLineCount(BaseDirPath As String) As Long
    Dim f() As String
    Dim c As Long
    Dim i As Long
    Dim l As Long
    c = GetFileList(BaseDirPath, f)
    For i = 1 To c
        Select Case (Right(f(i), 4))
            Case ".bas", ".cls", ".frm", ".ctl", ".VBP"
                l = l + CodeLineCount(ReadFile(f(i)))
            Case Else
                'DO NOTHING
        End Select
    Next i
    GetDirCodeLineCount = l
End Function

