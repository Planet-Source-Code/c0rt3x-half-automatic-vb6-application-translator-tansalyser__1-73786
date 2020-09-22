Attribute VB_Name = "mStrings"
Option Explicit

Private Const Base64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789%&"

Public Function Base64Encode(Text As String) As String
On Error GoTo Error

    Dim c1, c2, c3 As Integer
    Dim w1 As Integer
    Dim w2 As Integer
    Dim w3 As Integer
    Dim w4 As Integer
    Dim n As Integer
    Dim retry As String
    
    For n = 1 To Len(Text) Step 3
        c1 = Asc(Mid$(Text, n, 1))
        c2 = Asc(Mid$(Text, n + 1, 1) + Chr$(0))
        c3 = Asc(Mid$(Text, n + 2, 1) + Chr$(0))
        w1 = Int(c1 / 4)
        w2 = (c1 And 3) * 16 + Int(c2 / 16)
        If Len(Text) >= n + 1 Then w3 = (c2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
        If Len(Text) >= n + 2 Then w4 = c3 And 63 Else w4 = -1
        retry = retry + MimeEncode(w1) + MimeEncode(w2) + MimeEncode(w3) + MimeEncode(w4)
    Next
    
    Base64Encode = retry

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function

Public Function Base64Decode(Text As String) As String
On Error GoTo Error

    Dim w1 As Integer
    Dim w2 As Integer
    Dim w3 As Integer
    Dim w4 As Integer
    Dim n As Integer
    Dim retry As String
    
    For n = 1 To Len(Text) Step 4
        w1 = MimeDecode(Mid$(Text, n, 1))
        w2 = MimeDecode(Mid$(Text, n + 1, 1))
        w3 = MimeDecode(Mid$(Text, n + 2, 1))
        w4 = MimeDecode(Mid$(Text, n + 3, 1))
        If w2 >= 0 Then retry = retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
        If w3 >= 0 Then retry = retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
        If w4 >= 0 Then retry = retry + Chr$(((w3 * 64 + w4) And 255))
    Next
    
    Base64Decode = retry

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function

Private Function MimeEncode(W As Integer, Optional ByVal Key As String) As String
On Error GoTo Error

    If IsMissing(Key) Then Key = Base64
    If W >= 0 Then MimeEncode = Mid$(Base64, W + 1, 1) Else MimeEncode = ""
    
Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function

Private Function MimeDecode(a As String, Optional ByVal Key) As Integer
On Error GoTo Error

    If IsMissing(Key) Then Key = Base64
    If Len(a) = 0 Then MimeDecode = -1: Exit Function
    MimeDecode = InStr(Base64, a) - 1

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function

