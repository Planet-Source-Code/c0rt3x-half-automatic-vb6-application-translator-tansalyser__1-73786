Attribute VB_Name = "VBStringDecryptionModule"
Option Explicit

Private Const VBStringBase64Key As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789%&"


Public Function VBStringBase64Decode(Text As String) As String
On Error GoTo Error

    Dim w1 As Integer
    Dim w2 As Integer
    Dim w3 As Integer
    Dim w4 As Integer
    Dim n As Integer
    Dim retry As String
    
    For n = 1 To Len(Text) Step 4
        w1 = VBStringMimeDecode(Mid$(Text, n, 1))
        w2 = VBStringMimeDecode(Mid$(Text, n + 1, 1))
        w3 = VBStringMimeDecode(Mid$(Text, n + 2, 1))
        w4 = VBStringMimeDecode(Mid$(Text, n + 3, 1))
        If w2 >= 0 Then retry = retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
        If w3 >= 0 Then retry = retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
        If w4 >= 0 Then retry = retry + Chr$(((w3 * 64 + w4) And 255))
    Next
    
    VBStringBase64Decode = retry

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function


Private Function VBStringMimeDecode(a As String) As Integer
On Error GoTo Error

    If Len(a) = 0 Then
        VBStringMimeDecode = -1
        Exit Function
    End If
    VBStringMimeDecode = InStr(VBStringBase64Key, a) - 1

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function
