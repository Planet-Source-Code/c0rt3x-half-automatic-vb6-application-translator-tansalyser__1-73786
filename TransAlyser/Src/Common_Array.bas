Attribute VB_Name = "mArray"
Option Explicit

Public Function lDeleteArrayValue(lArray() As Long, lValue As Long) As Long
On Error GoTo Error

    Dim i&, n&
    
    For i = LBound(lArray) To UBound(lArray)
        If lArray(i) = lValue Then
            For n = i To UBound(lArray) - 1
                lArray(n) = lArray(n) + 1
            Next n
            ReDim Preserve lArray(LBound(lArray) To UBound(lArray) - 1)
        End If
    Next i

Exit Function
Error:
    Assert , "ArrayModule.lDeleteArrayValue", Err.Number, Err.Description, "lValue: '" & CStr(lValue) & "'"
    Resume Next
End Function

Public Function lSumOfArray(lArray&()) As Long
On Error GoTo Error

    Dim i&
    
    For i = LBound(lArray) To UBound(lArray)
        lSumOfArray = lSumOfArray + lArray(i)
    Next i
    
Exit Function
Error:
    Assert , "ArrayModule.lSumOfArray", Err.Number, Err.Description
    Resume Next
End Function

Public Function sArraySize(sArray$()) As Long
On Error GoTo Error
    
    Dim i As Long
    
    For i = LBound(sArray) To UBound(sArray)
        sArraySize = sArraySize + 1
    Next i
    
Exit Function
Error:
    Assert , "ArrayModule.sArraySize", Err.Number, Err.Description
    Resume Next
End Function

Public Function JoinArray(sArray$()) As String
On Error GoTo Error

    Dim lSize&
    Dim i&
    Dim lIndex&
    Dim sTOC$
    Dim sData$
    lSize = sArraySize(sArray)
    For i = 0 To lSize - 1
        lIndex = LBound(sArray) + i
        sTOC = sTOC & NumberToString(Len(sArray(lIndex))) & Chr(254 + Abs(CInt(i = lSize - 1)))
        sData = sData & sArray(lIndex)
    Next i
    JoinArray = sTOC & sData
    
Exit Function
Error:
    Assert , "ArrayModule.JoinArray", Err.Number, Err.Description
    Resume Next
End Function

Public Function SplitArray(sString$, sArray$())
On Error GoTo Error
    
    Dim lLenght&()
    Dim i&
    Dim lCount&
    Dim lStart&
    Dim ChrNumber&
    lStart = 1
    For i = 1 To Len(sString$)
        ChrNumber = Asc(Mid(sString$, i, 1))
        If (ChrNumber = 254) Or (ChrNumber = 255) Then
            lCount = lCount + 1
            ReDim Preserve lLenght(1 To lCount)
            lLenght(lCount) = StringToNumber((Mid(sString$, lStart, i - lStart)))
            lStart = i + 1
            If ChrNumber = 255 Then
                Exit For
            End If
        End If
    Next i
    ReDim Preserve sArray(1 To lCount)
    For i = 1 To lCount
        sArray(i) = Mid(sString$, lStart, lLenght(i))
        lStart = lStart + lLenght(i)
    Next i
    
Exit Function
Error:
    Assert , "ArrayModule.SplitArray", Err.Number, Err.Description
    Resume Next
End Function

Public Function JoinArrayRev(sArray$())
On Error GoTo Error
    
    Dim lSize&
    Dim i&
    Dim lIndex&
    Dim sTOC$
    Dim sData$
    lSize = sArraySize(sArray)
    For i = 0 To lSize - 1
        lIndex = LBound(sArray) + i
        sTOC = Chr(254 + Abs(CInt(i = lSize - 1))) & NumberToString(Len(sArray(lIndex))) & sTOC
        sData = sData & sArray(lIndex)
    Next i
    JoinArrayRev = sData & sTOC
    
Exit Function
Error:
    Assert , "ArrayModule.JoinArrayRev", Err.Number, Err.Description
    Resume Next
End Function

Public Function SplitArrayRev(sString$, sArray$())
On Error GoTo Error
    
    Dim lLenght&()
    Dim c&, i&
    Dim lCount&
    Dim lStart&
    Dim lLen&
    Dim bAsc As Byte
    
    For c = 0 To Len(sString) - 1
        i = Len(sString) - c
        bAsc = CByte(Asc(Mid(sString, i, 1)))
        If (bAsc = 254) Or (bAsc = 255) Then
            lCount = lCount + 1
            ReDim Preserve lLenght(1 To lCount)
            lLenght(lCount) = StringToNumber(Mid(sString, i + 1, lLen))
            lLen = 0
            If bAsc = 255 Then
                Exit For
            End If
        Else
            lLen = lLen + 1
        End If
    Next c
    ReDim Preserve sArray(1 To lCount)
    lStart = 1
    For i = 1 To lCount
        sArray(i) = Mid(sString$, lStart, lLenght(i))
        lStart = lStart + lLenght(i)
    Next i
    
Exit Function
Error:
    Assert , "ArrayModule.SplitArrayRev", Err.Number, Err.Description
    Resume Next
End Function

Public Function vArraySize(vArray()) As Long
On Error GoTo Error

    Dim i As Long
    
    For i = LBound(vArray) To UBound(vArray)
        vArraySize = vArraySize + 1
    Next i

Exit Function
Error:
    Assert , "ArrayModule.vArraySize", Err.Number, Err.Description
    Resume Next
End Function

Public Function lConvertArray(sArray$(), lArray&()) As Long
On Error GoTo Error
    
    Dim i&
    
    lConvertArray = CLng(sArray(LBound(sArray)))
    ReDim lArray(LBound(sArray) To UBound(sArray))
    
    For i = LBound(sArray) To UBound(sArray)
        lArray(i) = CLng(sArray(i))
    Next i

Exit Function
Error:
    Assert , "ArrayModule.lConvertArray", Err.Number, Err.Description
    Resume Next
End Function

Public Function sConvertArray(vArray(), sArray()) As String
On Error GoTo Error
    
    Dim i&
    sConvertArray = CStr(vArray(LBound(vArray)))
    ReDim sArray(LBound(vArray) To UBound(vArray))
    
    For i = LBound(vArray) To UBound(vArray)
        sArray(i) = CStr(vArray(i))
    Next i

Exit Function
Error:
    Assert , "ArrayModule.sConvertArray", Err.Number, Err.Description
    Resume Next
End Function

Public Function vConvertArray(sArray(), vArray()) As Variant
On Error GoTo Error

    Dim i&
    
    vConvertArray = CVar(sArray(LBound(sArray)))
    ReDim vArray(LBound(sArray) To UBound(sArray))
    
    For i = LBound(sArray) To UBound(sArray)
        vArray(i) = CVar(sArray(i))
    Next i

Exit Function
Error:
    Assert , "ArrayModule.vConvertArray", Err.Number, Err.Description
    Resume Next
End Function

Sub SortArray(a() As Variant)
On Error GoTo Error

    Dim u&, i&, j&, K&, H As Variant
    
    u = UBound(a)
    K = u \ 2
    While K > 0
        For i = 0 To u - K
            j = i
            While (j >= 0) And (a(j) > a(j + K))
                H = a(j)
                a(j) = a(j + K)
                a(j + K) = H
                If j > K Then
                    j = j - K
                Else
                    j = 0
                End If
            Wend
        Next i
        K = K \ 2
    Wend

Exit Sub
Error:
    Assert , "ArrayModule.SortArray", Err.Number, Err.Description
    Resume Next
End Sub

Sub SortStrArray(a() As String)
On Error GoTo Error

    Dim u&, i&, j&, K&, H As String
    
    If UBound(a) = 1 Then
        If a(0) > a(1) Then
            H = a(0)
            a(0) = a(1)
            a(1) = H
        End If
        Exit Sub
    End If
    
    u = UBound(a)
    K = u \ 2
    While K > 0
        For i = 0 To u - K
            j = i
            While (j >= 0) And (a(j) > a(j + K))
                H = a(j)
                a(j) = a(j + K)
                a(j + K) = H
                If j > K Then
                    j = j - K
                Else
                    j = 0
                End If
            Wend
        Next i
        K = K \ 2
    Wend

Exit Sub
Error:
    Assert , "ArrayModule.SortStrArray", Err.Number, Err.Description
    Resume Next
End Sub

Function InArray(Value As Variant, a() As Variant) As Long
On Error GoTo Error

    Dim i&
    
    For i = LBound(a) To UBound(a)
        
        If Value = a(i) Then
            
            InArray = i
            Exit Function
        
        End If
    
    Next i

Exit Function
Error:
    Assert , "ArrayModule.InArray", Err.Number, Err.Description, "Value: '" & CStr(Value) & "'"
    Resume Next
End Function

Function InStrArray(Value As String, a() As String) As Long
On Error GoTo Error

    Dim i&
    
    For i = LBound(a) To UBound(a)
        
        If LCase(Value) = LCase(a(i)) Then
            
            InStrArray = i
            Exit Function
        
        End If
    
    Next i
    
    InStrArray = -1

Exit Function
Error:
    Assert , "ArrayModule.InArray", Err.Number, Err.Description, "Value: '" & CStr(Value) & "'"
    Resume Next
End Function
