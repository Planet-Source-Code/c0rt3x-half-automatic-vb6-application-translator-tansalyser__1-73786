Attribute VB_Name = "mRandom"
Function RandomName(MaxLen As Long) As String
    Dim l As Long
    Dim i As Long
    l = RandomNumber(1, MaxLen)
    For i = 1 To l
        If i = 1 Then
            RandomName = RandomName & UCase(RandomLetter)
        Else
            RandomName = RandomName & RandomLetter
        End If
        
    Next i
End Function

Function RandomLetter()
    RandomLetter = Chr(RandomNumber(97, 122))
End Function

Function RandomNumber(Optional Min As Long = 0, Optional Max As Long = 1, Optional AllowFraction As Boolean) As Double
On Error GoTo Error
    
    Randomize Timer
    
    If AllowFraction Then
        RandomNumber = (Rnd * (Max - Min)) + Min
    Else
        RandomNumber = Int((Rnd * (Max - Min)) + Min + 0.5)
    End If

Exit Function
Error:
    Assert , "RandomModule.RandomNumber", Err.Number, Err.Description, "Min: '" & CStr(Min) & "', Max: '" & CStr(Max) & "', AllowFraction: '" & CStr(AllowFraction) & "'"
    Resume Next
End Function

Public Function RandomHex(Lenght As Long) As String
On Error GoTo Error

    Dim i&
    
    For i = 1 To Lenght
        RandomHex = RandomHex & Hex(RandomNumber(0, 15))
    Next i
    
Exit Function
Error:
    Assert , "RandomModule.RandomHex", Err.Number, Err.Description, "Lenght: '" & CStr(Lenght) & "'"
    Resume Next
End Function

Public Function RandomString(Lenght As Long) As String
On Error GoTo Error
    
    Dim i&
    
    For i = 1 To Lenght
        RandomString = RandomString & Chr(RandomNumber(0, 255))
    Next i
    
Exit Function
Error:
    Assert , "RandomModule.RandomString", Err.Number, Err.Description, "Lenght: '" & CStr(Lenght) & "'"
    Resume Next
End Function

Public Function RandomText(Lenght As Long) As String
On Error GoTo Error
    
    Dim i As Long
    
    For i = 1 To Lenght
        Select Case RandomNumber(1, 3)
            Case 1
                RandomText = RandomText & Chr(RandomNumber(97, 122))
            Case 2
                RandomText = RandomText & Chr(RandomNumber(65, 90))
            Case 3
                RandomText = RandomText & CStr(RandomNumber(0, 9))
        End Select
    Next i

Exit Function
Error:
    Assert , "RandomModule.RandomText", Err.Number, Err.Description, "Lenght: '" & CStr(Lenght) & "'"
    Resume Next
End Function

Function RandomFileName(Optional Extension$ = ".tmp", Optional Lenght% = 8) As String
On Error GoTo Error
    
    RandomFileName = RandomText(CLng(Lenght)) & Extension

Exit Function
Error:
    Assert , "RandomModule.RandomFileName", Err.Number, Err.Description, "Extension: '" & Extension & "', Lenght: '" & CStr(Lenght) & "'"
    Resume Next
End Function

Public Function RandomOrder(vArray() As Variant)
On Error GoTo Error

    Dim i As Long, n As Long
    Dim vTmp() As Variant
    Dim vTmp2() As Variant
    Dim lNumber As Long
    Dim lTmp
    
    vTmp = vArray
    For i = LBound(vArray) To UBound(vArray)
        lNumber = RandomNumber(LBound(vTmp), UBound(vTmp))
        vArray(i) = vTmp(lNumber)
        For n = lNumber To UBound(vTmp) - 1
            vTmp(n) = vTmp(n + 1)
        Next n
        lTmp = UBound(vTmp) - 1
        If i < UBound(vArray) Then
            ReDim vTmp2(LBound(vTmp) To lTmp)
            For n = LBound(vTmp) To lTmp
                vTmp2(n) = vTmp(n)
            Next n
            ReDim vTmp(LBound(vTmp) To lTmp)
            vTmp = vTmp2
        End If
    Next i

Exit Function
Error:
    Assert , "RandomModule.RandomOrder", Err.Number, Err.Description
    Resume Next
End Function

Public Function lRandomOrder(lArray() As Long)
On Error GoTo Error

    Dim i As Long, n As Long
    Dim vTmp() As Long
    Dim vTmp2() As Long
    Dim lNumber As Long
    Dim lTmp
    
    vTmp = lArray
    For i = LBound(lArray) To UBound(lArray)
        lNumber = RandomNumber(LBound(vTmp), UBound(vTmp))
        lArray(i) = vTmp(lNumber)
        For n = lNumber To UBound(vTmp) - 1
            vTmp(n) = vTmp(n + 1)
        Next n
        lTmp = UBound(vTmp) - 1
        If i < UBound(lArray) Then
            ReDim vTmp2(LBound(vTmp) To lTmp)
            For n = LBound(vTmp) To lTmp
                vTmp2(n) = vTmp(n)
            Next n
            ReDim vTmp(LBound(vTmp) To lTmp)
            vTmp = vTmp2
        End If
    Next i

Exit Function
Error:
    Assert , "RandomModule.lRandomOrder", Err.Number, Err.Description
    Resume Next
End Function

Function Likelihood(Likelihoods() As Double) As Long
    Dim i As Long
    Dim Sum As Double
    Dim p As Double
    For i = LBound(Likelihoods()) To UBound(Likelihoods())
        Sum = Sum + Likelihoods(i)
    Next i
    p = RandomNumber(0, CLng(Sum), True)
    Sum = 0
    For i = LBound(Likelihoods()) To UBound(Likelihoods())
        Sum = Sum + Likelihoods(i)
        If Sum >= p Then
            Likelihood = i
            Exit Function
        End If
    Next i
    Likelihood = UBound(Likelihoods())
End Function

Function LikelihoodTest()
    Dim a(1 To 3) As Double
    Dim r(1 To 3) As Long
    Dim rtn As Long
    Dim i As Long
    
    For i = 1 To 1000
        'a(i) = RandomNumber(0, 100, True)
        a(1) = 10
        a(2) = 100
        a(3) = 1000
        
        rtn = Likelihood(a)
        
        r(rtn) = r(rtn) + 1
        
    Next i
    
    Debug.Print r(1), r(2), r(3)
    
End Function
