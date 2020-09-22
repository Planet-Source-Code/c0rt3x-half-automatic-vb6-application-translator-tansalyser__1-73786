Attribute VB_Name = "mPseudoRnd"
Option Explicit

Function GetKeyNum(Key As String, Optional Min As Long = 0, Optional Max As Long = 255) As Long
    
    Dim i As Long
    Dim Num As Long
    Dim n As Long
    
    If Min <> Max Then
    
        For i = 1 To Len(Key)
            Num = Num + CLng(Asc(Mid(Key, i, 1)))
        Next i
    
        n = Min + (Num Mod Max)
    
    Else
        
        n = Max ' = Max
        
    End If
    
    
    GetKeyNum = n
    
End Function

Sub testkeynums()

    Dim i As Long
    Dim r As Long
    
    Dim Key As String
    
    Dim s As String
    Dim s2 As String
    
    Dim Arr() As Long
    
    Key = CStr(Timer) '"abcde"
    
    GetKeySequence 8, Key, Arr()
    
    For i = 1 To UBound(Arr)
        
        s = s & CStr(Arr(i))
        
    Next i
    
    
    GetKeySequence 8, Key, Arr()
    
    For i = 1 To UBound(Arr)
        
        s2 = s2 & CStr(Arr(i))
        
    Next i
    
    Debug.Print s, (s = s2)
    
End Sub

Function DelArrVal(Arr() As Long, ValIndex As Long)
    
    Dim i As Long
    
    For i = ValIndex To UBound(Arr) - 1
        Arr(i) = Arr(i + 1)
    Next i
    
    ReDim Preserve Arr(LBound(Arr) To (UBound(Arr) - 1))
    
End Function

Function GetKeySequence(Numbers As Long, Key As String, RetArr() As Long)
    
    Dim i As Long
    Dim x As Long
    
    Dim ArrNum() As Long
    
    Dim IndexKey As String
    Dim MinIndex As Long
    Dim MaxIndex As Long
    
    ReDim ArrNum(1 To Numbers)
    ReDim RetArr(1 To Numbers)
    
    For i = 1 To Numbers
        ArrNum(i) = i
    Next i
    
    For i = 1 To Numbers
        
        'IndexKey = CStr(Asc(CStr(i))) & Key & CStr(i)
        MinIndex = 1
        MaxIndex = UBound(ArrNum)
        
        
        x = GetKeyNum(Key, MinIndex, MaxIndex)
        
        RetArr(i) = ArrNum(x)
        
        If UBound(ArrNum) > 1 Then
            DelArrVal ArrNum, x
        End If
        
    Next i
    
End Function

