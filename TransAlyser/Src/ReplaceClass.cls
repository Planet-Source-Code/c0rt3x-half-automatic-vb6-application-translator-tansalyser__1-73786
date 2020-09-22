VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplaceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' by Olaf Schmidt, os@datenhaus.de, 20010106
Option Explicit

Private src%(), saSrc&(5)
Private Fnd%(), saFnd&(5)
Private Rep%(), saRep&(5)
Private Out%(), saOut&(5)
Private PosArr&(), UBPosArr&

' VB5 -> msvbvm50.dll
Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)
Private Declare Function SysAllocStringByteLen& Lib "oleaut32" (ByVal olestr&, ByVal BLen&)

Friend Function ReplaceEx(Text As String, sOld As String, sNew As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Count As Long = 2147483647, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As String
  Dim lenSrc&, lenFnd&, LenRep&, LenOut&
  Dim i&, j&, InPos&, OutPos&, CFnd&, Dist&, LCmp&, LFnd&, Fnd0%
  
  lenSrc = Len(Text)
  If lenSrc = 0 Then Exit Function
  
  lenFnd = Len(sOld): LenRep = Len(sNew)
  If lenFnd = 0 Then ReplaceEx = Text: Exit Function
  
  saRep(3) = StrPtr(sNew)
  saSrc(3) = StrPtr(Text)
  saFnd(3) = StrPtr(sOld): Fnd0 = Fnd(0)
  
  If lenFnd = LenRep Then
    RtlMoveMemory ByVal VarPtr(ReplaceEx), SysAllocStringByteLen(saSrc(3), lenSrc + lenSrc), 4
    saOut(3) = StrPtr(ReplaceEx)
  End If
  
  If Compare = vbBinaryCompare Then
    
    For i = Start - 1 To lenSrc - 1
      If src(i) <> Fnd0 Then 'Inline-Cascading for first Char
        i = i + 1
        If src(i) <> Fnd0 Then
          i = i + 1
          If src(i) <> Fnd0 Then
            i = i + 1
            If src(i) <> Fnd0 Then
              i = i + 1
              If src(i) <> Fnd0 Then
                i = i + 1
                If src(i) <> Fnd0 Then
                  i = i + 1
                  If src(i) <> Fnd0 Then
                    i = i + 1
                    If src(i) <> Fnd0 Then GoTo nxt_i
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
  
      For j = 1 To lenFnd - 1 'Search all others
        If src(i + j) <> Fnd(j) Then GoTo nxt_i
      Next j
  
      'Found at Position i (0 based)
      If i >= lenSrc Then Exit For
      CFnd = CFnd + 1
      If lenFnd = LenRep Then
        For j = 0 To LenRep - 1: Out(i + j) = Rep(j): Next j
      Else
        If CFnd > UBPosArr Then
          ReDim Preserve PosArr(UBPosArr + 512): UBPosArr = UBound(PosArr)
        End If
        PosArr(CFnd) = i
      End If
      If CFnd = Count Then Exit For
      i = i + lenFnd - 1
nxt_i: Next i

  Else 'vbStringCompare
    
    If Fnd0 > 64& And Fnd0 < 91& Or Fnd0 > 191& And Fnd0 < 223& Then Fnd0 = Fnd0 + 32&
    
    For i = Start - 1 To lenSrc - 1
      LCmp = src(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
      If LCmp <> Fnd0 Then  'Inline-Cascading for first Char
        i = i + 1: LCmp = src(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
        If LCmp <> Fnd0 Then
          i = i + 1: LCmp = src(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
          If LCmp <> Fnd0 Then
            i = i + 1: LCmp = src(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
            If LCmp <> Fnd0 Then GoTo nxt_ii
          End If
        End If
      End If
  
      For j = 0 To lenFnd - 1 'Search all others
        LCmp = src(i + j): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
        LFnd = Fnd(j): If LFnd > 64& And LFnd < 91& Or LFnd > 191& And LFnd < 223& Then LFnd = LFnd + 32&
        If LCmp <> LFnd Then GoTo nxt_ii
      Next j
      
      'Found at Position i (0 based)
      If i >= lenSrc Then Exit For
      CFnd = CFnd + 1
      If lenFnd = LenRep Then
        For j = 0 To LenRep - 1: Out(i + j) = Rep(j): Next j
      Else
        If CFnd > UBPosArr Then
          ReDim Preserve PosArr(UBPosArr + 512): UBPosArr = UBound(PosArr)
        End If
        PosArr(CFnd) = i
      End If
      If CFnd = Count Then Exit For
      i = i + lenFnd - 1
nxt_ii: Next i
  End If
  
  'Generate Output
  If lenFnd <> LenRep Then
    If CFnd = 0 Then
      ReplaceEx = Text
    Else
      LenOut = lenSrc + (LenRep - lenFnd) * CFnd
      RtlMoveMemory ByVal VarPtr(ReplaceEx), SysAllocStringByteLen(0, LenOut + LenOut), 4
      saOut(3) = StrPtr(ReplaceEx)
      
      OutPos = 0: InPos = 0
      For i = 1 To CFnd
        Dist = PosArr(i) - InPos
        If Dist > 100 Then
          RtlMoveMemory Out(OutPos), src(InPos), Dist + Dist
        ElseIf Dist > 0 Then
          j = 0
          Do 'Inline-Cascading
            Out(OutPos + j) = src(InPos + j): j = j + 1: If j = Dist Then Exit Do
            Out(OutPos + j) = src(InPos + j): j = j + 1: If j = Dist Then Exit Do
            Out(OutPos + j) = src(InPos + j): j = j + 1: If j = Dist Then Exit Do
            Out(OutPos + j) = src(InPos + j): j = j + 1: If j = Dist Then Exit Do
          Loop
        End If
        OutPos = OutPos + Dist
        InPos = PosArr(i) + lenFnd

        If LenRep > 100 Then
          RtlMoveMemory Out(OutPos), Rep(0), LenRep + LenRep
        ElseIf LenRep > 0 Then
          j = 0
          Do 'Inline-Cascading
            Out(OutPos + j) = Rep(j): j = j + 1: If j = LenRep Then Exit Do
            Out(OutPos + j) = Rep(j): j = j + 1: If j = LenRep Then Exit Do
            Out(OutPos + j) = Rep(j): j = j + 1: If j = LenRep Then Exit Do
            Out(OutPos + j) = Rep(j): j = j + 1: If j = LenRep Then Exit Do
          Loop
        End If
        OutPos = OutPos + LenRep
      Next i
      If (lenSrc - InPos) > 0 Then
        RtlMoveMemory Out(OutPos), src(InPos), (lenSrc - InPos) + (lenSrc - InPos)
      End If
    End If
  End If
End Function

Private Sub Class_Initialize()
  ReDim PosArr(512): UBPosArr = UBound(PosArr)
  
  saSrc(0) = 1: saSrc(1) = 2: saSrc(4) = 2147483647
  RtlMoveMemory ByVal ArrPtr(src), VarPtr(saSrc(0)), 4
  
  saFnd(0) = 1: saFnd(1) = 2: saFnd(4) = 2147483647
  RtlMoveMemory ByVal ArrPtr(Fnd), VarPtr(saFnd(0)), 4
    
  saRep(0) = 1: saRep(1) = 2: saRep(4) = 2147483647
  RtlMoveMemory ByVal ArrPtr(Rep), VarPtr(saRep(0)), 4
  
  saOut(0) = 1: saOut(1) = 2: saOut(4) = 2147483647
  RtlMoveMemory ByVal ArrPtr(Out), VarPtr(saOut(0)), 4
End Sub

Private Sub Class_Terminate()
  RtlMoveMemory ByVal ArrPtr(src), 0&, 4
  RtlMoveMemory ByVal ArrPtr(Fnd), 0&, 4
  RtlMoveMemory ByVal ArrPtr(Rep), 0&, 4
  RtlMoveMemory ByVal ArrPtr(Out), 0&, 4
End Sub


