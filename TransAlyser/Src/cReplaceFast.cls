VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReplaceFast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'******************************************************************************
'*
'* Original by Olaf Schmidt, www.datenhaus.de, 2001.01.06
'* Modified by Ulrich Korndörfer, www.prosource.de
'*
'* (c) 2002 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Based on CReplace10.cls, which has been contributed to the VB speed contest
'* website at www.xbeat.net\vbspeed and there it was the clear winner in the
'* category Replace.
'*
'* External dependencies: none except some APIs (see below)
'*
'* Version history
'*
'* Version 1.1 from 2003.09.18 (disclaimer added, comments added)
'*
'* Version 1.0 from 2002.09.15
'*
'* Replaced type declaration characters
'* Changed naming, started beautifying, added and corrected comments
'* Did some minor code changes
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************


Option Explicit


'************************************************************************
'* APIs
'************************************************************************

Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" _
               (ByRef ArrVar() As Any) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" _
               (ByRef Dest As Any, ByRef Source As Any, ByVal NumBytes As Long)

Private Declare Function SysAllocStringByteLen Lib "oleaut32" _
               (ByVal OleStr As Long, ByVal ByteLen As Long) As Long


'************************************************************************
'* Private members and consts
'************************************************************************


'The following dummy arrays are used to fake a pointer like access to
'strings. If proper set to a string, the individual characters of the
'string can be accessed by adressing an array element.
'As UNICODE16 strings are to be accessed, which do have two byte characters,
'the arrays must be of type integer, which has two bytes too.
'The benefit of doing so is very fast access to the individual characters,
'compared to traditional access using functions like Mid$.

'For example, if everything is proper set up for a string "Source", access is
'possible in the following way:

'SourceArr(0) -> first character of "Source" (char at position 1 of the string)

'To change the value of an individual character:

'SourceArr(i) = AscW("x")    (versus Mid$(Source, i + 1, 1) = "x")
'SourceArr(i) = &H78         (versus Mid$(Source, i + 1, 1) = ChrW$(&H78))

'To copy a character from one string to another:

'ResultArr(i) = SourceArr(j) (versus Mid$(Result, i + 1, 1) = Mid$(Source, j + 1, 1))

'To search for a character (for example search for an "x"):

'If SourceArr(i) = &H78 Then (versus if mid$(Source,i+1,1) = "x")

Private SourceArr()  As Integer 'Gives access to the string Source
Private FindArr()    As Integer 'Gives access to the string Find
Private ReplaceArr() As Integer 'Gives access to the string Replace
Private ResultArr()  As Integer 'Gives access to the string ReplaceFast, the
                                'implicit function var, which reveives the result
                                'of the function (the replaced string)

'For proper setting up the fake we need SafeArray structures.
'For short one can use an array too, such faking a SafeArray structure.

'Official structure for an onedimensional array and its mapping to the
'faking arrays:

'Public Type SafeArray1D
'  cDims       As Integer 'SA(0) high word, setup to 1
'  fFeatures   As Integer 'SA(0) low word,  setup to 0
'  cbElements  As Long    'SA(1),           setup to 2 (integer type)
'  cLocks      As Long    'SA(2),           setup to 0
'  pvData      As Long    'SA(3),           setup to string data (StrPtr(string))
'  cElements   As Long    'SA(4),           setup to 2147483647
'  lLBound     As Long    'SA(5),           setup to 0
'End Type

Private SourceSA(0 To 5)  As Long 'SafeArray fake for SourceArr
Private FindSA(0 To 5)    As Long 'SafeArray fake for FindArr
Private ReplaceSA(0 To 5) As Long 'SafeArray fake for ReplaceArr
Private ResultSA(0 To 5)  As Long 'SafeArray fake for ResultArr

Private FoundPos() As Long 'Array holding all starting positions for all occurrences of
                           'the Find string
Private FoundPosMaxIndex As Long 'UBound(FoundPos)

Private Const CHUNKSIZE As Long = 512 'Step size for increasing FoundPos


'************************************************************************
'* Public main function
'************************************************************************

Public Function ReplaceFast(ByRef Source As String, _
                            ByRef Find As String, _
                            ByRef Replace As String, _
                            Optional ByVal Start As Long = 1, _
                            Optional ByVal Count As Long = 2147483647, _
                            Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
                          ) As String

Dim SourceLen As Long, FindLen As Long, ReplaceLen As Long
Dim FindCount As Long, FindFirstChar As Integer
Dim SourcePos As Long, ResultPos As Long
Dim Dist As Long, LCmp As Long, LFnd As Long
Dim i As Long, j As Long

SourceLen = Len(Source)

'Source is empty, nothing to do, implicitly return null string
If SourceLen = 0 Then Exit Function

FindLen = Len(Find): ReplaceLen = Len(Replace)

'Nothing to find, so nothing to replace, return the source
If FindLen = 0 Then ReplaceFast = Source: Exit Function

'Make input strings pointer-like accessible through dummy arrays
SourceSA(3) = StrPtr(Source)
FindSA(3) = StrPtr(Find)
ReplaceSA(3) = StrPtr(Replace)

'When later the occurrences of the Find string are searched, speed is mandatory.
'The speedy algorithm used often needs the first char of the Find string, so
'set FindFirstChar to the first char of the Find string.
FindFirstChar = FindArr(0)

'If Find string and Replace string have the same length, the search can do
'an in place replace during searching. Therefore copy the Source string to
'the implicit function var ReplaceFast (the result string).
If FindLen = ReplaceLen Then
  'Originally following assignement was used:
  'RtlMoveMemory ByVal VarPtr(ReplaceFast), SysAllocStringByteLen(SourceSA(3), SourceLen + SourceLen), 4
  'As this is just a copy, it is replaced by the following, which is faster:
  ReplaceFast = Source
  'Make result string pointer-like accessible through dummy array
  ResultSA(3) = StrPtr(ReplaceFast)
End If

'Now search for all occurrences of the Find string in Source.
'This is done using separate algorithms for the two different compare modes.
'If ReplaceLen = FindLen, replace during search in place (in the Result string)
'the Find substring with the Replace substring.
'If ReplaceLen <> FindLen, remember the starting positions of all found occurrences
'of the Find string in the Source string using the array FoundPos. Later
'use these positions to replace the Find substring with the Replace substring.

If Compare = vbBinaryCompare Then

  For i = Start - 1 To SourceLen - 1

    If SourceArr(i) <> FindFirstChar Then 'Inline-Cascading for first char
      i = i + 1
      If SourceArr(i) <> FindFirstChar Then
        i = i + 1
        If SourceArr(i) <> FindFirstChar Then
          i = i + 1
          If SourceArr(i) <> FindFirstChar Then
            i = i + 1
            If SourceArr(i) <> FindFirstChar Then
              i = i + 1
              If SourceArr(i) <> FindFirstChar Then
                i = i + 1
                If SourceArr(i) <> FindFirstChar Then
                  i = i + 1
                  If SourceArr(i) <> FindFirstChar Then GoTo Continue_BinaryCompare
                End If
              End If
            End If
          End If
        End If
      End If
    End If

    For j = 1 To FindLen - 1 'Search all others
      If SourceArr(i + j) <> FindArr(j) Then GoTo Continue_BinaryCompare
    Next j

    'Found at Position i (0 based)

    If i >= SourceLen Then Exit For
    FindCount = FindCount + 1

    If FindLen = ReplaceLen Then

      'Replace now
      For j = 0 To ReplaceLen - 1: ResultArr(i + j) = ReplaceArr(j): Next j

    Else

      'First check for enough room in FoundPos, making FoundPos greater if necessary
      If FindCount > FoundPosMaxIndex Then
        FoundPosMaxIndex = FoundPosMaxIndex + CHUNKSIZE
        ReDim Preserve FoundPos(0 To FoundPosMaxIndex)
      End If
      'Remember found position in FoundPos and replace later
      FoundPos(FindCount) = i

    End If

    If FindCount = Count Then Exit For
    i = i + FindLen - 1

Continue_BinaryCompare:

  Next i


Else 'Compare = vbBinaryCompare (now assume vbTextCompare mode)

  If FindFirstChar > 64& And FindFirstChar < 91& Or FindFirstChar > 191& And FindFirstChar < 223& Then FindFirstChar = FindFirstChar + 32&

  For i = Start - 1 To SourceLen - 1

    LCmp = SourceArr(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
    If LCmp <> FindFirstChar Then  'Inline-Cascading for first Char
      i = i + 1: LCmp = SourceArr(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
      If LCmp <> FindFirstChar Then
        i = i + 1: LCmp = SourceArr(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
        If LCmp <> FindFirstChar Then
          i = i + 1: LCmp = SourceArr(i): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
          If LCmp <> FindFirstChar Then GoTo Continue_TextCompare
        End If
      End If
    End If

    For j = 0 To FindLen - 1 'Search all others
      LCmp = SourceArr(i + j): If LCmp > 64& And LCmp < 91& Or LCmp > 191& And LCmp < 223& Then LCmp = LCmp + 32&
      LFnd = FindArr(j): If LFnd > 64& And LFnd < 91& Or LFnd > 191& And LFnd < 223& Then LFnd = LFnd + 32&
      If LCmp <> LFnd Then GoTo Continue_TextCompare
    Next j

    'Found at Position i (0 based)
    'Following code is identical to the equivalent portion in vbBinaryCompare mode

    If i >= SourceLen Then Exit For
    FindCount = FindCount + 1

    If FindLen = ReplaceLen Then

      'Replace now
      For j = 0 To ReplaceLen - 1: ResultArr(i + j) = ReplaceArr(j): Next j

    Else

      'First check for enough room in FoundPos, making FoundPos greater if necessary
      If FindCount > FoundPosMaxIndex Then
        FoundPosMaxIndex = FoundPosMaxIndex + CHUNKSIZE
        ReDim Preserve FoundPos(0 To FoundPosMaxIndex)
      End If
      'Remember found position in FoundPos and replace later
      FoundPos(FindCount) = i

    End If

    If FindCount = Count Then Exit For
    i = i + FindLen - 1

Continue_TextCompare:

  Next i

End If


'If FindLen = ReplaceLen, we are finished at this point.

'If FindLen <> ReplaceLen, FoundPos now contains all positions where the Find
'string has been found. Now generate the result string by using these positions
'and insert the Replace string at these positions.

If FindLen <> ReplaceLen Then

  If FindCount = 0 Then

    'Nothing found, so nothing to replace: just copy the source

    ReplaceFast = Source

  Else 'FindCount <> 0

    'Replace the Find substrings with the Replace substrings

    RtlMoveMemory ByVal VarPtr(ReplaceFast), _
                  SysAllocStringByteLen(0, (SourceLen + (ReplaceLen - FindLen) * FindCount) * 2), 4
    ResultSA(3) = StrPtr(ReplaceFast)

    ResultPos = 0: SourcePos = 0

    For i = 1 To FindCount

      Dist = FoundPos(i) - SourcePos

      If Dist > 100 Then

        RtlMoveMemory ResultArr(ResultPos), SourceArr(SourcePos), Dist + Dist

      ElseIf Dist > 0 Then

        j = 0
        Do 'Inline-Cascading
          ResultArr(ResultPos + j) = SourceArr(SourcePos + j): j = j + 1: If j = Dist Then Exit Do
          ResultArr(ResultPos + j) = SourceArr(SourcePos + j): j = j + 1: If j = Dist Then Exit Do
          ResultArr(ResultPos + j) = SourceArr(SourcePos + j): j = j + 1: If j = Dist Then Exit Do
          ResultArr(ResultPos + j) = SourceArr(SourcePos + j): j = j + 1: If j = Dist Then Exit Do
        Loop

      End If

      ResultPos = ResultPos + Dist
      SourcePos = FoundPos(i) + FindLen

      If ReplaceLen > 100 Then

        RtlMoveMemory ResultArr(ResultPos), ReplaceArr(0), ReplaceLen + ReplaceLen

      ElseIf ReplaceLen > 0 Then

        j = 0
        Do 'Inline-Cascading
          ResultArr(ResultPos + j) = ReplaceArr(j): j = j + 1: If j = ReplaceLen Then Exit Do
          ResultArr(ResultPos + j) = ReplaceArr(j): j = j + 1: If j = ReplaceLen Then Exit Do
          ResultArr(ResultPos + j) = ReplaceArr(j): j = j + 1: If j = ReplaceLen Then Exit Do
          ResultArr(ResultPos + j) = ReplaceArr(j): j = j + 1: If j = ReplaceLen Then Exit Do
        Loop

      End If

      ResultPos = ResultPos + ReplaceLen

    Next i

    If (SourceLen - SourcePos) > 0 Then
      RtlMoveMemory ResultArr(ResultPos), SourceArr(SourcePos), (SourceLen - SourcePos) + (SourceLen - SourcePos)
    End If

  End If 'FindCount = 0

End If 'FindLen <> ReplaceLen

'Clean up

If FoundPosMaxIndex > CHUNKSIZE Then
  FoundPosMaxIndex = CHUNKSIZE
  ReDim FoundPos(0 To FoundPosMaxIndex)
End If

End Function


'************************************************************************
'* Class de/construction
'************************************************************************

Private Sub Class_Initialize()

'Setup starting space for the position array. Index 0 is never used

FoundPosMaxIndex = CHUNKSIZE
ReDim FoundPos(0 To FoundPosMaxIndex)

'Now prepare the faked pointer access to the strings

'First prepare the fake SafeArray structures

SourceSA(0) = 1: SourceSA(1) = 2: SourceSA(4) = 2147483647
FindSA(0) = 1: FindSA(1) = 2: FindSA(4) = 2147483647
ReplaceSA(0) = 1: ReplaceSA(1) = 2: ReplaceSA(4) = 2147483647
ResultSA(0) = 1: ResultSA(1) = 2: ResultSA(4) = 2147483647

'Then let the array vars of the dummy arrays point to them
'(make them their safe array structs). Until now, the array vars
'of the dummy arrays point to nothing (the dummy array vars value is 0).

RtlMoveMemory ByVal ArrPtr(SourceArr), VarPtr(SourceSA(0)), 4
RtlMoveMemory ByVal ArrPtr(FindArr), VarPtr(FindSA(0)), 4
RtlMoveMemory ByVal ArrPtr(ReplaceArr), VarPtr(ReplaceSA(0)), 4
RtlMoveMemory ByVal ArrPtr(ResultArr), VarPtr(ResultSA(0)), 4

'Thats (nearly) all! Really nice trick.

'The dummy arrays safe arrays fakes actually do not point to string data
'and such are useless in this state.
'To make an dummy array actually point to a string, eg. to a string
'named "Source", the element with index 3 of its safearray fake must
'be set to the adress of the strings data. This would be done so:

'SourceSA(3) = StrPtr(Source)

'This can be repeated so often as necessary. No cleanup is needed
'during calls to ReplaceFast. Only when the class terminates
'a little clean up is necessary.

End Sub

Private Sub Class_Terminate()

'We must clean up, before the class finishes.
'As the dummy array vars initially pointed to zero,
'we must reestablish this. If we would not, VB would think,
'there is allocated memory for the array data, which has to be freed.
'The pointer to this array data (which resides in the safe array fake
'at element index 3) immediately after setup in the class Initialize-event
'pointed to zero, and later may have been set (by using the ReplaceFast-function)
'to point to some string data. In both cases it points to memory, which
'the arrays did not allocate themselves. So they should not ask for deallocation.
'If they would - bang!

'Setting the dummy arrays array vars to zero lets VB think that the arrays
'never have been used or have been used and erased afterwards.

RtlMoveMemory ByVal ArrPtr(SourceArr), 0&, 4
RtlMoveMemory ByVal ArrPtr(FindArr), 0&, 4
RtlMoveMemory ByVal ArrPtr(ReplaceArr), 0&, 4
RtlMoveMemory ByVal ArrPtr(ResultArr), 0&, 4

End Sub
