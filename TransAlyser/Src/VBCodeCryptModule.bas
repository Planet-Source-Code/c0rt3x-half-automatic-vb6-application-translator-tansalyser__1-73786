Attribute VB_Name = "VBTransAlysertModule"
Option Explicit

Public My As MyApp



Public Enum VBAliasTypes
    
    vbFile = 2 ^ 1
    
    VBProject = 2 ^ 2
    VBModule = 2 ^ 3
    
    VBfrm = 2 ^ 4
    VBbas = 2 ^ 5
    VBcls = 2 ^ 6
    VBctl = 2 ^ 7
    VBpag = 2 ^ 8
    VBdsr = 2 ^ 9
    
    VBControl = 2 ^ 10
    VBVariable = 2 ^ 11
    VBDeclaration = 2 ^ 12
    VBEvent = 2 ^ 13
    
    VBWithEvents = 2 ^ 14
    
    VBProcedure = 2 ^ 15
    
    VBSub = 2 ^ 16
    VBFunction = 2 ^ 17
    VBProperty = 2 ^ 18
    
    VBParameter = 2 ^ 19
    VBControlProperty = 2 ^ 20
    VBType = 2 ^ 21
    
End Enum

Public Enum VBScopes
    VBLocal = 2 ^ 1
    VBPrivate = 2 ^ 2
    VBPublic = 2 ^ 3
End Enum

Public Enum VBModuleCodeBlocks
    A_Head = 1
    B_Properties = 2
    C_Attributes = 3
    D_Declinations = 4
    E_Subs = 5
End Enum

Public Type VBQuote
    StartPos As Long
    EndPos As Long
    Length As Long
    Quote As String
End Type

Public Enum VBSpecialModuleTypes
    [TranslationModule] = 1
    [TranslationForm] = 2
End Enum


Sub Main()
    
    Set My = New MyApp
    
    My.Start
    
End Sub

Sub Shutdown()
    
    Set My = Nothing
    
    End
    
End Sub

'Public Type VBAlias
'    ID                  As Long
'    Name                As String
'    NewName             As String
'    AliasType           As VBAliasTypes
'    Scope               As VBScopes
'    ModuleID            As Long
'    SubID               As Long
'    VariableID          As Long
'End Type

'Public Type VBCodeBlock
'
'    ID                  As Long
'    ParentID            As Long
'
'    ChildCount          As Long
'    ChildIDs()          As Long
'
'    Code                As String
'    Protected           As Boolean
'
'    StartPos            As Long
'    EndPos              As Long
'    Length              As Long
'
'End Type

'Public Type VBControl
'
'    ID                  As Long
'    ModuleID            As Long
'    BlockID             As Long
'    AliasID             As Long
'
'    Name                As String
'
'End Type

'Public Type VBSub
'
'    ID                  As Long
'    ModuleID            As Long
'    AliasID             As Long
'    BlockID             As Long
'
'    Name                As String
'    NewName             As String
'
'    SubType             As String
'    Scope               As VBScopes
'
'    FirstLine           As Long
'    LastLine            As Long
'    Code                As String
'
'    VariableCount       As Long
'    Variables()         As VBAlias
'
'End Type

'Public Type VBModule
'
'    ID                  As Long
'    Name                As String
'    NewName             As String
'
'    FilePath            As String
'    FileName            As String
'    FileExt             As String
'
'    ResFilePath         As String
'    ResFileName         As String
'
'    DataType            As String
'
'    Encrypt             As Boolean
'
'    Code                As String
'    Blocks()        As VBCodeBlock
'
'    ProtectedAreaCount  As Long
'    ProtectedAreas()    As VBCodeBlock
'
'    ControlCount        As Long
'    Controls()          As VBControl
'
'    DeclinationCount    As Long
'    Declinations()      As VBAlias
'
'    VariableCount       As Long
'    Variables()         As VBAlias
'
'    SubCount            As Long
'    Subs()              As VBSub
'
'End Type

'Public Type VBProject
'    Name                As String
'    FilePath            As String
'    FileName            As String
'    Code                As String
'    ModuleCount         As Long
'    Modules()           As VBModule
'End Type

Function IsDelimiter(Char As String, Optional IgnoredDelimiters As String) As Boolean
    
    Dim i As Long
    Dim Chars As String

    For i = Asc("a") To Asc("z")
        Chars = Chars & Chr(i)
    Next i
    For i = 0 To 9
        Chars = Chars & CStr(i)
    Next i
    
    Chars = Chars & "_" '& Chr(34)
    
    IsDelimiter = (InStr(1, Chars, Char, vbTextCompare) = 0)
    If IsDelimiter Then
        IsDelimiter = (InStr(1, IgnoredDelimiters, Char, vbTextCompare) = 0)
    End If

End Function


Function InText(Start As Long, Text As String, Word As String, Optional IgnoredDelimitersL As String, Optional IgnoredDelimitersR As String, Optional Reverse As Boolean, Optional WholeWordsOnly As Boolean = True) As Long
    
    Const DisallowedNeighboursLString As String = "VB."
    Const DisallowedNeighboursRString As String = ""
    
    Dim lStart                  As Long
    Dim x                       As Long
    Dim i                       As Long
    Dim a                       As Long
    Dim z                       As Long
    Dim Char                    As String
    Dim d1                      As Boolean
    Dim d2                      As Boolean
    Dim DisallowedNeighboursL() As String
    Dim DisallowedNeighboursR() As String
    Dim Lenght                  As Long
    Dim StartPos                As Long
    Dim EndPos                  As Long
    Dim Neighbour               As String
    
    DisallowedNeighboursL = Split(DisallowedNeighboursLString, "|")
    DisallowedNeighboursR = Split(DisallowedNeighboursRString, "|")
    
    lStart = Start
    
    If Word = "" Then
        Debug.Print "InText() --> Word=''"
        Exit Function
    End If
    
    Do
    
        If Reverse Then
            x = InStrRev(Text, Word, lStart, vbTextCompare)
        Else
            x = InStr(lStart, Text, Word, vbTextCompare)
        End If
        
        If x = 0 Then Exit Function
        
        If WholeWordsOnly Then
            
            a = x - 1
            z = x + Len(Word)
            
            If a > 0 Then
                Char = Mid(Text, a, 1)
                'If Char = Chr(34) Then
                '    Debug.Print GetLine(Text, a)
                '    Debug.Assert False
                'End If
                d1 = IsDelimiter(Char, IgnoredDelimitersL)
                If d1 Then
                    For i = 0 To UBound(DisallowedNeighboursL)
                        Lenght = Len(DisallowedNeighboursL(i))
                        StartPos = x - Lenght
                        If StartPos > 0 Then
                            Neighbour = Mid(Text, StartPos, Lenght)
                            If Neighbour = DisallowedNeighboursL(i) Then
                                StartPos = StartPos - 1
                                If StartPos > 0 Then
                                    Char = Mid(Text, StartPos, 1)
                                    If IsDelimiter(Char) Then
                                        d1 = False
                                    End If
                                Else
                                    d1 = False
                                End If
                                Exit For
                            End If
                        End If
                    Next i
                End If
            Else
                d1 = True
            End If
            
            If z <= Len(Text) Then
                Char = Mid(Text, z, 1)
                'If Char = Chr(34) Then
                '    Debug.Print GetLine(Text, z)
                '    Debug.Assert False
                'End If
                d2 = IsDelimiter(Char, IgnoredDelimitersR)
                If d2 Then
                    For i = 0 To UBound(DisallowedNeighboursR)
                        Lenght = Len(DisallowedNeighboursR(i))
                        StartPos = x + Len(Word)
                        EndPos = StartPos + Lenght
                        If EndPos < Len(Text) Then
                            If Mid(Text, x - Lenght, Lenght) = DisallowedNeighboursR(i) Then
                                d1 = False
                                Exit For
                            End If
                        End If
                    Next i
                End If
            Else
                d2 = True
            End If
        
        End If
        
        If (d1 And d2) Or (Not WholeWordsOnly) Then
            InText = x
            Exit Function
        End If
        
        If Reverse Then
            lStart = x - Len(Word)
        Else
            lStart = x + Len(Word)
        End If
    
    Loop While lStart < Len(Text)
    
    
    
End Function

Function AnyInText(Text As String, Words() As String, Optional WholeWordsOnly As Boolean = True, Optional Start As Long = 1) As Long
    
    Dim i As Long
    Dim c As Long
    
    For i = LBound(Words) To UBound(Words)
        c = c + 1
        If InText(Start, Text, Words(i), , , , WholeWordsOnly) <> 0 Then
            AnyInText = c
            Exit Function
        End If
        
    Next i
    
End Function

Function LineBeginIsIn(TxtLine As String, Words() As String) As Long
    
    Dim i As Long
    
    For i = LBound(Words) To UBound(Words)
       
        If Left(LTrim(TxtLine), Len(Words(i))) = Words(i) Then
            
            LineBeginIsIn = i
            Exit Function
            
        End If
        
    Next i
    
    LineBeginIsIn = -1
    
End Function

Function GetWord(Start As Long, Text As String) As String
    
    Dim i As Long
    Dim Char As String
    Dim r As Boolean
    Dim FirstChar As Long
    Dim LastChar As Long
    
    For i = Start To Len(Text) + 1
        
        Char = Mid(Text, i, 1)
        
        If i > Len(Text) Then
            r = True
        Else
            r = IsDelimiter(Char, ".")
        End If
        
        If FirstChar = 0 Then
            If r = False Then
                FirstChar = i
            End If
        Else
            If r = True Then
                LastChar = i
                Exit For
            End If
        End If
        
    Next i
    
    GetWord = Mid(Text, FirstChar, LastChar - FirstChar)
    
End Function


Function JoinLines(Lines() As String, FirstLine As Long, LastLine As Long) As String
    Dim l As Long
    Dim s As String
    For l = FirstLine To LastLine
        s = s & Lines(l)
        If l <> LastLine Then
            s = s & vbNewLine
        End If
    Next l
    JoinLines = s
End Function


Function GetFileExt(FileName As String) As String
    Dim x As Long
    x = InStrRev(FileName, ".")
    If x <> 0 Then
        GetFileExt = Mid(FileName, x, Len(FileName) - x + 1)
    End If
End Function



'Sub ptest()
'    Dim a() As String
'    Dim c As Long
'    Dim i As Long
'
'    c = GetSubParameters("Function ReplaceWords(Text As String, _" & vbNewLine & "OldWord As String, _" & vbNewLine & "      NewWord As String _" & vbNewLine & ", Optional IgnoredDelimitersL As String, _  " & vbNewLine & "  Optional IgnoredDelimitersR As String)  _" & vbNewLine & "     As String", a)
'
'    For i = 1 To c
'        Debug.Print a(i)
'    Next i
'
'End Sub


Function RandomAlias(Lenght As Long) As String
    
    Dim RndName As String
    Dim x As Long

    x = RandomNumber(2, Lenght)
    RndName = RandomName(Lenght - 1)
    RndName = Left(RndName, x - 1) & RandomNumber(0, 9) & Mid(RndName, x)
    RandomAlias = RndName
    
End Function

Function GetLine(Text As String, x As Long) As String
    Dim a As Long
    Dim z As Long
    
    a = InStrRev(Text, vbNewLine, x)
    If a = 0 Then
        a = 1
    Else
        a = a + 2
    End If
    
    z = InStr(x, Text, vbNewLine)
    If z = 0 Then
        z = Len(Text) + 1
    End If
    
    GetLine = Mid(Text, a, z - a)
    
End Function

Function Quote(Text As String) As String
    Quote = Chr(34) & Text & Chr(34)
End Function

Function ReplaceQuotations(Text As String, OldWord As String, NewWord As String, Optional ReplaceQuotes As Boolean = True) As String
    Dim s As String
    s = Text
    If ReplaceQuotes Then
        s = ReplaceWords(s, Quote(OldWord), Quote(NewWord))
    End If
    s = ReplaceWords(s, OldWord, NewWord)
    ReplaceQuotations = s
End Function


Function RemoveComments(Code As String) As String

    Dim Lines() As String
    Dim i As Long
    Dim x As Long
    
    Lines() = Split(Code, vbNewLine)
    
    For i = 0 To UBound(Lines)
        
        x = InStrRev(Lines(i), "'")
        
        If x <> 0 Then
            If Not IsQuoted(Lines(i), x, 1) Then
                Lines(i) = RTrim(Left(Lines(i), x - 1))
            Else
                'Debug.Print Lines(i)
            End If
        End If
        
'        If Left(Lines(i), Len("Attribute")) = "Attribute" Then
'            If InStr(1, Lines(i), "VB_Description", vbTextCompare) <> 0 Then
'                Lines(i) = ""
'            End If
'        End If
        
    Next i
    
    RemoveComments = Join(Lines, vbNewLine)

End Function

Function RemoveEmptyLines(Code As String) As String

    Dim Lines() As String
    Dim i As Long
    Dim s As String
    
    Lines() = Split(Code, vbNewLine)
    
    For i = 0 To UBound(Lines)
        
        If Trim(Lines(i)) <> "" Then
            s = s & Lines(i) & vbNewLine
        End If
        
    Next i
    
    RemoveEmptyLines = s

End Function

Function RemoveSpaces(Code As String) As String

    Dim Lines() As String
    Dim i As Long
    Dim s As String
    
    Lines() = Split(Code, vbNewLine)
    
    For i = 0 To UBound(Lines)
        
        s = s & LTrim(Lines(i)) & vbNewLine
        
    Next i
    
    RemoveSpaces = s

End Function

Function ReplaceWords(Text As String, OldWord As String, NewWord As String, Optional IgnoredDelimitersL As String = "", Optional IgnoredDelimitersR As String = "", Optional ReplaceInQuotes As Boolean = False) As String
    
    Dim i As Long
    Dim Char As String
    Dim r As Boolean
    Dim FirstChar As Long
    Dim LastChar As Long
    Dim s As String
    
    Dim LSide As String
    Dim RSide As String
    
    Dim StartPos As Long
    Dim EndPos As Long
    
    StartPos = 1
    s = Text
    
'    If InStr(1, OldWord, "Obfuscate", vbTextCompare) <> 0 Then
'
'        If InStr(1, NewWord, "aObfuscateZ", vbTextCompare) <> 0 Then
'
'            If InStr(1, Text, OldWord, vbTextCompare) <> 0 Then
'
'                Debug.Print Text
'                Debug.Assert False
'
'            End If
'
'        End If
'
'    End If
    
    Do
        
        StartPos = InText(StartPos, s, OldWord, IgnoredDelimitersL, IgnoredDelimitersR)
        
        If StartPos <> 0 Then
            
            'If InStrArray(OldWord, VBProtectedProperties) <> -1 Then
            '    ReplaceWords = Text
            '    Exit Function
            'End If
            
            If (Not IsQuoted(s, StartPos, Len(OldWord))) Or ReplaceInQuotes Then
            
                LSide = Left(s, StartPos - 1)
                RSide = Right(s, Len(s) - StartPos - Len(OldWord) + 1)
                
                StartPos = StartPos + Len(NewWord)
                
                s = LSide & NewWord & RSide
            
            Else
                
                StartPos = StartPos + Len(OldWord)
                'Debug.Assert False
                
            End If
            
        End If
    
    Loop While StartPos <> 0
    
    ReplaceWords = s
    
End Function


Function ReplaceWordsEx(Module As VBModule, Text As String, OldWord As String, NewWord As String, Optional IgnoredDelimitersL As String, Optional IgnoredDelimitersR As String) As String
    
    Dim i As Long
    Dim Char As String
    Dim r As Boolean
    Dim FirstChar As Long
    Dim LastChar As Long
    Dim s As String
    
    Dim LSide As String
    Dim RSide As String
    
    Dim StartPos As Long
    Dim EndPos As Long
    
    'If InStr(1, OldWord, "Form") <> 0 Then
    '    Debug.Assert False
    'End If
    
    StartPos = 1
    s = Text
    
    Do
        
        StartPos = InText(StartPos, s, OldWord, IgnoredDelimitersL, IgnoredDelimitersR)
        
        If StartPos <> 0 Then
            
            For i = 1 To Module.BlockCount
                With Module.Blocks(i)
                    
                    If .Protected Then
                
                        
                                    
                    End If
                                    
                End With
            Next i
            
            LSide = Left(s, StartPos - 1)
            RSide = Right(s, Len(s) - StartPos - Len(OldWord) + 1)
            
            StartPos = StartPos + Len(NewWord)
            
            s = LSide & NewWord & RSide
        
        End If
    
    Loop While StartPos <> 0
    
    ReplaceWordsEx = s
    
End Function





Function IsQuoted(CodeLine As String, Start As Long, Length As Long) As Boolean
    
    Dim i As Long
    Dim c As Long
    Dim Quotes() As VBQuote
    Dim s As String
    
    c = SplitQuotes(CodeLine, Quotes)
    
    For i = 1 To c
        
        With Quotes(i)
            
            If (.StartPos <= Start) And (.EndPos >= Start + Length) Then
                IsQuoted = True
            End If
            
        End With
        
    Next i
    
    
End Function

Function CountChar(Char As String, Text As String, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Long

    Dim x As Long
    Dim c As Long
    
    Do
        x = InStr(x + 1, Text, Char, CompareMode)
        If x <> 0 Then
            c = c + 1
        End If
    Loop While x <> 0

    CountChar = c

End Function


'Function qtest()
'
'    Dim i As Long
'    Dim c As Long
'    Dim s As String
'    Dim q() As VBQuote
'
'    s = ReadFile("D:\vb.txt")
'
'    Debug.Print s
'    c = SplitQuotes(s, q)
'
'    For i = 1 To c
'        Debug.Print "|" & q(i).Quote & "|"
'    Next i
'
'
'End Function


Function SplitQuotes(Text As String, Quotes() As VBQuote) As Long

    'Dim i As Long
    Dim x As Long
    Dim r As Long
    Dim c As Long
    Dim Counter As Long
    
    Dim a As Long
    Dim z As Long
    
    Dim CharA As String
    Dim CharZ As String
    
    Dim Char As String
    Dim StartPos As Long
    Dim EndPos As Long
    
    Dim Found As Boolean
    
    ReDim Quotes(0)
    
    Do
        x = x + 1
        
        Char = Mid(Text, x, 1)
        
        If Char = Chr(34) Then
            
            If StartPos = 0 Then
                
                StartPos = x + 1
                Counter = 1
                
            Else
                
                Counter = Counter + 1
                
                a = x - 1
                z = x + 1
                CharA = Mid(Text, a, 1)
                CharZ = Mid(Text, z, 1)
                
                If (Counter Mod 2) = 0 Then
                
                    If CharZ <> Chr(34) Then
                        Found = True
                    End If
                
                End If
                
                
                If Found Then
                    EndPos = x
                    c = c + 1
                    ReDim Preserve Quotes(c)
                    With Quotes(c)
                        .StartPos = StartPos
                        .EndPos = EndPos
                        .Length = (EndPos - StartPos)
                        .Quote = Mid(Text, StartPos, EndPos - StartPos)
                    End With
                    'Quotes(c) = Mid(Text, StartPos, EndPos - StartPos)
                    'Debug.Print "'" & Quotes(c) & "'"
                    StartPos = 0
                    Found = False
                End If
                
                
            End If
            
        End If
        
    Loop While x <= Len(Text)
    
    SplitQuotes = c

End Function

Function NotInStr(Start As Long, Text As String, SkippedStr As String)
    Dim i As Long
    For i = Start To (Len(Text) - Len(SkippedStr) + 1)
        If Mid(Text, i, Len(SkippedStr)) <> SkippedStr Then
            NotInStr = i
            Exit Function
        End If
    Next i
End Function


Function AdaptCode(Code As String) As String

    Dim StrCode As String
    
    
    StrCode = Code
    
    
    With My.Config
        
         If .RemoveCodeLayout Then
             
            If .RemoveComments Then
                StrCode = RemoveComments(StrCode)
            End If
            
            If .RemoveUnderscores Then
                StrCode = RemoveUnderscores(StrCode)
            End If
            
            If .RemoveColons Then
                StrCode = RemoveColons(StrCode)
            End If
             
            If .RemoveEmptyLines Then
                StrCode = RemoveEmptyLines(StrCode)
            End If
            
            If .RemoveSpaces Then
                StrCode = RemoveSpaces(StrCode)
            End If
             

        
         Else
             
           
            StrCode = RemoveComments(StrCode)
                
            StrCode = RemoveUnderscores(StrCode)
           
             
         End If
        
    End With
    
    AdaptCode = StrCode
       
End Function


Function RemoveColons(Code As String) As String

    Dim Lines() As String
    Dim i As Long
    Dim x As Long
    Dim z As Long
    Dim LineStr As String
    Dim SpaceCount As Long
    Dim NewLine As String
    Dim LeftPart As String
    Dim RightPart As String
    
    If InStr(1, Code, ":") = 0 Then
        RemoveColons = Code
        Exit Function
    End If
    
    Lines() = Split(Code, vbNewLine)

    For i = 0 To UBound(Lines)
        
        x = InStr(1, Lines(i), ":")
        If (x <> 0) Then
            
            LineStr = RTrim(Lines(i))
            
            If x <> Len(LineStr) Then
                
                SpaceCount = Len(Lines(i)) - Len(LTrim(Lines(i)))
                
                Do
        
                    z = InStr(z + 1, Lines(i), ":")
                    If z <> 0 Then
                        
                        If Not IsQuoted(Lines(i), z, 1) Then
                
                            LeftPart = Left(Lines(i), z - 1)
                            RightPart = Mid(Lines(i), z + 1)
                            
                            NewLine = LeftPart & vbNewLine & String(SpaceCount, " ") & LTrim(RightPart)
                            Lines(i) = NewLine
                
                        End If
            
                    End If
                
                Loop While z <> 0
            
            End If
                
        End If
            
    Next i

    RemoveColons = Join(Lines, vbNewLine)


End Function


Function GetDataType(VarStr As String) As String
    
    Const DataNameList As String = "Byte|Boolean|Integer|Long|Single|Double|Currency|Date|String|Object|Variant"
    Const DataSymbolList As String = "||%|&|!|#|@||$||"
    
    Dim DataNames() As String
    Dim DataSymbols() As String
    Dim i As Long
    Dim x As Long
    Dim s As String
    
    DataNames = Split(DataNameList, "|")
    DataSymbols = Split(DataSymbolList, "|")
    
    For i = 0 To UBound(DataNames)
        
        If DataSymbols(i) <> "" Then
            If InStr(1, VarStr, DataSymbols(i), vbTextCompare) <> 0 Then
                GetDataType = DataNames(i)
                Exit Function
            End If
        End If
        
        s = "As " & DataNames(i)
        If InText(1, VarStr, s) <> 0 Then
            GetDataType = DataNames(i)
            Exit Function
        End If
       
    
    Next i
    
    x = InText(1, VarStr, "As")
    If x = 0 Then
        GetDataType = "Variant"
    Else
        GetDataType = GetWord(x + Len("As") + 1, VarStr)
    End If
    
    'Debug.Assert (InStrRev(VarStr, "MSComctlLib.Node") = 0)
    
End Function



Function SelectScope(DeclarationStatement As String) As VBScopes
    
    Select Case DeclarationStatement
        Case "Private", "Dim"
            SelectScope = VBLocal + VBPrivate
        Case "Public", "Global"
            SelectScope = VBLocal + VBPrivate + VBPublic
    End Select
    
End Function


Function RemoveUnderscores(CodeStr As String) As String
    
    Dim Lines() As String
    Dim Line As String
    Dim LastLine As Long
    Dim s As String
    Dim tmp As String
    Dim i As Long
    Dim l As Long
    
    'Debug.Print SubCode
    
     Lines = Split(CodeStr, vbNewLine)
     
    If InStr(1, CodeStr, "_") = 0 Then
        RemoveUnderscores = CodeStr
        Exit Function
    End If
    
    Do
    
        If Right(RTrim(Lines(l)), 1) = "_" Then
        
            For i = l To UBound(Lines)
                Line = RTrim(Lines(i))
                If Right(Line, 1) <> "_" Then
                    LastLine = i
                    Exit For
                End If
            Next i
        
            For i = l To LastLine
                Line = RTrim(Lines(i))
                
                'If i <> LastLine Then
                    
                    If i <> LastLine Then
                        Line = Left(Line, Len(Line) - 1)
                    End If
                    Line = RTrim(LTrim(Line))
                    If i = l Then
                        s = s & Line
                    Else
                       ' If Left(LTrim(Lines(i + 1)), 1) = "," Then
                        '    s = s & Line
                        'Else
                        If Left(Line, 1) <> "," And Left(Line, 1) <> ")" Then
                            s = s & " "
                        End If
                        s = s & Line
                        'End If
                    End If
                    If i = LastLine Then
                        s = s & vbNewLine
                    End If
                    
                    'Debug.Print "|" & s & "|"
                'Else
                '
                '    s = s & RTrim(LTrim(Line)) & vbNewLine
                '    's = s & Line
                '    Exit For
                'End If
                
            Next i
            
            l = LastLine
            
        Else
        
            s = s & Lines(l)
            If l <> UBound(Lines) Then
                s = s & vbNewLine
            End If
        
        End If
    
        l = l + 1
    Loop While l <= UBound(Lines)
    
    RemoveUnderscores = s
    
End Function

Function Hug(Text As String, Optional HugSign1 As String = "[", Optional HugSign2 As String = "]") As String
    
    Hug = HugSign1 & Text & HugSign2
    
End Function


Function adapttest()

    Dim f1 As String
    Dim f2 As String
    Dim s As String

    f1 = "R:\GradientClass.cls"
    f2 = "R:\GradientClass2.cls"
    s = RemoveColons(RemoveComments(ReadFile(f1)))

    If FileExist(f2) Then Kill f2
    WriteFile f2, s

End Function



Function ReplaceInLines(Lines() As String, FirstLine As Long, LastLine As Long, OldWord As String, NewWord As String, Optional ReplaceInQuotes As Boolean = False)
            
    Dim i As Long

    
    For i = FirstLine To LastLine
        
'        If ((OldWord = "MimeEncode") And (InStr(1, Lines(i), "MimeEncode") <> 0)) Then
'            Debug.Assert False
'        End If
    
'        If (InStr(1, Lines(i), Chr(34) & "Obfuscate" & Chr(34), vbTextCompare) <> 0) Then
'
'
'
'        End If

        'Debug.Assert (i <> 42)
        
        'Debug.Print Lines(i)
        Lines(i) = ReplaceWords(Lines(i), OldWord, NewWord, , , ReplaceInQuotes)
        'Debug.Print Lines(i)
    
        
       
    
    Next i
        
End Function



'Public Function FastReplace(Text As String, Find As String, Replace As String, Optional Start As Long = 1, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As String
'
'    Static cReplace As cReplaceFast
'
'    If cReplace Is Nothing Then
'        Set cReplace = New cReplaceFast
'    End If
'
'    With cReplace
'
'        FastReplace = .ReplaceFast(Text, Find, Replace, Start, , CompareMode)
'
'    End With
'
'
'End Function


'Public Function RegEx(Text As String, RegExp As String, Results() As String) As Long
'
'    Dim myRegExp As RegExp
'    Dim myMatches As MatchCollection
'    Dim myMatch As Match
'    Dim MyStrings() As String
'    Dim i As Long
'
'    Set myRegExp = New RegExp
'
'    myRegExp.IgnoreCase = True
'    myRegExp.Global = True
'    myRegExp.Pattern = RegExp
'
'    Set myMatches = myRegExp.Execute(Text)
'
'    ReDim MyStrings(myMatches.Count)
'
'    For Each myMatch In myMatches
'        i = i + 1
'        MyStrings(i) = myMatches.Item(i - 1)
'    Next
'
'    Results = MyStrings
'    RegEx = myMatches.Count
'
'End Function
'
'Sub testRegEx()
'
'    Dim Str As String
'    Dim RegExp As String
'    Dim s() As String
'    Dim c As Long
'    Dim i As Long
'
'    Str = "abc"
'    RegExp = "\wabc"
'
'    c = RegEx(Str, RegExp, s)
'
'    For i = 1 To c
'
'        Debug.Print s(i)
'
'    Next i
'
'End Sub
