Attribute VB_Name = "MainModule"
Option Explicit

Public Enum VBModuleTypes
    VBbas
    VBcls
    VBfrm
    VBctl
    VBres
    VBpag
    VBdsr
End Enum

Public Enum VBAliasTypes
    
    VBFile = 2 ^ 1
    VBProject = 2 ^ 2
    VBModule = 2 ^ 3
    
    VBControl = 2 ^ 4
    
    VBSub = 2 ^ 5
    VBFunction = 2 ^ 6
    VBProperty = 2 ^ 7
    
    VBParam = 2 ^ 8
    
    VBVariable = 2 ^ 9
    
End Enum

Public Enum VBAliasScopes
    VBPublic = 2 ^ 1
    VBPrivate = 2 ^ 2
    VBLocal = 2 ^ 3
End Enum
 
Public Type VBAlias
    ID As Long
    Name As String
    NewName As String
    AliasType As VBAliasTypes
    Scope As VBAliasScopes
    ModuleID As Long
    SubID As Long
    VariableID As Long
End Type

Public Type VBSub
    ID As Long
    Start As Long
    End As Long
    Code As String
    Name As String
    NewName As String
    ModuleID As Long
    VariableCount As Long
    Variables() As VBAlias
End Type

Public Type VBModule
    ID As Long
    Name As String
    NewName As String
    
    FilePath As String
    FileName As String
    FileExt As String
    
    TypeName As String
    
    Encrypt As Boolean
    Code As String
    
    SubCount As Long
    Subs() As VBSub
    
    VariableCount As Long
    Variables() As VBAlias
    
End Type

Function IsDelimiter(Char As String) As Boolean
    
    Dim i As Long
    Dim Chars As String

    For i = Asc("a") To Asc("z")
        Chars = Chars & Chr(i)
    Next i
    For i = 0 To 9
        Chars = Chars & CStr(i)
    Next i
    
    IsDelimiter = (InStr(1, Chars, Char, vbTextCompare) = 0)

End Function


Function InText(Start As Long, Text As String, Word As String) As Long
    
    Dim lStart As Long
    Dim x As Long
    Dim i As Long
    Dim a As Long
    Dim z As Long
    Dim Char As String
    Dim D1 As Boolean
    Dim D2 As Boolean
    
    lStart = Start
    
    Do
    
        x = InStr(lStart, Text, Word, vbTextCompare)
        If x = 0 Then Exit Function
    
        a = x - 1
        z = x + Len(Word)
    
        If a > 0 Then
            Char = Mid(Text, a, 1)
            D1 = IsDelimiter(Char)
        Else
            D1 = True
        End If
        
        If z < Len(Text) Then
            Char = Mid(Text, z, 1)
            D2 = IsDelimiter(Char)
        Else
            D2 = True
        End If
        
        If D1 And D2 Then
            InText = x
            Exit Function
        End If
    
        lStart = x + Len(Word)
    
    Loop While lStart < Len(Text)
    
    
    
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
            r = IsDelimiter(Char)
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

Function ReplaceWord(Start As Long, Text As String, Replacement As String) As String
    
    Dim i As Long
    Dim Char As String
    Dim r As Boolean
    Dim FirstChar As Long
    Dim LastChar As Long
    Dim s As String
    
    For i = Start To Len(Text) + 1
        
        Char = Mid(Text, i, 1)
        
        If i > Len(Text) Then
            r = True
        Else
            r = IsDelimiter(Char)
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
    
    s = Left(Text, FirstChar - 1) & Replacement & Right(Text, Len(Text) - LastChar + 1)
    
    ReplaceWord = s
    
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

