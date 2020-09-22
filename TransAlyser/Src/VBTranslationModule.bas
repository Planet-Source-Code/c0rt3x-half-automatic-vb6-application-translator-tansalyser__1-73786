Attribute VB_Name = "VBTranslationModule"
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const SORT_DEFAULT As Integer = &H0
Private Const LANG_NEUTRAL As Integer = &H0
Private Const SUBLANG_DEFAULT As Integer = &H1
Private Const SUBLANG_SYS_DEFAULT As Integer = &H2

Private Const LANG_SYSTEM_DEFAULT As Long = (SUBLANG_SYS_DEFAULT * 1024&) Or LANG_NEUTRAL
Private Const LANG_USER_DEFAULT As Long = (SUBLANG_DEFAULT * 1024&) Or LANG_NEUTRAL

Private Const LOCALE_SYSTEM_DEFAULT As Long = (SORT_DEFAULT * 65536) Or LANG_SYSTEM_DEFAULT
Private Const LOCALE_USER_DEFAULT As Long = (SORT_DEFAULT * 65536) Or LANG_USER_DEFAULT

Private Const LOCALE_NOUSEROVERRIDE = &H80000000 '// do not use user overrides
Private Const LOCALE_USE_CP_ACP = &H40000000 '// use the system ACP
Private Const LOCALE_RETURN_NUMBER = &H20000000 '// return number instead of string

Private Const LOCALE_ILANGUAGE = &H1& '// language id
Private Const LOCALE_SLANGUAGE = &H2& '// localized name of language


Private Const Base64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789%&"

Private Function Base64Encode(Text As String) As String
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
    
End Function

Private Function Base64Decode(Text As String) As String
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
    
End Function

Private Function MimeEncode(W As Integer, Optional ByVal Key As String) As String
On Error GoTo Error

    If IsMissing(Key) Then Key = Base64
    If W >= 0 Then MimeEncode = Mid$(Base64, W + 1, 1) Else MimeEncode = ""
    
Exit Function
Error:
    
End Function

Private Function MimeDecode(a As String, Optional ByVal Key) As Integer
On Error GoTo Error

    If IsMissing(Key) Then Key = Base64
    If Len(a) = 0 Then MimeDecode = -1: Exit Function
    MimeDecode = InStr(Base64, a) - 1

Exit Function
Error:
    
End Function


Private Function GetLocaleLanguage() As String
On Error Resume Next
    
    Dim StrRet As String
    Dim x As Long

    StrRet = String$(1024, 0)

    If GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLANGUAGE, StrRet, Len(StrRet)) Then
        StrRet = Left$(StrRet, InStr(StrRet, vbNullChar) - 1)
        x = InStr(1, StrRet, " ", vbTextCompare)
        If x <> 0 Then
            StrRet = Left(StrRet, x - 1)
        End If
    End If
    
    GetLocaleLanguage = StrRet
    
End Function

Private Function FixPath(Path As String, Optional Terminate As Boolean = True) As String
On Error GoTo Error

    If Right(Path, 1) = "\" Then
        FixPath = Left(Path, Len(Path) - 1)
    Else
        FixPath = Path
    End If
    If Terminate Then FixPath = FixPath & "\"

Exit Function
Error:
    Resume Next
End Function

Private Function DirExists(BaseDirPath As String) As Boolean
On Error GoTo Error
    If InStr(1, BaseDirPath, "\\") <> 0 Then GoTo Error
    If (GetAttr(BaseDirPath) And vbDirectory) = vbDirectory Then
        DirExists = True
    End If
Exit Function
Error:
     DirExists = False
End Function

Private Function FileExists(FilePath As String) As Boolean
On Error GoTo Error
    If DirExists(FilePath) Then GoTo Error
    FileLen FilePath
    FileExists = True
Exit Function
Error:

End Function

Private Function LoadFile(FilePath As String) As String
On Error GoTo Error

    Dim FileNum As Long
    Dim Buffer As String

    FileNum = FreeFile
    
    Open FilePath For Binary As #FileNum
        
        Buffer = String(LOF(FileNum), " ")
        
        Get #FileNum, 1, Buffer
        
    Close FileNum
    
    LoadFile = Buffer
    
Exit Function
Error:
    LoadFile = ""
End Function

Function SaveFile(FilePath As String, FileData As String)
On Error GoTo Error

    Dim FileNum As Long
    Dim Buffer As String
    Dim FileSize As Long
    
    FileNum = FreeFile
      
    If FileExists(FilePath) Then
        Kill FilePath
    End If
  
    Open FilePath For Binary As FileNum
        
        Put FileNum, 1, FileData
        
    Close FileNum
        
Exit Function
Error:
    Resume Next
End Function

Private Function GetIniValue(sIni As String, Section As Variant, Key As Variant) As String
On Error GoTo Error

Dim sSection$, sKey$
Dim lSectionStart&, lSectionEnd&, lKeyStart&, lValueEnd&

    sSection = "[" & CStr(Section) & "]"
    sKey = CStr(Key) & "="
    
    lSectionStart = InStr(1, sIni, sSection)
    
    If lSectionStart < 1 Then GoTo Error
    
    lSectionStart = lSectionStart + Len(sSection)
    
    lKeyStart = InStr(lSectionStart, sIni, vbNewLine & sKey) + 2
    lSectionEnd = InStr(lSectionStart, sIni, "[")
    
    If lSectionEnd < 1 Then lSectionEnd = Len(sIni)
    If lKeyStart < 3 Or lKeyStart > lSectionEnd Then GoTo Error
    
    lValueEnd = InStr(lKeyStart, sIni & vbNewLine, vbNewLine)
    
    If lValueEnd < 1 Then GoTo Error
    
    GetIniValue = Mid(sIni, lKeyStart + Len(sKey), lValueEnd - (lKeyStart + Len(sKey)))

Exit Function
Error:
    GetIniValue = ""
    
End Function

Private Function SetIniValue(sIni As String, Section As Variant, Key As Variant, Value As Variant)
On Error GoTo Error

    Dim pINI$, sSection$, sKey$, sValue$, Line$(), sSectionChunk$, sLChunk$, sRChunk$
    Dim lSectionStart&, lSectionEnd&, lKeyStart&, lValueStart&, lValueEnd&, i&

    sSection = "[" & CStr(Section) & "]"
    sKey = vbNewLine & CStr(Key) & "="
    sValue = CStr(Value)
    pINI = sIni
    If InStr(1, pINI, sSection) < 1 Then
        If pINI <> "" Then
            Do While Right(pINI, 4) <> vbNewLine & vbNewLine
                pINI = pINI & vbNewLine
            Loop
        End If
        pINI = pINI & sSection
    End If
    lSectionStart = InStr(1, pINI, sSection) + Len(sSection)
    lSectionEnd = lSectionStart
    sSectionChunk = Mid(pINI, lSectionStart)
    Line = Split(sSectionChunk, vbNewLine)
    For i = LBound(Line()) To UBound(Line())
        If Left(Line(i), 1) = "[" And Right(Line(i), 1) = "]" Then Exit For
        If Len(Line(i)) > 0 Then
            lSectionEnd = lSectionEnd + Len(Line(i)) + 2
        End If
    Next i
    lKeyStart = InStr(lSectionStart, pINI, sKey)
    If lKeyStart < 1 Or lKeyStart > lSectionEnd Then
        sLChunk = Left(pINI, lSectionEnd - 1)
        sRChunk = Mid(pINI, lSectionEnd)
        If Right(sLChunk, 2) = vbNewLine Then sLChunk = Left(sLChunk, Len(sLChunk) - 2)
        If Left(sRChunk, 2) <> vbNewLine Then sRChunk = vbNewLine & sRChunk
        If Left(sRChunk, 3) = vbNewLine & "[" Then sRChunk = vbNewLine & sRChunk
        pINI = sLChunk & sKey & sRChunk
        lKeyStart = InStr(lSectionStart, pINI, sKey)
    End If
    lValueStart = lKeyStart + Len(sKey)
    lValueEnd = InStr(lValueStart, pINI, vbNewLine)
    If lValueEnd < 1 Then lValueEnd = lValueStart + 1 '
    sLChunk = Left(pINI, lValueStart - 1) '
    sRChunk = Mid(pINI, lValueEnd)
    pINI = sLChunk & sValue & sRChunk
    Do While Right(pINI, 2) = vbNewLine
        pINI = Left(pINI, Len(pINI) - 2)
    Loop
    pINI = pINI & vbNewLine
    sIni = pINI
    
Exit Function
Error:
    
End Function


Private Function GetLanguageFilePath(Optional LanguageName As String = "") As String
On Error Resume Next
    
    Const LanguageFileExtension As String = ".lng"
    
    Dim FilePath As String
    
    FilePath = FixPath(App.Path)
    
    If LanguageName = "" Then
        FilePath = FilePath & GetLocaleLanguage & LanguageFileExtension
    Else
        FilePath = FilePath & LanguageName & LanguageFileExtension
    End If
    
    GetLanguageFilePath = FilePath
    
End Function

Private Function GetLanguageConfigFilePath() As String
    
    GetLanguageConfigFilePath = FixPath(App.Path) & "Language.ini"
    
End Function


Public Function GetTxt(ID As String) As String
    
    Static ConfigFileData As String
    Static LanguageName As String
    Static TranslationFileData As String
    
    Dim fLanguage As VBTranslationForm
    Dim sIni As String
    
    Dim StrTxt As String
    
    If ConfigFileData = "" Then
    
        If FileExists(GetLanguageConfigFilePath) Then
            ConfigFileData = LoadFile(GetLanguageConfigFilePath)
            LanguageName = GetIniValue(ConfigFileData, "Translation", "Language")
        End If
        
    End If
    
    If LanguageName = "" Then
        
        Set fLanguage = New VBTranslationForm
        Load fLanguage
        If fLanguage.cbLanguage.Text <> "" Then
            fLanguage.Show 1
            LanguageName = fLanguage.cbLanguage.Text
            SetIniValue ConfigFileData, "Translation", "Language", LanguageName
            SaveFile GetLanguageConfigFilePath, ConfigFileData
        End If
        Set fLanguage = Nothing
        
    End If
    
    If TranslationFileData = "" Then
        If FileExists(GetLanguageFilePath(LanguageName)) Then
            TranslationFileData = LoadFile(GetLanguageFilePath(LanguageName))
        End If
    End If
    
    StrTxt = GetIniValue(TranslationFileData, "Translation", Base64Encode(ID))
    
    If StrTxt = "" Then
        GetTxt = ID
    Else
        GetTxt = Base64Decode(StrTxt)
    End If
    
End Function


