Attribute VB_Name = "mDebug"

Function LogMsg(Msg As String)
    
    Static LogCount As Long
    
    Dim FilePath As String
    Dim FileData As String
    
    FilePath = App.Path
    If Right(App.Path, 1) <> "\" Then
        FilePath = FilePath & "\"
    End If
    FilePath = FilePath & App.EXEName & ".log"
    
    If LogFileExist(FilePath) Then
        If LogCount = 0 Or LogCount > 1000 Then
            Kill FilePath
        Else
            FileData = LogReadFile(FilePath)
        End If
    End If
    
    FileData = FileData & Msg & vbNewLine
    LogWriteFile FilePath, FileData
    LogCount = LogCount + 1
    
End Function

Private Function LogFileExist(FilePath As String) As Boolean
On Error GoTo Error
    FileLen FilePath
    LogFileExist = True
Exit Function
Error:

End Function

Private Function LogReadFile(FilePath As String, Optional Start As Long, Optional Lenght As Long, Optional Reverse As Boolean = False) As String

    Dim FileNum As Long
    Dim Buffer As String
    Dim lLen As Long
    Dim lStart As Long
    Dim lLOF As Long

    lLen = Lenght
    lStart = Start
    
    FileNum = FreeFile
    Open FilePath For Binary As #FileNum
        lLOF = LOF(FileNum)
        If lStart > lLOF Then
            Err.Description = "Startpoint > Filelenght"
        End If
        If lStart = 0 Then lStart = 1
        If lLen = 0 Or (lStart + lLen - 1) > lLOF Then lLen = (lLOF - lStart + 1)
        If Reverse Then lStart = lLOF - lStart - lLen + 2
        Buffer = String(lLen, " ")
        Get #FileNum, lStart, Buffer
    Close FileNum
    LogReadFile = Buffer
    

End Function

Private Function LogWriteFile(sFilePath$, sData$, Optional ByVal lStart&, Optional Insert As Boolean, Optional NoError As Boolean)

    Dim FileNum As Long
    Dim sBuffer As String
    Dim FileExists As Boolean
    Dim FileSize As Long
    Dim Buffer As String
    
    
    If NoError Then On Error Resume Next
    
    If FileExists Then
        FileSize = FileLen(sFilePath)
    End If
    If lStart < 1 Then
        If FileExists Then lStart = FileSize + 1
        If lStart < 1 Then lStart = 1
    End If
    FileNum = FreeFile
    If FileExists And Insert Then
        Buffer = sData & LogReadFile(sFilePath, lStart)
        Open sFilePath For Binary As FileNum
        Put FileNum, lStart, Buffer
        Close FileNum
    Else
        Open sFilePath For Binary As FileNum
        Put FileNum, lStart, sData
        Close FileNum
    End If
        

End Function

