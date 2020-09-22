Attribute VB_Name = "mDebug"
Option Explicit

Public DebugMode    As Boolean

Public LogFilePath  As String
Public LogFileData  As String

'Customize these the following constants
Public Const DebugAppPath As String = "" '"D:\FirstCode\BrainTraineR\Release\"
Public Const DebugCommandLineOption = "-debug"

Sub Assert(Optional Message$, Optional Source$, Optional Number&, Optional Description$, Optional Info$, Optional Break As Boolean = True, Optional Transfer As Boolean)

   
        
        If Message <> Empty Then
            Log Message
            Exit Sub
        End If
        
        'Log "Error occurred at: " & (Date + Time)
        'Log "Source: " & Source
        'Log "Number: " & Number
        'Log "Description: " & Description
        'If Not Info = Empty Then Log "Info: " & Info
        'Log ""
        
        DoEvents
        
        If Message = "" Then Debug.Print Description: Debug.Assert (Break = False)
        
        If Not DebugMode And Not InIDE Then
            Msg "Error. " & NewLine(2) & "Ursache: " & Description & NewLine(1) & "Nummer: " & CStr(Number) & NewLine(1) & "Quelle: " & Source, vbCritical
        End If
        
        If Transfer Then
            Err.Raise Number, Source, Description
        End If
    

    
End Sub

'Sub AssertErr(Optional Source$, Optional Number&, Optional Description$, Optional Info$, Optional Transfer As Boolean)
'
'    With DebugForm
'
'        .LogError Source, Description, Number, Info
'
'        If Transfer Then
'            Err.Raise Number, Source, Description
'        End If
'
'    End With
'
'End Sub

Public Sub Log(Text As String)
On Error Resume Next

    If Len(LogFileData) > 1048576 Then LogFileData = ""

    LogFileData = LogFileData & Text & vbNewLine

    If Not InIDE Then WriteFile LogFilePath, LogFileData

End Sub

Public Function InIDE() As Boolean
On Error GoTo Error
    Debug.Assert 1 / 0
Exit Function
Error:
    InIDE = True
End Function


Sub XCopyExt(SrcDir As String, DstDir As String, Ext As String)
On Error GoTo Error

    Dim i As Long
    Dim x As Long
    Dim Dirs As Long
    Dim d() As String
    Dim c As Long
    Dim f() As String
    Dim src As String
    Dim Dst As String
    
    If Not DirExist(DstDir) Then
        BuildPath DstDir
    End If
    
    Dirs = GetSubDirList(SrcDir, d())
    
    For x = 1 To Dirs
    
        c = GetFileList(d(x), f, True, True)
        
        For i = 1 To c
            src = f(i)
            If LCase(Right(src, 4)) = Ext Then
                Dst = FixPath(DstDir) & GetFileName(f(i))
                
              
                
                Debug.Print src, "-->", Dst
                FileCopy src, Dst
                
            End If
            
            
        Next i
    
    Next x

Exit Sub
Error:
    Resume
End Sub
