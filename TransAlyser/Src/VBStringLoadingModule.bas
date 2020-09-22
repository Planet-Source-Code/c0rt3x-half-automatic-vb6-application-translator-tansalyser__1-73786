Attribute VB_Name = "VBStringLoadingModule"
Option Explicit

Public Function VBGetTxt(VBStringID As Long) As String
    
    Static FileData As String
    Static Lines() As String
    
    Dim LineStr As String
    
    If FileData = "" Then
    
        FileData = VBLoadFile(App.ExeName & ".str")
        Lines = Split(FileData, vbNewLine)
    
    End If
    
    LineStr = Lines(VBStringID - 1)
    
    VBGetTxt = LineStr
    
End Function

Function VBLoadFile(FilePath As String, Optional Start As Long, Optional Lenght As Long, Optional Reverse As Boolean = False) As String
On Error GoTo Error

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
            GoTo Error
        End If
        If lStart = 0 Then lStart = 1
        If lLen = 0 Or (lStart + lLen - 1) > lLOF Then lLen = (lLOF - lStart + 1)
        If Reverse Then lStart = lLOF - lStart - lLen + 2
        Buffer = String(lLen, " ")
        Get #FileNum, lStart, Buffer
    Close FileNum
    VBLoadFile = Buffer
    
Exit Function
Error:
    MsgBox Err.Description
    Resume Next
End Function

Function CamelCase(Text As String) As String
    
    Dim i As Long
    Dim s As String
    Dim Char As String
    
    i = 1
    
    Do While i <= Len(Text)
        
        
        Char = Mid(Text, i, 1)
        
        If Char = " " Then
            
            s = s & UCase(Mid(Text, i + 1, 1))
            i = i + 2
        
        Else
            
            s = s & Char
            i = i + 1
            
        End If
        
       
        
    Loop
    
    CamelCase = s
    
End Function

