VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule As VBModule
Private pParentControl As VBControl
Private pLineID As Long
Private pLineStr As String
Private pName As String
Private pValue As String
Private pResFileOffset As Long

Private pCaptionID As Long

Private pChildCount As Long
Private pChilds() As VBProperty

Private pList() As String


Public Checked As Boolean


Private Type FrxStringHeader
    
    StringLen As Byte
    StringLen2 As Integer
    
End Type


Property Get ParentModule() As VBModule
    Set ParentModule = pParentModule
End Property

Property Set ParentModule(NewObject As VBModule)
    Set pParentModule = NewObject
End Property

Property Get ParentControl() As VBControl
    Set ParentControl = pParentControl
End Property

Property Set ParentControl(NewObject As VBControl)
    Set pParentControl = NewObject
End Property

Property Get PropertyLine() As String
    PropertyLine = pLineStr
End Property

Property Let PropertyLine(NewValue As String)
    pLineStr = NewValue
End Property


Property Get LineID() As Long
    ChildCount = pChildCount
End Property


Property Get Name() As String
    Name = pName
End Property


Property Let Name(NewValue As String)
    pName = NewValue
End Property


Property Get Value() As String
    Value = pValue
End Property


Property Let Value(NewValue As String)
    pValue = NewValue
End Property

Property Get ResFileOffset() As Long
    ResFileOffset = pResFileOffset
End Property

Property Get ChildCount() As Long
    ChildCount = pChildCount
End Property

Property Let ChildCount(NewValue As Long)
    ReDim Preserve pChilds(NewValue)
    pChildCount = NewValue
End Property


Property Get Childs(Index As Long) As VBProperty
    Set Childs = pChilds(Index)
End Property

Property Set Childs(Index As Long, NewObject As VBProperty)
    Set pChilds(Index) = NewObject
End Property

Property Get List(Index As Long) As String
    List = pList(Index)
End Property

Property Let List(Index As Long, NewVal As String)
    pList(Index) = NewVal
End Property

Property Get ListCount() As Long
    ListCount = UBound(pList)
End Property

Function LoadLine(LineID As Long)
        
    Dim Words() As String
    Dim s As String
    Dim x As Long
    
    pLineID = LineID
    pLineStr = ParentModule.Lines(LineID)
    
    If InText(1, pLineStr, "BeginProperty") = 0 Then
    
        Words = Split(pLineStr, "=")
    
        Words(0) = RTrim(LTrim(Words(0)))
        Words(1) = RTrim(LTrim(Words(1)))
    
        Name = Words(0)
    
        'If InStr(1, Words(1), Chr(34)) <> 0 Then
        '    pIsString = True
        '    Value = FilterString(Words(1), Chr(34))
        'Else
        '    pIsString = False
           
                Value = Words(1)
            
        'End If
        
        
        If IsCaption Then
            pCaptionID = Me.ParentModule.AddCaption(Me)
        End If
    
    Else
        
        s = RTrim(LTrim(pLineStr))
        Name = Mid(s, InStrRev(s, " ") + 1)
        
    End If
    
    
    x = InStr(1, pLineStr, Chr(34) & ":")
    
    If x <> 0 Then
        
        
        s = Mid(pLineStr, 2, Len(pLineStr) - 2)
        
        Words() = Split(pLineStr, Chr(34))
        
        Me.ParentModule.ResFileName = Words(1)
        Me.ParentModule.ResFilePath = My.Project.BaseDirPath & Words(1)
        
        
        On Error Resume Next
        
        s = Mid(pLineStr, InStrRev(pLineStr, ":") + 1)
        pResFileOffset = CLng("&H" & s)
        
        If Err.Number = 0 Then
            pValue = ""
        Else
            pResFileOffset = -1
        End If
        
        On Error GoTo 0
    
    Else
        pResFileOffset = -1
    End If
    
    If pResFileOffset <> -1 Then
        
        pValue = GetBinaryProperty(Me.ParentModule.ResFilePath, pResFileOffset)
        
    End If
    
    
End Function

Function Crypt()
        
    If pLineID <> 0 Then
        
        ParentModule.Lines(pLineID) = Rebuild
    
    End If
    
End Function

Function Rebuild() As String
    
    Dim SpaceCount
    Dim s As String
    Dim i As Long
    
    SpaceCount = Len(pLineStr) - Len(LTrim(pLineStr))
    
    
    s = String(SpaceCount, " ") & Name & " = " & Value
    
    
    'If Name = "ImageList" Then
    '    Debug.Assert False
    'End If

        
    For i = 1 To Me.ParentModule.ControlCount
        
        If Me.ParentModule.Controls(i).Alias.Name = FilterString(Value, Chr(34)) Then
            
            s = String(SpaceCount, " ") & Name & " = " & Chr(34) & Me.ParentModule.Controls(i).Alias.NewName & Chr(34)
            
        End If
        
    Next i

    
    
    If My.Config.EncryptControlCaptions Then
    
        If IsCaption Then
            s = String(SpaceCount, " ") & Name & " = " & Chr(34) & Chr(34)
        End If
    
    End If
    
    Rebuild = s
    
End Function


Private Function IsCaption() As Boolean
    
    Const ReplacedProperties As String = "Text|Caption|ToolTipText|List"
    
    Dim s() As String
    Dim i As Long
    
    s = Split(ReplacedProperties, "|")
    
    For i = 0 To UBound(s)
        
        If s(i) = Me.Name Then
            IsCaption = True
        End If
        
    Next i
    
End Function

Private Function GetNextResStart(LastOffset As Long) As Long
    
    Dim i As Long
    Dim x As String
    Dim NextOffset As Long
    
    
    For i = 1 To Me.ParentModule.PropertyCount
        
        If Me.ParentModule.Properties(i).ResFileOffset > LastOffset Then
            
            If NextOffset > 0 Then
                If Me.ParentModule.Properties(i).ResFileOffset < NextOffset Then
                    NextOffset = Me.ParentModule.Properties(i).ResFileOffset
                End If
            Else
                NextOffset = Me.ParentModule.Properties(i).ResFileOffset
            End If
            
        End If
        
    Next i
    
    
    
    
    If NextOffset = 0 Then
        GetNextResStart = FileLen(Me.ParentModule.ResFilePath)
    Else
        GetNextResStart = NextOffset
    End If
    
End Function

Private Function GetBinaryProperty(FilePath As String, Offset As Long)
    
    Dim n As Long
    Dim s As String
    Dim i As Long
  
    Dim FrxFilePath As String
    Dim StrStart As String
    Dim StartPos As Long
    
    Dim ValueList() As String
    

    n = InStr(pLineStr, "=")
    
    FrxFilePath = FilePath
    
    If n Then
    
        If InStr(pLineStr, GetFileTitle(FilePath)) Then
        
            Select Case True
            
                Case (InStr(1, LTrim(pLineStr), "Text") = 1) Or (InStr(1, LTrim(pLineStr), "Caption") = 1)
            
                
                    pLineStr = Trim$(Mid$(pLineStr, n + 1))
             
                    'If InStr(pLineStr, "$") = 1 Then
                
                   
                    'Debug.Print FrxFilePath, FrxFilePath
                    
                    StrStart = Mid(pLineStr, InStr(1, pLineStr, ":") + 1)
                    StartPos = CLng(("&H" & StrStart))
                    
                    's = ExtractString(pLineStr, FrxFilePath)
                    s = GetFrxString(FrxFilePath, StartPos)
        

                Case (InStr(1, LTrim(pLineStr), "List") = 1) Or (InStr(1, LTrim(pLineStr), "ItemData") = 1)
                
           
            
                    GetFrxList pLineStr, FrxFilePath, ValueList()
                    
                    ReDim pList(1 To UBound(ValueList))
                    
                    For i = 1 To UBound(ValueList())
                            
                        pList(i) = ValueList(i)
                        
                    Next i
                    
                    s = Join(ValueList(), "|")
                
                Case (InStr(1, LTrim(pLineStr), "Picture") = 1) Or (InStr(1, LTrim(pLineStr), "Image") = 1) Or (InStr(1, LTrim(pLineStr), "ToolboxBitmap") = 1)
                    
                    'Debug.Print "TODO: Read Image From: ", pLineStr
                    
                Case Else
                     
                    'Debug.Print pLineStr
                    'Debug.Assert False
            
            End Select
            
        Else
        
'            If chkBinaryOnly.Value = vbUnchecked Then
'
'                s = Trim$(Mid$(pLineStr, n + 1))
'                s = Mid$(s, 2)
'                s = Left$(s, Len(s) - 1)
'
'            End If
      
        End If
    
    End If
  
    GetBinaryProperty = s

    
End Function


Function GetFrxString(FilePath As String, StartPos As Long) As String
    
    Dim FileNum As Long
    Dim Header As FrxStringHeader
    Dim Numbers() As Long
    Dim i As Long
    Dim c As Long
    Dim StrLen As Long
    Dim Buffer As String
    Dim StrStartPos As Long
    
    StrStartPos = StartPos + 1
    
    FileNum = FreeFile
    
    Open FilePath For Binary As FileNum
        
    Get FileNum, StrStartPos, Header

    If Header.StringLen < 255 Then

        StrLen = Header.StringLen
        StrStartPos = StrStartPos + 1

    Else
        
        
        
        If StrLen = 0 Then
            
            If GetNextResStart(StartPos) <> 0 Then
                
                StrLen = GetNextResStart(StartPos) - StartPos
            
            Else
            
                 StrLen = FileLen(Me.ParentModule.ResFilePath) - (FileLen(Me.ParentModule.ResFilePath) - 4)
                 
            End If
            
           
            StrStartPos = StartPos + 4
            
        Else
        
            StrLen = Header.StringLen2
            StrStartPos = StrStartPos + 3

        End If
        

    End If
        
    Buffer = String(StrLen, " ")
        
    Get FileNum, (StrStartPos), Buffer
        
    Close FileNum
    
    GetFrxString = Buffer

End Function


Private Function GetFrxList(ByVal sString As String, ByVal sSourceFile As String, List() As String)
    
    On Error GoTo Error
  
    Dim nOffset   As Long
    Dim sFile     As String
    Dim n         As Long
    Dim abFile()  As Byte
    Dim arrData() As Variant
    Dim FF        As Long
    Dim i         As Long
    Dim Members   As Long
    Dim Longest   As Long
    Dim Found     As Long
    Dim j         As Long
    Dim nFileSize As Long
    Dim c As Long

    'Dim nSize     As Long
    'Dim sData     As String
    'Dim sBytes    As String
    'Dim bFileOpen As Boolean
    'TODO
    'On Error GoTo EI_ErrorHandler
    n = InStr(sString, ":")

    If n <> 0 Then
  
        sFile = sSourceFile 'AttachPath(StripQuotes(Left$(sString, n - 1)), ExtractPath(sSourceFile))
        
        'Debug.Print sFile
        If FileExist(sFile) Then
    
            sString = "&H" & Trim$(Mid$(sString, n + 1))
            nOffset = CLng(sString) ' + 1 '+ PrevOffset
            FF = FreeFile
      
            'On Error GoTo Out
            Open sFile For Binary As FF
      
            nFileSize = LOF(FF)
            'On Error GoTo 0
            ReDim abFile(nFileSize - 1) As Byte
            Get #FF, , abFile
        
            Close FF
      
            For i = nOffset To nFileSize - 1

                If abFile(i) <> 0 Then

                    Members = Members + abFile(i)

                Else

                    If Members Then

                        nOffset = i
                        Exit For

                    End If
                End If

            Next i
      
            ReDim arrData(Members - 1) As Variant
      
            For i = nOffset To nFileSize - 1

                If abFile(i) <> 0 Then

                    Longest = Longest + abFile(i)

                Else

                    If Longest Then

                        nOffset = i
                        Exit For

                    End If
                End If

            Next i

            '  For I = nOffset To nFileSize - 1
            '  If abFile(I) <> 0 Then
            '  GetFrxList = GetFrxList & Chr(abFile(I))
            '  Else
            '  If LenB(GetFrxList) Then
            '  Found = Found + 1
            '  If Found > Members Then
            '  Exit For
            '  End If
            '  End If
            '
            '  End If
            '  Next
            '  GetFrxList = ""
            i = nOffset + 1

            Do
      
                DoEvents
        
                If abFile(i) <> 0 Then
            
                    For j = i + 2 To i + 1 + abFile(i)
          
                        If j <= UBound(abFile) Then
            
                            arrData(Found) = arrData(Found) & Chr$(abFile(j))
              
                        Else
             
                            Exit For
              
                        End If
            
                    Next j

                    Found = Found + 1
          
                Else
        
                    Exit Do
                    '  GetFrxList = GetFrxList & Chr(0)
          
                End If
        
                i = j
        
            Loop While Found < Members
      
        End If
    
    End If
  
    'GetFrxList = Join(arrData, "|")
    ReDim List(1 To UBound(arrData))
    For i = LBound(arrData) To UBound(arrData)
        
        c = c + 1
        ReDim Preserve List(1 To c)
        List(c) = arrData(i)
        
    Next i
    
  
    Exit Function
Error:
    Debug.Print Err.Description
    'Debug.Assert False
    'Resume
    GetFrxList = ""
End Function
