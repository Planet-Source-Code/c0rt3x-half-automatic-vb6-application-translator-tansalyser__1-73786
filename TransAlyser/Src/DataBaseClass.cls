VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent                       As DataBaseClass

Private pProperty()                 As DataBase_PropertyClass
Private pChild()                    As DataBaseClass

Private Const FixPropertyCount      As Long = 3
Private Const FixDataCount          As Long = 1

Public Event ValueChanging(cData As DataBaseClass, eDataProperty As eDataProperties, vOldValue, vNewValue)
Public Event ValueChanged(cData As DataBaseClass, eDataProperty As eDataProperties, vNewValue)

'Public Event LoadingProgress(cData As DataBaseClass, Progress As Double)
'Public Event SavingProgress(cData As DataBaseClass, Progress As Double)

Private Sub Class_Initialize()
On Error GoTo Error
    
    Dim i&

    ReDim pProperty(PropertyLBound To 0)
    For i = PropertyLBound To 0
        Set pProperty(i) = New DataBase_PropertyClass
    Next i
    With pProperty([eDataProperties - -2 - Name])
        Set .Parent = Me
        .SetProperty [PropertyPropertiesEnum - 01 - Name], "Name"
        .SetProperty [PropertyPropertiesEnum - 02 - DataType], [eDataType - 09 - String]
    End With
    With pProperty([eDataProperties - -1 - PropertyCount])
        Set .Parent = Me
        .SetProperty [PropertyPropertiesEnum - 01 - Name], "PropertyCount"
        .SetProperty [PropertyPropertiesEnum - 02 - DataType], [eDataType - 04 - Long]
    End With
    With pProperty([eDataProperties - 00 - ChildCount])
        Set .Parent = Me
        .SetProperty [PropertyPropertiesEnum - 01 - Name], "ChildCount"
        .SetProperty [PropertyPropertiesEnum - 02 - DataType], [eDataType - 04 - Long]
    End With
    ReDim pChild(DataLBound To 0)

Exit Sub
Error:
    Assert , "DataBaseClass.Class_Initialize", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub Class_Terminate()
On Error GoTo Error
    Set Parent = Nothing
    Erase pProperty
    Erase pChild
Exit Sub
Error:
    Assert , "DataBaseClass.Class_Terminate", Err.Number, Err.Description
    Resume Next
End Sub

Public Property Get Path() As String
On Error GoTo Error

    If Not Parent Is Nothing Then
        Path = Parent.Path & "\" & Name
    Else
        Path = Name
    End If

Exit Property
Error:
    Assert , "DataBaseClass.Path (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Private Sub InitializeDB(DB As DataBaseClass)
On Error GoTo Error

    If DB Is Nothing Then Set DB = New DataBaseClass

Exit Sub
Error:
    Assert , "DataBaseClass.InitalizeDB", Err.Number, Err.Description
    Resume Next
End Sub

Public Property Get PropertyLBound() As Long
On Error GoTo Error

    PropertyLBound = 1 - FixPropertyCount

Exit Property
Error:
    Assert , "DataBaseClass.PropertyLBound (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Function DataLBound() As Long
On Error GoTo Error
    DataLBound = 1 - FixDataCount
Exit Function
Error:
    Assert , "DataBaseClass.DataLBound (Property Get)", Err.Number, Err.Description
    Resume Next
End Function

Public Function AddNewProperty(sName$, Optional DataType As DataTypesEnum = [eDataType - 09 - String], Optional Value = "")
On Error GoTo Error

    PropertyCount = PropertyCount + 1
    
    With Properties(PropertyCount)
        .Name = sName
        .DataType = DataType
        .Value = Value
    End With

Exit Function
Error:
    Assert , "DataBaseClass.AddNewProperty", Err.Number, Err.Description, "Name: '" & sName & "', DataType: '" & DataType & "', Value: '" & Value & "'"
    Resume Next
End Function

Public Function AddProperty(cProp As DataBase_PropertyClass)
On Error GoTo Error
    
    PropertyCount = PropertyCount + 1
    Properties(PropertyCount).DataString = cProp.DataString
    
Exit Function
Error:
    Assert , "DataBaseClass.AddProperty", Err.Number, Err.Description
    Resume Next
End Function

Public Function AddNewChild(sName$)
On Error GoTo Error

    ChildCount = ChildCount + 1
    Children(ChildCount).Name = sName

Exit Function
Error:
    Assert , "DataBaseClass.AddNewChild", Err.Number, Err.Description
    Resume Next
End Function

Public Function AddChild(cDat As DataBaseClass)
On Error GoTo Error
    ChildCount = ChildCount + 1
    Children(ChildCount) = cDat
Exit Function
Error:
    Assert , "DataBaseClass.AddChild", Err.Number, Err.Description
    Resume Next
End Function

Public Function AddChildFromFile(FilePath As String)
On Error GoTo Error
    AddNewChild ChildCount + 1
    Children(ChildCount).Load FilePath
    Exit Function
Error:
    Assert , "DataBaseClass.AddChildFromFile", Err.Number, Err.Description, "FilePath: '" & FilePath & "'"
    Resume Next
End Function

Public Function DeleteChild(Optional lIndex As Long, Optional sName$)
On Error GoTo Error
    Dim lIndex2
    Dim i&
    If lIndex = 0 Then
        lIndex2 = GetChildIndex(sName)
    Else
        lIndex2 = lIndex
    End If
    If lIndex2 <> ChildCount Then
        For i = lIndex2 To ChildCount - 1
            Children(i).DataString = Children(i + 1).DataString
        Next i
    End If
    ChildCount = ChildCount - 1
    Exit Function
Error:
    Assert , "DataBaseClass.DeleteChild", Err.Number, Err.Description, "lIndex: '" & lIndex & "', sNAme: '" & sName & "'"
    Resume Next
End Function

Public Function DeleteProperty(Optional lIndex&, Optional sName$)
On Error GoTo Error
    Dim lIndex2&
    Dim bFound As Boolean
    Dim i&
    If lIndex = 0 Then
        lIndex2 = GetPropertyIndex(sName)
    Else
        lIndex2 = lIndex
    End If
    If lIndex2 <> ChildCount Then
        For i = lIndex2 To PropertyCount - 1
            Properties(i).DataString = Properties(i + 1).DataString
        Next i
    End If
    PropertyCount = PropertyCount - 1
    Exit Function
Error:
    Assert , "DataBaseClass.DeleteProperty", Err.Number, Err.Description, "lIndex: '" & lIndex & "', sNAme: '" & sName & "'"
    Resume Next
End Function

Public Function GetPropertyIndex(sName$, Optional CaseSensitive As Boolean) As Long
Dim i&
Dim S1$, S2$
On Error GoTo Error
    If CaseSensitive Then S1 = sName Else S1 = LCase(sName)
    For i = PropertyLBound To PropertyCount
        If CaseSensitive Then S2 = Properties(i).Name Else S2 = LCase(Properties(i).Name)
        If S1 = S2 Then
            GetPropertyIndex = i
            Exit Function
        End If
    Next i
    GetPropertyIndex = -666
Exit Function
Error:
    Assert , "DataBaseClass.GetPropertyIndex", Err.Number, Err.Description, "sName: '" & sName & "', CaseSensitive: '" & CaseSensitive & "'"
    Resume Next
End Function

Public Function GetChildIndex(sName As String) As Long
Dim i&
On Error GoTo Error
    For i = DataLBound To ChildCount
        If Not Children(i) Is Nothing Then
            If Children(i).Name = sName Then
                GetChildIndex = i
                Exit Function
            End If
        End If
    Next i
Exit Function
Error:
    Assert , "DataBaseClass.GetChildIndex", Err.Number, Err.Description, "sName: '" & sName & "'"
    Resume Next
End Function

Public Function GetParentIndex() As Long
On Error GoTo Error
    GetParentIndex = Parent.GetChildIndex(Name)
Exit Function
Error:
    Assert , "DataBaseClass.GetParentIndex", Err.Number, Err.Description
    Resume Next
End Function

Public Property Get ChildTemplate() As DataBaseClass
On Error GoTo Error
    Set ChildTemplate = pChild(0)
Exit Property
Error:
    Assert , "DataBaseClass.ChildTemplate (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Set ChildTemplate(cNewData As DataBaseClass)
On Error GoTo Error
    Set pChild(0) = cNewData
Exit Property
Error:
    Assert , "DataBaseClass.ChildTemplate (Property Set)", Err.Number, Err.Description
    Resume Next
End Property

Public Function GetProperty(sPropertyName$, Optional sDataName$, Optional IncludeChildren As Boolean, Optional CaseSensitive As Boolean, Optional bContinue As Boolean) As DataBase_PropertyClass
On Error GoTo Error
    Static i&
    
    Dim bName As Boolean
    Dim lIndex As Long
    
    Dim S1$, S2$
    
    If CaseSensitive Then S1 = sPropertyName Else S1 = LCase(sPropertyName)
    
    If sDataName <> "" Then bName = True
    
    If bName = False Or Name = sDataName Then
        lIndex = GetPropertyIndex(sPropertyName, CaseSensitive)
        If lIndex <> -666 Then
            If lIndex <> 0 Then
                Set GetProperty = Properties(lIndex)
                Exit Function
            End If
        End If
    End If
    
    If Not IncludeChildren Then
        Exit Function
    End If
    
    If Not bContinue Then i = 0
    
    For i = i + 1 To ChildCount
        If bName = False Or Children(i).Name = sDataName Then
            If bName Then
                Set GetProperty = Children(i).GetProperty(sPropertyName, sDataName, , CaseSensitive, bContinue)
            Else
                Set GetProperty = Children(i).GetProperty(sPropertyName, , CaseSensitive, bContinue)
            End If
            If CaseSensitive Then S2 = GetProperty.Name Else S2 = LCase(GetProperty.Name)
            If S1 = S2 Then Exit Function
        End If
    Next i
    
    Exit Function
Error:
    Assert , "DataBaseClass.PropertyValueChanging", Err.Number, Err.Description, "PropertyName: '" & sPropertyName & "', DataName: '" & sDataName & "', CaseSensitive: '" & CaseSensitive & "', Continue: '" & bContinue & "'"
    Resume Next
End Function

Public Property Get Property(sName As String) As DataBase_PropertyClass
On Error GoTo Error

    Set Property = GetProperty(sName)

Exit Property
Error:
    Assert , "DataBaseClass.Property (Property Get)", Err.Number, Err.Description, "sName: '" & sName & "'"
    Resume Next
End Property

Public Property Get Properties(DataProperty As eDataProperties) As DataBase_PropertyClass
On Error GoTo Error

    Set Properties = pProperty(DataProperty)

Exit Property
Error:
    Assert , "DataBaseClass.Properties (Property Get)", Err.Number, Err.Description, "DataProperty: '" & DataProperty & "'"
    Resume Next
End Property

Public Property Set Properties(DataProperty As eDataProperties, cNewProperty As DataBase_PropertyClass)
On Error GoTo Error
    
    Set pProperty(DataProperty) = cNewProperty
    
Exit Property
Error:
    Assert , "DataBaseClass.Properties (Property Get)", Err.Number, Err.Description, "DataProperty: '" & DataProperty & "'"
    Resume Next
End Property

Public Function PropertyValueChanging(cData As DataBaseClass, eDataProperty As eDataProperties, vOldValue, vNewValue)
On Error GoTo Error

    Dim i As Long
    
    'If vNewValue <> pProperty(eDataProperty).Value Then
        RaiseEvent ValueChanging(cData, eDataProperty, vOldValue, vNewValue)
        If Not Parent Is Nothing Then
            Parent.PropertyValueChanging cData, eDataProperty, vOldValue, vNewValue
        End If
    'End If
    If cData.Path <> Path Then
        Exit Function
    End If
    Select Case eDataProperty
        Case [eDataProperties - -1 - PropertyCount]
            ReDim Preserve pProperty(PropertyLBound To vNewValue)
            If CLng(vNewValue) > PropertyCount Then
                For i = PropertyCount + 1 To vNewValue
                    Set pProperty(i) = New DataBase_PropertyClass
                    Set pProperty(i).Parent = Me
                Next i
            End If
        Case [eDataProperties - 00 - ChildCount]
            ReDim Preserve pChild(DataLBound To CLng(vNewValue))
            If CLng(vNewValue) > ChildCount Then
                For i = ChildCount + 1 To CLng(vNewValue)
                    Set pChild(i) = New DataBaseClass
                    Set pChild(i).Parent = Me
                    If Not ChildTemplate Is Nothing Then
                        pChild(i).DataString = ChildTemplate.DataString
                    End If
                Next i
            End If
    End Select
    
Exit Function
Error:
    Assert , "DataBaseClass.PropertyValueChanging", Err.Number, Err.Description, "DataProperty: '" & eDataProperty & "'"
    Resume Next
End Function

Public Function PropertyValueChanged(cData As DataBaseClass, eDataProperty As eDataProperties, vNewValue)
On Error GoTo Error
    RaiseEvent ValueChanged(cData, eDataProperty, vNewValue)
    If Not Parent Is Nothing Then
        Parent.PropertyValueChanged cData, eDataProperty, vNewValue
    End If
Exit Function
Error:
    Assert , "DataBaseClass.PropertyValueChanged", Err.Number, Err.Description
    Resume Next
End Function

Public Property Get PropertyCount() As Long
On Error GoTo Error
    PropertyCount = Properties([eDataProperties - -1 - PropertyCount])
Exit Property
Error:
    Assert , "DataBaseClass.PropertyCount (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let PropertyCount(lNewPropertyCount&)
On Error GoTo Error
    Properties([eDataProperties - -1 - PropertyCount]) = lNewPropertyCount
Exit Property
Error:
    Assert , "DataBaseClass.PropertyCount (Property Let)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Get ChildCount() As Long
On Error GoTo Error
    ChildCount = Properties([eDataProperties - 00 - ChildCount])
Exit Property
Error:
    Assert , "DataBaseClass.ChildCount (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let ChildCount(lValue As Long)
On Error GoTo Error
    Properties([eDataProperties - 00 - ChildCount]) = lValue
Exit Property
Error:
    Assert , "DataBaseClass.ChildCount (Property Let)", Err.Number, Err.Description, "lValue: '" & lValue & "'"
    Resume Next
End Property

Public Property Get Name() As String
Attribute Name.VB_MemberFlags = "200"
On Error GoTo Error
    Name = Properties([eDataProperties - -2 - Name])
Exit Property
Error:
    Assert , "DataBaseClass.Name (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let Name(sValue As String)
On Error GoTo Error
    Properties([eDataProperties - -2 - Name]) = sValue
Exit Property
Error:
    Assert , "DataBaseClass.Name (Property Let)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Get Children(lIndex As Long) As DataBaseClass
On Error GoTo Error

    'InitializeDB pChild(lIndex)
    Set Children = pChild(lIndex)
    
Exit Property
Error:
    Assert , "DataBaseClass.Children (Property Get)", Err.Number, Err.Description, "lIndex: '" & lIndex & "'"
    Resume Next
End Property

Public Property Get Child(sName As String) As DataBaseClass
On Error GoTo Error

    Set Child = FindChild("Name", sName)
    
Exit Property
Error:
    Assert , "DataBaseClass.Child (Property Get)", Err.Number, Err.Description, "sName: '" & sName & "'"
    Resume Next
End Property

Public Function FindChild(PropertyName As String, PropertyValue As String, Optional SearchDeep As Long = 1, Optional CaseSensitive As Boolean, Optional bContinue As Boolean, Optional bIncludeMe As Boolean) As DataBaseClass
On Error GoTo Error

    Static i As Long
    Dim S1$, S2$
        
        If CaseSensitive Then S1 = PropertyValue Else S1 = LCase(PropertyValue)
        If CaseSensitive Then S2 = Property(PropertyName) Else S2 = LCase(Property(PropertyName))
        If bIncludeMe Then
            If S1 = S2 Then
                Set FindChild = Me
                Exit Function
            End If
        End If
        If Not bContinue Then i = 0
        If SearchDeep < 1 Then
            'Set FindChild = New DataBaseClass
            Exit Function
        End If
        For i = i + 1 To ChildCount
            Set FindChild = pChild(i).FindChild(PropertyName, PropertyValue, SearchDeep - 1, CaseSensitive, bContinue, True)
            If Not FindChild Is Nothing Then
                If CaseSensitive Then S2 = FindChild.Property(PropertyName) Else S2 = LCase(FindChild.Property(PropertyName))
                If S1 = S2 Then Exit Function
            End If
        Next i
        'Set FindChild = New DataBaseClass
        
Exit Function
Error:
    Assert , "DataBaseClass.FindChild", Err.Number, Err.Description, "PropertyName: '" & PropertyName & "', PropertyValue: '" & PropertyValue & "', SearchDeep: '" & SearchDeep & "', CaseSensitive: '" & CaseSensitive & "', bContinue: '" & bContinue & "', bIncludeMe: '" & bIncludeMe & "'"
    Resume Next
End Function

Public Function FindPath(sPath As String) As DataBaseClass
On Error GoTo Error

    Dim Piece() As String
    Dim SubPath As String
    Dim i As Long
    
    If Not Child(sPath) Is Nothing Then
        Set FindPath = Child(sPath)
        Exit Function
    End If
    Piece = Split(sPath, "\")
    For i = 1 To UBound(Piece)
        SubPath = SubPath & Piece(i)
        If i <> UBound(Piece) Then
            SubPath = SubPath & "\"
        End If
    Next i
    If Not FindChild("Name", Piece(0), , , , True) Is Nothing Then Set FindPath = FindChild("Name", Piece(0), , , , True).FindPath(SubPath)
    
    Exit Function
Error:
    Assert , "DataBaseClass.FindPath", Err.Number, Err.Description, "sPath: '" & sPath & "'"
    Resume Next
End Function

Function ExcludeChildren() As DataBaseClass
On Error GoTo Error
    
    Dim p As Long
    
    Set ExcludeChildren = New DataBaseClass
    For p = PropertyLBound To PropertyCount
        ExcludeChildren.Properties(p).DataString = Properties(p).DataString
    Next p
    If Not ChildTemplate Is Nothing Then
        Set ExcludeChildren.ChildTemplate = New DataBaseClass
        ExcludeChildren.ChildTemplate = ChildTemplate
    End If
    
Exit Function
Error:
    Assert , "DataBaseClass.ExcludeChildren", Err.Number, Err.Description
    Resume Next
End Function

Public Function GetDataString(Optional UseCompression As Boolean = True, Optional UseEncryption As Boolean = True, Optional Password As String = "FirstCode DataBase Version 1.0") As String
On Error GoTo Error

    Dim Data As String
    
    Data = DataString
    If UseCompression Then Data = Compress(Data)
    If UseEncryption Then Data = RC4(Data, Password)
    GetDataString = Data
    
Exit Function
Error:
    Assert , "DataBaseClass.GetDataString", Err.Number, Err.Description, "UseCompression: '" & UseCompression & "', UseEncryption: '" & UseEncryption & "', Password:= '" & Password & "'"
    Resume Next
End Function

Public Function SetDataString(Data As String, Optional UseCompression As Boolean = True, Optional UseEncryption As Boolean = True, Optional Password As String = "FirstCode DataBase Version 1.0") As Boolean
On Error GoTo Error

    Dim sData As String

    If UseEncryption Then
        sData = RC4(Data, Password)
    Else
        sData = Data
    End If
    If UseCompression Then
        sData = Decompress(sData)
    End If
    DataString = sData
    SetDataString = True
    
Exit Function
Error:
   Assert , "DataBaseClass.SetDataString", Err.Number, Err.Description
   Resume Next
End Function

Public Property Get DataString() As String
Attribute DataString.VB_UserMemId = 0
On Error GoTo Error

    Dim i&
    Dim sProperty$()
    Dim sData$()

    ReDim sProperty(PropertyLBound To PropertyCount)
    ReDim sData(DataLBound - 1 To ChildCount)
    For i = PropertyLBound To PropertyCount
        sProperty(i) = Properties(i).DataString
    Next i
    sData(DataLBound - 1) = JoinArray(sProperty)

    For i = DataLBound To ChildCount
        If Not Children(i) Is Nothing Then
            sData(i) = Children(i).DataString
        End If
    Next i

    DataString = JoinArray(sData)
    
    Exit Property
Error:
    Assert , "DataBaseClass.DataString (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let DataString(sDataString As String)
On Error GoTo Error

    Dim i&
    Dim c As Long
    Dim sProperty$()
    Dim sData$()

    If sDataString = "" Then Exit Property
    
    SplitArray sDataString, sData()
    
    c = c + 1
    SplitArray sData(c), sProperty 'Decompress(sData(c)), sProperty
    For i = 1 To UBound(sProperty)
        Properties(i - FixPropertyCount).DataString = sProperty(i)
    Next i
    c = c + 1
    If sData(c) <> "" Then
        InitializeDB pChild(0)
        pChild(0).DataString = sData(c)
    End If
    
    For i = 1 To ChildCount
        c = c + 1
        Children(i).DataString = sData(c)
    Next i
    
Exit Property
Error:
    Assert , "DataBaseClass.DataString (Property Let)", Err.Number, Err.Description
    If InIDE Then Resume Next
End Property

Function SwapChilds(Child1 As Long, Child2 As Long)
On Error GoTo Error
    
    Dim s As String
    
    s = Children(Child1)
    Children(Child1) = Children(Child2)
    Children(Child2) = s

Exit Function
Error:
    Assert , "DataBaseClass.SwapChilds", Err.Number, Err.Description, "Child1: '" & Child1 & "', Child2: '" & Child2 & "'"
    Resume Next
End Function

Function AddChilds(dbSource As DataBaseClass)
On Error GoTo Error

    Dim c As Long
    
    For c = 1 To dbSource.ChildCount
        AddChild dbSource.Children(c)
    Next c

Exit Function
Error:
    Assert , "DataBaseClass.AddChilds", Err.Number, Err.Description
    Resume Next
End Function

Function Save(FilePath As String)
On Error GoTo Error

    If FileExist(FilePath) Then DeleteFile FilePath
    WriteFile FilePath, GetDataString
    Exit Function

Error:
    Assert , "DataBaseClass.Save", Err.Number, Err.Description, "FilePath: '" & FilePath & "'"
    Resume Next
End Function

Function Load(FilePath As String)
On Error GoTo Error

    SetDataString ReadFile(FilePath)
    Exit Function

Error:
   Assert , "DataBaseClass.Load", Err.Number, Err.Description, "FilePath: '" & FilePath & "'"
   Resume Next
End Function

Public Function DebugPrint(Optional IncludeChildren As Boolean, Optional sTab As String)
On Error GoTo Error

    Dim i As Long
    
    Debug.Print sTab & MultiStr("#", 100)
    For i = PropertyLBound To PropertyCount
        Debug.Print sTab, i, "PropertyName: " & pProperty(i).Name, "", "Value: " & pProperty(i).Value
    Next i
    Debug.Print sTab & MultiStr("#", 100) & vbNewLine
    If Not IncludeChildren Then Exit Function
    On Error Resume Next
    For i = DataLBound To ChildCount
        Children(i).DebugPrint True, sTab & "----"
    Next i
    
Exit Function
Error:
    Assert , "DataBaseClass.DebugPrint", Err.Number, Err.Description, "IncludeChildren: '" & IncludeChildren & "'"
    Resume Next
End Function
