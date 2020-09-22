VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBase_PropertyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Parent As DataBaseClass

Private pValue() As String

Private Sub Class_Initialize()
On Error GoTo Error:
    
    Dim i&
    ReDim pValue(1 To 3)
    DataType = [eDataType - 09 - String]

Exit Sub
Error:
    Assert , "DataBase_PropertyClass.Class_Initialize", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub Class_Terminate()
On Error GoTo Error

    Set Parent = Nothing
    Erase pValue()
    
Exit Sub
Error:
    Assert , "DataBase_PropertyClass.Class_Terminate", Err.Number, Err.Description
    Resume Next
End Sub

Public Property Get Properties(eProperty As PropertyPropertiesEnum) As Variant
On Error GoTo Error

    Select Case eProperty
        Case [PropertyPropertiesEnum - 01 - Name]
            Properties = pValue(eProperty)
        Case [PropertyPropertiesEnum - 02 - DataType]
            Properties = CLng(pValue(eProperty))
        Case [PropertyPropertiesEnum - 03 - Value]
            If IsNumber And pValue([PropertyPropertiesEnum - 03 - Value]) = "" Then
                pValue([PropertyPropertiesEnum - 03 - Value]) = "0"
            End If
            Select Case CLng(pValue([PropertyPropertiesEnum - 02 - DataType]))
                Case [eDataType - 01 - Byte]
                    Properties = CByte(Asc(pValue(eProperty)))
                Case [eDataType - 02 - Boolean]
                    Properties = CBool(pValue(eProperty))
                Case [eDataType - 03 - Integer]
                    Properties = CInt(pValue(eProperty))
                Case [eDataType - 04 - Long]
                    Properties = CLng(pValue(eProperty))
                Case [eDataType - 05 - Single]
                    Properties = CSng(pValue(eProperty))
                Case [eDataType - 06 - Double]
                    Properties = CDbl(pValue(eProperty))
                Case [eDataType - 07 - Currency]
                    Properties = CCur(pValue(eProperty))
                Case [eDataType - 08 - Date]
                    If pValue(eProperty) = "" Then
                        Dim NewDate As Date
                        Properties = NewDate
                    Else
                        Properties = CDate(pValue(eProperty))
                    End If
                Case [eDataType - 09 - String]
                    Properties = CStr(pValue(eProperty))
                Case [eDataType - 10 - Variant]
                    Properties = CVar(pValue(eProperty))
                Case Else
                    Properties = pValue(eProperty)
            End Select
        End Select
        
Exit Property
Error:
    Assert , "DataBase_PropertyClass.Properties (Property Get)", Err.Number, Err.Description, "Property: '" & Property & "'"
    Resume Next
End Property

Public Property Let Properties(eProperty As PropertyPropertiesEnum, vNewValue)
On Error GoTo Error

    Dim lPropertyIndex As Long
    
    Select Case eProperty
        Case [PropertyPropertiesEnum - 01 - Name]
            pValue(eProperty) = CStr(vNewValue)
        Case [PropertyPropertiesEnum - 02 - DataType]
            pValue(eProperty) = CStr(vNewValue)
        Case [PropertyPropertiesEnum - 03 - Value]
            lPropertyIndex = Parent.GetPropertyIndex(Name)
            If lPropertyIndex <> -666 Then Parent.PropertyValueChanging Parent, lPropertyIndex, Properties(eProperty), vNewValue
            Select Case CLng(pValue([PropertyPropertiesEnum - 02 - DataType]))
                Case [eDataType - 01 - Byte]
                    If vNewValue = "" Then
                        pValue(eProperty) = 0
                    Else
                        pValue(eProperty) = Chr(vNewValue)
                    End If
                Case [eDataType - 02 - Boolean]
                    If Not vNewValue = "" Then
                        pValue(eProperty) = CStr(Abs(CInt(CBool(vNewValue))))
                    Else
                        pValue(eProperty) = "0"
                    End If
                Case [eDataType - 03 - Integer]
                    pValue(eProperty) = CStr(vNewValue)
                Case [eDataType - 04 - Long]
                    pValue(eProperty) = CStr(vNewValue)
                Case [eDataType - 05 - Single]
                    pValue(eProperty) = CStr(vNewValue)
                Case [eDataType - 06 - Double]
                    pValue(eProperty) = CStr(vNewValue)
                Case [eDataType - 07 - Currency]
                    pValue(eProperty) = CStr(vNewValue)
                Case [eDataType - 08 - Date]
                    If vNewValue = "" Then
                        Dim NewDate As Date
                        pValue(eProperty) = CStr(NewDate)
                    Else
                        pValue(eProperty) = CStr(vNewValue)
                    End If
                Case [eDataType - 09 - String]
                     pValue(eProperty) = CStr(vNewValue)
                Case [eDataType - 10 - Variant]
                     pValue(eProperty) = CStr(vNewValue)
                Case Else
                     pValue(eProperty) = CStr(vNewValue)
            End Select
            If lPropertyIndex <> -666 Then Parent.PropertyValueChanged Parent, lPropertyIndex, vNewValue
    End Select
    
Exit Property
Error:
    Assert , "DataBase_PropertyClass.Properties (Property Let)", Err.Number, Err.Description, "Property: '" & Property & "', NewValue: '" & vNewValue & "'"
    Resume Next
End Property

Public Function SetProperty(Property As PropertyPropertiesEnum, NewValue As String)
On Error GoTo Error

    pValue(Property) = NewValue

Exit Function
Error:
    Assert , "DataBase_PropertyClass.SetProperty", Err.Number, Err.Description, "Property: '" & Property & "', NewValue: '" & NewValue & "'"
    Resume Next
End Function

Public Function GetProperty(Property As PropertyPropertiesEnum) As String
On Error GoTo Error

    GetProperty = pValue(Property)

Exit Function
Error:
    Assert , "DataBase_PropertyClass.GetProperty", Err.Number, Err.Description, "Property: '" & Property & "'"
    Resume Next
End Function

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
On Error GoTo Error
    Value = Properties([PropertyPropertiesEnum - 03 - Value])
Exit Property
Error:
    Assert , "DataBase_PropertyClass.Value (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let Value(sNewValue As String)
On Error GoTo Error
    Properties([PropertyPropertiesEnum - 03 - Value]) = sNewValue
Exit Property
Error:
    Assert , "DataBase_PropertyClass.Value (Property Let)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Get Name() As String
On Error GoTo Error
    Name = Properties([PropertyPropertiesEnum - 01 - Name])
    Exit Property
Error:
    Assert , "DataBase_PropertyClass.Name (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let Name(NewValue As String)
On Error GoTo Error

    Properties([PropertyPropertiesEnum - 01 - Name]) = NewValue

Exit Property
Error:
    Assert , "DataBase_PropertyClass.Name (Property Let)", Err.Number, Err.Description, "NewValue: '" & NewValue & "'"
    Resume Next
End Property

Public Property Get DataType() As DataTypesEnum
On Error GoTo Error
    
    DataType = CLng(Properties([PropertyPropertiesEnum - 02 - DataType]))

Exit Property
Error:
    Assert , "DataBase_PropertyClass.DataType (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Let DataType(NewValue As DataTypesEnum)
On Error GoTo Error

    Properties([PropertyPropertiesEnum - 02 - DataType]) = NewValue

Exit Property
Error:
    Assert , "DataBase_PropertyClass.DataType (Property Let)", Err.Number, Err.Description, "NewValue: '" & NewValue & "'"
    Resume Next
End Property

Public Property Get IsNumber() As Boolean
On Error GoTo Error

    IsNumber = (pValue([PropertyPropertiesEnum - 02 - DataType]) < 8)
    
Exit Property
Error:
    Assert , "DataBase_PropertyClass.IsNumber (Property Get)", Err.Number, Err.Description
    Resume Next
End Property

Public Property Get DataString() As String
On Error Resume Next

    Dim sProperty$(1 To 3)
    
    sProperty(1) = Properties(1)
    sProperty(2) = Chr(CLng(Properties(2)))
    Select Case Properties(2)
        Case [eDataType - 02 - Boolean]
            sProperty(3) = CStr(Abs(CInt(Properties(3))))
        Case Else
            sProperty(3) = Properties(3)
    End Select
    DataString = JoinArray(sProperty)

End Property

Public Property Let DataString(sNewValue As String)
On Error Resume Next
    Dim sValue$()
    Dim i&
    SplitArray sNewValue, sValue
    For i = 1 To UBound(sValue)
        Select Case i
            Case [PropertyPropertiesEnum - 02 - DataType]
                Properties(i) = Asc(sValue(i))
            Case Else
                Properties(i) = sValue(i)
        End Select
    Next i
End Property
