Attribute VB_Name = "DataBaseModule"
Option Explicit

Public Enum DataTypesEnum
    [eDataType - 01 - Byte] = 1
    [eDataType - 02 - Boolean]
    [eDataType - 03 - Integer]
    [eDataType - 04 - Long]
    [eDataType - 05 - Single]
    [eDataType - 06 - Double]
    [eDataType - 07 - Currency]
    [eDataType - 08 - Date]
    [eDataType - 09 - String]
    [eDataType - 10 - Variant]
End Enum

Public Enum eDataProperties
    [eDataProperties - 00 - ChildCount] = 0
    [eDataProperties - -1 - PropertyCount] = -1
    [eDataProperties - -2 - Name] = -2
End Enum

Public Enum PropertyPropertiesEnum
    [PropertyPropertiesEnum - 01 - Name] = 1
    [PropertyPropertiesEnum - 02 - DataType]
    [PropertyPropertiesEnum - 03 - Value]
End Enum

Public Function DataTypeCount() As Long
    DataTypeCount = 10
End Function

Public Function DataTypeName(eDataType As DataTypesEnum) As String
On Error GoTo Error
    Select Case eDataType
        Case 1
            DataTypeName = "Byte"
        Case 2
            DataTypeName = "Boolean"
        Case 3
            DataTypeName = "Integer"
        Case 4
            DataTypeName = "Long"
        Case 5
            DataTypeName = "Single"
        Case 6
            DataTypeName = "Double"
        Case 7
            DataTypeName = "Currency"
        Case 8
            DataTypeName = "Date"
        Case 9
            DataTypeName = "String"
        Case 10
            DataTypeName = "Variant"
    End Select
Exit Function
Error:
    Assert , "DataBaseModule.DataTypeName", Err.Number, Err.Description
    Resume Next
End Function

Sub dbSortArrayAsDouble(a() As DataBaseClass, PropertyIndex As Long)
On Error GoTo Error

    Dim u&, i&, j&, K&, h As DataBaseClass

    u = UBound(a)
    K = u \ 2
    
    While K > 0
        For i = 0 To u - K
            j = i
            While (j >= 0) And CDbl((a(j).Properties(PropertyIndex)) > CDbl(a(j + K).Properties(PropertyIndex)))
                Set h = a(j)
                Set a(j) = a(j + K)
                Set a(j + K) = h
                If j > K Then
                    j = j - K
                Else
                    j = 0
                End If
            Wend
        Next i
        K = K \ 2
    Wend
    
Exit Sub
Error:
    Assert , "DataBaseModule.dbSortArrayAsDouble", Err.Number, Err.Description
    Resume Next
End Sub
