VERSION 5.00
Begin VB.Form fRes 
   Caption         =   "ResSaver"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    SaveRes "DLLS.ZIP", "BIN", App.Path & "\DLLs.zip"
    
    End
    
End Sub


Function SaveRes(id, Typ, FilePath As String)
On Error GoTo Error
    
    Dim lFileNum As Long
    Dim b() As Byte
    
    On Error Resume Next
    lFileNum = FreeFile
    
    Open FilePath For Binary As lFileNum
        b = LoadResData(id, Typ)
        Put lFileNum, , b()
    Close lFileNum

Exit Function
Error:
    Debug.Print Err.Number, Err.Description
    Resume Next
End Function

