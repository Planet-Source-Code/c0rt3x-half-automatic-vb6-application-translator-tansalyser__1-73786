VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyPaths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Property Get AppPath() As String
    
    If InIDE Then
        AppPath = FixPath(App.Path) & "Bin\"
    Else
        AppPath = FixPath(App.Path)
    End If

End Property

Property Get DemoProjectPath() As String
    
    'DemoProjectPath = AppPath & "DemoProjects\Sudoku\Sudoku.vbp"
    DemoProjectPath = AppPath & "DemoProjects\TranslationDemo\TranslationDemo.vbp"
End Property

