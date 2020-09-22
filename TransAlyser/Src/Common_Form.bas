Attribute VB_Name = "mForm"
Option Explicit

Const UseAnimation As Boolean = False


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nindex As Integer) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nindex As Integer, ByVal dwnewlong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Const GWL_STYLE = (-16)
Const WS_THICKFRAME = &H40000

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Declare Function DeactivateWindowTheme Lib "uxtheme" _
         Alias "SetWindowTheme" ( _
     ByVal hwnd As Long, _
     Optional ByRef pszSubAppName As String = " ", _
     Optional ByRef pszSubIdList As String = " ") As Long
     
Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Boolean

Enum SidesEnum
    LeftSide
    RightSide
    TopSide
    ButtomSide
End Enum

Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
    
    Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hwnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
    
End Sub

Function NoClose(f As Form)
    
    RemoveMenus f, False, False, _
        False, False, False, True, True
    
End Function

Public Function GetForm(Name As String) As Form
Dim f As Form
    For Each f In Forms
        If LCase(f.Name) = LCase(Name) Then
            Set GetForm = f
            Exit Function
        End If
    Next f
End Function

Function CenterInForm(fParent As Form, fChild As Form)
    If fChild.MDIChild Then
        fChild.Move (fParent.Width - fChild.Width) / 2, (fParent.Height - fChild.Height) / 2
    Else
        fChild.Move (fParent.Width - fChild.Width) / 2 + fParent.Left, (fParent.Height - fChild.Height) / 2 + fParent.Top
    End If
End Function

Function CenterInFrame(c As Control, f As Frame)
    
    c.Move (f.Width - c.Width) / 2, (f.Height - c.Height) / 2

End Function

Function RemoveBorder(hwnd As Long)
    Dim CurStyle As Long
    Dim NewStyle As Long
    CurStyle = GetWindowLong(hwnd, GWL_STYLE)
    NewStyle = SetWindowLong(hwnd, GWL_STYLE, 0)
End Function

'Function SetGlobalFont(Font As StdFont)
'Dim f As Form
'Dim c As Control
'    On Error Resume Next
'    For Each f In VB.Forms
'        For Each c In f.Controls
'            If TypeOf c Is StyleButton Then
'                Set c.Font = Font
'            End If
'        Next c
'    Next f
'End Function

Function TwipsX(Pixels As Long) As Long
    TwipsX = Pixels * Screen.TwipsPerPixelX
End Function

Function TwipsY(Pixels As Long) As Long
    TwipsY = Pixels * Screen.TwipsPerPixelY
End Function

Function UsingXPTheme() As Boolean
On Error Resume Next
    UsingXPTheme = IsThemeActive
End Function







Function Msg(Prompt, Optional Buttons As VbMsgBoxStyle, Optional Title, Optional HelpFile, Optional Context) As VbMsgBoxResult
    
    Dim f As Form
    
    For Each f In Forms
        If f.Tag = "OnTop" Then
            AllwaysOnTop f, False
            f.Tag = "OnTop=0"
        End If
    Next f
    
    Msg = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    
    For Each f In Forms
        If f.Tag = "OnTop=0" Then
            AllwaysOnTop f, True
        End If
    Next f
    
End Function


Function AllwaysOnTop(Window As Form, Active As Boolean)
On Error Resume Next
    If Active Then
        SetWindowPos Window.hwnd, -1, 0, 0, 0, 0, 3
        Window.Tag = "OnTop"
    Else
        SetWindowPos Window.hwnd, -2, 0, 0, 0, 0, 3
        Window.Tag = ""
    End If
    DoEvents
End Function
