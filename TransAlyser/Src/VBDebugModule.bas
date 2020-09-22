Attribute VB_Name = "VBDebugModule"
Option Explicit
 
Function TestErrHandler()
On Error GoTo OnErr
    
    Debug.Print 1 / 0

Exit Function
OnErr:
    
'<ErrHandler>
Select Case MsgBox( _
                        "Error: " & CStr(Err.Number) _
                        & vbNewLine & vbNewLine _
                        & "Description: " & Err.Description _
                        & vbNewLine & vbNewLine _
                        & "Source: <Module>.<Sub>" _
                            , vbCritical + vbAbortRetryIgnore)
    
    
    Case VbMsgBoxResult.vbRetry
            
        Resume
            
    Case VbMsgBoxResult.vbIgnore
            
        Resume Next
        
    Case VbMsgBoxResult.vbAbort
            
        'Exit
            
End Select

'</ErrHandler>
    
End Function
