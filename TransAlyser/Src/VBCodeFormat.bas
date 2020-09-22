Attribute VB_Name = "VBCodeFormat"
Option Explicit

Private vbKW() As String    ' vb Key Words
Private vbKWCount As Long

Private Const Alfabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

Private Enum NextItemIs

   None
   Word
   NewLine
   Comment
   QuotString
   Spaces
   Others
   
End Enum

' aids the VBCodeToHTM-function
' to determinate the sections(routines) in the code
Private Function CheckRow(ByVal txt As String, pos As Long) As Long
   If Mid(txt, pos, 1) = "'" Or Mid(txt, pos, 3) = "Rem" Then
      CheckRow = 1
   ElseIf Mid(txt, pos, 11) = "Private Sub" Then
      CheckRow = 2
   ElseIf Mid(txt, pos, 10) = "Public Sub" Then
      CheckRow = 2
   ElseIf Mid(txt, pos, 16) = "Private Function" Then
      CheckRow = 2
   ElseIf Mid(txt, pos, 15) = "Public Function" Then
      CheckRow = 2
   ElseIf Mid(txt, pos, 16) = "Private Property" Then
      CheckRow = 2
   ElseIf Mid(txt, pos, 15) = "Public Property" Then
      CheckRow = 2
   Else
      CheckRow = 0
   End If
End Function

' returns the type of a code item, its start and length byref
Private Function GetNextItem(txt As String, SelStart As Long, SelLen As Long) As NextItemIs
   
   Dim p As Long
   Dim np As Long
   Dim K As String * 1
   
   p = SelStart + SelLen + 1
   If p >= Len(txt) Then GetNextItem = None: Exit Function
   
   K = Mid(txt, p, 1)
   SelStart = p - 1
   np = p + 1

   ' Comments
   If K = "'" Or Mid(txt, p, 3) = "Rem" Then
      Do While np < Len(txt)
         K = Mid(txt, np, 1)
         If Asc(K) < 32 Then Exit Do
         np = np + 1
         DoEvents
      Loop
      SelLen = np - SelStart - 1
      GetNextItem = Comment
      Exit Function
   
   ' Word
   ElseIf InStr(Alfabet, K) > 0 Then
      Do While np < Len(txt)
         K = Mid(txt, np, 1)
         If Asc(K) < 33 Or InStr("():,", K) > 0 Then Exit Do
         np = np + 1
         DoEvents
      Loop
      SelLen = np - SelStart - 1
      GetNextItem = Word
      Exit Function
   
   ' Quot String
   ElseIf K = Chr(34) Then
   
      Do While np < Len(txt)
         K = Mid(txt, np, 1)
         If Asc(K) = 34 Then Exit Do
         np = np + 1
         DoEvents
      Loop
      
      SelLen = np - SelStart
      GetNextItem = QuotString
      
      Exit Function
   
   ' Spaces
   ElseIf K = " " Then
      Do While np < Len(txt)
         K = Mid(txt, np, 1)
         If Asc(K) <> 32 Then Exit Do
         np = np + 1
         DoEvents
      Loop
      SelLen = np - SelStart - 1
      GetNextItem = Spaces
      Exit Function
   
   ' Break - NewLine
   ElseIf K = Chr(13) Then
      SelLen = 2
      GetNextItem = NewLine
      Exit Function
   
   ' Others
   Else
   SelLen = np - SelStart - 1
   GetNextItem = Others
  
   End If
End Function

' determinates if a code-item is a keyword
Private Function IsKeyWord(ByVal txt As String) As Boolean
   
    Dim i As Long
   
    For i = LBound(vbKW) To UBound(vbKW)
   
        If txt = vbKW(i) Then
            IsKeyWord = True
            Exit Function
        End If
    
    Next i
   
End Function

' loads the keywordlist at the start of the program
Private Sub GetKeyWordList()
'   Dim ch As Long
'   Dim txt As String
'   Dim file As String
'   Dim k As String * 1
'   Dim Part As String
'   Dim P As Long, I As Long
'   ReDim vbKW(10) As String
'
'   file = App.Path & "\VBKeyW.txt"
'   ch = FreeFile
'   Open file For Input As ch
'   txt = Input(FileLen(file), ch)
'   Close ch
'
'   For P = 1 To Len(txt)
'      k = Mid(txt, P, 1)
'      Select Case k
'      Case ","
'         If Part <> "" Then
'            vbKW(I) = Part
'            If I = vbKWCount Then vbKWCount = vbKWCount + 10: ReDim Preserve vbKW(vbKWCount)
'            I = I + 1
'            Part = ""
'            End If
'      Case Is < " "
'      Case Else
'         Part = Part & k
'      End Select
'   Next P
'   vbKWCount = I

    Dim s As String
    
    s = fOptions.txtVBKeywords.Text
    
    s = Replace(s, vbNewLine, "")
   
    vbKW() = Split(s, ",")
    vbKWCount = UBound(vbKW) + 1
   
End Sub

Private Function GetFilePath(ByVal FullFilename As Variant) As Variant
   Dim i As Integer
   Dim K As String * 1

   For i = Len(FullFilename) To 1 Step -1
      K = Mid(FullFilename, i, 1)
      If K = "\" Then Exit For
   Next i
   GetFilePath = Left(FullFilename, i - 1)
End Function

Private Function GetFileTitle(ByVal FullFilename As Variant) As Variant
   Dim i As Integer
   Dim nm As String
   Dim K As String * 1

   For i = Len(FullFilename) To 1 Step -1
     K = Mid(FullFilename, i, 1)
     If K <> "\" Then nm = K + nm Else Exit For
   Next i
   GetFileTitle = nm
End Function

Private Sub ShowWithExplorer(ByVal file As String)

   Dim ipb As String, def As String
   On Error Resume Next
   
   def = "C:\Program Files\Internet Explorer\iexplore.exe "
   ipb = InputBox("Openen met onderstaande browser", "HTML teksten", def)
   
   If ipb = "" Then Exit Sub
   
   Shell def & file, vbNormalFocus
   
End Sub

' sets vbcolors to the text in a RichTextBox
Public Sub FormatVBCode(rtb As RichTextBox)
   
   Dim pss As Long
   Dim sS As Long
   Dim se As Long
   Dim nI As NextItemIs
   
   GetKeyWordList
   
   nI = Word
   
   With rtb
      
      '.Visible = False
      pss = .SelStart
      
      .SelStart = 0
      .SelLength = Len(.Text)
      .SelColor = 0
      .SelLength = 0
      
      While nI <> None
        
         sS = .SelStart
         se = .SelLength
         
         nI = GetNextItem(.Text, sS, se)
         
         .SelStart = sS
         .SelLength = se
         
         Select Case nI
         
            Case Word
               
               If IsKeyWord(.SelText) = True Then
                   
                   rtb.SelColor = RGB(0, 0, 128)
               
               Else
                   
                   rtb.SelColor = RGB(0, 0, 0)
               
               End If
            
            Case QuotString
                
                rtb.SelBold = True
                
            Case Comment
                
               rtb.SelColor = RGB(0, 128, 0)
               
            Case NewLine
                
               .SelStart = .SelStart + 2
               
            Case Else
               
         End Select
         
      Wend
      
      .SelStart = pss
      '.Visible = True
      
   End With
   
End Sub

Public Sub VBCodeToHTM(ByVal Code As String, ByVal file As String, ByVal Title As String)
   Dim ch As Long                ' file channel
   Dim nI As NextItemIs
   Dim sS As Long, sL As Long    ' Start & Length of an item
   Dim txt As String             ' item text to add to html-file row
   Dim ntxt As String            ' html-file row to print
   Dim CurType As Long           ' 1=normal, 2=keyw, 3=comment
   Dim pos As Long               ' character position within Code
   Dim Row As Long               ' current/counter row
   Dim CommentRows As Long       ' comments belonging to following routine
   ReDim rtd(20) As Long         ' storage sections
   Dim rtds As Long              ' sections count
   Dim rtdsmax As Long, r As Long  ' ubound and storage position
   
   Screen.MousePointer = 11
   
   ' count sections (routines) and store them
   If InStr(Code, vbCrLf) > 0 Then
      rtdsmax = 20               ' starting with a maximum sections of 20
      pos = 1: Row = 0           ' start positions
      Do
         pos = InStr(pos + 2, Code, vbCrLf)           ' find end of row
         If pos > 0 Then                              ' found
            Row = Row + 1                             ' new row
            Select Case CheckRow(Code, pos + 2)       ' row type?
               Case 0: CommentRows = 0                ' normal type, reset CommentRows-count
               Case 1: CommentRows = CommentRows + 1  ' comment type
               Case 2: rtd(rtds) = Row - CommentRows  ' new section type
                       rtds = rtds + 1                ' section counter
                       If rtds = rtdsmax Then rtdsmax = rtdsmax + 20: ReDim Preserve rtd(rtdsmax)
                       CommentRows = 0                ' reset CommentRows-count
            End Select
            Else
            Exit Do
            End If
      Loop
      Row = 0
      End If
      
   CurType = 1
   ch = FreeFile
   Open file For Output As ch
   ' html header + body start
   txt = ""
   txt = txt & "<!doctype html public '-//w3c//dtd html 4.0 transitional//en'>" & vbCrLf
   txt = txt & "<HTML>" & vbCrLf
   txt = txt & "  <HEAD>" & vbCrLf
   txt = txt & "     <META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=windows-1252'>" & vbCrLf
   txt = txt & "     <STYLE>" & vbCrLf
   txt = txt & "       <!--" & vbCrLf
   txt = txt & "        {font-family: Courier New;" & vbCrLf
   txt = txt & "            font-size: 10pt;}" & vbCrLf
   txt = txt & "        .RM  {color: #008000};" & vbCrLf
   txt = txt & "        .KW  {color: #000080};" & vbCrLf
   txt = txt & "        //-->" & vbCrLf
   txt = txt & "     </STYLE>" & vbCrLf
   txt = txt & "     <TITLE>" & Title & "</TITLE>" & vbCrLf
   txt = txt & "  </HEAD>" & vbCrLf
   txt = txt & "  <BODY BGCOLOR=#FFFFFF>" & vbCrLf
   txt = txt & "    <PRE>" & vbCrLf & "      " ' you could use <P> too
   Print #ch, txt;: txt = ""
   
   ' fill the body with colored vbcode text
   nI = Word
   sS = 0
   sL = 0
   While nI <> None
      nI = GetNextItem(Code, sS, sL)
      txt = Mid(Code, sS + 1, sL)
      ntxt = ""
      Select Case nI
      Case Word
         If IsKeyWord(txt) = True Then
            If CurType = 3 Then ntxt = "</SPAN><SPAN CLASS='KW'>"
            If CurType = 1 Then ntxt = "<SPAN CLASS='KW'>"
            CurType = 2
            Else
            If CurType <> 1 Then ntxt = "</SPAN>": CurType = 1
            End If
         ntxt = ntxt & txt
      Case Comment
         If CurType = 2 Then ntxt = "</SPAN><SPAN CLASS='RM'>"
         If CurType = 1 Then ntxt = "<SPAN CLASS='RM'>"
         CurType = 3
         ntxt = ntxt & txt
      Case NewLine
         Row = Row + 1
         If Row = rtd(r) Then
            ntxt = "<HR>" & vbCrLf & "      " ' new section
            r = r + 1
            Else
            ntxt = vbCrLf & "      " ' "<BR>" &   insert this when <P> is used
            End If
      Case Spaces
'                  when <P> is used, remove following Rem's
'         If Len(txt) > 1 Then
'            For I = 1 To Len(txt): ntxt = ntxt & "&nbsp;": Next I
'            Else
            ntxt = txt
'            End If
      Case Else
         If CurType <> 1 Then ntxt = "</SPAN>": CurType = 1
         ntxt = ntxt & txt
      End Select
      Print #ch, ntxt;: ntxt = ""
   Wend
   
   ' html - footer
   txt = txt & "     </PRE>" & vbCrLf ' you could use </P> too
   txt = txt & "  </BODY>" & vbCrLf
   txt = txt & "</HTML>" & vbCrLf
   Print #ch, txt;: txt = ""
   
   Close ch
   
   Screen.MousePointer = 0
End Sub


