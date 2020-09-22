Attribute VB_Name = "Common_ControlModule"
Public Declare Function SendMessage Lib "user32" _
 Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)

Option Explicit


'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte

  lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte

  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
  ntmFlags As Long
  ntmSizeEM As Long
  ntmCellHeight As Long
  ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

' tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1

Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

' EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

Declare Function EnumFontFamilies Lib "gdi32" Alias _
   "EnumFontFamiliesA" _
   (ByVal hdc As Long, ByVal lpszFamily As String, _
   ByVal lpEnumFontFamProc As Long, lParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
   ByVal hdc As Long) As Long

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
   ByVal FontType As Long, lParam As ListBox) As Long
Dim FaceName As String
Dim FullName As String
  FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
  EnumFontFamProc = 1

End Function

Sub FillComboWithFonts(cb As ComboBox)
Dim hdc As Long
  cb.Clear
  hdc = GetDC(cb.hwnd)
  EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, cb
  ReleaseDC cb.hwnd, hdc
End Sub


Function SizeListView(List As ListView)
    
    Dim i As Long
    
    With List
        
        For i = 1 To .ColumnHeaders.Count
            .ColumnHeaders(i).Width = (.Width - TwipsX(40)) / 5
        Next i
    
    End With

End Function

