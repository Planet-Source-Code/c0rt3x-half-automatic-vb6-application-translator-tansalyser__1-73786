VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open project"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fOpen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   2520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView List 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      Icons           =   "ilList"
      SmallIcons      =   "ilList"
      ColHdrIcons     =   "ilList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Browse..."
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   8760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fOpen.frx":000C
            Key             =   "Project"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgProject 
      Height          =   480
      Left            =   120
      Picture         =   "fOpen.frx":0166
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "fOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()
    
    Hide
    DoEvents
    
    With Dialog1
    
        .Filter = "VB-Projektdateien (*.VBP)|*.VBP"
        .InitDir = "C:\"
        .ShowOpen
    
        If FileExist(.FileName) Then
            fMain.LoadProject .FileName
            Unload Me
        Else
            Show
            DoEvents
        End If
    
    End With
    
End Sub

Private Sub Form_Load()

    Dim Reg As MyRegistryClass
    Set Reg = New MyRegistryClass
    Dim Name() As String
    
    Dim c As Long ' Number of Projects
    Dim e As Long ' Number of existing Projects
    Dim i As Long
    
    Reg.ClassKey = HKEY_CURRENT_USER
    Reg.SectionKey = "Software\Microsoft\Visual Basic\6.0\RecentFiles"
    Reg.EnumerateValues Name(), c
    
    With List
    
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, "K1", "Projekt", Width * 0.92
        .ListItems.Clear
        
        For i = 1 To c
        
            Reg.ValueKey = Name(i)
            
            If Mid(Reg.Value, 2, 1) = ":" Then
            
                If FileExist(Reg.Value) Then
                
                    e = e + 1
                    
                    .ListItems.Add e, "K" & i, Reg.Value, , "Project"
                    
                End If
            
            End If
            
        Next i
        
    End With
    
End Sub

Private Sub List_DblClick()
    Hide
    DoEvents
    fMain.LoadProject List.SelectedItem.Text
End Sub
