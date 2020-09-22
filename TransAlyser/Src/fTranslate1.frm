VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form fTranslate1 
   Caption         =   "Translate"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picText 
      Height          =   375
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin vbalIml6.vbalImageList ilGrid 
      Left            =   240
      Top             =   6720
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   2296
      Images          =   "fTranslate1.frx":0000
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin VB.Frame FrameBottom 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   7680
      Width           =   11535
      Begin VB.CommandButton cmdBack 
         Caption         =   "< &Back"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >"
         Height          =   375
         Left            =   10200
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid Grid 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11456
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "fTranslate1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    
    
    Me.picText.Font = Me.Grid.Font
    
    LoadStrings
    
End Sub

Private Sub Form_Resize()

    With Grid
        
        .Left = 0
        .Top = 0
        
        
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - (Me.FrameBottom.Height) - (Me.FrameBottom.Left * 2)
        
    End With
    
        With Me.FrameBottom
        
       .Top = Me.Grid.Height + .Left
       .Width = Me.ScaleWidth - (.Left * 2)
       
    End With
    
    With cmdNext
        
        .Left = Me.FrameBottom.Width - .Width - Me.FrameBottom.Left
        
    End With
    
End Sub

Function LoadStrings()
    
    Dim i As Long
    Dim TxtWidth As Long
    
    With Grid
        
        .Clear True
        
        .AddColumn "ID", "ID", ecgHdrTextALignCentre
        .AddColumn "Module", "Module"
        .AddColumn "String", "String"
        
        .ImageList = Me.ilGrid
        
        .GridLines = True
        
    End With
    
    For i = 1 To My.Project.StringCount
        
        With Grid
            
            .AddRow
            
            .CellIcon(.Rows, 1) = 0
            
            .CellText(.Rows, 1) = Format(i, String(Len(CStr(My.Project.StringCount)), "0"))
            
            .CellText(.Rows, 2) = My.Project.Strings(i).ParentModule.Name
            
            .CellText(.Rows, 3) = My.Project.Strings(i).Value
            
            TxtWidth = GetTxtWidth(.CellText(.Rows, 2))
            If .ColumnWidth(2) < TxtWidth Then
                .ColumnWidth(2) = TxtWidth
            End If
            
            TxtWidth = GetTxtWidth(.CellText(.Rows, 3))
            If .ColumnWidth(3) < TxtWidth Then
                .ColumnWidth(3) = TxtWidth
            End If
            
        End With
        
    Next i
    
    
    
End Function


Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    
    If KeyCode = vbKeySpace Then
        
        With Grid
            
            If .CellIcon(.SelectedRow, 1) = 0 Then
                'ClearSelection
                .CellIcon(.SelectedRow, 1) = 1
            Else
                .CellIcon(.SelectedRow, 1) = 0
            End If
            
        End With
        
    End If
    
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean)
    
    Dim CurRow As Long
    Dim CurCol As Long
    
    If Button = 1 Then
    
         With Grid
            
            .CellFromPoint X \ 15, Y \ 15, CurRow, CurCol
            
            If CurCol = 1 Then
            
                If .CellIcon(CurRow, 1) = 0 Then
                    'ClearSelection
                    .CellIcon(CurRow, 1) = 1
                Else
                    .CellIcon(CurRow, 1) = 0
                End If
            
            End If
        
        End With
    
    End If
    
End Sub

Private Function GetTxtWidth(Text As String) As Long
    
    GetTxtWidth = (Me.picText.TextWidth(Text) / Screen.TwipsPerPixelX) + 8

    
End Function

