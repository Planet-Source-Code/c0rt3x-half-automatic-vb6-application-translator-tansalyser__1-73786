VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fTranslate 
   Caption         =   "Translation"
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
   Icon            =   "fTranslation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Code"
      TabPicture(0)   =   "fTranslation.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListCode"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Controls"
      TabPicture(1)   =   "fTranslation.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListControls"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView ListCode 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListControls 
         Height          =   5895
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   5213
      TabIndex        =   6
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "-->|"
      Height          =   375
      Left            =   10920
      TabIndex        =   5
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<--"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<--"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "-->"
      Height          =   375
      Left            =   9360
      TabIndex        =   2
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox txtDst 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "<Translation>"
      Top             =   7320
      Width           =   11655
   End
   Begin VB.TextBox txtSrc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "<Original>"
      Top             =   6840
      Width           =   11655
   End
End
Attribute VB_Name = "fTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    
    With Me.ListCode
        
        .ColumnHeaders.Add , "ID", "ID"
        .ColumnHeaders.Add , "String", "String"
        .ColumnHeaders.Add , "Translation", "Translation"
        
    End With
    
    With Me.ListControls
        
        .ColumnHeaders.Add , "ID", "Control"
        .ColumnHeaders.Add , "String", "String"
        .ColumnHeaders.Add , "Translation", "Translation"
        
    End With
    
    LoadStrings

End Sub


Function LoadStrings()
    
    Dim x As Long
    Dim y As Long
    Dim z As Long
    
    Dim s As String
    
    Dim Item As ListItem
    
    
    
    For x = 1 To My.Project.StringCount
        
        'For y = 1 To My.Project.Modules(x).StringCount
            
            'Set Item = Me.ListCode.ListItems.Add(, , CStr(x) & "-" & CStr(y))
            
            'Item.SubItems(1) = My.Project.Modules(x).Strings(y).Value
            
        'Next y
    
        Set Item = Me.ListCode.ListItems.Add(, , CStr(x))
            
        Item.SubItems(1) = My.Project.Strings(x).Value
    
    Next x
    
    For x = 1 To My.Project.ModuleCount
        
        For y = 1 To My.Project.Modules(x).ControlCount
                
            For z = 1 To My.Project.Modules(x).Controls(y).PropertyCount
            
                Select Case My.Project.Modules(x).Controls(y).Properties(z).Name
                    
                    Case "Caption", "Text", "ToolTipText"
                                     
                        s = My.Project.Modules(x).Name
                        s = s & "." & My.Project.Modules(x).Controls(y).Name
                        
                        If My.Project.Modules(x).Controls(y).GetPropertyIndex("Index") <> 0 Then
                            
                            s = s & "(" & My.Project.Modules(x).Controls(y).Properties(My.Project.Modules(x).Controls(y).GetPropertyIndex("Index")).Value & ")"
                            
                        End If
                        
                        
                        s = s & "." & My.Project.Modules(x).Controls(y).Properties(z).Name
                        
                        Set Item = Me.ListControls.ListItems.Add(, , s)
                        
                        s = My.Project.Modules(x).Controls(y).Properties(z).Value
                        
                        s = Mid(s, 2, Len(s) - 2)
                        
                        Item.SubItems(1) = s
                        
                    Case Else
                    
                    
                End Select
                
                
            
            Next z
            
            
        
        Next y
        
    Next x
    
End Function


Private Sub Form_Resize()
    
    Dim i As Long
    
    For i = 1 To ListCode.ColumnHeaders.Count
        
        Me.ListCode.ColumnHeaders.Item(i).Width = (Me.ListCode.Width - (Me.ListCode.Width / 100 * 5)) / ListCode.ColumnHeaders.Count
        
    Next i
    
    For i = 1 To ListControls.ColumnHeaders.Count
        
        Me.ListControls.ColumnHeaders.Item(i).Width = (Me.ListControls.Width - (Me.ListControls.Width / 100 * 5)) / ListCode.ColumnHeaders.Count
        
    Next i
    
End Sub

Private Sub ListCode_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Me.txtSrc.Text = Item.SubItems(1)
    Me.txtDst.Text = Item.SubItems(2)
    
End Sub


Private Sub cmdSave_Click()
    Me.ListCode.SelectedItem.SubItems(2) = Me.txtDst.Text
End Sub

