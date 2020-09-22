VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form fOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
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
   Icon            =   "fOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Gblobal"
      TabPicture(0)   =   "fOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHighLight"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtVBKeywords"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtIgnore"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Keywords"
      TabPicture(1)   =   "fOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtIgnore 
         Height          =   1575
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "fOptions.frx":0044
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtVBKeywords 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "fOptions.frx":0081
         Top             =   4320
         Width           =   8895
      End
      Begin VB.Frame Frame8 
         Caption         =   "Protected words"
         Height          =   5295
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   8775
         Begin VB.TextBox txtKeywords 
            Height          =   1455
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Text            =   "fOptions.frx":062E
            Top             =   600
            Width           =   8295
         End
         Begin VB.TextBox txtProperties 
            Height          =   1575
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Text            =   "fOptions.frx":0A4E
            Top             =   2520
            Width           =   8295
         End
         Begin VB.TextBox txtMiscKeywords 
            Height          =   615
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "fOptions.frx":4E96
            Top             =   4440
            Width           =   8295
         End
         Begin VB.Label Label1 
            Caption         =   "Protected Keywords:"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Protected Properties:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Misc Keywords:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   4200
            Width           =   1935
         End
      End
      Begin VB.Label lblHighLight 
         Caption         =   "HighLight:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
End
Attribute VB_Name = "fOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Me.Hide
    
End Sub

Private Sub cmdCancel_Click()
    
    Me.Hide
    
End Sub

Private Sub cmdOK_Click()
        
    SaveConfig
    
    Me.Hide

End Sub


Function SaveConfig()
    
    With My.Config
    
         

    End With
    
    My.Config.SaveConfig

End Function


Function LoadConfig()
On Error Resume Next

    With My.Config
    
        
    End With
    
End Function

