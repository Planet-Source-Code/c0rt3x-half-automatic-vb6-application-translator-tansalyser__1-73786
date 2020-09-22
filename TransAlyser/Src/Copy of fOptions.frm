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
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Gblobal"
      TabPicture(0)   =   "fOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Declarations"
      TabPicture(1)   =   "fOptions.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Functions"
      TabPicture(2)   =   "fOptions.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Forms"
      TabPicture(3)   =   "fOptions.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Advanced"
      TabPicture(4)   =   "fOptions.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Keywords"
      TabPicture(5)   =   "fOptions.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame8"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame8 
         Caption         =   "Protected words"
         Height          =   5295
         Left            =   -74760
         TabIndex        =   43
         Top             =   480
         Width           =   8775
         Begin VB.TextBox txtProperties 
            Height          =   1695
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Text            =   "fOptions.frx":00B4
            Top             =   3360
            Width           =   8295
         End
         Begin VB.TextBox txtKeywords 
            Height          =   2175
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Text            =   "fOptions.frx":0131
            Top             =   600
            Width           =   8295
         End
         Begin VB.Label Label2 
            Caption         =   "Protected Properties:"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Protected Keywords:"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Misc"
         Height          =   5295
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   8775
         Begin VB.CheckBox chkCrackCheck 
            Caption         =   "Use CrackCheck"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1080
            Width           =   4935
         End
         Begin VB.CheckBox chkExportStrings 
            Caption         =   "Export strings to file"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   4935
         End
         Begin VB.TextBox txtOnError 
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "fOptions.frx":0551
            Top             =   2640
            Width           =   8295
         End
         Begin VB.TextBox txtErrorHandler 
            Height          =   1095
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   32
            Text            =   "fOptions.frx":0567
            Top             =   3240
            Width           =   8295
         End
         Begin VB.CheckBox chkAddErrorHandlers 
            Caption         =   "Add Error Handlers:"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   2040
            Width           =   4935
         End
         Begin VB.CheckBox chkDebugMode 
            Caption         =   "Debug Mode: (Adds debug code)"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1680
            Width           =   4935
         End
         Begin VB.CheckBox chkEncryptStrings 
            Caption         =   "Encrypt hardcoded strings"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Forms"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   25
         Top             =   480
         Width           =   8895
         Begin VB.CheckBox chkEncryptControlCaptions 
            Caption         =   "encrypt control captions"
            Height          =   255
            Left            =   480
            TabIndex        =   42
            Top             =   1200
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt control names"
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   27
            Top             =   840
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt controls"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Functions/Subs/Properties"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   8895
         Begin VB.CheckBox Check 
            Caption         =   "encrypt function parameters"
            Height          =   255
            Index           =   16
            Left            =   480
            TabIndex        =   24
            Top             =   1200
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt function names"
            Height          =   255
            Index           =   15
            Left            =   480
            TabIndex        =   23
            Top             =   840
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt functions (and subs/property handlers)"
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Declarations"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   8895
         Begin VB.CheckBox Check 
            Caption         =   "encrypt declaration sub/function parameter names"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   20
            Top             =   1200
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt declaration sub/function names"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   19
            Top             =   840
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt declarations"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   4935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Enums"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   13
         Top             =   4080
         Width           =   8895
         Begin VB.CheckBox Check 
            Caption         =   "encrypt Enums"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt Enum names"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   15
            Top             =   720
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt Enum member value names"
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   14
            Top             =   1080
            Width           =   4935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Types"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   9
         Top             =   2280
         Width           =   8895
         Begin VB.CheckBox Check 
            Caption         =   "encrypt types"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt type names"
            Height          =   255
            Index           =   9
            Left            =   480
            TabIndex        =   11
            Top             =   840
            Width           =   4935
         End
         Begin VB.CheckBox Check 
            Caption         =   "encrypt type member variable names"
            Height          =   255
            Index           =   10
            Left            =   480
            TabIndex        =   10
            Top             =   1200
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Global"
         Height          =   5295
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   8775
         Begin VB.CheckBox chkRemoveSpaces 
            Caption         =   "Remove spaces"
            Height          =   255
            Left            =   720
            TabIndex        =   39
            Top             =   4320
            Width           =   4935
         End
         Begin VB.CheckBox chkRemoveComments 
            Caption         =   "Remove comments"
            Height          =   255
            Left            =   720
            TabIndex        =   38
            Top             =   2880
            Value           =   1  'Checked
            Width           =   4935
         End
         Begin VB.CheckBox chkRemoveCodeLayout 
            Caption         =   "Remove code layout"
            Height          =   255
            Left            =   480
            TabIndex        =   37
            Top             =   2400
            Value           =   1  'Checked
            Width           =   4935
         End
         Begin VB.CheckBox chkRemoveColons 
            Caption         =   "Remove colons"
            Height          =   255
            Left            =   720
            TabIndex        =   36
            Top             =   3600
            Width           =   4935
         End
         Begin VB.CheckBox chkRemoveUnderscores 
            Caption         =   "Remove underscores"
            Height          =   255
            Left            =   720
            TabIndex        =   35
            Top             =   3240
            Value           =   1  'Checked
            Width           =   4935
         End
         Begin VB.CheckBox chkRemoveEmptyLines 
            Caption         =   "Remove empty lines"
            Height          =   255
            Left            =   720
            TabIndex        =   34
            Top             =   3960
            Width           =   4935
         End
         Begin VB.CheckBox chkObfuscate 
            Caption         =   "Obfuscation enabled"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   4935
         End
         Begin VB.CheckBox chkReplaceFileNames 
            Caption         =   "Replace file names"
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   840
            Width           =   3135
         End
         Begin VB.CheckBox chkReplaceProjectNames 
            Caption         =   "Replace project names"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox chkReplaceModuleNames 
            Caption         =   "Replace module/form/class names"
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   4935
         End
         Begin VB.CheckBox chkReplacePublicVariableNames 
            Caption         =   "Replace public variable names"
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   1920
            Width           =   4935
         End
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
    
         .Obfuscate = CBool(Me.chkObfuscate.Value)
        
            .ReplaceFileNames = CBool(Me.chkReplaceFileNames.Value)
            .ReplaceProjectNames = CBool(Me.chkReplaceProjectNames.Value)
            .ReplaceModuleNames = CBool(Me.chkReplaceModuleNames.Value)
        
        
        .RemoveCodeLayout = CBool(Me.chkRemoveCodeLayout.Value)
            
            .RemoveComments = CBool(Me.chkRemoveComments.Value)
            .RemoveUnderscores = CBool(Me.chkRemoveUnderscores.Value)
            
            .RemoveColons = CBool(Me.chkRemoveColons.Value)
            .RemoveEmptyLines = CBool(Me.chkRemoveEmptyLines.Value)
            .RemoveSpaces = CBool(Me.chkRemoveSpaces.Value)
    
          
        .EncryptStrings = CBool(Me.chkEncryptStrings.Value)
            
            .EncryptControlCaptions = CBool(Me.chkEncryptControlCaptions.Value)
            .ExportStrings = CBool(Me.chkExportStrings.Value)
            .EnableCrackCheck = CBool(Me.chkCrackCheck.Value)
            
        .EnableDebugMode = CBool(Me.chkDebugMode.Value)
        .AddErrorHandlers = CBool(Me.chkAddErrorHandlers.Value)

    End With

End Function


Function LoadConfig()
On Error Resume Next

    With My.Config
    
        Me.chkObfuscate.Value = Abs(CLng(.Obfuscate))
        
        Me.chkReplaceFileNames.Value = Abs(CLng(.ReplaceFileNames))
        Me.chkReplaceProjectNames.Value = Abs(CLng(.ReplaceProjectNames))
        Me.chkReplaceModuleNames.Value = Abs(CLng(.ReplaceModuleNames))
        
        Me.chkRemoveCodeLayout.Value = Abs(CLng(.RemoveCodeLayout))
        
        Me.chkRemoveComments.Value = Abs(CLng(.RemoveComments))
        Me.chkRemoveUnderscores.Value = Abs(CLng(.RemoveUnderscores))
        Me.chkRemoveColons.Value = Abs(CLng(.RemoveColons))
        Me.chkRemoveEmptyLines.Value = Abs(CLng(.RemoveEmptyLines))
        Me.chkRemoveSpaces.Value = Abs(CLng(.RemoveSpaces))
        
        Me.chkEncryptStrings.Value = Abs(CLng(.EncryptStrings))
        Me.chkEncryptControlCaptions.Value = Abs(CLng(.EncryptControlCaptions))
        Me.chkExportStrings.Value = Abs(CLng(.ExportStrings))
        Me.chkCrackCheck.Value = Abs(CLng(.EnableCrackCheck))
        
        Me.chkDebugMode.Value = Abs(CLng(.EnableDebugMode))
        Me.chkAddErrorHandlers.Value = Abs(CLng(.AddErrorHandlers))
        
    End With
    
End Function
