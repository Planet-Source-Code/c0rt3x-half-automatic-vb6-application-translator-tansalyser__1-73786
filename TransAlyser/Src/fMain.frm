VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fMain 
   Caption         =   "TransAlyser"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTreeNode 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   6120
   End
   Begin VB.TextBox txtEdit 
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrResizeH 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10680
      Top             =   3600
   End
   Begin VB.PictureBox picDivH 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3840
      MousePointer    =   7  'Size N S
      ScaleHeight     =   255
      ScaleWidth      =   6615
      TabIndex        =   7
      Top             =   3720
      Width           =   6615
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   1455
      Left            =   3840
      TabIndex        =   6
      Top             =   4200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"fMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picText 
      Height          =   375
      Left            =   4920
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer tmrResizeW 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   6240
   End
   Begin VB.PictureBox picDivW 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   3360
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4575
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin vbAcceleratorSGrid6.vbalGrid Grid 
      Height          =   1935
      Left            =   3840
      TabIndex        =   3
      Top             =   1440
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3413
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   8130
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   3240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":04C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":119C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":1E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":2B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":2FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":33F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":40CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   953
      ButtonWidth     =   873
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilTreeView 
      Left            =   240
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":49A8
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":4B02
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":4C5C
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":4DB6
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":4F10
            Key             =   "UserControl"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":506A
            Key             =   "PropertyPage"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":51C4
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":575E
            Key             =   "Sub"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":58B8
            Key             =   "ClosedFolder"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fMain.frx":5E52
            Key             =   "OpenFolder"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7858
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ilTreeView"
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
      OLEDropMode     =   1
   End
   Begin vbalIml6.vbalImageList ilGrid 
      Left            =   4200
      Top             =   6720
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   2296
      Images          =   "fMain.frx":63EC
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type



Private Const MyDefaulWestSideWidthInPercent As Double = 25

Private MyWestSideWidth As Long
Private MyNorthEastSideHeight As Long

Private MyWestSideWidthInPercent As Double
Private MyNorthEastSideHeightInPercent As Double


Public WithEvents Project As VBProject
Attribute Project.VB_VarHelpID = -1


Public Loading As Boolean
Public Cancelled As Boolean



Private Sub Form_Load()
    
'    Dim StartTime As Double
'
'    StartTime = Timer

    MyWestSideWidthInPercent = 25
    MyNorthEastSideHeightInPercent = 80
    
    Dim Button As Button
    
    With Toolbar
        
        Set .ImageList = Me.ilToolbar
        
        With .Buttons
            
            .Add , "Open", "Open", , 1
            .Add , "Save", "Save", , 6
            .Add , , , 3
            .Add , "Translate", "Translate", , 5
            '.Add , , , 3
            
            Set Button = .Add(, "Filter", "Filter", , 7)
            Button.Style = tbrCheck
            
            .Add , , , 3
            
            .Add , "Settings", "Settings", , 3
        
        End With
        
    End With
    
    Me.picDivW.Width = 45
    Me.picDivH.Height = 45
        
    MyWestSideWidth = Me.ScaleWidth / 100 * MyDefaulWestSideWidthInPercent
    MyNorthEastSideHeight = (Me.ScaleHeight - Me.Toolbar.Height - Me.picDivH.Height - StatusBar.Height) / 100 * 80
    
    Me.StatusBar.SimpleText = ""
    
    DoEvents
    Form_Resize
    
    Me.Show
    
    'fOpen.Show
    'DoEvents
    
    'fProgress.Show
    'DoEvents
    
    'Me.LoadProject My.Paths.DemoProjectPath
    'mnuCrypt_Click
    
    'My.Project.Compile
    
    'Debug.Print "OK", Format((Timer - StartTime), "fixed")
    'End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Shutdown
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    'Tree.Move 0, Me.Toolbar.Height, Me.ScaleWidth * 0.25, Me.ScaleHeight - Me.Toolbar.Height - Me.StatusBar.Height
    
    'Grid.Move Tree.Width, Me.Toolbar.Height, Me.ScaleWidth * 0.75, Me.ScaleHeight - Me.Toolbar.Height - Me.StatusBar.Height

    With Me.Tree
        
        .Left = 0
        .Top = Me.Toolbar.Height
        
        .Width = (Me.ScaleWidth - Me.picDivW.Width) / 100 * MyWestSideWidthInPercent
        .Height = (Me.ScaleHeight - Toolbar.Height - StatusBar.Height)
        
    End With
    
    With picDivW
        
        .Left = Me.Tree.Width
        .Top = Me.Toolbar.Height
        
        '.Width = Me.picDivW.Width
        .Height = (Me.ScaleHeight - Toolbar.Height - StatusBar.Height)
        
    End With
    
    With Grid
        
        .Left = (Me.Tree.Width + Me.picDivW.Width)
        .Top = Me.Toolbar.Height
        
        .Width = (Me.ScaleWidth - (Me.Tree.Width + Me.picDivW.Width))
        .Height = (Me.ScaleHeight - Me.Toolbar.Height - Me.StatusBar.Height) / 100 * MyNorthEastSideHeightInPercent
        
    End With
    
    With picDivH
        
        .Left = (Me.Tree.Width + Me.picDivW.Width)
        .Top = Me.Grid.Top + Me.Grid.Height
        
        .Width = (Me.ScaleWidth - (Tree.Width + Me.picDivW.Width))
        '.Height = Me.picDivH.Height
        
    End With
    
    With rtfCode
        
        .Left = Me.Tree.Width + Me.picDivW.Width
        .Top = Me.Grid.Top + Me.Grid.Height + Me.picDivH.Height
        
        .Width = (Me.ScaleWidth - (Me.Tree.Width + Me.picDivW.Width))
        .Height = (Me.ScaleHeight - Me.Toolbar.Height - Me.Grid.Height - Me.picDivH.Height - Me.StatusBar.Height)
        
    End With
    
    
        
    DoEvents
        
    

End Sub


Function LoadProject(Path As String)
On Error GoTo Error
    
    
    Dim RootNode As Node
    Dim ModRootNode As Node
    Dim ChildNode As Node
    Dim i As Long
        
    fProgress.Show
        
    Set Project = My.Project
    
    Project.LoadProjectFile Path
    
    Caption = "TransAlyser - " & Path
    
    With Tree
        
        .Nodes.Clear
        Set RootNode = .Nodes.Add(, , "Project", Project.Name, "Project", "Project")
        RootNode.Tag = Project.FileName
        
        RootNode.Checked = True
        
        If Project.ModuleCount > 0 Then
            If Project.ModuleTypeCount("Module") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "Modules", "Modules", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "Module" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "Module", "Module")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
        End If
            
            If Project.ModuleTypeCount("Class") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "Classes", "Classes", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "Class" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "Class", "Class")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
            End If
            
            If Project.ModuleTypeCount("Form") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "Forms", "Forms", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "Form" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "Form", "Form")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
            End If
            
            If Project.ModuleTypeCount("UserControl") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "UserControls", "User controls", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "UserControl" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "UserControl", "UserControl")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
            End If
            
            If Project.ModuleTypeCount("PropertyPage") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "PropertyPages", "Property pages", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "PropertyPage" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "PropertyPage", "PropertyPage")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
            End If
            
            If Project.ModuleTypeCount("Designer") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "Designer", "Designer", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "Designer" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "ClosedFolder", "OpenFolder")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
            End If
            
            If Project.ModuleTypeCount("ResFile32") > 0 Then
                Set ModRootNode = .Nodes.Add(RootNode, tvwChild, "ResFile", "ResFile", "ClosedFolder", "OpenFolder")
                ModRootNode.Checked = True
                For i = 1 To Project.ModuleCount
                    If Project.Modules(i).DataType = "ResFile32" Then
                        Set ChildNode = .Nodes.Add(ModRootNode, tvwChild, Project.Modules(i).Name, Project.Modules(i).Name, "ClosedFolder", "OpenFolder")
                        ChildNode.EnsureVisible
                        ChildNode.Checked = True
                    End If
                Next i
            End If
            
        End If
        
       
        
    End With
    
    fProgress.Hide
    
    LoadProjectData
    
    Set Tree.SelectedItem = Tree.Nodes(1)
    Tree_NodeClick Tree.SelectedItem
    
Exit Function
Error:
    
    Debug.Print Err.Description
    
End Function




Private Sub Grid_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, NewValue As Variant, bStayInEditMode As Boolean)
    
    
    Dim StringID As Long
    
    
    StringID = CLng(Grid.CellText(lRow, 1))
    
    
    If (txtEdit.Text = "") Then
        
        bStayInEditMode = True
    
    Else
      
        Grid.CellText(Grid.EditRow, Grid.EditCol) = txtEdit.Text
        
        My.Project.Strings(StringID).Translation = Grid.CellText(lRow, 4)
        My.Project.Strings(StringID).Checked = (Grid.CellText(lRow, 4) <> "")
        
'        If Tree.SelectedItem.Key = "Project" Then
'
'
'            My.Project.Strings(CLng(Grid.CellText(lRow, 1))).Translation = Grid.CellText(lRow, 4)
'            My.Project.Strings(CLng(Grid.CellText(lRow, 1))).Checked = (Grid.CellText(lRow, 4) <> "")
'
'        Else
'
'            My.Project.Modules(My.Project.GetModuleIndex(Tree.SelectedItem.Text)).Strings(CLng(Grid.CellText(lRow, 1))).Translation = Grid.CellText(lRow, 4)
'            My.Project.Modules(My.Project.GetModuleIndex(Tree.SelectedItem.Text)).Strings(CLng(Grid.CellText(lRow, 1))).Checked = (Grid.CellText(lRow, 4) <> "")
'
'        End If
        
        Grid.CellIcon(Grid.SelectedRow, 1) = IIf((Grid.CellText(lRow, 4) <> ""), 1, 0)
        
    End If
    
End Sub

Private Sub Grid_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    Dim sText As String
   
   
   If Grid.SelectedCol <> 4 Then
      bCancel = True
      Exit Sub
   End If
   
   Grid.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
   
   If Not IsMissing(Grid.CellText(lRow, lCol)) Then
      sText = Grid.CellFormattedText(lRow, lCol)
   Else
      sText = ""
   End If
   
   If Not (iKeyAscii = 0) Then
      sText = Chr$(iKeyAscii) & sText
      txtEdit.Text = sText
      txtEdit.SelStart = 1
      txtEdit.SelLength = Len(sText)
   Else
      txtEdit.Text = sText
      txtEdit.SelStart = 0
      txtEdit.SelLength = Len(sText)
   End If
   

   Set txtEdit.Font = Grid.CellFont(lRow, lCol)
   If Grid.CellBackColor(lRow, lCol) = -1 Then
      txtEdit.BackColor = Grid.BackColor
   Else
      txtEdit.BackColor = Grid.CellBackColor(lRow, lCol)
   End If
   
   
   txtEdit.Move lLeft + Grid.Left, lTop + Grid.Top + Screen.TwipsPerPixelY, lWidth, lHeight
   txtEdit.Visible = True
   txtEdit.ZOrder
   txtEdit.SetFocus
   
    
End Sub



Private Sub tmrTreeNode_Timer()
    
    
    Me.tmrTreeNode.Enabled = False
    
    Tree_NodeClick Me.Tree.Nodes(CLng(Me.tmrTreeNode.Tag))
    
    
    
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If (KeyCode = vbKeyReturn) Then
     
      Grid.EndEdit
      txtEdit.Visible = False
      
   ElseIf (KeyCode = vbKeyEscape) Then
     
      Grid.CancelEdit
      txtEdit.Visible = False
      
   ElseIf (Grid.SingleClickEdit) Then
      Select Case KeyCode
      
      End Select
   End If
   
End Sub

Private Sub picDivH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.tmrResizeH.Enabled = True
End Sub

Private Sub picDivH_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.tmrResizeH.Enabled = False
End Sub

Private Sub Project_FileProgress(Progress As Double)
    fProgress.pbFile.Value = Progress
    DoEvents
End Sub

Private Sub Project_ProjectProgress(Progress As Double)
    fProgress.pbProject.Value = Progress
    DoEvents
End Sub





Private Sub tmrResizeH_Timer()
    
    Dim HeightInPerCent As Double
    
    If MouseButtonPressed(1) = False Then
        Me.tmrResizeH.Enabled = False
    End If
    
    HeightInPerCent = 100 / Me.ScaleHeight * MyMousePos.y
    MyNorthEastSideHeightInPercent = HeightInPerCent
    
    If (HeightInPerCent < 20) Or (HeightInPerCent > 90) Then
        Exit Sub
    End If
    
    Me.NorthEastSideHeight = MyMousePos.y - Me.Toolbar.Height
    
    'Debug.Print HeightInPerCent
    
End Sub

Private Sub tmrResizeW_Timer()
    
     Dim WidthInPerCent As Double
    
    If MouseButtonPressed(1) = False Then
        Me.tmrResizeW.Enabled = False
    End If
    
    WidthInPerCent = 100 / Me.ScaleWidth * MyMousePos.x
    
    MyWestSideWidthInPercent = WidthInPerCent
    
    If (WidthInPerCent < 10) Or (WidthInPerCent > 90) Then
        Exit Sub
    End If
    
    Me.WestSideWidth = MyMousePos.x
    
    'Debug.Print WidthInPerCent
    
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        
        Case "Open"
            
            fOpen.Show 1, Me
        
        Case "Save"
            
               SaveProjectData
        
'        Case "Objects"
'
'            fObjects.Show 1, Me
        
        Case "Translate"
        
            fProgress.Show
            DoEvents
                
            My.SaveProject Project.BaseDirPath & "Translated\"
                
            fProgress.Hide
            DoEvents
        
        Case "Filter"
        
            Tree_NodeClick Tree.SelectedItem
        
        Case "Settings"
            
            fOptions.LoadConfig
            fOptions.Show 1, Me
            
        Case Else
        
            Debug.Print Button.Key
        
    End Select
    
End Sub


Private Sub picDivW_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.tmrResizeW.Enabled = True
End Sub

Private Sub picDivW_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.tmrResizeW.Enabled = False
End Sub



Public Property Get WestSideWidth() As Long
    WestSideWidth = MyWestSideWidth
End Property

Public Property Let WestSideWidth(NewVal As Long)
    MyWestSideWidth = NewVal
    Form_Resize
End Property

Public Property Get NorthEastSideHeight() As Long
    NorthEastSideHeight = MyNorthEastSideHeight
End Property

Public Property Let NorthEastSideHeight(NewVal As Long)
    MyNorthEastSideHeight = NewVal
    Form_Resize
End Property

Private Property Get MyMousePos() As POINTAPI
    
    Dim ScreenMousePos As POINTAPI
    Dim FormMousePos As POINTAPI
    
    GetCursorPos ScreenMousePos
    
    FormMousePos.x = (ScreenMousePos.x * Screen.TwipsPerPixelX) - Me.Left - (Me.Width - Me.ScaleWidth)
    FormMousePos.y = (ScreenMousePos.y * Screen.TwipsPerPixelY) - Me.Top - (Me.Height - Me.ScaleHeight)
    
    MyMousePos = FormMousePos
    
End Property


Function LoadStrings()
            

    Dim i As Long
    Dim TxtWidth As Long
    
    Dim bAddRow As Boolean
    
    
    Loading = True
    
    
    With Grid
        
        .Clear True
        
        .AddColumn "ID", "ID", ecgHdrTextALignCentre
        .AddColumn "Sub", "Sub"
        .AddColumn "String", "String"
        .AddColumn "Translation", "Translation"
        
        .ImageList = Me.ilGrid
        
        .GridLines = True
        
    End With
    
    For i = 1 To My.Project.StringCount
        
        
        If Cancelled Then
            
            Loading = False
            
            Exit Function
            
        End If
        
        
        Me.StatusBar.SimpleText = "Loading String " & Format(i, String(Len(CStr(My.Project.StringCount)), "0")) & " of " & My.Project.StringCount
        DoEvents
        
        With Grid
            
            If Toolbar.Buttons("Filter").Value = tbrPressed Then
                If My.Project.Strings(i).Checked Then
                    bAddRow = True
                Else
                    bAddRow = False
                End If
            Else
                bAddRow = True
            End If
            
            If bAddRow Then
            
                .AddRow
                
                .CellIcon(.Rows, 1) = Abs(CLng(My.Project.Strings(i).Checked))
                
                .CellText(.Rows, 1) = Format(i, String(Len(CStr(My.Project.StringCount)), "0"))
                
                .CellText(.Rows, 2) = My.Project.Strings(i).ParentModule.Name & "." & My.Project.Strings(i).ParentSub.Name
                
                .CellText(.Rows, 3) = My.Project.Strings(i).Value
                
                .CellText(.Rows, 4) = My.Project.Strings(i).Translation
                
                TxtWidth = GetTxtWidth(.CellText(.Rows, 2))
                If .ColumnWidth(2) < TxtWidth Then
                    .ColumnWidth(2) = TxtWidth
                End If
    
                TxtWidth = GetTxtWidth(.CellText(.Rows, 3))
                If .ColumnWidth(3) < TxtWidth Then
                    .ColumnWidth(3) = TxtWidth
                    .ColumnWidth(4) = TxtWidth
                End If
                
            End If
            
        End With
        
    Next i
    
    Me.StatusBar.SimpleText = ""
    
    
    Loading = False
    
End Function

Function LoadModuleStrings(ModuleName As String)
    
    Dim n As Long
    Dim i As Long
    Dim TxtWidth As Long
    
    Dim bAddRow As Boolean
    
    
    Loading = True
    
    
    With Grid
        
        .Clear True
        
        .AddColumn "ID", "ID", ecgHdrTextALignCentre
        .AddColumn "Sub", "Sub"
        .AddColumn "String", "String"
        .AddColumn "Translation", "Translation"
        
        .ImageList = Me.ilGrid
        
        .GridLines = True
        
    End With
    
    For n = 1 To My.Project.ModuleCount
        
        
        If My.Project.Modules(n).Name = ModuleName Then
        
            For i = 1 To My.Project.Modules(n).StringCount
                
                
                If Cancelled Then
            
                    Loading = False
            
                    Exit Function
            
                End If
                
                
                Me.StatusBar.SimpleText = "Loading String " & Format(i, String(Len(CStr(My.Project.Modules(n).StringCount)), "0")) & " of " & My.Project.Modules(n).StringCount
                DoEvents
            
            
                With Grid
                
                     If Toolbar.Buttons("Filter").Value = tbrPressed Then
                        If My.Project.Modules(n).Strings(i).Checked Then
                            bAddRow = True
                        Else
                            bAddRow = False
                        End If
                    Else
                        bAddRow = True
                    End If
                    
                    If bAddRow Then
                    
                        .AddRow
                        
                        .CellIcon(.Rows, 1) = Abs(CLng(My.Project.Modules(n).Strings(i).Checked))
                        
                        
                        
                         .CellText(.Rows, 1) = Format(My.Project.Modules(n).Strings(i).GlobalID, String(Len(CStr(My.Project.StringCount)), "0"))
                        
                        '.CellText(.Rows, 1) = Format(My.Project.Modules(n).Strings(i).Index, String(Len(CStr(My.Project.Modules(n).StringCount)), "0"))
                        
                        .CellText(.Rows, 2) = My.Project.Modules(n).Strings(i).ParentSub.Name
                        
                        .CellText(.Rows, 3) = My.Project.Modules(n).Strings(i).Value
                        
                        .CellText(.Rows, 4) = My.Project.Modules(n).Strings(i).Translation
                        
                        TxtWidth = GetTxtWidth(.CellText(.Rows, 2))
                        If .ColumnWidth(2) < TxtWidth Then
                            .ColumnWidth(2) = TxtWidth
                        End If
                        
                        TxtWidth = GetTxtWidth(.CellText(.Rows, 3))
                        If .ColumnWidth(3) < TxtWidth Then
                            .ColumnWidth(3) = TxtWidth
                            .ColumnWidth(4) = TxtWidth
                        End If
                    
                    End If
                        
                End With
            
            
            
            Next i
        
        End If
        
    Next n
    
    Me.StatusBar.SimpleText = ""
    
    
    Loading = False
    
    
End Function


Private Function GetTxtWidth(Text As String) As Long
    
    Me.picText.Font.Name = Me.Grid.Font.Name
    Me.picText.Font.Size = Me.Grid.Font.Size
    GetTxtWidth = (Me.picText.TextWidth(Text) / Screen.TwipsPerPixelX) + 8

    
End Function

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
    
    
    
    
    If Loading Then

        Cancelled = True
        
        Me.tmrTreeNode.Tag = Node.Index
        Me.tmrTreeNode.Enabled = True
        
        Exit Sub
        
'        Do
'
'            DoEvents
'
'        Loop While Loading
'
'        Cancelled = False

    End If
    
    Cancelled = False
    
    
    txtEdit.Visible = False
    
    If Node Is Nothing Then
        Exit Sub
    End If
    
    Select Case Node.Key
        
        Case "Project"
            
            'If Me.Toolbar.Buttons("Code").Value = tbrPressed Then
                LoadStrings
            'Else
            '    LoadControlStrings
            'End If
            
        Case Else
            
            'If Me.Toolbar.Buttons("Code").Value = tbrPressed Then
                LoadModuleStrings Node.Key
            'Else
            '    LoadControlStrings Tree.SelectedItem.Key
            'End If
            
            
    End Select
    
End Sub


Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    
    If KeyCode = vbKeySpace Then
        
        With Grid
            
            If .CellIcon(.SelectedRow, 1) = 0 Then
                'ClearSelection
                .CellIcon(.SelectedRow, 1) = 1
            Else
                .CellIcon(.SelectedRow, 1) = 0
            End If
            
            If Tree.SelectedItem.Key = "Project" Then
                My.Project.Strings(CLng(.CellText(.SelectedRow, 1))).Checked = CBool(.CellIcon(.SelectedRow, 1))
            Else
                My.Project.Modules(My.Project.GetModuleIndex(Tree.SelectedItem.Text)).Strings(.SelectedRow).Checked = CBool(.CellIcon(.SelectedRow, 1))
            End If
            
            
        End With
        
    End If
    
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, bDoDefault As Boolean)
    
    
    Dim CurRow As Long
    Dim CurCol As Long
    
    Dim GlobalStringID As Long
    Dim ModuleStringID As Long
    
    
    If Button = 1 Then
    
         With Grid
            
            .CellFromPoint x \ 15, y \ 15, CurRow, CurCol
            
            If CurCol = 1 Then
            
                If .CellIcon(CurRow, 1) = 0 Then
                    'ClearSelection
                    .CellIcon(CurRow, 1) = 1
                Else
                    .CellIcon(CurRow, 1) = 0
                End If
                
                
                
'                If Tree.SelectedItem.Key = "Project" Then
'
'
'
'                    My.Project.Strings(CLng(.CellText(CurRow, 1))).Checked = CBool(.CellIcon(CurRow, 1))
'
'                Else
                    
                    
                    GlobalStringID = CLng(Me.Grid.CellText(CurRow, 1))
                    
                    'ModuleStringID = My.Project.Strings(GlobalStringID).Index
                    
                    
                    My.Project.Strings(GlobalStringID).Checked = CBool(.CellIcon(CurRow, 1))
                
                    'My.Project.Modules(My.Project.GetModuleIndex(Tree.SelectedItem.Text)).Strings(ModuleStringID).Checked = CBool(.CellIcon(CurRow, 1))
                    
                    
                'End If
            
            End If
        
        End With
    
    End If
    
    
End Sub


Private Sub Grid_ColumnClick(ByVal lCol As Long)
On Error Resume Next

    With Grid
        
        If .SortObject.SortOrder(1) = CCLOrderAscending Then
            .SortObject.SortOrder(1) = CCLOrderDescending
        Else
            .SortObject.SortOrder(1) = CCLOrderAscending
        End If
        
        .SortObject.SortType(1) = CCLSortString
        
        .SortObject.SortColumn(1) = lCol
        '.SortObject.GroupBy(0) = 2
        
        .Sort
        
    End With
    
End Sub


Private Sub Grid_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
On Error Resume Next

    Dim StringID As Long
    Dim LineID As Long
    Dim ModuleID As Long
    
    If Me.Toolbar.Buttons("Code").Value = tbrPressed Then
    
        
        'If Me.Tree.SelectedItem.Key = "Project" Then
            
            StringID = CLng(Me.Grid.CellText(lRow, 1))
            LineID = My.Project.Strings(StringID).LineID
            ModuleID = My.Project.Strings(StringID).ParentModule.ID
            
            'Me.StatusBar.SimpleText = My.Project.Modules(ModuleID).Lines(LineID)
            
            LoadLine ModuleID, LineID
            
            
            
'        Else
'
'            ModuleID = My.Project.GetModuleIndex(Me.Tree.SelectedItem.Key)
'            If ModuleID <> 0 Then
'
'                StringID = CLng(Me.Grid.CellText(lRow, 1))
'                StringID = My.Project.Strings(StringID).Index
'
'                LineID = My.Project.Modules(ModuleID).Strings(StringID).LineID
'
'                'Me.StatusBar.SimpleText = My.Project.Modules(ModuleID).Lines(LineID)
'
'                LoadLine ModuleID, LineID
'
'    '            Me.rtfCode.Text = My.Project.Modules(ModuleID).Lines(LineID)
'    '
'    '            FormatVBCode Me.rtfCode
'
'            End If
'
'        End If
        
    ElseIf Me.Toolbar.Buttons("Controls").Value = tbrPressed Then
    
        Me.rtfCode.Text = ""
        
    End If
    
End Sub

Function LoadLine(ModuleID As Long, LineID As Long)
    
    Me.rtfCode.Text = LTrim(My.Project.Modules(ModuleID).Lines(LineID))
        
    FormatVBCode Me.rtfCode
    
End Function


'Function LoadControlStrings(Optional ModuleName As String = "")
'
'    Dim x As Long
'    Dim y As Long
'    Dim z As Long
'
'    Dim StringCount As Long
'    Dim c As Long
'
'    Dim s As String
'
'    Dim TxtWidth As Long
'
'    With Grid
'
'        .Clear True
'
'        .AddColumn "ID", "ID", ecgHdrTextALignCentre
'        .AddColumn "Property", "Property"
'        .AddColumn "String", "String"
'
'        .ImageList = Me.ilGrid
'
'        .GridLines = True
'
'    End With
'
'
'    For x = 1 To My.Project.ModuleCount
'
'        If (My.Project.Modules(x).Name = ModuleName) Or ModuleName = "" Then
'
'            For y = 1 To My.Project.Modules(x).ControlCount
'
'                For z = 1 To My.Project.Modules(x).Controls(y).PropertyCount
'
'                    Select Case My.Project.Modules(x).Controls(y).Properties(z).Name
'
'                        Case "Caption", "Text", "ToolTipText"
'
'
'                                StringCount = StringCount + 1
'
'
'                        Case Else
'
'
'                    End Select
'
'
'
'                Next z
'
'
'
'            Next y
'
'        End If
'
'    Next x
'
'
'
'    For x = 1 To My.Project.ModuleCount
'
'        If (My.Project.Modules(x).Name = ModuleName) Or ModuleName = "" Then
'
'            For y = 1 To My.Project.Modules(x).ControlCount
'
'                For z = 1 To My.Project.Modules(x).Controls(y).PropertyCount
'
'                    Select Case My.Project.Modules(x).Controls(y).Properties(z).Name
'
'                        Case "Caption", "Text", "ToolTipText"
'
'                            With Grid
'
'                                c = c + 1
'
'                                .AddRow
'
'                                .CellIcon(.Rows, 1) = 0
'
'                                .CellText(.Rows, 1) = Format(c, String(Len(CStr(StringCount)), "0"))
'
'
'                                s = My.Project.Modules(x).Name
'                                s = s & "." & My.Project.Modules(x).Controls(y).Name
'
'                                If My.Project.Modules(x).Controls(y).GetPropertyIndex("Index") <> 0 Then
'
'                                    s = s & "(" & My.Project.Modules(x).Controls(y).Properties(My.Project.Modules(x).Controls(y).GetPropertyIndex("Index")).Value & ")"
'
'                                End If
'
'
'                                s = s & "." & My.Project.Modules(x).Controls(y).Properties(z).Name
'
'                                .CellText(.Rows, 2) = s
'
'                               ' Set Item = Me.ListControls.ListItems.Add(, , s)
'
'                                s = My.Project.Modules(x).Controls(y).Properties(z).Value
'
'                                s = Mid(s, 2, Len(s) - 2)
'
'                                .CellText(.Rows, 3) = s
'
'
'
'                                TxtWidth = GetTxtWidth(.CellText(.Rows, 2))
'                                If .ColumnWidth(2) < TxtWidth Then
'                                    .ColumnWidth(2) = TxtWidth
'                                End If
'
'                                TxtWidth = GetTxtWidth(.CellText(.Rows, 3))
'                                If .ColumnWidth(3) < TxtWidth Then
'                                    .ColumnWidth(3) = TxtWidth
'                                End If
'
'
'                            End With
'
'                        Case Else
'
'
'                    End Select
'
'
'
'                Next z
'
'
'
'            Next y
'
'        End If
'
'    Next x
'
'End Function


Function SaveProjectData()
    
    Dim i As Long
    
    Dim FilePath As String
    Dim FileData As String
    
    SetIniValue FileData, "Project", "AppName", My.Project.Name
    SetIniValue FileData, "Project", "SelectedModule", Tree.SelectedItem.Text
    
    For i = 1 To My.Project.StringCount
        
        SetIniValue FileData, "Strings", Base64Encode(My.Project.Strings(i).Value), Abs(CLng(My.Project.Strings(i).Checked))
        SetIniValue FileData, "Translations", Base64Encode(My.Project.Strings(i).Value), Base64Encode(My.Project.Strings(i).Translation)
        
    Next i
    
    FilePath = My.Path & My.Project.Name & ".ini"
    
    If FileExist(FilePath) Then
        DeleteFile FilePath
    End If
    
    WriteFile FilePath, FileData
    
End Function

Function LoadProjectData()
        
    Dim i As Long
    
    Dim FilePath As String
    Dim FileData As String
    Dim Value As String
    
    
    FilePath = My.Path & My.Project.Name & ".ini"
    
    If Not FileExist(FilePath) Then
        Exit Function
    End If
    
    FileData = ReadFile(FilePath)
    
    For i = 1 To My.Project.StringCount
        
        Me.StatusBar.SimpleText = "Loading String Translation " & CStr(i) & " of " & CStr(My.Project.StringCount)
        
        DoEvents
        
        
        Value = GetIniValue(FileData, "Strings", Base64Encode(My.Project.Strings(i).Value))
        If Value <> "" Then
            My.Project.Strings(i).Checked = CBool(Value)
        End If
        
        Value = GetIniValue(FileData, "Translations", Base64Encode(My.Project.Strings(i).Value))
        If Value <> "" Then
            My.Project.Strings(i).Translation = Base64Decode(Value)
        End If
        
    Next i
    
    Value = GetIniValue(FileData, "Project", "SelectedModule")
    
    If Value <> "" Then
        Set Tree.SelectedItem = Tree.Nodes(Value)
    End If
    
    
End Function
