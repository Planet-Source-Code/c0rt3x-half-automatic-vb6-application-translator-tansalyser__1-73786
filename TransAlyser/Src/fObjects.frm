VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form fObjects 
   Caption         =   "Object Browser"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "fObjects.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   6390
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.ListBox lstTypeInfos 
      Height          =   840
      ItemData        =   "fObjects.frx":000C
      Left            =   3840
      List            =   "fObjects.frx":000E
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstMembers 
      Height          =   840
      ItemData        =   "fObjects.frx":0010
      Left            =   3840
      List            =   "fObjects.frx":0012
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4683
      _Version        =   393217
      Indentation     =   500
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imlTool"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlTool 
      Left            =   3840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":0014
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":0366
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":06B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":0A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":0D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":10AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":1400
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":1752
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":1AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":1DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":250E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fObjects.frx":2860
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Object browser control 1.1
'
' Copyright 2005, E.V.I.C.T. B.V.
' Website: http:\\www.evict.nl
' Support: mailto:evict@vermeer.nl
'
'Purpose:
' This control lets you have a look at the structure of any OLE/COM object
' Based on the article : <a href=http://msdn.microsoft.com/msdnmag/issues/1200/TypeLib/default.aspx>MSDN Magazine, December 2000 by Jason Fisher</a>
'
'License:
' GPL - The GNU General Public License
' Permits anyone the right to use and modify the software without limitations
' as long as proper credits are given and the original and modified source code
' are included. Requires that the final product, software derivate from the
' original source or any software utilizing a GPL component, such as this,
' is also licensed under the GPL license.
' For more information see http://www.gnu.org/licenses/gpl.txt
'
'License adition:
' You are permitted to use the software in a non-commercial context free of
' charge as long as proper credits are given and the original unmodified source
' code is included.
' For more information see http://www.evict.nl/licenses.html
'
'License exeption:
' If you would like to obtain a commercial license then please contact E.V.I.C.T. B.V.
' For more information see http://www.evict.nl/licenses.html
'
'Terms:
' This software is provided "as is", without warranty of any kind, express or
' implied, including  but not limited to the warranties of merchantability,
' fitness for a particular purpose and noninfringement. In no event shall the
' authors or copyright holders be liable for any claim, damages or other
' liability, whether in an action of contract, tort or otherwise, arising
' from, out of or in connection with the software or the use or other
' dealings in the software.
'
'History:
' 2002 : Created and added to the sharware library siteskinner
' jan 2005 : Changed the licensing from shareware to opensource
' feb 2005 : Added the SetSockLinger and getascip for the Improved connection method


Dim tlitypelibinfo As TLI.TypeLibInfo   ' A reference to the type library information that we are processing
Dim strReferedObjects()
' Property variables
Dim psName As String
Dim psFile As String
Dim psVersion As String
Dim psHelpString As String
Dim psHelpFile As String
Dim psSystem As String
Dim psGuid As String
Dim plSettings As Long

'Purpose:
' These settings can be used to customize the content of the ObjectBrowser
Private Enum ObjectBrowserSettings
    OB_GroupMemberType = 1           ' This will add an extra level to the tree. All properties will be in a property node, all methods will be in a method node and all events will be in an event node.
    OB_EnumInRoot = 2                ' If you use ExtendInnerObjects, then Enumerations will be shown where they are used. Then you will probably not want to see enumeration objects in the root
    OB_ExtendInnerObjects = 4        ' If a members returns an object/enum that is in this library, then that object can be put under this member.
    OB_ExtendParent = 8              ' When Objects are extended then you might want to disable extending the parrent since that is already done a level up in the tree.
    OB_RemoveReferedObjects = 16     ' When inner objects are extended then you can remove the objects that are an extended object since they are probably only accessed through that other object. This will make the tree smaller and it will only show you the main objects that you will use.
End Enum

'Purpose:
' The event that is called when a node is selected
Event NodeSelect(strNode As String, strCall As String, strDescription As String, strHelp As String, strHelpFile As String, Node As MSComctlLib.Node, strHelpPage As String)

'Purpose:
' The event that is called when a node is double clicked
Event NodeExecute(strNode As String, strCall As String, strDescription As String, strHelp As String, strHelpFile As String, Node As MSComctlLib.Node, strHelpPage As String)

'Purpose:
' To show us where we are in case we are processing a big library
Event Progress(strCurrentNode As String)

'Purpose:
' The Name for the selected object.

Private Property Get settings() As ObjectBrowserSettings

    settings = plSettings

End Property

Private Property Let settings(lngSetting As ObjectBrowserSettings)

    plSettings = lngSetting

End Property

'Purpose:
' The Name for the selected object.

'Public Property Get Name() As String
'
'    Name = psName
'
'End Property

'Purpose:
' The File for the selected object.

Public Property Get File() As String

    File = psFile

End Property

'Purpose:
' The HelpString for the selected object.

Public Property Get Version() As String

    Version = psVersion

End Property

'Purpose:
' The HelpString for the selected object.

Public Property Get HelpString() As String

    HelpString = psHelpString

End Property

'Purpose:
' The HelpFile for the selected object.

Public Property Get HelpFile() As String

    HelpFile = psHelpFile

End Property

'Purpose:
' The system where the selected object is for (Probably Win32 ;-)

Public Property Get System() As String

    System = psSystem

End Property

'Purpose:
' The Guid of the selected object

Public Property Get Guid() As String

    Guid = psGuid

End Property

'Purpose:
' The Treeview

Public Property Get Nodes() As MSComctlLib.Nodes

    Set Nodes = Tree.Nodes

End Property

'Purpose:
' Clear the treeview

Public Sub Clear()

    Tree.Nodes.Clear

End Sub

'Purpose:
' The tlb for this file will be added to the tree.

Public Function AddFromFile(strFileName As String) As Long

    On Error Resume Next
        Set tlitypelibinfo = New TLI.TypeLibInfo
    
        tlitypelibinfo.ContainingFile = strFileName
        AddFromFile = Err.Number
        If Err.Number = 0 Then
            On Error GoTo 0
            processTypeLib
        End If

End Function

'Purpose:
' You can add the members of an object in a different library to any node. This is used where the return value of one property/method is an object

Public Function AddForeignMembersToNode(Node As MSComctlLib.Node, strFileName As String, strObject As String) As Long

On Error Resume Next

    Dim lngLoop1 As Long
    Dim tliTypeinfo As TypeInfo
    Dim nodx As Node

        tlitypelibinfo.ContainingFile = strFileName
        AddForeignMembersToNode = Err.Number
        If Err.Number <> 0 Then Exit Function

        processTypeLibInfo

        ' Get all the objects and find the one
        tlitypelibinfo.GetTypesDirect lstTypeInfos.hwnd, tliWtListBox, tliStAll
        For lngLoop1 = 0 To lstTypeInfos.ListCount - 1
            If lstTypeInfos.List(lngLoop1) = strObject Then
                Set tliTypeinfo = tlitypelibinfo.GetTypeInfo(lstTypeInfos.List(lngLoop1))
                ' Now add all the members to this object
                If AddMembersToObject(lngLoop1, Replace(Node.FullPath, "\", ".")) Then
                    nodx.Tag = "Public WithEvents " & lstTypeInfos.List(lngLoop1) & " As " & tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1) & vbCrLf & tliTypeinfo.HelpString
                Else
                    nodx.Tag = "Public " & lstTypeInfos.List(lngLoop1) & " As New " & tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1) & vbCrLf & tliTypeinfo.HelpString
                End If
                Exit For
            End If
        Next lngLoop1
        nodx.EnsureVisible

End Function

Public Sub LoadDescriptions()

End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
    End If
    Me.Hide
End Sub

Private Sub Tree_Click()
On Error Resume Next
    
    Me.StatusBar.SimpleText = Tree.SelectedItem.Tag

End Sub

Private Sub Tree_DblClick()
On Error Resume Next

    MsgBox Tree.SelectedItem.Tag, vbInformation

End Sub



'Purpose:
' Get the typelib info and set the property variables

Private Sub processTypeLibInfo()
On Error Resume Next

    psName = tlitypelibinfo.Name
    psFile = LCase(tlitypelibinfo.ContainingFile)
    psVersion = tlitypelibinfo.MajorVersion & "." & tlitypelibinfo.MinorVersion
    psHelpString = tlitypelibinfo.HelpString
    psHelpFile = LCase(tlitypelibinfo.HelpFile)
    
    Select Case tlitypelibinfo.SysKind
    Case SYS_MAC

        psSystem = "Macintosh"
    Case SYS_WIN16

        psSystem = "Win16"
    Case SYS_WIN32

        psSystem = "Win32"
    End Select
    psGuid = tlitypelibinfo.Guid

    'Clear lists
    lstTypeInfos.Clear
    lstMembers.Clear

End Sub

'Purpose:
' This will add an entire type library to the tree

Private Sub processTypeLib()

On Error Resume Next
    
    Dim lngLoop1 As Long
    Dim lngLoop2 As Long
    Dim tliTypeinfo As TypeInfo
    Dim nodx As Node
    Dim nodr As Node
    Dim blnAdded As Boolean
    Dim strDescription As String
    Dim strTemp As String

        processTypeLibInfo
        ReDim strReferedObjects(0)

        tlitypelibinfo.GetTypesDirect lstTypeInfos.hwnd, tliWtListBox, tliStAll

        ' Adding the root to the tree
        Set nodr = Tree.Nodes.Add(, , psName, psName, 10, 10)
        nodr.Tag = tlitypelibinfo.HelpString & " object library  " & vbCrLf & "Reference=*\G" & psGuid & "#" & psVersion & "#" & tlitypelibinfo.ContainingFile & "#" & tlitypelibinfo.HelpString & vbCrLf & psHelpFile
        nodr.Sorted = True

        ' Adding the object to the tree
        For lngLoop1 = 0 To lstTypeInfos.ListCount - 1
            
            ' Depending on the type of object we will add it with it's specific icon.
            
            Set tliTypeinfo = tlitypelibinfo.GetTypeInfo(lstTypeInfos.List(lngLoop1))
            blnAdded = False
            
            Select Case tliTypeinfo.TypeKindString
            
                Case "coclass", "interface", "dispinterface"
                    
                    Set nodx = Tree.Nodes.Add(tlitypelibinfo.Name, tvwChild, tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1), lstTypeInfos.List(lngLoop1), 1, 1)
                    blnAdded = True
                    strDescription = lstTypeInfos.List(lngLoop1) & " is an object in the " & tlitypelibinfo.Name & " library."
                
                Case "module"
                    
                    Set nodx = Tree.Nodes.Add(tlitypelibinfo.Name, tvwChild, tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1), lstTypeInfos.List(lngLoop1), 2, 2)
                    blnAdded = True
                    strDescription = lstTypeInfos.List(lngLoop1) & " is an module in the " & tlitypelibinfo.Name & " library."
                
                Case "enum"
                    
                    If plSettings And OB_EnumInRoot Then
                        Set nodx = Tree.Nodes.Add(tlitypelibinfo.Name, tvwChild, tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1), lstTypeInfos.List(lngLoop1), 4, 4)
                        blnAdded = True
                        strDescription = lstTypeInfos.List(lngLoop1) & " is an enumeration in the " & tlitypelibinfo.Name & " library."
                    End If
                    
                Case "record"
                    
                    Set nodx = Tree.Nodes.Add(tlitypelibinfo.Name, tvwChild, tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1), lstTypeInfos.List(lngLoop1), 3, 3)
                    blnAdded = True
                    strDescription = lstTypeInfos.List(lngLoop1) & " is a record structure in the " & tlitypelibinfo.Name & " library."
            
                Case "alias", "union"
                    
                    'Debug.Print "tliTypeinfo.TypeKindString = " & tliTypeinfo.TypeKindString & ", for " & lstTypeInfos.List(lngLoop1)
                
                Case Else
                    
                    Debug.Print "tliTypeinfo.TypeKindString = " & tliTypeinfo.TypeKindString & ", for " & lstTypeInfos.List(lngLoop1)
            
            End Select
            
            ' If we did add the object, then add all the members to this object
            If blnAdded Then
                nodx.Sorted = True
                If AddMembersToObject(lngLoop1, tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1)) Then
                    strTemp = "Public WithEvents " & lstTypeInfos.List(lngLoop1) & " As " & tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1)
                    nodx.Tag = strDescription & vbCrLf & strTemp & vbCrLf & tliTypeinfo.HelpString
                Else
                    strTemp = "Public " & lstTypeInfos.List(lngLoop1) & " As New " & tlitypelibinfo.Name & "." & lstTypeInfos.List(lngLoop1)
                    nodx.Tag = strDescription & vbCrLf & strTemp & vbCrLf & tliTypeinfo.HelpString
                End If
            End If
        Next lngLoop1
        nodx.EnsureVisible

        ' Now remove the refered objects
        If plSettings And OB_RemoveReferedObjects Then
            For lngLoop1 = 1 To UBound(strReferedObjects)
                Set nodx = nodr.Child
                For lngLoop2 = 1 To nodr.Children
                    If nodx.Text = strReferedObjects(lngLoop1) Then
                        Tree.Nodes.Remove nodx.Index
                        Exit For
                    End If
                    Set nodx = nodx.Next
                Next lngLoop2
            Next lngLoop1
        End If

End Sub

'Add all the members for an object.

Private Function AddMembersToObject(lngObject As Long, strNode As String) As Boolean

On Error Resume Next

Dim lngLoop2 As Long
Dim nodx As Node
Dim tliInvokeKinds As InvokeKinds
Dim tliTypeinfo As TypeInfo
Dim Default As Boolean
Dim strGMT As String
Dim strDescription As String
Dim strType As String

        AddMembersToObject = False
        Set tliTypeinfo = tlitypelibinfo.GetTypeInfo(lstTypeInfos.List(lngObject))
        tlitypelibinfo.GetMembersDirect lstTypeInfos.ItemData(lngObject), lstMembers.hwnd, tliWtListBox, tliIdtInvokeKinds, False
        For lngLoop2 = 0 To lstMembers.ListCount - 1
            ' each member will be added (type specific)
            tliInvokeKinds = lstMembers.ItemData(lngLoop2)
            strType = ""
            If tliInvokeKinds = INVOKE_FUNC Then
                If plSettings And OB_GroupMemberType Then
                    Set nodx = Tree.Nodes.Add(strNode, tvwChild, strNode & ".Method", "Method", 7, 7)
                    If Err.Number = 0 Then nodx.Tag = vbCrLf & "All methods for the " & strNode & " object."
                    nodx.Sorted = True
                    strGMT = ".Method"
                End If
                Set nodx = Tree.Nodes.Add(strNode & strGMT, tvwChild, strNode & ".Method." & lstMembers.List(lngLoop2), lstMembers.List(lngLoop2), 7, 7)
                nodx.Tag = CallingMember(lstTypeInfos.ItemData(lngObject), lstMembers.ItemData(lngLoop2), lstMembers.List(lngLoop2), , lngObject, strNode)
                nodx.Sorted = True
                strDescription = lstMembers.List(lngLoop2) & " method for the " & tliTypeinfo.Name & " object in the " & tlitypelibinfo.Name & " library."
                strType = "Method"
            ElseIf tliInvokeKinds = INVOKE_EVENTFUNC Then
                AddMembersToObject = True
                If plSettings And OB_GroupMemberType Then
                    Set nodx = Tree.Nodes.Add(strNode, tvwChild, strNode & ".Event", "Event", 6, 6)
                    If Err.Number = 0 Then nodx.Tag = vbCrLf & "All Events for the " & strNode & " object."
                    nodx.Sorted = True
                    strGMT = ".Event"
                End If
                Set nodx = Tree.Nodes.Add(strNode & strGMT, tvwChild, strNode & ".Event." & lstMembers.List(lngLoop2), lstMembers.List(lngLoop2), 6, 6)
                nodx.Tag = PrototypeMember(lstTypeInfos.ItemData(lngObject), lstMembers.ItemData(lngLoop2), lstMembers.List(lngLoop2), lstTypeInfos.List(lngObject))
                nodx.Sorted = True
                strDescription = lstMembers.List(lngLoop2) & " event for the " & tliTypeinfo.Name & " object in the " & tlitypelibinfo.Name & " library."
                strType = "Event"
            Else
                If tliTypeinfo.TypeKindString = "enum" Then
                    Set nodx = Tree.Nodes.Add(strNode, tvwChild, strNode & ".Enum." & lstMembers.List(lngLoop2), lstMembers.List(lngLoop2), 9, 9)
                    nodx.Tag = CallingMember(lstTypeInfos.ItemData(lngObject), lstMembers.ItemData(lngLoop2), lstMembers.List(lngLoop2), , lngObject, strNode)
                    nodx.Sorted = True
                    strDescription = lstMembers.List(lngLoop2) & " constant for the " & tliTypeinfo.Name & " enumeration in the " & tlitypelibinfo.Name & " library."
                    strType = "Enum"
                Else
                    If plSettings And OB_GroupMemberType Then
                        Set nodx = Tree.Nodes.Add(strNode, tvwChild, strNode & ".Property", "Property", 5, 5)
                        If Err.Number = 0 Then nodx.Tag = vbCrLf & "All Properties for the " & strNode & " object."
                        nodx.Sorted = True
                        strGMT = ".Property"
                    End If
                    ' hmm... how can we see if a property is the default property of the object
                    Default = False
                    On Error Resume Next
                    If Default Then
                        Set nodx = Tree.Nodes.Add(strNode & strGMT, tvwChild, strNode & ".Property." & lstMembers.List(lngLoop2), lstMembers.List(lngLoop2), 8, 8)
                    Else
                        Set nodx = Tree.Nodes.Add(strNode & strGMT, tvwChild, strNode & ".Property." & lstMembers.List(lngLoop2), lstMembers.List(lngLoop2), 5, 5)
                    End If
                    If Err.Number = 0 Then
                        nodx.Tag = CallingMember(lstTypeInfos.ItemData(lngObject), lstMembers.ItemData(lngLoop2), lstMembers.List(lngLoop2), , lngObject, strNode & ".Property." & lstMembers.List(lngLoop2))
                        nodx.Sorted = True
                        strDescription = lstMembers.List(lngLoop2) & " property for the " & tliTypeinfo.Name & " object in the " & tlitypelibinfo.Name & " library."
                    End If
                    strType = "Property"
                End If
            End If
            nodx.Tag = strDescription & vbCrLf & nodx.Tag
            If Len(nodx.Key) > 0 Then RaiseEvent Progress(nodx.Key)
        Next lngLoop2

End Function

' This is how the member can be called

Private Function CallingMember(ByVal SearchData As Long, ByVal InvokeKinds As InvokeKinds, Optional ByVal MemberName As String, Optional ByVal ParentName As String, Optional ByVal lngObject As Long, Optional ByVal strNode As String) As String

On Error Resume Next

Dim tliParameterInfo As ParameterInfo
Dim tliTypeinfo As TypeInfo
Dim tliResolvedTypeInfo As TypeInfo
Dim tliTypeKinds As TypeKinds

Dim ConstVal As Variant
Dim bisconstant As Boolean
Dim strReturn As String
Dim bFirstParameter As Boolean
Dim bParamArray As Boolean
Dim intVarTypeCur As Integer
Dim strTypeName As String
Dim bOptional As Boolean
Dim bDefault As Boolean
Dim strSep As String
Dim lngLoop As Long
Dim strSearchObject As String
Dim strPrev As String

        With tlitypelibinfo
            'First, determine the type of member we're dealing with
            bisconstant = GetSearchType(SearchData) And tliStConstants
            With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)

                'Now add the name of the member
                strReturn = lstTypeInfos.List(lngObject) & "." & .Name & " "

                'Do we have a return value
                If Not ((InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC) And (.ReturnType.VarType = VT_VOID Or .ReturnType.VarType = VT_HRESULT)) Then
                    Select Case .ReturnType.VarType
                    Case VT_VARIANT, VT_VOID, VT_HRESULT
                    Case Else
                        If InvokeKinds = INVOKE_EVENTFUNC Then strPrev = "Set "
                        ' Ok, the return varialbe is an object, so look if we can/have to extend this
                        If IsEmpty(.ReturnType.TypedVariant) Then
                            If plSettings And OB_ExtendInnerObjects Then
                                If .Name <> "Parent" Or (plSettings And OB_ExtendParent) Then
                                    On Error Resume Next
                                    strSearchObject = .ReturnType.TypeInfo
                                On Error GoTo 0
                                If strSearchObject <> "" Then
                                    ' The typeinfo of an object starts with an underscore. So we have to remove it first
                                    If Left(.ReturnType.TypeInfo, 1) = "_" Then
                                        strSearchObject = Mid(.ReturnType.TypeInfo, 2)
                                    Else
                                        strSearchObject = .ReturnType.TypeInfo
                                    End If
                                    ' Look if we refer to an object in the current library. Then extend it
                                    For lngLoop = 0 To lstTypeInfos.ListCount - 1
                                        If lstTypeInfos.List(lngLoop) = strSearchObject Then
                                            If InStr(1, strNode, "." & strSearchObject & ".") = 0 Then
                                                ReDim Preserve strReferedObjects(UBound(strReferedObjects) + 1)
                                                strReferedObjects(UBound(strReferedObjects)) = strSearchObject
                                                AddMembersToObject lngLoop, strNode
                                                tlitypelibinfo.GetMembersDirect lstTypeInfos.ItemData(lngObject), lstMembers.hwnd, tliWtListBox, tliIdtInvokeKinds, False
                                                Exit For
                                            End If
                                        End If
                                    Next lngLoop
                                End If
                            End If
                        End If
                        strReturn = "ReturnValue = " & strReturn
                        On Error Resume Next
                            strReturn = .ReturnType.TypeInfo & strReturn
                        On Error GoTo 0
                        strReturn = strPrev & strReturn
                    Else
                        strReturn = TypeName(.ReturnType.TypedVariant) & "ReturnValue = " & strReturn
                    End If

                    If InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then strReturn = strReturn & "("
                End Select
            End If
            'Process the member's parameters
            With .Parameters
                If .Count Then

                    bFirstParameter = True
                    bParamArray = .OptionalCount = -1
                    strSep = ""
                    For Each tliParameterInfo In .Me
                        'Determine whether parameter is default, optional, etc.            •••
                        With tliParameterInfo.VarTypeInfo
                            'The seperator for multiple variables
                            strReturn = strReturn & strSep

                            ' mark optional values
                            If tliParameterInfo.Optional Then strReturn = strReturn & "["

                            ' Variable name
                            strReturn = strReturn & tliParameterInfo.Name

                            ' mark optional values with default value
                            If tliParameterInfo.Default Then
                                strReturn = strReturn & " = " & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                            End If
                            ' mark optional values
                            If tliParameterInfo.Optional Then strReturn = strReturn & "]"

                        End With
                        strSep = ", "
                    Next tliParameterInfo

                End If
            End With
            If Not ((InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC) And (.ReturnType.VarType = VT_VOID Or .ReturnType.VarType = VT_HRESULT)) Then
                Select Case .ReturnType.VarType
                Case VT_VARIANT, VT_VOID, VT_HRESULT
                Case Else
                    If InvokeKinds = INVOKE_FUNC Or InvokeKinds = INVOKE_EVENTFUNC Then strReturn = strReturn & ")"
                End Select
            End If

            CallingMember = strReturn & vbCrLf & .HelpString

            If GetSearchType(SearchData) And tliStConstants Then
                CallingMember = CallingMember & " (Value = " & .Value & ")"
            End If
        End With
    End With

End Function

' This is the prototype of the member

Private Function PrototypeMember(ByVal SearchData As Long, ByVal InvokeKinds As InvokeKinds, Optional ByVal MemberName As String, Optional ByVal ParentName As String) As String

On Error Resume Next

    Dim tliParameterInfo As ParameterInfo
    Dim tliTypeinfo As TypeInfo
    Dim tliResolvedTypeInfo As TypeInfo
    Dim tliTypeKinds As TypeKinds

    Dim ConstVal As Variant
    Dim bisconstant As Boolean
    Dim strReturn As String
    Dim bFirstParameter As Boolean
    Dim bParamArray As Boolean
    Dim intVarTypeCur As Integer
    Dim strTypeName As String
    Dim bOptional As Boolean
    Dim bDefault As Boolean
    Dim strSep As String

        With tlitypelibinfo
            'First, determine the type of member we're dealing with
            bisconstant = GetSearchType(SearchData) And tliStConstants
            With .GetMemberInfo(SearchData, InvokeKinds, , MemberName)
                'Now add the name of the member
                strReturn = "Sub " & ParentName & "_" & .Name

                'Process the member's parameters
                With .Parameters
                    If .Count Then
                        strReturn = strReturn & " ("
                        bFirstParameter = True
                        bParamArray = .OptionalCount = -1
                        strSep = ""
                        For Each tliParameterInfo In .Me
                            'Determine whether parameter is default, optional, etc.            •••
                            With tliParameterInfo.VarTypeInfo
                                'The seperator for multiple variables
                                strReturn = strReturn & strSep

                                ' mark optional values
                                If tliParameterInfo.Optional Then strReturn = strReturn & "Optional "

                                ' Varialbe referencing
                                If .PointerLevel = 0 Then
                                    strReturn = strReturn & "ByVal "
                                Else
                                    strReturn = strReturn & "ByRef "
                                End If

                                ' Variable name
                                strReturn = strReturn & tliParameterInfo.Name

                                ' Variable type information
                                Set tliResolvedTypeInfo = Nothing
                                tliTypeKinds = TKIND_MAX
                                intVarTypeCur = .VarType
                                If (intVarTypeCur And Not (VT_ARRAY Or VT_VECTOR)) = 0 Then
                                    strReturn = strReturn & " As " & Mid(.TypeInfo.Name, 2)
                                Else
                                    If intVarTypeCur <> vbVariant Then
                                        strTypeName = TypeName(.TypedVariant)
                                        If intVarTypeCur And (VT_ARRAY Or VT_VECTOR) Then
                                            strReturn = strReturn & "() As " & Left$(strTypeName, Len(strTypeName) - 2)
                                        Else
                                            strReturn = strReturn & " As " & strTypeName
                                        End If
                                    End If
                                End If

                                ' mark optional values with default value
                                If tliParameterInfo.Default Then
                                    strReturn = strReturn & ProduceDefaultValue(tliParameterInfo.DefaultValue, tliResolvedTypeInfo)
                                End If
                            End With
                            strSep = ", "
                        Next tliParameterInfo

                        strReturn = strReturn & ")"
                    End If
                End With

                If bisconstant Then
                    ConstVal = .Value
                    strReturn = strReturn & " = " & ConstVal
                Else
                    Select Case .ReturnType.VarType
                    Case VT_VARIANT, VT_VOID, VT_HRESULT
                    Case Else
                        If IsEmpty(.ReturnType.TypedVariant) Then
                            strReturn = strReturn & " As " & Mid(.ReturnType.TypeInfo, 2)
                        Else
                            strReturn = strReturn & " As " & TypeName(.ReturnType.TypedVariant)
                        End If
                    End Select
                End If

                PrototypeMember = strReturn & vbCrLf & .HelpString
            End With
        End With

End Function

' Returns the default value of a variable. Is used in the PrototypeMember function

Private Function ProduceDefaultValue(DefVal As Variant, ByVal TI As TypeInfo) As String

On Error Resume Next

Dim lTrackVal As Long
Dim MI As MemberInfo
Dim TKind As TypeKinds

    If TI Is Nothing Then
        Select Case VarType(DefVal)
        Case VBString
            If Len(DefVal) Then
                ProduceDefaultValue = """" & DefVal & """"
            End If
        Case vbBoolean 'Always show for Boolean
            ProduceDefaultValue = DefVal
        Case vbDate
            If DefVal Then
                ProduceDefaultValue = "#" & DefVal & "#"
            End If
        Case Else 'Numeric Values
            If DefVal <> 0 Then
                ProduceDefaultValue = DefVal
            End If
        End Select
    Else
        'See if we have an enum and track the matching member
        'If the type is an object, then there will never be a
        'default value other than Nothing
        TKind = TI.TypeKind
        Do While TKind = TKIND_ALIAS
            TKind = TKIND_MAX
            On Error Resume Next
                Set TI = TI.ResolvedType
                If Err = 0 Then TKind = TI.TypeKind
            On Error GoTo 0
        Loop
        If TI.TypeKind = TKIND_ENUM Then
            lTrackVal = DefVal
            For Each MI In TI.Members
                If MI.Value = lTrackVal Then
                    ProduceDefaultValue = MI.Name
                    Exit For
                End If
            Next MI
        End If
    End If

End Function

Private Function GetSearchType(ByVal SearchData As Long) As TliSearchTypes

On Error Resume Next

    If SearchData And &H80000000 Then
        GetSearchType = ((SearchData And &H7FFFFFFF) \ &H1000000 And &H7F&) Or &H80
    Else
        GetSearchType = SearchData \ &H1000000 And &HFF&
    End If

End Function


Private Sub Form_Resize()

On Error Resume Next

    Me.Tree.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Me.StatusBar.Height
    
End Sub

Private Sub Form_Load()

On Error Resume Next

    'Public Enum ObjectBrowserSettings
    '    OB_GroupMemberType = 1           ' This will add an extra level to the tree. All properties will be in a property node, all methods will be in a method node and all events will be in an event node.
    '    OB_EnumInRoot = 2                ' If you use ExtendInnerObjects, then Enumerations will be shown where they are used. Then you will probably not want to see enumeration objects in the root
    '    OB_ExtendInnerObjects = 4        ' If a members returns an object/enum that is in this library, then that object can be put under this member.
    '    OB_ExtendParent = 8              ' When Objects are extended then you might want to disable extending the parrent since that is already done a level up in the tree.
    '    OB_RemoveReferedObjects = 16     ' When inner objects are extended then you can remove the objects that are an extended object since they are probably only accessed through that other object. This will make the tree smaller and it will only show you the main objects that you will use.
    'End Enum
    
    
    settings = OB_EnumInRoot
    
    
    'AddVBDefaultObjects
    
        
    'AddFromFile "C:\WINDOWS\system32\mscomctl.ocx"
    
    'AddFromFile "olelib.tlb"
    
   'GetObjPath ""
    
    'Clipboard.SetText GetWordList
        
End Sub




Sub AddVBDefaultObjects()

On Error Resume Next

   AddFromFile ("vb6.olb")
   AddFromFile ("msvbvm60.dll\1")
   'AddFromFile ("msvbvm60.dll\2")
   AddFromFile ("msvbvm60.dll\3")
   AddFromFile ("stdole2.tlb")
    
End Sub

Private Function InStrArr(Value As String, a() As String) As Long
On Error Resume Next

    Dim i&
    
    For i = LBound(a) To UBound(a)
        
        If Value = a(i) Then
            
            InStrArr = i
            Exit Function
        
        End If
    
    Next i

End Function

Public Function GetWordList() As String

    Dim i As Long
    Dim s As String
    
    Dim Word As String
    Dim Words() As String
    
    ReDim Words(0)
    
    For i = 1 To Me.Tree.Nodes.Count
        
        Word = Me.Tree.Nodes.Item(i).Text
        
        If InStrArr(Word, Words) = 0 Then
        
            ReDim Preserve Words(UBound(Words) + 1)
            
            Words(UBound(Words)) = Word
        
        End If
        
    
    Next i
    
    'Debug.Print UBound(Words)
    
    GetWordList = Join(Words, " ")
    
End Function


Function GetObjPath(ClsID As String) As String
    
    Dim Reg As RegistryClass
    Dim Sections() As String
    Dim c As Long
    Dim FilePath As String
    
    Set Reg = New RegistryClass
    
    With Reg
        
        .ClassKey = HKEY_CLASSES_ROOT
        
      
        .SectionKey = "TypeLib\" & ClsID
                
        .EnumerateSections Sections, c
                
        For i = 1 To c
            
            .SectionKey = "TypeLib\" & ClsID & "\" & Sections(i) & "\0\win32"
            FilePath = .Value
            If FilePath <> "" Then
                GetObjPath = FilePath
            End If
                    
        Next i
                
    End With
    
    
    
End Function


Function AddFromID(ID As String)

    Dim FilePath As String
    
    FilePath = GetObjPath(ID)
    
    If FileExist(FilePath) Then
        AddFromFile GetObjPath(ID)
    Else
        Debug.Print "Can't find library: " & ID
    End If
    
End Function
