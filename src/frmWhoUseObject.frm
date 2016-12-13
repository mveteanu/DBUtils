VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWhoUseObject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Who use the object"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmWhoUseObject.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageCombo cboObjects 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imlIcons"
   End
   Begin MSComctlLib.TreeView treeObjects 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlIcons"
      Appearance      =   1
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoUseObject.frx":0442
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoUseObject.frx":059C
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWhoUseObject.frx":06F6
            Key             =   "IcoProc"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Childs of selected object:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "DB Objects:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmWhoUseObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strInitialObject As String
Public strInitialObjectType As String

Private SQLObjects As New Collection
Private objCatLocal As ADOX.Catalog
Private FirstTime As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnClose_Click
End Sub


Private Sub Form_Load()
    Dim t As ADOX.Table
    Dim v As ADOX.View
    Dim p As ADOX.Procedure
    Dim i As Integer
    
    Set objCatLocal = fMainForm.ActiveForm.objDBCon.objCat
    For Each t In objCatLocal.Tables
        DoEvents
        If t.Type = dbuTypeTable Then
            With cboObjects.ComboItems.Add(, "t" & t.Name, t.Name, DBObjectIcon(dbuTypeTable))
                If (strInitialObjectType = dbuTypeTable) And (t.Name = strInitialObject) Then i = .Index
            End With
        End If
    Next
    For Each v In objCatLocal.Views
        DoEvents
        SQLObjects.Add Array(v.Name, dbuTypeView, GetObjectText(objCatLocal, v.Name, dbuTypeView)), "v" & v.Name
        With cboObjects.ComboItems.Add(, "v" & v.Name, v.Name, DBObjectIcon(dbuTypeView))
            If (strInitialObjectType = dbuTypeView) And (v.Name = strInitialObject) Then i = .Index
        End With
    Next
    For Each p In objCatLocal.Procedures
        DoEvents
        SQLObjects.Add Array(p.Name, dbuTypeProc, GetObjectText(objCatLocal, p.Name, dbuTypeProc)), "p" & p.Name
        With cboObjects.ComboItems.Add(, "p" & p.Name, p.Name, DBObjectIcon(dbuTypeProc))
            If (strInitialObjectType = dbuTypeProc) And (p.Name = strInitialObject) Then i = .Index
        End With
    Next
    FirstTime = True
    FillFormForListIndex i
End Sub

Private Sub FillFormForListIndex(idx As Integer)
    Dim strName As String
    Dim strType As String
    
    With cboObjects.ComboItems.Item(idx)
        strName = .Text
        Select Case Left(.Key, 1)
            Case "t": strType = dbuTypeTable
            Case "v": strType = dbuTypeView
            Case "p": strType = dbuTypeProc
        End Select
        .Selected = True
    End With
    
    FirstTime = False
    
    Screen.MousePointer = vbHourglass
    If Not BuildTreeView(strName, strType, "", treeObjects) Then
        MsgBox "DB objects hierarhy is too deep to be rendered!" & vbCrLf & vbCrLf & _
               "Reasons: " & vbCrLf & _
               " - tables and queries have strange names that complicates the life of this tool, or" & vbCrLf & _
               " - you have bugs in your queries that produce an infinite recursive call", vbExclamation + vbOKOnly, "DB Utils"
               
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Function BuildTreeView(ByVal strName As String, ByVal strType As String, ByVal strParentKey As String, ByVal TV As TreeView) As Boolean
    Dim strKey As String
    Dim v As Variant
    
    strKey = strParentKey & strType & strName
    If Len(strKey) > 1024 Then
        BuildTreeView = False
        Exit Function
    End If
    
    If Len(strParentKey) = 0 Then
        TV.Nodes.Clear
        With TV.Nodes.Add(, , strKey, strName, DBObjectIcon(strType))
            .Bold = True
        End With
    Else
        TV.Nodes.Add strParentKey, tvwChild, strKey, strName, DBObjectIcon(strType)
    End If
    TV.Nodes(TV.Nodes.Count).Expanded = True
    
    For Each v In SQLObjects
        If strName <> v(0) Then
            If WordInsideText(strName, v(2), 1) Then
                If Not BuildTreeView(v(0), v(1), strKey, TV) Then
                    BuildTreeView = False
                    Exit Function
                End If
            End If
        End If
    Next
    BuildTreeView = True
End Function

Private Sub cboObjects_Click()
    If FirstTime Then Exit Sub
    FillFormForListIndex cboObjects.SelectedItem.Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SQLObjects = Nothing
    Set objCatLocal = Nothing
End Sub

Private Sub Form_Activate()
    btnClose.SetFocus
End Sub

Private Sub treeObjects_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim LI As Node
    Dim strName As String
    Dim strType As String
    
    Set LI = treeObjects.HitTest(x, y)
    If LI Is Nothing Then Exit Sub
    
    If Button And vbRightButton Then
        LI.Selected = True
        strName = LI.Text
        strType = DBIconOf(LI.Image)
        Call HandleContextMenu(ShowPopupMenu(Me, strType), strName, strType)
    End If
End Sub

