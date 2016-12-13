VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBObjectsInFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DB Objects used in file"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmDBObjectsInFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin MSComctlLib.ImageCombo cboFiles 
      Height          =   330
      Left            =   120
      TabIndex        =   3
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
   Begin MSComctlLib.ListView lstObjects 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "DBObject"
         Text            =   "DB Object"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBObjectsInFile.frx":0442
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBObjectsInFile.frx":059C
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBObjectsInFile.frx":06F6
            Key             =   "IcoProc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDBObjectsInFile.frx":0850
            Key             =   "IcoFile"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "DB Objects in file:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Text files:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmDBObjectsInFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strInitialFileName As String

Private SQLObjects As New Collection
Private objConLocal As clsDBConnection
Private FirstTime As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnClose_Click
End Sub

Private Sub Form_Activate()
    btnClose.SetFocus
End Sub

Private Sub Form_Load()
    Dim t As ADOX.Table
    Dim v As ADOX.View
    Dim p As ADOX.Procedure
    Dim f As Variant
    Dim i As Integer
    
    Set objConLocal = fMainForm.ActiveForm.objDBCon
    
    For Each f In objConLocal.Files
        If Not f(5) Then
            With cboFiles.ComboItems.Add(, , f(4), DBObjectIcon(dbuTypeFile))
                If f(4) = strInitialFileName Then i = .Index
            End With
        End If
    Next
    For Each t In objConLocal.objCat.Tables
        If t.Type = dbuTypeTable Then
            SQLObjects.Add Array(t.Name, dbuTypeTable), "t" & t.Name
        End If
    Next
    For Each v In objConLocal.objCat.Views
        SQLObjects.Add Array(v.Name, dbuTypeView), "v" & v.Name
    Next
    For Each p In objConLocal.objCat.Procedures
        SQLObjects.Add Array(p.Name, dbuTypeProc), "p" & p.Name
    Next
    FirstTime = True
    FillFormForListIndex i
End Sub

Private Sub FillFormForListIndex(idx As Integer)
    Dim strName As String
    Dim v As Variant
    
    With cboFiles.ComboItems.Item(idx)
        strName = .Text
        .Selected = True
    End With
    
    FirstTime = False
    
    Screen.MousePointer = vbHourglass
    lstObjects.ListItems.Clear
    For Each v In SQLObjects
        If WordInsideText(v(0), GetTextFileContent(strName), 2) Then
            With lstObjects.ListItems.Add(, , v(0), , DBObjectIcon(v(1)))
                .ListSubItems.Add , , v(1)
            End With
        End If
    Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub cboFiles_Click()
    If FirstTime Then Exit Sub
    FillFormForListIndex cboFiles.SelectedItem.Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SQLObjects = Nothing
    Set objConLocal = Nothing
End Sub

Private Sub lstObjects_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim LI As ListItem
    Dim strName As String
    Dim strType As String
    
    Set LI = lstObjects.HitTest(x, y)
    If LI Is Nothing Then Exit Sub
    
    If Button And vbRightButton Then
        LI.Selected = True
        strName = LI.Text
        strType = LI.ListSubItems(1).Text
        Call HandleContextMenu(ShowPopupMenu(Me, strType), strName, strType)
    End If
End Sub

