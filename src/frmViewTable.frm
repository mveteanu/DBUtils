VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table structure viewer"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmViewTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageCombo cboTables 
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
   Begin MSComctlLib.ListView lstColumns 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Field Name"
         Text            =   "Field Name"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "DataType"
         Text            =   "Data Type"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Attributes"
         Text            =   "Attributes"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Description"
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
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
            Picture         =   "frmViewTable.frx":0442
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTable.frx":059C
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewTable.frx":06F6
            Key             =   "IcoProc"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Columns:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "DB Tables:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmViewTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strInitialObject As String

Private objCatLocal As ADOX.Catalog
Private FirstTime As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cboTables_Click()
    If FirstTime Then Exit Sub
    FillFormForListIndex cboTables.SelectedItem.Index
End Sub

Private Sub Form_Activate()
    btnClose.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnClose_Click
End Sub

Private Sub Form_Load()
    Dim t As ADOX.Table
    Dim i As Integer
    
    Set objCatLocal = fMainForm.ActiveForm.objDBCon.objCat
    For Each t In objCatLocal.Tables
        If t.Type = dbuTypeTable Then
            With cboTables.ComboItems.Add(, "t" & t.Name, t.Name, DBObjectIcon(dbuTypeTable))
                If t.Name = strInitialObject Then i = .Index
            End With
        End If
    Next
    FirstTime = True
    FillFormForListIndex i
End Sub

Private Sub FillFormForListIndex(idx As Integer)
    Dim tbl As ADOX.Table
    Dim col As ADOX.Column
    
    With cboTables.ComboItems.Item(idx)
        Set tbl = objCatLocal.Tables(.Text)
        .Selected = True
    End With
    FirstTime = False
    
    lstColumns.ListItems.Clear
    
    For Each col In tbl.Columns
        With lstColumns.ListItems.Add(, , col.Name)
            .ListSubItems.Add , , DataTypeEnum(col.Type) & IIf(col.Type = adVarWChar, " (" & col.DefinedSize & ")", "")
            .ListSubItems.Add , , IIf(col.Properties("Autoincrement").Value, " Autonumber", "")
            .ListSubItems.Add , , col.Properties("Description").Value
        End With
    Next
    
    Set tbl = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objCatLocal = Nothing
End Sub
