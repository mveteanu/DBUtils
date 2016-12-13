VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   5850
   Begin VB.PictureBox picTab 
      Height          =   1215
      Index           =   4
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   4395
      TabIndex        =   10
      Top             =   5880
      Width           =   4455
      Begin MSComctlLib.ListView lstDBO 
         Height          =   855
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1508
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListBig"
         SmallIcons      =   "imlToolbarIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Modified"
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Created"
            Text            =   "Created"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Type"
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "Path"
            Text            =   "Full path"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgListBig 
      Left            =   5040
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":0442
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":075C
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":0A76
            Key             =   "IcoProc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":0D90
            Key             =   "IcoFile"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTab 
      Height          =   1215
      Index           =   3
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   4395
      TabIndex        =   5
      Top             =   4560
      Width           =   4455
      Begin MSComctlLib.ListView lstDBO 
         Height          =   855
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1508
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListBig"
         SmallIcons      =   "imlToolbarIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Modified"
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Created"
            Text            =   "Created"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Type"
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   1215
      Index           =   2
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   3240
      Width           =   4455
      Begin MSComctlLib.ListView lstDBO 
         Height          =   855
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1508
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListBig"
         SmallIcons      =   "imlToolbarIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Modified"
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Created"
            Text            =   "Created"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Type"
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   1095
      Index           =   1
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   4395
      TabIndex        =   3
      Top             =   2040
      Width           =   4455
      Begin MSComctlLib.ListView lstDBO 
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1508
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListBig"
         SmallIcons      =   "imlToolbarIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Modified"
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Created"
            Text            =   "Created"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Type"
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   1095
      Index           =   0
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   840
      Width           =   4455
      Begin MSComctlLib.ListView lstDBO 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1508
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imgListBig"
         SmallIcons      =   "imlToolbarIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Name"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Modified"
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "Created"
            Text            =   "Created"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "Type"
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            Object.Tag             =   "1"
            ImageKey        =   "View Details"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            Object.Tag             =   "1"
            ImageKey        =   "View Large Icons"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            Object.Tag             =   "1"
            ImageKey        =   "View List"
            Value           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            Object.Tag             =   "1"
            ImageKey        =   "View Small Icons"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   12091
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tables"
            Key             =   "Tables"
            Object.ToolTipText     =   "Database tables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Queries"
            Key             =   "Queries"
            Object.ToolTipText     =   "Database procedures and views"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Views"
            Key             =   "Views"
            Object.ToolTipText     =   "Database views"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Procedures"
            Key             =   "Procedures"
            Object.ToolTipText     =   "Database stored procs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            Key             =   "Files"
            Object.ToolTipText     =   "Project source code files"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5040
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":10AA
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":11BC
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":12CE
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":13E0
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":14F2
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1604
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":175E
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18B8
            Key             =   "IcoProc"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1A12
            Key             =   "IcoFile"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public objDBCon As clsDBConnection
Private ItemClickName As String

Private Sub AddListViewItem(lst As ListView, ByVal nume As String, ByVal datemod As Variant, ByVal datecre As Variant, ByVal tip As String, ByVal icon As Variant)
    Dim strDatemod As String, strDatecre As String
    
    If VarType(datemod) = vbDate Then strDatemod = FormatDateTime(datemod, vbGeneralDate)
    If VarType(datecre) = vbDate Then strDatecre = FormatDateTime(datecre, vbGeneralDate)
    With lst.ListItems.Add(, , nume, icon, icon)
        .ListSubItems.Add , , strDatemod
        .ListSubItems.Add , , strDatecre
        .ListSubItems.Add , , tip
    End With
End Sub

Private Sub FillWithObjects()
    Dim tbl As ADOX.Table
    Dim prc As ADOX.Procedure
    Dim viw As ADOX.View
    Dim fil As Variant
    
    If Not objDBCon.ConnectToDB Then
        MsgBox "Error connecting to database!", vbOKOnly + vbCritical
    Else
        fMainForm.sbStatusBar.Panels(1).Text = "Loading table list..."
        For Each tbl In objDBCon.objCat.Tables
            If tbl.Type = "TABLE" Then
                AddListViewItem lstDBO(0), tbl.Name, tbl.datemodified, tbl.datecreated, dbuTypeTable, "IcoTable"
            End If
            DoEvents
        Next
        fMainForm.sbStatusBar.Panels(1).Text = "Loading views list..."
        For Each viw In objDBCon.objCat.Views
            AddListViewItem lstDBO(1), viw.Name, viw.datemodified, viw.datecreated, dbuTypeView, "IcoView"
            AddListViewItem lstDBO(2), viw.Name, viw.datemodified, viw.datecreated, dbuTypeView, "IcoView"
            DoEvents
        Next
        fMainForm.sbStatusBar.Panels(1).Text = "Loading procedures list..."
        For Each prc In objDBCon.objCat.Procedures
            AddListViewItem lstDBO(1), prc.Name, prc.datemodified, prc.datecreated, dbuTypeProc, "IcoProc"
            AddListViewItem lstDBO(3), prc.Name, prc.datemodified, prc.datecreated, dbuTypeProc, "IcoProc"
            DoEvents
        Next
        
        If Len(objDBCon.strFolderPath) > 0 Then
            fMainForm.sbStatusBar.Panels(1).Text = "Loading files list..."
            If objDBCon.ConnectToFiles Then
                For Each fil In objDBCon.Files
                    With lstDBO(4).ListItems.Add(, , fil(0), "IcoFile", "IcoFile")
                        .ListSubItems.Add , , fil(2)
                        .ListSubItems.Add , , fil(1)
                        .ListSubItems.Add , , fil(3)
                        .ListSubItems.Add , , fil(4)
                    End With
                Next
            Else
                MsgBox "Error connecting to files!", vbOKOnly + vbCritical
            End If
        End If
        fMainForm.sbStatusBar.Panels(1).Text = "Ready"
    End If
End Sub

Private Sub Form_Activate()
    Static DontConnect As Boolean
    If Not DontConnect Then
        Screen.MousePointer = vbHourglass
        Call FillWithObjects
        Screen.MousePointer = vbDefault
        DontConnect = True
    End If
End Sub

Private Sub Form_Load()
    Form_Resize
    TabStrip1_Click
End Sub

Private Sub Form_Resize()
    Dim pic As PictureBox
    Dim i As Integer
    
    On Error Resume Next
    TabStrip1.Move 0, 420, Me.ScaleWidth, Me.ScaleHeight - 420
    
    For i = 0 To picTab.Count - 1
        picTab(i).Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
        picTab(i).BorderStyle = 0
        lstDBO(i).Move 0, 0, picTab(i).Width, picTab(i).Height
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objDBCon = Nothing
End Sub

Private Sub lstDBO_DblClick(Index As Integer)
    Dim ItemName As String
    Dim ItemType As String
    
    With lstDBO(Index).SelectedItem
        ItemName = .Text
        ItemType = .ListSubItems(3).Text
    End With
    If ItemName <> ItemClickName Then Exit Sub
    ItemClickName = ""
    
    Select Case ItemType
        Case dbuTypeTable
            Call ShowTableStructure(ItemName)
        Case dbuTypeView
            Call ShowObjectCode(ItemName, dbuTypeView)
        Case dbuTypeProc
            Call ShowObjectCode(ItemName, dbuTypeProc)
    End Select
End Sub

Private Sub lstDBO_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    ItemClickName = Item.Text
End Sub

Private Sub lstDBO_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ItemClickName = lstDBO(Index).SelectedItem.Text
        Call lstDBO_DblClick(Index)
    End If
End Sub

Private Sub lstDBO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim LI As ListItem
    Dim strName As String
    Dim strType As String
    Dim bFileIsBin As Boolean
    
    Set LI = lstDBO(Index).HitTest(x, y)
    If LI Is Nothing Then Exit Sub
    
    If Button And vbRightButton Then
        LI.Selected = True
        strName = LI.Text
        strType = LI.ListSubItems(3).Text
        
        If strType = dbuTypeFile Then
            strName = LI.ListSubItems(4).Text
            bFileIsBin = objDBCon.Files(strName)(5)
        End If
        Call HandleContextMenu(ShowPopupMenu(Me, strType, bFileIsBin), strName, strType)
    End If
End Sub

Private Sub TabStrip1_Click()
    Dim pic As PictureBox
    For Each pic In picTab
        pic.Visible = (pic.Index = TabStrip1.SelectedItem.Index - 1)
    Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "View Details", "View Large Icons", "View List", "View Small Icons"
            ApplyView Button.Key
        Case "Properties"
            Call ShowDBProperties
    End Select
End Sub

Private Sub ApplyView(ButtonKey As String)
    Dim NewView As Integer
    Dim lst As ListView
    Dim but As MSComctlLib.Button
    
    Select Case ButtonKey
        Case "View Details"
            NewView = lvwReport
        Case "View Large Icons"
            NewView = lvwIcon
        Case "View List"
            NewView = lvwList
        Case "View Small Icons"
            NewView = lvwSmallIcon
    End Select
    
    For Each but In Toolbar1.Buttons
        If but.Tag = 1 Then but.Value = tbrUnpressed
    Next
    Toolbar1.Buttons(ButtonKey).Value = tbrPressed
    Toolbar1.Refresh
    
    For Each lst In lstDBO
        lst.View = NewView
    Next
End Sub


