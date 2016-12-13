VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "VMASOFT DBUtils"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New connection"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            Object.ToolTipText     =   "Properties"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search in database"
            ImageKey        =   "Search"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help contents"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2925
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2593
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2/1/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "5:18 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1800
      Top             =   1200
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
            Picture         =   "frmMain.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0554
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06AE
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":07C0
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuContextualMenus 
      Caption         =   "Contextual menus"
      Visible         =   0   'False
      Begin VB.Menu mnuTableMenu 
         Caption         =   "Table Menu"
         Begin VB.Menu mnuTableItem 
            Caption         =   "View Table Structure"
            Index           =   1
         End
         Begin VB.Menu mnuTableItem 
            Caption         =   "View Contents"
            Index           =   2
         End
         Begin VB.Menu mnuTableItem 
            Caption         =   "Who Use Table"
            Index           =   3
         End
         Begin VB.Menu mnuTableItem 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuTableItem 
            Caption         =   "In which files is used"
            Index           =   5
         End
      End
      Begin VB.Menu mnuViewMenu 
         Caption         =   "View Menu"
         Begin VB.Menu mnuViewItem 
            Caption         =   "View SQL Code"
            Index           =   1
         End
         Begin VB.Menu mnuViewItem 
            Caption         =   "View Contents"
            Index           =   2
         End
         Begin VB.Menu mnuViewItem 
            Caption         =   "Who Use View"
            Index           =   3
         End
         Begin VB.Menu mnuViewItem 
            Caption         =   "What Use View"
            Index           =   4
         End
         Begin VB.Menu mnuViewItem 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuViewItem 
            Caption         =   "In which files is used"
            Index           =   6
         End
      End
      Begin VB.Menu mnuProcMenu 
         Caption         =   "Proc Menu"
         Begin VB.Menu mnuProcItem 
            Caption         =   "View SQL Code"
            Index           =   1
         End
         Begin VB.Menu mnuProcItem 
            Caption         =   "Who Use Procedure"
            Index           =   2
         End
         Begin VB.Menu mnuProcItem 
            Caption         =   "What Use Procedure"
            Index           =   3
         End
         Begin VB.Menu mnuProcItem 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuProcItem 
            Caption         =   "In which files is used"
            Index           =   5
         End
      End
      Begin VB.Menu mnuFileMenu 
         Caption         =   "File menu"
         Begin VB.Menu mnuFileItem 
            Caption         =   "Open with registred program"
            Index           =   1
         End
         Begin VB.Menu mnuFileItem 
            Caption         =   "Open with Notepad"
            Index           =   2
         End
         Begin VB.Menu mnuFileItem 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuFileItem 
            Caption         =   "What DB object use"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSQLView 
         Caption         =   "SQLView menu"
         Begin VB.Menu mnuSQLViewItem 
            Caption         =   "Copy text"
            Index           =   1
         End
         Begin VB.Menu mnuSQLViewItem 
            Caption         =   "Copy formated text"
            Index           =   2
         End
         Begin VB.Menu mnuSQLViewItem 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuSQLViewItem 
            Caption         =   "Search selected text"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuTopFileMenu 
      Caption         =   "&File"
      Begin VB.Menu mnuTopFileItem 
         Caption         =   "&New"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuTopFileItem 
         Caption         =   "&Close"
         Index           =   2
      End
      Begin VB.Menu mnuTopFileItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTopFileItem 
         Caption         =   "Propert&ies"
         Index           =   4
      End
      Begin VB.Menu mnuTopFileItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTopFileItem 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTopView 
      Caption         =   "&View"
      Begin VB.Menu mnuTopViewItem 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuTopViewItem 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
         Index           =   2
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowItem 
         Caption         =   "&New Window"
         Index           =   1
      End
      Begin VB.Menu mnuWindowItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuWindowItem 
         Caption         =   "&Cascade"
         Index           =   3
      End
      Begin VB.Menu mnuWindowItem 
         Caption         =   "Tile &Horizontal"
         Index           =   4
      End
      Begin VB.Menu mnuWindowItem 
         Caption         =   "Tile &Vertical"
         Index           =   5
      End
      Begin VB.Menu mnuWindowItem 
         Caption         =   "&Arrange Icons"
         Index           =   6
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsItems 
         Caption         =   "Search text in database"
         Index           =   1
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuTopHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTopHelpItem 
         Caption         =   "&Contents"
         Index           =   1
      End
      Begin VB.Menu mnuTopHelpItem 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTopHelpItem 
         Caption         =   "&About "
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
    Static DontShow As Boolean
    If Not DontShow Then
        OpenNewDatabase
        DontShow = True
    End If
End Sub

Private Sub OpenNewDatabase()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    Dim objCon As clsDBConnection
    
    Set objCon = New clsDBConnection
    If objCon.ShowCreateConnectionDialog = vbOK Then
        lDocumentCount = lDocumentCount + 1
        Set frmD = New frmDocument
        frmD.Caption = "Database " & lDocumentCount
        Set frmD.objDBCon = objCon
        If Len(objCon.strFolderPath) = 0 Then frmD.TabStrip1.Tabs.Remove (5)
        frmD.Show
    End If
End Sub

Private Sub MDIForm_Load()
    With New clsPersistentSettings
        .LoadWindowPosition Me
    End With
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    With New clsPersistentSettings
        .SaveWindowPosition Me
    End With
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            Call OpenNewDatabase
        Case "Properties"
            If Not (ActiveForm Is Nothing) Then
                Call ShowDBProperties
            End If
        Case "Search"
            If Not (ActiveForm Is Nothing) Then
                Call ShowSearchText
            End If
        Case "Help"
            Call mnuTopHelpItem_Click(1)
    End Select
End Sub


' Meniul View
Private Sub mnuTopViewItem_Click(Index As Integer)
    Select Case Index
        Case 1
            mnuTopViewItem(1).Checked = Not mnuTopViewItem(1).Checked
            tbToolBar.Visible = mnuTopViewItem(1).Checked
        Case 2
            mnuTopViewItem(2).Checked = Not mnuTopViewItem(2).Checked
            sbStatusBar.Visible = mnuTopViewItem(2).Checked
    End Select
End Sub



' Meniul Tools
Private Sub mnuTools_Click()
    mnuToolsItems(1).Enabled = Not (ActiveForm Is Nothing)
End Sub
Private Sub mnuToolsItems_Click(Index As Integer)
    Select Case Index
        Case 1
            If Not (ActiveForm Is Nothing) Then
                Call ShowSearchText
            End If
    End Select
End Sub


' Meniul Window
Private Sub mnuWindowItem_Click(Index As Integer)
    Select Case Index
        Case 1: Call OpenNewDatabase
        Case 3: Me.Arrange vbCascade
        Case 4: Me.Arrange vbTileHorizontal
        Case 5: Me.Arrange vbTileVertical
        Case 6: Me.Arrange vbArrangeIcons
    End Select
End Sub


' Meniul File
Private Sub mnuTopFileMenu_Click()
    mnuTopFileItem(2).Enabled = Not (ActiveForm Is Nothing)
    mnuTopFileItem(4).Enabled = Not (ActiveForm Is Nothing)
End Sub
Private Sub mnuTopFileItem_Click(Index As Integer)
    Select Case Index
        Case 1: Call OpenNewDatabase
        Case 2: If Not (ActiveForm Is Nothing) Then Unload ActiveForm
        Case 4: If Not (ActiveForm Is Nothing) Then Call ShowDBProperties
        Case 6: Unload Me
    End Select
End Sub


' Meniul Help
Private Sub mnuTopHelpItem_Click(Index As Integer)
    Select Case Index
        Case 1
            On Error Resume Next
            Call HtmlHelp(Me.hwnd, App.Path & "\DBUtils.chm", 0, 0)
            If Err Then
                MsgBox Err.Description, vbCritical, "Help error"
            End If
            On Error GoTo 0
        Case 3
            frmAbout.Show vbModal, Me
    End Select
End Sub



' Meniurile Contextuale
Private Sub mnuTableItem_Click(Index As Integer)
    PopupID = Index
End Sub
Private Sub mnuViewItem_Click(Index As Integer)
    PopupID = 100 + Index
End Sub
Private Sub mnuProcItem_Click(Index As Integer)
    PopupID = 200 + Index
End Sub
Private Sub mnuFileItem_Click(Index As Integer)
    PopupID = 300 + Index
End Sub
Private Sub mnuSQLViewItem_Click(Index As Integer)
    PopupID = 400 + Index
End Sub

