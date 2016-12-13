VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search text in database objects"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmSearchText.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Search results"
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   7215
      Begin MSComctlLib.ListView lstObjects 
         Height          =   3495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6165
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
         NumItems        =   3
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
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comments"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Text to search"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.CheckBox chkSearchFiles 
         Caption         =   "Search in text Files"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox chkSearchQueries 
         Caption         =   "Search in Queries"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkSearchTables 
         Caption         =   "Search in structure of Tables"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox txtWords 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Type here the text you want to search. You can use AND and NOT operators to limit your search results."
         Top             =   240
         Width           =   6015
      End
      Begin VB.CommandButton btnFind 
         Caption         =   "&Find"
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkWordsOnly 
         Caption         =   "Words only"
         Height          =   255
         Left            =   4920
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   6000
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
            Picture         =   "frmSearchText.frx":0442
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchText.frx":059C
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchText.frx":06F6
            Key             =   "IcoProc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearchText.frx":0850
            Key             =   "IcoFile"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearchText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strInitialText As String
Public lngSearchIn As Long

Private SQLObjects As New Collection
Private objConLocal As clsDBConnection

Const commPrefix As String = "Found in "
Const commTable As String = "table definition"
Const commQuery As String = "SQL code"
Const commFile As String = "text file "

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnFind_Click()
    Call FillWithSearchResults(Trim(txtWords.Text))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnClose_Click
End Sub

Private Sub Form_Load()
    Dim t As ADOX.Table
    Dim v As ADOX.View
    Dim p As ADOX.Procedure
    Dim f As Variant
    Dim s As String, tcol As ADOX.Column
    
    Set objConLocal = fMainForm.ActiveForm.objDBCon
    
    Me.Caption = "Search text in database (" & fMainForm.ActiveForm.Caption & ")"
    If (lngSearchIn And 4) = 4 Then chkSearchFiles.Value = 1 Else chkSearchFiles.Value = 0
    If (lngSearchIn And 2) = 2 Then chkSearchQueries.Value = 1 Else chkSearchQueries.Value = 0
    If (lngSearchIn And 1) = 1 Then chkSearchTables.Value = 1 Else chkSearchTables.Value = 0
    
    For Each t In objConLocal.objCat.Tables
        DoEvents
        If t.Type = dbuTypeTable Then
            s = t.Name
            For Each tcol In t.Columns
                s = s & " " & tcol.Name & " (" & DataTypeEnum(tcol.Type) & ")"
            Next
            SQLObjects.Add Array(t.Name, dbuTypeTable, s, commTable), "t" & t.Name
        End If
    Next
    For Each v In objConLocal.objCat.Views
        DoEvents
        SQLObjects.Add Array(v.Name, dbuTypeView, GetObjectText(objConLocal.objCat, v.Name, dbuTypeView), commQuery), "v" & v.Name
    Next
    For Each p In objConLocal.objCat.Procedures
        DoEvents
        SQLObjects.Add Array(p.Name, dbuTypeProc, GetObjectText(objConLocal.objCat, p.Name, dbuTypeProc), commQuery), "p" & p.Name
    Next
    If Len(objConLocal.strFolderPath) > 0 Then
        For Each f In objConLocal.Files
            DoEvents
            If Not f(5) Then
                SQLObjects.Add Array(f(0), dbuTypeFile, f(4), commFile & f(4)), "f" & f(4)
            End If
        Next
    Else
        chkSearchFiles.Value = 0
        chkSearchFiles.Enabled = False
    End If
    txtWords.Text = strInitialText
    FillWithSearchResults strInitialText
End Sub

Private Sub FillWithSearchResults(strWords)
    Dim v As Variant
    Dim s As String
    Dim b As Boolean
    
    lstObjects.ListItems.Clear
    If Len(Trim(strWords)) = 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    btnFind.Enabled = False
    
    For Each v In SQLObjects
        s = v(2)
        Select Case v(1)
            Case dbuTypeTable
                b = (chkSearchTables.Value = 1)
            Case dbuTypeView
                b = (chkSearchQueries.Value = 1)
            Case dbuTypeProc
                b = (chkSearchQueries.Value = 1)
            Case dbuTypeFile
                b = (chkSearchFiles.Value = 1)
                s = GetTextFileContent(s)
        End Select
        
        If b Then
            If WordsInsideText(strWords, s, (chkWordsOnly.Value = 1)) Then
                With lstObjects.ListItems.Add(, , v(0), , DBObjectIcon(v(1)))
                    .ListSubItems.Add , , v(1)
                    .ListSubItems.Add , , commPrefix & v(3)
                End With
            End If
        End If
    Next
    
    btnFind.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SQLObjects = Nothing
    Set objConLocal = Nothing
End Sub

Private Sub Form_Activate()
    txtWords.SetFocus
End Sub

Private Sub txtWords_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call btnFind_Click
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
        
        If strType = dbuTypeFile Then
            strName = LI.ListSubItems(2).Text
            strName = Right(strName, Len(strName) - Len(commFile) - Len(commPrefix))
        End If
        Call HandleContextMenu(ShowPopupMenu(Me, strType, False), strName, strType)
    End If
End Sub

