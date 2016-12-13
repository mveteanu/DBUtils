VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuildConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Build connection"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmBuildConnection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmBuildConnection"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTab 
      Height          =   3255
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   4000
      Begin VB.ComboBox txtFileExt 
         Height          =   315
         ItemData        =   "frmBuildConnection.frx":0442
         Left            =   120
         List            =   "frmBuildConnection.frx":0444
         TabIndex        =   18
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CheckBox chkIncBinary 
         Caption         =   "Include binary files"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CheckBox chkIncHidden 
         Caption         =   "Include hidden files and folders"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CommandButton btnBrowseForCode 
         Caption         =   "..."
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   480
         Width           =   315
      End
      Begin VB.TextBox txtCodeFolder 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Miscellaneous:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "File extensions:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Source Code Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame FrameTab 
      Height          =   3255
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Width           =   4000
      Begin VB.TextBox txtDBFile 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton btnBrowseForDB 
         Caption         =   "..."
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         Top             =   480
         Width           =   315
      End
      Begin VB.TextBox txtDBConn 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Text            =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%MDB"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "JET Database File:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Connection string template:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   3615
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Database connection"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Project source code"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgBox 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "frmBuildConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strDBFilename As String
Public strDBConTemplate As String
Public strFolderPath As String
Public arFilesExtensions As Variant
Public bIncludeHiddenFiles As Boolean
Public bIncludeBinaryFiles As Boolean
Public bCancelPressed As Boolean

' Intoarce calea catre un folder urmata de "\"
' Exemplu: msgbox BrowseForFolder(Me.hWnd, "Alegeti un director", BIF_OnlyIncludeDirs, BSF_MyComputer)
Private Function BrowseForFolder(strTitle As String, intOptions As Long, strRootPath, Optional lngHWnd As Long = 0) As String
    Dim objShell As Shell32.Shell, objFolder As Shell32.Folder
    Dim re As String, p As Long
    
    Set objShell = New Shell32.Shell
    Set objFolder = objShell.BrowseForFolder(lngHWnd, strTitle, intOptions, strRootPath)
    On Error Resume Next
    re = objFolder.ParentFolder.ParseName(objFolder.Title).Path & "\"
    If Err.Number <> 0 Then
        Err.Clear
        re = objFolder.Title
        p = InStr(re, ":")
        If p > 0 Then re = Mid(re, p - 1, 2) & "\"
    End If
    On Error GoTo 0
    Set objFolder = Nothing
    Set objShell = Nothing
    BrowseForFolder = CStr(re)
End Function

Private Sub btnBrowseForCode_Click()
    Const BIF_OnlyIncludeDirs As Long = &H1
    Const BSF_MyComputer = 17
    Dim strFld As String

    strFld = BrowseForFolder("Source code folder:", BIF_OnlyIncludeDirs, BSF_MyComputer, Me.hwnd)
    If Len(strFld) <> 0 Then txtCodeFolder.Text = strFld
End Sub

Private Sub btnBrowseForDB_Click()
    With dlgBox
        .DialogTitle = "Open database"
        .Filter = "Microsoft Access Databases (*.mdb)|*.mdb"
        .ShowOpen
        txtDBFile.Text = .filename
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function FileExtArray(strFileExt) As String()
    Dim re As Variant
    Dim re2() As String, re2ubound As Integer
    Dim i As Integer, p As Integer
    re = Split(strFileExt, ";", -1, vbTextCompare)
    For i = LBound(re) To UBound(re)
        p = InStr(1, re(i), ".", vbTextCompare)
        If p > 0 Then
            re(i) = Trim(Right(re(i), Len(re(i)) - p))
            re2ubound = re2ubound + 1
        Else
            re(i) = ""
        End If
    Next
    ReDim re2(re2ubound)
    p = 0
    For i = LBound(re) To UBound(re)
        If Len(re(i)) > 0 Then
            re2(p) = CStr(re(i))
            p = p + 1
        End If
    Next
    FileExtArray = re2
End Function

Private Function ValidConnection() As Boolean
    Dim re As Boolean
    re = False
    If Not FileExists(txtDBFile.Text) Then
        MsgBox "Invalid database file name.", vbOKOnly + vbCritical, "DBUtils"
    ElseIf Not TestConnectionString(Replace(txtDBConn.Text, "%MDB", txtDBFile.Text, , , vbTextCompare)) Then
        MsgBox "Invalid connection.", vbOKOnly + vbCritical, "DBUtils"
    ElseIf (Len(txtCodeFolder.Text) > 0) And (Not FolderExists(txtCodeFolder.Text)) Then
        MsgBox "Invalid source code folder.", vbOKOnly + vbCritical, "DBUtils"
    ElseIf (UBound(FileExtArray(txtFileExt.Text)) = 0) Then
        MsgBox "No file extension specified.", vbOKOnly + vbCritical, "DBUtils"
    Else
        re = True
    End If
    ValidConnection = re
End Function

Private Sub btnOK_Click()
    If ValidConnection Then
        strDBFilename = txtDBFile.Text
        strDBConTemplate = txtDBConn.Text
        strFolderPath = txtCodeFolder.Text
        arFilesExtensions = FileExtArray(txtFileExt.Text)
        bIncludeHiddenFiles = (chkIncHidden.Value = 1)
        bIncludeBinaryFiles = (chkIncBinary.Value = 1)
        bCancelPressed = False
        Unload Me
    End If
End Sub

Private Sub btnSave_Click()
    Dim r As clsDBConnection
    Dim s As String
    If ValidConnection Then
        s = InputBox("Enter a name to identify later this connection", "Save connection parameters")
        If Len(Trim(s)) = 0 Then Exit Sub
        Set r = New clsDBConnection
        With r
            .strDBFilename = txtDBFile.Text
            .strDBConnTemplate = txtDBConn.Text
            .strFolderPath = txtCodeFolder.Text
            .arFilesExtensions = FileExtArray(txtFileExt.Text)
            .bIncludeHiddenFiles = (chkIncHidden.Value = 1)
            .bIncludeBinaryFiles = (chkIncBinary.Value = 1)
            .SaveConnection s
        End With
        Set r = Nothing
    End If
End Sub


Private Sub Form_Activate()
    txtDBFile.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnCancel_Click
End Sub

Private Sub Form_Load()
    Dim fra As Frame
    bCancelPressed = True
    With txtFileExt
        .AddItem "*.*"
        .AddItem "*.asp;*.inc;*.vbs"
        .AddItem "*.frm;*.bas;*.cls"
        .AddItem "*.dfm;*.pas"
        .ListIndex = 0
    End With
    For Each fra In FrameTab
        fra.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, _
            TabStrip1.ClientWidth, TabStrip1.ClientHeight
        fra.BorderStyle = 0
    Next
    Call TabStrip1_Click
End Sub

Private Sub TabStrip1_Click()
    Dim fra As Frame
    For Each fra In FrameTab
        fra.Visible = (fra.index = TabStrip1.SelectedItem.index - 1)
    Next
End Sub
