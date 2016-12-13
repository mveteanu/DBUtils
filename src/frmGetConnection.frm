VERSION 5.00
Begin VB.Form frmGetConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VMASOFT DBUtils"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmGetConnection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select database connection"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ListBox lstSavedCon 
         Height          =   1425
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   3735
      End
      Begin VB.OptionButton chkLoad 
         Caption         =   "Use a saved connection"
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Width           =   3135
      End
      Begin VB.OptionButton chkBuild 
         Caption         =   "Build a new connection"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "frmGetConnection.frx":0442
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmGetConnection.frx":133C
         Top             =   1320
         Width           =   480
      End
   End
   Begin VB.Menu mnuSavedConnections 
      Caption         =   "Saved Connections"
      Visible         =   0   'False
      Begin VB.Menu mnuSavedConnectionsItems 
         Caption         =   "Detele selected connection"
      End
   End
End
Attribute VB_Name = "frmGetConnection"
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

Private Sub btnOK_Click()
    Dim f As frmBuildConnection
    Dim c As clsDBConnection
    
    If chkBuild.Value Then
        Set f = New frmBuildConnection
        With f
            .Show vbModal
            strDBFilename = .strDBFilename
            strDBConTemplate = .strDBConTemplate
            strFolderPath = .strFolderPath
            arFilesExtensions = .arFilesExtensions
            bIncludeHiddenFiles = .bIncludeHiddenFiles
            bIncludeBinaryFiles = .bIncludeBinaryFiles
            bCancelPressed = .bCancelPressed
        End With
        Set f = Nothing
        Call FillListWithSavedCons
    ElseIf chkLoad.Value Then
        If lstSavedCon.ListIndex <> -1 Then
            Set c = New clsDBConnection
            c.LoadConnection lstSavedCon.List(lstSavedCon.ListIndex)
            With c
                strDBFilename = .strDBFilename
                strDBConTemplate = .strDBConnTemplate
                strFolderPath = .strFolderPath
                arFilesExtensions = .arFilesExtensions
                bIncludeHiddenFiles = .bIncludeHiddenFiles
                bIncludeBinaryFiles = .bIncludeBinaryFiles
                bCancelPressed = False
            End With
            Set c = Nothing
        End If
    End If
    If Not bCancelPressed Then Unload Me
End Sub

Private Sub FillListWithSavedCons()
    Dim r As clsPersistentSettings, ar As Variant, arit As Variant
    
    Set r = New clsPersistentSettings
    ar = r.GetAllConnections
    If VarType(ar) <> vbEmpty Then
        lstSavedCon.Clear
        For Each arit In ar
            lstSavedCon.AddItem arit
        Next
    End If
    Set r = Nothing
End Sub


Private Sub Form_Load()
    chkBuild.Value = True
    bCancelPressed = True
    Call FillListWithSavedCons
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnCancel_Click
End Sub

Private Sub lstSavedCon_Click()
    chkLoad.Value = True
End Sub

Private Sub lstSavedCon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbRightButton Then Me.PopupMenu mnuSavedConnections
End Sub

Private Sub mnuSavedConnectionsItems_Click()
    Dim r As clsPersistentSettings
    Dim s As String
    
    If lstSavedCon.ListIndex = -1 Then Exit Sub
    s = lstSavedCon.List(lstSavedCon.ListIndex)
    If MsgBox("Are you sure you want to delete the settings for" & vbCrLf & s, vbQuestion + vbYesNo, "Delete connection") = vbYes Then
        Set r = New clsPersistentSettings
        r.DeleteConnection s
        Set r = Nothing
        Call FillListWithSavedCons
    End If
End Sub
