VERSION 5.00
Begin VB.Form frmDBProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database properties"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "frmDBProperties.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Files"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   5655
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lbl 
         Caption         =   "Files loaded:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Folder:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtView 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lbl 
         Caption         =   "Procedures:"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Views:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Queries:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Tables:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "Connection String:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl 
         Caption         =   "File Name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmDBProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objConLocal As clsDBConnection

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
    Dim lngViews As Long
    Dim lngProcs As Long
    
    If fMainForm.ActiveForm Is Nothing Then
        Unload Me
    Else
        Set objConLocal = fMainForm.ActiveForm.objDBCon
    End If
    
    With objConLocal
        txtView(0).Text = .strDBFilename
        txtView(1).Text = .strConnectionString
        txtView(2).Text = .CountObjects(dbuTypeTable)
        lngViews = .CountObjects(dbuTypeView)
        lngProcs = .CountObjects(dbuTypeProc)
        txtView(3).Text = lngViews + lngProcs
        txtView(4).Text = lngViews
        txtView(5).Text = lngProcs
        txtView(6).Text = .strFolderPath
        txtView(7).Text = .CountObjects(dbuTypeFile)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objConLocal = Nothing
End Sub

