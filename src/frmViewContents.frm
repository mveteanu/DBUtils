VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmViewContents 
   Caption         =   "Contents of: "
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   Icon            =   "frmViewContents.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid DG 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7011
      _Version        =   393216
      BackColorBkg    =   -2147483636
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmViewContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strInitialObject As String
Public strInitialObjectType As String

Private objLocalCon As ADODB.Connection
Private objRs As ADODB.Recordset

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Err
    Me.Caption = "Contents of " & strInitialObjectType & ": " & strInitialObject
    
    Set objLocalCon = fMainForm.ActiveForm.objDBCon.objCon
    
    Set objRs = New ADODB.Recordset
    Set objRs.ActiveConnection = objLocalCon
    objRs.Open strInitialObject, , adOpenStatic, adLockOptimistic, adCmdTableDirect
 
    Set DG.DataSource = objRs
    
    Exit Sub
Err:
    Call MsgBox("Error obtaining data for selected table/view!", vbCritical + vbOKOnly, "DBUtils")
End Sub

Private Sub Form_Resize()
    DG.Width = Me.ScaleWidth
    DG.Height = Me.ScaleHeight
    DG.ColWidth(0) = DG.RowHeight(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRs = Nothing
    Set objLocalCon = Nothing
End Sub
