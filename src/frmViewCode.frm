VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Code Viewer"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmViewCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmViewCode.frx":0442
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
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
            Picture         =   "frmViewCode.frx":04C4
            Key             =   "IcoTable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":061E
            Key             =   "IcoView"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmViewCode.frx":0778
            Key             =   "IcoProc"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "DB Queries:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "SQL Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmViewCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strInitialObject As String
Public strInitialObjectType As String

Private objCatLocal As ADOX.Catalog
Private FirstTime As Boolean

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cboObjects_Click()
    If FirstTime Then Exit Sub
    FillFormForListIndex cboObjects.SelectedItem.Index
End Sub

Private Sub Form_Activate()
    btnClose.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call btnClose_Click
End Sub

Private Sub Form_Load()
    Dim v As ADOX.View
    Dim p As ADOX.Procedure
    Dim i As Integer
    
    Set objCatLocal = fMainForm.ActiveForm.objDBCon.objCat
    
    For Each v In objCatLocal.Views
        With cboObjects.ComboItems.Add(, "v" & v.Name, v.Name, DBObjectIcon(dbuTypeView))
            If (strInitialObjectType = dbuTypeView) And (v.Name = strInitialObject) Then i = .Index
        End With
    Next
    For Each p In objCatLocal.Procedures
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
            Case "v": strType = dbuTypeView
            Case "p": strType = dbuTypeProc
        End Select
        .Selected = True
    End With
    FirstTime = False
    With txtCode
        .Text = GetObjectText(objCatLocal, strName, strType)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = vbBlack
    End With
    HilightSyntax txtCode, Array("PARAMETERS", "SELECT", "INSERT", "UPDATE", "DELETE", "FROM", "INTO", "VALUES", "SET", "WHERE", "HAVING", "GROUP BY", "ORDER BY", "ASC", "DESC", "INNER JOIN", "LEFT JOIN", "RIGHT JOIN", "OUTER", "ON", "AS", "UNION", "AND", "OR", "NOT", "IN", "BETWEEN", "ALL", "DISTINCT", "DISTINCTROW", "LIKE", "IS", "NULL"), vbBlue
    HilightSyntax txtCode, Array("SUM", "COUNT", "AVG", "First", "Last", "Max", "Min", "Var", "STDEV", "Now", "DateDiff", "Round", "IIf", "DateAdd", "CDate", "CBool", "CStr", "CLng", "CInt", "Len", "InStr", "Exists"), vbMagenta
    HilightSyntax txtCode, Array("Bit", "Byte", "Short", "Long", "Currency", "IEEESingle", "IEEEDouble", "DateTime", "Binary", "Text", "LongBinary", "Guid", "Value"), vbRed
End Sub

Private Sub HilightSyntax(objRTF As RichTextBox, arKeyWords, KeyWordsColor)
    Dim strKeyWord As Variant
    Dim posFound As Integer
    
    With objRTF
        For Each strKeyWord In arKeyWords
            .SelStart = 0
            .SelLength = 0
            posFound = 0
            Do
                posFound = .Find(strKeyWord, posFound, , rtfWholeWord) ' + rtfMatchCase
                If posFound = -1 Then Exit Do
                    .SelColor = KeyWordsColor
                    posFound = posFound + Len(strKeyWord)
                Loop
        Next
        .SelLength = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsObject(objCatLocal) Then
        Set objCatLocal = Nothing
    End If
End Sub


Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a As Variant
    
    If Button And vbRightButton Then
        a = Array(txtCode.SelText, txtCode.SelRTF)
        Call HandleContextMenu(ShowPopupMenu(Me, "RTFSELECTION", a), "", "", a)
    End If
End Sub
