Attribute VB_Name = "modWindows"
Option Explicit

Public PopupID As Integer

' Se ocupa cu Dispatch-ul actiunilor selectate din meniurile contextuale in functie de actiunea intMenuIndex si de
' numele si tipul obiectului pentru care s-a apelat meniul contextual. Optional pot fi transmise informatii suplimentare prin variantul Tag
Public Sub HandleContextMenu(ByVal intMenuIndex As Integer, ByVal strItemName As String, ByVal strItemType As String, Optional ByVal Tag As Variant)
    Dim s As String
    
    If intMenuIndex = -1 Then Exit Sub
    If (fMainForm.ActiveForm Is Nothing) Then Exit Sub
    
    With fMainForm.ActiveForm.objDBCon
        Select Case intMenuIndex
            ' Meniul contextual pentru Tables
            Case 1: Call ShowTableStructure(strItemName)
            Case 2: Call ShowTableViewContent(strItemName, strItemType)
            Case 3: Call ShowWhoUseObject(strItemName, strItemType)
            Case 5: Call ShowSearchText(strItemName, 4)
            'Meniul contextual pentru Views
            Case 101: Call ShowObjectCode(strItemName, dbuTypeView)
            Case 102: Call ShowTableViewContent(strItemName, strItemType)
            Case 103: Call ShowWhoUseObject(strItemName, strItemType)
            Case 104: Call ShowWhatUseObject(strItemName, strItemType)
            Case 106: Call ShowSearchText(strItemName, 4)
            ' Meniul contextual pentru Procedures
            Case 201: Call ShowObjectCode(strItemName, dbuTypeProc)
            Case 202: Call ShowWhoUseObject(strItemName, strItemType)
            Case 203: Call ShowWhatUseObject(strItemName, strItemType)
            Case 205: Call ShowSearchText(strItemName, 4)
            ' Meniul contextual pentru Files
            Case 301: Call ShellExecute(fMainForm.hwnd, "open", strItemName, 0, 0, 1)
            Case 302: Call Shell("Notepad " & strItemName, vbNormalFocus)
            Case 304: Call ShowDBObjectsInFile(strItemName)
            ' Meniul contextual pentru RTF-ul cu View SQL Code
            Case 401
                Clipboard.Clear
                Clipboard.SetText Tag(0), vbCFText
            Case 402
                Clipboard.Clear
                Clipboard.SetText Tag(1), vbCFRTF
            Case 404
                s = Tag(0)
                Do While Right(s, Len(vbCrLf)) = vbCrLf
                    s = Left(s, Len(s) - Len(vbCrLf))
                Loop
                Call ShowSearchText(Trim(s))
        End Select
    End With
End Sub

' Afiseaza in fereastra objWnd un meniu contextual, potrivit pentru un obiect de tipul strObjType
' Suplimentar mai pot fi pasate informatii in variantul Tag
Public Function ShowPopupMenu(ByVal objWnd As Form, ByVal strObjType As String, Optional Tag As Variant) As Integer
    Dim bExistsFiles As Boolean
    Dim bFileIsBin As Boolean
    
    With fMainForm
        If (.ActiveForm Is Nothing) Then Exit Function
        bExistsFiles = (Len(.ActiveForm.objDBCon.strFolderPath) > 0)
        PopupID = -1
        Select Case strObjType
            Case dbuTypeTable
                .mnuTableItem(5).Enabled = bExistsFiles
                objWnd.PopupMenu .mnuTableMenu
            Case dbuTypeView
                .mnuViewItem(6).Enabled = bExistsFiles
                objWnd.PopupMenu .mnuViewMenu
            Case dbuTypeProc
                .mnuProcItem(5).Enabled = bExistsFiles
                objWnd.PopupMenu .mnuProcMenu
            Case dbuTypeFile
                If IsMissing(Tag) Then
                    bFileIsBin = True
                Else
                    bFileIsBin = CBool(Tag)
                End If
                .mnuFileItem(2).Enabled = Not bFileIsBin
                .mnuFileItem(4).Enabled = Not bFileIsBin
                objWnd.PopupMenu .mnuFileMenu
            Case "RTFSELECTION"
                .mnuSQLViewItem(1).Enabled = (Len(Tag(0)) > 0)
                .mnuSQLViewItem(2).Enabled = (Len(Tag(1)) > 0)
                .mnuSQLViewItem(4).Enabled = (Len(Tag(0)) > 0)
                objWnd.PopupMenu .mnuSQLView
        End Select
    End With
    ShowPopupMenu = PopupID
End Function

Public Sub ShowObjectCode(ByVal strObjectName As String, ByVal strObjectType As String)
    With New frmViewCode
        .strInitialObject = strObjectName
        .strInitialObjectType = strObjectType
        .Show vbModal
    End With
End Sub

Public Sub ShowTableStructure(ByVal strObjectName As String)
    With New frmViewTable
        .strInitialObject = strObjectName
        .Show vbModal
    End With
End Sub

Public Sub ShowTableViewContent(ByVal strObjectName As String, ByVal strObjectType As String)
    With New frmViewContents
        .strInitialObject = strObjectName
        .strInitialObjectType = strObjectType
        On Error Resume Next
        .Show vbModeless
        If Err.Number <> 0 Then
            Err.Clear
            .Show vbModal
        End If
        On Error GoTo 0
    End With
End Sub

Public Sub ShowWhoUseObject(ByVal strObjectName As String, ByVal strObjectType As String)
    With New frmWhoUseObject
        .strInitialObject = strObjectName
        .strInitialObjectType = strObjectType
        Screen.MousePointer = vbHourglass
        .Show vbModal
    End With
End Sub

Public Sub ShowWhatUseObject(ByVal strObjectName As String, ByVal strObjectType As String)
    With New frmWhatUseObject
        .strInitialObject = strObjectName
        .strInitialObjectType = strObjectType
        Screen.MousePointer = vbHourglass
        .Show vbModal
    End With
End Sub

' strTextToFind = textul care se cauta la deschiderea ferestrei.. daca lipseste se asteapta introducerea unui text de cautat
' lngSearchIn = specifica unde se cauta prin bitii setati: b2b1b0 - b2 = Files, b1 = Queries, b0 = Tables (implicit in queries si tables)
Public Sub ShowSearchText(Optional ByVal strTextToFind As String = "", Optional ByVal lngSearchIn As Long = 3)
    With New frmSearchText
        .strInitialText = strTextToFind
        .lngSearchIn = lngSearchIn
        Screen.MousePointer = vbHourglass
        .Show vbModal
    End With
End Sub

Public Sub ShowDBObjectsInFile(ByVal strFilename As String)
    With New frmDBObjectsInFile
        .strInitialFileName = strFilename
        Screen.MousePointer = vbHourglass
        .Show vbModal
    End With
End Sub

Public Sub ShowDBProperties()
    With New frmDBProperties
        .Show vbModal
    End With
End Sub
