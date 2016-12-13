VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public strDBFilename As String
Public strDBConnTemplate As String
Public objCon As ADODB.Connection
Public objCat As ADOX.Catalog

Public strFolderPath As String
Public arFilesExtensions As Variant
Public bIncludeHiddenFiles As Boolean
Public bIncludeBinaryFiles As Boolean
Public Files As Collection

Private Sub Class_Initialize()
    strDBFilename = ""
    strDBConnTemplate = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%MDB"
    
    strFolderPath = ""
    arFilesExtensions = Array("*")
    bIncludeHiddenFiles = False
    bIncludeBinaryFiles = False
    Set Files = New Collection
End Sub

Private Sub Class_Terminate()
    Set Files = Nothing
    Set objCat = Nothing
    If Not (objCon Is Nothing) Then
        If objCon.State = adStateOpen Then objCon.Close
    End If
    Set objCon = Nothing
End Sub

Public Function strConnectionString() As String
    strConnectionString = Replace(strDBConnTemplate, "%MDB", strDBFilename, 1, -1, vbTextCompare)
End Function

Public Function ConnectToDB() As Boolean
    On Error Resume Next
    Set objCon = New ADODB.Connection
    objCon.Open strConnectionString
    Set objCat = New ADOX.Catalog
    Set objCat.ActiveConnection = objCon
    ConnectToDB = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function ConnectToFiles() As Boolean
    On Error Resume Next
    Call BuildCatalog(strFolderPath, Files)
    ConnectToFiles = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub BuildCatalog(strParentFolder As String, colFiles As Collection)
    Dim fso As Scripting.FileSystemObject
    Dim fld As Scripting.Folder, subfld As Scripting.Folder
    Dim fil As Scripting.File
    Dim filebinary As Boolean, fileext As String, filename As String, datecreated As Date, datemodified As Date, filepath As String
    Dim bInclude As Boolean
    
    Set fso = New Scripting.FileSystemObject
    Set fld = fso.GetFolder(strParentFolder)
    For Each fil In fld.Files
        filename = fil.Name
        datecreated = fil.datecreated
        datemodified = fil.DateLastModified
        filepath = fil.Path
        fileext = fso.GetExtensionName(fil.Name)
        
        bInclude = True
        If Not bIncludeHiddenFiles Then
            If ((fil.Attributes And Hidden) = Hidden) Then bInclude = False
            If ((fil.Attributes And System) = System) Then bInclude = False
        End If
        If ((fil.Attributes And Directory) = Directory) Then bInclude = False
        If Not ExtInExtArray(fileext, arFilesExtensions) Then bInclude = False
        filebinary = FileIsBinary(filepath)
        If ((Not bIncludeBinaryFiles) And filebinary) Then bInclude = False

        If bInclude Then colFiles.Add Array(filename, datecreated, datemodified, "FILE", filepath, filebinary), filepath
    Next
    For Each subfld In fld.SubFolders
        bInclude = True
        If Not bIncludeHiddenFiles Then
            If ((subfld.Attributes And Hidden) = Hidden) Then bInclude = False
        End If
        If bInclude Then Call BuildCatalog(subfld.Path, colFiles)
    Next
    Set fld = Nothing
    Set fso = Nothing
End Sub

Private Function ExtInExtArray(strExt As String, arExtArray As Variant) As Boolean
    Dim e As Variant
    Dim le As String
    
    le = LCase(strExt)
    For Each e In arExtArray
        If (e = "*") Or (e = le) Then
            ExtInExtArray = True
            Exit Function
        End If
    Next
    ExtInExtArray = False
End Function

Public Sub SaveConnection(strName As String)
    Dim strSection As String
    strSection = dbuPersistentConnections & strName
    
    SaveSetting dbuAppName, strSection, "strDBFileName", strDBFilename
    SaveSetting dbuAppName, strSection, "strDBConnTemplate", strDBConnTemplate
    SaveSetting dbuAppName, strSection, "strFolderPath", strFolderPath
    SaveSetting dbuAppName, strSection, "arFilesExtensions", Join(arFilesExtensions, "|")
    SaveSetting dbuAppName, strSection, "bIncludeHiddenFiles", bIncludeHiddenFiles
    SaveSetting dbuAppName, strSection, "bIncludeBinaryFiles", bIncludeBinaryFiles
End Sub

Public Sub LoadConnection(strName As String)
    Dim strSection As String
    strSection = dbuPersistentConnections & strName
    
    strDBFilename = GetSetting(dbuAppName, strSection, "strDBFileName", strDBFilename)
    strDBConnTemplate = GetSetting(dbuAppName, strSection, "strDBConnTemplate", strDBConnTemplate)
    strFolderPath = GetSetting(dbuAppName, strSection, "strFolderPath", strFolderPath)
    arFilesExtensions = Split(GetSetting(dbuAppName, strSection, "arFilesExtensions", Join(arFilesExtensions, "|")), "|", -1, vbTextCompare)
    bIncludeHiddenFiles = GetSetting(dbuAppName, strSection, "bIncludeHiddenFiles", bIncludeHiddenFiles)
    bIncludeBinaryFiles = GetSetting(dbuAppName, strSection, "bIncludeBinaryFiles", bIncludeBinaryFiles)
End Sub


Public Function ShowWizardConnectionDialog() As VbMsgBoxResult
    Dim objWnd As frmBuildConnection
    Set objWnd = New frmBuildConnection
    
    objWnd.Show vbModal
    If Not objWnd.bCancelPressed Then
        strDBFilename = objWnd.strDBFilename
        strDBConnTemplate = objWnd.strDBConTemplate
        strFolderPath = objWnd.strFolderPath
        arFilesExtensions = objWnd.arFilesExtensions
        bIncludeHiddenFiles = objWnd.bIncludeHiddenFiles
        bIncludeBinaryFiles = objWnd.bIncludeBinaryFiles
        ShowWizardConnectionDialog = vbOK
    Else
        ShowWizardConnectionDialog = vbCancel
    End If
    
    Set objWnd = Nothing
End Function


Public Function ShowCreateConnectionDialog() As VbMsgBoxResult
    Dim objWnd As frmGetConnection
    Set objWnd = New frmGetConnection
    
    objWnd.Show vbModal
    If Not objWnd.bCancelPressed Then
        strDBFilename = objWnd.strDBFilename
        strDBConnTemplate = objWnd.strDBConTemplate
        strFolderPath = objWnd.strFolderPath
        arFilesExtensions = objWnd.arFilesExtensions
        bIncludeHiddenFiles = objWnd.bIncludeHiddenFiles
        bIncludeBinaryFiles = objWnd.bIncludeBinaryFiles
        ShowCreateConnectionDialog = vbOK
    Else
        ShowCreateConnectionDialog = vbCancel
    End If
    
    Set objWnd = Nothing
End Function

Public Function CountObjects(strObjType As String) As Long
    Dim re As Long
    Dim t As ADOX.Table
    
    Select Case strObjType
        Case dbuTypeTable
            re = 0
            For Each t In objCat.Tables
                If t.Type = dbuTypeTable Then re = re + 1
            Next
        Case dbuTypeView
            re = objCat.Views.Count
        Case dbuTypeProc
            re = objCat.Procedures.Count
        Case dbuTypeFile
            re = Files.Count
    End Select
    CountObjects = re
End Function
