Attribute VB_Name = "modDBUtils"
Option Explicit

Public Function DBObjectIcon(ByVal strObjType As String) As String
    Dim re As String
    Select Case strObjType
        Case dbuTypeTable: re = "IcoTable"
        Case dbuTypeView:  re = "IcoView"
        Case dbuTypeProc:  re = "IcoProc"
        Case dbuTypeFile:  re = "IcoFile"
    End Select
    DBObjectIcon = re
End Function

Public Function DBIconOf(ByVal strIconType As String) As String
    Dim re As String
    Select Case strIconType
        Case "IcoTable": re = dbuTypeTable
        Case "IcoView":  re = dbuTypeView
        Case "IcoProc":  re = dbuTypeProc
        Case "IcoFile":  re = dbuTypeFile
    End Select
    DBIconOf = re
End Function

Public Function GetObjectText(ByVal objDBCat As ADOX.Catalog, ByVal strObjectName As String, ByVal strType As String) As String
    Dim re As String
    Select Case strType
        Case dbuTypeView
            re = objDBCat.Views(strObjectName).Command.CommandText
        Case dbuTypeProc
            re = objDBCat.Procedures(strObjectName).Command.CommandText
    End Select
    GetObjectText = re
End Function

Public Function DataTypeEnum(t As Long) As String
    Dim re As String
    Select Case t
        Case adBigInt: re = "DBTYPE_I8"
        Case adBinary: re = "DBTYPE_BYTES"
        Case adBoolean: re = "Yes/No"
        Case adBSTR: re = "DBTYPE_BSTR"
        Case adChapter: re = "DBTYPE_HCHAPTER"
        Case adChar: re = "DBTYPE_STR"
        Case adCurrency: re = "DBTYPE_CY"
        Case adDate: re = "Date/Time"
        Case adDBDate: re = "DBTYPE_DBDATE"
        Case adDBTime: re = "DBTYPE_DBTIME"
        Case adDBTimeStamp: re = "DBTYPE_DBTIMESTAMP"
        Case adDecimal: re = "DBTYPE_DECIMAL"
        Case adDouble: re = "Double"
        Case adEmpty: re = "DBTYPE_EMPTY"
        Case adError: re = "DBTYPE_ERROR"
        Case adFileTime: re = "DBTYPE_FILETIME"
        Case adGUID: re = "DBTYPE_GUID"
        Case adIDispatch: re = "DBTYPE_IDISPATCH"
        Case adInteger: re = "Long Integer"
        Case adIUnknown: re = "DBTYPE_IUNKNOWN"
        Case adNumeric: re = "DBTYPE_NUMERIC"
        Case adPropVariant: re = "DBTYPE_PROP_VARIANT"
        Case adSingle: re = "Single"
        Case adSmallInt: re = "Integer"
        Case adTinyInt: re = "DBTYPE_I1"
        Case adUnsignedBigInt: re = "DBTYPE_UI8"
        Case adUnsignedInt: re = "DBTYPE_UI4"
        Case adUnsignedSmallInt: re = "DBTYPE_UI2"
        Case adUnsignedTinyInt: re = "Byte"
        Case adUserDefined: re = "DBTYPE_UDT"
        Case adVariant: re = "DBTYPE_VARIANT"
        Case adWChar: re = "DBTYPE_WSTR"
        Case adLongVarBinary: re = "Long Binary Data"
        Case adLongVarChar: re = "adLongVarChar"
        Case adLongVarWChar: re = "Memo"
        Case adVarBinary: re = "adVarBinary"
        Case adVarChar: re = "adVarChar"
        Case adVarNumeric: re = "adVarNumeric"
        Case adVarWChar: re = "Text"
        Case Else: re = CStr(t)
    End Select
    DataTypeEnum = re
End Function

Public Function TestConnectionString(strCon As String) As Boolean
    Dim objCon As ADODB.Connection
    On Error Resume Next
    Set objCon = New ADODB.Connection
    objCon.Open strCon
    If objCon.State = adStateOpen Then objCon.Close
    Set objCon = Nothing
    TestConnectionString = CBool(Err.Number = 0)
    On Error GoTo 0
End Function

