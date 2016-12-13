VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPersistentSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
                           
Public Sub SaveWindowPosition(objWnd As Object)
    Dim strSection As String
    If objWnd.WindowState <> vbMinimized Then
        With objWnd
            strSection = dbuPersistentWindows & .Name
            SaveSetting dbuAppName, strSection, "Left", .Left
            SaveSetting dbuAppName, strSection, "Top", .Top
            SaveSetting dbuAppName, strSection, "Width", .Width
            SaveSetting dbuAppName, strSection, "Height", .Height
        End With
    End If
End Sub

Public Sub LoadWindowPosition(objWnd As Object)
    Dim strSection As String
    With objWnd
        strSection = dbuPersistentWindows & .Name
        .Left = GetSetting(dbuAppName, strSection, "Left", 1000)
        .Top = GetSetting(dbuAppName, strSection, "Top", 1000)
        .Width = GetSetting(dbuAppName, strSection, "Width", 6500)
        .Height = GetSetting(dbuAppName, strSection, "Height", 6500)
    End With
End Sub

Public Sub DeleteConnection(strName As String)
    On Error Resume Next
    Call RegDeleteKey(HKEY_CURRENT_USER, dbuVBReg & dbuAppName & "\" & dbuPersistentConnections & strName)
    On Error GoTo 0
End Sub

Public Function GetAllConnections() As Variant
    Dim s As String
    s = dbuVBReg & dbuAppName & "\" & dbuPersistentConnections
    GetAllConnections = EnumRegistryKeys(HKEY_CURRENT_USER, s)
End Function
