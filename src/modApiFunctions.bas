Attribute VB_Name = "modApiFunctions"
Option Explicit

Public Const HKEY_CURRENT_USER = &H80000001

Private Const READ_CONTROL = &H20000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" _
    Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, ByVal cbName As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As _
     Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) _
    As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Public Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As Variant
    Dim handle As Long, Index As Long, length As Long
    ReDim result(0 To 100) As String

    If Len(KeyName) > 0 Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
            Exit Function
        End If
        hKey = handle
    End If
    
    For Index = 0 To 999999
        If Index > UBound(result) Then
            ReDim Preserve result(Index + 99) As String
        End If
        length = 260                   ' Max length for a key name.
        result(Index) = Space$(length)
        If RegEnumKey(hKey, Index, result(Index), length) Then Exit For
        result(Index) = Left$(result(Index), InStr(result(Index), vbNullChar) - 1)
    Next

    If handle Then RegCloseKey handle
    If Index > 0 Then
        ReDim Preserve result(Index - 1) As String
        EnumRegistryKeys = result()
    End If
End Function

