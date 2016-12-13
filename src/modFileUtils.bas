Attribute VB_Name = "modFileUtils"
Option Explicit

Public Function FileExists(strFilename As String) As Boolean
    Dim objFS As Scripting.FileSystemObject
    Set objFS = New Scripting.FileSystemObject
    FileExists = objFS.FileExists(strFilename)
    Set objFS = Nothing
End Function

Public Function FolderExists(strFoldername As String) As Boolean
    Dim objFS As Scripting.FileSystemObject
    Set objFS = New Scripting.FileSystemObject
    FolderExists = objFS.FolderExists(strFoldername)
    Set objFS = Nothing
End Function

' Un fisier se considera ca este binar daca in continutul sau
' apar bytes de 0x00. Citirea se face in blocuri din motive de viteza.
Public Function FileIsBinary(strFilename As String) As Boolean
    Const BufLength As Long = 4096
    Dim bIsBin As Boolean
    Dim Buf(0 To BufLength - 1) As Byte
    Dim i As Long, J As Long
    Dim lngFile As Long, lngBlocks As Long, lngRest As Long
    
    bIsBin = False
    Open strFilename For Binary Access Read Shared As #1
    lngFile = LOF(1)
    lngBlocks = lngFile \ BufLength
    lngRest = lngFile Mod BufLength
    
    For J = 1 To lngBlocks
        DoEvents
        If bIsBin Then Exit For
        Get #1, ((J - 1) * BufLength) + 1, Buf
        For i = 0 To UBound(Buf)
            If Buf(i) = 0 Then
                bIsBin = True
                Exit For
            End If
        Next
    Next
    If Not bIsBin Then
        Get #1, (lngBlocks * BufLength) + 1, Buf
        For i = 0 To lngRest - 1
            If Buf(i) = 0 Then
                bIsBin = True
                Exit For
            End If
        Next
    End If
    Close #1
    FileIsBinary = bIsBin
End Function

' Intoarce sub forma de String continutul unui fisier Text...
' Pentru simplificarea codului continutul este citit in intregime in memorie
' presupunandu-se ca intr-un proiect sursele sunt suficient de mici pentru a
' putea fi citite in memorie una cate una
Public Function GetTextFileContent(ByVal strName As String) As String
    Dim fso As Scripting.FileSystemObject
    Dim fil As Scripting.File
    
    Set fso = New Scripting.FileSystemObject
    Set fil = fso.GetFile(strName)
    With fil.OpenAsTextStream(ForReading, TristateFalse)
        GetTextFileContent = .ReadAll
    End With
    Set fil = Nothing
    Set fso = Nothing
End Function
