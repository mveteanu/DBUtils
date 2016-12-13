Attribute VB_Name = "modSearchText"
Option Explicit

' Intoarce true daca strWord care este numele unei tabele sau query sau field se afla
' in interiorul lui strText care este textul unui query
' intWordsMode = 0 - se cauta textul asa cum se da
' intWordsMode = 1 - se cauta nume de tabele sau query-uri
' intWordsMode = 2 - se cauta nume de tabele, query-uri, campuri... si in fisiere
Public Function WordInsideText(ByVal strWord As String, ByVal strText As String, ByVal intWordsMode) As Boolean
    Dim p As Long
    Dim re As Boolean
    Dim let1 As String, let2 As String
    Dim ar1 As Variant, ar2 As Variant, arc As Variant, cap As Integer
    
    p = InStr(1, strText, strWord, vbTextCompare)
    If intWordsMode = 0 Then
        WordInsideText = (p > 0)
        Exit Function
    End If
    If p > 0 Then
        re = True
        If p > 1 Then
            let1 = Mid(strText, p - 1, 1)
        Else
            let1 = " "
        End If
        let2 = Mid(strText, p + Len(strWord), 1)
        If Len(Trim(let2)) = 0 Then let2 = " "
        Select Case intWordsMode
            Case 1
                ar1 = Array(" ", vbCr, vbLf, vbTab, "[", ",", "(", "&", "*", "+", "-", "=", "/", "<", ">", "!", "`")
                ar2 = Array(" ", vbCr, vbLf, vbTab, ".", "!", "]", ",", ")")
            Case 2
                ar1 = Array(" ", vbCr, vbLf, vbTab, ".", ":", ";", "[", ",", "(", "&", "*", "+", "-", "=", "/", "<", ">", "!", "`", "'", """")
                ar2 = Array(" ", vbCr, vbLf, vbTab, ".", ":", ";", "!", "[", "]", ",", "(", ")", "&", "*", "+", "-", "=", "/", ">", "<", "'", """")
        End Select
        cap = 0
        For Each arc In ar1
            If let1 = arc Then cap = cap + 1
        Next
        For Each arc In ar2
            If let2 = arc Then cap = cap + 1
        Next
        re = (cap = 2)
    Else
        re = False
    End If
    
    WordInsideText = re
End Function


' Cauta sa vada daca cuvintele specificate de strWords se afla in interiorul textului
' ca separator se foloseste operatorul "and"... pentru negatie se foloseste "not"
' Exemplu: marian and veteanu
'          pitesti and not spectralims
Public Function WordsInsideText(ByVal strWords As String, ByVal strText As String, ByVal bWordsOnly As Boolean) As Boolean
    Const strAnd As String = " and "
    Const strNot As String = "not "
    Dim arWords As Variant
    Dim i As Integer
    Dim allFound As Boolean, bComp As Boolean
    
    arWords = Split(strWords, strAnd, -1, vbTextCompare)
    For i = LBound(arWords) To UBound(arWords)
        arWords(i) = Trim(arWords(i))
    Next
    
    allFound = True
    For i = LBound(arWords) To UBound(arWords)
        If LCase(Left(arWords(i), 4)) = strNot Then
            arWords(i) = Trim(Right(arWords(i), Len(arWords(i)) - Len(strNot)))
            bComp = True
        End If
        If WordInsideText(arWords(i), strText, IIf(bWordsOnly, 2, 0)) = bComp Then
            allFound = False
            Exit For
        End If
    Next
    
    WordsInsideText = allFound
End Function

