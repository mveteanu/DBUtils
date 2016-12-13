Attribute VB_Name = "modStart"
Option Explicit

Public Const dbuTypeTable As String = "TABLE"
Public Const dbuTypeView As String = "VIEW"
Public Const dbuTypeProc As String = "PROC"
Public Const dbuTypeFile As String = "FILE"

Public Const dbuVBReg As String = "Software\VB and VBA Program Settings\"
Public Const dbuAppName As String = "VMASOFT DBUtils"
Public Const dbuPersistentWindows As String = "Windows\"
Public Const dbuPersistentConnections As String = "Connections\"

Public fMainForm As frmMain

Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub

