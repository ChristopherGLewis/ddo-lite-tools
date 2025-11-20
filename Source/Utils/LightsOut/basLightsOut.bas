Attribute VB_Name = "basLightsOut"
' Written by Ellis Dee
Option Explicit

Public cfg As clsConfig

Public Const ErrorIgnore As Long = 37001

Sub Main()
    If App.PrevInstance Then Exit Sub
    Set cfg = New clsConfig
    InitHelp
    frmLightsOut.Show
End Sub

Public Sub CloseApp()
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
    Set cfg = Nothing
End Sub

Public Function DataPath()
    If DebugMode() Then
        DataPath = App.Path & "\..\..\..\Utils\"
    Else
        DataPath = App.Path & "\"
    End If
End Function

Public Sub OpenForm(pstrForm As String)
    Dim frm As Form

    Select Case pstrForm
        Case "frmHelp": frmHelp.Show
        Case Else: MsgBox "Feature under construction.", vbInformation, "Sorry..."
    End Select
End Sub

