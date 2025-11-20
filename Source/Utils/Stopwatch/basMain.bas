Attribute VB_Name = "basMain"
Option Explicit

Public cfg As clsConfig

Sub Main()
    If App.PrevInstance Then Exit Sub
    Set cfg = New clsConfig
    frmStopwatch.Show
End Sub

Public Sub CloseApp()
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
    Next
    Set cfg = Nothing
End Sub

Public Function DataPath() As String
    If DebugMode() Then
        DataPath = App.Path & "\..\..\..\Utils\"
    Else
        DataPath = App.Path & "\"
    End If
End Function
