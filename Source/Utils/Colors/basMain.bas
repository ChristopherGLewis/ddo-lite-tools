Attribute VB_Name = "basMain"
Option Explicit

Public Const ErrorIgnore As Long = 37001

Public glngActiveColor As Long

Public PixelX As Long
Public PixelY As Long

Public xp As clsWindowsXP
Public cfg As clsConfig

Sub Main()
    If App.PrevInstance Then Exit Sub
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    Set xp = New clsWindowsXP
    Set cfg = New clsConfig
    InitHelp
    If Not CommandLine() Then frmColors.Show
End Sub

Public Sub CloseApp()
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
    Next
    Set cfg = Nothing
    Set xp = Nothing
End Sub

Public Function DataPath() As String
    If xp.DebugMode Then DataPath = App.Path & "\..\..\..\Utils\" Else DataPath = App.Path & "\"
End Function

Public Function SettingsPath() As String
    If xp.DebugMode Then SettingsPath = App.Path & "\..\..\..\Settings\" Else SettingsPath = App.Path & "\..\Settings\"
End Function

Public Sub ColorChange()
    cfg.SaveSettings
    Broadcast "Colors"
End Sub

Public Sub OpenForm(pstrForm As String)
    Dim frm As Form

    Select Case pstrForm
        Case "frmHelp": frmHelp.Show
        Case Else: MsgBox "Feature under construction.", vbInformation, "Sorry..."
    End Select
End Sub

Private Function CommandLine() As Boolean
    Dim strFile As String
    
    If Len(Command) = 0 Then Exit Function
    strFile = SettingsPath() & Command & ".screen"
    If Not xp.File.Exists(strFile) Then Exit Function
    cfg.LoadColorFile strFile
    ColorChange
    CommandLine = True
    CloseApp
End Function
