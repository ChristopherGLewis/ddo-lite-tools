Attribute VB_Name = "basMain"
Option Explicit

Public Const ErrorIgnore = 37001
Public glngActiveColor As Long

Public PixelX As Long
Public PixelY As Long

Public xp As clsWindowsXP
Public cfg As clsConfig

Sub Main()
    If App.PrevInstance Then
        MsgBox "Only one Compendium may be running at a time.", vbInformation, "Notice"
        Exit Sub
    End If
    Load frmMessages
    Randomize
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    Set xp = New clsWindowsXP
    Set cfg = New clsConfig
    
    InitData
    frmCompendium.Show
End Sub

Public Sub CloseApp()
    Dim frm As Form
    
    SaveAllData
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
    Set cfg = Nothing
    Set xp = Nothing
End Sub

Public Sub OpenForm(pstrForm As String)
    Dim frm As Form
    
    Select Case pstrForm
        Case "frmTools": Set frm = frmTools
        Case "frmCharacters": Set frm = frmCharacter
        Case "frmChallenges": Set frm = frmChallenges
        Case "frmSagas": Set frm = frmSagas
        Case "frmColorPreview": Set frm = frmColorPreview
        Case "frmPatrons": Set frm = frmPatrons
        Case "frmCharacter": Set frm = frmCharacter
        Case "frmAbout": Set frm = frmAbout
        Case "frmCritCalculator": Set frm = frmCritCalculator
        Case "frmWilderness": Set frm = frmWilderness
        Case Else: Exit Sub
    End Select
    If cfg.ChildWindows Then
        frm.Show vbModeless, frmCompendium
    Else
        frm.Show
    End If
    Set frm = Nothing
End Sub

Public Function RandomNumber(plngUpper As Long, Optional plngLower As Long = 1) As Long
    RandomNumber = Int((plngUpper - plngLower + 1) * Rnd + plngLower)
End Function

Public Sub RunCommand(ptypCommand As MenuCommandType)
    Dim strFile As String
    
    If Len(ptypCommand.Target) = 0 Then
        MsgBox "No target specified.", vbInformation, "Notice"
        Exit Sub
    End If
    Select Case ptypCommand.Style
        Case mceLink
            xp.OpenURL ptypCommand.Target
        Case mceShortcut
            strFile = ptypCommand.Target
            If InStr(strFile, ":") = 0 And Left$(strFile, 2) <> "\\" Then
                If Left$(strFile, 1) = "\" Then strFile = App.Path & strFile Else strFile = App.Path & "\" & strFile
            End If
            If Not xp.File.Exists(strFile) Then
                MsgBox "File not found:" & vbNewLine & vbNewLine & ptypCommand.Target, vbInformation, "Notice"
                Exit Sub
            End If
            If Len(ptypCommand.Param) Then
                xp.File.RunParams strFile, ptypCommand.Param
            Else
                xp.File.Run strFile
            End If
    End Select
End Sub

Public Function RunLeftClickCommand(plngCharacter As Long) As Boolean
    Dim strCommand As String
    Dim blnFound As Boolean
    Dim i As Long
    
    With db.Character(plngCharacter)
        strCommand = .LeftClick
        With .ContextMenu
            For i = 1 To .Commands
                With .Command(i)
                    If .Caption = strCommand Then
                        blnFound = True
                        Exit For
                    End If
                End With
            Next
        End With
    End With
    If blnFound Then
        RunLeftClickCommand = True
        RunCommand db.Character(plngCharacter).ContextMenu.Command(i)
    End If
End Function

Public Sub LightsOut()
    Dim strFile As String
    
    strFile = App.Path & "\Utils\LightsOut.exe"
    If xp.File.Exists(strFile) Then xp.File.Run strFile
End Sub

Public Sub ADQRiddle()
    Dim strFile As String
    
    strFile = App.Path & "\Utils\ADQ.exe"
    If xp.File.Exists(strFile) Then xp.File.Run strFile
End Sub


' ************* DEV TOOLS *************


Public Sub PatronWidths()
    Dim strLine() As String
    Dim strName As String
    Dim strFile As String
    Dim i As Long
    
    ReDim strLine(1 To db.Patrons)
    For i = 1 To db.Patrons
        strName = db.Patron(i).Abbreviation
        strLine(i) = strName & "," & frmCompendium.GetTextWidth(strName)
    Next
    strFile = App.Path & "\Patrons.csv"
    xp.File.SaveStringAs strFile, Join(strLine, vbNewLine)
End Sub

Public Sub PackWidths()
    Dim strLine() As String
    Dim strName As String
    Dim strFile As String
    Dim i As Long
    
    ReDim strLine(db.Packs)
    strName = "Free to Play"
    strLine(0) = strName & "," & frmCompendium.GetTextWidth(strName)
    For i = 1 To db.Packs
        strName = db.Pack(i).Abbreviation
        strLine(i) = strName & "," & frmCompendium.GetTextWidth(strName)
    Next
    strFile = App.Path & "\Packs.csv"
    xp.File.SaveStringAs strFile, Join(strLine, vbNewLine)
End Sub

