Attribute VB_Name = "basDataSave"
Option Explicit

Private mstrLine() As String
Private mlngLines As Long


' ************* GENERAL *************


Public Sub DirtyFlag(penFlag As DirtyFlagEnum)
    frmCompendium.tmrAutoSave.Enabled = False
    gblnDirtyFlag(dfeAny) = True
    gblnDirtyFlag(penFlag) = True
    frmCompendium.SetCaption
    frmCompendium.tmrAutoSave.Enabled = True
End Sub

Public Sub AutoSave()
    frmCompendium.tmrAutoSave.Enabled = False
    If gblnDirtyFlag(dfeSettings) Then cfg.SaveSettings
    If gblnDirtyFlag(dfeData) Then SaveCompendiumFile
    If gblnDirtyFlag(dfeNotes) Then SaveNotes
    If gblnDirtyFlag(dfeLinks) Then SaveLinkLists
    If xp.DebugMode Then ReportActions
    Erase gblnDirtyFlag
    frmCompendium.SetCaption
End Sub

Private Sub ReportActions()
    Dim strMessage As String
    
    If Not gblnDirtyFlag(0) Then Exit Sub
    If gblnDirtyFlag(dfeSettings) Then strMessage = "Settings, "
    If gblnDirtyFlag(dfeData) Then strMessage = strMessage & "Data, "
    If gblnDirtyFlag(dfeLinks) Then strMessage = strMessage & "Links, "
    If gblnDirtyFlag(dfeNotes) Then strMessage = strMessage & "Notes, "
    If Len(strMessage) Then strMessage = Left(strMessage, Len(strMessage) - 2)
    Debug.Print "AutoSaved " & strMessage & " at " & Format(Now(), "Long Time")
End Sub

Public Sub SaveCompendiumFile()
    If Len(cfg.DataFile) = 0 Then Exit Sub
    CalculateFavor
    ResetLines 511
    SaveCharacterSettings
    SaveCharacters
    SaveQuests
    SaveChallenges
    SaveLines CompendiumFile()
    Erase mstrLine
End Sub

Public Sub RenameCompendiumFile(pstrNew As String)
    Dim strOld As String
    Dim strNew As String
    Dim strOldFile As String
    
    AutoSave
    strOldFile = CompendiumFile()
    strOld = CompendiumBackupFile()
    cfg.DataFile = pstrNew
    SaveCompendiumFile
    strNew = CompendiumBackupFile()
    frmCompendium.SetCaption
    If xp.File.Exists(strNew) Then xp.File.Delete strNew
    If xp.File.Exists(strOld) Then xp.File.Move strOld, strNew
    If xp.File.Exists(strOldFile) Then xp.File.Delete strOldFile
End Sub

Public Sub SaveAllData()
    cfg.SaveSettings
    ' Compendium
    If Len(cfg.DataFile) Then
        SaveCompendiumFile
        DeleteBackup CompendiumBackupFile()
    End If
    ' Notes
    SaveNotes
    DeleteBackup NotesBackupFile()
    ' Links
    SaveLinkLists
    DeleteBackup LinkListsBackupFile()
End Sub

Private Sub DeleteBackup(pstrFile As String)
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
End Sub

Private Sub ResetLines(plngBuffer As Long)
    ReDim mstrLine(plngBuffer)
    mlngLines = -1
End Sub

Private Sub AddLine(pstrNewLine As String)
    mlngLines = mlngLines + 1
    If mlngLines > UBound(mstrLine) Then ReDim Preserve mstrLine(((mlngLines - 1) * 3) \ 2)
    mstrLine(mlngLines) = pstrNewLine
End Sub

Private Sub SaveLines(pstrFile As String)
    If mlngLines < 0 Then Exit Sub
    If mlngLines < UBound(mstrLine) Then ReDim Preserve mstrLine(mlngLines)
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
    xp.File.SaveStringAs pstrFile, Join(mstrLine, vbNewLine)
End Sub


' ************* LINKLISTS *************


Private Sub SaveLinkLists()
    Dim strFile As String
    Dim blnMultiple As Boolean
    Dim lngCount As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim i As Long
    Dim c As Long
    
    blnMultiple = (CountCompendiums() > 1)
    ResetLines 127
    For i = 1 To db.LinkLists
        frmCompendium.SaveLinkList i ' Copy the in-control link list data to global structure gtypMenu
        With gtypMenu
            If Not .Deleted Then
                AddLine "LinkList: " & .Title
                frmCompendium.GetMenuCoords i, lngLeft, lngTop
                AddLine "Left: " & lngLeft
                AddLine "Top: " & lngTop
                For c = 1 To .Commands
                    With .Command(c)
                        AddLine "Menu: " & GetMenuStyleName(.Style) & vbTab & .Caption & vbTab & .Target & vbTab & .Param
                    End With
                Next
                mlngLines = mlngLines + 1
                lngCount = lngCount + 1
            End If
        End With
    Next
    strFile = LinkListsFile()
    If lngCount Then
        SaveLines strFile
    Else
        If xp.File.Exists(strFile) Then xp.File.Delete strFile
    End If
End Sub


' ************* SECTIONS *************


Private Function OutputCoords(ptypCoords As CoordsType) As String
    With ptypCoords
        OutputCoords = .Left & vbTab & .Top & vbTab & .Width & vbTab & .Height
    End With
End Function

Private Sub SaveCharacterSettings()
    AddLine "SectionName: Settings"
    mlngLines = mlngLines + 1
    AddLine "BackColor: " & cfg.GetColor(cgeControls, cveBackground)
    mlngLines = mlngLines + 1
End Sub

Private Sub SaveCharacters()
    Dim strNotes() As String
    Dim i As Long
    Dim c As Long
    
    If db.Characters = 0 Then Exit Sub
    AddLine "SectionName: Characters"
    mlngLines = mlngLines + 1
    For c = 1 To db.Characters
        With db.Character(c)
            AddLine "Character: " & .Character
            If .Level Then AddLine "Level: " & .Level
            AddLine "CustomColor: " & .CustomColor
            AddLine "GeneratedColor: " & GetColorName(.GeneratedColor)
            AddLine "BackColor: " & .BackColor
            AddLine "DimColor: " & .DimColor
            AddLine "LeftClick: " & .LeftClick
            With .ContextMenu
                For i = 1 To .Commands
                    With .Command(i)
                        AddLine "Menu: " & GetMenuStyleName(.Style) & vbTab & .Caption & vbTab & .Target & vbTab & .Param
                    End With
                Next
            End With
            AddLine "TomeStat: " & ArrayToString(.Tome.Stat)
            AddLine "TomeSkill: " & ArrayToString(.Tome.Skill)
            AddLine "TomeHeroicXP: " & .Tome.HeroicXP
            AddLine "TomeEpicXP: " & .Tome.EpicXP
            AddLine "TomeRacialAP: " & .Tome.RacialAP
            AddLine "TomeUniversalAP: " & .Tome.UniversalAP
            AddLine "TomeFate: " & .Tome.Fate
            AddLine "TomeDestiny: " & .Tome.Destiny
            AddLine "TomePower: " & ArrayToString(.Tome.Power)
            AddLine "TomeRR: " & ArrayToString(.Tome.RR)
            AddLine "PastLifeClass: " & ArrayToString(.PastLife.Class)
            AddLine "PastLifeRace: " & ArrayToString(.PastLife.Racial)
            AddLine "PastLifeIconic: " & ArrayToString(.PastLife.Iconic)
            AddLine "PastLifeEpic: " & ArrayToString(.PastLife.Epic)
            For i = 1 To db.Sagas
                AddLine "Saga: " & db.Saga(i).SagaName & vbTab & CharacterSagaData(c, i)
            Next
            If Len(.Notes) Then
                strNotes = Split(.Notes, vbNewLine)
                For i = 0 To UBound(strNotes)
                    AddLine "Notes: " & strNotes(i)
                Next
            End If
        End With
        mlngLines = mlngLines + 1
    Next
End Sub

Private Function ArrayToString(plngArray() As Long) As String
    Dim strReturn As String
    Dim i As Long
    Dim iMax As Long
    
    iMax = UBound(plngArray)
    strReturn = Space$(iMax)
    For i = 1 To iMax
        Mid$(strReturn, i, 1) = plngArray(i)
    Next
    ArrayToString = strReturn
End Function

Private Function CharacterSagaData(plngCharacter As Long, plngSaga As Long) As String
    Dim strReturn As String
    Dim i As Long
    
    strReturn = Space$(db.Saga(plngSaga).Quests)
    With db.Character(plngCharacter).Saga(plngSaga)
        For i = 1 To db.Saga(plngSaga).Quests
            Mid$(strReturn, i, 1) = GetProgressLetter(.Progress(i))
        Next
    End With
    CharacterSagaData = strReturn
End Function

Private Sub SaveQuests()
    Dim strQuest As String
    Dim strProgress As String
    Dim strFlags As String
    Dim q As Long
    Dim c As Long
    
    AddLine "SectionName: Quests"
    mlngLines = mlngLines + 1
    For q = 1 To db.Quests
        With db.Quest(q)
            strProgress = vbNullString
            For c = 1 To db.Characters
                strProgress = strProgress & GetProgressLetter(.Progress(c))
            Next
            strFlags = vbNullString
            If .Skipped Then strFlags = strFlags & "k"
            AddLine .ID & vbTab & strProgress & vbTab & strFlags
        End With
    Next
    mlngLines = mlngLines + 1
End Sub

Private Sub SaveChallenges()
    Dim strChallenge As String
    Dim strProgress As String
    Dim i As Long
    Dim c As Long
    
    AddLine "SectionName: Challenges"
    mlngLines = mlngLines + 1
    For i = 1 To db.Challenges
        With db.Challenge(i)
            strProgress = vbNullString
            For c = 1 To db.Characters
                strProgress = strProgress & .Stars(c)
            Next
            AddLine .ID & vbTab & strProgress
        End With
    Next
End Sub
