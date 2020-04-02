Attribute VB_Name = "basCompendium"
Option Explicit


' ************* DATA *************


Public Sub InitData()
    LoadTomeData
    BackupFiles
    LoadDataFiles
    CalculateFavor
    InitHelp
End Sub

Public Sub UpdateFavor()
    Dim frm As Form
    
    If GetForm(frm, "frmPatrons") Then frm.Recalculate
    If GetForm(frm, "frmCharacter") Then frm.FavorTotals
End Sub


' ************* FILES *************


Public Function DataPath() As String
    DataPath = App.Path & "\Data\Compendium\"
End Function

Private Function SavePath() As String
    SavePath = cfg.CompendiumPath & "\"
End Function

Public Function LinkListsFile() As String
    LinkListsFile = SavePath() & "Compendium.linklists"
End Function

Public Function LinkListsBackupFile(Optional pblnNoPath As Boolean = False) As String
    Dim strReturn As String
    
    strReturn = "Compendium_Backup.linklists"
    If Not pblnNoPath Then strReturn = SavePath() & strReturn
    LinkListsBackupFile = strReturn
End Function

Public Function CompendiumFile() As String
    CompendiumFile = DataFileToFileName(cfg.DataFile)
End Function

Public Function CompendiumBackupFile() As String
    CompendiumBackupFile = DataFileToFileName(cfg.DataFile, True)
End Function

Public Function NotesFile() As String
    NotesFile = SavePath() & "Notes.txt"
End Function

Public Function NotesBackupFile(Optional pblnNoPath As Boolean = False) As String
    Dim strReturn As String
    
    strReturn = "Notes_Backup.txt"
    If Not pblnNoPath Then strReturn = SavePath() & strReturn
    NotesBackupFile = strReturn
End Function

Public Sub BackupCompendium()
    BackupFile CompendiumFile(), CompendiumBackupFile()
End Sub

Public Sub BackupFiles()
    BackupFile CompendiumFile(), CompendiumBackupFile()
    BackupFile NotesFile(), NotesBackupFile()
    BackupFile LinkListsFile(), LinkListsBackupFile()
End Sub

Private Sub BackupFile(pstrFile As String, pstrBackup As String)
    If xp.File.Exists(pstrFile) Then
        If Not xp.File.Exists(pstrBackup) Then xp.File.Copy pstrFile, pstrBackup
    Else
        If xp.File.Exists(pstrBackup) Then xp.File.Copy pstrBackup, pstrFile
    End If
End Sub

Public Function ProperName(pstrDataFile As String) As String
    Dim strProper As String
    
    strProper = Trim$(xp.File.MakeNameDOS(pstrDataFile))
    If LCase$(Right$(strProper, 11)) = ".compendium" Then strProper = Trim$(Left$(strProper, Len(strProper) - 11))
    ProperName = StrConv(strProper, vbProperCase)
End Function

Public Function DataFileToFileName(pstrDataFile As String, Optional pblnBackup As Boolean = False) As String
    Dim strBackup As String
    
    If pblnBackup Then strBackup = "_Backup"
    DataFileToFileName = SavePath() & pstrDataFile & strBackup & ".compendium"
End Function

Public Function FileNameToDataFile(pstrFileName As String) As String
    FileNameToDataFile = ProperName(GetNameFromFilespec(pstrFileName))
End Function

Public Function ChooseEXE() As Boolean
    Dim strFile As String
    
    strFile = xp.ShowOpenDialog(vbNullString, "Programs|*.exe|All Files|*.*", "*.exe")
    If Len(strFile) Then
        cfg.PlayEXE = strFile
        ChooseEXE = True
        DirtyFlag dfeSettings
    End If
End Function


' ************* BACKUPS *************


Public Sub RevertCharacters()
    If Not ValidateRevert(CompendiumFile(), CompendiumBackupFile(), "Character", "character data") Then Exit Sub
    OpenCompendiumFile False
End Sub

Public Sub RevertNotes()
    If Not ValidateRevert(NotesFile(), NotesBackupFile(), "Notes") Then Exit Sub
    frmCompendium.usrtxtNotes.Text = xp.File.LoadToString(NotesFile())
End Sub

Public Sub RevertLinks()
    If Not ValidateRevert(LinkListsFile(), LinkListsBackupFile(), "Links") Then Exit Sub
    frmCompendium.ReloadLinkLists
End Sub

Private Function ValidateRevert(pstrFile As String, pstrBackup As String, pstrName As String, Optional pstrDescrip As String) As Boolean
    AutoSave
    If Len(pstrDescrip) = 0 Then pstrDescrip = pstrName
    If xp.File.Exists(pstrBackup) Then
        If MsgBox("Restore " & pstrDescrip & " from backup?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") = vbYes Then
            If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
            xp.File.Copy pstrBackup, pstrFile
            ValidateRevert = True
        End If
    Else
        MsgBox pstrName & " backup not found", vbInformation, "Notice"
    End If
End Function

Public Sub RefreshBackups()
    AutoSave
    If MsgBox("Overwrite all backups with current data?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    RefreshBackup CompendiumFile(), CompendiumBackupFile()
    RefreshBackup NotesFile(), NotesBackupFile()
    RefreshBackup LinkListsFile(), LinkListsBackupFile()
    MsgBox "New backups created successfully", vbInformation, "Notice"
End Sub

Private Sub RefreshBackup(pstrFile As String, pstrBackup As String)
    If Not xp.File.Exists(pstrFile) Then Exit Sub
    If xp.File.Exists(pstrBackup) Then xp.File.Delete pstrBackup
    xp.File.Copy pstrFile, pstrBackup
End Sub

Public Sub DeleteBackups()
    AutoSave
    DeleteFile CompendiumBackupFile()
    DeleteFile NotesBackupFile()
    DeleteFile LinkListsBackupFile()
End Sub

Private Sub DeleteFile(pstrFile As String)
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
End Sub


' ************* COLORS *************


Public Sub MatchColors(pblnApply As Boolean)
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim lngMid As Long
    Dim i As Long
    
    cfg.CompendiumBackColor = cfg.GetColor(cgeControls, cveBackground)
    ' RGB
    xp.ColorToRGB cfg.GetColor(cgeControls, cveBackground), lngRed, lngGreen, lngBlue
    lngMid = Int((lngRed + lngGreen + lngBlue) / 3 + 0.5)
    lngMid = ((lngMid + 2) \ 5) * 5
    Select Case lngMid
        Case Is < 60: lngMid = 70
        Case Is < 85: lngMid = 85
        Case Is > 230: lngMid = 230
    End Select
    cfg.NamedHigh = lngMid + 25
    cfg.NamedMed = lngMid
    cfg.NamedLow = lngMid - 25
    ' Dim
    Select Case lngMid
        Case 0 To 85: cfg.NamedDim = 75
        Case 86 To 99: cfg.NamedDim = 80
        Case 100 To 124: cfg.NamedDim = 85
        Case 125 To 255: cfg.NamedDim = 90
    End Select
    ' Apply
    If pblnApply Then
        For i = 1 To db.Characters
            With db.Character(i)
                If Not .CustomColor Then
                    .BackColor = GetColorValue(.GeneratedColor)
                    .DimColor = GetColorDim(.BackColor)
                End If
            End With
        Next
    End If
End Sub


' ************* NOTES *************


Public Function LoadNotes() As String
    Dim strFile As String
    Dim strBackup As String
    
    strFile = NotesFile()
    strBackup = NotesBackupFile()
    If xp.File.Exists(strFile) Then
        If Not xp.File.Exists(strBackup) Then xp.File.Copy strFile, strBackup
        LoadNotes = xp.File.LoadToString(strFile)
    Else
        If xp.File.Exists(strBackup) Then
            xp.File.Copy strBackup, strFile
            LoadNotes = xp.File.LoadToString(strFile)
        End If
    End If
End Function

Public Sub SaveNotes()
    Dim strFile As String
    
    strFile = NotesFile()
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    xp.File.SaveStringAs strFile, frmCompendium.usrtxtNotes.Text
    Broadcast "Notes"
End Sub


' ************* FAVOR *************


Public Sub CalculateFavor()
    Dim lngFavor As Long
    Dim q As Long
    Dim c As Long
    
    For c = 1 To db.Characters
        db.Character(c).ChallengeFavor = 0
        db.Character(c).QuestFavor = 0
        db.Character(c).TotalFavor = 0
    Next
    For q = 1 To db.Quests
        For c = 1 To db.Characters
            lngFavor = QuestFavor(q, db.Quest(q).Progress(c))
            With db.Character(c)
                .QuestFavor = .QuestFavor + lngFavor
                .TotalFavor = .TotalFavor + lngFavor
            End With
        Next
    Next
    For q = 1 To db.Challenges
        For c = 1 To db.Characters
            lngFavor = db.Challenge(q).Stars(c)
            With db.Character(c)
                .ChallengeFavor = .ChallengeFavor + lngFavor
                .TotalFavor = .TotalFavor + lngFavor
            End With
        Next
    Next
End Sub

Public Function QuestFavor(plngQuest As Long, ByVal penDifficulty As ProgressEnum) As Long
    Select Case db.Quest(plngQuest).Style
        Case qeQuest: If penDifficulty = peSolo Then penDifficulty = peCasual
        Case qeRaid: If penDifficulty = peCasual Then penDifficulty = peNormal
        Case qeSolo: If penDifficulty <> peNone Then penDifficulty = peSolo
    End Select
    Select Case penDifficulty
        Case peElite: QuestFavor = db.Quest(plngQuest).Favor * 3
        Case peNone: QuestFavor = 0
        Case peSolo, peNormal: QuestFavor = db.Quest(plngQuest).Favor
        Case peHard: QuestFavor = db.Quest(plngQuest).Favor * 2
        Case peCasual: QuestFavor = db.Quest(plngQuest).Favor \ 2
    End Select
End Function

Public Sub PatronWiki(pstrPatron As String)
    Dim lngIndex As Long
    
    lngIndex = SeekPatron(pstrPatron)
    If lngIndex Then xp.OpenURL MakeWiki("Favor#" & db.Patron(lngIndex).Wiki)
End Sub

Public Sub Reincarnate(plngCharacter As Long)
    Dim frm As Form
    Dim i As Long
    
    If plngCharacter = 0 Then Exit Sub
    If MsgBox("Clear all quest, saga and challenge progress for " & db.Character(plngCharacter).Character & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    For i = 1 To db.Quests
        db.Quest(i).Progress(plngCharacter) = peNone
    Next
    For i = 1 To db.Challenges
        db.Challenge(i).Stars(plngCharacter) = 0
    Next
    With db.Character(plngCharacter)
        .QuestFavor = 0
        .ChallengeFavor = 0
        .TotalFavor = 0
        For i = 1 To db.Sagas
            ReDim .Saga(i).Progress(1 To db.Saga(i).Quests)
        Next
    End With
    If GetForm(frm, "frmCharacter") Then frm.FavorTotals
    If GetForm(frm, "frmChallenges") Then frm.ReQueryData plngCharacter
    If GetForm(frm, "frmPatrons") Then frm.RedrawForm
    If GetForm(frm, "frmSagas") Then frm.Redraw
    frmCompendium.RedrawQuests
    frmCompendium.FavorChange plngCharacter
    DirtyFlag dfeData
End Sub


' ************* MENUS *************


Public Sub QuestWiki(plngQuest As Long)
    xp.OpenURL MakeWiki(db.Quest(plngQuest).Wiki)
End Sub

Public Sub QuestLink(plngQuest As Long, pstrCaption As String)
    Dim typCommand As MenuCommandType
    Dim i As Long
    
    With db.Quest(plngQuest)
        For i = 1 To .Links
            If .Link(i).FullName = pstrCaption Then
                typCommand.Caption = .Link(i).FullName
                typCommand.Style = .Link(i).Style
                typCommand.Target = .Link(i).Target
                RunCommand typCommand
                Exit For
            End If
        Next
    End With
End Sub

Public Sub PackWiki(pstrPack As String)
    Dim lngPack As Long
    Dim strWiki As String
    
    If pstrPack = "Free to Play" Or Len(pstrPack) = 0 Then
        strWiki = "Guide to Free to Play#Quest_list"
    Else
        lngPack = SeekPack(pstrPack)
        If lngPack Then strWiki = db.Pack(lngPack).Wiki
    End If
    If Len(strWiki) Then xp.OpenURL MakeWiki(strWiki)
End Sub

Public Sub PackLink(plngPack As Long, pstrCaption As String)
    Dim typCommand As MenuCommandType
    Dim i As Long
    
    If plngPack < 1 Then Exit Sub
    With db.Pack(plngPack)
        For i = 1 To .Links
            If .Link(i).FullName = pstrCaption Then
                typCommand.Caption = .Link(i).FullName
                typCommand.Style = .Link(i).Style
                typCommand.Target = .Link(i).Target
                RunCommand typCommand
                Exit For
            End If
        Next
    End With
End Sub

