Attribute VB_Name = "basExport"
Option Explicit

Private mstrLine() As String
Private mlngLines As Long
Private mlngBuffer As Long


' ************* GENERAL *************


Private Function ExportCharacterFile(pstrPath As String, pstrExtension As String, plngCharacter As Long) As String
    Dim strCharacter As String
    Dim strFile As String
    
    If plngCharacter = 0 Or plngCharacter > db.Characters Then
        Notice "Choose a character to export"
        Exit Function
    End If
    strCharacter = db.Character(plngCharacter).Character
    strFile = strCharacter & "." & pstrExtension
    strFile = xp.ShowSaveAsDialog(pstrPath, strFile, "Build Files|*." & pstrExtension, "*." & pstrExtension)
    If Len(strFile) Then
        If xp.File.Exists(strFile) Then
            If Not AskAlways(GetFileFromFilespec(strFile) & " already exists. Overwrite?") Then Exit Function
        End If
    End If
    ExportCharacterFile = strFile
End Function

Private Sub InitLines()
    mlngLines = 0
    mlngBuffer = 255
    ReDim mstrLine(1 To mlngBuffer)
End Sub

Private Sub AddLine(pstrLine As String, Optional plngExtraLines As Long = 0)
    mlngLines = mlngLines + 1
    If mlngLines + plngExtraLines > mlngBuffer Then
        mlngBuffer = (mlngBuffer * 3) \ 2
        ReDim Preserve mstrLine(1 To mlngBuffer)
    End If
    mstrLine(mlngLines) = pstrLine
    mlngLines = mlngLines + plngExtraLines
End Sub

Private Sub BlankLine(Optional plngLines As Long = 1)
    mlngLines = mlngLines + plngLines
End Sub

Private Sub SaveLines(pstrFile As String)
    If mlngLines Then
        ReDim Preserve mstrLine(1 To mlngLines)
        xp.File.SaveStringAs pstrFile, Join(mstrLine, vbNewLine)
    End If
    mlngLines = 0
    mlngBuffer = 0
    Erase mstrLine
End Sub


' ************* CHARACTER BUILDER LITE *************


Public Sub ExportCharacterLite(plngCharacter As Long)
    Dim strProgram As String
    Dim strFile As String
    Dim strLine() As String
    Dim lngCount As Long
    Dim strText As String
    Dim i As Long
    
    strFile = ExportCharacterFile(cfg.LitePath, "build", plngCharacter)
    If Len(strFile) = 0 Then Exit Sub
    cfg.LitePath = GetPathFromFilespec(strFile)
    InitLines
    With db.Character(plngCharacter)
        ' Overview
        AddLine "[Overview]", 1
        AddLine "Name: " & GetNameFromFilespec(strFile)
        AddLine "MaxLevels: 30", 1
        ' Notes
        If Len(.Notes) Then
            strLine = Split(.Notes, vbNewLine)
            For i = 0 To UBound(strLine)
                AddLine "Notes: " & strLine(i)
            Next
            BlankLine
        End If
        BlankLine
        AddLine "[Stats]", 1
        ' Past Lives
        lngCount = 0
        For i = 1 To UBound(.PastLife.Class)
            lngCount = lngCount + .PastLife.Class(i)
        Next
        For i = 1 To UBound(.PastLife.Racial)
            lngCount = lngCount + .PastLife.Racial(i)
        Next
        For i = 1 To UBound(.PastLife.Iconic)
            lngCount = lngCount + .PastLife.Iconic(i)
        Next
        Select Case lngCount
            Case 0: If cfg.BuildPoints = beAdventurer Then strText = "Adventurer" Else strText = "Champion"
            Case 1: strText = "Hero"
            Case Else: strText = "Legend"
        End Select
        AddLine "Preferred: " & strText, 1
        ' Stat tomes
        For i = 1 To UBound(.Tome.Stat)
            If .Tome.Stat(i) Then Exit For
        Next
        If i <= UBound(.Tome.Stat) Then
            AddLine ";    Advn  Chmp  Hero  Lgnd  Tome"
            AddLine ";    ----  ----  ----  ----  ----"
            For i = 1 To 6
                AddLine LiteStatTome(i, .Tome.Stat(i))
            Next
            AddLine ";    ----  ----  ----  ----"
            AddLine ";     28    32    34    36", 1
        End If
        BlankLine
        ' Skill tomes
        For i = 1 To UBound(.Tome.Skill)
            If .Tome.Skill(i) Then Exit For
        Next
        If i <= UBound(.Tome.Skill) Then
            AddLine "[Skills]", 1
            AddLine ";         1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  Tome"
            AddLine ";        ------------------------------------------------------------------------------------"
            For i = 1 To UBound(.Tome.Skill)
                AddLine LiteSkillTome(i, .Tome.Skill(i))
            Next
            BlankLine
        End If
        BlankLine
        ' RacialAP
        lngCount = .Tome.RacialAP
        For i = 1 To UBound(.PastLife.Racial)
            If .PastLife.Racial(i) = 3 Then lngCount = lngCount + 1
        Next
        If lngCount Then
            AddLine "[Enhancements]", 1
            AddLine "RacialAP: " & lngCount, 1
        End If
    End With
    SaveLines strFile
    strProgram = App.Path & "\CharacterBuilderLite.exe"
    If xp.File.Exists(strProgram) Then
        If AskAlways("Open " & GetFileFromFilespec(strFile) & " now?") Then xp.File.RunParams strProgram, strFile
    End If
End Sub

Private Function LiteStatTome(plngStat As Long, plngTome As Long) As String
    Dim strStat As String
    Dim strTome As String
    
    Select Case plngStat
        Case 1: strStat = "STR"
        Case 2: strStat = "DEX"
        Case 3: strStat = "CON"
        Case 4: strStat = "INT"
        Case 5: strStat = "WIS"
        Case 6: strStat = "CHA"
    End Select
    If plngTome = 0 Then strTome = "  " Else strTome = Right$(" " & CStr(plngTome), 2)
    LiteStatTome = strStat & ":                          " & strTome
End Function

Private Function LiteSkillTome(plngSkill As Long, plngTome As Long) As String
    Dim strSkill As String
    Dim strTome As String
    
    Select Case plngSkill
        Case 1: strSkill = "Balance"
        Case 2: strSkill = "Bluff"
        Case 3: strSkill = "Concent"
        Case 4: strSkill = "Diplo"
        Case 5: strSkill = "Disable"
        Case 6: strSkill = "Haggle"
        Case 7: strSkill = "Heal"
        Case 8: strSkill = "Hide"
        Case 9: strSkill = "Intim"
        Case 10: strSkill = "Jump"
        Case 11: strSkill = "Listen"
        Case 12: strSkill = "Move Si"
        Case 13: strSkill = "Open Lo"
        Case 14: strSkill = "Perform"
        Case 15: strSkill = "Repair"
        Case 16: strSkill = "Search"
        Case 17: strSkill = "Spellcr"
        Case 18: strSkill = "Spot"
        Case 19: strSkill = "Swim"
        Case 20: strSkill = "Tumble"
        Case 21: strSkill = "UMD"
    End Select
    strSkill = Left$(strSkill & "       ", 7)
    If plngTome Then strTome = plngTome Else strTome = " "
    LiteSkillTome = strSkill & ":                                                                                   " & strTome
End Function


' ************* DDO BUILDER *************


Public Sub ExportDDOBuilder(plngCharacter As Long)
    Dim strProgram As String
    Dim strFile As String
    Dim i As Long
    
    strFile = ExportCharacterFile(cfg.BuilderPath, "ddocp", plngCharacter)
    If Len(strFile) = 0 Then Exit Sub
    cfg.BuilderPath = GetPathFromFilespec(strFile)
    InitLines
    With db.Character(plngCharacter)
        AddLine "<DDOCharacterData>"
        AddLine "  <Character>"
        AddLine "    <Name>" & GetNameFromFilespec(strFile) & "</Name>"
        AddLine "    <Alignment>Lawful Good</Alignment>"
        AddLine "    <Race>Human</Race>"
        AddLine "    <AbilitySpend>"
        AddLine "      <AvailableSpend>36</AvailableSpend>"
        AddLine "      <StrSpend>0</StrSpend>"
        AddLine "      <DexSpend>0</DexSpend>"
        AddLine "      <ConSpend>0</ConSpend>"
        AddLine "      <IntSpend>0</IntSpend>"
        AddLine "      <WisSpend>0</WisSpend>"
        AddLine "      <ChaSpend>0</ChaSpend>"
        AddLine "    </AbilitySpend>"
        AddLine "    <GuildLevel>0</GuildLevel>"
        For i = 1 To 6
            AddLine BuilderStatTome(i, .Tome.Stat(i))
        Next
        AddLine "    <SkillTomes>"
        For i = 1 To 21
            AddLine BuilderSkillTome(i, .Tome.Skill(i))
        Next
        AddLine "    </SkillTomes>"
        AddBuilderSpecialFeats .Tome, .PastLife
    End With
    AddLine "    <ActiveStances>"
    AddLine "      <Stances>Unarmed</Stances>"
    AddLine "      <Stances>Cloth Armor</Stances>"
    AddLine "      <Stances>Centered</Stances>"
    AddLine "      <Stances>Lawful</Stances>"
    AddLine "      <Stances>Good</Stances>"
    AddLine "    </ActiveStances>"
    AddLine "    <Level4>Strength</Level4>"
    AddLine "    <Level8>Strength</Level8>"
    AddLine "    <Level12>Strength</Level12>"
    AddLine "    <Level16>Strength</Level16>"
    AddLine "    <Level20>Strength</Level20>"
    AddLine "    <Level24>Strength</Level24>"
    AddLine "    <Level28>Strength</Level28>"
    AddLine "    <Class1>Unknown</Class1>"
    AddLine "    <Class2>Unknown</Class2>"
    AddLine "    <Class3>Unknown</Class3>"
    AddLine "    <SelectedEnhancementTrees>"
    AddLine "      <TreeName>Human</TreeName>"
    AddLine "      <TreeName>Falconry</TreeName>"
    AddLine "      <TreeName>Harper Agent</TreeName>"
    AddLine "      <TreeName>Vistani Knife Fighter</TreeName>"
    AddLine "      <TreeName>No selection</TreeName>"
    AddLine "      <TreeName>No selection</TreeName>"
    AddLine "      <TreeName>No selection</TreeName>"
    AddLine "    </SelectedEnhancementTrees>"
    For i = 1 To 30
        AddLine "    <LevelTraining>"
        If i > 20 Then AddLine "      <Class>Epic</Class>"
        AddLine "      <SkillPointsAvailable>0</SkillPointsAvailable>"
        AddLine "      <SkillPointsSpent>0</SkillPointsSpent>"
        AddLine "      <TrainedFeats/>"
        AddLine "    </LevelTraining>"
    Next
    AddLine "    <ActiveEpicDestiny/>"
    AddLine "    <FatePoints>5</FatePoints>"
    For i = 1 To 5
        AddLine "    <TwistOfFate>"
        AddLine "      <Tier>0</Tier>"
        AddLine "    </TwistOfFate>"
    Next
    AddLine "    <ActiveGear>Standard</ActiveGear>"
    AddLine "    <EquippedGear>"
    AddLine "      <Name>Standard</Name>"
    AddLine "    </EquippedGear>"
    AddLine "  </Character>"
    AddLine "</DDOCharacterData>"
    SaveLines strFile
'    ' Can't immediately run DDOBuilder.exe; something still locks the file so DDOBuilder throws errors.
'    ' This problem persists even if we wait and run it from a timer, so sadly we just have to disable the feature.
'    strProgram = SearchForBuilderEXE(GetPathFromFilespec(strFile))
'    If Len(strProgram) Then
'        If AskAlways("Open " & GetFileFromFilespec(strFile) & " now?") Then xp.File.RunParams strProgram, strFile
'    End If
End Sub

Private Function SearchForBuilderEXE(ByVal pstrPath As String) As String
    Dim strFile As String
    Dim lngPos As Long
    
    Do
        strFile = pstrPath & "\DDOBuilder.exe"
        If xp.File.Exists(strFile) Then
            SearchForBuilderEXE = strFile
            Exit Function
        End If
        lngPos = InStrRev(pstrPath, "\")
        If lngPos Then pstrPath = Left$(pstrPath, lngPos - 1)
    Loop While InStr(pstrPath, "\")
End Function

Private Function BuilderStatTome(plngStat As Long, plngTome As Long) As String
    Dim strStat As String
    
    Select Case plngStat
        Case 1: strStat = "Str"
        Case 2: strStat = "Dex"
        Case 3: strStat = "Con"
        Case 4: strStat = "Int"
        Case 5: strStat = "Wis"
        Case 6: strStat = "Cha"
    End Select
    BuilderStatTome = "    <" & strStat & "Tome>" & plngTome & "</" & strStat & "Tome>"
End Function

Private Function BuilderSkillTome(plngSkill As Long, plngTome As Long) As String
    Dim strSkill As String
    
    Select Case plngSkill
        Case 1: strSkill = "Balance"
        Case 2: strSkill = "Bluff"
        Case 3: strSkill = "Concentration"
        Case 4: strSkill = "Diplomacy"
        Case 5: strSkill = "DisableDevice"
        Case 6: strSkill = "Haggle"
        Case 7: strSkill = "Heal"
        Case 8: strSkill = "Hide"
        Case 9: strSkill = "Intimidate"
        Case 10: strSkill = "Jump"
        Case 11: strSkill = "Listen"
        Case 12: strSkill = "MoveSilently"
        Case 13: strSkill = "OpenLock"
        Case 14: strSkill = "Perform"
        Case 15: strSkill = "Repair"
        Case 16: strSkill = "Search"
        Case 17: strSkill = "SpellCraft"
        Case 18: strSkill = "Spot"
        Case 19: strSkill = "Swim"
        Case 20: strSkill = "Tumble"
        Case 21: strSkill = "UMD"
    End Select
    BuilderSkillTome = "      <" & strSkill & ">" & plngTome & "</" & strSkill & ">"
End Function

Private Sub AddBuilderSpecialFeats(ptypTome As TomeType, ptypPL As PastLifeType)
    Dim blnSpecial As Boolean
    Dim strArray() As String
    Dim i As Long
    Const HeroicList As String = "Heroic,Barbarian,Bard,Cleric,Fighter,Paladin,Ranger,Rogue,Sorcerer,Wizard,Monk,Favored Soul,Artificer,Druid,Warlock,Alchemist"
    Const RacialList As String = "Racial,Aasimar,Dragonborn,Drow,Dwarf,Elf,Gnome,Halfling,Half-Elf,Half-Orc,Human,Warforged,Shifter"
    Const IconicList As String = "Iconic,Bladeforged,Deep Gnome,Morninglord,Purple Dragon Knight,Aasimar Scourge,Shadar-Kai"
    Const EpicList As String = "Epic,Arcane Sphere: Arcane Alacrity,Arcane Sphere: Energy Criticals,Arcane Sphere: Enchant Weapon,Divine Sphere: Brace,Divine Sphere: Power Over Life and Death,Divine Sphere: Block Energy,Martial Sphere: Doublestrike,Martial Sphere: Fortification,Martial Sphere: Skill Mastery,Primal Sphere: Doubleshot,Primal Sphere: Fast Healing,Primal Sphere: Colors of the Queen,Arcane Sphere: Ancient Knowledge,Divine Sphere: Ancient Blessings,Martial Sphere: Ancient Tactics,Primal Sphere: Ancient Power"
    ' Note ^^^ is out of order to support older load files
    
    strArray = Split(HeroicList, ",")
    For i = 1 To UBound(ptypPL.Class)
        AddBuilderPastLife strArray(0), strArray(i), ptypPL.Class(i), blnSpecial
    Next
    strArray = Split(RacialList, ",")
    For i = 1 To UBound(ptypPL.Racial)
        AddBuilderPastLife strArray(0), strArray(i), ptypPL.Racial(i), blnSpecial
    Next
    strArray = Split(IconicList, ",")
    For i = 1 To UBound(ptypPL.Iconic)
        AddBuilderPastLife strArray(0), strArray(i), ptypPL.Iconic(i), blnSpecial
    Next
    strArray = Split(EpicList, ",")
    For i = 1 To UBound(ptypPL.Epic)
        AddBuilderPastLife strArray(0), strArray(i), ptypPL.Epic(i), blnSpecial
    Next
    strArray = Split("Inherent,Melee Power,Ranged Power,Spell Power", ",")
    For i = 1 To UBound(ptypTome.Power)
        AddBuilderInherent strArray(i), ptypTome.Power(i), blnSpecial
    Next
    AddBuilderInherent "Fate Point", ptypTome.Fate, blnSpecial
    AddBuilderInherent "Physical Resistance", ptypTome.RR(1), blnSpecial
    AddBuilderInherent "Magical Resistance", ptypTome.RR(2), blnSpecial
    AddBuilderInherent "Racial Action Point", ptypTome.RacialAP, blnSpecial
    If blnSpecial Then
        AddLine "    </SpecialFeats>"
    Else
        AddLine "    <SpecialFeats/>"
    End If
End Sub

Private Sub AddBuilderPastLife(pstrType As String, pstrName As String, plngCount As Long, pblnSpecial As Boolean)
    Dim i As Long
    
    If plngCount > 0 And pblnSpecial = False Then
        AddLine "    <SpecialFeats>"
        pblnSpecial = True
    End If
    For i = 1 To plngCount
        AddLine "      <TrainedFeat>"
        AddLine "        <FeatName>Past Life: " & pstrName & "</FeatName>"
        AddLine "        <Type>" & pstrType & "PastLife</Type>"
        AddLine "        <LevelTrainedAt>0</LevelTrainedAt>"
        AddLine "      </TrainedFeat>"
    Next
End Sub

Private Sub AddBuilderInherent(pstrName As String, plngCount As Long, pblnSpecial As Boolean)
    Dim i As Long
    
    If plngCount > 0 And pblnSpecial = False Then
        AddLine "    <SpecialFeats>"
        pblnSpecial = True
    End If
    For i = 1 To plngCount
        AddLine "      <TrainedFeat>"
        AddLine "        <FeatName>Inherent " & pstrName & "</FeatName>"
        AddLine "        <Type>SpecialFeat</Type>"
        AddLine "        <LevelTrainedAt>0</LevelTrainedAt>"
        AddLine "      </TrainedFeat>"
    Next
End Sub


' ************* RON'S PLANNER *************


Public Sub ExportCharacterRon(plngCharacter As Long)
    Dim strProgram As String
    Dim strFile As String
    Dim strText As String
    Dim i As Long
    
    strFile = ExportCharacterFile(cfg.RonPath, "txt", plngCharacter)
    If Len(strFile) = 0 Then Exit Sub
    cfg.RonPath = GetPathFromFilespec(strFile)
    InitLines
    With db.Character(plngCharacter)
        ' Overview
        AddLine "NAME: " & .Character & ", ;"
        AddLine "TOMERAISE: "
        For i = 1 To 6
            AddLine .Tome.Stat(i) & ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,"
        Next
        AddLine ";"
        strText = vbNullString
        AddLine "PASTLIFE: "
        For i = 1 To UBound(.PastLife.Class)
            strText = strText & .PastLife.Class(RonClass(i)) & ", "
        Next
        AddLine strText & ";"
        strText = vbNullString
        AddLine "ICONICPL: "
        For i = 1 To UBound(.PastLife.Iconic)
            strText = strText & .PastLife.Iconic(RonIconic(i)) & ", "
        Next
        AddLine strText & ";"
        strText = vbNullString
        AddLine "EPICPL: "
        For i = 1 To UBound(.PastLife.Epic)
            strText = strText & .PastLife.Epic(RonEpic(i)) & ", "
        Next
        AddLine strText & ";"
        strText = vbNullString
        AddLine "RACEPL: "
        For i = 1 To UBound(.PastLife.Racial)
            strText = strText & .PastLife.Racial(RonRace(i)) & ", "
        Next
        AddLine strText & ";"
    End With
    SaveLines strFile
End Sub

' Send this a position used by Ron's planner and it returns the position used by Compendium
Private Function RonClass(plngClass As Long) As Long
    Select Case plngClass
        Case 1: RonClass = 4
        Case 2: RonClass = 5
        Case 3: RonClass = 1
        Case 4: RonClass = 10
        Case 5: RonClass = 7
        Case 6: RonClass = 6
        Case 7: RonClass = 3
        Case 8: RonClass = 9
        Case 9: RonClass = 8
        Case 10: RonClass = 2
        Case 11: RonClass = 11
        Case 12: RonClass = 12
        Case 13: RonClass = 13
        Case 14: RonClass = 14
        Case Else: RonClass = plngClass
    End Select
End Function

' Send this a position used by Ron's planner and it returns the position used by Compendium
Private Function RonRace(plngRace As Long) As Long
    Select Case plngRace
        Case 1: RonRace = 10
        Case 2: RonRace = 5
        Case 3: RonRace = 7
        Case 4: RonRace = 4
        Case 5: RonRace = 11
        Case 6: RonRace = 3
        Case 7: RonRace = 8
        Case 8: RonRace = 9
        Case 9: RonRace = 6
        Case 10: RonRace = 2
        Case 11: RonRace = 1
        Case Else: RonRace = plngRace
    End Select
End Function

' Send this a position used by Ron's planner and it returns the position used by Compendium
Private Function RonIconic(plngIconic As Long) As Long
    Select Case plngIconic
        Case 1: RonIconic = 1
        Case 2: RonIconic = 3
        Case 3: RonIconic = 4
        Case 4: RonIconic = 6
        Case 5: RonIconic = 2
        Case 6: RonIconic = 5
        Case Else: RonIconic = plngIconic
    End Select
End Function

Private Function RonEpic(plngEpic As Long) As Long
    Select Case plngEpic
        Case 1: RonEpic = 2
        Case 2: RonEpic = 3
        Case 3: RonEpic = 1
        Case 4: RonEpic = 5
        Case 5: RonEpic = 4
        Case 6: RonEpic = 6
        Case 7: RonEpic = 7
        Case 8: RonEpic = 9
        Case 9: RonEpic = 8
        Case 10: RonEpic = 10
        Case 11: RonEpic = 11
        Case 12: RonEpic = 12
        Case Else: RonEpic = plngEpic
    End Select
End Function

