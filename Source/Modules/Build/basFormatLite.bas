Attribute VB_Name = "basFormatLite"
Option Explicit

Private Enum SectionEnum
    secUnknown
    secOverview
    secStats
    secSkills
    secFeats
    secSpells
    secEnhancements
    secLevelingGuide
    secDestiny
    secTwists
End Enum

Private Type GuideTreeLookupType
    Tree As Long
    Class As ClassEnum
    Display As String
End Type

Private mstrLine() As String
Private mlngLines As Long
Private mlngBuffer As Long

Private menSection As SectionEnum
Private mstrField As String
Private mstrValue As String
Private mstrList() As String
Private mlngListMax As Long

Private mlngSpellSlot() As Long

Private mtypLookup() As GuideTreeLookupType
Private mlngLookups As Long

Public Sub LiteFeatList()
    Dim strLine() As String
    Dim lngCurrent As Long
    Dim i As Long
    Dim s As Long
    Dim strFile As String
    
    ReDim strLine(db.Feats - 1)
    For i = 1 To db.Feats
        With db.Feat(i)
            If .Selectors = 0 Then
                strLine(i - 1) = .FeatName
            Else
                For s = 1 To .Selectors
                    strLine(i - 1) = strLine(i - 1) & .FeatName & ": " & .Selector(s).SelectorName
                    If s < .Selectors Then strLine(i - 1) = strLine(i - 1) & vbNewLine
                Next
            End If
        End With
    Next
    strFile = DataPath() & "FeatsLite.txt"
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    xp.File.SaveStringAs strFile, Join(strLine, vbNewLine)
End Sub


' ************* SAVE *************


Public Sub SaveFileLite(pstrFile As String)
    InitArray
    SaveOverviewLite
    SaveStatsLite
    SaveSkillsLite
    SaveFeatsLite
    SaveSpellsLite
    SaveEnhancementsLite
    SaveLevelingGuideLite
    SaveDestinyLite
    TrimArray
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
    xp.File.SaveStringAs pstrFile, Join(mstrLine, vbNewLine)
    Erase mstrLine
End Sub


' ************* ARRAY *************


Private Sub InitArray()
    mlngBuffer = 511
    mlngLines = 0
    ReDim mstrLine(mlngBuffer)
End Sub

Private Sub AddLine(pstrLine As String, Optional plngLines As Long = 1)
    mstrLine(mlngLines) = pstrLine
    GrowArray plngLines
End Sub

Private Sub AddSection(pstrSection As String)
    Dim lngBlanks As Long
    
    If mlngLines > 2 Then
        lngBlanks = 2
        If Len(mstrLine(mlngLines - 1)) = 0 Then
            lngBlanks = 1
            If Len(mstrLine(mlngLines - 2)) = 0 Then lngBlanks = 0
        End If
        If lngBlanks Then BlankLine lngBlanks
    End If
    AddLine "[" & pstrSection & "]", 2
End Sub

Private Sub BlankLine(Optional plngLines As Long = 1)
    GrowArray plngLines
End Sub

Private Sub GrowArray(plngLines As Long)
    mlngLines = mlngLines + plngLines
    If mlngLines > mlngBuffer Then
        mlngBuffer = mlngBuffer + 128
        ReDim Preserve mstrLine(mlngBuffer)
    End If
End Sub

Private Sub TrimArray()
    If mlngLines <> mlngBuffer Then ReDim Preserve mstrLine(mlngLines)
End Sub

Private Sub RevertToLine(plngLines As Long)
    Dim i As Long
    
    For i = plngLines To mlngLines
        mstrLine(i) = vbNullString
    Next
    mlngLines = plngLines
End Sub

Private Function GetSectionName() As String
    Select Case menSection
        Case secOverview: GetSectionName = "Overview"
        Case secStats: GetSectionName = "Stats"
        Case secSkills: GetSectionName = "Skills"
        Case secFeats: GetSectionName = "Feats"
        Case secSpells: GetSectionName = "Spells"
        Case secEnhancements: GetSectionName = "Enhancements"
        Case secLevelingGuide: GetSectionName = "Leveling Guide"
        Case secDestiny: GetSectionName = "Destiny"
        Case secTwists: GetSectionName = "Twists"
        Case Else: GetSectionName = "Unknown"
    End Select
End Function



' ************* SAVE - OVERVIEW *************


Private Sub SaveOverviewLite()
    Dim i As Long
    
    ' General
    AddSection "Overview"
    AddLine "Name: " & build.BuildName
    If build.Race <> reAny Then AddLine "Race: " & GetRaceName(build.Race)
    If build.Alignment <> aleAny Then AddLine "Alignment: " & GetAlignmentName(build.Alignment)
    AddLine "MaxLevels: " & build.MaxLevels
    ' Notes
    SaveNotesLite
    ' Build Classes
    If build.BuildClass(0) <> ceAny Then BlankLine
    For i = 0 To 2
        If build.BuildClass(i) <> ceAny Then AddLine "Class: " & GetClassName(build.BuildClass(i))
    Next
    SaveClassLevelsLite
End Sub

Private Sub SaveNotesLite()
    Dim strLine() As String
    Dim i As Long
    Dim iMax As Long
    
    If Len(build.Notes) = 0 Then Exit Sub
    strLine = Split(build.Notes, vbNewLine)
    For iMax = UBound(strLine) To 0 Step -1
        If Len(Trim$(strLine(iMax))) <> 0 Then Exit For
    Next
    BlankLine
    For i = 0 To iMax
        AddLine "Notes: " & strLine(i)
    Next
End Sub

Private Sub SaveClassLevelsLite()
    Dim i As Long
    
    For i = 2 To 20
        If build.Class(i) <> build.Class(1) Then Exit For
    Next
    If i > 20 Then Exit Sub
    BlankLine
    For i = 1 To 20
        AddLine "Level: " & i & vbTab & GetClassName(build.Class(i))
    Next
End Sub


' ************* SAVE - STATS *************


Private Sub SaveStatsLite()
    Dim lngLines As Long
    Dim blnInclude As Boolean
    Dim i As Long
    
    AddSection "Stats"
    AddLine "Preferred: " & GetBuildPointsName(build.BuildPoints), 2
    For i = 0 To 3
        If build.IncludePoints(i) = 0 Then AddLine GetBuildPointsName(i) & ": No"
    Next
    ' Build Points Table
    lngLines = mlngLines
    AddLine ";    Advn  Chmp  Hero  Lgnd  Tome"
    AddLine ";    ----  ----  ----  ----  ----"
    For i = 1 To 6
        AddLine MakeStatLine(i, False, blnInclude)
    Next
    AddLine ";    ----  ----  ----  ----"
    AddLine MakeStatLine(0, True, blnInclude), 2
    If Not blnInclude Then RevertToLine lngLines
    ' Levelups
    lngLines = mlngLines
    blnInclude = False
    For i = 0 To 7
        AddLine "Levelup: " & i * 4 & vbTab & GetStatName(build.Levelups(i))
        If build.Levelups(i) <> aeAny Then blnInclude = True
    Next
    If Not blnInclude Then RevertToLine lngLines
End Sub

Private Function MakeStatLine(ByVal penStat As StatEnum, pblnTotal As Boolean, pblnInclude As Boolean) As String
    Dim strReturn As String
    Dim lngStat As Long
    Dim strStat As String
    Dim i As Long
    
    If pblnTotal Then strReturn = ";    " Else strReturn = UCase$(GetStatName(penStat, True)) & ": "
    For i = 0 To 3
        lngStat = build.StatPoints(i, penStat)
        If lngStat = 0 Then
            strStat = "      "
        Else
            strStat = " " & Right$("  " & lngStat, 2) & "   "
            pblnInclude = True
        End If
        strReturn = strReturn & strStat
    Next
    If Not pblnTotal Then
        If build.Tome(penStat) = 0 Then
            strStat = "   "
        Else
            strStat = " " & Right$("  " & build.Tome(penStat), 2)
            pblnInclude = True
        End If
        strReturn = strReturn & strStat
    End If
    MakeStatLine = strReturn
End Function

Private Function NoZero(ByVal plngValue As Long) As String
    If plngValue Then NoZero = plngValue Else NoZero = " "
End Function

Private Function GetBuildPointsName(ByVal penBuildPoints As BuildPointsEnum) As String
    Select Case penBuildPoints
        Case beAdventurer: GetBuildPointsName = "Adventurer"
        Case beChampion: GetBuildPointsName = "Champion"
        Case beHero: GetBuildPointsName = "Hero"
        Case beLegend: GetBuildPointsName = "Legend"
    End Select
End Function


' ************* SAVE - SKILLS *************


Private Sub SaveSkillsLite()
    Dim lngLines As Long
    Dim blnInclude As Boolean
    Dim i As Long
    
    lngLines = mlngLines
    AddSection "Skills"
    AddLine ";         1   2   3   4   5   6   7   8   9  10  11  12  13  14  15  16  17  18  19  20  Tome"
    AddLine ";        ------------------------------------------------------------------------------------"
    For i = 1 To seSkills - 1
        AddLine MakeSkillsLine(i, blnInclude)
    Next
    If Not blnInclude Then RevertToLine lngLines
End Sub

Private Function MakeSkillsLine(penSkill As SkillsEnum, pblnInclude As Boolean) As String
    Dim strReturn As String
    Dim strTome As String
    Dim i As Long
    
    strReturn = Left$(GetSkillName(penSkill, True) & "       ", 7) & ": "
    For i = 1 To 20
        If build.Skills(penSkill, i) = 0 Then
            strReturn = strReturn & "    "
        Else
            strReturn = strReturn & Right$(" " & build.Skills(penSkill, i), 2) & "  "
            pblnInclude = True
        End If
    Next
    If build.SkillTome(penSkill) = 0 Then
        strTome = "   "
    Else
        strTome = "  " & build.SkillTome(penSkill)
        pblnInclude = True
    End If
    MakeSkillsLine = strReturn & strTome
End Function


' ************* SAVE - FEATS *************


Private Sub SaveFeatsLite()
    Dim lngLines As Long
    Dim blnInclude As Boolean
    Dim i As Long
    
    lngLines = mlngLines
    AddSection "Feats"
    For i = bftStandard To bftFeatTypes - 1
        If AddFeatGroup(i) Then
            BlankLine
            blnInclude = True
        End If
    Next
    If Not blnInclude Then RevertToLine lngLines
End Sub

Private Function AddFeatGroup(penType As BuildFeatTypeEnum) As Boolean
    Dim strField As String
    Dim strSlot As String
    Dim strLevel As String
    Dim strFeat As String
    Dim lngFeat As Long
    Dim strLine As String
    Dim blnBlankLine As Boolean
    Dim i As Long
    
    If build.Feat(penType).Feats = 0 Then Exit Function
    strField = GetBuildFeatGroup(penType)
    With build.Feat(penType)
        For i = 1 To .Feats
            With .Feat(i)
                strFeat = .FeatName
                If .Selector > 0 Then
                    lngFeat = SeekFeat(strFeat)
                    If lngFeat <> 0 Then
                        If db.Feat(lngFeat).Selectors >= .Selector Then strFeat = strFeat & ": " & db.Feat(lngFeat).Selector(.Selector).SelectorName
                    End If
                End If
                Select Case penType
                    Case bftAlternate, bftExchange
                        strSlot = GetBuildFeatSlot(build.Feat(.ChildType).Feat(.Child))
                        If penType = bftExchange Then strLevel = .Level & vbTab Else strLevel = vbNullString
                        strLine = strField & ": " & strSlot & vbTab & strLevel & strFeat
                    Case Else
                        strSlot = GetBuildFeatSlot(build.Feat(penType).Feat(i))
                        If Len(strFeat) Then strLine = strField & ": " & strSlot & vbTab & strFeat Else strLine = vbNullString
                End Select
                If Len(strLine) Then
                    AddLine strLine
                    AddFeatGroup = True
                End If
            End With
        Next
    End With
End Function

Private Function GetBuildFeatGroup(penType As BuildFeatTypeEnum) As String
    Dim strReturn As String
    
    Select Case penType
        Case bftGranted: strReturn = "Granted"
        Case bftStandard: strReturn = "Standard"
        Case bftLegend: strReturn = "Legend"
        Case bftRace: strReturn = "Race"
        Case bftClass1, bftClass2, bftClass3: strReturn = "Class"
        Case bftDeity: strReturn = "Deity"
        Case bftAlternate: strReturn = "Alternate"
        Case bftExchange: strReturn = "Exchange"
    End Select
    GetBuildFeatGroup = strReturn
End Function

Private Function GetBuildFeatSlot(ptypBuildFeat As BuildFeatType) As String
    Dim strReturn As String
    
    With ptypBuildFeat
        Select Case .Type
            Case bftGranted: strReturn = "Granted"
            Case bftStandard
                Select Case .Level
                    Case Is < 20: strReturn = "Heroic " & .Level
                    Case 21, 24, 27, 30: strReturn = "Epic " & .Level
                    Case 26, 28, 29: strReturn = "Destiny " & .Level
                End Select
            Case bftLegend: strReturn = "Legend " & .Level
            Case bftRace: strReturn = GetRaceName(build.Race) & " " & .Level
            Case bftClass1: strReturn = GetClassName(build.BuildClass(0)) & " " & .ClassLevel
            Case bftClass2: strReturn = GetClassName(build.BuildClass(1)) & " " & .ClassLevel
            Case bftClass3: strReturn = GetClassName(build.BuildClass(2)) & " " & .ClassLevel
            Case bftDeity: strReturn = "Deity" & " " & .Level
            Case bftAlternate: strReturn = "Alternate" & " " & .Level
            Case bftExchange: strReturn = "Exchange" & " " & .Level
        End Select
    End With
    GetBuildFeatSlot = strReturn
End Function


' ************* SAVE - SPELLS *************


Private Sub SaveSpellsLite()
    Dim lngLines As Long
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngSlot As Long
    Dim strClass As String
    Dim blnInclude As Boolean
    Dim blnBlankLine As Boolean
    
    If build.CanCastSpell(1) = 0 Then Exit Sub ' Can't cast level 1 spells
    lngLines = mlngLines
    AddSection "Spells"
    For enClass = 1 To ceClasses - 1
        strClass = GetClassName(enClass)
        With build.Spell(enClass)
            For lngLevel = 1 To .MaxSpellLevel
                With .Level(lngLevel)
                    For lngSlot = 1 To .Slots
                        With .Slot(lngSlot)
                            AddLine "Spell: " & strClass & " " & lngLevel & vbTab & .Spell
                            If Len(.Spell) Then
                                If .SlotType = sseStandard Or .SlotType = sseFree Then blnInclude = True
                            End If
                        End With
                    Next
                    If .Slots <> 0 Then BlankLine
                End With
            Next
        End With
    Next
    If Not blnInclude Then RevertToLine lngLines
End Sub


' ************* SAVE - ENHANCEMENTS *************


Private Sub SaveEnhancementsLite()
    Dim i As Long
    
    If Len(build.Tier5) = 0 And build.RacialAP = 0 And build.UniversalAP = 0 And build.Trees = 0 Then Exit Sub
    AddSection "Enhancements"
    ' Tier5
    If Len(build.Tier5) Then AddLine "Tier5: " & build.Tier5
    ' Racial AP
    If build.RacialAP <> 0 Then AddLine "RacialAP: " & build.RacialAP
    If build.UniversalAP <> 0 Then AddLine "UniversalAP: " & build.UniversalAP
    If Len(build.Tier5) <> 0 Or build.RacialAP <> 0 Or build.UniversalAP <> 0 Then BlankLine
    ' Trees
    For i = 1 To build.Trees
        With build.Tree(i)
            AddLine "Tree: " & .TreeName
            AddLine "Type: " & GetTreeStyleName(.TreeType)
            If .Source <> 0 Then AddLine "Source: " & GetClassName(.Source)
            If .ClassLevels <> 0 Then AddLine "ClassLevels: " & .ClassLevels
            AddAbilityLite build.Tree(i), peEnhancement
            BlankLine
        End With
    Next
End Sub

Private Sub AddAbilityLite(ptypBuildTree As BuildTreeType, penTreeType As PointerEnum)
    Dim strLine As String
    Dim lngTree As Long
    Dim i As Long
    
    If ptypBuildTree.Abilities = 0 Then Exit Sub
    lngTree = SeekTree(ptypBuildTree.TreeName, penTreeType)
    If lngTree = 0 Then Exit Sub
    For i = 1 To ptypBuildTree.Abilities
        Select Case penTreeType
            Case peEnhancement: strLine = GetAbilityLine(ptypBuildTree.Ability(i), db.Tree(lngTree))
            Case peDestiny: strLine = GetAbilityLine(ptypBuildTree.Ability(i), db.Destiny(lngTree))
        End Select
        AddLine strLine
    Next
End Sub

Private Function GetAbilityLine(ptypAbility As BuildAbilityType, ptypTree As TreeType) As String
On Error GoTo GetAbilityLineErr
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRank As Long
    Dim strReturn As String
    
    With ptypAbility
        lngTier = .Tier
        lngAbility = .Ability
        lngSelector = .Selector
        lngRank = .Rank
    End With
    With ptypTree.Tier(lngTier).Ability(lngAbility)
        strReturn = "Ability: Tier " & lngTier & ": " & .AbilityName
        If lngSelector Then strReturn = strReturn & ": " & .Selector(lngSelector).SelectorName
        If .Ranks > 1 Then strReturn = strReturn & " (Rank " & lngRank & ")"
    End With
    
GetAbilityLineExit:
    GetAbilityLine = strReturn
    Exit Function
    
GetAbilityLineErr:
    strReturn = "Error"
    Resume GetAbilityLineExit
End Function


' ************* SAVE - LEVELING GUIDE *************


Private Sub SaveLevelingGuideLite()
    Dim strLine As String
    Dim lngTree As Long
    Dim i As Long
    
    If build.Guide.Trees = 0 And build.Guide.Enhancements = 0 Then Exit Sub
    AddSection "LevelingGuide"
    With build.Guide
        If .Enhancements <> 0 Then
            For i = 1 To .Enhancements
                AddLine "Guide: " & GetGuideLine(build.Guide, i)
            Next
        End If
    End With
End Sub

Private Function GetGuideLine(ptypBuild As BuildGuideType, plngEnhancement As Long) As String
On Error GoTo GetGuideLineErr
    Dim strReturn As String
    Dim lngID As Long
    Dim strTree As String
    Dim lngTree As Long
    Dim typGuide As BuildGuideEnhancementType
    Dim typAbility As AbilityType
    
    typGuide = ptypBuild.Enhancement(plngEnhancement)
    lngID = typGuide.ID
    Select Case lngID
        Case 0
            strReturn = "Unknown"
        Case 1 To 99
            GuideTreeInfo plngEnhancement, lngTree, strTree
            With Guide.Tree(Guide.Enhancement(plngEnhancement).GuideTreeID)
                lngTree = .TreeID
                strTree = .Display
            End With
            strReturn = strTree & " Tier " & typGuide.Tier & ": "
            typAbility = db.Tree(lngTree).Tier(typGuide.Tier).Ability(typGuide.Ability)
            strReturn = strReturn & typAbility.AbilityName
            If typGuide.Selector <> 0 Then strReturn = strReturn & ": " & typAbility.Selector(typGuide.Selector).SelectorName
            If typAbility.Ranks > 1 Then strReturn = strReturn & " (Rank " & typGuide.Rank & ")"
        Case 100
            strReturn = "Reset All Trees"
        Case 101 To 199
            GuideTreeInfo plngEnhancement, lngTree, strTree
            strReturn = "Reset Tree: " & strTree
        Case 200
            strReturn = "Bank AP"
    End Select
    
GetGuideLineExit:
    GetGuideLine = strReturn
    Exit Function
    
GetGuideLineErr:
    strReturn = "Error"
    Resume GetGuideLineExit
End Function

Private Sub GuideTreeInfo(plngEnhancement As Long, plngTreeID As Long, pstrDisplay As String)
    With Guide.Tree(Guide.Enhancement(plngEnhancement).GuideTreeID)
        plngTreeID = .TreeID
        pstrDisplay = .Display
    End With
End Sub


' ************* SAVE - DESTINY *************


Private Sub SaveDestinyLite()
    Dim i As Long
    
    If Len(build.DestinyTier5) = 0 And build.DestinyAP = 0 Then Exit Sub
    AddSection "Destiny"
    ' Tier5
    If Len(build.DestinyTier5) Then AddLine "DestinyTier5: " & build.DestinyTier5
    ' Destiny AP
    If build.DestinyAP <> 0 Then AddLine "DestinyAP: " & build.DestinyAP
    If Len(build.DestinyTier5) <> 0 Or build.DestinyAP <> 0 Then BlankLine
    ' Destinies
    For i = 1 To build.Destinies
        With build.Destiny(i)
            AddLine "Destiny: " & .TreeName
            AddLine "Type: " & GetTreeStyleName(.TreeType)
            'TODO
            'If .Source <> 0 Then AddLine "Source: " & GetClassName(.Source)
            'If .ClassLevels <> 0 Then AddLine "ClassLevels: " & .ClassLevels
            AddAbilityLite build.Destiny(i), peDestiny
            BlankLine
        End With
    Next



End Sub


' ************* LOAD *************


Public Function LoadFileLite(pstrFile As String) As Boolean
On Error GoTo LoadFileLiteErr
    Dim strLine As String
    Dim lngLine As Long
    Dim i As Long
    
    ClearBuild False
    SetBuildDefaults
    ReDim mlngSpellSlot(1 To ceClasses - 1, 1 To 20)
    Erase mtypLookup
    mlngLookups = 0
    LoadArray pstrFile
    For lngLine = 0 To mlngLines
        strLine = mstrLine(lngLine)
        If ParseLine(strLine) Then
            Select Case menSection
                Case secOverview: LoadOverviewText
                Case secStats: LoadStatsText
                Case secSkills: LoadSkillsText
                Case secFeats: LoadFeatsText
                Case secSpells: LoadSpellsText
                Case secEnhancements: LoadEnhancementsText
                Case secLevelingGuide: LoadLevelingGuideText
                Case secDestiny: LoadDestinyText
            End Select
        End If
    Next
    LoadFileLite = True
    
LoadFileLiteExit:
    Erase mstrLine
    Erase mlngSpellSlot
    Erase mtypLookup
    frmMain.tmrDeprecate.Enabled = True
    Exit Function
    
LoadFileLiteErr:
    Dim strMessage As String
    strMessage = "Unexpected error: " & Err.Description & vbNewLine & vbNewLine
    strMessage = strMessage & "Section: " & GetSectionName() & vbNewLine
    strMessage = strMessage & "Field: " & mstrField & vbNewLine
    If mlngListMax = 0 Then
        strMessage = strMessage & "Value: " & mstrValue & vbNewLine
    Else
        For i = 0 To mlngListMax
            strMessage = strMessage & "Value " & i + 1 & ": " & mstrList(i) & vbNewLine
        Next
    End If
    strMessage = strMessage & vbNewLine & "File: " & GetFileFromFilespec(pstrFile)
    strMessage = strMessage & vbNewLine & "Line Number: " & lngLine + 1
    strMessage = strMessage & vbNewLine & strLine & vbNewLine
    strMessage = strMessage & vbNewLine & "Ignore error and load anyway?"
    If MsgBox(strMessage, vbQuestion + vbYesNo, "Error " & Err.Number) <> vbYes Then
        cfg.Dirty = False
        Resume LoadFileLiteExit
    Else
        Resume Next
    End If
End Function

Private Sub LoadArray(pstrFile As String)
    Dim strRaw As String
    
    strRaw = xp.File.LoadToString(pstrFile)
    mstrLine = Split(strRaw, vbNewLine)
    mlngLines = UBound(mstrLine)
End Sub

Private Function ParseLine(pstrRaw As String) As Boolean
    Dim lngPos As Long
    
    Select Case Left$(pstrRaw, 1)
        Case ";", " "
        Case "["
            If Right$(pstrRaw, 1) <> "]" Then Exit Function
            ' Finalize previous section
            Select Case menSection
                Case secOverview: FinalizeOverview
            End Select
            ' Start new section
            Select Case LCase$(pstrRaw)
                Case "[overview]": menSection = secOverview
                Case "[stats]": menSection = secStats
                Case "[skills]": menSection = secSkills
                Case "[feats]": menSection = secFeats
                Case "[spells]": menSection = secSpells
                Case "[enhancements]": menSection = secEnhancements
                Case "[levelingguide]": menSection = secLevelingGuide
                Case "[destiny]": menSection = secDestiny
                Case "[twists]": menSection = secTwists
                Case Else: menSection = secSkills
            End Select
        Case Else
            lngPos = InStr(pstrRaw, ": ")
            If lngPos = 0 Then Exit Function
            mstrField = LCase$(Trim$(Left$(pstrRaw, lngPos - 1)))
            mstrValue = Mid$(pstrRaw, lngPos + 2)
            mstrList = Split(mstrValue, vbTab)
            mlngListMax = UBound(mstrList)
            ParseLine = True
    End Select
End Function

Private Function SplitCombo(pstrCombo As String, pstrText As String, plngNumber As Long) As Boolean
    Dim lngPos As Long
    Dim strNumber As String
    
    lngPos = InStrRev(pstrCombo, " ")
    If lngPos < 2 Then Exit Function
    pstrText = LCase$(Left$(pstrCombo, lngPos - 1))
    strNumber = Trim$(Mid$(pstrCombo, lngPos + 1))
    If Not IsNumeric(strNumber) Then Exit Function
    plngNumber = val(strNumber)
    SplitCombo = True
End Function


' ************* LOAD - OVERVIEW *************


Private Sub LoadOverviewText()
    Select Case mstrField
        Case "name": build.BuildName = mstrValue
        Case "race": build.Race = GetRaceID(mstrValue)
        Case "alignment": build.Alignment = GetAlignmentID(mstrValue)
        Case "maxlevels": build.MaxLevels = val(mstrValue)
        Case "notes": If Len(build.Notes) = 0 Then build.Notes = mstrValue Else build.Notes = build.Notes & vbNewLine & mstrValue
        Case "class": SetBuildClass
        Case "level": SetBuildLevel
    End Select
End Sub

Private Sub SetBuildClass()
    Dim enClass As ClassEnum
    Dim i As Long
    
    enClass = GetClassID(mstrValue)
    If enClass = ceAny Then Exit Sub
    For i = 0 To 2
        If build.BuildClass(i) = ceAny Then Exit For
    Next
    If i < 3 Then build.BuildClass(i) = enClass
End Sub

Private Sub SetBuildLevel()
    Dim lngLevel As Long
    Dim enClass As ClassEnum
    
    If mlngListMax <> 1 Then Exit Sub
    lngLevel = val(mstrList(0))
    If lngLevel < 1 Or lngLevel > 20 Then Exit Sub
    enClass = GetClassID(mstrList(1))
    If enClass <> ceAny Then build.Class(lngLevel) = enClass
End Sub

Private Sub FinalizeOverview()
    Dim lngLookup() As Long
    Dim enClass As ClassEnum
    Dim lngBuildClass As Long
    Dim i As Long
    
    ' Assign missing BuildClasses
    If build.BuildClass(0) = ceAny Then
        ReDim lngLookup(ceClasses - 1)
        For i = 1 To 20
            enClass = build.Class(i)
            If enClass <> ceAny Then
                If lngLookup(enClass) = 0 Then
                    lngLookup(0) = lngLookup(0) + 1
                    lngLookup(enClass) = lngLookup(0)
                End If
                lngBuildClass = lngLookup(enClass)
                If lngBuildClass > 0 And lngBuildClass < 4 Then build.BuildClass(lngBuildClass - 1) = enClass
            End If
        Next
    End If
    ' Assign missing ClassLevels
    If build.BuildClass(0) <> ceAny Then
        For i = 1 To 20
            If build.Class(i) = ceAny Then build.Class(i) = build.BuildClass(0)
        Next
    End If
    ' Calculate BAB
    CalculateBAB
    ' Initialize feat slots
    InitBuildFeats
    ' Initialize spells
    InitBuildSpells
End Sub


' ************* LOAD - STATS *************


Private Sub LoadStatsText()
    Dim i As Long
    
    Select Case mstrField
        Case "preferred"
            build.BuildPoints = GetBuildPointsID(mstrValue)
        Case "levelup"
            If mlngListMax = 1 Then
                i = val(mstrList(0)) \ 4
                If i >= 0 And i <= 7 Then build.Levelups(i) = GetStatID(mstrList(1))
            End If
        Case Else
            ' Include Build points
            For i = 0 To 3
                If mstrField = LCase$(GetBuildPointsName(i)) Then
                    If LCase$(Trim$(mstrValue)) = "no" Then build.IncludePoints(i) = 0
                    Exit Sub
                End If
            Next
            ' Stats and Tomes
            For i = 1 To 6
                If mstrField = LCase$(GetStatName(i, True)) Then
                    ParseStatLine i
                    Exit Sub
                End If
            Next
    End Select
End Sub

Private Function GetBuildPointsID(ByVal pstrBuildPoints As String) As BuildPointsEnum
    Select Case LCase$(pstrBuildPoints)
        Case "adventurer": GetBuildPointsID = beAdventurer
        Case "champion": GetBuildPointsID = beChampion
        Case "hero": GetBuildPointsID = beHero
        Case "legend": GetBuildPointsID = beLegend
    End Select
End Function

Private Sub ParseStatLine(penStat As StatEnum)
    Dim i As Long
    
    For i = 0 To 3
        build.StatPoints(i, penStat) = val(Trim$(Mid$(mstrValue, (i * 6) + 1, 4)))
        build.StatPoints(i, 0) = build.StatPoints(i, 0) + build.StatPoints(i, penStat)
    Next
    build.Tome(penStat) = val(Trim$(Mid$(mstrValue, 25, 4)))
    If build.Tome(penStat) > tomes.Stat.Max Then build.Tome(penStat) = tomes.Stat.Max
End Sub


' ************* LOAD - SKILLS *************


Private Sub LoadSkillsText()
    Dim strValue As String
    Dim enSkill As SkillsEnum
    Dim lngLevel As Long
    
    For enSkill = 1 To seSkills - 1
        If mstrField = LCase$(GetSkillName(enSkill, True)) Then
            For lngLevel = 1 To 20
                build.Skills(enSkill, lngLevel) = val(Mid$(mstrValue, (lngLevel - 1) * 4 + 1, 3))
            Next
            build.SkillTome(enSkill) = val(Mid$(mstrValue, 82, 3))
            If build.SkillTome(enSkill) > tomes.Skill.Max Then build.SkillTome(enSkill) = tomes.Skill.Max
            Exit Sub
        End If
    Next
End Sub


' ************* LOAD - FEATS *************


Private Sub LoadFeatsText()
    Dim strFeat As String
    Dim strSelector As String
    Dim lngSelector As Long
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim enChildType As BuildFeatTypeEnum
    Dim lngChildIndex As Long
    
    If mlngListMax < 1 Then Exit Sub
    enChildType = bftUnknown
    Do
        ' Start by identifying actual type
        Select Case mstrField
            Case "alternate": enType = bftAlternate
            Case "exchange": enType = bftExchange
            Case Else: If Not IdentifySlot(mstrList(0), enType, lngIndex) Then Exit Sub Else Exit Do
        End Select
        ' Identify effective type for alternates and echanges
        If Not IdentifySlot(mstrList(0), enChildType, lngChildIndex) Then Exit Sub
        lngIndex = AddSlot(enType)
        build.Feat(enType).Feat(lngIndex).ChildType = enChildType
        build.Feat(enType).Feat(lngIndex).Child = lngChildIndex
        If enType = bftExchange Then
            build.Feat(enType).Feat(lngIndex).Level = val(mstrList(1))
        Else
            build.Feat(enType).Feat(lngIndex).Level = build.Feat(enChildType).Feat(lngChildIndex).Level
        End If
        build.Feat(enType).Feat(lngIndex).ClassLevel = build.Feat(enChildType).Feat(lngChildIndex).ClassLevel
        build.Feat(enType).Feat(lngIndex).Source = build.Feat(enChildType).Feat(lngChildIndex).Source
    Loop Until True
    ' Identify feat
    Select Case FindFeat(mstrList(mlngListMax), strFeat, strSelector, lngSelector)
        Case 0
            DeprecateFeat enType, lngIndex, strFeat, strSelector, enChildType, lngChildIndex
            Exit Sub
        Case 2
            DeprecateFeat enType, lngIndex, strFeat, strSelector, enChildType, lngChildIndex
            Exit Sub
    End Select
    build.Feat(enType).Feat(lngIndex).FeatName = strFeat
    build.Feat(enType).Feat(lngIndex).Selector = lngSelector
End Sub

Private Function FindFeat(pstrRaw As String, pstrFeat As String, pstrSelector As String, plngSelector As Long) As Long
    Dim lngFeat As Long
    Dim lngPos As Long
    
    lngPos = InStr(pstrRaw, ":")
    If lngPos = 0 Then
        pstrFeat = pstrRaw
        pstrSelector = vbNullString
    Else
        pstrFeat = Left$(pstrRaw, lngPos - 1)
        pstrSelector = Trim$(Mid$(pstrRaw, lngPos + 1))
        '2020.10.07 - hack to fix AchStudies which has a ":" in the feat name
        If (pstrFeat = "Alchemical Studies") Then
            pstrFeat = Trim(pstrRaw)
        End If
    End If
    lngFeat = DeprecateGetFeatID(pstrFeat)
    If lngFeat Then
        pstrFeat = db.Feat(lngFeat).FeatName
        With db.Feat(lngFeat)
            For plngSelector = .Selectors To 1 Step -1
                If LCase$(.Selector(plngSelector).SelectorName) = LCase$(pstrSelector) Then Exit For
            Next
            If plngSelector = 0 And .Selectors > 0 Then FindFeat = 2 Else FindFeat = 1
        End With
    End If
End Function

Private Function IdentifySlot(pstrRaw As String, penType As BuildFeatTypeEnum, plngIndex As Long) As Boolean
    Dim strText As String
    Dim lngNumber As Long
    Dim i As Long
    
    If Not SplitCombo(pstrRaw, strText, lngNumber) Then Exit Function
    penType = IdentifySlotType(strText)
    Select Case penType
        Case bftUnknown
            Exit Function
        Case bftAlternate, bftExchange
            plngIndex = 0
        Case bftClass1, bftClass2, bftClass3
            For plngIndex = 1 To build.Feat(penType).Feats
                If build.Feat(penType).Feat(plngIndex).ClassLevel = lngNumber Then Exit For
            Next
        Case Else
            For plngIndex = 1 To build.Feat(penType).Feats
                If build.Feat(penType).Feat(plngIndex).Level = lngNumber Then Exit For
            Next
    End Select
    If plngIndex > build.Feat(penType).Feats Then penType = bftUnknown Else IdentifySlot = True
End Function

Private Function IdentifySlotType(pstrText As String) As BuildFeatTypeEnum
    Dim enType As BuildFeatTypeEnum
    Dim bytClass As Byte
    
    Select Case pstrText
        Case "standard", "heroic", "epic", "destiny": enType = bftStandard
        Case "legend": enType = bftLegend
        Case "race": enType = bftRace
        Case "deity": enType = bftDeity
        Case "alternate": enType = bftAlternate
        Case "exchange": enType = bftExchange
        Case Else
            ' Weird bug. If I don't explicitly convert this value to a byte, the comparisons below will
            ' fail to work properly, but only for the compiled exe.
            bytClass = GetClassID(pstrText)
            Select Case bytClass
                Case 0: enType = bftUnknown
                Case build.BuildClass(0): enType = bftClass1
                Case build.BuildClass(1): enType = bftClass2
                Case build.BuildClass(2): enType = bftClass3
                Case Else: enType = bftUnknown
            End Select
            If enType = bftUnknown Then
                 If GetRaceID(pstrText) <> reAny Then enType = bftRace
            End If
    End Select
    IdentifySlotType = enType
End Function

Private Function AddSlot(penType As BuildFeatTypeEnum) As Long
    If penType = bftExchange And mlngListMax <> 2 Then Exit Function
    With build.Feat(penType)
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats).Type = penType
        AddSlot = .Feats
    End With
End Function


' ************* LOAD - SPELLS *************


Private Sub LoadSpellsText()
    Dim strClass As String
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngSlot As Long
    Dim strSpell As String
    Dim i As Long
    
    If mstrField <> "spell" Then Exit Sub
    If Not mlngListMax = 1 Then Exit Sub
    If Not SplitCombo(mstrList(0), strClass, lngLevel) Then Exit Sub
    enClass = GetClassID(strClass)
    If enClass = ceAny Then Exit Sub
    If lngLevel < 1 Or lngLevel > build.Spell(enClass).MaxSpellLevel Then Exit Sub
    lngSlot = mlngSpellSlot(enClass, lngLevel) + 1
    mlngSpellSlot(enClass, lngLevel) = lngSlot
    strSpell = mstrList(1)
    With build.Spell(enClass).Level(lngLevel)
        If CheckFree(enClass, strSpell) Then
            .FreeSlots = .FreeSlots + 1
            .Slots = .Slots + 1
            ReDim Preserve .Slot(1 To .Slots)
        End If
        With .Slot(lngSlot)
            If Len(.Spell) = 0 Then .Spell = strSpell
        End With
    End With
End Sub


' ************* LOAD - TREES *************


Private Sub LoadEnhancementsText()
    Select Case mstrField
        Case "tier5"
            build.Tier5 = mstrValue
        Case "racialap"
            build.RacialAP = val(mstrValue)
        Case "universalap"
            build.UniversalAP = val(mstrValue)
        Case "tree"
            build.Trees = build.Trees + 1
            ReDim Preserve build.Tree(1 To build.Trees)
            build.Tree(build.Trees).TreeName = mstrValue
        Case "type"
            If build.Trees <> 0 Then build.Tree(build.Trees).TreeType = GetTreeStyleID(mstrValue)
        Case "source"
            If build.Trees <> 0 Then build.Tree(build.Trees).Source = GetClassID(mstrValue)
        Case "classlevels"
            If build.Trees <> 0 Then build.Tree(build.Trees).ClassLevels = val(mstrValue)
        Case "ability"
            If build.Trees <> 0 Then AddTreeAbility mstrValue, build.Tree(build.Trees), peEnhancement
    End Select
End Sub

Private Sub LoadDestinyText()
    Select Case mstrField
        Case "destinyap"
            build.DestinyAP = val(mstrValue)
        Case "permdestinypoints"
            build.PermDestinyPoints = val(mstrValue)
        Case "destiny"
            build.Destinies = build.Destinies + 1
            ReDim Preserve build.Destiny(1 To build.Destinies)
            build.Destiny(build.Destinies).TreeName = mstrValue
            
        Case "ability"
            AddTreeAbility mstrValue, build.Destiny(build.Destinies), peDestiny
    End Select
End Sub

Private Sub AddTreeAbility(ByVal pstrRaw As String, ptypBuild As BuildTreeType, penType As PointerEnum)
    Dim typAbility As DeprecateAbilityType
    
    typAbility.PointerType = penType
    If Not ParseAbility(typAbility, ptypBuild.TreeName & " " & pstrRaw) Then
        DeprecateAbility typAbility
    ElseIf Not ResolveAbility(typAbility) Then
        DeprecateAbility typAbility
    Else
        AddBuildAbility ptypBuild, typAbility
    End If
End Sub

Private Function ParseAbility(ptypAbility As DeprecateAbilityType, ByVal pstrRaw As String) As Boolean
    Dim strNumber As String
    Dim lngPos As Long
    
    ptypAbility.Raw = pstrRaw
    lngPos = InStr(pstrRaw, "Tier ")
    If lngPos = 0 Then Exit Function
    strNumber = Mid$(pstrRaw, lngPos + 5, 1)
    If Not IsNumeric(strNumber) Then Exit Function
    ptypAbility.Tier = val(strNumber)
    If lngPos > 1 Then
        ptypAbility.TreeName = Left$(pstrRaw, lngPos - 2)
        ptypAbility.GuideTreeDisplay = ptypAbility.TreeName
    End If
    pstrRaw = Mid$(pstrRaw, lngPos + 8)
    lngPos = InStr(pstrRaw, "(Rank ")
    If lngPos = 0 Then
        ptypAbility.Rank = 1
    Else
        strNumber = Mid$(pstrRaw, lngPos + 6, 1)
        If Not IsNumeric(strNumber) Then Exit Function
        ptypAbility.Rank = val(strNumber)
        pstrRaw = Left$(pstrRaw, lngPos - 2)
    End If
    lngPos = InStr(pstrRaw, ": ")
    If lngPos = 0 Then
        ptypAbility.AbilityName = pstrRaw
    Else
        ptypAbility.AbilityName = Left$(pstrRaw, lngPos - 1)
        ptypAbility.SelectorName = Mid$(pstrRaw, lngPos + 2)
    End If
    ptypAbility.Parsed = True
    ParseAbility = True
End Function

Private Function ResolveAbility(ptypAbility As DeprecateAbilityType) As Boolean
    Dim lngTree As Long
    
    lngTree = SeekTree(ptypAbility.TreeName, ptypAbility.PointerType)
    If lngTree = 0 Then Exit Function
    Select Case ptypAbility.PointerType
        Case peEnhancement: ResolveAbility = ResolveAbilityTree(ptypAbility, db.Tree(lngTree))
        Case peDestiny: ResolveAbility = ResolveAbilityTree(ptypAbility, db.Destiny(lngTree))
    End Select
End Function

Private Function ResolveAbilityTree(ptypAbility As DeprecateAbilityType, ptypTree As TreeType) As Boolean
    If ptypAbility.Tier > ptypTree.Tiers Then Exit Function
    ResolveAbilityTree = FindAbilityOnTier(ptypAbility, ptypTree.Tier(ptypAbility.Tier))
End Function

Private Function FindAbilityOnTier(ptypAbility As DeprecateAbilityType, ptypTier As TierType) As Boolean
    Dim a As Long
    Dim s As Long
    
    For a = 1 To ptypTier.Abilities
        If ptypTier.Ability(a).AbilityName = ptypAbility.AbilityName Then
            If ptypTier.Ability(a).SelectorStyle = sseNone Then
                If Len(ptypAbility.SelectorName) <> 0 Then Exit Function Else Exit For
            Else
                If Len(ptypAbility.SelectorName) = 0 Then Exit Function
                For s = 1 To ptypTier.Ability(a).Selectors
                    If ptypTier.Ability(a).Selector(s).SelectorName = ptypAbility.SelectorName Then Exit For
                Next
                If s > ptypTier.Ability(a).Selectors Then Exit Function Else Exit For
            End If
        ' Some abilities have a colon in their name (eg: Warpriest's "Inflame: Energy Absorption")
        ' Parsing will split those into, for example: Ability=Inflame, Selector=Energy Absorption
        ElseIf ptypTier.Ability(a).AbilityName = ptypAbility.AbilityName & ": " & ptypAbility.SelectorName Then
            Exit For
        End If
    Next
    If a > ptypTier.Abilities Then Exit Function
    If ptypAbility.Rank > ptypTier.Ability(a).Ranks Then Exit Function
    ptypAbility.Ability = a
    ptypAbility.Selector = s
    FindAbilityOnTier = True
End Function

Private Sub AddBuildAbility(ptypBuild As BuildTreeType, ptypAbility As DeprecateAbilityType)
    With ptypBuild
        .Abilities = .Abilities + 1
        ReDim Preserve .Ability(1 To .Abilities)
        With .Ability(.Abilities)
            .Tier = ptypAbility.Tier
            .Ability = ptypAbility.Ability
            .Selector = ptypAbility.Selector
            .Rank = ptypAbility.Rank
        End With
    End With
End Sub


' ************* LOAD - LEVELING GUIDE *************


Private Sub LoadLevelingGuideText()
    Dim typAbility As DeprecateAbilityType
    Dim lngLookup As Long
    
    If gtypDeprecate.LevelingGuide.Deprecated Then Exit Sub
    If mstrField <> "guide" Then Exit Sub
    typAbility.Deprecated = True
    typAbility.PointerType = peEnhancement
    If mstrValue = "Bank AP" Then
        AddGuideID 200
    ElseIf mstrValue = "Reset All Trees" Then
        AddGuideID 100
    ElseIf Left$(mstrValue, 12) = "Reset Tree: " Then
        typAbility.GuideTreeDisplay = Mid$(mstrValue, 13)
        If Not GuideTreeLookup(typAbility, lngLookup) Then DeprecateLevelingGuide typAbility Else AddGuideID 100 + lngLookup
    Else
        If Not ParseAbility(typAbility, mstrValue) Then
            DeprecateLevelingGuide typAbility
        ElseIf Not GuideTreeLookup(typAbility, lngLookup) Then
            DeprecateLevelingGuide typAbility
        ElseIf Not ResolveAbility(typAbility) Then
            DeprecateLevelingGuide typAbility
        Else
            With build.Guide
                .Enhancements = .Enhancements + 1
                ReDim Preserve .Enhancement(1 To .Enhancements)
                With .Enhancement(.Enhancements)
                    .ID = lngLookup
                    .Tier = typAbility.Tier
                    .Ability = typAbility.Ability
                    .Selector = typAbility.Selector
                    .Rank = typAbility.Rank
                End With
            End With
        End If
    End If
End Sub

Private Sub AddGuideID(plngID As Long)
    With build.Guide
        .Enhancements = .Enhancements + 1
        ReDim Preserve .Enhancement(1 To .Enhancements)
        .Enhancement(.Enhancements).ID = plngID
    End With
End Sub

Private Function GuideTreeLookup(ptypAbility As DeprecateAbilityType, plngLookup As Long) As Boolean
    For plngLookup = mlngLookups To 1 Step -1
        If mtypLookup(plngLookup).Display = ptypAbility.GuideTreeDisplay Then Exit For
    Next
    If plngLookup = 0 Then plngLookup = AddTreeLookup(ptypAbility.GuideTreeDisplay)
    If plngLookup Then
        ptypAbility.GuideTreeID = plngLookup
        ptypAbility.TreeName = db.Tree(mtypLookup(plngLookup).Tree).TreeName
        GuideTreeLookup = True
    End If
End Function

Private Function AddTreeLookup(ByVal pstrDisplay As String) As Long
    Dim lngTree As Long
    Dim enClass As ClassEnum
    
    If Not FindTreeByAbbreviation(pstrDisplay, lngTree) Then
        If Mid$(pstrDisplay, Len(pstrDisplay) - 5, 2) <> " (" Or Right$(pstrDisplay, 1) <> ")" Then Exit Function
        If Not FindTreeByAbbreviation(Left$(pstrDisplay, Len(pstrDisplay) - 6), lngTree) Then Exit Function
        If Not FindClassByInitials(Mid$(pstrDisplay, Len(pstrDisplay) - 3, 3), enClass) Then Exit Function
    End If
    If Not GetGuideTreeClass(lngTree, enClass) Then Exit Function
    mlngLookups = mlngLookups + 1
    ReDim Preserve mtypLookup(1 To mlngLookups)
    With mtypLookup(mlngLookups)
        .Tree = lngTree
        .Class = enClass
        .Display = pstrDisplay
    End With
    With build.Guide
        .Trees = .Trees + 1
        ReDim Preserve .Tree(1 To .Trees)
        With .Tree(.Trees)
            .TreeName = db.Tree(lngTree).TreeName
            .Class = enClass
        End With
    End With
    AddTreeLookup = mlngLookups
End Function

Private Function FindTreeByAbbreviation(pstrAbbreviation As String, plngTree As Long) As Boolean
    Dim i As Long
    
    For plngTree = db.Trees To 1 Step -1
        If db.Tree(plngTree).Abbreviation = pstrAbbreviation Then Exit For
    Next
    FindTreeByAbbreviation = (plngTree <> 0)
End Function

Private Function FindClassByInitials(pstrInitials As String, penClass As ClassEnum) As Boolean
    Dim i As Long
    
    For penClass = ceClasses - 1 To 1 Step -1
        If db.Class(penClass).Initial(3) = pstrInitials Then Exit For
    Next
    FindClassByInitials = (penClass <> ceAny)
End Function

Private Function GetGuideTreeClass(plngTree As Long, penClass As ClassEnum) As Boolean
    Dim typClassSplit() As ClassSplitType
    Dim c As Long
    Dim t As Long
    
    If db.Tree(plngTree).TreeType <> tseClass Or penClass <> ceAny Then
        GetGuideTreeClass = True
    Else
        For c = 0 To GetClassSplit(typClassSplit) - 1
            penClass = typClassSplit(c).ClassID
            For t = 1 To db.Class(penClass).Trees
                If db.Class(penClass).Tree(t) = db.Tree(plngTree).TreeName Then
                    GetGuideTreeClass = True
                    Exit Function
                End If
            Next
        Next
    End If
End Function

