Attribute VB_Name = "basBuild"
' Written by Ellis Dee
' These are public functions used by the dialogs, the output routine and the import screen
Option Explicit

' Used for enhancement/destiny trees
Private Type LongType
    Value As Long
End Type

Private Type AbilityIDType
    Tier As Byte
    Ability As Byte
    Ranks As Byte
    Selector As Byte
End Type


' ************* POINTERS *************


Public Function PointerDisplay(ptypPointer As PointerType, pblnAbbreviate As Boolean, Optional plngTree As Long) As String
    With ptypPointer
        Select Case .Style
            Case peFeat: PointerDisplay = GetFeatDisplay(.Feat, .Selector, pblnAbbreviate, (plngTree <> 0))
            Case peEnhancement: PointerDisplay = GetEnhancementDisplay(ptypPointer, pblnAbbreviate, plngTree)
            Case peDestiny: PointerDisplay = GetDestinyDisplay(ptypPointer, pblnAbbreviate, plngTree)
        End Select
    End With
End Function

Public Function GetFeatDisplay(plngFeat As Long, ByVal plngSelector As Long, pblnAbbreviate As Boolean, pblnFeatPrefix As Boolean) As String
    Dim strFeat As String
    Dim strSelector As String
    Dim strDisplay As String
    
    If plngFeat = 0 Then Exit Function
    With db.Feat(plngFeat)
        If pblnAbbreviate Then strFeat = .Abbreviation Else strFeat = .FeatName
        If plngSelector <> 0 Then
            strSelector = .Selector(plngSelector).SelectorName
            If .SelectorOnly Then strDisplay = strSelector Else strDisplay = strFeat & ": " & strSelector
        Else
            strDisplay = strFeat
        End If
    End With
    If pblnFeatPrefix Then strDisplay = "Feat: " & strDisplay
    GetFeatDisplay = strDisplay
End Function

Private Function GetEnhancementDisplay(ptypPointer As PointerType, pblnAbbreviate As Boolean, plngTree As Long) As String
    Dim strTree As String
    Dim strTier As String
    Dim strAbility As String
    Dim strRank As String
    
    With ptypPointer
        If .Tree <> plngTree Then strTree = db.Tree(.Tree).Abbreviation & " "
        strTier = "Tier " & .Tier & ": "
        With db.Tree(.Tree).Tier(.Tier).Ability(.Ability)
            If ptypPointer.Selector = 0 Then
                If pblnAbbreviate Then strAbility = .Abbreviation Else strAbility = .AbilityName
            ElseIf .SelectorOnly Then
                strAbility = .Selector(ptypPointer.Selector).SelectorName
            Else
                If pblnAbbreviate Then strAbility = .Abbreviation Else strAbility = .AbilityName
                strAbility = strAbility & ": " & .Selector(ptypPointer.Selector).SelectorName
            End If
        End With
    End With
    If ptypPointer.Rank > 0 Then strRank = " (Rank " & ptypPointer.Rank & ")"
    GetEnhancementDisplay = strTree & strTier & strAbility & strRank
End Function

Private Function GetDestinyDisplay(ptypPointer As PointerType, pblnAbbreviate As Boolean, plngTree As Long) As String
    Dim strTree As String
    Dim strTier As String
    Dim strAbility As String
    Dim strRank As String
    
    With ptypPointer
        If .Tree <> plngTree Then strTree = db.Destiny(.Tree).Abbreviation & " "
        strTier = "Tier " & .Tier & ": "
        With db.Destiny(.Tree).Tier(.Tier).Ability(.Ability)
            If ptypPointer.Selector = 0 Then
                If pblnAbbreviate Then strAbility = .Abbreviation Else strAbility = .AbilityName
            ElseIf .SelectorOnly Then
                strAbility = .Selector(ptypPointer.Selector).SelectorName
            Else
                If pblnAbbreviate Then strAbility = .Abbreviation Else strAbility = .AbilityName
                strAbility = strAbility & ": " & .Selector(ptypPointer.Selector).SelectorName
            End If
        End With
    End With
    If ptypPointer.Rank > 0 Then strRank = " (Rank " & ptypPointer.Rank & ")"
    GetDestinyDisplay = strTree & strTier & strAbility & strRank
End Function


' ************* CALCULATIONS *************


Public Function HeroicLevels() As Long
    If build.MaxLevels > 19 Then HeroicLevels = 20 Else HeroicLevels = build.MaxLevels
End Function

Public Function GetClassSplit(ptypClassSplit() As ClassSplitType) As Long
    Dim lngNext As Long
    Dim strInitial() As String
    Dim lngInitial As Long
    Dim enColor As ColorValueEnum
    Dim lngSplash As Long
    Dim i As Long
    Dim j As Long
    
    ReDim ptypClassSplit(ceClasses - 1)
    ' Count how many levels taken in each class
    For i = 1 To HeroicLevels()
        ptypClassSplit(build.Class(i)).Levels = ptypClassSplit(build.Class(i)).Levels + 1
    Next
    ' Move classes taken to front of list
    For i = 1 To ceClasses - 1
        If ptypClassSplit(i).Levels <> 0 Then
            With ptypClassSplit(lngNext)
                .Levels = ptypClassSplit(i).Levels
                .ClassID = i
                .ClassName = db.Class(i).ClassName
                .Color = db.Class(i).Color
            End With
            lngNext = lngNext + 1
        End If
    Next
    ' Remove classes not taken
    If lngNext = 0 Then Exit Function
    ReDim Preserve ptypClassSplit(lngNext - 1)
    ' Sort the class manually (there's only 3 at most)
    If lngNext <> 1 Then ClassSplitCompare ptypClassSplit, 0, 1
    If lngNext = 3 Then
        ClassSplitCompare ptypClassSplit, 1, 2
        ClassSplitCompare ptypClassSplit, 0, 1
    End If
    ' Assign colors
    ptypClassSplit(0).Color = -1
    If lngNext = 3 Then
        For i = 0 To 2
            For j = 0 To 2
                If ptypClassSplit(i).ClassID = build.BuildClass(j) Then
                    ptypClassSplit(i).BuildClass = j
                    Exit For
                End If
            Next
        Next
        If ptypClassSplit(1).BuildClass > ptypClassSplit(2).BuildClass Then lngSplash = 1 Else lngSplash = 2
        ptypClassSplit(lngSplash).Color = ComplementColor(ptypClassSplit(3 - lngSplash).Color, ptypClassSplit(lngSplash).Color)
    End If
    ' Initials
    If lngNext > 1 Then
        If db.Class(ptypClassSplit(0).ClassID).Initial(0) = db.Class(ptypClassSplit(1).ClassID).Initial(0) Then lngInitial = 1
    End If
    If lngNext = 3 And lngInitial = 0 Then
        Select Case db.Class(ptypClassSplit(2).ClassID).Initial(0)
            Case db.Class(ptypClassSplit(0).ClassID).Initial(0), db.Class(ptypClassSplit(1).ClassID).Initial(0): lngInitial = 1
        End Select
    End If
    For i = 0 To lngNext - 1
        ptypClassSplit(i).Initial = db.Class(ptypClassSplit(i).ClassID).Initial(lngInitial)
    Next
    ' Return how many classes we have (remember that array is zero-based)
    GetClassSplit = lngNext
End Function

Private Sub ClassSplitCompare(ptypClassSplit() As ClassSplitType, plngOne As Long, plngTwo As Long)
    Dim typHold As ClassSplitType
    
    If ptypClassSplit(plngOne).Levels < ptypClassSplit(plngTwo).Levels Then
        typHold = ptypClassSplit(plngOne)
        ptypClassSplit(plngOne) = ptypClassSplit(plngTwo)
        ptypClassSplit(plngTwo) = typHold
    End If
End Sub

Public Function ComplementColor(penColor As ColorValueEnum, penPreferred As ColorValueEnum) As Long
    Select Case penColor
        Case cveRed
            Select Case penPreferred
                Case cveRed, cveOrange, cvePurple: ComplementColor = cveBlue
                Case Else: ComplementColor = penPreferred
            End Select
        Case cveBlue
            Select Case penPreferred
                Case cveBlue, cvePurple: ComplementColor = cveRed
                Case Else: ComplementColor = penPreferred
            End Select
        Case cveGreen
            Select Case penPreferred
                Case cveGreen, cveYellow: ComplementColor = cveOrange
                Case Else: ComplementColor = penPreferred
            End Select
        Case cveOrange
            Select Case penPreferred
                Case cveOrange, cveRed, cveYellow: ComplementColor = cveGreen
                Case Else: ComplementColor = penPreferred
            End Select
        Case cveYellow
            Select Case penPreferred
                Case cveYellow, cveOrange, cveGreen: ComplementColor = cvePurple
                Case Else: ComplementColor = penPreferred
            End Select
        Case cvePurple
            Select Case penPreferred
                Case cvePurple, cveBlue, cveRed: ComplementColor = cveYellow
                Case Else: ComplementColor = penPreferred
            End Select
    End Select
End Function

Public Function CalculateBaseStat(plngBase As Long, ByVal plngPoints As Long) As Long
    Select Case plngPoints
        Case 0 To 6: CalculateBaseStat = plngBase + plngPoints
        Case 8: CalculateBaseStat = plngBase + 7
        Case 10: CalculateBaseStat = plngBase + 8
        Case 13: CalculateBaseStat = plngBase + 9
        Case 16: CalculateBaseStat = plngBase + 10
    End Select
End Function

Public Function CalculateStat(penStat As StatEnum, ByVal plngLevel As Long, Optional pblnMod As Boolean = False, Optional pblnFred As Boolean = False) As Long
    Dim lngTotal As Long
    Dim lngTome As Long
    Dim i As Long
    
    If plngLevel = 0 Then plngLevel = 1 ' Only needed for Stat Schedule grid
    If build.Race = reAny Or penStat = aeAny Then Exit Function
    ' Base
    lngTotal = CalculateBaseStat(db.Race(build.Race).Stats(penStat), build.StatPoints(build.BuildPoints, penStat))
    ' Levelups
    lngTotal = lngTotal + CalculateLevelup(penStat, plngLevel)
    ' Tome
    lngTotal = lngTotal + CalculateTome(penStat, plngLevel, pblnFred)
    ' Return Mod or Stat
    If pblnMod Then
        ' Throwing away the decimal incorrectly converts -0.5 into 0, so subtract 11 instead of 10 to offset if stat < 10
        If lngTotal < 10 Then CalculateStat = (lngTotal - 11) \ 2 Else CalculateStat = (lngTotal - 10) \ 2
    Else
        CalculateStat = lngTotal
    End If
End Function

Public Function CalculateTome(penStat As StatEnum, plngLevel As Long, pblnFred As Boolean)
    Dim lngTome As Long
    
    If plngLevel > 1 Or pblnFred = True Then
        lngTome = TomeLevel(plngLevel)
        If build.Tome(penStat) < lngTome Then lngTome = build.Tome(penStat)
    End If
    CalculateTome = lngTome
End Function

Public Function CalculateLevelup(penStat As StatEnum, plngLevel As Long)
    Dim lngLevelUp As Long
    Dim i As Long
    
    For i = 1 To 7
        If plngLevel >= i * 4 And build.Levelups(i) = penStat Then lngLevelUp = lngLevelUp + 1
    Next
    CalculateLevelup = lngLevelUp
End Function

Public Function TomeLevel(ByVal plngLevel As Long) As Long
    Dim i As Long
    
    For i = tomes.Stat.Max To 1 Step -1
        If tomes.Stat.Level(i) > 0 And tomes.Stat.Level(i) <= plngLevel Then Exit For
    Next
    TomeLevel = i
End Function

Public Function CalculateMod(plngStat As Long) As Long
    CalculateMod = (plngStat - 10) \ 2
End Function

Public Function CalculateSkill(penSkill As SkillsEnum, ByVal plngLevel As Long, pblnSkillTome As Boolean) As Long
    Dim lngTotal As Long
    Dim lngMax As Long
    Dim lngTome As Long
    Dim i As Long
    
    lngMax = HeroicLevels()
    If lngMax > plngLevel Then lngMax = plngLevel
    For i = 1 To lngMax
        lngTotal = lngTotal + build.Skills(penSkill, i) * Skill.grid(penSkill, i).Native
    Next
    lngTotal = lngTotal \ 2
    ' Tome
    If pblnSkillTome Then
        lngTome = SkillTomeLevel(plngLevel)
        If build.SkillTome(penSkill) < lngTome Then lngTome = build.SkillTome(penSkill)
        lngTotal = lngTotal + lngTome
    End If
    CalculateSkill = lngTotal
End Function

Public Function SkillTomeLevel(ByVal plngLevel As Long) As Long
    Dim i As Long
    
    For i = tomes.Skill.Max To 1 Step -1
        If tomes.Skill.Level(i) > 0 And tomes.Skill.Level(i) <= plngLevel Then Exit For
    Next
    SkillTomeLevel = i
End Function

Public Sub CalculateBAB()
    Dim lngLevel() As Long
    Dim enClass As ClassEnum
    Dim enBAB As BABEnum
    Dim lngBAB As Long
    Dim i As Long
    
    ReDim lngLevel(ceClasses - 1)
    For i = 1 To 20
        enClass = build.Class(i)
        If enClass <> ceAny Then
            enBAB = db.Class(enClass).BAB
            lngLevel(enClass) = lngLevel(enClass) + 1
            If GetBAB(enBAB, lngLevel(enClass)) Then lngBAB = lngBAB + 1
            build.BAB(i) = lngBAB
        End If
    Next
    For i = 21 To MaxLevel
        If GetBAB(enBAB, i) Then lngBAB = lngBAB + 1
        build.BAB(i) = lngBAB
    Next
End Sub

Public Function GetBAB(penBAB As BABEnum, plngLevel As Long) As Boolean
    Select Case plngLevel
        Case 21, 23, 25, 27, 29: GetBAB = True
        Case 22, 24, 26, 28, 30: GetBAB = False
        Case Else
            Select Case penBAB
                Case beFull: GetBAB = True
                Case beThreeQuarters: If (plngLevel - 1) Mod 4 <> 0 Then GetBAB = True
                Case beHalf: If plngLevel Mod 2 = 0 Then GetBAB = True
            End Select
    End Select
End Function


' ************* SKILLS *************


Public Sub InitBuildSkills()
    Dim blnEverNative(1 To 21) As Boolean
    Dim lngSkillPoints As Long
    Dim lngSkill As Long '  Row
    Dim lngLevel As Long ' Column
    Dim typClassSplit() As ClassSplitType
    Dim lngClasses As Long
    Dim strInitial() As String
    Dim lngColor() As Long
    Dim blnThief As Boolean
    Dim i As Long
    
    ' Initialize array
    With Skill
        Erase .Row
        Erase .Col
        Erase .grid
    End With
    ' Identify Initials and Colos if multiclass
    lngClasses = GetClassSplit(typClassSplit)
    If lngClasses = 0 Then Exit Sub
    ReDim strInitial(ceClasses - 1)
    ReDim lngColor(ceClasses - 1)
    For i = 0 To lngClasses - 1
        With typClassSplit(i)
            strInitial(.ClassID) = .Initial
            If .Color = -1 Then lngColor(.ClassID) = -1 Else lngColor(.ClassID) = cfg.GetColor(cgeOutput, .Color)
        End With
    Next
    For lngLevel = 1 To HeroicLevels()
        With Skill.Col(lngLevel)
            .Class = build.Class(lngLevel)
            .Initial = strInitial(.Class)
            .Color = lngColor(.Class)
            If .Class = ceRogue Or .Class = ceArtificer Then blnThief = True
            .Thief = blnThief
        End With
    Next
    ' Calculate max ranks for each interior cell
    For lngSkill = 1 To 21
        For lngLevel = 1 To HeroicLevels()
            With Skill.grid(lngSkill, lngLevel)
                ' Is skill native for this level? Has it ever been native?
                If db.Class(build.Class(lngLevel)).NativeSkills(lngSkill) Then .Native = 2 Else .Native = 1
                If .Native = 2 Then blnEverNative(lngSkill) = True
                ' Ranks
                .Ranks = build.Skills(lngSkill, lngLevel) * .Native
                .MaxRanks = lngLevel + 3
                If blnEverNative(lngSkill) Then .MaxRanks = .MaxRanks * 2
            End With
        Next
    Next
    ' Calculate level totals
    For lngLevel = 1 To HeroicLevels()
        ' Points spent
        Skill.Col(lngLevel).Points = 0
        For lngSkill = 1 To 21
            Skill.Col(lngLevel).Points = Skill.Col(lngLevel).Points + build.Skills(lngSkill, lngLevel)
        Next
        lngSkillPoints = db.Class(build.Class(lngLevel)).SkillPoints + db.Race(build.Race).SkillPoints + CalculateStat(aeInt, lngLevel, True)
        If lngSkillPoints < 1 Then lngSkillPoints = 1
        If lngLevel = 1 Then lngSkillPoints = lngSkillPoints * 4
        Skill.Col(lngLevel).MaxPoints = lngSkillPoints
    Next
    ' Calculate skill totals
    For lngSkill = 1 To 21
        With Skill.Row(lngSkill)
            ' Ranks spent
            .Ranks = 0
            For lngLevel = 1 To HeroicLevels()
                .Ranks = .Ranks + build.Skills(lngSkill, lngLevel) * Skill.grid(lngSkill, lngLevel).Native
            Next
            ' Max allowed
            .MaxRanks = HeroicLevels() + 3
            If blnEverNative(lngSkill) Then .MaxRanks = .MaxRanks * 2
        End With
    Next
End Sub


' ************* FEATS *************


Public Sub InitBuildFeats()
    InitGrantedFeats
    InitStandardFeats
    InitLegendFeats
    InitRaceFeats
    InitClassFeats
    InitDeityFeats
    InitFeatList
End Sub

Private Sub InitGrantedFeats()
    Dim enClass As ClassEnum
    Dim lngClassLevel As Long
    Dim lngLevel As Long
    Dim lngGranted As Long
    Dim lngClassLevels(1 To 20)  As Long
    Dim blnDwarvenAxe As Boolean
    
    If build.Class(1) = ceAny Then Exit Sub
    If build.Race = reDwarf Then blnDwarvenAxe = True
    ' Allocate enough space that we only have to resize once at the end
    With build.Feat(bftGranted)
        ReDim .Feat(1 To 32)
        .Feats = 0
    End With
    For lngLevel = 1 To 20
        ' Race granted feats
        With db.Race(build.Race)
            For lngGranted = 1 To .GrantedFeats
                If .GrantedFeat(lngGranted).Tier = lngLevel Then GrantFeat .GrantedFeat(lngGranted), lngLevel, bfsRace, build.Feat(bftGranted).Feats
            Next
        End With
        ' Class granted feats
        enClass = build.Class(lngLevel)
        lngClassLevels(enClass) = lngClassLevels(enClass) + 1
        With db.Class(enClass)
            For lngGranted = 1 To .GrantedFeats
                If .GrantedFeat(lngGranted).Tier = lngClassLevels(enClass) Then GrantFeat .GrantedFeat(lngGranted), lngLevel, bfsClass, build.Feat(bftGranted).Feats
            Next
        End With
        ' Dwarven axes
        If blnDwarvenAxe Then
            Select Case build.Class(lngLevel)
                Case ceBarbarian, ceFighter, cePaladin, ceRanger
                    GrantDwarvenAxe lngLevel, build.Feat(bftGranted).Feats
                    blnDwarvenAxe = False
            End Select
        End If
    Next
    With build.Feat(bftGranted)
        If .Feats = 0 Then Erase .Feat Else ReDim Preserve .Feat(1 To .Feats)
    End With
End Sub

Private Sub GrantFeat(ptypPointer As PointerType, plngLevel As Long, penSource As BuildFeatSourceEnum, pbytIndex As Byte)
    Dim blnFound As Boolean
    Dim i As Long
    
    If ptypPointer.Feat = 0 Then Exit Sub
    ' Don't grant feats that are already granted
    For i = 1 To pbytIndex - 1
        With build.Feat(bftGranted).Feat(i)
            If .FeatName = db.Feat(ptypPointer.Feat).FeatName And .Selector = ptypPointer.Selector Then blnFound = True
        End With
        If blnFound Then Exit Sub
    Next
    ' Granted Feats store the level they're granted in Tier
    pbytIndex = pbytIndex + 1
    With build.Feat(bftGranted).Feat(pbytIndex)
        .FeatName = db.Feat(ptypPointer.Feat).FeatName
        .Selector = ptypPointer.Selector
        .Source = penSource
        .Level = plngLevel
        .Type = bftGranted
    End With
End Sub

Private Sub GrantDwarvenAxe(plngLevel As Long, pbytIndex As Byte)
    Dim typPointer As PointerType
    Dim lngFeat As Long
    Dim i As Long
    
    lngFeat = SeekFeat("Exotic Weapon")
    If lngFeat = 0 Then Exit Sub
    With db.Feat(lngFeat)
        For i = 1 To .Selectors
            If .Selector(i).SelectorName = "Dwarven Axe" Then
                typPointer.Feat = lngFeat
                typPointer.Selector = i
                typPointer.Style = peFeat
                GrantFeat typPointer, plngLevel, bfsRace, pbytIndex
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitStandardFeats()
    Dim lngIndex As Long
    
    AllocateFeatSlots bftStandard, 14
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 1
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 3
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 6
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 9
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 12
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 15
    InitBuildFeatSlot bftStandard, bfsHeroic, lngIndex, 18
    InitBuildFeatSlot bftStandard, bfsEpic, lngIndex, 21
    InitBuildFeatSlot bftStandard, bfsEpic, lngIndex, 24
    InitBuildFeatSlot bftStandard, bfsDestiny, lngIndex, 26
    InitBuildFeatSlot bftStandard, bfsEpic, lngIndex, 27
    InitBuildFeatSlot bftStandard, bfsDestiny, lngIndex, 28
    InitBuildFeatSlot bftStandard, bfsDestiny, lngIndex, 29
    InitBuildFeatSlot bftStandard, bfsEpic, lngIndex, 30
End Sub

Private Sub InitLegendFeats()
    AllocateFeatSlots bftLegend, 1
    InitBuildFeatSlot bftLegend, bfsLegend, 0, 30
End Sub

Private Sub AllocateFeatSlots(penType As BuildFeatTypeEnum, plngFeats As Long, Optional pblnEraseFirst As Boolean = False)
    With build.Feat(penType)
        If pblnEraseFirst Then
            .Feats = 0
            Erase .Feat
        End If
        If .Feats <> plngFeats Then
            If plngFeats = 0 Then Erase .Feat Else ReDim Preserve .Feat(1 To plngFeats)
            .Feats = plngFeats
        End If
    End With
End Sub

Private Sub InitBuildFeatSlot(penType As BuildFeatTypeEnum, penSource As BuildFeatSourceEnum, plngIndex As Long, plngCharacterLevel As Long, Optional plngClassLevel As Long)
    Dim lngFeat As Long
    Dim enClass As ClassEnum
    Dim blnValidClassFeat As Boolean
    
    plngIndex = plngIndex + 1
    With build.Feat(penType)
        If .Feats < plngIndex Then
            .Feats = plngIndex
            ReDim Preserve .Feat(1 To .Feats)
        End If
        With .Feat(plngIndex)
            .Level = plngCharacterLevel
            .ClassLevel = plngClassLevel
            .Type = penType
            .Source = penSource
            ' Blank out any illegal class/race bonus feats
            Select Case penType
                ' Class
                Case bftClass1, bftClass2, bftClass3
                    lngFeat = SeekFeat(.FeatName)
                    enClass = build.BuildClass(penType - bftClass1)
                    If enClass <> ceAny And lngFeat <> 0 Then
                        If db.Feat(lngFeat).ClassBonus(enClass) Then
                            blnValidClassFeat = True
                        ElseIf db.Feat(lngFeat).ClassOnly Then
                            blnValidClassFeat = db.Feat(lngFeat).ClassOnlyClasses(enClass)
                        End If
                        If Not blnValidClassFeat Then
                            .FeatName = vbNullString
                            .Selector = 0
                        End If
                    End If
                ' Race
                Case bftRace
                    lngFeat = SeekFeat(.FeatName)
                    If build.Race <> reAny And lngFeat <> 0 Then
                        If Not db.Feat(lngFeat).RaceBonus(build.Race) Then
                            .FeatName = vbNullString
                            .Selector = 0
                        End If
                    End If
            End Select
        End With
    End With
End Sub

Private Sub InitRaceFeats()
    Dim i As Long
    
    If db.Race(build.Race).BonusFeat Then
        AllocateFeatSlots bftRace, 1
        InitBuildFeatSlot bftRace, bfsRace, 0, 1
    ElseIf build.Feat(bftRace).Feats = 1 Then
        build.Feat(bftRace).Feat(1).Level = 99
    End If
End Sub

Private Sub InitClassFeats()
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngClassLevel As Long
    Dim lngClasses As Long
    Dim lngFeats As Long
    Dim lngIndex As Long
    Dim enSource As BuildFeatSourceEnum
    Dim typTemp As BuildFeatListType
    Dim i As Long
    
    For i = 0 To 2
        enClass = build.BuildClass(i)
        If enClass = ceAny Then Exit For
        lngIndex = 0
        ' Make a temporary copy
        typTemp = build.Feat(bftClass1 + i)
        ' Define class feat slots
        AllocateFeatSlots bftClass1 + i, ClassBonusFeatCount(enClass), True
        ' Apply class levels
        For lngClassLevel = 1 To 20
            enSource = IsClassBonusFeat(enClass, lngClassLevel)
            If enSource <> 0 Then
                lngLevel = GetCharacterLevel(enClass, lngClassLevel)
                InitBuildFeatSlot bftClass1 + i, enSource, lngIndex, lngLevel, lngClassLevel
            End If
        Next
        ' Now go through and copy in the original feats to the appropriate new slots
        MapClassFeats bftClass1 + i, typTemp
    Next
End Sub

Private Function ClassBonusFeatCount(penClass As ClassEnum) As Long
    Dim lngReturn As Long
    Dim i As Long
    
    For i = 1 To 20
        If IsClassBonusFeat(penClass, i) <> 0 Then lngReturn = lngReturn + 1
    Next
    ClassBonusFeatCount = lngReturn
End Function

Private Function IsClassBonusFeat(penClass As ClassEnum, plngClassLevel As Long) As BuildFeatSourceEnum
    Dim enSource As BuildFeatSourceEnum
    
    enSource = db.Class(penClass).BonusFeat(plngClassLevel)
    Select Case enSource
        Case bfsClass, bfsClassOnly: IsClassBonusFeat = enSource
        Case Else: Exit Function
    End Select
    Select Case plngClassLevel
        Case 1, 3, 6, 12, 20
        Case Else: Exit Function
    End Select
    Select Case penClass
        Case ceCleric, ceFavoredSoul, cePaladin: IsClassBonusFeat = False
    End Select
End Function

Private Function GetCharacterLevel(penClass, plngClassLevel) As Long
    Dim lngLevel As Long
    Dim i As Long
    
    For i = 1 To HeroicLevels()
        If build.Class(i) = penClass Then lngLevel = lngLevel + 1
        If lngLevel = plngClassLevel Then
            GetCharacterLevel = i
            Exit Function
        End If
    Next
    GetCharacterLevel = 99
End Function

Private Sub MapClassFeats(penType As BuildFeatTypeEnum, ptypTemp As BuildFeatListType)
    Dim blnExitFor As Boolean
    Dim i As Long
    Dim j As Long
    
    For i = 1 To ptypTemp.Feats
        blnExitFor = False
        For j = 1 To build.Feat(penType).Feats
            With build.Feat(penType).Feat(j)
                If .Level = ptypTemp.Feat(i).Level And .ClassLevel = ptypTemp.Feat(i).ClassLevel Then
                    .FeatName = ptypTemp.Feat(i).FeatName
                    .Selector = ptypTemp.Feat(i).Selector
                    blnExitFor = True
                End If
            End With
            If blnExitFor Then Exit For
        Next
    Next
End Sub

' Note: The commented out "lngCount = lngCount + 1" lines
' ended up double-booking each deity feat, leaving a blank slot
' before each actual slot and causing the array to be twice as big
' as needed. The CompactDeityFeats() function corrects this
' for builds saved with v2.4 and earlier.
Private Function InitDeityFeats()
    Dim lngClassLevels() As Long
    Dim blnOne As Boolean
    Dim blnSix As Boolean
    Dim lngLevel As Long
    Dim lngCount As Long
    Dim enClass As ClassEnum
    
    CompactDeityFeats
    ReDim lngClassLevels(1 To ceClasses - 1)
    For lngLevel = 1 To 20
        enClass = build.Class(lngLevel)
        If enClass <> ceAny Then lngClassLevels(enClass) = lngClassLevels(enClass) + 1
        Select Case enClass
            Case ceCleric, ceFavoredSoul, cePaladin
                Select Case lngClassLevels(enClass)
                    Case 1
                        If Not blnOne Then
'                            lngCount = lngCount + 1
                            InitBuildFeatSlot bftDeity, bfsDeity, lngCount, lngLevel, lngClassLevels(enClass)
                            blnOne = True
                        End If
                    Case 6
                        If Not blnSix Then
'                            lngCount = lngCount + 1
                            InitBuildFeatSlot bftDeity, bfsDeity, lngCount, lngLevel, lngClassLevels(enClass)
                            blnSix = True
                        End If
                    Case 20 ' 3, 12, 20
                        If enClass = ceFavoredSoul Then
'                            lngCount = lngCount + 1
                            InitBuildFeatSlot bftDeity, bfsDeity, lngCount, lngLevel, lngClassLevels(enClass)
                        End If
                End Select
        End Select
    Next
    With build.Feat(bftDeity)
        If .Feats <> lngCount Then
            .Feats = lngCount
            If .Feats = 0 Then Erase .Feat Else ReDim Preserve .Feat(1 To .Feats)
        End If
    End With
End Function

' Up through v2.4, deity feats were being double-initialized, resulting in
' the deity feat array being twice as big as needed and having a blank
' feat slot added before every actual slot. This didn't cause problems
' because blank slots are ignored, but it's sloppy.
' To fix it, we need to move all the actual slots to the front of the array,
' moving all the blank slots to the end of the array. Initialization will
' then redim away all the blanks.
Private Sub CompactDeityFeats()
    Dim typBlank As BuildFeatType
    Dim i As Long
    Dim j As Long
    
    With build.Feat(bftDeity)
        For i = 1 To .Feats
            If .Feat(i).ClassLevel = 3 Or .Feat(i).ClassLevel = 12 Then .Feat(i).Level = 0
        Next
        For i = 1 To .Feats
            If j = 0 And .Feat(i).Level = 0 Then
                j = i
            ElseIf j <> 0 And .Feat(i).Level <> 0 Then
                .Feat(j) = .Feat(i)
                .Feat(i) = typBlank
                For j = j + 1 To i
                    If .Feat(j).Level = 0 Then Exit For
                Next
            End If
        Next
    End With
End Sub

' Returns new index in Feat.List()
Public Function AddSpecialFeat(penSpecialType As BuildFeatTypeEnum, penParentType As BuildFeatTypeEnum, plngParentIndex As Long) As Long
    Dim typTemp As BuildFeatType
    Dim i As Long
    Dim j As Long
    
    With build.Feat(penParentType).Feat(plngParentIndex)
        typTemp.Child = plngParentIndex
        typTemp.ChildType = penParentType
        typTemp.ClassLevel = .ClassLevel
        typTemp.FeatName = vbNullString
        If penSpecialType = bftAlternate Then typTemp.Level = .Level Else typTemp.Level = glngLevel
        typTemp.Selector = 0
        typTemp.Source = .Source
        typTemp.Type = penSpecialType
    End With
    With build.Feat(penSpecialType)
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats) = typTemp
        For i = .Feats To 2 Step -1
            If .Feat(i).Level < .Feat(i - 1).Level Then .Feat(i) = .Feat(i - 1) Else Exit For
        Next
        .Feat(i) = typTemp
    End With
    AddSpecialFeat = InitFeatList(penSpecialType, i)
End Function

Public Sub DeleteFeatSlot(penType As BuildFeatTypeEnum, plngIndex As Long)
    Dim i As Long
    
    With build.Feat(penType)
        For i = plngIndex To .Feats - 1
            .Feat(i) = .Feat(i + 1)
        Next
        .Feats = .Feats - 1
        If .Feats = 0 Then Erase .Feat Else ReDim Preserve .Feat(1 To .Feats)
    End With
    InitFeatList
End Sub

Public Function InitFeatList(Optional penNewType As BuildFeatTypeEnum, Optional plngNewIndex As Long) As Long
    Dim typBlank As FeatListType
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim lngReturn As Single
    Dim lngNew As Single
    
    Feat = typBlank
    ReDim Feat.ChannelCount(fceChannels)
    If build.Class(1) = ceAny Then Exit Function
    ReDim Feat.List(1 To 48)
    For enType = bftGranted To bftExchange
        For lngIndex = 1 To build.Feat(enType).Feats
            AddFeatToList enType, lngIndex, (enType = penNewType And lngIndex = plngNewIndex)
        Next
    Next
    ReDim Preserve Feat.List(1 To Feat.Count)
    SortFeatList
    SlotIndexes
    CheckSlotErrors
    For lngIndex = 1 To Feat.Count
        ' Return index for the newly added exchange slot
        If Feat.List(lngIndex).Flag Then InitFeatList = lngIndex
    Next
    InitFeatChannels
End Function

Private Sub SlotIndexes()
    Dim lngGranted As Long
    Dim lngSelected As Long
    Dim i As Long
    Dim j As Long

    With Feat
        For i = 1 To .Count
            With .List(i)
'                If .ActualType = bftGranted Then
'                    lngGranted = lngGranted + 1
'                    .Slot = lngGranted
'                Else
'                    lngSelected = lngSelected + 1
'                    .Slot = lngSelected
'                End If
                If .ActualType = bftExchange Then
                    For j = 1 To Feat.Count
                        If Feat.List(j).ActualType = .EffectiveType And Feat.List(j).Index = .ParentIndex Then
                            .ExchangeIndex = j
                            .SourceFilter = Feat.List(j).SourceFilter
                            Feat.List(j).ExchangeIndex = i
                            Exit For
                        End If
                    Next
                End If
            End With
        Next
    End With
End Sub

'Private Sub SlotIndexes()
'    Dim lngGranted As Long
'    Dim lngSelected As Long
'    Dim i As Long
'    Dim j As Long
'
'    With Feat
'        For i = 1 To .Count
'            With .List(i)
'                If .ActualType = bftGranted Then
'                    lngGranted = lngGranted + 1
'                    .Slot = lngGranted
'                Else
'                    lngSelected = lngSelected + 1
'                    .Slot = lngSelected
'                End If
'                If .ActualType = bftExchange Then
'                    For j = 1 To Feat.Count
'                        If Feat.List(j).ActualType = .EffectiveType And Feat.List(j).Index = .ParentIndex Then
'                            .ExchangeIndex = j
'                            .SourceFilter = Feat.List(j).SourceFilter
'                            Feat.List(j).ExchangeIndex = i
'                            Exit For
'                        End If
'                    Next
'                End If
'            End With
'        Next
'    End With
'End Sub

Private Sub InitFeatChannels()
    Dim enClass As ClassEnum
    Dim lngClassLevel As Long
    Dim i As Long
    
    For i = 1 To Feat.Count
        With Feat.List(i)
            If .ActualType = bftExchange Then
                enClass = Feat.List(.ExchangeIndex).Class
                lngClassLevel = Feat.List(.ExchangeIndex).ClassLevel
            Else
                enClass = .Class
                lngClassLevel = .ClassLevel
            End If
            ReDim .ChannelSlot(fceChannels - 1)
            .Channel = fceGeneral
            Select Case .EffectiveType
                Case bftGranted: .Channel = fceGranted
                Case bftRace
                    Select Case build.Race
                        Case reHalfElf, reDragonborn, reAasimar: .Channel = fceRacial
                    End Select
                Case bftDeity: .Channel = fceDeity
                Case bftClass1, bftClass2, bftClass3
                    Select Case enClass
                        Case ceCleric: .Channel = fceCleric
                        Case ceDruid: .Channel = fceWildShape
                        Case ceFavoredSoul
                            Select Case lngClassLevel
                                Case 5, 10, 15: .Channel = fceEnergy
                                Case 2, 7: .Channel = fceFavoredSoul
                            End Select
                        Case ceMonk: If lngClassLevel = 3 Then .Channel = fceMonk
                        Case ceRanger: .Channel = fceFavoredEnemy
                        Case ceRogue: .Channel = fceRogue
                        Case ceWarlock: .Channel = fceWarlock
                    End Select
            End Select
            Feat.ChannelCount(.Channel) = Feat.ChannelCount(.Channel) + 1
            .ChannelSlot(.Channel) = Feat.ChannelCount(.Channel)
            If .Channel <> fceGranted Then
                Feat.ChannelCount(fceSelected) = Feat.ChannelCount(fceSelected) + 1
                .ChannelSlot(fceSelected) = Feat.ChannelCount(fceSelected)
            End If
        End With
    Next
End Sub

Private Function AddFeatToList(penType As BuildFeatTypeEnum, plngIndex As Long, pblnFlag As Boolean) As Long
    Dim typBuild As BuildFeatType
    Dim typNew As FeatDetailType
    Dim lngFeat As Long
    Dim strDisplay As String
    Dim lngLevel As Long
    
    typBuild = build.Feat(penType).Feat(plngIndex)
    If typBuild.Level < 1 Or typBuild.Level > build.MaxLevels Then Exit Function
    With typNew
        .Flag = pblnFlag
        .ActualType = penType
        If penType = bftAlternate Or penType = bftExchange Then
            .EffectiveType = typBuild.ChildType
            .ParentIndex = typBuild.Child
        Else
            .EffectiveType = .ActualType
        End If
        .Index = plngIndex
        .Level = typBuild.Level
        If .Level > 0 And .Level < 21 Then
            .Class = build.Class(.Level)
            .ClassLevel = GetClassLevel(.Class, .Level)
        End If
        .FeatName = typBuild.FeatName
        .Selector = typBuild.Selector
        If Len(.FeatName) Then .FeatID = SeekFeat(.FeatName)
        GetDisplayNames typNew
        lngLevel = typBuild.Level
        Select Case penType
            Case bftAlternate
                .SourceForm = "Alternate"
                .SourceOutput = " OR "
            Case bftExchange
                .SourceForm = "Exchange"
                .SourceOutput = "Swap"
                lngLevel = build.Feat(typBuild.ChildType).Feat(typBuild.Child).Level
            Case Else
                Select Case typBuild.Source
                    Case bfsHeroic
                    Case bfsEpic: .SourceForm = "Epic"
                    Case bfsDestiny: .SourceForm = "Destiny"
                    Case bfsRace: .SourceForm = db.Race(build.Race).Abbreviation
                    Case bfsClass, bfsClassOnly: If penType = bftGranted Then .SourceForm = db.Class(build.Class(.Level)).ClassName Else .SourceForm = db.Class(build.Class(.Level)).Initial(2)
                    Case bfsDeity: .SourceForm = "Deity"
                    Case bfsLegend
                        .SourceForm = "Legendary"
                        .SourceOutput = "Legend"
                End Select
                If typBuild.Source <> bfsLegend Then .SourceOutput = .SourceForm
        End Select
        Select Case typBuild.Source
            Case bfsHeroic: .SourceFilter = "Level " & .Level
            Case bfsEpic: .SourceFilter = "Epic " & .Level
            Case bfsDestiny: .SourceFilter = "Destiny " & .Level
            Case bfsRace: .SourceFilter = db.Race(build.Race).Abbreviation
            Case bfsClass, bfsClassOnly, bfsDeity: .SourceFilter = db.Class(build.Class(lngLevel)).Initial(2) & " " & .ClassLevel
            Case bfsLegend: .SourceFilter = "Legend 30"
        End Select
    End With
    With Feat
        If penType = bftGranted Then
            .Granted = .Granted + 1
        Else
            .Selected = .Selected + 1
        End If
        .Count = .Count + 1
        If .Count > UBound(.List) Then ReDim Preserve .List(1 To .Count + 12)
        .List(.Count) = typNew
    End With
End Function

Public Sub GetDisplayNames(ptypDetail As FeatDetailType)
    Dim strFeat As String
    Dim strSelector As String
    Dim strDisplay As String
    
    If ptypDetail.FeatID = 0 Then
        ptypDetail.Display = ptypDetail.FeatName
        If ptypDetail.ActualType = bftAlternate Then ptypDetail.DisplayAlternate = ptypDetail.FeatName
        Exit Sub
    End If
    With db.Feat(ptypDetail.FeatID)
        strFeat = .FeatName
        If ptypDetail.Selector <> 0 Then
            strSelector = .Selector(ptypDetail.Selector).SelectorName
            If .SelectorOnly Then strDisplay = strSelector Else strDisplay = strFeat & ": " & strSelector
        Else
            strDisplay = strFeat
        End If
    End With
    With ptypDetail
        .FeatName = strFeat
        .Display = strDisplay
        If .ActualType = bftAlternate Then
            If strFeat = build.Feat(.EffectiveType).Feat(.ParentIndex).FeatName Then
                .DisplayAlternate = strSelector
            Else
                .DisplayAlternate = strDisplay
            End If
        End If
    End With
End Sub

Private Function GetClassLevel(penClass As ClassEnum, plngLevel As Long) As Long
    Dim lngReturn As Long
    Dim i As Long
    
    For i = 1 To plngLevel
        If build.Class(i) = penClass Then lngReturn = lngReturn + 1
    Next
    GetClassLevel = lngReturn
End Function

Private Sub SortFeatList()
    Dim i As Long
    Dim j As Long
    Dim typSwap As FeatDetailType
    
    With Feat
        ' Sort by level
        For i = 2 To .Count
            typSwap = .List(i)
            For j = i To 2 Step -1
                If CompareSlot(typSwap, .List(j - 1)) = -1 Then .List(j) = .List(j - 1) Else Exit For
            Next j
            .List(j) = typSwap
        Next
    End With
End Sub

Private Function CompareSlot(ptypSlot1 As FeatDetailType, ptypSlot2 As FeatDetailType) As Long
    If ptypSlot1.Level < ptypSlot2.Level Then
        CompareSlot = -1
    ElseIf ptypSlot1.Level > ptypSlot2.Level Then
        CompareSlot = 1
    ElseIf ptypSlot1.EffectiveType < ptypSlot2.EffectiveType Then
        CompareSlot = -1
    ElseIf ptypSlot1.EffectiveType > ptypSlot2.EffectiveType Then
        CompareSlot = 1
    Else
        CompareSlot = 0
    End If
End Function

Public Sub CheckSlotErrors()
    Dim i As Long
    
    For i = 1 To Feat.Count
        With Feat.List(i)
            .ErrorState = False
            .ErrorText = vbNullString
            If .ActualType <> bftGranted Then
                If .ActualType = bftExchange Then
                    .ErrorState = SlotLocked(i)
                    .ErrorText = gstrError
                End If
                If Not .ErrorState Then
                    .ErrorState = (CheckFeatSlot(db.Feat(.FeatID), build.Feat(.ActualType).Feat(.Index)) <> dsCanDrop)
                    .ErrorText = gstrError
                End If
            End If
        End With
    Next
End Sub

Public Function CheckFeatSlot(ptypFeat As FeatType, ptypSlot As BuildFeatType, Optional pblnSelectorRequired As Boolean = True) As DropStateEnum
    Dim blnSelector As Boolean
    Dim typSlot As BuildFeatType
    Dim lngActualLevel As Long
    Dim enClass As ClassEnum
    Dim blnFred As Boolean
    Dim s As Long
    Dim i As Long
    
    gstrError = vbNullString
    CheckFeatSlot = dsDefault
    If Len(ptypFeat.FeatName) = 0 Then
        CheckFeatSlot = dsCanDrop
        Exit Function
    End If
    ' Apply target slot info to exchanged feats
    typSlot = ptypSlot
    lngActualLevel = typSlot.Level
    If typSlot.Type = bftExchange Then
        typSlot = build.Feat(typSlot.ChildType).Feat(typSlot.Child)
        typSlot.FeatName = ptypSlot.FeatName
        typSlot.Selector = ptypSlot.Selector
        blnFred = True
    End If
    ' Unselectable (and not granted)
    If ptypFeat.Selectable = False And typSlot.Type <> bftGranted Then
        gstrError = "Not a selectable feat."
        Exit Function
    End If
    ' Granted by class levels
    If ptypFeat.GrantedBy.Class <> ceAny Then
        enClass = ptypFeat.GrantedBy.Class
        s = 0
        For i = 1 To 20
            If build.Class(i) = enClass Then s = s + 1
        Next
        If s >= ptypFeat.GrantedBy.ClassLevels Then
            gstrError = "Chosen at " & GetClassName(enClass) & " level " & ptypFeat.GrantedBy.ClassLevels
            Exit Function
        End If
    End If
    ' Invalid selector?
    If pblnSelectorRequired Then
        If ptypFeat.SelectorStyle <> sseNone Then
            If typSlot.Selector = 0 Then
                gstrError = "A selector is required but wasn't specified."
                Exit Function
            End If
            s = typSlot.Selector
            blnSelector = True
        End If
    End If
    ' Slot type
    Select Case typSlot.Source
        Case bfsHeroic
            If Not ptypFeat.Group(feHeroic) Then
                gstrError = "Can't be taken as a Heroic feat."
                Exit Function
            End If
        Case bfsEpic
            If Not (ptypFeat.Group(feHeroic) Or ptypFeat.Group(feEpic)) Then
                gstrError = "Can't be taken as an Epic feat."
                Exit Function
            End If
        Case bfsDestiny
            If Not ptypFeat.Group(feDestiny) Then
                gstrError = "Can't be taken as a Destiny feat."
                Exit Function
            End If
        Case bfsLegend
            If Not ptypFeat.Group(feLegend) Then
                gstrError = "Can't be taken as a Legendary feat."
                Exit Function
            End If
        Case bfsRace
            If Not ptypFeat.RaceBonus(build.Race) Then
                gstrError = "Not a valid " & db.Race(build.Race).Abbreviation & " bonus feat."
                Exit Function
            End If
        Case bfsDeity
            If Not ptypFeat.Deity Then
                gstrError = "Can't be taken as a Deity feat."
                Exit Function
            End If
        Case bfsClassOnly
            If Not ptypFeat.ClassOnly Then
                gstrError = "Can't be taken as a class feat."
                Exit Function
            End If
        Case bfsClass
            If typSlot.Level > 20 Then
                gstrError = "Class feats must be taken during class levels."
                Exit Function
            End If
            enClass = build.Class(typSlot.Level)
            If blnSelector Then
                If Not (ptypFeat.Selector(s).ClassBonus(enClass) Or ptypFeat.ClassOnly) Then
                    gstrError = "Not a valid " & db.Class(enClass).Abbreviation & " bonus feat. (This may not apply to all selectors.)"
                    Exit Function
                End If
            Else
                If Not (ptypFeat.ClassBonus(enClass) Or ptypFeat.ClassOnly) Then
                    gstrError = "Not a valid " & db.Class(enClass).Abbreviation & " bonus feat."
                    Exit Function
                End If
            End If
    End Select
    ' Class only
    If ptypFeat.ClassOnly Then
        If typSlot.Level > 20 Then
            gstrError = "Can't be taken in epic levels."
            Exit Function
        End If
        If Not ptypFeat.ClassOnlyClasses(build.Class(typSlot.Level)) Then
            gstrError = "Can only be taken by certain classes as a class bonus feat."
            Exit Function
        End If
        If Not ptypFeat.ClassOnlyLevels(typSlot.ClassLevel) Then
            gstrError = "Can only be taken at certain class levels."
            Exit Function
        End If
    End If
    ' Level
    If typSlot.Level < ptypFeat.Level Or typSlot.Level > build.MaxLevels Then
        gstrError = "Requires " & ptypFeat.Level & " character levels."
        Exit Function
    End If
    ' BAB
    If build.BAB(typSlot.Level) < ptypFeat.BAB Then
        gstrError = "Requires BAB " & ptypFeat.BAB & "."
        Exit Function
    End If
    ' Race only
    If ptypFeat.RaceOnly And typSlot.Source <> bfsRace Then
        gstrError = "Can only be taken as a racial bonus feat."
        Exit Function
    End If
    ' Class Bonus Level
    If ptypFeat.ClassBonusLevel.Class <> ceAny Then
        If typSlot.Source = bfsClass Then
            enClass = build.Class(typSlot.Level)
            If enClass = ptypFeat.ClassBonusLevel.Class And typSlot.ClassLevel <> ptypFeat.ClassBonusLevel.ClassLevels Then
                gstrError = "Available as " & GetClassName(enClass) & " bonus feat only at " & GetClassName(enClass) & " level " & ptypFeat.ClassBonusLevel.ClassLevels
                Exit Function
            End If
        End If
    End If
    ' Can cast spells?
    If ptypFeat.CanCastSpell Then
        If build.CanCastSpell(ptypFeat.CanCastSpellLevel) = 0 Or build.CanCastSpell(ptypFeat.CanCastSpellLevel) > typSlot.Level Then
            If ptypFeat.CanCastSpellLevel = 0 Then
                gstrError = "Can't cast healing spells yet."
            Else
                gstrError = "Can't cast Level " & ptypFeat.CanCastSpellLevel & " spells yet."
            End If
            Exit Function
        End If
    End If
    ' Class Level (unlike other checks, both the feat and the selector versions should be checked)
    If ptypFeat.Class(0) Then
        If Not CheckClassLevels(typSlot.Level, ptypFeat.Class, ptypFeat.ClassLevel) Then
            gstrError = "Does not meet class level requirements."
            Exit Function
        End If
    End If
    If blnSelector Then
        If ptypFeat.Selector(s).Class(0) Then
            If Not CheckClassLevels(typSlot.Level, ptypFeat.Selector(s).Class, ptypFeat.Selector(s).ClassLevel) Then
                gstrError = "Does not meet class level requirements."
                Exit Function
            End If
        End If
    End If
    ' Race restricted
    If blnSelector Then
        If RaceRestricted(ptypFeat.Selector(s).Race) Then
            gstrError = "Race restricted. (May not apply to all selectors.)"
            Exit Function
        End If
    Else
        If RaceRestricted(ptypFeat.Race) Then
            gstrError = "Race restricted."
            Exit Function
        End If
    End If
    ' Some checks don't apply to Alternate feats
    If typSlot.Type <> bftAlternate Then
        ' Selectors reqs
        If blnSelector Then
            ' Stat
            If ptypFeat.Selector(s).Stat <> aeAny Then
                If CalculateStat(ptypFeat.Selector(s).Stat, typSlot.Level, False, blnFred) < ptypFeat.Selector(s).StatValue Then
                    gstrError = "Requires " & GetStatName(ptypFeat.Selector(s).Stat) & " " & ptypFeat.Selector(s).StatValue & " (May not apply to all selectors.)"
                    Exit Function
                End If
            End If
            ' Skill
            If ptypFeat.Selector(s).Skill <> seAny Then
                If CalculateSkill(ptypFeat.Selector(s).Skill, typSlot.Level, ptypFeat.SkillTome) < ptypFeat.Selector(s).SkillValue Then
                    gstrError = "Requires " & GetSkillName(ptypFeat.Selector(s).Skill, False) & " " & ptypFeat.Selector(s).SkillValue & " (skill tomes "
                    If ptypFeat.SkillTome Then gstrError = gstrError & "apply)" Else gstrError = gstrError & "do not apply)"
                    Exit Function
                End If
            End If
        Else ' Not a selector feat
            ' Stat
            If ptypFeat.Stat <> aeAny Then
                If CalculateStat(ptypFeat.Stat, typSlot.Level, False, blnFred) < ptypFeat.StatValue Then
                    gstrError = "Requires " & GetStatName(ptypFeat.Stat) & " " & ptypFeat.StatValue
                    Exit Function
                End If
            End If
            ' Skill
            If ptypFeat.Skill <> seAny Then
                If CalculateSkill(ptypFeat.Skill, typSlot.Level, ptypFeat.SkillTome) < ptypFeat.SkillValue Then
                    gstrError = "Requires " & GetSkillName(ptypFeat.Skill, False) & " " & ptypFeat.SkillValue & " (skill tomes "
                    If ptypFeat.SkillTome Then gstrError = gstrError & "apply)" Else gstrError = gstrError & "do not apply)"
                    Exit Function
                End If
            End If
        End If
        ' Past lives
        If ptypFeat.PastLife And build.BuildPoints < beHero Then
            gstrError = "This is a first life build."
            Exit Function
        End If
        If ptypFeat.Legend And build.BuildPoints < beLegend Then
            gstrError = "This is not a legend build."
            Exit Function
        End If
    End If
    ' Passed the hard reqs, now check soft (ignore these for alternate feats)
    If typSlot.Type = bftAlternate Then
        CheckFeatSlot = dsCanDrop
    Else
        CheckFeatSlot = dsCanDropError
        ' Alignment
        If blnSelector Then
            If ptypFeat.Selector(s).Alignment(0) Then
                If Not ptypFeat.Selector(s).Alignment(build.Alignment) Then
                    gstrError = "Incompatible alignment."
                    Exit Function
                End If
            End If
        End If
        ' Feat reqs (also selector reqs)
        If CheckFeatReqs(ptypFeat, typSlot, lngActualLevel) Then CheckFeatSlot = dsCanDrop
    End If
    ' NotClass
    ' For Pacts and Domains that can't be taken by lawful characters, prevent those choices when monk levels taken
    If blnSelector Then
        enClass = ptypFeat.Selector(s).NotClass
        If enClass <> ceAny Then
            For i = 1 To 20
                If build.Class(i) = enClass Then Exit For
            Next
            If i <= 20 Then
                gstrError = "Cannot be taken by characters with " & db.Class(enClass).ClassName & " levels."
                CheckFeatSlot = dsCanDropError
                Exit Function
            End If
        End If
    End If
End Function

' Returns TRUE if not allowed
Public Function RaceRestricted(plngRace() As Long) As Boolean
    RaceRestricted = True
    Select Case plngRace(0)
        Case rreAny: RaceRestricted = False
        Case rreRequired: If plngRace(build.Race) = 1 Then RaceRestricted = False
        Case rreNotAllowed: If plngRace(build.Race) = 0 Then RaceRestricted = False
        Case rreStandard: If (db.Race(build.Race).Type = rteFree Or db.Race(build.Race).Type = rtePremium) And plngRace(build.Race) = 0 Then RaceRestricted = False
        Case rreIconic: If db.Race(build.Race).Type = rteIconic And plngRace(build.Race) = 0 Then RaceRestricted = False
    End Select
End Function

Public Function CheckClassLevels(ByVal plngLevel As Long, pblnClass() As Boolean, plngClassLevel() As Long) As Boolean
    Dim lngClassLevels() As Long
    Dim lngMax As Long
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    
    ' Count class levels taken by the level of this slot
    If plngLevel > 20 Then lngMax = 20 Else lngMax = plngLevel
    ReDim lngClassLevels(1 To ceClasses - 1)
    For lngLevel = 1 To lngMax
        lngClassLevels(build.Class(lngLevel)) = lngClassLevels(build.Class(lngLevel)) + 1
    Next
    ' Now check to see if we have enough levels for any required class
    For enClass = 1 To ceClasses - 1
        If pblnClass(enClass) Then
            If lngClassLevels(enClass) >= plngClassLevel(enClass) Then
                CheckClassLevels = True
                Exit Function
            End If
        End If
    Next
End Function

' Returns TRUE if we successfully pass all featreqs
Public Function CheckFeatReqs(ptypFeat As FeatType, ptypSlot As BuildFeatType, plngLevel As Long) As Boolean
    Dim typTaken() As FeatTakenType
    Dim enReqGroup As ReqGroupEnum
    
    IdentifyTakenFeats typTaken, plngLevel
    If ptypFeat.SelectorStyle = sseShared And ptypSlot.Selector <> 0 And ptypFeat.FeatName = ptypSlot.FeatName Then
        If typTaken(ptypFeat.Parent.Feat).Selector(ptypSlot.Selector) = False Then
            gstrError = "Requires " & GetFeatDisplay(ptypFeat.Parent.Feat, ptypSlot.Selector, True, False)
            Exit Function
        End If
    End If
    For enReqGroup = rgeAll To rgeNone
        If CheckFeatReq(ptypFeat.Req(enReqGroup), enReqGroup, ptypSlot.Level, typTaken) Then Exit Function
    Next
    ' Check selector reqs if (and only if) the feat in ptypFeat is already slotted into ptypSlot
    If ptypSlot.Selector <> 0 And ptypFeat.SelectorStyle <> sseNone Then
        If ptypFeat.FeatName = ptypSlot.FeatName Then
            For enReqGroup = rgeAll To rgeNone
                If CheckFeatReq(ptypFeat.Selector(ptypSlot.Selector).Req(enReqGroup), enReqGroup, ptypSlot.Level, typTaken) Then Exit Function
            Next
        End If
    End If
    CheckFeatReqs = True
End Function

' Returns TRUE if we fail this featreq
Public Function CheckFeatReq(ptypReq As ReqListType, penReqGroup As ReqGroupEnum, ByVal plngLevel As Long, ptypTaken() As FeatTakenType) As Boolean
    Dim lngMatches As Long
    Dim strTaken As String
    Dim strMissing As String
    Dim i As Long

    If ptypReq.Reqs = 0 Then Exit Function
    For i = 1 To ptypReq.Reqs
        With ptypReq.Req(i)
            If .Selector = 0 Then
                If ptypTaken(.Feat).Times <> 0 Then
                    If Len(strTaken) = 0 Then strTaken = db.Feat(.Feat).Abbreviation
                    lngMatches = lngMatches + 1
                Else
                    If Len(strMissing) = 0 Then strMissing = db.Feat(.Feat).Abbreviation
                End If
            Else
                If ptypTaken(.Feat).Selector(.Selector) Then
                    If Len(strTaken) = 0 Then strTaken = db.Feat(.Feat).Abbreviation
                    lngMatches = lngMatches + 1
                Else
                    If Len(strMissing) = 0 Then strMissing = db.Feat(.Feat).Abbreviation
                End If
            End If
        End With
    Next
    Select Case penReqGroup
        Case rgeAll
            If lngMatches < ptypReq.Reqs Then
                gstrError = "Requires " & strMissing
                CheckFeatReq = True
            End If
        Case rgeOne
            If lngMatches < 1 Then
                gstrError = "Requires a feat from the 'One of' list."
                CheckFeatReq = True
            End If
        Case rgeNone
            If lngMatches > 0 Then
                gstrError = "Antireq for " & strTaken
                CheckFeatReq = True
            End If
    End Select
End Function

Public Sub IdentifyTakenFeats(ptypTaken() As FeatTakenType, ByVal plngLevel As Long)
    Dim lngFeat As Long
    Dim lngExchange As Long
    Dim lngSelector As Long
    Dim lngMaxLevel As Long
    Dim i As Long
    Dim s As Long
    
    If plngLevel = 0 Then lngMaxLevel = build.MaxLevels Else lngMaxLevel = plngLevel
    ' Create a map of all feats and selectors
    ReDim ptypTaken(db.Feats)
    For i = 1 To db.Feats
        If db.Feat(i).Selectors <> 0 Then
            With ptypTaken(i)
                ReDim .Selector(db.Feat(i).Selectors)
            End With
        End If
    Next
    ' Flag all taken feats and selectors
    For i = 1 To Feat.Count
        ' We're done if we outleveled the level range
        If Feat.List(i).Level > lngMaxLevel Then Exit Sub
        ' Ignore blank slots, alternate feats, and exchange feat sources (exchanged feats are tracked from their target slots only)
        If Feat.List(i).FeatID <> 0 And Feat.List(i).ActualType <> bftAlternate And Feat.List(i).ActualType <> bftExchange Then
            ' Start with the feat and selector of this literal slot
            lngFeat = Feat.List(i).FeatID
            lngSelector = Feat.List(i).Selector
            ' Exchange feat
            If Feat.List(i).ExchangeIndex <> 0 Then
                lngExchange = Feat.List(i).ExchangeIndex
                ' Only exchange feat if the exchange happens inside the specified level range
                If Feat.List(lngExchange).Level <= lngMaxLevel Then
                    lngFeat = Feat.List(lngExchange).FeatID
                    lngSelector = Feat.List(lngExchange).Selector
                End If
            End If
            ' Now that we know our effective feat and selector (after any exchanges), count it as taken
            With ptypTaken(lngFeat)
                .Times = .Times + 1
                If lngSelector Then
                    ' Is this selector "All" selectors?
                    If db.Feat(lngFeat).Selector(lngSelector).All Then
                        For s = 1 To UBound(.Selector)
                            If .Selector(s) = False Then
                                .Selector(s) = True
                                .SelectorsTaken = .SelectorsTaken + 1
                            End If
                        Next
                    Else
                        If .Selector(lngSelector) = False Then
                            .Selector(lngSelector) = True
                            .SelectorsTaken = .SelectorsTaken + 1
                        End If
                    End If
                End If
            End With
        End If
    Next
End Sub

Public Function SlotLocked(plngExchangeSlot As Long) As Boolean
    Dim lngLevel As Long
    Dim lngSlotIndex As Long
    Dim lngCheck As Long
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim lngReq As Long
    Dim blnLocked As Boolean
    Dim i As Long
    
    If plngExchangeSlot = 0 Then Exit Function
    lngLevel = Feat.List(plngExchangeSlot).Level
    lngSlotIndex = Feat.List(plngExchangeSlot).ExchangeIndex
    If lngSlotIndex = 0 Then Exit Function
    Do
        lngFeat = Feat.List(lngSlotIndex).FeatID
        If lngFeat = 0 Then
            gstrError = "Nothing to exchange."
            Exit Do
        End If
        lngSelector = Feat.List(lngSlotIndex).Selector
        If Feat.List(lngSlotIndex).ActualType = bftDeity Then
            gstrError = "Deity feats cannot be exchanged."
            Exit Do
        End If
        If db.Feat(lngFeat).Pact Then
            gstrError = "Warlock Pacts cannot be exchanged."
            Exit Do
        End If
        If db.Feat(lngFeat).Domain Then
            gstrError = "Cleric Domains cannot be exchanged."
            Exit Do
        End If
        ' Check "All"
        For i = lngSlotIndex + 1 To Feat.Count
            If Feat.List(i).Level > lngLevel Then Exit For
            If Feat.List(i).ActualType <> bftGranted And Feat.List(i).ActualType <> bftAlternate Then
                lngCheck = Feat.List(i).FeatID
                If lngCheck <> 0 Then
                    With db.Feat(lngCheck)
                        For lngReq = 1 To .Req(rgeAll).Reqs
                            With .Req(rgeAll).Req(lngReq)
                                If .Feat = lngFeat Then
                                    If db.Feat(lngFeat).SelectorStyle = sseShared Then
                                        If .Selector = lngSelector Then blnLocked = True
                                    ElseIf .Selector = 0 Or .Selector = lngSelector Then
                                        blnLocked = True
                                    End If
                                End If
                            End With
                            If blnLocked Then Exit For
                        Next
                    End With
                End If
                If blnLocked Then Exit For
            End If
        Next
        If blnLocked Then
            gstrError = Feat.List(lngSlotIndex).Display & " is locked by " & Feat.List(i).Display
            Exit Do
        End If
        ' Check "One"
        For i = lngSlotIndex + 1 To Feat.Count
            If Feat.List(i).Level > lngLevel Then Exit For
            If Feat.List(i).ActualType <> bftGranted Then
                If OneCount(Feat.List(i).FeatID, Feat.List(i).Selector, Feat.List(i).Level, lngFeat, lngSelector) Then
                    gstrError = Feat.List(lngSlotIndex).Display & " is locked by " & Feat.List(i).Display
                    Exit Do
                End If
            End If
        Next
        ' Check repeats (can't exchange toughness unless it's last toughness taken)
        If db.Feat(lngFeat).Times > 1 Then
            For i = lngSlotIndex + 1 To Feat.Count
                If Feat.List(i).Level > lngLevel Then Exit For
                Select Case Feat.List(i).ActualType
                    Case bftGranted, bftAlternate
                    Case bftExchange
                        If Feat.List(Feat.List(i).ExchangeIndex).Level > Feat.List(lngSlotIndex).Level Then
                            If Feat.List(Feat.List(i).ExchangeIndex).FeatID = lngFeat Then
                                gstrError = "Only the last " & Feat.List(lngSlotIndex).Display & " can be exchanged."
                                Exit Do
                            End If
                        End If
                    Case Else
                        If Feat.List(i).FeatID = lngFeat Then
                            gstrError = "Only the last " & Feat.List(lngSlotIndex).Display & " can be exchanged."
                            Exit Do
                        End If
                End Select
            Next
        End If
        Exit Function
    Loop Until True
    SlotLocked = True
End Function

Private Function OneCount(plngOneFeat As Long, plngOneSelector As Long, plngOneLevel, plngFeat As Long, plngSelector As Long) As Boolean
    Dim lngCount As Long
    Dim blnFound As Boolean
    Dim lngMaxLevel As Long
    Dim i As Long
    Dim r As Long
    
    If plngOneFeat = 0 Then Exit Function
    With db.Feat(plngOneFeat).Req(rgeOne)
        If .Reqs <> 0 Then
            For i = 1 To .Reqs
                If .Req(i).Feat = plngFeat Then
                    If db.Feat(plngOneFeat).SelectorStyle = sseShared Then
                        If .Req(i).Selector = plngSelector Then
                            blnFound = True
                            Exit For
                        End If
                    ElseIf .Req(i).Selector = 0 Or .Req(i).Selector = plngSelector Then
                        blnFound = True
                        Exit For
                    End If
                End If
            Next
        End If
        If blnFound Then
            For i = 1 To Feat.Count
                If Feat.List(i).ActualType <> bftAlternate And Feat.List(i).Level <= plngOneLevel Then
                    For r = 1 To .Reqs
                        If Feat.List(i).FeatID = .Req(r).Feat Then
                            If .Req(r).Selector = 0 Or .Req(r).Selector = Feat.List(i).Selector Then lngCount = lngCount + 1
                        End If
                    Next
                End If
            Next
            If lngCount < 2 Then OneCount = True
        End If
    End With
End Function


' ************* SPELLS *************


Public Sub InitBuildSpells()
    Dim enClass As ClassEnum
    Dim lngSpellLevel As Long
    
    If build.Class(1) = ceAny Then Exit Sub
    ' Calculate class levels and CanCastSpell info
    InitBuildSpellsClassLevels
    ' Define spell slot arrays for each class
    For enClass = 1 To ceClasses - 1
        With build.Spell(enClass)
            If .ClassLevels > 0 And db.Class(enClass).CanCastSpell(1) > 0 Then
                ' Calculate MaxSpellLevel
                InitBuildSpellsMaxSpellLevel enClass, build.Spell(enClass)
                ' Spell slots
                For lngSpellLevel = 1 To .MaxSpellLevel
                    ' Define slots
                    InitBuildSpellsSlots enClass, .ClassLevels, lngSpellLevel, .Level(lngSpellLevel)
                    ' Populate slots
                    InitBuildSpellsSlotLevels enClass, lngSpellLevel, .Level(lngSpellLevel), .ClassLevels
                Next
            End If
        End With
    Next
    ' Add mandatory spells
    If build.Spell(ceCleric).ClassLevels > 0 Then AddClericCures
    If build.Spell(ceWarlock).ClassLevels > 1 Then AddWarlockPactSpells
End Sub

Private Sub InitBuildSpellsClassLevels()
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngSpellLevel As Long
    
    ' Initialize CanCastSpell and spell class levels
    Erase build.CanCastSpell
    ReDim Preserve build.Spell(1 To ceClasses - 1)
    For enClass = 1 To ceClasses - 1
        build.Spell(enClass).Class = enClass
        build.Spell(enClass).ClassLevels = 0
        build.Spell(enClass).MaxSpellLevel = 0
    Next
    For lngLevel = 1 To HeroicLevels()
        enClass = build.Class(lngLevel)
        ' Class levels
        build.Spell(enClass).ClassLevels = build.Spell(enClass).ClassLevels + 1
        For lngSpellLevel = 0 To 9
            If build.CanCastSpell(lngSpellLevel) = 0 Then
                ' CanCastSpell
                If build.Spell(enClass).ClassLevels = db.Class(enClass).CanCastSpell(lngSpellLevel) Then build.CanCastSpell(lngSpellLevel) = lngLevel
            End If
        Next
    Next
    ' Enable Edit menu
    frmMain.mnuEdit(4).Enabled = (build.CanCastSpell(1) <> 0)
End Sub

Private Sub InitBuildSpellsMaxSpellLevel(penClass As ClassEnum, ptypClass As BuildClassSpellType)
    Dim lngSpellLevel As Long
    
    With ptypClass
        For lngSpellLevel = db.Class(penClass).MaxSpellLevel To 1 Step -1
            If db.Class(penClass).SpellSlots(.ClassLevels, lngSpellLevel) > 0 Then
                .MaxSpellLevel = lngSpellLevel
                ReDim Preserve .Level(1 To lngSpellLevel)
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitBuildSpellsSlots(penClass As ClassEnum, ByVal plngClassLevels As Long, plngSpellLevel As Long, ptypLevel As BuildClassLevelType)
    With ptypLevel
        ' Base slots
        .BaseSlots = db.Class(penClass).SpellSlots(plngClassLevels, plngSpellLevel)
        ' Add mandatory slots
        .Mandatory = sseStandard
        Select Case penClass
            Case ceCleric: If plngSpellLevel < 9 Then .Mandatory = sseClericCure
            Case ceWarlock: If plngClassLevels >= WarlockPactLevel(plngSpellLevel) Then .Mandatory = sseWarlockPact
        End Select
        If .Mandatory <> sseStandard Then .BaseSlots = .BaseSlots + 1
        ' Create spell slots
        .Slots = .BaseSlots + .FreeSlots
        ReDim Preserve .Slot(1 To .Slots)
    End With
End Sub

Private Sub InitBuildSpellsSlotLevels(penClass As ClassEnum, plngSpellLevel As Long, ptypLevel As BuildClassLevelType, ByVal plngMaxClassLevels As Long)
    Dim enSlotType As SpellSlotEnum
    Dim lngEffectiveSlot As Long
    Dim lngSlot As Long
    Dim lngStep As Long
    Dim lngClassLevel As Long
    
    With ptypLevel
        lngEffectiveSlot = 1
        For lngSlot = 1 To .Slots
            lngStep = 0
            ' Slot Type
            If lngSlot <= .FreeSlots Then
                enSlotType = sseFree
            ElseIf lngSlot = 1 And .Mandatory = sseClericCure Then
                enSlotType = sseClericCure
            ElseIf lngSlot = .Slots And .Mandatory = sseWarlockPact Then
                enSlotType = sseWarlockPact
            Else
                enSlotType = sseStandard
                lngStep = 1
            End If
            .Slot(lngSlot).SlotType = enSlotType
            ' Class level this slot is acquired
            If enSlotType = sseWarlockPact Then
                lngClassLevel = WarlockPactLevel(plngSpellLevel)
            Else
                lngClassLevel = GetSpellSlotClassLevel(penClass, plngSpellLevel, lngEffectiveSlot, plngMaxClassLevels)
            End If
            ' Build level when we reach that many class levels
            .Slot(lngSlot).Level = GetCharacterLevel(penClass, lngClassLevel)
            ' Increment slot counter, ignoring free and mandatory slots
            lngEffectiveSlot = lngEffectiveSlot + lngStep
        Next
    End With
End Sub

Private Function GetSpellSlotClassLevel(penClass As ClassEnum, plngSpellLevel As Long, plngSlot As Long, ByVal plngMaxClassLevels As Long) As Long
    Dim lngClassLevel As Long
    
    For lngClassLevel = 1 To plngMaxClassLevels
        If db.Class(penClass).SpellSlots(lngClassLevel, plngSpellLevel) >= plngSlot Then
            GetSpellSlotClassLevel = lngClassLevel
            Exit Function
        End If
    Next
End Function

Private Sub AddClericCures()
    Dim lngLevel As Long
    Dim lngMax As Long
    
    With build.Spell(ceCleric)
        If .MaxSpellLevel > 7 Then lngMax = 8 Else lngMax = .MaxSpellLevel
        For lngLevel = 1 To lngMax
            .Level(lngLevel).Slot(1).Spell = db.Class(ceCleric).MandatorySpell(lngLevel)
        Next
    End With
End Sub

Private Sub AddWarlockPactSpells()
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim strPact As String
    Dim lngPact As Long
    Dim i As Long
    
    ' Find chosen pact in feats list
    For i = 1 To Feat.Count
        lngFeat = Feat.List(i).FeatID
        lngSelector = Feat.List(i).Selector
        If lngFeat <> 0 And lngSelector <> 0 Then
            If db.Feat(lngFeat).FeatName = "Pact" Then
                strPact = db.Feat(lngFeat).Selector(lngSelector).SelectorName
                Exit For
            End If
        End If
    Next
    If Len(strPact) Then
        ' Identify pact index
        With db.Class(ceWarlock)
            For lngPact = 1 To .Pacts
                If .Pact(lngPact).PactName = strPact Then Exit For
            Next
            If lngPact > .Pacts Then lngPact = 0
        End With
    End If
    ' Apply pact spells to any/all pact slots, or blank out pact slots if no feat was chosen
    With build.Spell(ceWarlock)
        For i = 1 To .MaxSpellLevel
            With .Level(i)
                If .Slots > 0 Then
                    With .Slot(.Slots)
                        If .SlotType = sseWarlockPact Then
                            If lngPact = 0 Then .Spell = vbNullString Else .Spell = db.Class(ceWarlock).Pact(lngPact).Spells(i)
                        End If
                    End With
                End If
            End With
        Next
    End With
End Sub

Public Sub AddFreeSpellSlot(penClass As ClassEnum, plngSpellLevel As Long)
    Dim i As Long
    
    With build.Spell(penClass).Level(plngSpellLevel)
        .FreeSlots = .FreeSlots + 1
        .Slots = .BaseSlots + .FreeSlots
        ReDim Preserve .Slot(1 To .Slots)
        For i = .Slots To 2 Step -1
            .Slot(i) = .Slot(i - 1)
        Next
        .Slot(1).SlotType = sseFree
        .Slot(1).Spell = vbNullString
        .Slot(1).Level = GetSpellSlotClassLevel(penClass, plngSpellLevel, 1, build.Spell(penClass).ClassLevels)
    End With
End Sub

Public Sub RemoveEmptyFreeSlots(penClass As ClassEnum, plngSpellLevel As Long)
    Dim lngSlot As Long
    Dim i As Long
    
    With build.Spell(penClass).Level(plngSpellLevel)
        Do
            lngSlot = lngSlot + 1
            If .Slot(lngSlot).SlotType = sseFree And Len(.Slot(lngSlot).Spell) = 0 Then
                For i = lngSlot To .Slots - 1
                    .Slot(i) = .Slot(i + 1)
                Next
                .FreeSlots = .FreeSlots - 1
                .Slots = .BaseSlots + .FreeSlots
            End If
        Loop While lngSlot < .Slots
        ReDim Preserve .Slot(1 To .Slots)
    End With
End Sub

Public Function WarlockPactLevel(plngSpellLevel As Long) As Long
    Select Case plngSpellLevel
        Case 1: WarlockPactLevel = 2
        Case 2: WarlockPactLevel = 5
        Case 3: WarlockPactLevel = 9
        Case 4: WarlockPactLevel = 14
        Case 5: WarlockPactLevel = 17
        Case 6: WarlockPactLevel = 19
    End Select
End Function

Public Function CheckFree(penClass As ClassEnum, pstrSpell As String) As Boolean
    Dim i As Long

    With db.Class(penClass)
        For i = 1 To .FreeSpells
            If .FreeSpell(i) = pstrSpell Then
                CheckFree = True
                Exit For
            End If
        Next
    End With
End Function


' ************* RACIAL AP *************

'calculate the Enhancement tree available/spent - all the plngs are altered in this sub
Public Sub GetPointsSpentAndMax(plngSpentBase As Long, plngSpentRPLBonus As Long, plngSpentUniBonus As Long, plngMaxBase As Long, plngMaxRPLBonus As Long, plngMaxUniBonus As Long)
    Dim lngTree As Long
    Dim lngBuildTree As Long
    Dim lngAbility As Long
    Dim lngPoints As Long
    
    plngSpentBase = 0
    plngSpentRPLBonus = 0
    plngSpentUniBonus = 0
    'evaluate each tree
    For lngBuildTree = 1 To build.Trees
        With build.Tree(lngBuildTree)
            lngTree = SeekTree(.TreeName, peEnhancement)
            For lngAbility = 1 To .Abilities
                With .Ability(lngAbility)
                    If .Ability <> 0 Then
                        'Get our points spent on each ability in a tree
                        lngPoints = GetPoints(db.Tree(lngTree).Tier(.Tier).Ability(.Ability), .Selector, .Rank)
                        'Pull from appropriate bonus pool first
                        If db.Tree(lngTree).TreeType = tseRace Then
                            plngSpentRPLBonus = plngSpentRPLBonus + lngPoints
                        ElseIf db.Tree(lngTree).TreeType = tseGlobal Then
                           plngSpentUniBonus = plngSpentUniBonus + lngPoints
                         Else
                            plngSpentBase = plngSpentBase + lngPoints
                        End If
                    End If
                End With
            Next
        End With
    Next
    'set our max spend
    If plngSpentRPLBonus > build.RacialAP Then
        plngSpentBase = plngSpentBase + plngSpentRPLBonus - build.RacialAP
        plngSpentRPLBonus = build.RacialAP
    End If
    If plngSpentUniBonus > build.UniversalAP Then
        plngSpentBase = plngSpentBase + plngSpentUniBonus - build.UniversalAP
        plngSpentUniBonus = build.UniversalAP
    End If
    plngMaxBase = HeroicLevels() * 4
    plngMaxRPLBonus = build.RacialAP
    plngMaxUniBonus = build.UniversalAP
End Sub


' ************* ENHANCEMENTS *************


Public Sub InitBuildTrees()
    Dim strRace As String
    Dim lngClassLevels() As Long
    Dim enClass As ClassEnum
    Dim lngTreeLevels() As Long
    Dim lngTree As Long
    Dim lngBuildTree As Long
    Dim i As Long
    Dim j As Long
    
    ReDim lngTreeLevels(db.Trees)
    ' Flag racial tree, adding it as a build tree if necessary
    If build.Race <> reAny Then
        strRace = db.Race(build.Race).RaceName
        lngTree = SeekTree(strRace, peEnhancement)
        lngTreeLevels(lngTree) = build.MaxLevels
        If FindBuildTree(strRace) = 0 Then
            With build
                .Trees = .Trees + 1
                ReDim Preserve .Tree(1 To .Trees)
                With .Tree(.Trees)
                    .TreeName = strRace
                    .TreeType = tseRace
                End With
            End With
        End If
    End If
    ' Count up class levels taken
    ReDim lngClassLevels(ceClasses - 1)
    For i = 1 To HeroicLevels()
        lngClassLevels(build.Class(i)) = lngClassLevels(build.Class(i)) + 1
    Next
    ' Flag all class trees as valid options
    For i = 0 To 2
        enClass = build.BuildClass(i)
        If enClass <> ceAny Then
            For j = 1 To db.Class(enClass).Trees
                lngTree = SeekTree(db.Class(enClass).Tree(j), peEnhancement)
                If lngTreeLevels(lngTree) < lngClassLevels(enClass) Then lngTreeLevels(lngTree) = lngClassLevels(enClass)
            Next
        End If
    Next
    ' Now check build trees, updating class levels and deleting invalid trees
    For i = build.Trees To 1 Step -1
        lngTree = SeekTree(build.Tree(i).TreeName, peEnhancement)
        ' Set "class" levels for Racial Class trees (Elf-AA) and Global trees (Harper)
        If lngTree <> 0 Then
            Select Case db.Tree(lngTree).TreeType
                Case tseRaceClass, tseGlobal: lngTreeLevels(lngTree) = build.MaxLevels
            End Select
        End If
        If lngTree = 0 Or lngTreeLevels(lngTree) = 0 Then
            For j = i To build.Trees - 1
                build.Tree(j) = build.Tree(j + 1)
            Next
            build.Trees = build.Trees - 1
            If build.Trees = 0 Then Erase build.Tree Else ReDim Preserve build.Tree(1 To build.Trees)
        Else
            build.Tree(i).ClassLevels = lngTreeLevels(lngTree)
        End If
    Next
End Sub

Public Function FindBuildTree(pstrTreeName As String) As Long
    Dim i As Long
    
    For i = 1 To build.Trees
        If build.Tree(i).TreeName = pstrTreeName Then
            FindBuildTree = i
            Exit Function
        End If
    Next
End Function


' ************* DESTINY *************


Public Function GetTwistDisplayName(plngTwist As Long) As String
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim strDisplay As String
    Dim strSelector As String
    
    If build.Twist(plngTwist).Tier < 1 Then Exit Function
    With build.Twist(plngTwist)
        lngDestiny = SeekTree(.DestinyName, peDestiny)
        lngTier = .Tier
        lngAbility = .Ability
        lngSelector = .Selector
    End With
    If lngDestiny > db.Destinies Then Exit Function
    With db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility)
        strDisplay = .AbilityName
        If lngSelector Then
            strSelector = db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility).Selector(lngSelector).SelectorName
            If .SelectorOnly Then strDisplay = strSelector Else strDisplay = strDisplay & ": " & strSelector
        End If
    End With
    GetTwistDisplayName = strDisplay
End Function

Public Function CalculateFatePoints(plngTwist As Long, ByVal plngTier As Long) As Long
    Dim lngTotal As Long
    Dim lngStep As Long
    Dim i As Long
    
    lngStep = plngTwist
    For i = 1 To plngTier
        lngTotal = lngTotal + lngStep
        lngStep = lngStep + 1
    Next
    CalculateFatePoints = lngTotal
End Function

Public Function MaxFatePoints() As Long
    Select Case build.MaxLevels
        Case 20 To 28: MaxFatePoints = 32
        Case 29: MaxFatePoints = 34
        Case 30: MaxFatePoints = 37
    End Select
End Function

Public Function MaxTwistSlots() As Long
    If build.MaxLevels = 30 Then MaxTwistSlots = 5 Else MaxTwistSlots = 4
End Function


' ************* TREES *************


Public Function CostDescrip(ptypAbility As AbilityType, plngSelector As Long, Optional pblnPerRank As Boolean = True) As String
    Dim lngCost As Long
    Dim i As Long
    
    If plngSelector = 0 Then
        lngCost = ptypAbility.Cost
        For i = 1 To ptypAbility.Selectors
            If ptypAbility.Selector(i).Cost <> lngCost Then
                CostDescrip = "Cost: Varies"
                Exit Function
            End If
        Next
    Else
        lngCost = ptypAbility.Selector(plngSelector).Cost
    End If
    If ptypAbility.Ranks > 1 And pblnPerRank Then
        CostDescrip = "Cost: " & lngCost & " AP per rank"
    Else
        CostDescrip = "Cost: " & lngCost & " AP"
    End If
End Function

Public Function GetInsertionPoint(ptypBuildTree As BuildTreeType, ByVal plngTier As Long, ByVal plngAbility As Long) As Long
    Dim i As Long
    
    With ptypBuildTree
        For i = 1 To .Abilities
            If .Ability(i).Tier > plngTier Then
                Exit For
            ElseIf .Ability(i).Tier = plngTier And .Ability(i).Ability >= plngAbility Then
                Exit For
            End If
        Next
    End With
    GetInsertionPoint = i
End Function

Public Sub RemoveBlanks(ptypBuildTree As BuildTreeType)
    Dim i As Long
    Dim j As Long
    
    With ptypBuildTree
        For i = 1 To .Abilities
            If .Ability(i).Ability <> 0 Then
                j = j + 1
                If j <> i Then .Ability(j) = .Ability(i)
            End If
        Next
        If .Abilities <> j Then
            .Abilities = j
            If .Abilities = 0 Then Erase .Ability Else ReDim Preserve .Ability(1 To .Abilities)
        End If
    End With
End Sub

Public Function AbilityTaken(ptypBuildTree As BuildTreeType, plngTier As Long, plngAbility As Long) As Boolean
    Dim blnTaken As Boolean
    Dim i As Long
    
    For i = 1 To ptypBuildTree.Abilities
        With ptypBuildTree.Ability(i)
            If .Tier = plngTier And .Ability = plngAbility Then blnTaken = True
        End With
        If blnTaken Then
            AbilityTaken = True
            Exit Function
        End If
    Next
End Function

' The key difference between ability and feat selectors is that any given ability can only be taken once.
' eg: Improved Critical can be taken multiple times, but Tier 1 Elemental Arrows can only ever be taken once.
Public Sub GetSelectors(ptypTree As TreeType, plngTier As Long, plngAbility As Long, pblnSelector() As Boolean)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim typTaken() As FeatTakenType
    Dim lngFeat As Long
    Dim i As Long

    With ptypTree.Tier(plngTier).Ability(plngAbility)
        ReDim pblnSelector(.Selectors)
        Select Case ptypTree.Tier(plngTier).Ability(plngAbility).SelectorStyle
            Case sseRoot
                For i = 1 To .Selectors
                    pblnSelector(i) = True
                Next
            Case sseShared
                If .Parent.Style = peFeat Then
                    ' Call IdentifyTakenFeats() to find all selectors taken
                    lngFeat = .Parent.Feat
                    IdentifyTakenFeats typTaken, build.MaxLevels
                    For i = 1 To .Selectors
                        pblnSelector(i) = typTaken(lngFeat).Selector(i)
                    Next
                Else
                    ' Flag the only selector taken
                    lngSelector = GetSelector(.Parent)
                    pblnSelector(lngSelector) = True
                End If
            Case sseExclusive
                ' Initialize all choices as valid
                For i = 1 To .Selectors
                    pblnSelector(i) = True
                Next
                ' Parent choice is already taken
                lngSelector = GetSelector(.Parent)
                pblnSelector(lngSelector) = False
                ' Siblings are also taken
                For i = 1 To .Siblings
                    lngSelector = GetSelector(.Sibling(i))
                    pblnSelector(lngSelector) = False
                Next
        End Select
    End With
End Sub

Private Function GetSelector(ptypPointer As PointerType) As Long
    Dim lngBuildTree As Long
    
    With ptypPointer
        Select Case .Style
            Case peEnhancement
                lngBuildTree = FindBuildTree(db.Tree(.Tree).TreeName)
                GetSelector = GetSelectorChosen(build.Tree(lngBuildTree), ptypPointer)
            Case peDestiny
                GetSelector = GetSelectorChosen(build.Destiny, ptypPointer)
        End Select
    End With
End Function

Private Function GetSelectorChosen(ptypBuildTree As BuildTreeType, ptypPointer As PointerType) As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim i As Long
    
    With ptypPointer
        lngTier = .Tier
        lngAbility = .Ability
    End With
    With ptypBuildTree
        For i = 1 To .Abilities
            If .Ability(i).Tier = lngTier Then
                If .Ability(i).Ability = lngAbility Then
                    GetSelectorChosen = .Ability(i).Selector
                    Exit For
                End If
            End If
        Next
    End With
End Function

' Combine 4 bytes into a long
Public Function CreateAbilityID(ByVal plngTier As Long, ByVal plngAbility As Long, Optional ByVal plngRanks As Long, Optional ByVal plngSelector As Long) As Long
    Dim typAbilityID As AbilityIDType
    Dim typLong As LongType
    
    typAbilityID.Tier = plngTier
    typAbilityID.Ability = plngAbility
    typAbilityID.Ranks = plngRanks
    typAbilityID.Selector = plngSelector
    LSet typLong = typAbilityID
    CreateAbilityID = typLong.Value
End Function

' Split a long into 4 bytes
Public Sub SplitAbilityID(plngLong As Long, plngTier As Long, plngAbility As Long, Optional plngRanks As Long, Optional plngSelector As Long)
    Dim typAbilityID As AbilityIDType
    Dim typLong As LongType
    
    typLong.Value = plngLong
    LSet typAbilityID = typLong
    plngTier = typAbilityID.Tier
    plngAbility = typAbilityID.Ability
    plngRanks = typAbilityID.Ranks
    plngSelector = typAbilityID.Selector
End Sub

Public Function GetSpentInTree(ptypTree As TreeType, ptypBuildTree As BuildTreeType, plngSpent() As Long, plngTotal As Long) As Boolean
    Dim lngPoints As Long
    Dim lngTier As Long
    Dim i As Long
    Dim j As Long
    
    plngTotal = 0
    ReDim plngSpent(6)
    ' Total up points spent per tier
    With ptypBuildTree
        If .TreeType = tseDestiny Then GetSpentInTree = True
        For i = 1 To .Abilities
            With .Ability(i)
                If .Tier = 0 And .Ability = 1 Then GetSpentInTree = True
                If .Ability <> 0 Then
                    lngPoints = GetPoints(ptypTree.Tier(.Tier).Ability(.Ability), .Selector, .Rank)
                    lngTier = GetTier(.Tier, .Ability, ptypTree.TreeType)
                    For j = lngTier To 6
                        plngSpent(j) = plngSpent(j) + lngPoints
                    Next
                    plngTotal = plngTotal + lngPoints
                End If
            End With
        Next
    End With
End Function

' Send 0 to get total for all trees
Public Function QuickSpentInTree(plngBuildTree As Long) As Long
    Dim lngTree As Long
    Dim lngBuildTree As Long
    Dim lngTotal As Long
    Dim lngAbility As Long
    
    For lngBuildTree = 1 To build.Trees
        If plngBuildTree = 0 Or plngBuildTree = lngBuildTree Then
            With build.Tree(lngBuildTree)
                lngTree = SeekTree(.TreeName, peEnhancement)
                For lngAbility = 1 To .Abilities
                    With .Ability(lngAbility)
                        If .Ability <> 0 Then lngTotal = lngTotal + GetPoints(db.Tree(lngTree).Tier(.Tier).Ability(.Ability), .Selector, .Rank)
                    End With
                Next
            End With
            If plngBuildTree <> 0 Then Exit For
        End If
    Next
    QuickSpentInTree = lngTotal
End Function

' Treat class tree cores as tier = core # for Spent In Tree purposes
Public Function GetTier(ByVal plngTier As Long, ByVal plngAbility As Long, penTreeStyle As TreeStyleEnum) As Long
    If penTreeStyle = tseRace Or plngTier > 0 Then
        GetTier = plngTier
    ElseIf plngAbility = 1 Then
        GetTier = 0
    Else
        GetTier = plngAbility
    End If
End Function

' Returns TRUE if this ability fails spent in tree
Public Function CheckSpentInTreeAbility(ptypTree As TreeType, ptypAbility As BuildAbilityType, plngSpent() As Long) As Boolean
    Dim lngTier As Long
    
    lngTier = GetTier(ptypAbility.Tier, ptypAbility.Ability, ptypTree.TreeType)
    If lngTier > 0 Then
        If plngSpent(lngTier - 1) < GetSpentReq(ptypTree.TreeType, ptypAbility.Tier, ptypAbility.Ability) Then
            gstrError = "Not enough spent in tree"
            CheckSpentInTreeAbility = True
        End If
    End If
End Function

' Returns TRUE if this ability fails class/character levels
Public Function CheckLevels(ptypBuildTree As BuildTreeType, ByVal plngTier As Long, ByVal plngAbility As Long) As Boolean
    Dim lngLevels As Long
    Dim lngClassLevels As Long
    
    GetLevelReqs ptypBuildTree.TreeType, plngTier, plngAbility, lngLevels, lngClassLevels
    If build.MaxLevels < lngLevels Then
        gstrError = "Not enough character levels"
        CheckLevels = True
    ElseIf ptypBuildTree.ClassLevels < lngClassLevels Then
        gstrError = "Not enough class levels"
        CheckLevels = True
    End If
End Function

Public Sub GetLevelReqs(ByVal penTreeStyle As TreeStyleEnum, ByVal plngTier As Long, ByVal plngAbility As Long, plngLevels As Long, plngClassLevels As Long)
    plngLevels = 0
    plngClassLevels = 0
    If plngTier = 0 Then
        GetLevelReqsCore penTreeStyle, plngAbility, plngLevels, plngClassLevels
    Else
        GetLevelReqsTier penTreeStyle, plngTier, plngLevels, plngClassLevels
    End If
End Sub

Public Function GetBuildLevelReq(plngLevels As Long, plngClassLevels As Long, penClass As ClassEnum) As Long
    Dim lngClassLevels As Long
    Dim lngBuildLevels As Long
    
    ' What build level when we reach class level?
    If plngClassLevels Then
        For lngBuildLevels = 1 To HeroicLevels()
            If build.Class(lngBuildLevels) = penClass Then
                lngClassLevels = lngClassLevels + 1
                If lngClassLevels = plngClassLevels Then Exit For
            End If
        Next
        If lngBuildLevels > HeroicLevels() Then
            GetBuildLevelReq = -1
            Exit Function
        End If
    End If
    ' Is character level requirenent higher than class level requirement?
    If plngLevels > lngBuildLevels Then lngBuildLevels = plngLevels
    ' Do we have enough levels?
    If lngBuildLevels > build.MaxLevels Then lngBuildLevels = -1
    GetBuildLevelReq = lngBuildLevels
End Function

Private Function GetLevelReqsCore(penTreeStyle As TreeStyleEnum, plngAbility As Long, plngLevels As Long, plngClassLevels As Long) As Long
    Dim lngMin As Long
    
    Select Case penTreeStyle
        Case tseClass, tseGlobal
            Select Case plngAbility
                Case 1: lngMin = 1
                Case 2: lngMin = 3
                Case 3: lngMin = 6
                Case 4: lngMin = 12
                Case 5: lngMin = 18
                Case 6: lngMin = 20
            End Select
        Case tseRace
            Select Case plngAbility
                Case 1: lngMin = 1
                Case 2: lngMin = 4
                Case 3: lngMin = 7
                Case 4: lngMin = 11
                Case 5: lngMin = 16
            End Select
        Case tseRaceClass
            Select Case plngAbility
                Case 1: lngMin = 1
                Case 2: lngMin = 4
                Case 3: lngMin = 8
                Case 4: lngMin = 15
                Case 5: lngMin = 22
                Case 6: lngMin = 25
            End Select
    End Select
    If penTreeStyle = tseClass Then plngClassLevels = lngMin Else plngLevels = lngMin
End Function

Private Function GetLevelReqsTier(penTreeStyle As TreeStyleEnum, plngTier As Long, plngLevels As Long, plngClassLevels As Long) As Long
    Select Case penTreeStyle
        Case tseClass
            plngClassLevels = plngTier
            If plngTier = 5 Then plngLevels = 12
        Case tseGlobal, tseRaceClass
            If plngTier = 5 Then plngLevels = 12 Else plngLevels = plngTier
    End Select
End Function

Public Function GetSpentReq(ByVal penTreeStyle As TreeStyleEnum, ByVal plngTier As Long, ByVal plngAbility As Long) As Long
    Select Case penTreeStyle
        Case tseRace
            Select Case plngTier
                Case 1: GetSpentReq = 1
                Case 2: GetSpentReq = 5
                Case 3: GetSpentReq = 10
                Case 4: GetSpentReq = 15
            End Select
        Case tseClass, tseGlobal, tseRaceClass
            Select Case plngTier
                Case 1: GetSpentReq = 1
                Case 2: GetSpentReq = 5
                Case 3: GetSpentReq = 10
                Case 4: GetSpentReq = 20
                Case 5: GetSpentReq = 30
                Case 0
                    Select Case plngAbility
                        Case 2: GetSpentReq = 5
                        Case 3: GetSpentReq = 10
                        Case 4: GetSpentReq = 20
                        Case 5: GetSpentReq = 30
                        Case 6: GetSpentReq = 40
                    End Select
            End Select
        Case tseDestiny
            If plngTier > 1 Then GetSpentReq = (plngTier - 1) * 4
    End Select
End Function

Private Function GetPoints(ptypAbility As AbilityType, ByVal plngSelector As Long, ByVal plngRanks As Long) As Long
    Dim lngPoints As Long
    
    If plngSelector = 0 Or ptypAbility.Selectors = 0 Then lngPoints = ptypAbility.Cost Else lngPoints = ptypAbility.Selector(plngSelector).Cost
    If plngRanks > 1 Then lngPoints = lngPoints * plngRanks
    GetPoints = lngPoints
End Function

' Returns TRUE if errors found
Public Function CheckAbilityErrors(ptypTree As TreeType, ptypBuildTree As BuildTreeType, ptypAbility As BuildAbilityType, plngSpent() As Long) As Boolean
    Dim blnPassChecks As Boolean
    
    CheckAbilityErrors = True
    ' Levels (eg: 3 class levels for tier 3, 12 build levels for tier 5, 4 character levels for racial core 2, etc...)
    If CheckLevels(ptypBuildTree, ptypAbility.Tier, ptypAbility.Ability) Then Exit Function
    ' Points in tree
    If CheckSpentInTreeAbility(ptypTree, ptypAbility, plngSpent) Then Exit Function
    With ptypAbility
        Do
            ' Class checks (eg: Half-Orc Tier 4: Power Rage requires 2 barbarian levels for rank 1, 6 for rank 2)
            With ptypTree.Tier(.Tier).Ability(.Ability)
                Do
'                    If CheckAbilityClassLevels(.Class, .ClassLevel, 1) Then Exit Do
                    If .RankReqs And ptypAbility.Rank > 1 Then
                        If CheckAbilityClassLevels(.Rank(ptypAbility.Rank).Class, .Rank(ptypAbility.Rank).ClassLevel, ptypAbility.Rank) Then Exit Do
                    End If
                    blnPassChecks = True
                Loop Until True
            End With
            If Not blnPassChecks Then Exit Do
            blnPassChecks = False
            ' All/One/None
            If .Selector = 0 Then
                If Not CheckAbilityReqs(ptypTree.Tier(.Tier).Ability(.Ability).Req, .Rank, False) Then Exit Do
                ' Ranks
                If .Rank > 1 And ptypTree.Tier(.Tier).Ability(.Ability).RankReqs Then
                    If Not CheckAbilityReqs(ptypTree.Tier(.Tier).Ability(.Ability).Rank(.Rank).Req, .Rank, True) Then Exit Do
                    If .Selector <> 0 Then
                        If Not CheckAbilityReqs(ptypTree.Tier(.Tier).Ability(.Ability).Selector(.Selector).Rank(.Rank).Req, .Rank, True) Then Exit Do
                    End If
                End If
            Else
                If Not CheckAbilityReqs(ptypTree.Tier(.Tier).Ability(.Ability).Selector(.Selector).Req, .Rank, False) Then Exit Do
                ' Ranks
                If .Rank > 1 And ptypTree.Tier(.Tier).Ability(.Ability).Selector(.Selector).RankReqs Then
                    'Note this is checking at the selector level - we don't have data if the req is
                    'only on one selector (SS->t1:Studies->Magic/Music)
                    'We'd need to add something like Rank3Sel1None: to the ench.txt file
                    'and update the dataload stuff
                    If Not CheckAbilityReqs(ptypTree.Tier(.Tier).Ability(.Ability).Selector(.Selector).Rank(.Rank).Req, .Rank, True) Then Exit Do
                End If
            End If

            blnPassChecks = True
        Loop Until True
    End With
    CheckAbilityErrors = Not blnPassChecks
End Function

' Returns TRUE if we fail this check
Public Function CheckAbilityClassLevels(pblnClass() As Boolean, plngLevel() As Long, ByVal plngRank As Long) As Boolean
    Dim lngClassLevels() As Long
    Dim i As Long
    
    If Not pblnClass(0) Then Exit Function
    ReDim lngClassLevels(ceClasses - 1)
    For i = 1 To HeroicLevels()
        lngClassLevels(build.Class(i)) = lngClassLevels(build.Class(i)) + 1
    Next
    For i = 1 To ceClasses - 1
        If pblnClass(i) Then
            If lngClassLevels(i) >= plngLevel(i) Then Exit Function
        End If
    Next
    If plngRank > 1 Then gstrError = "Rank " & plngRank & " class requirements" Else gstrError = "Class requirements"
    CheckAbilityClassLevels = True
End Function

' Returns TRUE if we successfully pass all abilityreqs
Private Function CheckAbilityReqs(ptypReqList() As ReqListType, ByVal plngRanks As Long, pblnRankReq As Boolean) As Boolean
    Dim enReq As ReqGroupEnum
    
    For enReq = rgeAll To rgeNone
        If CheckAbilityReq(ptypReqList(enReq), enReq, plngRanks, pblnRankReq) Then Exit Function
    Next
    CheckAbilityReqs = True
End Function

' Returns TRUE if we fail any reqs
Private Function CheckAbilityReq(ptypReqList As ReqListType, penReq As ReqGroupEnum, plngRanks As Long, pblnRankReq As Boolean) As Boolean
    Dim lngMatches As Long
    Dim lngTree As Long
    Dim strTaken As String
    Dim strMissing As String
    Dim i As Long

    If ptypReqList.Reqs = 0 Then Exit Function
    For i = 1 To ptypReqList.Reqs
        With ptypReqList.Req(i)
            Select Case .Style
                Case peFeat
                    If CheckAbilityFeat(.Feat, .Selector) Then
                        If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, 1)
                        lngMatches = lngMatches + 1
                    Else
                        If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, 1)
                    End If
                Case peDestiny
                    If CheckAbility(build.Destiny, ptypReqList.Req(i), penReq, plngRanks) Then
                        If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                        lngMatches = lngMatches + 1
                    Else
                        If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                    End If
                Case peEnhancement
                    lngTree = FindBuildTree(db.Tree(.Tree).TreeName)
                    If lngTree <> 0 Then
                        If CheckAbility(build.Tree(lngTree), ptypReqList.Req(i), penReq, plngRanks) Then
                            If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                            lngMatches = lngMatches + 1
                        Else
                            If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                        End If
                    End If
            End Select
        End With
    Next
    Select Case penReq
        Case rgeAll
            If lngMatches < ptypReqList.Reqs Then
                If pblnRankReq Then gstrError = "Rank " & plngRanks & " requires " Else gstrError = "Requires "
                gstrError = gstrError & strMissing
                CheckAbilityReq = True
            End If
        Case rgeOne
            If lngMatches < 1 Then
                gstrError = "Nothing taken from the 'One of' list"
                CheckAbilityReq = True
            End If
        Case rgeNone
            If lngMatches > 0 Then
                If pblnRankReq Then gstrError = "Rank " & plngRanks & " excludes " Else gstrError = "Antireq for "
                gstrError = gstrError & strTaken
                CheckAbilityReq = True
            End If
    End Select
End Function

' Returns TRUE if errors found
Public Function CheckTwistSlot(plngTwist As Long) As Boolean
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    
    If plngTwist > build.Twists Then Exit Function
    With build.Twist(plngTwist)
        If Len(.DestinyName) Then lngDestiny = SeekTree(.DestinyName, peDestiny)
        lngTier = .Tier
        lngAbility = .Ability
    End With
    If lngDestiny = 0 Or lngTier = 0 Or lngAbility = 0 Then Exit Function
    If lngTier > 4 Then
        gstrError = "Can't twist tier " & lngTier & " abilities"
        CheckTwistSlot = True
    Else
        CheckTwistSlot = True
        If CheckTwistReq(db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility).Req(rgeAll), rgeAll) Then Exit Function
        If CheckTwistReq(db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility).Req(rgeNone), rgeNone) Then Exit Function
        CheckTwistSlot = False
    End If
End Function

' Returns TRUE if feat prereqs not satisfied
Public Function CheckTwistErrors(ptypAbility As AbilityType) As Boolean
    CheckTwistErrors = True
    If CheckTwistReq(ptypAbility.Req(rgeAll), rgeAll) Then Exit Function
    If CheckTwistReq(ptypAbility.Req(rgeNone), rgeNone) Then Exit Function
    CheckTwistErrors = False
End Function

' Returns TRUE if we fail any reqs
Private Function CheckTwistReq(ptypReqList As ReqListType, penReq As ReqGroupEnum) As Boolean
    Dim lngMatches As Long
    Dim lngTree As Long
    Dim strTaken As String
    Dim strMissing As String
    Dim lngTotal As Long
    Dim i As Long

    If ptypReqList.Reqs = 0 Then Exit Function
    lngTotal = ptypReqList.Reqs
    For i = 1 To ptypReqList.Reqs
        With ptypReqList.Req(i)
            If .Style = peFeat Then
                If CheckAbilityFeat(.Feat, .Selector) Then
                    If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, 1)
                    lngMatches = lngMatches + 1
                Else
                    If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                End If
            Else
                lngTotal = lngTotal - 1
            End If
        End With
    Next
    Select Case penReq
        Case rgeAll
            If lngMatches < lngTotal Then
                gstrError = "Requires " & strMissing
                CheckTwistReq = True
            End If
        Case rgeNone
            If lngMatches > 0 Then
                gstrError = "Antireq for " & strMissing
                CheckTwistReq = True
            End If
    End Select
End Function

Public Function CheckAbilityFeat(ByVal plngFeat As Long, ByVal plngSelector As Long) As Boolean
    Dim typTaken() As FeatTakenType
    Dim i As Long
    
    IdentifyTakenFeats typTaken, build.MaxLevels
    If plngSelector = 0 Then
        CheckAbilityFeat = (typTaken(plngFeat).Times <> 0)
    Else
        CheckAbilityFeat = typTaken(plngFeat).Selector(plngSelector)
    End If
End Function

' Returns TRUE if ability found
Public Function CheckAbility(ptypBuildTree As BuildTreeType, ptypReq As PointerType, penReq As ReqGroupEnum, plngRanks As Long) As Boolean
    Dim lngRanks As Long
    Dim blnFound As Boolean
    Dim lngMaxRanks As Long
    Dim i As Long
    
    If ptypReq.Rank = 0 Then lngRanks = plngRanks Else lngRanks = ptypReq.Rank
    With ptypBuildTree
        For i = 1 To .Abilities
            If .Ability(i).Tier = ptypReq.Tier And .Ability(i).Ability = ptypReq.Ability Then
                lngMaxRanks = GetAbilityMaxRanks(.TreeType, .TreeName, .Ability(i).Tier, .Ability(i).Ability)
                If ptypReq.Selector = 0 Then
                    blnFound = True
                    Exit For
                ElseIf .Ability(i).Selector = ptypReq.Selector Then
                    blnFound = True
                    Exit For
                End If
            End If
        Next
        If blnFound Then
            If lngRanks <> 0 Then
                lngMaxRanks = GetAbilityMaxRanks(.TreeType, .TreeName, .Ability(i).Tier, .Ability(i).Ability)
                ' lngRanks = the number of ranks we took in an ability we're checking prereqs for
                ' .Ability(i).Rank = for the prereq we found as taken, how many ranks were taken
                ' lngMaxRanks = for the prereq we found as taken, max ranks you can take
                If .Ability(i).Rank >= lngRanks Or lngMaxRanks < lngRanks Then CheckAbility = True
            Else
                CheckAbility = True
            End If
        End If
    End With
End Function

Private Function GetAbilityMaxRanks(ByVal penTreeType As TreeStyleEnum, pstrTreeName As String, ByVal plngTier As Long, ByVal plngAbility As Long) As Long
On Error GoTo GetAbilityMaxRanksErr
    Dim lngMaxRanks As Long
    Dim lngTree As Long
    
    lngTree = SeekTree(pstrTreeName, penTreeType)
    If lngTree Then
        If penTreeType = tseDestiny Then
            lngMaxRanks = db.Destiny(lngTree).Tier(plngTier).Ability(plngAbility).Ranks
        Else
            lngMaxRanks = db.Tree(lngTree).Tier(plngTier).Ability(plngAbility).Ranks
        End If
    End If
    
GetAbilityMaxRanksExit:
    If lngMaxRanks = 0 Then GetAbilityMaxRanks = 3 Else GetAbilityMaxRanks = lngMaxRanks
    Exit Function
    
GetAbilityMaxRanksErr:
    MsgBox "Unexpected Error: " & Err.Description & vbNewLine & vbNewLine & _
        "Module: basBuild.GetAbilityMaxRanks()" & vbNewLine & _
        "Ability: " & pstrTreeName & " Tier " & plngTier & " Ability " & plngAbility & vbNewLine & vbNewLine & _
        "Error can be ignored. Program resuming.", vbInformation, "Error #" & Err.Number
    Resume GetAbilityMaxRanksExit
End Function


' ************* LEVELING GUIDE *************


Public Sub InitLevelingGuide()
    InitGuideTrees
    InitGuideEnhancements
End Sub

Private Sub InitGuideTrees()
    Dim lngClassLevels() As Long
    Dim i As Long
    Dim j As Long
    
    Erase Guide.Tree, Guide.Enhancement, Guide.TreeLookup
    Guide.Enhancements = 0
    Guide.Trees = 0
    If build.Class(1) = ceAny Or build.Race = reAny Then Exit Sub
    ' Add racial tree
    AddGuideTree db.Race(build.Race).RaceName, ceAny, build.MaxLevels
    ' Add class trees
    ReDim lngClassLevels(ceClasses - 1)
    For i = 1 To HeroicLevels()
        lngClassLevels(build.Class(i)) = lngClassLevels(build.Class(i)) + 1
    Next
    For i = 1 To ceClasses - 1
        If lngClassLevels(i) > 0 Then
            With db.Class(i)
                For j = 1 To .Trees
                    AddGuideTree .Tree(j), i, lngClassLevels(i)
                Next
            End With
        End If
    Next
    ' Add racial class tree(s)
    With db.Race(build.Race)
        For i = 1 To .Trees
            AddGuideTree .Tree(i), ceAny, build.MaxLevels
        Next
    End With
    ' Add global tree(s)
    For i = 1 To db.Trees
        If db.Tree(i).TreeType = tseGlobal Then AddGuideTree db.Tree(i).TreeName, ceAny, build.MaxLevels
    Next
    ' Create lookup
    If build.Guide.Trees = 0 Then
        Erase Guide.TreeLookup
    Else
        ReDim Guide.TreeLookup(build.Guide.Trees)
        For i = 1 To build.Guide.Trees
            For j = 1 To Guide.Trees
                If build.Guide.Tree(i).TreeName = Guide.Tree(j).TreeName And build.Guide.Tree(i).Class = Guide.Tree(j).Class Then
                    Guide.TreeLookup(i) = j
                    Exit For
                End If
            Next
        Next
    End If
End Sub

Private Sub AddGuideTree(pstrTreeName As String, penClass As ClassEnum, ByVal plngClassLevels As Long)
    Dim blnDuplicate As Boolean
    Dim i As Long
    
    ' Already added?
    For i = 1 To Guide.Trees
        If Guide.Tree(i).TreeName = pstrTreeName Then
            Guide.Tree(i).Duplicate = True
            Guide.Tree(i).Display = db.Tree(Guide.Tree(i).TreeID).Abbreviation & " (" & db.Class(Guide.Tree(i).Class).Initial(3) & ")"
            blnDuplicate = True
            Exit For
        End If
    Next
    With Guide
        .Trees = .Trees + 1
        ReDim Preserve .Tree(1 To .Trees)
        With .Tree(.Trees)
            .GuideTreeID = i
            .TreeName = pstrTreeName
            .TreeID = SeekTree(.TreeName, peEnhancement)
            .Abbreviation = db.Tree(.TreeID).Abbreviation
            .TreeStyle = db.Tree(.TreeID).TreeType
            .Class = penClass
            .ClassLevels = plngClassLevels
            .MaxTier = .ClassLevels
            If .MaxTier > 5 Then .MaxTier = 5
            If .MaxTier = 5 And (.TreeStyle = tseRace Or build.MaxLevels < 12) Then .MaxTier = 4
            .Duplicate = blnDuplicate
            .Display = db.Tree(.TreeID).Abbreviation
            If .Duplicate Then .Display = .Display & " (" & db.Class(.Class).Initial(3) & ")"
            For i = 1 To build.Guide.Trees
                If build.Guide.Tree(i).TreeName = .TreeName And build.Guide.Tree(i).Class = .Class Then .BuildGuideTreeID = i
            Next
            For i = 1 To build.Trees
                If build.Tree(i).TreeName = .TreeName And build.Tree(i).Source = .Class Then .BuildTreeID = i
            Next
            .BuildTree.Source = .Class
            .BuildTree.ClassLevels = .ClassLevels
            .BuildTree.TreeName = .TreeName
            .BuildTree.TreeType = .TreeStyle
        End With
    End With
End Sub

Public Sub InitGuideEnhancements()
    Dim lngSpentInTree() As Long
    Dim blnSelectorOnly As Boolean
    Dim lngSpent As Long
    Dim lngSpentRacial As Long
    Dim blnRacialTree As Boolean
    Dim i As Long
    
    On Error GoTo 0
    Guide.Enhancements = build.Guide.Enhancements
    If build.Guide.Enhancements = 0 Then
        Erase Guide.Enhancement
        Exit Sub
    End If
    With Guide
        ReDim lngSpentInTree(.Trees)
        ReDim .Enhancement(.Enhancements)
        For i = 1 To .Enhancements
            With build.Guide.Enhancement(i)
                Guide.Enhancement(i).Tier = .Tier
                Guide.Enhancement(i).Ability = .Ability
                Guide.Enhancement(i).Selector = .Selector
                Guide.Enhancement(i).Rank = .Rank
                Guide.Enhancement(i).BuildGuideTreeID = build.Guide.Enhancement(i).ID
            End With
            With .Enhancement(i)
                blnRacialTree = False
                Select Case .BuildGuideTreeID
                    Case 1 To 99
                        .Style = geEnhancement
                        .GuideTreeID = Guide.TreeLookup(.BuildGuideTreeID)
                        blnRacialTree = (Guide.Tree(.GuideTreeID).TreeStyle = tseRace)
                    Case 100
                        .Style = geResetAllTrees
                        .BuildGuideTreeID = 0
                        ReDim lngSpentInTree(Guide.Trees)
                        lngSpent = 0
                        lngSpentRacial = 0
                        .Spent = 0
                        .SpentRacial = 0
                        .SpentInTree = 0
                    Case 101 To 199
                        .Style = geResetTree
                        .BuildGuideTreeID = .BuildGuideTreeID - 100
                        .GuideTreeID = Guide.TreeLookup(.BuildGuideTreeID)
                        lngSpent = lngSpent - lngSpentInTree(.GuideTreeID)
                        lngSpentInTree(.GuideTreeID) = 0
                        .Spent = lngSpent
                        .SpentInTree = 0
                        If Guide.Tree(.GuideTreeID).TreeStyle = tseRace Then lngSpentRacial = 0
                        .SpentRacial = lngSpentRacial
                    Case 200
                        .Style = geBankAP
                        .BuildGuideTreeID = 0
                        .Spent = lngSpent
                        .SpentInTree = lngSpent
                        .SpentRacial = lngSpentRacial
                    Case Else
                        .Style = geUnknown
                        .BuildGuideTreeID = 0
                        .Spent = 0
                        .SpentInTree = 0
                        .SpentRacial = 0
                        .Cost = 0
                        .ML = 0
                End Select
                If .BuildGuideTreeID <> 0 Then
                    .BuildTreeID = Guide.Tree(.GuideTreeID).BuildTreeID
                    .TreeID = Guide.Tree(.GuideTreeID).TreeID
                End If
                If .Style = geEnhancement Then
                    .ML = GuideAbilityML(.GuideTreeID, .Tier, .Ability, .Selector, .Rank)
                    If db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).SelectorStyle = sseNone Then
                        .Cost = db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Cost
                        .Display = db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Abbreviation
                    Else
                        .Cost = db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Selector(.Selector).Cost
                        If db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).SelectorOnly Then
                            .Display = db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Selector(.Selector).SelectorName
                        Else
                            Guide.Enhancement(i).Display = db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Abbreviation & ": " & db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Selector(.Selector).SelectorName
                        End If
                    End If
                    If db.Tree(.TreeID).Tier(.Tier).Ability(.Ability).Ranks > 1 Then .RankText = Left$(" III", .Rank + 1) Else .RankText = vbNullString
                    lngSpent = lngSpent + .Cost
                    lngSpentInTree(.GuideTreeID) = lngSpentInTree(.GuideTreeID) + .Cost
                    .SpentInTree = lngSpentInTree(.GuideTreeID)
                    .Spent = lngSpent
                    If blnRacialTree Then lngSpentRacial = lngSpentRacial + .Cost
                    .SpentRacial = lngSpentRacial
                End If
            End With
        Next
    End With
    For i = 1 To build.Guide.Enhancements
        AddGuideBuildAbility build.Guide.Enhancement(i)
    Next
    CalculateGuideLevels
    CheckGuideErrors
End Sub

' Actual minimum level for this build based on all factors
Public Function GuideAbilityML(plngGuideTree As Long, plngTier As Long, plngAbility As Long, ByVal plngSelector As Long, plngRank As Long) As Long
    Dim lngTree As Long
    Dim enClass As ClassEnum
    Dim lngLevels As Long
    Dim lngClassLevels As Long
    Dim lngReturn As Long
    Dim lngRecurse As Long
    Dim i As Long
    
    enClass = Guide.Tree(plngGuideTree).Class
    lngTree = Guide.Tree(plngGuideTree).TreeID
    ' Start with basic tree ML mechanics
    GetLevelReqs db.Tree(lngTree).TreeType, plngTier, plngAbility, lngLevels, lngClassLevels
    lngReturn = GetBuildLevelReq(lngLevels, lngClassLevels, enClass)
    Select Case lngReturn
        Case -1: lngReturn = 0
        Case 0: lngReturn = 1
    End Select
    ' Feat reqs
    With db.Tree(lngTree).Tier(plngTier).Ability(plngAbility)
        ' Ability feat & class reqs
        GuideFeatReqLevel .Req, lngReturn
'        GuideClassReqLevel .Class, .ClassLevel, lngReturn
        ' Rank feat & class reqs
        If .RankReqs = True And plngRank > 1 Then
            With .Rank(plngRank)
                GuideFeatReqLevel .Req, lngReturn
                GuideClassReqLevel .Class, .ClassLevel, lngReturn
            End With
        End If
        ' Selector feat & class reqs
        If plngSelector <> 0 Then
            With .Selector(plngSelector)
                GuideFeatReqLevel .Req, lngReturn
                ' Selector Rank feat & class reqs
                If .RankReqs = True And plngRank > 1 Then
                    With .Rank(plngRank)
                        GuideFeatReqLevel .Req, lngReturn
                        GuideClassReqLevel .Class, .ClassLevel, lngReturn
                    End With
                End If
            End With
        End If
        ' Recurse check "All" enhancements for their own feat reqs (for simplicity, only check if prereq is in same tree)
        If Not lngReturn < 1 Then
            With .Req(rgeAll)
                For i = 1 To .Reqs
                    With .Req(i)
                        If .Style = peEnhancement And .Tree = lngTree Then
                            lngRecurse = GuideAbilityML(plngGuideTree, .Tier, .Ability, .Selector, .Rank)
                            If lngRecurse = 0 Or lngRecurse > lngReturn Then lngReturn = lngRecurse
                        End If
                    End With
                Next
            End With
        End If
    End With
    GuideAbilityML = lngReturn
End Function

Private Sub GuideFeatReqLevel(ptypReq() As ReqListType, plngLevel As Long)
    Dim lngReq As Long
    Dim lngFeat As Long
    Dim blnNeed As Boolean
    Dim blnFound As Boolean

    If plngLevel < 1 Then Exit Sub
    For lngReq = 1 To ptypReq(rgeAll).Reqs
        If ptypReq(rgeAll).Req(lngReq).Style = peFeat Then
            blnNeed = True
            For lngFeat = 1 To Feat.Count
                If Feat.List(lngFeat).ActualType <> bftAlternate Then
                    If Feat.List(lngFeat).FeatID = ptypReq(rgeAll).Req(lngReq).Feat Then
                        blnFound = True
                        If plngLevel < Feat.List(lngFeat).Level Then plngLevel = Feat.List(lngFeat).Level
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    If blnNeed And Not blnFound Then plngLevel = 0
End Sub

Private Sub GuideClassReqLevel(pblnClass() As Boolean, plngClassLevel() As Long, plngLevel As Long)
    Dim lngReturn As Long
    Dim enClass As Long
    Dim lngLevel As Long

    If plngLevel < 1 Then Exit Sub
    If Not pblnClass(0) Then Exit Sub
    For enClass = 1 To ceClasses - 1
        If pblnClass(enClass) Then
            lngLevel = GetBuildLevelReq(plngLevel, plngClassLevel(enClass), enClass)
            If lngReturn < 1 Or (lngLevel > 0 And lngReturn > lngLevel) Then lngReturn = lngLevel
        End If
    Next
    If lngReturn < 1 Then
        plngLevel = 0
    ElseIf plngLevel < lngReturn Then
        plngLevel = lngReturn
    End If
End Sub

Private Sub CalculateGuideLevels()
    Dim lngLevel As Long
    Dim lngSpent As Long
    Dim lngSpentRacial As Long ' Amount of racial AP actually spent, capped at racial past life AP (ie: # of AP to be "ignored")
    Dim lngSpentInTree() As Long
    Dim i As Long
    
    ReDim lngSpentInTree(build.Guide.Trees)
    If db.Race(build.Race).Type = rteIconic Then
        If build.MaxLevels > 14 Then lngLevel = 14 Else lngLevel = build.MaxLevels
    Else
        lngLevel = 1
    End If
    For i = 1 To Guide.Enhancements
        With Guide.Enhancement(i)
            .Level = 0
            .Bank = 0
            .SpentInTree = 0
            Select Case .Style
                Case geUnknown
                Case geEnhancement
                    lngSpent = lngSpent + .Cost
                    lngSpentInTree(.BuildGuideTreeID) = lngSpentInTree(.BuildGuideTreeID) + .Cost
                    .SpentInTree = lngSpentInTree(.BuildGuideTreeID)
                    If .SpentRacial > build.RacialAP Then lngSpentRacial = build.RacialAP Else lngSpentRacial = .SpentRacial
                    If lngLevel < .ML Then lngLevel = .ML
                    Do While lngSpent > GetAP(lngLevel) + lngSpentRacial
                        lngLevel = lngLevel + 1
                        If lngLevel > build.MaxLevels Then
                            lngLevel = build.MaxLevels
                            Exit Do
                        End If
                    Loop
                    .Level = lngLevel
                Case geBankAP
                    Do
                        .Bank = GetAP(lngLevel) - lngSpent + lngSpentRacial
                        If .Bank < 1 Then
                            lngLevel = lngLevel + 1
                            If lngLevel > build.MaxLevels Then
                                lngLevel = build.MaxLevels
                                Exit Do
                            End If
                        End If
                    Loop Until .Bank > 0
                    .Level = lngLevel
                    lngLevel = lngLevel + 1
                    .Display = "Bank " & .Bank & " AP"
                Case geResetTree
                    If i = 1 Then .Level = 1 Else .Level = Guide.Enhancement(i - 1).Level + 1
                    lngLevel = .Level
                    lngSpent = lngSpent - lngSpentInTree(.BuildGuideTreeID)
                    lngSpentInTree(.BuildGuideTreeID) = 0
                    .SpentInTree = 0
                    If Guide.Tree(.GuideTreeID).TreeStyle = tseRace Then lngSpentRacial = 0
                    .Display = "Reset " & Guide.Tree(.GuideTreeID).TreeName
                Case geResetAllTrees
                    If i = 1 Then .Level = 1 Else .Level = Guide.Enhancement(i - 1).Level + 1
                    lngLevel = .Level
                    lngSpent = 0
                    lngSpentRacial = 0
                    ReDim lngSpentInTree(build.Guide.Trees)
                    .SpentInTree = 0
                    .Display = "Reset All Trees"
            End Select
        End With
    Next
End Sub

Private Function GetAP(plngLevel As Long) As Long
    If plngLevel > 20 Then GetAP = 80 Else GetAP = plngLevel * 4
End Function

Public Sub AddGuideBuildAbility(ptypEnhancement As BuildGuideEnhancementType)
    Dim lngGuideTree As Long
    Dim lngInsert As Long
    Dim blnFound As Boolean
    Dim i As Long
    
    Select Case ptypEnhancement.ID
        Case 1 To 99
            lngGuideTree = Guide.TreeLookup(ptypEnhancement.ID)
            With Guide.Tree(lngGuideTree).BuildTree
                ' Is a lower rank already there?
                For i = 1 To .Abilities
                    If .Ability(i).Tier = ptypEnhancement.Tier And .Ability(i).Ability = ptypEnhancement.Ability Then
                        If .Ability(i).Rank < ptypEnhancement.Rank Then .Ability(i).Rank = ptypEnhancement.Rank
                        blnFound = True
                        Exit For
                    End If
                Next
                If Not blnFound Then
                    lngInsert = GetInsertionPoint(Guide.Tree(lngGuideTree).BuildTree, ptypEnhancement.Tier, ptypEnhancement.Ability)
                    .Abilities = .Abilities + 1
                    ReDim Preserve .Ability(1 To .Abilities)
                    For i = .Abilities To lngInsert + 1 Step -1
                        .Ability(i) = .Ability(i - 1)
                    Next
                    With .Ability(lngInsert)
                        .Tier = ptypEnhancement.Tier
                        .Ability = ptypEnhancement.Ability
                        .Selector = ptypEnhancement.Selector
                        .Rank = ptypEnhancement.Rank
                    End With
                End If
            End With
        Case 100 ' Reset all trees
            For i = 1 To Guide.Trees
                Guide.Tree(i).BuildTree.Abilities = 0
                Erase Guide.Tree(i).BuildTree.Ability
            Next
        Case 101 To 199
            lngGuideTree = Guide.TreeLookup(ptypEnhancement.ID - 100)
            Guide.Tree(lngGuideTree).BuildTree.Abilities = 0
            Erase Guide.Tree(lngGuideTree).BuildTree.Ability
    End Select
End Sub

Private Sub CheckGuideErrors()
    Dim enTreeMap() As ClassEnum
    Dim lngTree As Long
    Dim lngClassLevel() As Long
    Dim lngCurrentLevel As Long
    Dim enClass As ClassEnum
    Dim lngLevels As Long
    Dim lngClassLevels As Long
    Dim blnRankReqs As Boolean
    Dim lngTier5 As Long
    Dim lngTreeCount As Long
    Dim blnTreeMap() As Boolean
    Dim i As Long

    ReDim enTreeMap(db.Trees) ' Used to prevent taking same tree from multiple classes at once (eg: Vanguard)
    ReDim blnTreeMap(db.Trees) ' Used to prevent taking more than 6 class trees at once
    ReDim lngSpentInTree(db.Trees)
    ReDim lngClassLevel(ceClasses - 1)
    For i = 1 To Guide.Enhancements
        With Guide.Enhancement(i)
            .ErrorState = False
            .ErrorText = vbNullString
            Do
                Select Case .Style
                    Case geEnhancement
                        ' Update class level counters
                        If lngCurrentLevel < .Level Then
                            For lngCurrentLevel = lngCurrentLevel + 1 To .Level
                                If lngCurrentLevel <= HeroicLevels() Then lngClassLevel(build.Class(lngCurrentLevel)) = lngClassLevel(build.Class(lngCurrentLevel)) + 1
                            Next
                            lngCurrentLevel = .Level ' Remember that For...Next control variables = Max+1 after loop
                        End If
                        ' Make sure same class tree isn't taken by two different classes
                        If db.Tree(.TreeID).TreeType = tseClass And enTreeMap(.TreeID) <> ceAny And enTreeMap(.TreeID) <> Guide.Tree(.GuideTreeID).Class Then
                            .ErrorText = "Can't take " & db.Tree(.TreeID).TreeName & " from both " & db.Class(enTreeMap(.TreeID)).ClassName & " and " & db.Class(Guide.Tree(.GuideTreeID).Class).ClassName & " at the same time"
                            Exit Do
                        End If
                        ' Enforce 6 tree limit on class trees
                        If db.Tree(.TreeID).TreeType <> tseRace Then
                            If blnTreeMap(.TreeID) = False Then
                                blnTreeMap(.TreeID) = True
                                lngTreeCount = lngTreeCount + 1
                                If lngTreeCount > 6 Then
                                    .ErrorText = "Can't use more than 6 class trees at any one time."
                                    Exit Do
                                End If
                            End If
                        End If
                        ' Commit this tree
                        If Guide.Tree(.GuideTreeID).TreeStyle = tseClass Then
                            enTreeMap(.TreeID) = Guide.Tree(.GuideTreeID).Class
                        Else
                            enTreeMap(.TreeID) = build.Race
                        End If
                        ' Check for tree lockouts (Savants, Eldritch Knight, AA/Racial AA)
                        If Len(db.Tree(.TreeID).Lockout) Then
                            lngTree = SeekTree(db.Tree(.TreeID).Lockout, peEnhancement)
                            If lngTree Then
                                If enTreeMap(lngTree) <> ceAny Then
                                    .ErrorText = db.Tree(.TreeID).TreeName & " is locked out by " & db.Tree(lngTree).TreeName
                                    Exit Do
                                End If
                            End If
                        End If
                        ' Check Spent in Tree
                        If .SpentInTree - .Cost < GetSpentReq(db.Tree(.TreeID).TreeType, .Tier, .Ability) Then
                            .ErrorText = "Requires " & GetSpentReq(db.Tree(.TreeID).TreeType, .Tier, .Ability) & " AP spent in tree, only " & .SpentInTree - .Cost & " AP spent"
                            Exit Do
                        End If
                        ' Check build and class level reqs from basic tree mechanics
                        GetLevelReqs db.Tree(.TreeID).TreeType, .Tier, .Ability, lngLevels, lngClassLevels
                        If lngCurrentLevel < lngLevels Then
                            .ErrorText = "Requires " & lngLevels & " character levels"
                            Exit Do
                        End If
                        enClass = Guide.Tree(.GuideTreeID).Class
                        If db.Tree(.TreeID).TreeType = tseClass Then
                            If lngClassLevel(enClass) < lngClassLevels Then
                                .ErrorText = "Requires " & lngClassLevels & " " & db.Class(enClass).ClassName & " levels"
                                Exit Do
                            End If
                        End If
                        ' Enforce Tier 5 lockouts
                        If .Tier = 5 Then
                            If lngTier5 = 0 Then
                                lngTier5 = .GuideTreeID
                            ElseIf lngTier5 <> .GuideTreeID Then
                                .ErrorText = "Can't take Tier 5s from both " & Guide.Tree(lngTier5).Display & " and " & Guide.Tree(.GuideTreeID).Display
                                Exit Do
                            End If
                        End If
'                        ' Advanced class level reqs (eg: Half-Orc Power Rage) --- this check includes rank-specific class reqs as well
'                        If CheckGuideErrorClassReq(db.Tree(.TreeID).Tier(.Tier).Ability(.Ability), lngClassLevel, .Rank) Then
'                            .ErrorText = "Class level requirements not met"
'                            Exit Do
'                        End If
                        ' Check basic reqs
                        If CheckGuideErrorReqs(i - 1, db.Tree(.TreeID).Tier(.Tier).Ability(.Ability), .Selector, .Rank, lngCurrentLevel) Then
                            .ErrorText = gstrError
                            Exit Do
                        End If
                        ' Check ranks
                        If .Rank > 1 Then
                            If Not FindGuideAbility(i - 1, .TreeID, .Tier, .Ability, .Selector, .Rank - 1) Then
                                .ErrorText = "Rank " & .Rank & " requires first taking Rank " & .Rank - 1
                                Exit Do
                            End If
                        End If
                    Case geResetTree
                        enTreeMap(.TreeID) = ceAny
                        blnTreeMap(.TreeID) = False
                        lngTreeCount = lngTreeCount - 1
                        If lngTier5 = .GuideTreeID Then lngTier5 = 0
                    Case geResetAllTrees
                        ReDim enTreeMap(db.Trees)
                        ReDim blnTreeMap(db.Trees)
                        lngTreeCount = 0
                        lngTier5 = 0
                End Select
            Loop Until True
            .ErrorState = Len(.ErrorText)
        End With
    Next
End Sub

Private Function FindGuideAbility(plngIndex As Long, plngTree As Long, plngTier As Long, plngAbility As Long, plngSelector As Long, plngRank As Long) As Boolean
    Dim blnFound As Boolean
    Dim i As Long
    
    For i = 1 To plngIndex
        With Guide.Enhancement(i)
            Select Case .Style
                Case geEnhancement
                    ' Nested If...Then chain because VB6 doesn't short-circuit
                    If .Ability = plngAbility Then ' Do these in order of most likely to fail first for efficiency
                        If .Tier = plngTier Then
                            If .TreeID = plngTree Then
                                If plngSelector = 0 Or .Selector = plngSelector Then
                                    If plngRank = 0 Or .Rank = plngRank Then
                                        blnFound = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case geResetTree
                    If .TreeID = plngTree Then blnFound = False
                Case geResetAllTrees
                    blnFound = False
            End Select
        End With
    Next
    FindGuideAbility = blnFound
End Function

' Returns TRUE if there are errors
Private Function CheckGuideErrorReqs(plngIndex As Long, ptypAbility As AbilityType, plngSelector As Long, plngRank As Long, plngLevel As Long) As Boolean
    Dim enReq As ReqGroupEnum
    
    gstrError = vbNullString
    CheckGuideErrorReqs = True
    If plngSelector = 0 Then
        For enReq = rgeAll To rgeNone
            If CheckGuideErrorReq(plngIndex, ptypAbility.Req(enReq), enReq, plngLevel) Then Exit Function
        Next
        If plngRank > 1 And ptypAbility.RankReqs Then
            For enReq = rgeAll To rgeNone
                If CheckGuideErrorReq(plngIndex, ptypAbility.Rank(plngRank).Req(enReq), enReq, plngLevel) Then Exit Function
            Next
        End If
    Else
        For enReq = rgeAll To rgeNone
            If CheckGuideErrorReq(plngIndex, ptypAbility.Selector(plngSelector).Req(enReq), enReq, plngLevel) Then Exit Function
        Next
        If plngRank > 1 And ptypAbility.Selector(plngSelector).RankReqs Then
            For enReq = rgeAll To rgeNone
                If CheckGuideErrorReq(plngIndex, ptypAbility.Selector(plngSelector).Rank(plngRank).Req(enReq), enReq, plngLevel) Then Exit Function
            Next
        End If
    End If
    CheckGuideErrorReqs = False
End Function

' Returns TRUE if we fail any reqs
Private Function CheckGuideErrorReq(plngIndex As Long, ptypReqList As ReqListType, penReq As ReqGroupEnum, plngLevel As Long) As Boolean
    Dim lngMatches As Long
    Dim lngTree As Long
    Dim strTaken As String
    Dim strMissing As String
    Dim i As Long

    If ptypReqList.Reqs = 0 Then Exit Function
    For i = 1 To ptypReqList.Reqs
        With ptypReqList.Req(i)
            Select Case .Style
                Case peFeat
                    If CheckGuideErrorFeat(.Feat, .Selector, plngLevel) Then
                        If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, 1)
                        lngMatches = lngMatches + 1
                    Else
                        If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, 1)
                    End If
                Case peEnhancement
                    If FindGuideAbility(plngIndex, .Tree, .Tier, .Ability, .Selector, .Rank) Then
                        If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                        lngMatches = lngMatches + 1
                    Else
                        If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                    End If
            End Select
        End With
    Next
    Select Case penReq
        Case rgeAll
            If lngMatches < ptypReqList.Reqs Then
                gstrError = "Requires " & strMissing
                CheckGuideErrorReq = True
            End If
        Case rgeOne
            If lngMatches < 1 Then
                gstrError = "Nothing taken from the 'One of' list"
                CheckGuideErrorReq = True
            End If
        Case rgeNone
            If lngMatches > 0 Then
                gstrError = "Antireq for " & strTaken
                CheckGuideErrorReq = True
            End If
    End Select
End Function

' Returns TRUE if feat is taken
Private Function CheckGuideErrorFeat(plngFeat As Long, plngSelector As Long, plngLevel As Long) As Boolean
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim lngExchange As Long
    Dim i As Long
    
    For i = 1 To Feat.Count
        If Feat.List(i).Level > plngLevel Then Exit Function
        Select Case Feat.List(i).ActualType
            Case bftAlternate, bftExchange
            Case Else
                lngExchange = Feat.List(i).ExchangeIndex
                If lngExchange Then
                    If Feat.List(lngExchange).Level > plngLevel Then lngExchange = 0
                End If
                If lngExchange Then
                    lngFeat = Feat.List(lngExchange).FeatID
                    lngSelector = Feat.List(lngExchange).Selector
                Else
                    lngFeat = Feat.List(i).FeatID
                    lngSelector = Feat.List(i).Selector
                End If
                If lngFeat = plngFeat Then
                    If plngSelector = 0 Or lngSelector = plngSelector Then
                        CheckGuideErrorFeat = True
                        Exit Function
                    End If
                End If
        End Select
    Next
End Function

'' Returns TRUE if error found
'Private Function CheckGuideErrorClassReq(ptypAbility As AbilityType, plngClassLevel() As Long, plngRank As Long) As Boolean
'    Dim blnPass As Boolean
'    Dim i As Long
'
'    If ptypAbility.Class(0) Then
'        For i = 1 To ceClasses - 1
'            If ptypAbility.Class(i) Then
'                If plngClassLevel(i) >= ptypAbility.ClassLevel(i) Then
'                    blnPass = True
'                    Exit For
'                End If
'            End If
'        Next
'        If Not blnPass Then
'            CheckGuideErrorClassReq = True
'            Exit Function
'        End If
'    End If
'    If plngRank > 1 And ptypAbility.RankReqs Then
'        blnPass = True
'        With ptypAbility.Rank(plngRank)
'             If .Class(0) Then
'                blnPass = False
'                For i = 1 To ceClasses - 1
'                    If .Class(i) Then
'                        If plngClassLevel(i) >= .ClassLevel(i) Then
'                            blnPass = True
'                            Exit For
'                        End If
'                    End If
'                Next
'            End If
'        End With
'        If Not blnPass Then
'            CheckGuideErrorClassReq = True
'            Exit Function
'        End If
'   End If
'End Function

