Attribute VB_Name = "basBuildDefs"
' Written by Ellis Dee
' The build structure is so big I decided to move it to its own module.
' Any versioning code for builds is also here
Option Explicit

' NOTE: When changing the structure of the build, when incrementing the version number
' in this definition, do a global search for the definition first. There are a couple places
' it's used and need to be changed. For example, when clearing a build, a blank
' structure is initialized and copied over it, and that blank structure's definition
' will also need to be upgraded.
Public build As BuildType4
Public Skill As SkillGridType ' Not saved to build file (calculated as needed)
Public Feat As FeatListType ' Not saved to build file (calculated as needed)
Public Guide As GuideType ' Not saved to build file (calculated as needed)

' Used when needing to know the current build's class split
Public Type ClassSplitType
    Levels As Long
    ClassID As ClassEnum
    ClassName As String
    Color As ColorValueEnum
    Initial As String
    BuildClass As Long
End Type

' **** SKILL GRID ****
Public Type SkillRowType
    Ranks As Long
    MaxRanks As Long
End Type

Public Type SkillColType
    Class As Long
    Initial As String
    Color As Long
    Thief As Boolean
    Points As Long
    MaxPoints As Long
End Type

Public Type SkillCellType
    Native As Long ' 1 = Cross class, 2 = Native (Ranks = Native * Points)
    Ranks As Long ' Ranks = 2 * Actual ranks (to avoid floating point math)
    MaxRanks As Long ' 2 * Actual MaxRanks
End Type

Public Type SkillMapType
    Skill As Long
    SkillName As String
End Type

Public Type SkillOutputType
    Skill As Long
    Ranks As Long
End Type

Public Type SkillGridType
    Row(1 To 21) As SkillRowType
    Col(1 To 20) As SkillColType
    grid(1 To 21, 1 To 20) As SkillCellType
    Map(1 To 21) As SkillMapType
    Out() As SkillOutputType
End Type
' **** END SKILL GRID ****

' *** FEAT LIST ***
Public Type FeatDetailType
    ActualType As BuildFeatTypeEnum
    EffectiveType As BuildFeatTypeEnum ' For alternates and exchanges, this is set to the type of the parent feat
    ParentIndex As Long ' build.Feat(EffectiveType).Feat(ParentIndex)
    Index As Long
'    Slot As Long
    Channel As FeatChannelEnum
    ChannelSlot() As Long ' List of slot positions in the various channels
    Level As Long
    Class As ClassEnum
    ClassLevel As Long
    SourceForm As String
    SourceOutput As String
    SourceFilter As String
    FeatID As Long
    Selector As Long
    FeatName As String
    Display As String
    DisplayAlternate As String
    ExchangeIndex As Long ' Points to index in this array
    ErrorState As Boolean
    ErrorText As String
    Flag As Boolean ' This is set to true only for newly added exchange feat
End Type

Public Type FeatListType
    List() As FeatDetailType
    Count As Long
    ChannelCount() As Long
    Granted As Long
    Selected As Long
End Type
' *** END FEAT LIST ***

Public Type BuildSpellListType
    Class As Byte
    ClassLevels As Byte
    MaxSpellLevel As Byte
    SpellList() As SpellListType ' Index = Spell Level
End Type

Public Type BuildSpellSlotType
    SlotType As Byte ' 0=Standard, 1=Free, 2=ClericCure, 3=WarlockPact
    Level As Byte
    Spell As String
End Type

Public Type BuildClassLevelType
    Slot() As BuildSpellSlotType
    Slots As Byte
    BaseSlots As Byte
    FreeSlots As Byte
    Mandatory As Byte ' MandatorySpellEnum: 0=None, (1=Free isn't used here), 2=Cleric Cure, 3=Warlock Pact
End Type

Public Type BuildClassSpellType
    Class As Byte
    ClassLevels As Byte
    MaxSpellLevel As Byte
    Level() As BuildClassLevelType
End Type

Public Type BuildAbilityType
    Tier As Byte ' 0 = core
    Ability As Byte ' Index in db (If trees change, existing builds become unpredictable, as opposed to invalid if we stored names)
    Rank As Byte
    Selector As Byte ' Index of selector (0 if no selector)
End Type

Public Type BuildTreeType
    TreeName As String
    TreeType As Byte
    Source As Byte ' ClassEnum, or 0 for global (eg: Harper) or racial class tree (eg: Elf-AA)
    ClassLevels As Byte
    Ability() As BuildAbilityType
    Abilities As Byte
End Type

Public Type TwistType
    DestinyName As String
    Tier As Byte
    Ability As Byte
    Selector As Byte
End Type

Public Type BuildFeatType
    Source As Byte ' BuildFeatSourceEnum - Tells us what to display in "Source" column of Feats screen
    Type As Byte ' BuildFeatTypeEnum - Tells us which index to access in Build.Feats(Index) array
    Level As Byte ' Character level
    ClassLevel As Byte
    FeatName As String
    Selector As Byte
    Child As Byte ' Index of the feat being replaced (or attaching an alternate)
    ChildType As Byte ' Child's type: 0 = Granted, 1 = Standard, 2 = Race, 3 = Class1, 4 = Class2, 5 = Class3, 6 = Deity, 7 = this feat is itself a child being exchanged
End Type

Public Type BuildFeatListType
    Feat() As BuildFeatType
    Feats As Byte
End Type

' Original version (0.9.4 and earlier)
Public Type BuildType1
    BuildName As String
    Race As Byte
    Alignment As Byte
    MaxLevels As Byte
    BuildClass(2) As Byte ' The three classes chosen in the Overview screen
    Class(1 To 20) As Byte
    BAB() As Byte ' 1 to MaxLevel
    StatPoints(3, 6) As Byte ' 28/32/34/36 versions, stored as Points Spent
    BuildPoints As Byte ' Which version is the "main"? (BuildPointsEnum; 0=28pt)
    IncludePoints(3) As Byte ' Which build point versions to include in display/export
    Levelups(7) As Byte
    Tome(6) As Byte
    Skills(1 To 21, 1 To 20) As Byte
    Feat() As BuildFeatListType
    CanCastSpell(9) As Byte  ' Index 0 = healing spells; if index 1 = 0, build can't cast spells at all (never gets level 1 spells)
    ClassSpell() As BuildSpellListType
    ClassSpells As Byte
    Tree() As BuildTreeType
    Trees As Byte
    Tier5 As String
    Destiny As BuildTreeType
    Twist() As TwistType
    Twists As Byte
End Type

' Version 2 (0.9.5) adds Author, Link and TomeEarly
Public Type BuildType2
    Author As String
    Link As String
    BuildName As String
    Race As Byte
    Alignment As Byte
    MaxLevels As Byte
    BuildClass(2) As Byte ' The three classes chosen in the Overview screen
    Class(1 To 20) As Byte
    BAB() As Byte ' 1 to MaxLevel
    StatPoints(3, 6) As Byte ' 28/32/34/36 versions, stored as Points Spent
    BuildPoints As Byte ' Which version is the "main"? (BuildPointsEnum; 0=28pt)
    IncludePoints(3) As Byte ' Which build point versions to include in display/export
    Levelups(7) As Byte
    Tome(6) As Byte
    TomeEarly As Byte
    Skills(1 To 21, 1 To 20) As Byte
    Feat() As BuildFeatListType
    CanCastSpell(9) As Byte  ' Index 0 = healing spells; if index 1 = 0, build can't cast spells at all (never gets level 1 spells)
    ClassSpell() As BuildSpellListType
    ClassSpells As Byte
    Tree() As BuildTreeType
    Trees As Byte
    Tier5 As String
    Destiny As BuildTreeType
    Twist() As TwistType
    Twists As Byte
End Type

' **** LEVELING GUIDE ****
' All possible trees that could be part of Leveling Guide
Public Type GuideTreeType
    GuideTreeID As Long ' Index in this array
    TreeID As Long ' db.Tree() index
    BuildTreeID As Long ' build.Tree() index
    BuildGuideTreeID As Long ' build.Guide.Tree() index
    TreeName As String
    Abbreviation As String
    Initial As String
    Color As ColorValueEnum
    Display As String ' for dups; eg: Warpriest (Clr)
    TreeStyle As TreeStyleEnum
    Class As ClassEnum
    ClassLevels As Long
    Duplicate As Boolean ' True if multiple classes in build grant this tree
    BuildTree As BuildTreeType ' Temporary BuildTree structure used for checking prereqs
    Spent As Long ' Used by Output
    MaxTier As Long
End Type

' Perfect mirror of build.Guide.Enhancement() udt array storing additional calculated info
Public Type GuideEnhancementType
    Style As GuideEnhancementEnum
    GuideTreeID As Long ' Guide.Tree() index
    TreeID As Long ' db.Tree() index
    BuildTreeID As Long ' build.Tree() index
    BuildGuideTreeID As Long ' build.Guide.Tree() index
    Tier As Long
    Ability As Long
    Selector As Long
    Rank As Long
    Display As String
    RankText As String
    Level As Long
    Cost As Long
    SpentInTree As Long
    Spent As Long
    SpentRacial As Long
    ML As Long
    Bank As Long
    ErrorState As Boolean
    ErrorText As String
    Selected As Boolean
End Type

Public Type GuideType
    Tree() As GuideTreeType
    Trees As Long
    Enhancement() As GuideEnhancementType
    Enhancements As Long
    TreeLookup() As Long ' Guide.Tree(TreeLookup(BuildGuideTreeID))
End Type
' **** END LEVELING GUIDE ****

Public Type BuildGuideTreeType
    TreeName As String
    Class As Byte
    OutputColor As Byte ' ColorValueEnum
End Type

Public Type BuildGuideEnhancementType
    ID As Byte ' 0=unknown, 1-99=BuildGuideTree index, 100=Reset all trees, 101-199=Reset tree (BGT index+100), 200=Bank remaining AP
    Tier As Byte
    Ability As Byte
    Selector As Byte
    Rank As Byte
End Type

Public Type BuildGuideType
    Tree() As BuildGuideTreeType
    Trees As Byte
    Enhancement() As BuildGuideEnhancementType
    Enhancements As Byte
End Type

' Version 3 (2.0) many changes
Public Type BuildType3
    Notes As String
    BuildName As String
    Race As Byte
    Alignment As Byte
    MaxLevels As Byte
    BuildClass(2) As Byte ' The three classes chosen in the Overview screen
    Class(1 To 20) As Byte
    BAB() As Byte ' 1 to MaxLevel
    StatPoints(3, 6) As Byte ' 28/32/34/36 versions, stored as Points Spent
    BuildPoints As Byte ' Which version is the "main"? (BuildPointsEnum; 0=28pt)
    IncludePoints(3) As Byte ' Which build point versions to include in display/export
    Levelups(7) As Byte
    Tome(6) As Byte
    Skills(1 To 21, 1 To 20) As Byte
    SkillTome(1 To 21) As Byte
    Feat() As BuildFeatListType
    CanCastSpell(9) As Byte  ' Index 0 = healing spells; if index 1 = 0, build can't cast spells at all (never gets level 1 spells)
    Spell() As BuildClassSpellType
    Tree() As BuildTreeType
    Trees As Byte
    Guide As BuildGuideType
    Tier5 As String
    Destiny As BuildTreeType
    Twist() As TwistType
    Twists As Byte
End Type

' Version 4 (2.2) add Racial AP
Public Type BuildType4
    Notes As String
    BuildName As String
    Race As Byte
    Alignment As Byte
    MaxLevels As Byte
    BuildClass(2) As Byte ' The three classes chosen in the Overview screen
    Class(1 To 20) As Byte
    BAB() As Byte ' 1 to MaxLevel
    StatPoints(3, 6) As Byte ' 28/32/34/36 versions, stored as Points Spent
    BuildPoints As Byte ' Which version is the "main"? (BuildPointsEnum; 0=28pt)
    IncludePoints(3) As Byte ' Which build point versions to include in display/export
    Levelups(7) As Byte
    Tome(6) As Byte
    Skills(1 To 21, 1 To 20) As Byte
    SkillTome(1 To 21) As Byte
    Feat() As BuildFeatListType
    CanCastSpell(9) As Byte  ' Index 0 = healing spells; if index 1 = 0, build can't cast spells at all (never gets level 1 spells)
    Spell() As BuildClassSpellType
    Tree() As BuildTreeType
    Trees As Byte
    Guide As BuildGuideType
    Tier5 As String
    RacialAP As Byte
    Destiny As BuildTreeType
    Twist() As TwistType
    Twists As Byte
End Type

Public Sub Version1To2(ptypBuild1 As BuildType1, ptypBuild2 As BuildType2)
    Dim i As Long
    Dim j As Long
    
    With ptypBuild2
        .Author = vbNullString
        .Link = vbNullString
        .BuildName = ptypBuild1.BuildName
        .Race = ptypBuild1.Race
        .Alignment = ptypBuild1.Alignment
        .MaxLevels = ptypBuild1.MaxLevels
        For i = 0 To 2
            .BuildClass(i) = ptypBuild1.BuildClass(i)
        Next
        For i = 1 To 20
            .Class(i) = ptypBuild1.Class(i)
        Next
        .BAB = ptypBuild1.BAB
        For i = 0 To 3
            For j = 0 To 6
                .StatPoints(i, j) = ptypBuild1.StatPoints(i, j)
            Next
        Next
        .BuildPoints = ptypBuild1.BuildPoints
        For i = 0 To 3
            .IncludePoints(i) = ptypBuild1.IncludePoints(i)
        Next
        For i = 0 To 7
            .Levelups(i) = ptypBuild1.Levelups(i)
        Next
        For i = 0 To 6
            .Tome(i) = ptypBuild1.Tome(i)
        Next
        Select Case .BuildPoints
            Case beAdventurer, beChampion: .TomeEarly = 0
            Case beHero, beLegend: .TomeEarly = 1
        End Select
        For i = 1 To 21
            For j = 1 To 20
                .Skills(i, j) = ptypBuild1.Skills(i, j)
            Next
        Next
        .Feat = ptypBuild1.Feat
        For i = 0 To 9
            .CanCastSpell(i) = ptypBuild1.CanCastSpell(i)
        Next
        .ClassSpell = ptypBuild1.ClassSpell
        .ClassSpells = ptypBuild1.ClassSpells
        .Tree = ptypBuild1.Tree
        .Trees = ptypBuild1.Trees
        .Tier5 = ptypBuild1.Tier5
        .Destiny = ptypBuild1.Destiny
        .Twist = ptypBuild1.Twist
        .Twists = ptypBuild1.Twists
    End With
End Sub

Public Sub Version2To3(ptyp2 As BuildType2, ptyp3 As BuildType3)
    Dim typBlank As BuildType3
    Dim i As Long
    Dim j As Long
    
    ptyp3 = typBlank
    With ptyp3
        If Len(ptyp2.Author) Then .Notes = "Build by " & ptyp2.Author & vbNewLine
        If Len(ptyp2.Link) Then .Notes = .Notes & ptyp2.Link
        .BuildName = ptyp2.BuildName
        .Race = ptyp2.Race
        .Alignment = ptyp2.Alignment
        .MaxLevels = ptyp2.MaxLevels
        If .MaxLevels = 28 Then .MaxLevels = 30
        For i = 0 To 2
            .BuildClass(i) = ptyp2.BuildClass(i)
        Next
        For i = 1 To 20
            .Class(i) = ptyp2.Class(i)
        Next
        .BAB = ptyp2.BAB
        ReDim Preserve .BAB(1 To 30)
        .BAB(29) = .BAB(28) + 1
        .BAB(30) = .BAB(29)
        For i = 0 To 3
            For j = 0 To 6
                .StatPoints(i, j) = ptyp2.StatPoints(i, j)
            Next
        Next
        .BuildPoints = ptyp2.BuildPoints
        For i = 0 To 3
            .IncludePoints(i) = ptyp2.IncludePoints(i)
        Next
        For i = 0 To 7
            .Levelups(i) = ptyp2.Levelups(i)
        Next
        For i = 0 To 6
            .Tome(i) = ptyp2.Tome(i)
        Next
        For i = 1 To 21
            For j = 1 To 20
                .Skills(i, j) = ptyp2.Skills(i, j)
            Next
        Next
        ReDim .Feat(bftFeatTypes - 1)
        For i = 0 To 9
            Select Case i
                Case 0 To 1: .Feat(i) = ptyp2.Feat(i)
                Case 3 To 9: .Feat(i) = ptyp2.Feat(i - 1)
            End Select
        Next
        For i = 0 To 9
            .CanCastSpell(i) = ptyp2.CanCastSpell(i)
        Next
        .Tree = ptyp2.Tree
        .Trees = ptyp2.Trees
        .Tier5 = ptyp2.Tier5
        .Destiny = ptyp2.Destiny
        .Twist = ptyp2.Twist
        .Twists = ptyp2.Twists
    End With
    InitBuildSpells
    If build.CanCastSpell(1) > 0 Then Spells2To3 build.Spell, ptyp2.ClassSpell, ptyp2.ClassSpells
End Sub

Private Sub Spells2To3(ptypSpell() As BuildClassSpellType, ptypClassSpell() As BuildSpellListType, ByVal plngClassSpells As Long)
    Dim lngClassSpell As Long
    Dim enClass As Long
    Dim lngLevel As Long
    Dim lngSlot As Long
    Dim strSpell As String
    
    For lngClassSpell = 1 To plngClassSpells
        enClass = ptypClassSpell(lngClassSpell).Class
        For lngLevel = 1 To ptypClassSpell(lngClassSpell).MaxSpellLevel
            If lngLevel <= ptypSpell(enClass).MaxSpellLevel Then
                ' First, add any used free slots
                For lngSlot = 1 To ptypClassSpell(lngClassSpell).SpellList(lngLevel).Spells
                    strSpell = ptypClassSpell(lngClassSpell).SpellList(lngLevel).Spell(lngSlot)
                    If strSpell = "<Any>" Then strSpell = vbNullString
                    If Len(strSpell) Then
                        If CheckFree(enClass, strSpell) Then
                            AddFreeSpellSlot enClass, lngLevel
                        End If
                    End If
                Next
                ' Now apply spells
                For lngSlot = 1 To ptypClassSpell(lngClassSpell).SpellList(lngLevel).Spells
                    If lngSlot <= ptypSpell(enClass).Level(lngLevel).Slots Then
                        Select Case ptypSpell(enClass).Level(lngLevel).Slot(lngSlot).SlotType
                            Case sseStandard, sseFree
                                strSpell = ptypClassSpell(lngClassSpell).SpellList(lngLevel).Spell(lngSlot)
                                If strSpell = "<Any>" Then strSpell = vbNullString
                                ptypSpell(enClass).Level(lngLevel).Slot(lngSlot).Spell = strSpell
                        End Select
                    End If
                Next
            End If
        Next
    Next
End Sub

Public Sub Version3To4(ptyp3 As BuildType3, ptyp4 As BuildType4)
    Dim typBlank As BuildType4
    Dim i As Long
    Dim j As Long
    
    ptyp4 = typBlank
    With ptyp4
        .Notes = ptyp3.Notes
        .BuildName = ptyp3.BuildName
        .Race = ptyp3.Race
        .Alignment = ptyp3.Alignment
        .MaxLevels = ptyp3.MaxLevels
        For i = 0 To 2
            .BuildClass(i) = ptyp3.BuildClass(i)
        Next
        For i = 1 To 20
            .Class(i) = ptyp3.Class(i)
        Next
        .BAB = ptyp3.BAB
        For i = 0 To 3
            For j = 0 To 6
                .StatPoints(i, j) = ptyp3.StatPoints(i, j)
            Next
        Next
        .BuildPoints = ptyp3.BuildPoints
        For i = 0 To 3
            .IncludePoints(i) = ptyp3.IncludePoints(i)
        Next
        For i = 0 To 7
            .Levelups(i) = ptyp3.Levelups(i)
        Next
        For i = 0 To 6
            .Tome(i) = ptyp3.Tome(i)
        Next
        For i = 1 To 21
            For j = 1 To 20
                .Skills(i, j) = ptyp3.Skills(i, j)
            Next
            .SkillTome(i) = ptyp3.SkillTome(i)
        Next
        .Feat = ptyp3.Feat
        For i = 0 To 9
            .CanCastSpell(i) = ptyp3.CanCastSpell(i)
        Next
        .Spell = ptyp3.Spell
        .Tree = ptyp3.Tree
        .Trees = ptyp3.Trees
        .Guide = ptyp3.Guide
        .Tier5 = ptyp3.Tier5
        .RacialAP = 0
        .Destiny = ptyp3.Destiny
        .Twist = ptyp3.Twist
        .Twists = ptyp3.Twists
    End With
End Sub

