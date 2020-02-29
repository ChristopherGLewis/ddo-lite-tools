Attribute VB_Name = "basEnum"
' Written by Ellis Dee
' I would normally put these Enumerations in the appropriate modules, but there's so many
' of them I decided to stash them all here to help reduce clutter everywhere else.
' And since they're all here together, I put the GetID() and GetName() functions here too.
Option Explicit

Public Enum StatEnum
    aeAny
    aeStr
    aeDex
    aeCon
    aeInt
    aeWis
    aeCha
End Enum

Public Enum AlignmentEnum
    aleAny
    aleTrueNeutral
    aleNeutralGood
    aleLawfulNeutral
    aleLawfulGood
    aleChaoticNeutral
    aleChaoticGood
End Enum

Public Enum RaceEnum
    reAny
    reHuman
    reDrow
    reDwarf
    reElf
    reHalfling
    reHalfElf
    reHalfOrc
    reWarforged
    reBladeforged
    rePurpleDragonKnight
    reMorninglord
    reShadarKai
    reGnome
    reDeepGnome
    reDragonborn
    reAasimar
    reScourge
    reWoodElf
    reTiefling
    reScoundrel
    reRaces
End Enum

Public Enum RaceTypeEnum
    rteUnknown
    rteFree
    rtePremium
    rteIconic
End Enum

Public Enum ClassEnum
    ceAny
    ceBarbarian
    ceBard
    ceCleric
    ceFighter
    cePaladin
    ceRanger
    ceRogue
    ceSorcerer
    ceWizard
    ceMonk
    ceFavoredSoul
    ceArtificer
    ceDruid
    ceWarlock
    ceAlchemist
    ceClasses
    ceEmpty
End Enum

Public Enum SkillsEnum
    seAny
    seBalance
    seBluff
    seConcentration
    seDiplomacy
    seDisableDevice
    seHaggle
    seHeal
    seHide
    seIntimidate
    seJump
    seListen
    seMoveSilently
    seOpenLock
    sePerform
    seRepair
    seSearch
    seSpellcraft
    seSpot
    seSwim
    seTumble
    seUMD
    seSkills
End Enum

Public Enum BABEnum
    beFull
    beThreeQuarters
    beHalf
End Enum

Public Enum SpellSlotEnum
    sseStandard
    sseFree
    sseClericCure
    sseWarlockPact
End Enum

Public Enum ReqEnum
    reqStat = 1
    reqSkill
    reqBAB
    reqLevel
    reqCastSpellLevel
    reqClassLevel
    reqRace
End Enum

Public Enum PointerEnum
    peError
    peFeat
    peEnhancement
    peDestiny
End Enum

Public Enum ReqGroupEnum
    rgeAll = 1
    rgeOne
    rgeNone
End Enum

Public Enum FilterEnum
    feAll
    feMelee
    feRange
    feSpellcasting
    feDefense
    feMisc
    feHeroic
    feEpic
    feDestiny
    feLegend
    feFilters
End Enum

Public Enum BuildFeatSourceEnum ' Source column in feats screen (display only)
    bfsHeroic
    bfsEpic
    bfsDestiny
    bfsRace
    bfsClass
    bfsClassOnly
    bfsDeity
    bfsLegend
End Enum

Public Enum BuildFeatTypeEnum ' Feat type: build.Feat(Type).Feat(Index)
    bftGranted
    bftStandard
    bftLegend
    bftRace
    bftClass1
    bftClass2
    bftClass3
    bftDeity
    bftAlternate
    bftExchange
    bftFeatTypes
    bftUnknown
End Enum

Public Enum FeatChannelEnum
    fceUnknown
    fceSelected
    fceGeneral
    fceRacial
    fceWarlock
    fceMonk
    fceRogue
    fceFavoredEnemy
    fceWildShape
    fceCleric
    fceFavoredSoul
    fceEnergy
    fceDeity
    fceGranted
    fceChannels
End Enum

Public Enum TreeStyleEnum
    tseRace = 1
    tseClass
    tseDestiny
    tseGlobal
    tseRaceClass
End Enum

Public Enum SelectorStyleEnum
    sseNone
    sseRoot
    sseShared ' Each ability must use same selector (eg: Magister school focus)
    sseExclusive ' Each ability must use a different selector (eg: Improved Defender Stance)
End Enum

Public Enum DragEnum
    dragNormal
    dragMouseDown
    dragMouseMove
End Enum

Public Enum CascadeChangeEnum
    cceAll
    cceMaxLevel
    cceRace
    cceAlignment
    cceClass
    cceStats
    cceSkill
    cceFeat
    cceEnhancements
End Enum

Public Enum GuideEnhancementEnum
    geUnknown
    geEnhancement
    geResetTree
    geResetAllTrees
    geBankAP
End Enum

Public Enum PlannerEnum
    peUnknown
    peLite
    peRon
    peBuilder
End Enum

Public Enum RaceReqEnum
    rreAny
    rreRequired
    rreNotAllowed
    rreStandard
    rreIconic
End Enum

Public Function GetBuildPoints(ByVal penBuildPoints As BuildPointsEnum) As Long
    Dim lngOffset As Long
    
    If build.Race = reDrow Then lngOffset = -4
    Select Case penBuildPoints
        Case beAdventurer: GetBuildPoints = 28
        Case beChampion: GetBuildPoints = 32 + lngOffset
        Case beHero: GetBuildPoints = 34 + lngOffset
        Case beLegend: GetBuildPoints = 36 + lngOffset
    End Select
End Function

Public Function GetAlignmentID(pstrAlignment As String) As AlignmentEnum
    Select Case LCase$(pstrAlignment)
        Case "true neutral": GetAlignmentID = aleTrueNeutral
        Case "neutral good": GetAlignmentID = aleNeutralGood
        Case "lawful neutral": GetAlignmentID = aleLawfulNeutral
        Case "lawful good": GetAlignmentID = aleLawfulGood
        Case "chaotic neutral": GetAlignmentID = aleChaoticNeutral
        Case "chaotic good": GetAlignmentID = aleChaoticGood
    End Select
End Function

Public Function GetAlignmentName(ByVal penAlignment As AlignmentEnum) As String
    Select Case penAlignment
        Case aleTrueNeutral: GetAlignmentName = "True Neutral"
        Case aleNeutralGood: GetAlignmentName = "Neutral Good"
        Case aleLawfulNeutral: GetAlignmentName = "Lawful Neutral"
        Case aleLawfulGood: GetAlignmentName = "Lawful Good"
        Case aleChaoticNeutral: GetAlignmentName = "Chaotic Neutral"
        Case aleChaoticGood: GetAlignmentName = "Chaotic Good"
    End Select
End Function

Public Function GetRaceID(pstrRace As String) As RaceEnum
    Select Case LCase$(pstrRace)
        Case "drow": GetRaceID = reDrow
        Case "dwarf": GetRaceID = reDwarf
        Case "elf": GetRaceID = reElf
        Case "gnome": GetRaceID = reGnome
        Case "halfling": GetRaceID = reHalfling
        Case "half-elf": GetRaceID = reHalfElf
        Case "half-orc": GetRaceID = reHalfOrc
        Case "human": GetRaceID = reHuman
        Case "warforged": GetRaceID = reWarforged
        Case "bladeforged": GetRaceID = reBladeforged
        Case "purple dragon knight": GetRaceID = rePurpleDragonKnight
        Case "morninglord": GetRaceID = reMorninglord
        Case "shadar-kai": GetRaceID = reShadarKai
        Case "deep gnome": GetRaceID = reDeepGnome
        Case "dragonborn": GetRaceID = reDragonborn
        Case "aasimar": GetRaceID = reAasimar
        Case "scourge", "scourge aasimar", "aasimar scourge": GetRaceID = reScourge
        Case "wood elf": GetRaceID = reWoodElf
        Case "tiefling": GetRaceID = reTiefling
        Case "scoundrel", "tiefling scoundrel": GetRaceID = reScoundrel
    End Select
End Function

Public Function GetRaceName(ByVal penRace As RaceEnum) As String
    Select Case penRace
        Case reDrow: GetRaceName = "Drow"
        Case reDwarf: GetRaceName = "Dwarf"
        Case reElf: GetRaceName = "Elf"
        Case reGnome: GetRaceName = "Gnome"
        Case reHalfling: GetRaceName = "Halfling"
        Case reHalfElf: GetRaceName = "Half-Elf"
        Case reHalfOrc: GetRaceName = "Half-Orc"
        Case reHuman: GetRaceName = "Human"
        Case reWarforged: GetRaceName = "Warforged"
        Case reBladeforged: GetRaceName = "Bladeforged"
        Case rePurpleDragonKnight: GetRaceName = "Purple Dragon Knight"
        Case reMorninglord: GetRaceName = "Morninglord"
        Case reShadarKai: GetRaceName = "Shadar-kai"
        Case reDeepGnome: GetRaceName = "Deep Gnome"
        Case reDragonborn: GetRaceName = "Dragonborn"
        Case reAasimar: GetRaceName = "Aasimar"
        Case reScourge: GetRaceName = "Aasimar Scourge"
        Case reWoodElf: GetRaceName = "Wood Elf"
        Case reTiefling: GetRaceName = "Tiefling"
        Case reScoundrel: GetRaceName = "Tiefling Scoundrel"
    End Select
End Function

Public Function GetRaceTypeID(pstrRace As String) As RaceTypeEnum
    Select Case LCase$(pstrRace)
        Case "free": GetRaceTypeID = rteFree
        Case "premium": GetRaceTypeID = rtePremium
        Case "iconic": GetRaceTypeID = rteIconic
    End Select
End Function

Public Function GetRaceTypeName(ByVal penType As RaceTypeEnum) As String
    Select Case penType
        Case rteFree: GetRaceTypeName = "Free to Play"
        Case rtePremium: GetRaceTypeName = "Premium"
        Case rteIconic: GetRaceTypeName = "Iconic"
    End Select
End Function

Public Function GetClassID(ByVal pstrClass As String) As ClassEnum
    Select Case LCase$(pstrClass)
        Case "alchemist": GetClassID = ceAlchemist
        Case "artificer": GetClassID = ceArtificer
        Case "barbarian": GetClassID = ceBarbarian
        Case "bard": GetClassID = ceBard
        Case "cleric": GetClassID = ceCleric
        Case "druid": GetClassID = ceDruid
        Case "favored soul": GetClassID = ceFavoredSoul
        Case "fighter": GetClassID = ceFighter
        Case "monk": GetClassID = ceMonk
        Case "paladin": GetClassID = cePaladin
        Case "ranger": GetClassID = ceRanger
        Case "rogue": GetClassID = ceRogue
        Case "sorcerer": GetClassID = ceSorcerer
        Case "warlock": GetClassID = ceWarlock
        Case "wizard": GetClassID = ceWizard
    End Select
End Function

Public Function GetClassName(ByVal penClass As ClassEnum, Optional pblnAbbreviation As Boolean) As String
    Select Case penClass
        Case ceAlchemist: GetClassName = "Alchemist"
        Case ceArtificer: GetClassName = "Artificer"
        Case ceBarbarian: GetClassName = "Barbarian"
        Case ceBard: GetClassName = "Bard"
        Case ceCleric: GetClassName = "Cleric"
        Case ceDruid: GetClassName = "Druid"
        Case ceFavoredSoul: GetClassName = "Favored Soul"
        Case ceFighter: GetClassName = "Fighter"
        Case ceMonk: GetClassName = "Monk"
        Case cePaladin: GetClassName = "Paladin"
        Case ceRanger: GetClassName = "Ranger"
        Case ceRogue: GetClassName = "Rogue"
        Case ceSorcerer: GetClassName = "Sorcerer"
        Case ceWarlock: GetClassName = "Warlock"
        Case ceWizard: GetClassName = "Wizard"
    End Select
End Function

Public Function GetClassResourceID(ByVal penClass As ClassEnum) As String
    Select Case penClass
        Case ceAlchemist: GetClassResourceID = "CLSALCHEMIST"
        Case ceArtificer: GetClassResourceID = "CLSARTIFICER"
        Case ceBarbarian: GetClassResourceID = "CLSBARBARIAN"
        Case ceBard: GetClassResourceID = "CLSBARD"
        Case ceCleric: GetClassResourceID = "CLSCLERIC"
        Case ceDruid: GetClassResourceID = "CLSDRUID"
        Case ceEmpty: GetClassResourceID = "CLSEMPTY"
        Case ceFavoredSoul: GetClassResourceID = "CLSFAVOREDSOUL"
        Case ceFighter: GetClassResourceID = "CLSFIGHTER"
        Case ceMonk: GetClassResourceID = "CLSMONK"
        Case cePaladin: GetClassResourceID = "CLSPALADIN"
        Case ceRanger: GetClassResourceID = "CLSRANGER"
        Case ceRogue: GetClassResourceID = "CLSROGUE"
        Case ceSorcerer: GetClassResourceID = "CLSSORCERER"
        Case ceWarlock: GetClassResourceID = "CLSWARLOCK"
        Case ceWizard: GetClassResourceID = "CLSWIZARD"
    End Select
End Function

Public Function GetSkillID(ByVal pstrSkill As String) As SkillsEnum
    Select Case LCase$(pstrSkill)
        Case "balance": GetSkillID = seBalance
        Case "bluff": GetSkillID = seBluff
        Case "concentration": GetSkillID = seConcentration
        Case "diplomacy": GetSkillID = seDiplomacy
        Case "disable device", "disabledevice": GetSkillID = seDisableDevice
        Case "haggle": GetSkillID = seHaggle
        Case "heal": GetSkillID = seHeal
        Case "hide": GetSkillID = seHide
        Case "intimidate": GetSkillID = seIntimidate
        Case "jump": GetSkillID = seJump
        Case "listen": GetSkillID = seListen
        Case "move silently", "movesilently": GetSkillID = seMoveSilently
        Case "open lock", "openlock": GetSkillID = seOpenLock
        Case "perform": GetSkillID = sePerform
        Case "repair": GetSkillID = seRepair
        Case "search": GetSkillID = seSearch
        Case "spellcraft": GetSkillID = seSpellcraft
        Case "spot": GetSkillID = seSpot
        Case "swim": GetSkillID = seSwim
        Case "tumble": GetSkillID = seTumble
        Case "umd", "use magic device": GetSkillID = seUMD
    End Select
End Function

Public Function GetSkillName(ByVal penSkill As SkillsEnum, Optional pblnAbbreviate As Boolean = False) As String
    Select Case penSkill
        Case seBalance: GetSkillName = "Balance"
        Case seBluff: GetSkillName = "Bluff"
        Case seConcentration: If pblnAbbreviate Then GetSkillName = "Concent" Else GetSkillName = "Concentration"
        Case seDiplomacy: If pblnAbbreviate Then GetSkillName = "Diplo" Else GetSkillName = "Diplomacy"
        Case seDisableDevice: If pblnAbbreviate Then GetSkillName = "Disable" Else GetSkillName = "Disable Device"
        Case seHaggle: GetSkillName = "Haggle"
        Case seHeal: GetSkillName = "Heal"
        Case seHide: GetSkillName = "Hide"
        Case seIntimidate: If pblnAbbreviate Then GetSkillName = "Intim" Else GetSkillName = "Intimidate"
        Case seJump: GetSkillName = "Jump"
        Case seListen: GetSkillName = "Listen"
        Case seMoveSilently: If pblnAbbreviate Then GetSkillName = "Move Si" Else GetSkillName = "Move Silently"
        Case seOpenLock: If pblnAbbreviate Then GetSkillName = "Open Lo" Else GetSkillName = "Open Lock"
        Case sePerform: GetSkillName = "Perform"
        Case seRepair: GetSkillName = "Repair"
        Case seSearch: GetSkillName = "Search"
        Case seSpellcraft: If pblnAbbreviate Then GetSkillName = "Spellcr" Else GetSkillName = "Spellcraft"
        Case seSpot: GetSkillName = "Spot"
        Case seSwim: GetSkillName = "Swim"
        Case seTumble: GetSkillName = "Tumble"
        Case seUMD: If pblnAbbreviate Then GetSkillName = "UMD" Else GetSkillName = "Use Magic Device"
    End Select
End Function

Public Function GetStatID(pstrStat As String) As Long
    Select Case LCase$(pstrStat)
        Case "strength", "str": GetStatID = aeStr
        Case "dexterity", "dex": GetStatID = aeDex
        Case "constitution", "con": GetStatID = aeCon
        Case "intelligence", "int": GetStatID = aeInt
        Case "wisdom", "wis": GetStatID = aeWis
        Case "charisma", "cha": GetStatID = aeCha
    End Select
End Function

Public Function GetStatName(ByVal penStat As StatEnum, Optional pblnAbbreviate As Boolean = False) As String
    Dim strReturn As String
    
    Select Case penStat
        Case aeStr: strReturn = "Strength"
        Case aeDex: strReturn = "Dexterity"
        Case aeCon: strReturn = "Constitution"
        Case aeInt: strReturn = "Intelligence"
        Case aeWis: strReturn = "Wisdom"
        Case aeCha: strReturn = "Charisma"
    End Select
    If pblnAbbreviate Then GetStatName = Left$(strReturn, 3) Else GetStatName = strReturn
End Function

Public Function GetReqGroupID(pstrReqGroup As String) As ReqGroupEnum
    Select Case pstrReqGroup
        Case "all": GetReqGroupID = rgeAll
        Case "one": GetReqGroupID = rgeOne
        Case "none": GetReqGroupID = rgeNone
    End Select
End Function

Public Function GetReqGroupName(ByVal penReqGroup As ReqGroupEnum) As String
    Select Case penReqGroup
        Case rgeAll: GetReqGroupName = "All"
        Case rgeOne: GetReqGroupName = "One"
        Case rgeNone: GetReqGroupName = "None"
    End Select
End Function

Public Function GetGroupID(pstrGroup As String) As FilterEnum
    Dim enGroup As FilterEnum
    
    Select Case LCase$(Trim$(pstrGroup))
        Case "heroic": enGroup = feHeroic
        Case "melee": enGroup = feMelee
        Case "range": enGroup = feRange
        Case "spellcasting": enGroup = feSpellcasting
        Case "defense": enGroup = feDefense
        Case "misc": enGroup = feMisc
        Case "epic": enGroup = feEpic
        Case "destiny": enGroup = feDestiny
        Case "legend": enGroup = feLegend
    End Select
    GetGroupID = enGroup
End Function

Public Function GetFeatGroupName(ByVal penFeatGroup As FilterEnum) As String
    Select Case penFeatGroup
        Case feAll: GetFeatGroupName = "Show All Feats"
        Case feHeroic: GetFeatGroupName = "Heroic"
        Case feMelee: GetFeatGroupName = "Melee"
        Case feRange: GetFeatGroupName = "Range"
        Case feSpellcasting: GetFeatGroupName = "Spellcasting"
        Case feDefense: GetFeatGroupName = "Defense"
        Case feMisc: GetFeatGroupName = "Misc"
        Case feEpic: GetFeatGroupName = "Epic"
        Case feDestiny: GetFeatGroupName = "Destiny"
        Case feLegend: GetFeatGroupName = "Legend"
    End Select
End Function

Public Function GetTreeStyleName(ByVal penTreeType As TreeStyleEnum) As String
    Select Case penTreeType
        Case tseClass: GetTreeStyleName = "Class"
        Case tseRace: GetTreeStyleName = "Race"
        Case tseDestiny: GetTreeStyleName = "Destiny"
        Case tseGlobal: GetTreeStyleName = "Global"
        Case tseRaceClass: GetTreeStyleName = "RaceClass"
    End Select
End Function

Public Function GetTreeStyleID(pstrStyle As String) As TreeStyleEnum
    Select Case LCase$(pstrStyle)
        Case "class": GetTreeStyleID = tseClass
        Case "race": GetTreeStyleID = tseRace
        Case "destiny": GetTreeStyleID = tseDestiny
        Case "global": GetTreeStyleID = tseGlobal
        Case "raceclass": GetTreeStyleID = tseRaceClass
    End Select
End Function

Public Function GetFeatChannelName(penChannel As FeatChannelEnum, Optional ByVal penRace As RaceEnum = reAny) As String
    Select Case penChannel
        Case fceSelected: GetFeatChannelName = "Selected"
        Case fceGeneral: GetFeatChannelName = "General"
        Case fceRacial
            Select Case penRace
                Case reHalfElf: GetFeatChannelName = "Dilettante"
                Case reDragonborn: GetFeatChannelName = "Dragon"
                Case reAasimar: GetFeatChannelName = "Bond"
                Case Else: GetFeatChannelName = "Racial"
            End Select
        Case fceWarlock: GetFeatChannelName = "Pact"
        Case fceMonk: GetFeatChannelName = "Monk"
        Case fceRogue: GetFeatChannelName = "Rogue"
        Case fceFavoredEnemy: GetFeatChannelName = "Ranger"
        Case fceWildShape: GetFeatChannelName = "Druid"
        Case fceCleric: GetFeatChannelName = "Domain"
        Case fceFavoredSoul: GetFeatChannelName = "Fav Soul"
        Case fceEnergy: GetFeatChannelName = "Energy"
        Case fceDeity: GetFeatChannelName = "Deity"
        Case fceGranted: GetFeatChannelName = "Granted"
    End Select
End Function

Public Function GetFeatChannelID(pstrChannel As String) As FeatChannelEnum
    Select Case pstrChannel
        Case "Selected": GetFeatChannelID = fceSelected
        Case "General": GetFeatChannelID = fceGeneral
        Case "Racial", "Dragon", "Dilettante", "Bond": GetFeatChannelID = fceRacial
        Case "Pact": GetFeatChannelID = fceWarlock
        Case "Monk": GetFeatChannelID = fceMonk
        Case "Rogue": GetFeatChannelID = fceRogue
        Case "Ranger": GetFeatChannelID = fceFavoredEnemy
        Case "Druid": GetFeatChannelID = fceWildShape
        Case "Domain": GetFeatChannelID = fceCleric
        Case "Fav Soul": GetFeatChannelID = fceFavoredSoul
        Case "Energy": GetFeatChannelID = fceEnergy
        Case "Deity": GetFeatChannelID = fceDeity
        Case "Granted": GetFeatChannelID = fceGranted
        Case Else: GetFeatChannelID = fceUnknown
    End Select
End Function

Public Function GetRaceReqID(pstrRaceReq As String) As RaceReqEnum
    Select Case LCase$(pstrRaceReq)
        Case "required": GetRaceReqID = rreRequired
        Case "notallowed": GetRaceReqID = rreNotAllowed
        Case "standard": GetRaceReqID = rreStandard
        Case "iconic": GetRaceReqID = rreIconic
        Case Else: GetRaceReqID = rreAny
    End Select
End Function
