Attribute VB_Name = "basDataDefs"
Option Explicit

Public db As DatabaseType

Public gtypMenu As MenuType ' Used to pass menu structure to / from modal form
Public gtypClipboard As ClipboardCommandType ' Clipboard for menu items

Public gblnDirtyFlag(4) As Boolean


' ************* ENUMERATIONS *************


Public Enum DirtyFlagEnum
    dfeAny
    dfeSettings
    dfeData
    dfeNotes
    dfeLinks
End Enum

Public Enum GeneratedColorEnum
    gceUnknown
    gceRed
    gceOrange
    gceYellow
    gceChartreuse
    gceGreen
    gceAqua
    gceTeal
    gceSky
    gceBlue
    gceOrchid
    gcePurple
    gcePink
    gceGray
    gceColors
End Enum

Public Enum QuestStyleEnum
    qeUnknown
    qeRaid
    qeSolo
    qeQuest
End Enum

Public Enum QuestStatusEnum
    qseShow
    qseDim
    qseHide
End Enum

Public Enum MenuCommandEnum
    mceLink
    mceShortcut
    mceSeparator
End Enum

Public Enum ClipboardStatusEnum
    cseEmpty
    cseCopied
    cseCut
End Enum

Public Enum TableColumnStyleEnum
    tcseNumeric
    tcseTextLeft
    tcseTextCenter
    tcseTextRight
End Enum

Public Enum NoteScopeEnum
    nsePublic
    nseShared
    nsePrivate
End Enum


' ************* USER-DEFINED TYPES *************


Public Type NotesType
    Text As String
    Scope As NoteScopeEnum
    TabName As String
    MenuName As String
    FontName As String
    FontSize As String
End Type

Public Type QuestLinkType
    FullName As String
    Abbreviation As String
    Style As MenuCommandEnum
    Target As String
End Type

Public Type QuestType
    Quest As String
    ID As String ' Original quest name text so that old compendium files still work if you correct typo in name
    SortName As String
    CompendiumName As String
    Wiki As String
    Order As Long
    Style As QuestStyleEnum
    Patron As String
    Favor As Long
    Pack As String
    Skipped As Boolean
    BaseLevel As Long
    EpicLevel As Long
    GroupLevel As Long
    SagaGroup(1) As Long ' 0=Heroic, 1=Epic
    SagaOrder(1) As Long ' 0=Heroic, 1=Epic
    Saga() As Long
    Progress() As ProgressEnum
    Status As QuestStatusEnum
    Link() As QuestLinkType
    Links As Long
End Type

Public Type RewardType
    Points As Long
    xp As Long
    Renown As Long
End Type

Public Type SagaType
    SagaName As String
    Abbreviation As String
    Wiki As String
    Tier As SagaTierEnum
    Order As Long
    Group As Long
    NPC() As String
    NPCs As Long
    Tome As Long
    Astrals As Long
    Reward(3) As RewardType
    Quests As Long
    Quest() As Long
End Type

Public Type PackType
    Pack As String
    Abbreviation As String
    Wiki As String
    Lowest As Long
    Highest As Long
    Cost As Long
    Patron As Long
    Order As Long
    Link() As QuestLinkType
    Links As Long
End Type

Public Type AreaType
    Area As String
    Wiki As String
    Map As String
    Lowest As Long
    Highest As Long
    Explorer As Long
    Pack As String
    Order As Long
    Links As Long
    Link() As QuestLinkType
End Type

Public Type ChallengeType
    Challenge As String
    ID As String ' Original challenge name text so that old compendium files still work if you correct typo in name
    Wiki As String
    Group As String
    Patron As String
    LevelLow As Long
    LevelHigh As Long
    GameOrder As Long
    GroupOrder As Long
    Stars() As Long
    MaxStars As Long
End Type

Public Type PatronType
    Patron As String
    Abbreviation As String
    Wiki As String
    Order As Long
End Type

Public Type MenuCommandType
    Caption As String
    Style As MenuCommandEnum
    Target As String
    Param As String
End Type

Public Type ClipboardCommandType
    Command As MenuCommandType
    ClipboardStatus As ClipboardStatusEnum
End Type

Public Type MenuType
    LinkList As Boolean
    Title As String
    Left As Long
    Top As Long
    Commands As Long
    Command() As MenuCommandType
    Accepted As Boolean
    Deleted As Boolean
    Selected As Long ' Index of item to be selected, -1 for title, or 0 for none
End Type

Public Type CharacterSagaType
    Progress() As ProgressEnum
End Type

Public Type TomeType
    Stat(1 To 6) As Long
    Skill(1 To 21) As Long
    RacialAP As Long
    UniversalAP As Long
    HeroicXP As String
    EpicXP As String
    Fate As Long
    Destiny As Long
    Power(1 To 3) As Long
    RR(1 To 2) As Long
End Type

'TODO Remove hard coding
Public Const maxClassCount = 22  '2024.03.17 DragonLord => 15 classes + 7 Archtypes
Public Const maxRaceCount = 14 ' 2022.04.29 Tabaxi
Public Const maxIconicCount = 9 ' 2022.04.29 Tabaxi Trailblazer
Public Const maxEpicCount = 17 ' 2022.04.29 Added Eclispe Power

Public Type PastLifeType
    Class(1 To maxClassCount) As Long
    Racial(1 To maxRaceCount) As Long
    Iconic(1 To maxIconicCount) As Long
    Epic(1 To maxEpicCount) As Long
End Type

Public Type CharacterType
    Character As String
    Level As Long
    Notes As String
    ContextMenu As MenuType
    LeftClick As String ' Menu item caption
    CustomColor As Boolean
    GeneratedColor As GeneratedColorEnum
    BackColor As Long
    DimColor As Long
    QuestFavor As Long
    ChallengeFavor As Long
    TotalFavor As Long
    Tome As TomeType
    PastLife As PastLifeType
    Saga() As CharacterSagaType
'    Notes As Long
'    Note() As NotesType
End Type

Public Type CoordsType
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Public Type TableColumnType
    Header As String
    Widest As String
    Style As TableColumnStyleEnum
    Left As Long
    Width As Long
    Right As Long
End Type

Public Type TableRowType
    Value() As Variant
    Group As Long
End Type

Public Type TableType
    TableID As String
    Title As String
    Headers As Boolean
    Group As Long
    Columns As Long
    Column() As TableColumnType
    Rows As Long
    Row() As TableRowType
End Type

Public Type DatabaseType
    Packs As Long
    Pack() As PackType
    Quests As Long
    Quest() As QuestType
    Sagas As Long
    Saga() As SagaType
    Areas As Long
    Area() As AreaType
    Challenge() As ChallengeType
    Challenges As Long
    Patrons As Long
    Patron() As PatronType
    Characters As Long
    Character() As CharacterType
    LinkLists As Long
    LinkList() As MenuType
    Templates As Long
    Template() As MenuType
    Tables As Long
    Table() As TableType
    FileNotes As Long
    FileNote() As NotesType
    PublicNotes As Long
    PublicNote() As NotesType
End Type


' ************* ENUMERATED VALUES *************


Public Function GetChallengeOrderName(penOrder As ChallengeOrderEnum) As String
    Select Case penOrder
        Case coeChallenge: GetChallengeOrderName = "Challenge"
        Case coeGame: GetChallengeOrderName = "Game"
        Case coeGroup: GetChallengeOrderName = "Group"
    End Select
End Function

Public Function GetChallengeOrderID(pstrOrder As String) As ChallengeOrderEnum
    Select Case LCase$(pstrOrder)
        Case "challenge": GetChallengeOrderID = coeChallenge
        Case "game": GetChallengeOrderID = coeGame
        Case "group": GetChallengeOrderID = coeGroup
    End Select
End Function

Public Function GetColorName(penColor As GeneratedColorEnum) As String
    Select Case penColor
        Case gceAqua: GetColorName = "Aqua"
        Case gceBlue: GetColorName = "Blue"
        Case gceChartreuse: GetColorName = "Chartreuse"
        Case gceGray: GetColorName = "Gray"
        Case gceGreen: GetColorName = "Green"
        Case gceOrange: GetColorName = "Orange"
        Case gceOrchid: GetColorName = "Orchid"
        Case gcePink: GetColorName = "Pink"
        Case gcePurple: GetColorName = "Purple"
        Case gceRed: GetColorName = "Red"
        Case gceSky: GetColorName = "Sky"
        Case gceTeal: GetColorName = "Teal"
        Case gceYellow: GetColorName = "Yellow"
    End Select
End Function

Public Function GetColorID(pstrColor As String) As GeneratedColorEnum
    Select Case LCase$(pstrColor)
        Case "aqua": GetColorID = gceAqua
        Case "blue": GetColorID = gceBlue
        Case "chartreuse": GetColorID = gceChartreuse
        Case "gray": GetColorID = gceGray
        Case "green": GetColorID = gceGreen
        Case "orange": GetColorID = gceOrange
        Case "orchid": GetColorID = gceOrchid
        Case "pink": GetColorID = gcePink
        Case "purple": GetColorID = gcePurple
        Case "red": GetColorID = gceRed
        Case "sky": GetColorID = gceSky
        Case "teal": GetColorID = gceTeal
        Case "yellow": GetColorID = gceYellow
    End Select
End Function

Public Function GetColorValue(penColor As GeneratedColorEnum, Optional pblnDim As Boolean = False) As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim lngHigh As Long
    Dim lngMed As Long
    Dim lngLow As Long
    Dim dblSkip As Double
    
    lngHigh = cfg.NamedHigh
    lngMed = cfg.NamedMed
    lngLow = cfg.NamedLow
    Select Case penColor
        Case gceAqua:         SetColorValue lngRed, lngGreen, lngBlue, lngLow, lngHigh, lngMed
        Case gceBlue:          SetColorValue lngRed, lngGreen, lngBlue, lngLow, lngLow, lngHigh
        Case gceChartreuse: SetColorValue lngRed, lngGreen, lngBlue, lngMed, lngHigh, lngLow
        Case gceGray:          SetColorValue lngRed, lngGreen, lngBlue, lngHigh, lngHigh, lngHigh
        Case gceGreen:        SetColorValue lngRed, lngGreen, lngBlue, lngLow, lngHigh, lngLow
        Case gceOrange:      SetColorValue lngRed, lngGreen, lngBlue, lngHigh, lngMed, lngLow
        Case gceOrchid:       SetColorValue lngRed, lngGreen, lngBlue, lngMed, lngLow, lngHigh
        Case gcePink:          SetColorValue lngRed, lngGreen, lngBlue, lngHigh, lngLow, lngMed
        Case gcePurple:       SetColorValue lngRed, lngGreen, lngBlue, lngHigh, lngLow, lngHigh
        Case gceRed:           SetColorValue lngRed, lngGreen, lngBlue, lngHigh, lngLow, lngLow
        Case gceSky:           SetColorValue lngRed, lngGreen, lngBlue, lngLow, lngMed, lngHigh
        Case gceTeal:           SetColorValue lngRed, lngGreen, lngBlue, lngLow, lngHigh, lngHigh
        Case gceYellow:       SetColorValue lngRed, lngGreen, lngBlue, lngHigh, lngHigh, lngLow
        Case Else: xp.ColorToRGB cfg.GetColor(cgeControls, cveBackground), lngRed, lngGreen, lngBlue
    End Select
    If pblnDim Then
        dblSkip = cfg.NamedDim / 100
        lngRed = lngRed * dblSkip
        lngGreen = lngGreen * dblSkip
        lngBlue = lngBlue * dblSkip
    End If
    GetColorValue = RGB(lngRed, lngGreen, lngBlue)
End Function

Private Sub SetColorValue(plngRed As Long, plngGreen As Long, plngBlue As Long, plngSetRed As Long, plngSetGreen As Long, plngSetBlue As Long)
    plngRed = plngSetRed
    plngGreen = plngSetGreen
    plngBlue = plngSetBlue
End Sub

Public Function GetColorDim(plngColor As Long) As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim dblSkip As Double
    
    xp.ColorToRGB plngColor, lngRed, lngGreen, lngBlue
    dblSkip = cfg.NamedDim / 100
    lngRed = lngRed * dblSkip
    lngGreen = lngGreen * dblSkip
    lngBlue = lngBlue * dblSkip
    GetColorDim = RGB(lngRed, lngGreen, lngBlue)
End Function

Public Function GetMenuStyleName(penStyle As MenuCommandEnum) As String
    Select Case penStyle
        Case mceLink: GetMenuStyleName = "Link"
        Case mceShortcut: GetMenuStyleName = "Shortcut"
        Case mceSeparator: GetMenuStyleName = "Separator"
    End Select
End Function

Public Function GetMenuStyleID(pstrStyle As String) As MenuCommandEnum
    Select Case LCase$(pstrStyle)
        Case "link": GetMenuStyleID = mceLink
        Case "shortcut": GetMenuStyleID = mceShortcut
        Case "separator": GetMenuStyleID = mceSeparator
    End Select
End Function

Public Function GetOrderID(pstrOrder As String) As CompendiumOrderEnum
    Select Case LCase$(pstrOrder)
        Case "level": GetOrderID = coeLevel
        Case "epic": GetOrderID = coeEpic
        Case "quest": GetOrderID = coeQuest
        Case "pack": GetOrderID = coePack
        Case "patron": GetOrderID = coePatron
        Case "favor": GetOrderID = coeFavor
        Case "style": GetOrderID = coeStyle
        Case Else: GetOrderID = coePack
    End Select
End Function

Public Function GetOrderName(penOrder As CompendiumOrderEnum) As String
    Select Case penOrder
        Case coeLevel: GetOrderName = "Level"
        Case coeEpic: GetOrderName = "Epic"
        Case coeQuest: GetOrderName = "Quest"
        Case coePack: GetOrderName = "Pack"
        Case coePatron: GetOrderName = "Patron"
        Case coeFavor: GetOrderName = "Favor"
        Case coeStyle: GetOrderName = "Style"
    End Select
End Function

Public Function GetPaneName(penPane As PaneEnum) As String
    Select Case penPane
        Case peQuests: GetPaneName = "Quests"
        Case peHome: GetPaneName = "Home"
        Case peXP: GetPaneName = "XP"
        Case peWilderness: GetPaneName = "Wilderness"
        Case peNotes: GetPaneName = "Notes"
        Case peLinks: GetPaneName = "Links"
    End Select
End Function

Public Function GetPaneID(pstrPane As String) As PaneEnum
    Select Case LCase$(pstrPane)
        Case "quests": GetPaneID = peQuests
        Case "home": GetPaneID = peHome
        Case "xp": GetPaneID = peXP
        Case "wilderness": GetPaneID = peWilderness
        Case "notes": GetPaneID = peNotes
        Case "links": GetPaneID = peLinks
    End Select
End Function

Public Function GetProgressID(pstrProgress As String) As ProgressEnum
    If Len(pstrProgress) = 1 Then
        GetProgressID = InStr("scnheav", pstrProgress)
    Else
        Select Case pstrProgress
            Case "Casual": GetProgressID = peCasual
            Case "Normal": GetProgressID = peNormal
            Case "Hard": GetProgressID = peHard
            Case "Elite": GetProgressID = peElite
            Case "VIP", "VIP Skip": GetProgressID = peVIP
            Case "Astrals", "Astral Shard Skip": GetProgressID = peAstrals
        End Select
    End If
End Function

Public Function GetProgressLetter(penProgress As ProgressEnum) As String
    Dim strProgress As String
    
    If penProgress <> peNone Then strProgress = Mid$("scnheav", penProgress, 1)
    If Len(strProgress) = 0 Then strProgress = " "
    GetProgressLetter = strProgress
End Function

Public Function GetProgressName(penProgress As ProgressEnum) As String
    Select Case penProgress
        Case peSolo: GetProgressName = "Solo"
        Case peCasual: GetProgressName = "Casual"
        Case peNormal: GetProgressName = "Normal"
        Case peHard: GetProgressName = "Hard"
        Case peElite: GetProgressName = "Elite"
        Case peAstrals: GetProgressName = "Astrals"
        Case peVIP: GetProgressName = "VIP"
    End Select
End Function

Public Function GetQuestStyleID(pstrStyle As String) As QuestStyleEnum
    Select Case LCase$(pstrStyle)
        Case "raid": GetQuestStyleID = qeRaid
        Case "solo": GetQuestStyleID = qeSolo
        Case Else: GetQuestStyleID = qeQuest
    End Select
End Function

Public Function GetQuestStyleName(penStyle As QuestStyleEnum)
    Select Case penStyle
        Case qeRaid: GetQuestStyleName = "Raid"
        Case qeSolo: GetQuestStyleName = "Solo"
    End Select
End Function

Public Function GetWindowSizeName(penSize As WindowSizeEnum) As String
    Select Case penSize
        Case wseMaximized: GetWindowSizeName = "Maximized"
        Case wseFillDesktop: GetWindowSizeName = "Fill Desktop"
        Case wseRemember: GetWindowSizeName = "Last Position"
    End Select
End Function

Public Function GetWindowSizeID(pstrSize As String) As WindowSizeEnum
    Select Case LCase$(pstrSize)
        Case "maximized": GetWindowSizeID = wseMaximized
        Case "fill desktop": GetWindowSizeID = wseFillDesktop
        Case "last position": GetWindowSizeID = wseRemember
    End Select
End Function
