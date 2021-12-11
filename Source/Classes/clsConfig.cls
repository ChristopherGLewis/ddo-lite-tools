VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private mstrMessages As String


' ************* GENERAL *************


Public Enum ExtensionEnum
    eeUnknown
    eeScreenColor
    eeOutputColor
    eeBBCodes
End Enum

Public Enum SectionEnum
    seAll
    seShared
    seColors
    seCharBuild
    seOutput
    seCompendium
    sePaths
    seLinkLists
End Enum

Public Enum UtilEnum
    ueColors
    ueLightsOut
    ueADQ
    ueStopwatch
End Enum

Private Type CoordsType
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type MainWindowType
    WindowSize As WindowSizeEnum
    Maximized As Boolean
    WindowCoords As CoordsType
End Type

Private mstrArray() As String
Private mlngLine As Long

Private mblnDirty As Boolean


' ************* PATH DEFS *************


Private Type PathType
    CharacterBuilderLite As String
    RonPath As String
    BuilderPath As String
    Compendium As String
    CannithCrafting As String
End Type

Private mtypPath As PathType
Private mblnOldPaths As Boolean ' If old-style CharBuild paths found on load, immediately save to new version


' ************* SHARED DEFS *************


Private Type SharedType
    ColorCoords As CoordsType
    ScreenColors As String
    OutputColors As String
    ColorVersion As Long
End Type

Private mtypShared As SharedType

' Increment this value to have user colors reset to defaults the next time they run any Lite tool
Private Const ColorVersion As Long = 1


' ************* COLOR DEFS *************


Public Enum ColorFileEnum
    cfeAll
    cfeScreen
    cfeOutput
End Enum

Public Enum ColorGroupEnum
    cgeWorkspace
    cgeControls
    cgeNavigation
    cgeDropSlots
    cgeOutput
End Enum

Public Enum ColorValueEnum
    cveBackground
    cveBackHighlight
    cveBackError
    cveBackRelated
    cveText
    cveTextError
    cveTextDim
    cveTextLink
    cveBorderInterior
    cveBorderExterior
    cveBorderHighlight
    cveLightGray
    cveRed
    cveYellow
    cveBlue
    cvePurple
    cveGreen
    cveOrange
    cveColorValues
End Enum

Private mlngColor(4, 17) As Long
Private mblnDarkColors As Boolean


' ************* CHARBUILD DEFS *************


Private Const StartMRU As Long = 11 ' Index of the "-" menu option

Public Enum BackColorEnum
    bceDefault
    bceOfficial
    bceCodeTags
    bceStylish
    bceCustom
End Enum

Public Enum BuildPointsEnum
    beAdventurer
    beChampion
    beHero
    beLegend
End Enum

Public Enum ContrastEnum
    ceNone
    ceLow
    ceMedium
    ceHigh
End Enum

Public Enum FeatOrderEnum
    foeAlphabetical
    foeGroupRelated
End Enum

Public Enum FeatOrderOutputEnum
    fooLevel
    fooChannel
End Enum

Public Enum OutputEnum
    oeAll
    oeOverview
    oeStats
    oeSkills
    oeFeats
    oeSpells
    oeEnhancements
    oeDestiny
    oeDestiny2
    oeGear
    oeRemember
    oeExport
    oeNone
End Enum

Public Enum SkillOrderOutputEnum
    sooAlphabetical
    sooRanksSpent
End Enum

Public Enum SkillOrderScreenEnum
    sosAlphabetical
    sosNativeFirst
End Enum

Public Enum WindowSizeEnum
    wseMaximized
    wseFillDesktop
    wseRemember
End Enum

Private Type CharBuildType
    MainWindow As MainWindowType
    DialogCoords As CoordsType
    OptionsCoords As CoordsType
    ConvertCoords As CoordsType
    ConvertOnStartup As Boolean
    ChildWindows As Boolean
    ShowErrors As Boolean
    MRUList() As String
    MRUCount As Long
    BuildPoints As BuildPointsEnum
    SkillOrderScreen As SkillOrderScreenEnum
    SkillOrderOutput As SkillOrderOutputEnum
    FeatOrder As FeatOrderEnum
    FeatChannels As Boolean
    FeatOrderOutput As FeatOrderOutputEnum
    FeatShowAll As Boolean '2021.08.06 added FeatShowAll flag
    IconOverview As Boolean
    IconSkills As Boolean
    Version As String
    ProcessTime As Long
    UseIcons As Boolean
    BBCodes As String
    OutputMargin As Long
    Confirm As Boolean
    Contrast As ContrastEnum
End Type

Private mtypCharBuild As CharBuildType
Private mstrForm(1 To 20) As String
Private mlngForms As Long
Private mblnRefresh As Boolean ' If changes were made while frmMain was minimized, refresh output when frmMain is restored


' ************* COMPENDIUM DEFS *************


Public Enum CompendiumOrderEnum
    coeLevel
    coeEpic
    coeQuest
    coePack
    coePatron
    coeFavor
    coeStyle
End Enum

Public Enum LevelSortEnum
    lseHeroic
    lseEpic
    lseGame
End Enum

Public Enum PaneEnum
    peQuests
    peHome
    peXP
    peWilderness
    peNotes
    peLinks
    pePaneCount
End Enum

Public Enum ProgressEnum
    peNone
    peSolo
    peCasual
    peNormal
    peHard
    peElite
    peAstrals
    peVIP
End Enum

Public Enum ChallengeOrderEnum
    coeChallenge
    coeGame
    coeGroup
End Enum

Public Enum SagaTierEnum
    steHeroic
    steEpic
End Enum

Public Enum SagaSortEnum
    ssePack ' Default
    sseLevel
    sseQuest
    sseUser
End Enum

Private Type CompendiumCoordsType
    Challenges As CoordsType
    Sagas As CoordsType
    SagaDetail As CoordsType
    Tools As CoordsType
    Characters As CoordsType
    Menus As CoordsType
    Patrons As CoordsType
    Colors As CoordsType
    CritCalculator As CoordsType
    Wilderness As CoordsType
End Type

Private Type CompendiumType
    DataFile As String
    PlayButton As Boolean
    PlayEXE As String
    RunBefore As Boolean
    WheelStep As Long
    MainWindow As MainWindowType
    ChildWindows As Boolean
    Coords As CompendiumCoordsType
    CompendiumOrder As CompendiumOrderEnum
    LevelSort As LevelSortEnum
    AbbreviateColumns As Boolean
    AbbreviatePacks As Boolean
    AbbreviatePatrons As Boolean
    LeftPane As PaneEnum
    RightPane As PaneEnum
    Character As Long
    BackColor As Long
    NamedHigh As Long
    NamedMed As Long
    NamedLow As Long
    NamedDim As Long
    Scroll As Long
    MarginX As Long
    MarginY As Long
    Sides As Long
    Bottom As Long
    Difficulty As ProgressEnum
    ChallengeOrder As ChallengeOrderEnum
    SagaTier As SagaTierEnum
    SagaSort As SagaSortEnum
    SagaBringToTops As Long
    SagaBringToTop() As String
End Type

Private mtypCompendium As CompendiumType


' ************* OUTPUT DEFS *************


Private Type TagsType
    Enabled As Boolean
    OpenTag As String
    CloseTag As String
End Type

Private Type ListTagsType
    Enabled As Boolean
    BulletOpen As String
    BulletClose As String
    NumberedOpen As String
    NumberedClose As String
End Type

Private Type OutputType
    Codes As Boolean
    Dots As Boolean
    TextColors As Boolean
    Bold As TagsType
    Underline As TagsType
    Fixed As TagsType
    Lists As ListTagsType
    Color As TagsType
    Wrapper As TagsType
    Reddit As Boolean
End Type

Private mtypOutput As OutputType
Private menOutputSection As OutputEnum


' ************* INITIALIZE *************


Private Sub Class_Initialize()
    mblnRefresh = False
    SetDefaults seAll
    LoadSettings
    If mblnOldPaths Then SaveSettings False
End Sub

Private Sub Class_Terminate()
    SaveSettings
End Sub

Private Function SettingsFile() As String
    SettingsFile = AppRootPath() & "\Settings.txt"
End Function

Public Property Get SharedData() As Long
    SharedData = AppRootPath() & "\Data\"
End Property

Private Function DefaultPath() As String
    DefaultPath = AppRootPath() & "\Save"
End Function

Public Function AppRootPath() As String
    Dim strPath As String
    
    If App.Title = "Colors" Then
        If xp.DebugMode Then
            strPath = App.Path & "\..\..\.."
        Else
            strPath = App.Path & "\.."
        End If
        AppRootPath = xp.Folder.RelativeToAbsolute(strPath)
    Else
        AppRootPath = App.Path
    End If
End Function

Public Sub SetDefaults(penSection As SectionEnum)
    Dim typBlank As OutputType
    
    If penSection = seCharBuild Or penSection = seAll Then
        With mtypCharBuild
            With .MainWindow
                .WindowSize = wseRemember
                .Maximized = True
            End With
            .ConvertOnStartup = False
            .ChildWindows = True
            .ShowErrors = True
            .MRUCount = 9
            ReDim .MRUList(1 To .MRUCount)
            .BuildPoints = beAdventurer
            .SkillOrderScreen = sosAlphabetical
            .SkillOrderOutput = sooRanksSpent
            .FeatOrder = foeAlphabetical
            .FeatChannels = False
            .FeatOrderOutput = fooChannel
            .FeatShowAll = False '2021.08.06 added FeatShowAll flag
            .IconOverview = True
            .IconSkills = True
            .UseIcons = True
            .BBCodes = "Default"
            .OutputMargin = 0
            .Confirm = True
            .Version = vbNullString
            .Contrast = ceMedium
        End With
    End If
    If penSection = seCompendium Or penSection = seAll Then
        With mtypCompendium
            .DataFile = "Main"
            With .MainWindow
                .WindowSize = wseRemember
                .Maximized = True
            End With
            .Character = 0
            .ChildWindows = True
            .PlayButton = True
            .PlayEXE = vbNullString
            .WheelStep = 6
            .RunBefore = False
            .SagaTier = steEpic
            .SagaSort = ssePack
            .SagaBringToTops = 0
            Erase .SagaBringToTop
            .Difficulty = peElite
            .LevelSort = lseHeroic
            .AbbreviateColumns = False
            .AbbreviatePacks = True
            .AbbreviatePatrons = True
            .LeftPane = peQuests
            .RightPane = peHome
            .Character = 0
            .CompendiumOrder = coePack
            .ChallengeOrder = coeGroup
            .MarginX = 4
            .MarginY = 2
            .Sides = 0
            .Bottom = 0
        End With
    End If
    If penSection = sePaths Or penSection = seAll Then
        mtypPath.CharacterBuilderLite = vbNullString
        mtypPath.Compendium = vbNullString
        mtypPath.CannithCrafting = vbNullString
        mtypPath.RonPath = vbNullString
        mtypPath.BuilderPath = vbNullString
    End If
    If penSection = seShared Or penSection = seAll Then
        With mtypShared
            .ScreenColors = "Black"
            .OutputColors = "Default"
        End With
    End If
    If penSection = seOutput Or penSection = seAll Then
        mtypOutput = typBlank
        InitFormat "Default", mtypOutput
    End If
    If penSection = seColors Or penSection = seAll Then
        Erase mlngColor
        InitColors cveBackground, 0, 3947580, 6316128, 6316128, 2236962
        InitColors cveBackHighlight, 19532, 32125, 24672, 32125, 19275
        InitColors cveBackError, 76, 4934575, 96, 4934575, 75
        InitColors cveBackRelated, 4980736, 11488075, 6291456, 11488075, 4915200
        InitColors cveText, 12632256, 15461355, 14803425, 16777215, 16777215
        InitColors cveTextError, 10132207, 10132207, 12632319, 12632319, 11579647
        InitColors cveTextDim, 8553090, 11579568, 12895428, 11579568, 12632256
        InitColors cveTextLink, 16769480, 16769480, 16777215, 16762026, 16769480
        InitColors cveBorderInterior, 8421504, 7039851, 3947580, 2171169, 8421504
        InitColors cveBorderExterior, 12632256, 11579568, 3947580, 11513775, 12632256
        InitColors cveBorderHighlight, 10526880, 16777215, 3947580, 16777215, 16777215
        InitColors cveLightGray, 10395294, 12434877, 10395294, 10395294, 12434877
        InitColors cveRed, 10132207, 10132207, 10132207, 10132207, 10132207
        InitColors cveYellow, 7795199, 7795199, 3927039, 3927039, 7795199
        InitColors cveBlue, 16370320, 16370320, 16370320, 16370320, 16370320
        InitColors cvePurple, 14193614, 14193614, 14193614, 14193614, 14193614
        InitColors cveGreen, 8701825, 8701825, 10999461, 10999461, 8701825
        InitColors cveOrange, 2533375, 5093375, 2533375, 2533375, 5093375
    End If
End Sub

Private Sub InitColors(penColor As ColorValueEnum, plngWrk As Long, plngCtl As Long, plngNav As Long, plngDrp As Long, plngOut As Long)
    mlngColor(cgeWorkspace, penColor) = plngWrk
    mlngColor(cgeControls, penColor) = plngCtl
    mlngColor(cgeNavigation, penColor) = plngNav
    mlngColor(cgeDropSlots, penColor) = plngDrp
    mlngColor(cgeOutput, penColor) = plngOut
End Sub


' ************* LOAD SETTINGS *************


Private Sub LoadSettings()
    Dim strFile As String
    Dim strRaw As String
    Dim strSection() As String
    Dim strTitle As String
    Dim i As Long
    
    mblnOldPaths = False
    strFile = SettingsFile()
    If Not xp.File.Exists(strFile) Then
        mtypShared.ColorVersion = ColorVersion
        Exit Sub
    End If
    strRaw = xp.File.LoadToString(strFile)
    strSection = Split(strRaw, "Section: ")
    For i = 1 To UBound(strSection)
        strTitle = Left$(strSection(i), InStr(strSection(i), vbNewLine) - 1)
        Select Case strTitle
            Case "Character Builder Lite": LoadCharBuild strSection(i)
            Case "Compendium": LoadCompendium strSection(i)
            Case "Paths": LoadPaths strSection(i)
            Case "Shared": LoadShared strSection(i)
            Case "Output": LoadOutput strSection(i)
            Case "Colors": LoadColors strSection(i), cfeAll
        End Select
    Next
End Sub

Public Sub ReQuery(penSection As SectionEnum)
    Dim strFile As String
    Dim strRaw As String
    Dim strSection() As String
    Dim strTitle As String
    Dim i As Long
    
    strFile = SettingsFile()
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strSection = Split(strRaw, "Section: ")
    For i = 1 To UBound(strSection)
        strTitle = Left$(strSection(i), InStr(strSection(i), vbNewLine) - 1)
        Select Case strTitle
            Case "Colors"
                If penSection = seColors Or penSection = seAll Then
                    LoadColors strSection(i), cfeAll
                    RefreshAllColors
                End If
            Case "Shared"
                If penSection = seColors Or penSection = seShared Or penSection = seAll Then
                    LoadShared strSection(i)
                End If
            Case "Paths"
                If penSection = sePaths Or penSection = seAll Then
                    LoadPaths strSection(i)
                End If
        End Select
    Next
End Sub

Private Sub LoadCharBuild(pstrRaw As String)
    Dim strField As String
    Dim strValue As String
    Dim strLine() As String
    Dim typBlank As CharBuildType
    Dim i As Long
    
    mtypCharBuild = typBlank
    SetDefaults seCharBuild
    With mtypCharBuild
        strLine = Split(pstrRaw, vbNewLine)
        For i = 1 To UBound(strLine)
            If ParseLine(strLine(i), strField, strValue) Then
                Select Case strField
                    Case "windowsize": .MainWindow.WindowSize = val(strValue)
                    Case "maximized": If LCase$(strValue) = "true" Then .MainWindow.Maximized = True Else .MainWindow.Maximized = False
                    Case "windowcoords": ParseCoords .MainWindow.WindowCoords, strValue
                    Case "dialogcoords": ParseCoords .DialogCoords, strValue
                    Case "optionscoords": ParseCoords .OptionsCoords, strValue
                    Case "convertcoords": ParseCoords .ConvertCoords, strValue
                    Case "convertonstartup": If LCase$(strValue) = "true" Then .ConvertOnStartup = True Else .ConvertOnStartup = False
                    Case "childwindows": If LCase$(strValue) = "true" Then .ChildWindows = True Else .ChildWindows = False
                    Case "showerrorsonstartup": If LCase$(strValue) = "true" Then .ShowErrors = True Else .ShowErrors = False
                    Case "iconoverview": If LCase$(strValue) = "true" Then .IconOverview = True Else .IconOverview = False
                    Case "iconskills": If LCase$(strValue) = "true" Then .IconSkills = True Else .IconSkills = False
                    Case "mrulist": SplitMRU strValue
                    Case "mrucount": .MRUCount = val(strValue)
                    Case "buildpoints": .BuildPoints = val(strValue)
                    Case "skillorderscreen": .SkillOrderScreen = val(strValue)
                    Case "skillorderoutput": .SkillOrderOutput = val(strValue)
                    Case "featorder": .FeatOrder = val(strValue)
                    Case "featchannels": If LCase$(strValue) = "true" Then .FeatChannels = True Else .FeatChannels = False
                    Case "featorderoutput": .FeatOrderOutput = val(strValue)
                    Case "featshowall": If LCase$(strValue) = "true" Then .FeatShowAll = True Else .FeatShowAll = False '2021.08.06 added FeatShowAll flag
                    Case "useicons": If LCase$(strValue) = "true" Then .UseIcons = True Else .UseIcons = False
                    Case "bbcodes": .BBCodes = strValue
                    Case "outputmargin": .OutputMargin = val(strValue)
                    Case "confirm": If LCase$(strValue) = "true" Then .Confirm = True Else .Confirm = False
                    Case "version": .Version = strValue
                    Case "processtime": .ProcessTime = val(strValue)
                    Case "contrast": .Contrast = val(strValue)
                    Case "litepath"
                        mtypPath.CharacterBuilderLite = strValue
                        mblnOldPaths = True
                    Case "ronpath"
                        mtypPath.RonPath = strValue
                        mblnOldPaths = True
                    Case "builderpath"
                        mtypPath.BuilderPath = strValue
                        mblnOldPaths = True
                End Select
            End If
        Next
        If .MRUCount = 0 Then Erase .MRUList Else ReDim Preserve .MRUList(1 To .MRUCount)
    End With
End Sub

Private Sub SplitMRU(pstrRaw As String)
    Dim strList() As String
    Dim i As Long
    Dim iMax As Long
    
    With mtypCharBuild
        If .MRUCount = 0 Then
            Erase .MRUList
        Else
            ReDim Preserve .MRUList(1 To .MRUCount)
            If Len(pstrRaw) <> 0 Then
                strList = Split(vbTab & pstrRaw, vbTab)
                iMax = UBound(strList)
                If iMax > .MRUCount Then iMax = .MRUCount
                For i = 1 To iMax
                    mtypCharBuild.MRUList(i) = strList(i)
                Next
            End If
        End If
    End With
End Sub

Private Sub LoadCompendium(pstrRaw As String)
    Dim strField As String
    Dim strValue As String
    Dim strLine() As String
    Dim typBlank As CompendiumType
    Dim i As Long
    
    mtypCompendium = typBlank
    SetDefaults seCompendium
    With mtypCompendium
        strLine = Split(pstrRaw, vbNewLine)
        For i = 1 To UBound(strLine)
            If ParseLine(strLine(i), strField, strValue) Then
                Select Case strField
                    Case "datafile": .DataFile = strValue
                    Case "runbefore": If LCase$(strValue) = "true" Then .RunBefore = True Else .RunBefore = False
                    Case "playbutton": If LCase$(strValue) = "true" Then .PlayButton = True Else .PlayButton = False
                    Case "playexe": .PlayEXE = strValue
                    Case "sortorder": .CompendiumOrder = val(strValue)
                    Case "levelsort": .LevelSort = val(strValue)
                    Case "abbreviatecolumns": If LCase$(strValue) = "true" Then .AbbreviateColumns = True Else .AbbreviateColumns = False
                    Case "abbreviatepacks": If LCase$(strValue) = "true" Then .AbbreviatePacks = True Else .AbbreviatePacks = False
                    Case "abbreviatepatrons": If LCase$(strValue) = "true" Then .AbbreviatePatrons = True Else .AbbreviatePatrons = False
                    Case "challengeorder": .ChallengeOrder = val(strValue)
                    Case "sagaepic": If LCase$(strValue) = "true" Then .SagaTier = steEpic Else .SagaTier = steHeroic ' For backwards compatibility
                    Case "sagatier": .SagaTier = val(strValue)
                    Case "sagasort": .SagaSort = val(strValue)
                    Case "sagabringtotop"
                        .SagaBringToTops = .SagaBringToTops + 1
                        ReDim Preserve .SagaBringToTop(1 To .SagaBringToTops)
                        .SagaBringToTop(.SagaBringToTops) = strValue
                    Case "scroll": .Scroll = val(strValue)
                    Case "wheelstep": .WheelStep = val(strValue)
                    Case "marginx": .MarginX = val(strValue)
                    Case "marginy": .MarginY = val(strValue)
                    Case "sides": .Sides = val(strValue)
                    Case "bottom": .Bottom = val(strValue)
                    Case "leftpane": .LeftPane = val(strValue)
                    Case "rightpane": .RightPane = val(strValue)
                    Case "character": .Character = val(strValue)
                    Case "difficulty": .Difficulty = val(strValue)
                    Case "backcolor": .BackColor = val(strValue)
                    Case "namedhigh": .NamedHigh = val(strValue)
                    Case "namedmed": .NamedMed = val(strValue)
                    Case "namedlow": .NamedLow = val(strValue)
                    Case "nameddim": .NamedDim = val(strValue)
                    Case "windowsize": .MainWindow.WindowSize = val(strValue)
                    Case "maximized": If LCase$(strValue) = "true" Then .MainWindow.Maximized = True Else .MainWindow.Maximized = False
                    Case "childwindows": If LCase$(strValue) = "false" Then .ChildWindows = False Else .ChildWindows = True
                    Case "windowcoords": ParseCoords .MainWindow.WindowCoords, strValue
                    Case "challengecoords": ParseCoords .Coords.Challenges, strValue
                    Case "sagacoords": ParseCoords .Coords.Sagas, strValue
                    Case "sagadetailcoords": ParseCoords .Coords.SagaDetail, strValue
                    Case "wildernesscoords": ParseCoords .Coords.Wilderness, strValue
                    Case "toolscoords": ParseCoords .Coords.Tools, strValue
                    Case "charactercoords": ParseCoords .Coords.Characters, strValue
                    Case "menucoords": ParseCoords .Coords.Menus, strValue
                    Case "patroncoords": ParseCoords .Coords.Patrons, strValue
                    Case "colorcoords": ParseCoords .Coords.Colors, strValue
                    Case "critcalculatorcoords": ParseCoords .Coords.CritCalculator, strValue
                End Select
            End If
        Next
        If Len(.DataFile) = 0 Then .DataFile = "Main"
    End With
End Sub

Private Sub ParseCoords(ptypCoords As CoordsType, pstrRaw As String)
    Dim strValue() As String
    
    strValue = Split(pstrRaw, vbTab)
    If UBound(strValue) <> 3 Then Exit Sub
    With ptypCoords
        .Left = val(strValue(0))
        .Top = val(strValue(1))
        .Width = val(strValue(2))
        .Height = val(strValue(3))
    End With
End Sub

Private Sub LoadPaths(pstrRaw As String)
    Dim strField As String
    Dim strValue As String
    Dim strLine() As String
    Dim i As Long
    
    If mblnOldPaths Then Exit Sub
    With mtypPath
        strLine = Split(pstrRaw, vbNewLine)
        For i = 1 To UBound(strLine)
            If ParseLine(strLine(i), strField, strValue) Then
                Select Case strField
                    Case "characterbuilderlite": .CharacterBuilderLite = strValue
                    Case "compendium": .Compendium = strValue
                    Case "cannithcrafting": .CannithCrafting = strValue
                    Case "ronplanner": .RonPath = strValue
                    Case "ddobuilder": .BuilderPath = strValue
                    Case Else: Debug.Print "Unknown Path: " & strValue
                End Select
            End If
        Next
    End With
End Sub

Private Function LoadShared(pstrRaw As String) As Boolean
    Dim strField As String
    Dim strValue As String
    Dim strLine() As String
    Dim blnValid As Boolean
    Dim typBlank As SharedType
    Dim i As Long
    
    With mtypShared
        strLine = Split(pstrRaw, vbNewLine)
        For i = 1 To UBound(strLine)
            If ParseLine(strLine(i), strField, strValue) Then
                blnValid = True
                Select Case strField
                    Case "colorcoords": ParseCoords .ColorCoords, strValue
                    Case "screencolors": .ScreenColors = strValue
                    Case "outputcolors": .OutputColors = strValue
                    Case "colorversion": .ColorVersion = val(strValue)
                    Case Else: blnValid = False
                End Select
                If blnValid Then LoadShared = True
            End If
        Next
    End With
End Function

Public Function LoadOutput(pstrRaw As String) As Boolean
    LoadOutput = ParseOutputData(pstrRaw, mtypOutput)
End Function

Private Function ParseOutputData(pstrRaw As String, ptypOutput As OutputType) As Boolean
    Dim strField As String
    Dim strValue As String
    Dim strLine() As String
    Dim blnValid As Boolean
    Dim i As Long
    
    With ptypOutput
        strLine = Split(pstrRaw, vbNewLine)
        For i = 0 To UBound(strLine)
            If ParseLine(strLine(i), strField, strValue) Then
                blnValid = True
                Select Case strField
                    Case "codes": If LCase$(strValue) = "true" Then .Codes = True Else .Codes = False
                    Case "dots": If LCase$(strValue) = "true" Then .Dots = True Else .Dots = False
                    Case "textcolors": If LCase$(strValue) = "true" Then .TextColors = True Else .TextColors = False
                    Case "bold": ParseTags .Bold, strValue
                    Case "underline": ParseTags .Underline, strValue
                    Case "fixed": ParseTags .Fixed, strValue
                    Case "lists": ParseListTags .Lists, strValue
                    Case "color": ParseTags .Color, strValue
                    Case "wrapper": ParseTags .Wrapper, strValue
                    Case "reddit": If LCase$(strValue) = "true" Then .Reddit = True Else .Reddit = False
                    Case Else: blnValid = False
                End Select
                If blnValid Then ParseOutputData = True
            End If
        Next
    End With
End Function

Private Function InitFormat(pstrFormatName As String, ptypFormat As OutputType) As Boolean
    With ptypFormat
        .Codes = True
        .Dots = True
        .TextColors = True
        .Reddit = False
        InitTags .Bold, True, "[b]", "[/b]"
        InitTags .Underline, True, "[u]", "[/u]"
        InitTags .Fixed, True, "[font=courier]", "[/font]"
        InitListTags .Lists, True, "[list]", "[/list]", "[list=1]", "[/list]"
        InitTags .Color, True, "[color=#$]", "[/color]"
        InitTags .Wrapper, False, "[code]", "[/code]"
        Select Case pstrFormatName
            Case "Default", vbNullString
            Case "Code Tags"
                .Dots = False
                .Fixed.Enabled = False
                .Wrapper.Enabled = True
            Case "Plain Text"
                .Codes = False
                .Dots = False
                .TextColors = False
                .Bold.Enabled = False
                .Underline.Enabled = False
                .Fixed.Enabled = False
                .Lists.Enabled = False
                .Color.Enabled = False
                .Wrapper.Enabled = False
            Case "Reddit"
                .Reddit = True
                .Dots = False
                InitTags .Bold, False, vbNullString, vbNullString
                InitTags .Underline, False, vbNullString, vbNullString
                InitTags .Fixed, False, vbNullString, vbNullString
                InitListTags .Lists, False, vbNullString, vbNullString, vbNullString, vbNullString
                InitTags .Color, False, vbNullString, vbNullString
                InitTags .Wrapper, False, vbNullString, vbNullString
        End Select
    End With
    InitFormat = LoadFormat(pstrFormatName, ptypFormat)
End Function

Private Function LoadFormat(pstrFormatName As String, ptypFormat As OutputType) As Boolean
    Dim strFile As String
    Dim strRaw As String
    
    strFile = App.Path & "\Settings\" & pstrFormatName & ".bbcodes"
    If xp.File.Exists(strFile) Then
        strRaw = xp.File.LoadToString(strFile)
        LoadFormat = ParseOutputData(strRaw, ptypFormat)
    Else
        Select Case LCase$(pstrFormatName)
            Case "default", "code tags", "plain text", "reddit": LoadFormat = True
        End Select
    End If
End Function

Private Sub InitTags(ptypTags As TagsType, pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    With ptypTags
        .Enabled = pblnEnabled
        .OpenTag = pstrOpen
        .CloseTag = pstrClose
    End With
End Sub

Private Sub InitListTags(ptypListTags As ListTagsType, pblnEnabled As Boolean, pstrBulletOpen As String, pstrBulletClose As String, pstrNumberedOpen As String, pstrNumberedClose As String)
    With ptypListTags
        .Enabled = pblnEnabled
        .BulletOpen = pstrBulletOpen
        .BulletClose = pstrBulletClose
        .NumberedOpen = pstrNumberedOpen
        .NumberedClose = pstrNumberedClose
    End With
End Sub

Private Sub ParseTags(ptypTags As TagsType, pstrRaw As String)
    Dim strValue() As String
    
    strValue = Split(pstrRaw, vbTab)
    If UBound(strValue) <> 2 Then Exit Sub
    With ptypTags
        If LCase$(strValue(0)) = "true" Then .Enabled = True Else .Enabled = False
        .OpenTag = strValue(1)
        .CloseTag = strValue(2)
    End With
End Sub

Private Sub ParseListTags(ptypListTags As ListTagsType, pstrRaw As String)
    Dim strValue() As String
    
    strValue = Split(pstrRaw, vbTab)
    If UBound(strValue) <> 4 Then Exit Sub
    With ptypListTags
        If LCase$(strValue(0)) = "true" Then .Enabled = True Else .Enabled = False
        .BulletOpen = strValue(1)
        .BulletClose = strValue(2)
        .NumberedOpen = strValue(3)
        .NumberedClose = strValue(4)
    End With
End Sub

Private Function LoadColors(pstrRaw As String, penColorFile As ColorFileEnum) As Boolean
    Dim enGroup As ColorGroupEnum
    Dim enColor As ColorValueEnum
    Dim strField As String
    Dim strValue As String
    Dim strLine() As String
    Dim blnApply As Boolean
    Dim lngValue As Long
    Dim i As Long
    
    If mtypShared.ColorVersion < ColorVersion Then
        mtypShared.ScreenColors = "Black"
        mtypShared.OutputColors = "Default"
        mtypShared.ColorVersion = ColorVersion
        SaveSettings False
        DeleteColorFiles
        Exit Function
    End If
    ' Parse settings
    strLine = Split(pstrRaw, vbNewLine)
    For i = 0 To UBound(strLine)
        If ParseLine(strLine(i), strField, strValue) Then
            If ParseColor(strField, enGroup, enColor) Then
                Select Case penColorFile
                    Case cfeScreen: blnApply = (enGroup <> cgeOutput)
                    Case cfeOutput: blnApply = (enGroup = cgeOutput)
                    Case cfeAll: blnApply = True
                End Select
                If blnApply Then
                    lngValue = ColorToValue(strValue)
                    If lngValue <> -1 Then mlngColor(enGroup, enColor) = lngValue
                End If
                LoadColors = True
            End If
        End If
    Next
End Function

' This is a one-time deletion
Private Sub DeleteColorFiles()
    Select Case ColorVersion
        Case 1
            DeleteFile "Anakin.screen"
            DeleteFile "Arctic.screen"
            DeleteFile "Blue.screen"
            DeleteFile "Gold.screen"
            DeleteFile "Green.screen"
            DeleteFile "Mac Rumors.screen"
            DeleteFile "McCarthy.screen"
            DeleteFile "Pastel.screen"
            DeleteFile "Prospect.screen"
            DeleteFile "Purple.screen"
            DeleteFile "Raspberry.screen"
            DeleteFile "Squarefactor.screen"
            DeleteFile "Sublime.screen"
            DeleteFile "White.screen"
    End Select
End Sub

Private Sub DeleteFile(pstrFile As String)
    Dim strFile As String
    
    strFile = App.Path & "\Settings\" & pstrFile
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
End Sub

Private Function ParseColor(pstrRaw As String, penGroup As ColorGroupEnum, penColor As ColorValueEnum) As Boolean
    Dim strGroup As String
    Dim strColor As String
    
    If Len(pstrRaw) < 4 Then Exit Function
    strGroup = Left$(pstrRaw, 3)
    strColor = Mid$(pstrRaw, 4)
    Select Case LCase$(strGroup)
        Case "wrk": penGroup = cgeWorkspace
        Case "ctl": penGroup = cgeControls
        Case "nav": penGroup = cgeNavigation
        Case "drp": penGroup = cgeDropSlots
        Case "out": penGroup = cgeOutput
        Case Else: Exit Function
    End Select
    Select Case LCase$(strColor)
        Case "background": penColor = cveBackground
        Case "backhighlight": penColor = cveBackHighlight
        Case "backerror": penColor = cveBackError
        Case "backsub": penColor = cveBackRelated
        Case "text": penColor = cveText
        Case "texterror": penColor = cveTextError
        Case "textdim": penColor = cveTextDim
        Case "textlink": penColor = cveTextLink
        Case "borderint": penColor = cveBorderInterior
        Case "borderext": penColor = cveBorderExterior
        Case "borderhighlight": penColor = cveBorderHighlight
        Case "gray": penColor = cveLightGray
        Case "red": penColor = cveRed
        Case "yellow": penColor = cveYellow
        Case "blue": penColor = cveBlue
        Case "purple": penColor = cvePurple
        Case "green": penColor = cveGreen
        Case "orange": penColor = cveOrange
        Case Else: Exit Function
    End Select
    ParseColor = True
End Function

Private Function ColorToValue(pstrRaw As String) As Long
    Dim lngValue As Long
    Dim strSplit() As String
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    lngValue = -1
    If InStr(pstrRaw, ",") Then
        strSplit = Split(pstrRaw, ",")
        If UBound(strSplit) = 2 Then
            lngRed = val(strSplit(0))
            lngGreen = val(strSplit(1))
            lngBlue = val(strSplit(2))
            lngValue = RGB(lngRed, lngGreen, lngBlue)
        End If
    ElseIf IsNumeric(pstrRaw) Then
        lngValue = val(pstrRaw)
    Else
        Select Case pstrRaw
            Case "Button Text": lngValue = vbButtonText
            Case "Button Face": lngValue = vbButtonFace
            Case "Button Shadow": lngValue = vbButtonShadow
            Case "Window Text": lngValue = vbWindowText
            Case "Window Background": lngValue = vbWindowBackground
            Case "Window Frame": lngValue = vbWindowFrame
            Case "Application Workspace": lngValue = vbApplicationWorkspace
            Case "Gray Text": lngValue = vbGrayText
            Case "Inactive Caption Text": lngValue = vbInactiveCaptionText
            Case "Menu Text": lngValue = vbMenuText
            Case "Highlight Text": lngValue = vbHighlightText
            Case "Highlight": lngValue = vbHighlight
            Case "Active Border": lngValue = vbActiveBorder
            Case "Inactive Border": lngValue = vbInactiveBorder
            Case "Info Text": lngValue = vbInfoText
            Case "Info Background": lngValue = vbInfoBackground
            Case "TitleBar Text": lngValue = vbTitleBarText
            Case "Active Title Bar": lngValue = vbActiveTitleBar
            Case "Inactive Title Bar": lngValue = vbInactiveTitleBar
            Case "Menu Bar": lngValue = vbMenuBar
            Case "Scroll Bars": lngValue = vbScrollBars
            Case "Desktop": lngValue = vbDesktop
            Case "3D Face": lngValue = vb3DFace
            Case "3D Light": lngValue = vb3DLight
            Case "3D Shadow": lngValue = vb3DShadow
            Case "3D Dark Shadow": lngValue = vb3DDKShadow
            Case "3D Highlight": lngValue = vb3DHighlight
        End Select
    End If
    ColorToValue = lngValue
End Function

Private Function ParseLine(ByVal pstrLine As String, pstrField As String, pstrValue As String) As Boolean
    Dim lngPos As Long
    Dim i As Long
    
    pstrField = vbNullString
    pstrValue = vbNullString
    If Len(pstrLine) = 0 Then Exit Function
    lngPos = InStr(pstrLine, "=")
    If lngPos = 0 Then Exit Function
    pstrField = LCase$(Trim$(Left$(pstrLine, lngPos - 1)))
    If lngPos < Len(pstrLine) Then pstrValue = Trim$(Mid$(pstrLine, lngPos + 1))
    ParseLine = True
End Function


' ************* SAVE SETTINGS *************


Public Sub SaveSettings(Optional pblnReloadPaths As Boolean = True)
    Dim strFile As String
    Dim strCharBuild As String
    Dim strCompendium As String
    Dim strPaths As String
    Dim strShared As String
    Dim strOutput As String
    Dim strColor As String
    
    If Not ReloadExternalSections(pblnReloadPaths) Then Exit Sub
    strCharBuild = MakeCharBuildSettings()
    strCompendium = MakeCompendiumSettings()
    strPaths = MakePathSettings()
    strOutput = "Section: Output" & vbNewLine & vbNewLine & MakeOutputSettings()
    strShared = MakeSharedSettings()
    strColor = "Section: Colors" & vbNewLine & vbNewLine & MakeColorSettings(cfeAll)
    strFile = SettingsFile()
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    xp.File.SaveStringAs strFile, strCharBuild & strCompendium & strPaths & strOutput & strShared & strColor
End Sub

Private Function ReloadExternalSections(pblnReloadPaths As Boolean) As Boolean
    Dim strFile As String
    Dim strRaw As String
    Dim strSection() As String
    Dim strTitle As String
    Dim i As Long
    
    Select Case App.Title
        Case "Character Builder Lite", "Compendium", "Colors", "Cannith Crafting Builder Lite"
        Case Else: Exit Function
    End Select
    ReloadExternalSections = True
    strFile = SettingsFile()
    If Not xp.File.Exists(strFile) Then Exit Function
    strRaw = xp.File.LoadToString(strFile)
    strSection = Split(strRaw, "Section: ")
    For i = 1 To UBound(strSection)
        strTitle = Left$(strSection(i), InStr(strSection(i), vbNewLine) - 1)
        Select Case strTitle
            Case "Character Builder Lite"
                If App.Title <> "Character Builder Lite" Then LoadCharBuild strSection(i)
            Case "Compendium"
                If App.Title <> "Compendium" Then LoadCompendium strSection(i)
            Case "Output"
                If App.Title <> "Character Builder Lite" Then LoadOutput strSection(i)
            Case "Paths"
                If mblnOldPaths Then
                    mblnOldPaths = False
                ElseIf pblnReloadPaths Then
                    LoadPaths strSection(i)
                End If
        End Select
    Next
End Function

Private Function MakeCharBuildSettings() As String
    InitString
    AddString "Section: Character Builder Lite", 1
    With mtypCharBuild
        AddString "WindowSize=" & .MainWindow.WindowSize
        AddString "Maximized=" & .MainWindow.Maximized
        AddString "WindowCoords=" & CoordsToString(.MainWindow.WindowCoords)
        AddString "DialogCoords=" & CoordsToString(.DialogCoords)
        AddString "OptionsCoords=" & CoordsToString(.OptionsCoords)
        AddString "ConvertCoords=" & CoordsToString(.ConvertCoords)
        AddString "ConvertOnStartup=" & .ConvertOnStartup
        AddString "ChildWindows=" & .ChildWindows
        AddString "ShowErrorsOnStartup=" & .ShowErrors
        AddString "UseIcons=" & .UseIcons
        AddString "IconOverview=" & .IconOverview
        AddString "IconSkills=" & .IconSkills
        AddString "MRUCount=" & .MRUCount
        AddString "MRUList=" & JoinMRU()
        AddString "BuildPoints=" & .BuildPoints
        AddString "SkillOrderScreen=" & .SkillOrderScreen
        AddString "SkillOrderOutput=" & .SkillOrderOutput
        AddString "FeatOrder=" & .FeatOrder
        AddString "FeatChannels=" & .FeatChannels
        AddString "FeatOrderOutput=" & .FeatOrderOutput
        AddString "FeatShowAll=" & .FeatShowAll '2021.08.06 added FeatShowAll flag
        AddString "BBCodes=" & .BBCodes
        AddString "OutputMargin=" & .OutputMargin
        AddString "Confirm=" & .Confirm
        AddString "Contrast=" & .Contrast
        AddString "Version=" & .Version
        AddString "ProcessTime=" & .ProcessTime
    End With
    MakeCharBuildSettings = GetString(1)
End Function

Private Function JoinMRU() As String
    Dim strJoin As String
    
    With mtypCharBuild
        If .MRUCount > 0 Then
            If UBound(.MRUList) <> .MRUCount Then ReDim Preserve .MRUList(1 To .MRUCount)
            strJoin = Join(.MRUList, vbTab)
            JoinMRU = strJoin
        End If
    End With
End Function

Private Function MakeCompendiumSettings() As String
    Dim i As Long
    
    InitString
    AddString "Section: Compendium", 1
    With mtypCompendium
        AddString "DataFile=" & .DataFile
        AddString "RunBefore=" & .RunBefore
        AddString "PlayButton=" & .PlayButton
        AddString "PlayEXE=" & .PlayEXE
        AddString "SortOrder=" & .CompendiumOrder
        AddString "LevelSort=" & .LevelSort
        AddString "AbbreviateColumns=" & .AbbreviateColumns
        AddString "AbbreviatePacks=" & .AbbreviatePacks
        AddString "AbbreviatePatrons=" & .AbbreviatePatrons
        AddString "ChallengeOrder=" & .ChallengeOrder
        AddString "SagaTier=" & .SagaTier
        AddString "SagaSort=" & .SagaSort
        For i = 1 To .SagaBringToTops
            AddString "SagaBringToTop=" & .SagaBringToTop(i)
        Next
        AddString "Scroll=" & .Scroll
        AddString "WheelStep=" & mtypCompendium.WheelStep
        AddString "MarginX=" & .MarginX
        AddString "MarginY=" & .MarginY
        AddString "Sides=" & .Sides
        AddString "Bottom=" & .Bottom
        AddString "LeftPane=" & .LeftPane
        AddString "RightPane=" & .RightPane
        AddString "Character=" & .Character
        AddString "Difficulty=" & .Difficulty
        AddString "BackColor=" & .BackColor
        AddString "NamedHigh=" & .NamedHigh
        AddString "NamedMed=" & .NamedMed
        AddString "NamedLow=" & .NamedLow
        AddString "NamedDim=" & .NamedDim
        AddString "WindowSize=" & .MainWindow.WindowSize
        AddString "Maximized=" & .MainWindow.Maximized
        AddString "ChildWindows=" & .ChildWindows
        AddString "WindowCoords=" & CoordsToString(.MainWindow.WindowCoords)
        AddString "ChallengeCoords=" & CoordsToString(.Coords.Challenges)
        AddString "SagaCoords=" & CoordsToString(.Coords.Sagas)
        AddString "SagaDetailCoords=" & CoordsToString(.Coords.SagaDetail)
        AddString "WildernessCoords=" & CoordsToString(.Coords.Wilderness)
        AddString "ToolsCoords=" & CoordsToString(.Coords.Tools)
        AddString "CharacterCoords=" & CoordsToString(.Coords.Characters)
        AddString "MenuCoords=" & CoordsToString(.Coords.Menus)
        AddString "PatronCoords=" & CoordsToString(.Coords.Patrons)
        AddString "ColorCoords=" & CoordsToString(.Coords.Colors)
        AddString "CritCalculatorCoords=" & CoordsToString(.Coords.CritCalculator)
    End With
    MakeCompendiumSettings = GetString(1)
End Function

Private Function MakePathSettings() As String
    Dim blnInclude As Boolean
    
    InitString
    With mtypPath
        AddPath "CharacterBuilderLite", .CharacterBuilderLite, DefaultPath(), blnInclude
        AddPath "Compendium", .Compendium, DefaultPath(), blnInclude
        AddPath "CannithCrafting", .CannithCrafting, DefaultPath(), blnInclude
        AddPath "RonPlanner", .RonPath, xp.Folder.UserDocs, blnInclude
        AddPath "DDOBuilder", .BuilderPath, xp.Folder.UserDocs, blnInclude
    End With
    If blnInclude Then MakePathSettings = GetString(1)
End Function

Private Sub AddPath(pstrName As String, pstrPath As String, pstrDefault As String, pblnInclude As Boolean)
    If Not CustomPath(pstrPath, pstrDefault) Then Exit Sub
    If Not pblnInclude Then
        AddString "Section: Paths", 1
        pblnInclude = True
    End If
    AddString pstrName & "=" & pstrPath
End Sub

Private Function CustomPath(pstrPath As String, pstrDefault As String) As Boolean
    If Len(pstrPath) Then
        If pstrPath <> pstrDefault Then
            If xp.Folder.Exists(pstrPath) Then CustomPath = True
        End If
    End If
End Function

Private Function MakeSharedSettings() As String
    InitString
    AddString "Section: Shared", 1
    With mtypShared
        AddString "ColorCoords=" & CoordsToString(.ColorCoords)
        AddString "ScreenColors=" & .ScreenColors
        AddString "OutputColors=" & .OutputColors
        AddString "ColorVersion=" & .ColorVersion
    End With
    MakeSharedSettings = GetString(1)
End Function

Public Function MakeOutputSettings() As String
    InitString
    With mtypOutput
        AddString "Codes=" & .Codes
        AddString "Dots=" & .Dots
        AddString "TextColors=" & .TextColors
        AddString "Bold=" & TagsToString(.Bold)
        AddString "Underline=" & TagsToString(.Underline)
        AddString "Fixed=" & TagsToString(.Fixed)
        AddString "Lists=" & ListTagsToString(.Lists)
        AddString "Color=" & TagsToString(.Color)
        AddString "Wrapper=" & TagsToString(.Wrapper)
        If .Reddit Then AddString "Reddit=True"
    End With
    MakeOutputSettings = GetString(1)
End Function

Public Function MakeColorSettings(penColorFile As ColorFileEnum) As String
    Dim lngGroup As Long
    Dim lngColor As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    
    Select Case penColorFile
        Case cfeAll
            lngFirst = 0
            lngLast = 4
        Case cfeOutput
            lngFirst = 4
            lngLast = 4
        Case cfeScreen
            lngFirst = 0
            lngLast = 3
    End Select
    ' Colors
    InitString
    For lngGroup = lngFirst To lngLast
        For lngColor = 0 To 17
            AddString ColorToString(lngGroup, lngColor)
        Next
    Next
    MakeColorSettings = GetString(1)
End Function

Private Function ColorToString(penGroup As ColorGroupEnum, penColor As ColorValueEnum) As String
    Dim strGroup As String
    Dim strValue As String
    Dim strColor As String
    
    Select Case penGroup
        Case cgeWorkspace: strGroup = "Wrk"
        Case cgeControls: strGroup = "Ctl"
        Case cgeNavigation: strGroup = "Nav"
        Case cgeDropSlots: strGroup = "Drp"
        Case cgeOutput: strGroup = "Out"
    End Select
    Select Case penColor
        Case cveBackground: strValue = "Background"
        Case cveBackHighlight: strValue = "BackHighlight"
        Case cveBackError: strValue = "BackError"
        Case cveBackRelated: strValue = "BackSub"
        Case cveText: strValue = "Text"
        Case cveTextError: strValue = "TextError"
        Case cveTextDim: strValue = "TextDim"
        Case cveTextLink: strValue = "TextLink"
        Case cveBorderInterior: strValue = "BorderInt"
        Case cveBorderExterior: strValue = "BorderExt"
        Case cveBorderHighlight: strValue = "BorderHighlight"
        Case cveLightGray: strValue = "Gray"
        Case cveRed: strValue = "Red"
        Case cveYellow: strValue = "Yellow"
        Case cveBlue: strValue = "Blue"
        Case cvePurple: strValue = "Purple"
        Case cveGreen: strValue = "Green"
        Case cveOrange: strValue = "Orange"
    End Select
    If mlngColor(penGroup, penColor) < 0 Then
        Select Case mlngColor(penGroup, penColor)
            Case vbButtonText: strColor = "Button Text"
            Case vbButtonFace: strColor = "Button Face"
            Case vbButtonShadow: strColor = "Button Shadow"
            Case vbWindowText: strColor = "Window Text"
            Case vbWindowBackground: strColor = "Window Background"
            Case vbWindowFrame: strColor = "Window Frame"
            Case vbApplicationWorkspace: strColor = "Application Workspace"
            Case vbGrayText: strColor = "Gray Text"
            Case vbInactiveCaptionText: strColor = "Inactive Caption Text"
            Case vbMenuText: strColor = "Menu Text"
            Case vbHighlightText: strColor = "Highlight Text"
            Case vbHighlight: strColor = "Highlight"
            Case vbActiveBorder: strColor = "Active Border"
            Case vbInactiveBorder: strColor = "Inactive Border"
            Case vbInfoText: strColor = "Info Text"
            Case vbInfoBackground: strColor = "Info Background"
            Case vbTitleBarText: strColor = "TitleBar Text"
            Case vbActiveTitleBar: strColor = "Active Title Bar"
            Case vbInactiveTitleBar: strColor = "Inactive Title Bar"
            Case vbMenuBar: strColor = "Menu Bar"
            Case vbScrollBars: strColor = "Scroll Bars"
            Case vbDesktop: strColor = "Desktop"
            Case vb3DFace: strColor = "3D Face"
            Case vb3DLight: strColor = "3D Light"
            Case vb3DShadow: strColor = "3D Shadow"
            Case vb3DDKShadow: strColor = "3D Dark Shadow"
            Case vb3DHighlight: strColor = "3D Highlight"
            Case Else: strColor = mlngColor(penGroup, penColor)
        End Select
    Else
        strColor = ColorToRGB(mlngColor(penGroup, penColor))
    End If
    ColorToString = strGroup & strValue & "=" & strColor
End Function

Private Function ColorToRGB(plngColor As Long) As String
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    xp.ColorToRGB plngColor, lngRed, lngGreen, lngBlue
    ColorToRGB = lngRed & "," & lngGreen & "," & lngBlue
End Function

Private Sub InitString()
    mlngLine = -1
    ReDim mstrArray(256)
End Sub

Private Sub AddString(pstrLine As String, Optional plngExtraBlankLines As Long = 0)
    mlngLine = mlngLine + 1
    mstrArray(mlngLine) = pstrLine
    mlngLine = mlngLine + plngExtraBlankLines
End Sub

Private Function GetString(Optional plngExtraBlankLines As Long = 0) As String
    ReDim Preserve mstrArray(mlngLine + 2 + plngExtraBlankLines)
    GetString = Join(mstrArray, vbNewLine)
End Function

Private Function CoordsToString(ptypCoords As CoordsType) As String
    With ptypCoords
        CoordsToString = .Left & vbTab & .Top & vbTab & .Width & vbTab & .Height
    End With
End Function

Private Function TagsToString(ptypTags As TagsType) As String
    With ptypTags
        TagsToString = .Enabled & vbTab & .OpenTag & vbTab & .CloseTag
    End With
End Function

Private Function ListTagsToString(ptypListTags As ListTagsType) As String
    With ptypListTags
        ListTagsToString = .Enabled & vbTab & .BulletOpen & vbTab & .BulletClose & vbTab & .NumberedOpen & vbTab & .NumberedClose
    End With
End Function


' ************* MESSAGES *************


Public Sub MessageAdd(pstrMessage As String)
    If Len(mstrMessages) Then mstrMessages = mstrMessages & vbNewLine
    mstrMessages = mstrMessages & pstrMessage
End Sub

Public Sub MessageShow()
    If Len(mstrMessages) = 0 Then Exit Sub
    MsgBox mstrMessages, vbInformation, "Notice"
    mstrMessages = vbNullString
End Sub


' ************* UTILS *************


Public Sub RunUtil(penUtil As UtilEnum)
    Dim strFile As String
    Dim strApp As String
    
    Select Case penUtil
        Case ueColors: strApp = "Colors"
        Case ueLightsOut: strApp = "LightsOut"
        Case ueADQ: strApp = "ADQ"
        Case ueStopwatch: strApp = "Stopwatch"
    End Select
    strFile = App.Path & "\Utils\" & strApp & ".exe"
    If Not xp.File.Exists(strFile) Then
        Notice "File not found:" & vbNewLine & vbNewLine & strFile
        Exit Sub
    End If
    xp.File.Run strFile
    ' If the app isn't already loaded, trying to activate it will throw a runtime error
    On Error Resume Next
    ' If the app is already running, activate it to bring it to the foreground
    AppActivate strApp
End Sub


' ************* WINDOWS *************


Public Property Get WindowSize() As WindowSizeEnum
    Select Case App.Title
        Case "Character Builder Lite": WindowSize = mtypCharBuild.MainWindow.WindowSize
        Case "Compendium": WindowSize = mtypCompendium.MainWindow.WindowSize
    End Select
End Property

Public Property Let WindowSize(ByVal penWindowSize As WindowSizeEnum)
    Dim frm As Form
    
    Select Case App.Title
        Case "Character Builder Lite": mtypCharBuild.MainWindow.WindowSize = penWindowSize
        Case "Compendium": mtypCompendium.MainWindow.WindowSize = penWindowSize
    End Select
    SizeWindow
End Property

Public Sub SizeWindow()
    Dim typMain As MainWindowType
    Dim blnMaximized As Boolean
    Dim blnMinimized As Boolean
    Dim enWindowSize As WindowSizeEnum
    Dim frm As Form
    
    Select Case App.Title
        Case "Character Builder Lite"
            typMain = mtypCharBuild.MainWindow
            If Not GetForm(frm, "frmMain") Then Exit Sub
        Case "Compendium"
            typMain = mtypCompendium.MainWindow
            If Not GetForm(frm, "frmCompendium") Then Exit Sub
    End Select
    blnMaximized = (frm.WindowState = vbMaximized)
    blnMinimized = (frm.WindowState = vbMinimized)
    With typMain
        If .WindowSize = wseMaximized Or (.WindowSize = wseRemember And .Maximized = True) Then
            If Not blnMaximized Then frm.WindowState = vbMaximized
        Else
            If blnMaximized Or blnMinimized Then frm.WindowState = vbNormal
            enWindowSize = .WindowSize
            With .WindowCoords
                If enWindowSize = wseFillDesktop Or .Width = 0 Or .Height = 0 Then xp.GetDesktop .Left, .Top, .Width, .Height
                frm.Move .Left, .Top, .Width, .Height
            End With
        End If
    End With
End Sub

Public Sub SaveWindowSize()
    Dim typMain As MainWindowType
    Dim frm As Form
    
    Select Case App.Title
        Case "Character Builder Lite"
            typMain = mtypCharBuild.MainWindow
            If Not GetForm(frm, "frmMain") Then Exit Sub
        Case "Compendium"
            typMain = mtypCompendium.MainWindow
            If Not GetForm(frm, "frmCompendium") Then Exit Sub
    End Select
    typMain.Maximized = (frm.WindowState = vbMaximized)
    If frm.WindowState = vbNormal Then
        With typMain.WindowCoords
            .Left = frm.Left
            .Top = frm.Top
            .Width = frm.Width
            .Height = frm.Height
        End With
    End If
    Select Case App.Title
        Case "Character Builder Lite": mtypCharBuild.MainWindow = typMain
        Case "Compendium": mtypCompendium.MainWindow = typMain
    End Select
End Sub

Public Sub MoveForm(pfrm As Form, Optional pblnCascade As Boolean = False)
    Dim typCoords As CoordsType
    Dim lngTop As Long
    Dim lngOffset As Long
    Dim frm As Form
    Dim i As Long
    
    If pblnCascade Then
        If Not GetMainForm(frm) Then Exit Sub
        lngTop = frm.Height - frm.ScaleHeight
        lngOffset = lngTop - frm.TextHeight("X")
        lngOffset = pfrm.Height - pfrm.ScaleHeight
        For i = 1 To mlngForms
            If pfrm.Name = mstrForm(i) Then
                typCoords.Left = i * lngOffset
                typCoords.Top = lngTop + (i - 1) * lngOffset
                Exit For
            End If
        Next
    Else
        Select Case pfrm.Name
            Case "frmOverview", "frmStats", "frmSkills", "frmFeats", "frmSpells", "frmEnhancements", "frmDestiny2", "frmExport", "frmDeprecate": typCoords = mtypCharBuild.DialogCoords
            Case "frmOptions": typCoords = mtypCharBuild.OptionsCoords
            Case "frmConvert": typCoords = mtypCharBuild.ConvertCoords
            Case "frmColors": typCoords = mtypShared.ColorCoords
            Case "frmFormat"
            Case "frmChallenges": typCoords = mtypCompendium.Coords.Challenges
            Case "frmSagas": typCoords = mtypCompendium.Coords.Sagas
            Case "frmSagaDetail": typCoords = mtypCompendium.Coords.SagaDetail
            Case "frmWilderness": typCoords = mtypCompendium.Coords.Wilderness
            Case "frmTools": typCoords = mtypCompendium.Coords.Tools
            Case "frmCharacter": typCoords = mtypCompendium.Coords.Characters
            Case "frmMenuEditor": typCoords = mtypCompendium.Coords.Menus
            Case "frmPatrons": typCoords = mtypCompendium.Coords.Patrons
            Case "frmColorPreview": typCoords = mtypCompendium.Coords.Colors
            Case "frmCritCalculator": typCoords = mtypCompendium.Coords.CritCalculator
        End Select
        If typCoords.Left = 0 And typCoords.Top = 0 Then
            If GetForm(frm, "frmCompendium") Then
                typCoords.Left = frm.Left + (frm.Width - pfrm.Width) \ 2
                typCoords.Top = frm.Top + (frm.Height - pfrm.Height) \ 2
            Else
                typCoords.Left = (Screen.Width - pfrm.Width) \ 2
                typCoords.Top = (Screen.Height - pfrm.Height) \ 2
            End If
        End If
    End If
    With typCoords
        If .Width = 0 Or .Height = 0 Then
            pfrm.Move .Left, .Top
        Else
            pfrm.Move .Left, .Top, .Width, .Height
        End If
    End With
End Sub

Public Sub SavePosition(pfrm As Form, Optional pblnDimensions As Boolean = False)
    Dim typCoords As CoordsType
    
    With typCoords
        .Left = pfrm.Left
        .Top = pfrm.Top
        If pblnDimensions Then
            .Width = pfrm.Width
            .Height = pfrm.Height
        End If
    End With
    Select Case pfrm.Name
        Case "frmOverview", "frmStats", "frmSkills", "frmFeats", "frmSpells", "frmEnhancements", "frmDestiny2", "frmExport", "frmDeprecate": mtypCharBuild.DialogCoords = typCoords
        Case "frmOptions": mtypCharBuild.OptionsCoords = typCoords
        Case "frmConvert": mtypCharBuild.ConvertCoords = typCoords
        Case "frmColors": mtypShared.ColorCoords = typCoords
        Case "frmChallenges": mtypCompendium.Coords.Challenges = typCoords
        Case "frmSagas": mtypCompendium.Coords.Sagas = typCoords
        Case "frmSagaDetail": mtypCompendium.Coords.SagaDetail = typCoords
        Case "frmWilderness": mtypCompendium.Coords.Wilderness = typCoords
        Case "frmTools": mtypCompendium.Coords.Tools = typCoords
        Case "frmCharacter": mtypCompendium.Coords.Characters = typCoords
        Case "frmMenuEditor": mtypCompendium.Coords.Menus = typCoords
        Case "frmPatrons": mtypCompendium.Coords.Patrons = typCoords
        Case "frmColorPreview": mtypCompendium.Coords.Colors = typCoords
        Case "frmCritCalculator": mtypCompendium.Coords.CritCalculator = typCoords
    End Select
End Sub

Public Property Get ChildWindows() As Boolean
    Select Case App.Title
        Case "Character Builder Lite": ChildWindows = mtypCharBuild.ChildWindows
        Case "Compendium": ChildWindows = mtypCompendium.ChildWindows
    End Select
End Property

Public Property Let ChildWindows(ByVal pblnChildWindows As Boolean)
    Select Case App.Title
        Case "Character Builder Lite": mtypCharBuild.ChildWindows = pblnChildWindows
        Case "Compendium": mtypCompendium.ChildWindows = pblnChildWindows
    End Select
End Property

Private Function GetMainForm(pfrm As Form) As Boolean
    Select Case App.Title
        Case "Character Builder Lite": GetMainForm = GetForm(pfrm, "frmMain")
        Case "Compendium": GetMainForm = GetForm(pfrm, "frmCompendium")
    End Select
End Function


' ************* SHARED *************


Public Sub Configure(pfrm As Form)
    MoveForm pfrm
    RefreshColors pfrm
End Sub

Public Function GetExtension(penExtension As ExtensionEnum) As String
    Select Case penExtension
        Case eeScreenColor: GetExtension = "screen"
        Case eeOutputColor: GetExtension = "output"
        Case eeBBCodes: GetExtension = "bbcodes"
    End Select
End Function

Public Property Get ScreenColors() As String
    ScreenColors = mtypShared.ScreenColors
End Property

Public Property Let ScreenColors(ByVal pstrScreenColors As String)
    mtypShared.ScreenColors = pstrScreenColors
End Property

Public Property Get OutputColors() As String
    OutputColors = mtypShared.OutputColors
End Property

Public Property Let OutputColors(ByVal pstrOutputColors As String)
    mtypShared.OutputColors = pstrOutputColors
End Property

Public Property Get OutputTextColors() As Boolean
    OutputTextColors = mtypOutput.TextColors
End Property

Public Property Let OutputTextColors(ByVal pblnOutputTextColors As Boolean)
    mtypOutput.TextColors = pblnOutputTextColors
End Property

Public Property Get OutputReddit() As Boolean
    OutputReddit = mtypOutput.Reddit
End Property

Public Property Let OutputReddit(ByVal pblnOutputReddit As Boolean)
    mtypOutput.Reddit = pblnOutputReddit
End Property


' ************* CHARACTER BUILDER LITE *************


Public Sub CascadeWindows()
    Dim frm As Form
    Dim i As Long
    
    Erase mstrForm
    mlngForms = 0
    If GetForm(frm, "frmOverview") Then AddForm "frmOverview"
    If GetForm(frm, "frmStats") Then AddForm "frmStats"
    If GetForm(frm, "frmSkills") Then AddForm "frmSkills"
    If GetForm(frm, "frmFeats") Then AddForm "frmFeats"
    If GetForm(frm, "frmSpells") Then AddForm "frmSpells"
    If GetForm(frm, "frmEnhancements") Then AddForm "frmEnhancements"
    If GetForm(frm, "frmDestiny2") Then AddForm "frmDestiny2"
    If GetForm(frm, "frmExport") Then AddForm "frmExport"
    For i = 1 To mlngForms
        GetForm frm, mstrForm(i)
        MoveForm frm, True
        frm.SetFocus
    Next
    Set frm = Nothing
End Sub

Private Sub AddForm(pstrForm As String)
    mlngForms = mlngForms + 1
    mstrForm(mlngForms) = pstrForm
End Sub

Public Property Get BuildPoints() As BuildPointsEnum
    BuildPoints = mtypCharBuild.BuildPoints
End Property

Public Property Let BuildPoints(ByVal penBuildPoints As BuildPointsEnum)
    mtypCharBuild.BuildPoints = penBuildPoints
End Property

Public Property Get Dirty() As Boolean
    Dirty = mblnDirty
End Property

Public Property Let Dirty(ByVal pblnDirty As Boolean)
    mblnDirty = pblnDirty
End Property

Public Property Get FeatOrder() As FeatOrderEnum
    FeatOrder = mtypCharBuild.FeatOrder
End Property

Public Property Let FeatOrder(ByVal penFeatOrder As FeatOrderEnum)
    mtypCharBuild.FeatOrder = penFeatOrder
End Property

Public Property Get FeatChannels() As Boolean
    FeatChannels = mtypCharBuild.FeatChannels
End Property
Public Property Let FeatChannels(ByVal val As Boolean)
    mtypCharBuild.FeatChannels = val
End Property

Public Property Get FeatOrderOutput() As FeatOrderOutputEnum
    FeatOrderOutput = mtypCharBuild.FeatOrderOutput
End Property
Public Property Let FeatOrderOutput(ByVal penFeatOrderOutput As FeatOrderOutputEnum)
    mtypCharBuild.FeatOrderOutput = penFeatOrderOutput
End Property

'2021.08.06 Added FeatShowAll flag
Public Property Get FeatShowAll() As Boolean
    FeatShowAll = mtypCharBuild.FeatShowAll
End Property
Public Property Let FeatShowAll(ByVal val As Boolean)
    mtypCharBuild.FeatShowAll = val
End Property



Public Property Get OutputSection() As OutputEnum
    OutputSection = menOutputSection
End Property

Public Property Let OutputSection(ByVal penOutputSection As OutputEnum)
    menOutputSection = penOutputSection
End Property

Public Property Get OutputRefresh() As Boolean
    OutputRefresh = mblnRefresh
End Property

Public Property Let OutputRefresh(ByVal pblnRefresh As Boolean)
    mblnRefresh = pblnRefresh
End Property

Public Property Get ProcessTime() As Single
    ProcessTime = mtypCharBuild.ProcessTime
End Property

Public Property Let ProcessTime(ByVal psngProcessTime As Single)
    mtypCharBuild.ProcessTime = psngProcessTime
End Property

Public Property Get ShowErrors() As Boolean
    ShowErrors = mtypCharBuild.ShowErrors
End Property

Public Property Let ShowErrors(ByVal pblnShowErrors As Boolean)
    mtypCharBuild.ShowErrors = pblnShowErrors
End Property

Public Property Get SkillOrderScreen() As SkillOrderScreenEnum
    SkillOrderScreen = mtypCharBuild.SkillOrderScreen
End Property

Public Property Let SkillOrderScreen(ByVal penSkillOrderScreen As SkillOrderScreenEnum)
    mtypCharBuild.SkillOrderScreen = penSkillOrderScreen
End Property

Public Property Get SkillOrderOutput() As SkillOrderOutputEnum
    SkillOrderOutput = mtypCharBuild.SkillOrderOutput
End Property

Public Property Let SkillOrderOutput(ByVal penSkillOrderOutput As SkillOrderOutputEnum)
    mtypCharBuild.SkillOrderOutput = penSkillOrderOutput
End Property

Public Property Get IconOverview() As Boolean
    IconOverview = mtypCharBuild.IconOverview
End Property

Public Property Let IconOverview(ByVal pblnIconOverview As Boolean)
    mtypCharBuild.IconOverview = pblnIconOverview
End Property

Public Property Get IconSkills() As Boolean
    IconSkills = mtypCharBuild.IconSkills
End Property

Public Property Let IconSkills(ByVal pblnIconSkills As Boolean)
    mtypCharBuild.IconSkills = pblnIconSkills
End Property

Public Property Get Version() As String
    Version = mtypCharBuild.Version
End Property

Public Property Let Version(ByVal pstrVersion As String)
    mtypCharBuild.Version = pstrVersion
End Property

Public Property Get BBCodes() As String
    BBCodes = mtypCharBuild.BBCodes
End Property

Public Property Let BBCodes(ByVal pstrBBCodes As String)
    mtypCharBuild.BBCodes = pstrBBCodes
End Property

Public Property Get OutputMargin() As Long
    OutputMargin = mtypCharBuild.OutputMargin
End Property

Public Property Let OutputMargin(ByVal plngOutputMargin As Long)
    mtypCharBuild.OutputMargin = plngOutputMargin
End Property

Public Property Get UseIcons() As Boolean
    UseIcons = mtypCharBuild.UseIcons
End Property

Public Property Let UseIcons(ByVal pblnUseIcons As Boolean)
    mtypCharBuild.UseIcons = pblnUseIcons
End Property

Public Property Get Confirm() As Boolean
    Confirm = mtypCharBuild.Confirm
End Property

Public Property Let Confirm(ByVal pblnConfirm As Boolean)
    mtypCharBuild.Confirm = pblnConfirm
End Property

Public Property Get ConvertOnStartup() As Boolean
    ConvertOnStartup = mtypCharBuild.ConvertOnStartup
End Property

Public Property Let ConvertOnStartup(ByVal pblnConvertOnStartup As Boolean)
    mtypCharBuild.ConvertOnStartup = pblnConvertOnStartup
End Property

Public Property Get Contrast() As ContrastEnum
    Contrast = mtypCharBuild.Contrast
End Property

Public Property Let Contrast(penContrast As ContrastEnum)
    mtypCharBuild.Contrast = penContrast
End Property


' ************* COMPENDIUM *************


Public Property Get PlayButton() As Boolean
    PlayButton = mtypCompendium.PlayButton
End Property

Public Property Let PlayButton(ByVal pblnPlayButton As Boolean)
    mtypCompendium.PlayButton = pblnPlayButton
End Property

Public Property Get PlayEXE() As String
    PlayEXE = mtypCompendium.PlayEXE
End Property

Public Property Let PlayEXE(pstrPlayEXE As String)
    mtypCompendium.PlayEXE = pstrPlayEXE
End Property

Public Property Let RunBefore(pblnRunBefore As Boolean)
    mtypCompendium.RunBefore = True
End Property

Public Property Get RunBefore() As Boolean
    RunBefore = mtypCompendium.RunBefore
End Property

Public Property Let DataFile(ByVal pstrDataFile As String)
    mtypCompendium.DataFile = pstrDataFile
End Property

Public Property Get DataFile() As String
'    If Len(mtypCompendium.DataFile) = 0 Then mtypCompendium.DataFile = "Main"
    DataFile = mtypCompendium.DataFile
End Property

Public Property Let CompendiumOrder(ByVal penCompendiumOrder As CompendiumOrderEnum)
    mtypCompendium.CompendiumOrder = penCompendiumOrder
    mblnDirty = True
End Property

Public Property Get CompendiumOrder() As CompendiumOrderEnum
    CompendiumOrder = mtypCompendium.CompendiumOrder
End Property

Public Property Let ChallengeOrder(ByVal penChallengeOrder As ChallengeOrderEnum)
    mtypCompendium.ChallengeOrder = penChallengeOrder
    mblnDirty = True
End Property

Public Property Get SagaTier() As SagaTierEnum
    SagaTier = mtypCompendium.SagaTier
End Property

Public Property Let SagaTier(ByVal penSagaTier As SagaTierEnum)
    mtypCompendium.SagaTier = penSagaTier
End Property

Public Property Get SagaSort() As SagaSortEnum
    SagaSort = mtypCompendium.SagaSort
End Property

Public Property Let SagaSort(ByVal penSagaSort As SagaSortEnum)
    mtypCompendium.SagaSort = penSagaSort
End Property

Public Sub AddSagaBringToTop(pstrSaga As String)
    Dim i As Long
    
    With mtypCompendium
        For i = 1 To .SagaBringToTops
            If .SagaBringToTop(i) = pstrSaga Then Exit For
        Next
        If i > .SagaBringToTops Then
            .SagaBringToTops = .SagaBringToTops + 1
            ReDim Preserve .SagaBringToTop(1 To .SagaBringToTops)
        Else
            For i = i To .SagaBringToTops - 1
                .SagaBringToTop(i) = .SagaBringToTop(i + 1)
            Next
        End If
        .SagaBringToTop(.SagaBringToTops) = pstrSaga
    End With
End Sub

Public Function GetSagaBringToTop(pstrList() As String) As Long
    With mtypCompendium
        pstrList = .SagaBringToTop
        GetSagaBringToTop = .SagaBringToTops
    End With
End Function

Public Sub ClearSagaBringToTop()
    With mtypCompendium
        Erase .SagaBringToTop
        .SagaBringToTops = 0
    End With
End Sub

Public Property Let LevelSort(ByVal penLevelSort As LevelSortEnum)
    mtypCompendium.LevelSort = penLevelSort
    mblnDirty = True
End Property

Public Property Get LevelSort() As LevelSortEnum
    LevelSort = mtypCompendium.LevelSort
End Property

Public Property Let AbbreviateColumns(pblnAbbreviateColumns As Boolean)
    mtypCompendium.AbbreviateColumns = pblnAbbreviateColumns
End Property

Public Property Get AbbreviateColumns() As Boolean
    AbbreviateColumns = mtypCompendium.AbbreviateColumns
End Property

Public Property Let AbbreviatePacks(pblnAbbreviatePacks As Boolean)
    mtypCompendium.AbbreviatePacks = pblnAbbreviatePacks
End Property

Public Property Get AbbreviatePacks() As Boolean
    AbbreviatePacks = mtypCompendium.AbbreviatePacks
End Property

Public Property Let AbbreviatePatrons(pblnAbbreviatePatrons As Boolean)
    mtypCompendium.AbbreviatePatrons = pblnAbbreviatePatrons
End Property

Public Property Get AbbreviatePatrons() As Boolean
    AbbreviatePatrons = mtypCompendium.AbbreviatePatrons
End Property

Public Property Get ChallengeOrder() As ChallengeOrderEnum
    ChallengeOrder = mtypCompendium.ChallengeOrder
End Property

Public Property Let Difficulty(ByVal penDifficulty As ProgressEnum)
    mtypCompendium.Difficulty = penDifficulty
    mblnDirty = True
End Property

Public Property Get Difficulty() As ProgressEnum
    Difficulty = mtypCompendium.Difficulty
End Property

Public Property Let LeftPane(ByVal penLeftPane As PaneEnum)
    mtypCompendium.LeftPane = penLeftPane
    mblnDirty = True
End Property

Public Property Get LeftPane() As PaneEnum
    LeftPane = mtypCompendium.LeftPane
End Property

Public Property Let RightPane(ByVal penRightPane As PaneEnum)
    mtypCompendium.RightPane = penRightPane
    mblnDirty = True
End Property

Public Property Get RightPane() As PaneEnum
    RightPane = mtypCompendium.RightPane
End Property

Public Property Let Character(ByVal plngCharacter As Long)
    mtypCompendium.Character = plngCharacter
End Property

Public Property Get Character() As Long
    Character = mtypCompendium.Character
End Property

Public Property Let CompendiumBackColor(ByVal plngBackColor As Long)
    mtypCompendium.BackColor = plngBackColor
End Property

Public Property Get CompendiumBackColor() As Long
    CompendiumBackColor = mtypCompendium.BackColor
End Property

Public Property Let CompendiumScroll(ByVal plngScroll As Long)
    mtypCompendium.Scroll = plngScroll
    mblnDirty = True
End Property

Public Property Get CompendiumScroll() As Long
    CompendiumScroll = mtypCompendium.Scroll
End Property

Public Property Let WheelStep(ByVal plngWheelStep As Long)
    mtypCompendium.WheelStep = plngWheelStep
    mblnDirty = True
End Property

Public Property Get WheelStep() As Long
    WheelStep = mtypCompendium.WheelStep
End Property

Public Property Let MarginX(ByVal plngMarginX As Long)
    mtypCompendium.MarginX = plngMarginX
    mblnDirty = True
End Property

Public Property Get MarginX() As Long
    MarginX = mtypCompendium.MarginX
End Property

Public Property Let MarginY(ByVal plngMarginY As Long)
    mtypCompendium.MarginY = plngMarginY
    mblnDirty = True
End Property

Public Property Get MarginY() As Long
    MarginY = mtypCompendium.MarginY
End Property

Public Property Let Sides(ByVal plngSides As Long)
    mtypCompendium.Sides = plngSides
    mblnDirty = True
End Property

Public Property Get Sides() As Long
    Sides = mtypCompendium.Sides
End Property

Public Property Let Bottom(ByVal plngBottom As Long)
    mtypCompendium.Bottom = plngBottom
    mblnDirty = True
End Property

Public Property Get Bottom() As Long
    Bottom = mtypCompendium.Bottom
End Property

Public Property Let NamedHigh(ByVal plngNamedHigh As Long)
    mtypCompendium.NamedHigh = plngNamedHigh
    mblnDirty = True
End Property

Public Property Get NamedHigh() As Long
    NamedHigh = mtypCompendium.NamedHigh
End Property

Public Property Let NamedMed(ByVal plngNamedMed As Long)
    mtypCompendium.NamedMed = plngNamedMed
    mblnDirty = True
End Property

Public Property Get NamedMed() As Long
    NamedMed = mtypCompendium.NamedMed
End Property

Public Property Let NamedLow(ByVal plngNamedLow As Long)
    mtypCompendium.NamedLow = plngNamedLow
    mblnDirty = True
End Property

Public Property Get NamedLow() As Long
    NamedLow = mtypCompendium.NamedLow
End Property

Public Property Let NamedDim(ByVal plngNamedDim As Long)
    mtypCompendium.NamedDim = plngNamedDim
    mblnDirty = True
End Property

Public Property Get NamedDim() As Long
    NamedDim = mtypCompendium.NamedDim
End Property


' ************* PATHS *************


Public Property Get LitePath() As String
    ReQuery sePaths
    LitePath = PathLookup(mtypPath.CharacterBuilderLite, DefaultPath(), DefaultPath() & "\Save\Binary")
End Property

Public Property Let LitePath(ByVal pstrPath As String)
    If mtypPath.CharacterBuilderLite = pstrPath Then Exit Property
    mtypPath.CharacterBuilderLite = pstrPath
    SaveSettings False
End Property

Public Property Get CraftingPath() As String
    ReQuery sePaths
    CraftingPath = PathLookup(mtypPath.CannithCrafting, DefaultPath())
End Property

Public Property Let CraftingPath(ByVal pstrPath As String)
    If mtypPath.CannithCrafting = pstrPath Then Exit Property
    mtypPath.CannithCrafting = pstrPath
    SaveSettings False
End Property

Public Property Get CompendiumPath() As String
    ReQuery sePaths
    CompendiumPath = PathLookup(mtypPath.Compendium, DefaultPath())
End Property

Public Property Let CompendiumPath(ByVal pstrPath As String)
    If mtypPath.Compendium = pstrPath Then Exit Property
    mtypPath.Compendium = pstrPath
    SaveSettings False
End Property

Public Property Get RonPath() As String
    ReQuery sePaths
    RonPath = PathLookup(mtypPath.RonPath, xp.Folder.UserDocs)
End Property

Public Property Let RonPath(ByVal pstrPath As String)
    If mtypPath.RonPath = pstrPath Then Exit Property
    mtypPath.RonPath = pstrPath
    SaveSettings False
End Property

Public Property Get BuilderPath() As String
    ReQuery sePaths
    BuilderPath = PathLookup(mtypPath.BuilderPath, xp.Folder.UserDocs)
End Property

Public Property Let BuilderPath(ByVal pstrPath As String)
    If mtypPath.BuilderPath = pstrPath Then Exit Property
    mtypPath.BuilderPath = pstrPath
    SaveSettings False
End Property

Private Function PathLookup(pstrPath As String, pstrDefault As String, Optional pstrPrevent As String) As String
    If Len(pstrPath) = 0 Then
        PathLookup = pstrDefault
    ElseIf pstrPath = pstrPrevent Then
        PathLookup = pstrDefault
    ElseIf Not xp.Folder.Exists(pstrPath) Then
        PathLookup = pstrDefault
    Else
        PathLookup = pstrPath
    End If
End Function


' ************* CANNITH CRAFTING *************


Public Property Get DarkColors() As Boolean
    DarkColors = mblnDarkColors
End Property

Public Property Let DarkColors(ByVal pblnDarkColors As Boolean)
    mblnDarkColors = pblnDarkColors
End Property


' ************* COLORS *************

' Compendium has the unique requirement of needing to call MatchColors() before refreshing any form colors.
' Since MatchColors() edits Compendium-specific data structures, we call it from frmCompendium.RefreshColors,
' late-bound so other programs can ignore it. Then we skip frmCompendium during main loop since it's already done.
Public Sub RefreshAllColors()
    Dim frm As Form
    
    xp.Mouse = msAppWait
    If GetForm(frm, "frmCompendium") Then frm.RefreshColors
    For Each frm In Forms
        RefreshColors frm
        Select Case frm.Name
            Case "frmGearset", "frmSagas": frm.Redraw
        End Select
    Next
    Set frm = Nothing
    xp.Mouse = msNormal
End Sub

Public Sub RefreshColors(pfrm As Form)
    Dim enGroup As ColorGroupEnum
    Dim ctl As Control
    
    Select Case pfrm.Name
        Case "frmCompendium", "frmScaling" ' Do nothing
            Exit Sub
        Case "frmMain", "frmCompendium", "frmColors"
            pfrm.RefreshColors
            Exit Sub
    End Select
    pfrm.BackColor = GetColor(cgeWorkspace, cveBackground)
    For Each ctl In pfrm.Controls
        Select Case ctl.Tag
            Case "nav": ApplyColors ctl, cgeNavigation
            Case "ctl": ApplyColors ctl, cgeControls
            Case "out": ApplyColors ctl, cgeOutput
            Case "wrk": ApplyColors ctl, cgeWorkspace
            Case "tip"
            Case Else
                Select Case TypeName(ctl)
                    Case "Frame", "Label", "Line", "Shape", "userMenu"
                        ApplyColors ctl, cgeWorkspace
                    Case "CheckBox", "ComboBox", "ListBox", "TextBox", "userSpinner", "userTextbox"
                        ApplyColors ctl, cgeControls
                    Case "userHeader"
                        ApplyColors ctl, cgeNavigation
                    Case "userCheckBox"
                        ctl.RefreshColors cgeWorkspace
                    Case "userDetails", "userList", "userRaceCombo", "userQuests", "userStats", "userGearSlots"
                        ctl.RefreshColors
                    Case "userIcon", "userAugment", "userAugSlots", "userToolbar", "userLinkLists", "userTab"
                        ctl.RefreshColors
                    Case "userInfo"
                        ctl.TitleForeColor = GetColor(cgeWorkspace, cveText)
                        ctl.TitleBackColor = GetColor(cgeWorkspace, cveBackground)
                        ctl.TextColor = GetColor(cgeControls, cveText)
                        ctl.ErrorColor = GetColor(cgeControls, cveTextError)
                        ctl.LinkColor = GetColor(cgeControls, cveTextLink)
                        ctl.BackColor = GetColor(cgeControls, cveBackground)
                        ctl.Redraw
                    Case "PictureBox"
                        If pfrm.Name = "frmColors" And ctl.Name = "picColor" Then
                            ctl.BackColor = GetColor(pfrm.ColorGroup, ctl.Index)
                        ElseIf ctl.Name = "picGrid" Then
                            ctl.BackColor = GetColor(cgeWorkspace, cveBackground)
                            ctl.FillColor = GetColor(cgeControls, cveBackground)
                        ElseIf ctl.Name = "picPicker" Then
                            ctl.BackColor = glngActiveColor
                        Else
                            ApplyColors ctl, cgeWorkspace
                        End If
                End Select
        End Select
        Set ctl = Nothing
    Next
    Select Case pfrm.Name
        Case "frmSkills": pfrm.DrawGrid
        Case "frmHelp": pfrm.DrawText
        Case "frmOverview", "frmEnhancements", "frmFormat", "frmExport", "frmColorFile", "frmAugments": pfrm.RefreshColors
        Case "frmPatrons", "frmChallenges", "frmWilderness", "frmColorPreview": pfrm.ReDrawForm
    End Select
End Sub

Public Sub LoadColorFile(pstrFile As String)
    Dim enColorFile As ColorFileEnum
    Dim strText As String
    Dim i As Long
    
    If Not xp.File.Exists(pstrFile) Then
        Notice GetFileFromFilespec(pstrFile) & " not found"
        Exit Sub
    End If
    Select Case LCase$(GetExtFromFilespec(pstrFile))
        Case "screen": enColorFile = cfeScreen
        Case "output": enColorFile = cfeOutput
        Case Else: enColorFile = cfeAll
    End Select
    strText = xp.File.LoadToString(pstrFile)
    If Not LoadColors(strText, enColorFile) Then
        Notice GetFileFromFilespec(pstrFile) & " is not a valid color file"
        Exit Sub
    End If
    RefreshAllColors
    Select Case enColorFile
        Case cfeScreen: Me.ScreenColors = GetNameFromFilespec(pstrFile)
        Case cfeOutput: Me.OutputColors = GetNameFromFilespec(pstrFile)
    End Select
End Sub

Public Sub ImportColors(pstrText As String, penColorFile As ColorFileEnum)
    If LoadColors(pstrText, penColorFile) Then RefreshAllColors
End Sub

Public Function GetColor(ByVal penGroup As ColorGroupEnum, ByVal penColor As ColorValueEnum) As Long
    If penColor = -1 Then GetColor = -1 Else GetColor = mlngColor(penGroup, penColor)
End Function

Public Sub SetColor(ByVal penGroup As ColorGroupEnum, ByVal penColor As ColorValueEnum, plngColor)
    mlngColor(penGroup, penColor) = plngColor
End Sub

Public Function ApplyColors(pctl As Control, penGroup As ColorGroupEnum)
    Select Case TypeName(pctl)
        Case "Label"
            If Left$(pctl.Name, 3) = "lnk" Then
                pctl.ForeColor = mlngColor(penGroup, cveTextLink)
            Else
                pctl.ForeColor = mlngColor(penGroup, cveText)
            End If
            pctl.BackColor = mlngColor(penGroup, cveBackground)
        Case "TextBox", "ListBox", "ComboBox", "PictureBox", "Frame", "CheckBox", "OptionButton", "userTextbox"
            pctl.ForeColor = mlngColor(penGroup, cveText)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
        Case "userHeader"
            pctl.TextColor = mlngColor(penGroup, cveText)
            pctl.LinkColor = mlngColor(penGroup, cveTextLink)
            pctl.ErrorColor = mlngColor(penGroup, cveTextError)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.BorderColor = mlngColor(penGroup, cveBorderExterior)
        Case "userMenu"
            pctl.ForeColor = mlngColor(penGroup, cveTextLink)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.BorderColor = mlngColor(penGroup, cveBorderExterior)
        Case "userCheckBox"
            pctl.RefreshColors penGroup
        Case "userSpinner"
            pctl.ForeColor = mlngColor(penGroup, cveText)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.DisabledColor = mlngColor(penGroup, cveTextDim)
            pctl.BorderColor = mlngColor(penGroup, cveBorderExterior)
            pctl.BorderInterior = mlngColor(penGroup, cveBorderInterior)
        Case "Shape"
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.BorderColor = mlngColor(penGroup, cveBorderHighlight)
        Case "Line"
            If pctl.Name = "linHeader" Then pctl.BorderColor = GetColor(penGroup, cveText) Else pctl.BorderColor = GetColor(penGroup, cveBorderExterior)
    End Select
End Function

Public Property Get BackgroundStyle() As BackColorEnum
    Select Case mlngColor(cgeOutput, cveBackground)
        Case RGB(51, 51, 51): BackgroundStyle = bceOfficial
        Case RGB(102, 102, 102): BackgroundStyle = bceCodeTags
        Case RGB(17, 17, 17): BackgroundStyle = bceStylish
        Case RGB(34, 34, 34): BackgroundStyle = bceDefault
        Case Else: BackgroundStyle = bceCustom
    End Select
End Property

Public Property Let BackgroundStyle(penStyle As BackColorEnum)
    Select Case penStyle
        Case bceOfficial: Me.SetColor cgeOutput, cveBackground, RGB(51, 51, 51)
        Case bceCodeTags: Me.SetColor cgeOutput, cveBackground, RGB(102, 102, 102)
        Case bceStylish: Me.SetColor cgeOutput, cveBackground, RGB(17, 17, 17)
        Case bceDefault: Me.SetColor cgeOutput, cveBackground, RGB(34, 34, 34)
    End Select
End Property


' ************* OUTPUT FORMAT *************


Public Function SetOutputFormat(pstrFormat As String) As Boolean
    If Len(pstrFormat) = 0 Then Exit Function
    If InitFormat(pstrFormat, mtypOutput) Then
        mtypCharBuild.BBCodes = pstrFormat
        SetOutputFormat = True
        ReadOutputFormat
    End If
End Function

Public Sub ReadOutputFormat()
    Dim frm As Form
    
    If Not GetForm(frm, "frmFormat") Then Exit Sub
    With mtypOutput
        frm.usrchkReddit.Value = .Reddit
        frm.usrchkBBCodes.Value = .Codes
        frm.usrchkUseDots.Value = .Dots
        frm.usrchkColoredText.Value = .TextColors
        frm.usrchkBold.Value = .Bold.Enabled
        frm.txtBoldOpen.Text = .Bold.OpenTag
        frm.txtBoldClose.Text = .Bold.CloseTag
        frm.usrchkUnderline.Value = .Underline.Enabled
        frm.txtUnderlineOpen.Text = .Underline.OpenTag
        frm.txtUnderlineClose.Text = .Underline.CloseTag
        frm.usrchkFixed.Value = .Fixed.Enabled
        frm.txtFixedOpen.Text = .Fixed.OpenTag
        frm.txtFixedClose.Text = .Fixed.CloseTag
        frm.usrchkLists.Value = .Lists.Enabled
        frm.txtBulletOpen.Text = .Lists.BulletOpen
        frm.txtBulletClose.Text = .Lists.BulletClose
        frm.txtNumberedOpen.Text = .Lists.NumberedOpen
        frm.txtNumberedClose.Text = .Lists.NumberedClose
        frm.usrchkColor.Value = .Color.Enabled
        frm.txtColorOpen.Text = .Color.OpenTag
        frm.txtColorClose.Text = .Color.CloseTag
        frm.usrchkWrapper.Value = .Wrapper.Enabled
        frm.txtWrapperOpen.Text = .Wrapper.OpenTag
        frm.txtWrapperClose.Text = .Wrapper.CloseTag
    End With
    frm.Dirty = False
    frm.EnableControls
End Sub

Public Sub WriteOutputFormat()
    Dim frm As Form
    
    If Not GetForm(frm, "frmFormat") Then Exit Sub
    With mtypOutput
        .Reddit = frm.usrchkReddit.Value
        .Codes = frm.usrchkBBCodes.Value
        .Dots = frm.usrchkUseDots.Value
        .TextColors = frm.usrchkColoredText.Value
        .Bold.Enabled = frm.usrchkBold.Value
        .Bold.OpenTag = frm.txtBoldOpen.Text
        .Bold.CloseTag = frm.txtBoldClose.Text
        .Underline.Enabled = frm.usrchkUnderline.Value
        .Underline.OpenTag = frm.txtUnderlineOpen.Text
        .Underline.CloseTag = frm.txtUnderlineClose.Text
        .Fixed.Enabled = frm.usrchkFixed.Value
        .Fixed.OpenTag = frm.txtFixedOpen.Text
        .Fixed.CloseTag = frm.txtFixedClose.Text
        .Lists.Enabled = frm.usrchkLists.Value
        .Lists.BulletOpen = frm.txtBulletOpen.Text
        .Lists.BulletClose = frm.txtBulletClose.Text
        .Lists.NumberedOpen = frm.txtNumberedOpen.Text
        .Lists.NumberedClose = frm.txtNumberedClose.Text
        .Color.Enabled = frm.usrchkColor.Value
        .Color.OpenTag = frm.txtColorOpen.Text
        .Color.CloseTag = frm.txtColorClose.Text
        .Wrapper.Enabled = frm.usrchkWrapper.Value
        .Wrapper.OpenTag = frm.txtWrapperOpen.Text
        .Wrapper.CloseTag = frm.txtWrapperClose.Text
    End With
End Sub

Public Function IdentifyOutputFormat() As String
    Dim typCompare As OutputType
    Dim strName() As String
    Dim i As Long
    
    mtypCharBuild.BBCodes = vbNullString
    GetFormatNames strName
    For i = 0 To UBound(strName)
        If InitFormat(strName(i), typCompare) Then
            If CompareFormats(mtypOutput, typCompare) Then
                IdentifyOutputFormat = strName(i)
                mtypCharBuild.BBCodes = strName(i)
                Exit Function
            End If
        End If
    Next
End Function

Public Sub GetFormatNames(pstrName() As String)
    Dim strFile As String
    Dim lngMax As Long
    Dim i As Long
    
    lngMax = 4
    ReDim pstrName(lngMax)
    pstrName(0) = "DDO Forums"
    pstrName(1) = "Code Tags"
    pstrName(2) = "Plain Text"
    i = 2
    strFile = Dir(App.Path & "\Settings\*.bbcodes")
    Do While Len(strFile)
        Select Case LCase$(strFile)
            Case "ddo forums.bbcodes", "code tags.bbcodes", "plain text.bbcodes"
            Case Else
                i = i + 1
                If i > lngMax Then
                    lngMax = lngMax + 1
                    ReDim Preserve pstrName(lngMax)
                End If
                pstrName(i) = GetNameFromFilespec(strFile)
        End Select
        strFile = Dir
    Loop
End Sub

Private Function CompareFormats(ptyp1 As OutputType, ptyp2 As OutputType) As Boolean
    If ptyp1.Codes <> ptyp2.Codes Then Exit Function
    If ptyp1.Dots <> ptyp2.Dots Then Exit Function
    If ptyp1.TextColors <> ptyp2.TextColors Then Exit Function
    If Not CompareTags(ptyp1.Bold, ptyp2.Bold) Then Exit Function
    If Not CompareTags(ptyp1.Underline, ptyp2.Underline) Then Exit Function
    If Not CompareTags(ptyp1.Fixed, ptyp2.Fixed) Then Exit Function
    If Not CompareTags(ptyp1.Color, ptyp2.Color) Then Exit Function
    If Not CompareTags(ptyp1.Wrapper, ptyp2.Wrapper) Then Exit Function
    If ptyp1.Lists.Enabled <> ptyp2.Lists.Enabled Then Exit Function
    If ptyp1.Lists.BulletOpen <> ptyp2.Lists.BulletOpen Then Exit Function
    If ptyp1.Lists.BulletClose <> ptyp2.Lists.BulletClose Then Exit Function
    If ptyp1.Lists.NumberedOpen <> ptyp2.Lists.NumberedOpen Then Exit Function
    If ptyp1.Lists.NumberedClose <> ptyp2.Lists.NumberedClose Then Exit Function
    CompareFormats = True
End Function

Private Function CompareTags(ptypTag1 As TagsType, ptypTag2 As TagsType) As Boolean
    If ptypTag1.Enabled <> ptypTag2.Enabled Then Exit Function
    If ptypTag1.OpenTag <> ptypTag2.OpenTag Then Exit Function
    If ptypTag1.CloseTag <> ptypTag2.CloseTag Then Exit Function
    CompareTags = True
End Function

Public Sub GetOutputStyle(pblnCodes As Boolean, pblnDots As Boolean, pblnTextColors As Boolean)
    With mtypOutput
        pblnCodes = .Codes
        pblnDots = .Dots
        pblnTextColors = .TextColors
    End With
End Sub

Public Sub GetOutputBold(pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    GetTags mtypOutput.Bold, pblnEnabled, pstrOpen, pstrClose
End Sub

Public Sub GetOutputUnderline(pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    GetTags mtypOutput.Underline, pblnEnabled, pstrOpen, pstrClose
End Sub

Public Sub GetOutputFixed(pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    GetTags mtypOutput.Fixed, pblnEnabled, pstrOpen, pstrClose
End Sub

Public Sub GetOutputColor(pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    GetTags mtypOutput.Color, pblnEnabled, pstrOpen, pstrClose
End Sub

Public Sub GetOutputWrapper(pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    GetTags mtypOutput.Wrapper, pblnEnabled, pstrOpen, pstrClose
End Sub

Private Sub GetTags(ptypTags As TagsType, pblnEnabled As Boolean, pstrOpen As String, pstrClose As String)
    pblnEnabled = ptypTags.Enabled
    pstrOpen = ptypTags.OpenTag
    pstrClose = ptypTags.CloseTag
End Sub

Public Sub GetOutputLists(pblnEnabled As Boolean, pstrBulletOpen As String, pstrBulletClose As String, pstrNumberedOpen As String, pstrNumberedClose As String)
    With mtypOutput.Lists
        pblnEnabled = .Enabled
        pstrBulletOpen = .BulletOpen
        pstrBulletClose = .BulletClose
        pstrNumberedOpen = .NumberedOpen
        pstrNumberedClose = .NumberedClose
    End With
End Sub

Public Sub SaveFormat(pstrFile As String)
    Dim strRaw As String
    
    strRaw = MakeOutputSettings()
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
    xp.File.SaveStringAs pstrFile, strRaw
End Sub


' ************* MRU LIST *************


Public Property Let MRUCount(ByVal plngMRUCount As Long)
    With mtypCharBuild
        If plngMRUCount < 1 Then
            .MRUCount = 0
            Erase .MRUList
        Else
            If plngMRUCount > 9 Then plngMRUCount = 9
            .MRUCount = plngMRUCount
            ReDim Preserve .MRUList(1 To .MRUCount)
        End If
    End With
End Property

Public Property Get MRUCount() As Long
    MRUCount = mtypCharBuild.MRUCount
End Property

Public Sub ShowMRU()
    Dim blnShow As Boolean
    Dim frm As Form
    Dim i As Long
    
    If Not GetForm(frm, "frmMain") Then Exit Sub
    With frm
        With .mnuFile(StartMRU)
            If mtypCharBuild.MRUCount > 0 Then
                If Len(mtypCharBuild.MRUList(1)) <> 0 Then .Visible = True Else .Visible = False
            Else
                .Visible = False
            End If
        End With
        For i = 1 To mtypCharBuild.MRUCount
            With .mnuFile(StartMRU + i)
                If mtypCharBuild.MRUCount = 0 Then blnShow = False Else blnShow = (Len(mtypCharBuild.MRUList(i)) <> 0)
                If blnShow Then
                    .Caption = "&" & i & ". " & Replace(GetNameFromFilespec(mtypCharBuild.MRUList(i)), "&", "&&")
                    .Visible = True
                Else
                    .Caption = "MRU" & i
                    .Visible = False
                End If
            End With
        Next
        For i = mtypCharBuild.MRUCount + 1 To 9
            With .mnuFile(StartMRU + i)
                .Caption = "MRU" & i
                .Visible = False
            End With
        Next
    End With
End Sub

Public Sub AddMRU(pstrFileName As String)
    Dim strTemp() As String
    Dim lngCurrent As Long
    Dim strDisplay As String
    Dim blnFound As Boolean
    Dim i As Long
    
    With mtypCharBuild
        If Len(pstrFileName) And .MRUCount > 0 Then
            For i = 1 To .MRUCount
                If .MRUList(i) = pstrFileName Then
                    blnFound = True
                    Exit For
                End If
            Next
            If blnFound Or LCase$(GetExtFromFilespec(pstrFileName)) = "build" Then
                strDisplay = GetNameFromFilespec(pstrFileName)
                strTemp = .MRUList
                .MRUList(1) = pstrFileName
                lngCurrent = 1
                For i = 1 To .MRUCount
                    If Len(strTemp(i)) <> 0 Then
                        If GetNameFromFilespec(strTemp(i)) <> strDisplay Then
                            lngCurrent = lngCurrent + 1
                            If lngCurrent > .MRUCount Then Exit For
                            .MRUList(lngCurrent) = strTemp(i)
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Function GetMRUTarget(ByVal pstrCaption As String) As String
    Dim i As Long
    
    pstrCaption = Mid(pstrCaption, 4)
    For i = 1 To mtypCharBuild.MRUCount
        If GetNameFromFilespec(mtypCharBuild.MRUList(i)) = pstrCaption Then
            GetMRUTarget = CheckExtension(mtypCharBuild.MRUList(i))
            Exit Function
        End If
    Next
End Function

Public Function CheckExtension(pstrFile As String) As String
    Dim strFile As String
    Dim lngPos As Long
    
    CheckExtension = pstrFile
    If xp.File.Exists(pstrFile) Then Exit Function
    If LCase$(Right$(pstrFile, 4)) <> ".bld" Then Exit Function
    strFile = Left$(pstrFile, Len(pstrFile) - 3) & "build"
    If xp.File.Exists(strFile) Then CheckExtension = strFile
End Function

Public Sub RemoveMRU(ByVal pstrCaption As String)
    Dim i As Long
    
    pstrCaption = Mid(pstrCaption, 4)
    With mtypCharBuild
        For i = 1 To .MRUCount
            If GetNameFromFilespec(.MRUList(i)) = pstrCaption Then Exit For
        Next
        If i <= .MRUCount Then
            For i = i To .MRUCount - 1
                .MRUList(i) = .MRUList(i + 1)
            Next
            .MRUList(.MRUCount) = vbNullString
        End If
    End With
    ShowMRU
End Sub

