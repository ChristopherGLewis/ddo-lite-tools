Attribute VB_Name = "basOutput"
' Written by Ellis Dee
Option Explicit

Private Const BulletPoint As String = "•"

Public Enum ValidEnum
    veSkip
    veEmpty
    veErrors
    veIncomplete
    veComplete
End Enum

Private Enum FitEnum
    feAlwaysFits
    feNeverFits
    feFitsWithoutScroll
End Enum

Private Enum SpecialCodeEnum
    sceNone
    sceCourier ' Adds both [Font=Courier] and [/Font] tags
    sceCourierBegin
    sceCourierEnd
End Enum

Private Type ListDetailType
    Numbered As Boolean
    Count As Long
End Type

Private Type ListType
    List() As ListDetailType
    Lists As Long
End Type

Public Type FeatOutputDetailType
    Level As Long
    Source As String
    FeatName As String
    Alternate() As String
    Alternates As Long
    Replaces As String
    Channel As FeatChannelEnum
    ' FeatType and Slot are used in frmImport
    FeatType As BuildFeatTypeEnum
    Slot As Long
End Type

Public Type FeatOutputType
    Feat() As FeatOutputDetailType
    Feats As Long
End Type

Private Type OutputTreeType
    ID As Long
    AP As Long
End Type

Public Type GuideOutputEnhancementType
    GuideTreeID As Long
    Tree As String
    Color As ColorValueEnum
    Tier As Long
    Enhancement As String
    Display As String
End Type

Public Type GuideOutputBlockType
    Display As String
    Color As Long
End Type

Public Type GuideOutputBuildTreeType
    GuideTreeID As Long
    Color As Long
    TreeName As String
    Ability() As BuildAbilityType
    Abilities As Long
    Tier() As String
    Tiers As Long
End Type

Public Type GuideOutputType
    Block() As GuideOutputBlockType
    Blocks As Long
    Enhancement() As GuideOutputEnhancementType
    Enhancements As Long
    Reset As String
    BuildTree() As GuideOutputBuildTreeType
    BuildTrees As Long
End Type

Private menOutput As OutputEnum
Private menRemember As OutputEnum
Private mblnDisplay As Boolean

Private mstrJoin() As String
Private mlngCurrent As Long
Private mlngMax As Long

Private mstrHideColor As String

Private mlngMaxWidth As Long
Private mlngMaxHeight As Long
Private mlngListIndent As Long ' Indent size in Twips for display only

Private mlngLineHeight As Long

Private mtypFeatOutput As FeatOutputType
Private mtypGuideOutput() As GuideOutputType

' Output format
Private mblnReddit As Boolean
Private mlngSectionLines As Long
Private mblnCodes As Boolean
Private mblnDots As Boolean
Private mblnTextColors As Boolean
Private mblnBold As Boolean
Private mstrBoldOpen As String
Private mstrBoldClose As String
Private mblnUnderline As Boolean
Private mstrUnderlineOpen As String
Private mstrUnderlineClose As String
Private mblnFixed As Boolean
Private mstrFixedOpen As String
Private mstrFixedClose As String
Private mblnLists As Boolean
Private mstrBulletOpen As String
Private mstrBulletClose As String
Private mstrNumberedOpen As String
Private mstrNumberedClose As String
Private mblnColor As Boolean
Private mstrColorOpen As String
Private mstrColorClose As String
Private mblnWrapper As Boolean
Private mstrWrapperOpen As String
Private mstrWrapperClose As String
Private mblnColorOverride As Boolean
' Display format
Private mstrDefaultFont As String
Private msngDefaultSize As Single
Private mstrCourierFont As String
Private msngCourierSize As Single

' Stack structure for handling nested [List]s
Private mtypList As ListType

' Max Destiny. 24 + PDP
Const MaxDestinyAP = 40  '4 per epic level


Public Function GenerateOutput(penOutput As OutputEnum) As String
    Dim strRaw As String
    Dim i As Long
    
    If Not BuildIsOpen() Then
        frmMain.picBuild.Cls
        frmMain.scrollHorizontal.Visible = False
        frmMain.scrollVertical.Visible = False
        frmMain.CanDrag = False
        Exit Function
    End If
    ' Initialize
    WriteOutputFormat
    If penOutput = oeRemember Then menOutput = menRemember Else menOutput = penOutput
    If menOutput <> oeExport Then menRemember = menOutput
    If menOutput = oeAll Then xp.Mouse = msSystemWait
    mblnDisplay = (menOutput <> oeExport)
    'Check for Reddit display
    If cfg.OutputReddit = True And mblnDisplay = False Then
        mblnReddit = True
    Else
        mblnReddit = False
    End If
    If mblnReddit Then mlngSectionLines = 1 Else mlngSectionLines = 2
    If mblnDisplay And frmMain.WindowState = vbMinimized Then
        cfg.OutputRefresh = True
        Exit Function
    End If
    Select Case penOutput
        Case oeAll, oeExport: Screen.MousePointer = vbHourglass
    End Select
    mlngCurrent = 0
    ReDim mstrJoin(255) ' Enough that we never need to worry about increasing it
    mlngLineHeight = frmMain.TextHeight("X")
    If mblnDisplay Then
        mlngMaxWidth = 0
        mlngMaxHeight = 0
        frmMain.picBuild.FontName = mstrDefaultFont
        frmMain.picBuild.FontSize = msngDefaultSize
        frmMain.picBuild.FontBold = False
        frmMain.picBuild.Cls
        frmMain.picBuild.Move 0, 0, PixelX * 1000, PixelY * 1000
        frmMain.picBuild.ForeColor = cfg.GetColor(cgeOutput, cveText)
        frmMain.picBuild.BackColor = cfg.GetColor(cgeOutput, cveBackground)
        mlngListIndent = frmMain.picBuild.TextWidth("XXXXX")
    ElseIf mblnCodes Then
        If mblnColor Then mstrHideColor = Replace(mstrColorOpen, "$", xp.ColorToHex(cfg.GetColor(cgeOutput, cveBackground)))
    End If
    ' Create each section
    OutputHeader
    OutputClassSplit
    OutputStats
    OutputSkills
    OutputFeats
    OutputSpells
    OutputEnhancements
    OutputDestinies
    ' Finish
    If mblnDisplay Then
        ResizeOutput
    Else
        TrimBlanks
        If mlngCurrent > -1 Then
            ReDim Preserve mstrJoin(mlngCurrent)
            For i = 0 To mlngCurrent
                mstrJoin(i) = RTrim$(mstrJoin(i))
                If mblnDots Then AddDots mstrJoin(i)
            Next
            If mblnCodes And mblnWrapper Then
                mstrJoin(0) = mstrWrapperOpen & mstrJoin(0)
                mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & mstrWrapperClose
            End If
            mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & vbNewLine
            strRaw = Join(mstrJoin, vbNewLine)
            GenerateOutput = strRaw
        End If
    End If
    Select Case penOutput
        Case oeAll, oeExport: Screen.MousePointer = vbNormal
    End Select
    If menOutput = oeAll Then xp.Mouse = msNormal
End Function

Private Sub TrimBlanks()
    Do While LastBlank()
        mlngCurrent = mlngCurrent - 1
        If mlngCurrent = 0 Then Exit Do
    Loop
End Sub

Private Function LastBlank() As Boolean
    If mlngCurrent = -1 Then Exit Function
    Do While Right$(mstrJoin(mlngCurrent), 2) = vbNewLine
        mstrJoin(mlngCurrent) = RTrim$(Left$(mstrJoin(mlngCurrent), Len(mstrJoin(mlngCurrent)) - 2))
    Loop
    If Len(RTrim$(mstrJoin(mlngCurrent))) = 0 Then LastBlank = True
End Function

Private Sub WriteOutputFormat()
    cfg.GetOutputStyle mblnCodes, mblnDots, mblnTextColors
    cfg.GetOutputBold mblnBold, mstrBoldOpen, mstrBoldClose
    cfg.GetOutputUnderline mblnUnderline, mstrUnderlineOpen, mstrUnderlineClose
    cfg.GetOutputLists mblnLists, mstrBulletOpen, mstrBulletClose, mstrNumberedOpen, mstrNumberedClose
    cfg.GetOutputFixed mblnFixed, mstrFixedOpen, mstrFixedClose
    cfg.GetOutputColor mblnColor, mstrColorOpen, mstrColorClose
    cfg.GetOutputWrapper mblnWrapper, mstrWrapperOpen, mstrWrapperClose
End Sub

Public Sub InitFonts(ppic As PictureBox)
    mstrDefaultFont = ppic.FontName
    msngDefaultSize = ppic.FontSize
    If Not SetFont(ppic, "Courier New") Then
        ppic.FontName = mstrDefaultFont
        ppic.FontSize = msngDefaultSize
        If Not SetFont(ppic, "Courier") Then
            ppic.FontName = mstrDefaultFont
            ppic.FontSize = msngDefaultSize
            frmMain.Visible = True
            Notice "Courier font not found. Please download and install either Courier New or Courier so that output displays properly on screen."
        End If
    End If
    mstrCourierFont = ppic.FontName
    msngCourierSize = ppic.FontSize
    ppic.FontName = mstrDefaultFont
    ppic.FontSize = msngDefaultSize
End Sub

Private Function SetFont(ppic As PictureBox, pstrFontName As String) As Boolean
On Error Resume Next
    ppic.FontName = pstrFontName
    If Err.Number = 0 Then SetFont = True
End Function


' ************* HEADER *************


Private Sub OutputHeader()
    Dim strText As String
    Dim typClassSplit() As ClassSplitType
    Dim lngClasses As Long
    Dim strLevels As String
    Dim strClassNames As String
    Dim lngEpic As Long
    Dim strPipe As String
    Dim i As Long
    
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeOverview) Then Exit Sub
    ' Name
    If mblnReddit Then strPipe = "|"
    If Len(build.BuildName) Then
        OutputText strPipe & build.BuildName & strPipe, , , , Not mblnReddit
        If mblnReddit Then OutputText "|:--|:--"
    End If
    ' Class split
    lngClasses = GetClassSplit(typClassSplit)
    If lngClasses <> 0 Then
        strLevels = typClassSplit(0).Levels
        strClassNames = typClassSplit(0).ClassName
        For i = 1 To lngClasses - 1
            strLevels = strLevels & "/" & typClassSplit(i).Levels
            strClassNames = strClassNames & "/" & typClassSplit(i).ClassName
        Next
        If lngClasses = 1 Then strText = strClassNames & " " & strLevels Else strText = strLevels & " " & strClassNames
        If build.MaxLevels <> MAX_LEVEL Then
            lngEpic = build.MaxLevels - 20
            If lngEpic > 0 Then strText = strText & ", Epic " & lngEpic
        End If
        If mblnReddit Then
            OutputText "|" & strText & "|"
            If mblnReddit = True And Len(build.BuildName) = 0 Then OutputText "|:--|:--"
        Else
            OutputText strText
        End If
    End If
    ' Alignment and Race
    strText = strPipe & Trim$(GetAlignmentName(build.Alignment) & " " & GetRaceName(build.Race)) & strPipe
    If Len(strText) Then OutputText strText
    If frmMain.picBuild.CurrentY <> 0 Then BlankLine mlngSectionLines
End Sub


' ************* CLASS SPLIT *************


Private Sub OutputClassSplit()
    Dim typClassSplit() As ClassSplitType
    Dim lngClasses As Long
    Dim strText As String
    Dim strItem As String
    Dim lngLevel As Long
    Dim lngLastLevel As Long
    Dim lngColor() As Long
    Dim lngColumns As Long
    Dim i As Long
    Dim j As Long
    Dim jMax As Long
    
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeOverview) Then Exit Sub
    If Not ValidClassSplit() Then Exit Sub
    lngClasses = GetClassSplit(typClassSplit)
    If lngClasses < 2 Then Exit Sub
    ReDim lngColor(ceClasses - 1)
    For i = 0 To lngClasses - 1
        lngColor(typClassSplit(i).ClassID) = typClassSplit(i).Color
    Next
    ' Title
    OutputTitle "Level Order"
    BlankLine
    lngLastLevel = HeroicLevels()
    If mblnReddit Then
        lngColumns = 1 + (build.MaxLevels - 1) \ 5
        If lngColumns > 4 Then lngColumns = 4
        OutputText Left("|||||", lngColumns + 1)
        OutputText Left("|:--|:--|:--|:--|:--", (lngColumns + 1) * 4)
    Else
        OutputText vbNullString, False, , sceCourierBegin
    End If
    For i = 1 To 5
        For j = 0 To 15 Step 5
            lngLevel = i + j
            If build.MaxLevels >= lngLevel Then
                If mblnReddit Then
                    OutputText "|" & lngLevel & ") " & GetClassName(build.Class(lngLevel)), False
                Else
                    ' Level (always standard color)
                    If j = 0 Then strText = CStr(lngLevel) Else strText = AlignText(CStr(lngLevel), 2, vbRightJustify)
                    OutputText strText & ". ", False
                    ' Class (use class color if splash)
                    strText = GetClassName(build.Class(lngLevel))
                    If lngLevel + 5 <= lngLastLevel Then strText = AlignText(strText, 15, vbLeftJustify)
                    OutputText strText, False, cfg.GetColor(cgeOutput, lngColor(build.Class(lngLevel)))
                End If
            End If
        Next
        If i = 5 Or i >= lngLastLevel Then
            If mblnReddit Then OutputText "|" Else OutputText vbNullString, , , sceCourierEnd
            Exit For
        Else
            If mblnReddit Then OutputText "|" Else OutputText vbNullString
        End If
    Next
    BlankLine mlngSectionLines
End Sub

Private Function ValidClassSplit() As Boolean
    If mblnDisplay And menOutput <> oeAll Then
        Select Case menOutput
            Case oeOverview, oeStats, oeSkills, oeFeats, oeSpells, oeEnhancements: ValidClassSplit = True
        End Select
        Exit Function
    End If
    ValidClassSplit = True
End Function


' ************* STATS *************


Private Sub OutputStats()
    Dim blnShow(5) As Boolean
    Dim lngStat As Long
    Dim lngPoints As Long
    Dim strText As String
    Dim strItem As String
    Dim lngColumns As Long
    Dim i As Long
    
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeStats) Then Exit Sub
    If Not ValidStats() Then Exit Sub

    For i = 0 To 3
        blnShow(i) = (build.IncludePoints(i) = 1 And build.StatPoints(i, 0) = GetBuildPoints(i))
'        If build.RacialAP > 0 And i < 3 Then blnShow(i) = False
        If blnShow(i) Then lngColumns = lngColumns + 1
    Next
    For i = 1 To MAX_STATS
        If CapTome(build.Tome(i)) > 0 Then
            blnShow(4) = True
            lngColumns = lngColumns + 1
            Exit For
        End If
    Next
    For i = 1 To MAX_LEVELUPS
        If build.Levelups(i) <> aeAny Then
            If CapLevelup(i * 4) > 0 Then
                blnShow(5) = True
                lngColumns = lngColumns + 1
            End If
            Exit For
        End If
    Next
    If lngColumns = 0 Then Exit Sub
    ' Title
    OutputTitle "Stats"
    ' Reddit?
    If mblnReddit Then
        OutputStatsReddit blnShow, lngColumns + 2
        Exit Sub
    End If
    ' Header
    OutputText Space$(15), False, , sceCourierBegin
    For lngPoints = 0 To 3
        If blnShow(lngPoints) Then OutputText GetBuildPoints(lngPoints) & "pt     ", False, StatColumnColor(lngPoints), , StatColumnBold(lngPoints)
    Next
    If blnShow(4) Then OutputText "Tome     ", False
    If blnShow(5) Then OutputText "Level Up", False
    OutputText vbNullString
    ' Underlines
    OutputText Space$(15), False
    For lngPoints = 0 To 4
        If blnShow(lngPoints) Then OutputText "----     ", False, StatColumnColor(lngPoints), , StatColumnBold(lngPoints)
    Next
    If blnShow(5) Then OutputText "--------", False
    OutputText vbNullString
    ' Body
    For lngStat = 1 To MAX_STATS
        OutputText Left$(GetStatName(lngStat) & Space$(15), 15), False
        ' Build Points
        For lngPoints = 0 To 3
            If blnShow(lngPoints) Then
                strItem = CStr(CalculateBaseStat(db.Race(build.Race).Stats(lngStat), build.StatPoints(lngPoints, lngStat)))
                OutputText AlignText(strItem, 3, vbRightJustify) & "      ", False, StatColumnColor(lngPoints), , StatColumnBold(lngPoints)
            End If
        Next
        ' Tomes
        If blnShow(4) Then
            If build.Tome(lngStat) <> 0 Then
                OutputText " +" & CapTome(build.Tome(lngStat)) & "      ", False
            Else
                OutputText "         ", False
            End If
        End If
        ' Levelups
        If build.Levelups(lngStat) <> aeAny And build.MaxLevels >= lngStat * 4 Then
            strItem = lngStat * 4 & ": " & UCase$(GetStatName(build.Levelups(lngStat), True))
            OutputText AlignText(strItem, 7, vbRightJustify), False, LevelupColor(lngStat)
        End If
        ' Output this line
        'If lngStat = 6 And Not blnShowMoreLevelups Then
        '    OutputText vbNullString, , , sceCourierEnd
        'Else
            OutputText vbNullString
        'End If
            
    Next
    'Loop for 25+ to get additional level ups
    For i = 28 To build.MaxLevels Step 4
        ' Even though this line is all blanks (dots), process it the same as the previous lines.
        ' This is because when displaying, bold text is (annoyingly) wider than regular text
        ' despite Courier being a fixed-width font. Fixed width? Lies!!!!!
        OutputText Space(15), False
        For lngStat = 0 To 4
            If blnShow(lngStat) Then
                OutputText Space(9), False, , , (build.BuildPoints = lngStat)
            End If
        Next
        strText = i & ": " & UCase$(GetStatName(build.Levelups(i / 4), True))
        OutputText strText, , LevelupColor(i / 4) ' , sceCourierEnd
        
    Next
    BlankLine
End Sub

Private Sub OutputStatsReddit(pblnShow() As Boolean, plngColumns As Long)
    Dim strText As String
    Dim strBold As String
    Dim lngStat As Long
    Dim strItem As String
    Dim lngPoints As Long
    Dim lngMainStat As Long
    Dim i As Long
    
    ' Header
    BlankLine
    strText = "||"
    For lngPoints = 0 To 3
        If pblnShow(lngPoints) Then
            If build.BuildPoints = lngPoints Then strBold = "**" Else strBold = vbNullString
            strText = strText & strBold & GetBuildPoints(lngPoints) & "pt" & strBold & "|"
        End If
    Next
    If pblnShow(4) Then strText = strText & "Tome|"
    If pblnShow(5) Then strText = strText & "Level Up|"
    OutputText strText
    ' Alignment
    strText = "|:--"
    For i = 0 To 3
        If pblnShow(i) Then strText = strText & "|--:"
    Next
    If pblnShow(4) Then strText = strText & "|:--:"
    If pblnShow(5) Then strText = strText & "|--:"
    strText = strText & "|:--"
    OutputText strText
    ' Rows
    lngMainStat = IdentifyMainStat()
    For lngStat = 1 To 6
        strText = "|" & GetStatName(lngStat) & "|"
        ' Build Points
        For lngPoints = 0 To 3
            If pblnShow(lngPoints) Then
                If build.BuildPoints = lngPoints Then strBold = "**" Else strBold = vbNullString
                strItem = CStr(CalculateBaseStat(db.Race(build.Race).Stats(lngStat), build.StatPoints(lngPoints, lngStat)))
                strText = strText & strBold & strItem & strBold & "|"
            End If
        Next
        ' Tomes
        If pblnShow(4) Then
            If build.Tome(lngStat) <> 0 Then strText = strText & "+" & CapTome(build.Tome(lngStat))
            strText = strText & "|"
        End If
        ' Levelups
        If pblnShow(5) Then
            If build.Levelups(lngStat) <> aeAny And build.MaxLevels >= lngStat * 4 Then
                If build.Levelups(lngStat) = lngMainStat Then
                    strBold = vbNullString
                Else
                    strBold = "**"
                End If
                strText = strText & strBold & lngStat * 4 & ": " & UCase$(GetStatName(build.Levelups(lngStat), True)) & strBold
            End If
            strText = strText & "|"
        End If
        ' Output this line
        OutputText strText
    Next
    ' output levelups more that stats
    'Loop for 25+ to get additional level ups
    For i = 28 To build.MaxLevels Step 4
        If build.Levelups(i / 4) = lngMainStat Then
            strBold = vbNullString
        Else
            strBold = "**"
        End If
        strText = strBold & i & ": " & UCase$(GetStatName(build.Levelups(i / 4), True)) & strBold
        OutputText Left$("||||||||", plngColumns - 1) & strText & "|"
    Next
    BlankLine
End Sub

Private Function IdentifyMainStat() As Long
    Dim lngStat(MAX_STATS) As Long
    Dim lngHigh As Long
    Dim lngMain As Long
    Dim i As Long
    
    For i = 1 To MAX_LEVELUPS
        lngStat(build.Levelups(i)) = lngStat(build.Levelups(i)) + 1
    Next
    For i = 1 To MAX_STATS
        If lngHigh < lngStat(i) Then
            lngHigh = lngStat(i)
            lngMain = i
        End If
    Next
    IdentifyMainStat = lngMain
End Function

Private Function StatColumnColor(plngColumn As Long) As Long
    If build.BuildPoints = plngColumn Or plngColumn > 3 Then
        StatColumnColor = cfg.GetColor(cgeOutput, cveText)
    Else
        StatColumnColor = cfg.GetColor(cgeOutput, cveTextDim)
    End If
End Function

Private Function StatColumnBold(plngColumn As Long) As Boolean
    If build.BuildPoints = plngColumn Then StatColumnBold = Not (mblnTextColors And mblnColor)
End Function

Private Function ValidStats() As Boolean
    If mblnDisplay And menOutput <> oeAll Then
        Select Case menOutput
            Case oeOverview, oeStats, oeSkills, oeFeats: ValidStats = True
        End Select
        Exit Function
    End If
    ValidStats = True
End Function

Private Function CapTome(pbytTome As Byte) As Long
    Dim lngCap As Long
    
    lngCap = TomeLevel(build.MaxLevels)
    If pbytTome > lngCap Then CapTome = lngCap Else CapTome = pbytTome
End Function

Private Function CapLevelup(pbytLevelup As Byte) As Long
    Dim lngCap As Long
    
    lngCap = build.MaxLevels \ 4
    If pbytLevelup > lngCap Then CapLevelup = lngCap Else CapLevelup = pbytLevelup
End Function

Private Function LevelupColor(plngIndex As Long) As Long
    If build.Levelups(0) <> aeAny Then
        LevelupColor = -1
    Else
        Select Case build.Levelups(plngIndex)
            Case aeStr: LevelupColor = cfg.GetColor(cgeOutput, cveOrange)
            Case aeDex: LevelupColor = cfg.GetColor(cgeOutput, cveRed)
            Case aeCon: LevelupColor = cfg.GetColor(cgeOutput, cveGreen)
            Case aeInt: LevelupColor = cfg.GetColor(cgeOutput, cveBlue)
            Case aeWis: LevelupColor = cfg.GetColor(cgeOutput, cveYellow)
            Case aeCha: LevelupColor = cfg.GetColor(cgeOutput, cvePurple)
        End Select
    End If
End Function


' ************* SKILLS *************


Private Sub OutputSkills()
    Dim blnClassRow As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngSkill As Long
    Dim strError As String
    Dim strItem As String
    Dim enFont As SpecialCodeEnum
    Dim lngHeroicLevels As Long
    Dim strDashLine As String
    Dim enValid As ValidEnum
    Dim i As Long

    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeSkills) Then Exit Sub
    enValid = ValidSkills(blnClassRow)
    If enValid = veSkip Or enValid = veEmpty Then Exit Sub
    lngHeroicLevels = HeroicLevels()
    strDashLine = "         " & String$(lngHeroicLevels * 3, 45)
    ' Title
    If enValid = veErrors Then strError = "(Errors)"
    OutputTitle "Skills", , strError
    enFont = sceCourierBegin
    ' Reddit
    If mblnReddit Then
        OutputSkillsReddit blnClassRow, lngHeroicLevels, (enValid = veComplete)
        Exit Sub
    End If
    ' Header
    If blnClassRow Then
        OutputText Space$(8), False, , enFont
        enFont = sceNone
        For lngCol = 1 To lngHeroicLevels
            With Skill.Col(lngCol)
                OutputText AlignText(.Initial, 3, vbRightJustify), (lngCol = lngHeroicLevels), .Color
            End With
        Next
    End If
    OutputText Space$(8), False, , enFont
    For lngCol = 1 To lngHeroicLevels
        With Skill.Col(lngCol)
            OutputText AlignText(CStr(lngCol), 3, vbRightJustify), (lngCol = lngHeroicLevels), .Color
        End With
    Next
    OutputText strDashLine
    ' Grid
    For lngRow = 0 To UBound(Skill.Out)
        lngSkill = Skill.Out(lngRow).Skill
        OutputText AlignText(GetSkillName(lngSkill, True), 9, vbLeftJustify), False
        For lngCol = 1 To lngHeroicLevels
            With Skill.grid(lngSkill, lngCol)
                OutputText FormatRanks(.Ranks), False, Skill.Col(lngCol).Color
            End With
        Next
        ' Skill total
        OutputText " " & FormatRanks(Skill.Row(lngSkill).Ranks)
    Next
    ' Footer
    OutputText strDashLine
    ' Spent
    OutputText Space$(8), False
    For lngCol = 1 To lngHeroicLevels - 1
        OutputText AlignText(CStr(Skill.Col(lngCol).Points), 3, vbRightJustify), False, Skill.Col(lngCol).Color
    Next
    If enValid = veComplete Then
        OutputText AlignText(CStr(Skill.Col(lngHeroicLevels).Points), 3, vbRightJustify), , Skill.Col(lngCol).Color, sceCourierEnd
    Else
        OutputText AlignText(CStr(Skill.Col(lngHeroicLevels).Points), 3, vbRightJustify), , Skill.Col(lngCol).Color
        ' Max
        OutputText AlignText("Max", 8, vbLeftJustify), False
        For lngCol = 1 To lngHeroicLevels - 1
            OutputText AlignText(CStr(Skill.Col(lngCol).MaxPoints), 3, vbRightJustify), False, Skill.Col(lngCol).Color
        Next
        OutputText AlignText(CStr(Skill.Col(lngHeroicLevels).MaxPoints), 3, vbRightJustify), , Skill.Col(lngCol).Color, sceCourierEnd
    End If
    BlankLine 2
End Sub

Private Sub OutputSkillsReddit(pblnClassRow As Boolean, plngLevels As Long, pblnComplete As Boolean)
    Dim strCell() As String
    Dim lngSkill As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    BlankLine
    ReDim strCell(plngLevels + 3)
    ' Header (level, not class)
    For i = 1 To plngLevels
        strCell(i + 1) = i
    Next
    OutputText Join(strCell, "|")
    ReDim strCell(plngLevels + 3)
    OutputText "|:--" & Join(strCell, "|--:")
    ' Class header
    If pblnClassRow Then
        ReDim strCell(plngLevels + 3)
        For i = 1 To plngLevels
            strCell(i + 1) = Skill.Col(i).Initial
        Next
        OutputText Join(strCell, "|")
    End If
    ' Grid
    For lngRow = 0 To UBound(Skill.Out)
        lngSkill = Skill.Out(lngRow).Skill
        strCell(1) = GetSkillName(lngSkill, True)
        For lngCol = 1 To plngLevels
            With Skill.grid(lngSkill, lngCol)
                strCell(lngCol + 1) = Trim$(FormatRanks(.Ranks))
            End With
        Next
        strCell(plngLevels + 2) = Trim$(FormatRanks(Skill.Row(lngSkill).Ranks))
        OutputText Join(strCell, "|")
    Next
    ' Spent
    ReDim strCell(plngLevels + 3)
    For i = 1 To plngLevels
        strCell(i + 1) = Skill.Col(i).Points
    Next
    OutputText Join(strCell, "|")
    ' Max
    If Not pblnComplete Then
        strCell(1) = "Max"
        For i = 1 To plngLevels
            strCell(i + 1) = Skill.Col(i).MaxPoints
        Next
        OutputText Join(strCell, "|")
    End If
    BlankLine
End Sub

Public Function ValidSkills(pblnClassRow As Boolean, Optional pblnImport As Boolean = False) As ValidEnum
    Dim typClassSplit() As ClassSplitType
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngSkill As Long
    Dim lngRanks As Long
    Dim enReturn As ValidEnum
    Dim i As Long
    
    ValidSkills = veSkip
    If mblnDisplay And Not pblnImport Then
        Select Case menOutput
            Case oeAll, oeOverview, oeSkills
            Case Else: Exit Function
        End Select
    End If
    ValidSkills = veEmpty
    ' Refresh grid stats
    InitBuildSkills
    ' Check for class row
    Select Case GetClassSplit(typClassSplit)
        Case 0: Exit Function
        Case 2, 3: pblnClassRow = True
    End Select
    ' Gather skills that have points assigned into an array
    ReDim Skill.Out(20)
    For lngRow = 1 To 21
        lngSkill = InitialSkillOrder(lngRow)
        With Skill.Row(lngSkill)
            If .Ranks > 0 Then
                Skill.Out(i).Skill = lngSkill
                Skill.Out(i).Ranks = .Ranks
                i = i + 1
            End If
        End With
    Next
    If i = 0 Then Exit Function
    enReturn = veComplete
    ' Discard unused skills
    ReDim Preserve Skill.Out(i - 1)
    ' Verify all points spent every level without going over
    For lngCol = 1 To HeroicLevels()
        If Skill.Col(lngCol).Points > Skill.Col(lngCol).MaxPoints Then
            enReturn = veErrors
            Exit For
        ElseIf Skill.Col(lngCol).Points < Skill.Col(lngCol).MaxPoints Then
            enReturn = veIncomplete
        End If
    Next
    ' Verify no skills exceeded their rolling caps (level+3)
    Do While enReturn <> veErrors
        For lngRow = 1 To 21
            lngRanks = 0
            For lngCol = 1 To HeroicLevels()
                lngRanks = lngRanks + Skill.grid(lngRow, lngCol).Ranks
                If lngRanks > Skill.grid(lngRow, lngCol).MaxRanks Then
                    enReturn = veErrors
                    Exit Do
                End If
            Next
        Next
        Exit Do
    Loop
    ' Sort by ranks spent
    If cfg.SkillOrderOutput = sooRanksSpent Then SortRows
    ' Finished
    ValidSkills = enReturn
End Function

Private Function InitialSkillOrder(plngRow As Long) As SkillsEnum
    If cfg.SkillOrderOutput = sooAlphabetical Then
        InitialSkillOrder = plngRow
    Else
        ' Custom skill order to be used for skills with the same number of ranks
        ' Concept here is to keep related skills together if they're maxed (caster skills, trapping skills, hide + move silently, etc...)
        Select Case plngRow
            Case 1: InitialSkillOrder = sePerform
            Case 2: InitialSkillOrder = seConcentration
            Case 3: InitialSkillOrder = seHeal
            Case 4: InitialSkillOrder = seRepair
            Case 5: InitialSkillOrder = seSpellcraft
            Case 6: InitialSkillOrder = seDisableDevice
            Case 7: InitialSkillOrder = seOpenLock
            Case 8: InitialSkillOrder = seSearch
            Case 9: InitialSkillOrder = seSpot
            Case 10: InitialSkillOrder = seListen
            Case 11: InitialSkillOrder = seBluff
            Case 12: InitialSkillOrder = seDiplomacy
            Case 13: InitialSkillOrder = seIntimidate
            Case 14: InitialSkillOrder = seBalance
            Case 15: InitialSkillOrder = seJump
            Case 16: InitialSkillOrder = seHide
            Case 17: InitialSkillOrder = seMoveSilently
            Case 18: InitialSkillOrder = seSwim
            Case 19: InitialSkillOrder = seTumble
            Case 20: InitialSkillOrder = seHaggle
            Case 21: InitialSkillOrder = seUMD
        End Select
    End If
End Function

' Insertion sort (stable algorithm, meaning skills retain order if have the same ranks)
Public Sub SortRows()
    Dim typHold As SkillOutputType
    Dim i As Long
    Dim j As Long

    For i = 1 To UBound(Skill.Out)
        typHold = Skill.Out(i)
        For j = i To 1 Step -1
            If typHold.Ranks > Skill.Out(j - 1).Ranks Then Skill.Out(j) = Skill.Out(j - 1) Else Exit For
        Next j
        Skill.Out(j) = typHold
    Next i
End Sub

Private Function FormatRanks(ByVal plngRanks As Long) As String
    Dim strTrail As String
    
    If plngRanks Mod 2 = 0 Then strTrail = " " Else strTrail = "½"
    Select Case plngRanks
        Case Is < 1: FormatRanks = "   "
        Case 1: FormatRanks = " ½ "
        Case 2 To 19: FormatRanks = " " & plngRanks \ 2 & strTrail
        Case Else: FormatRanks = plngRanks \ 2 & strTrail
    End Select
End Function


' ************* FEATS *************


Private Sub OutputFeats()
    Dim enValid As ValidEnum
    Dim enChannel As FeatChannelEnum
    Dim lngCount As Long
    Dim blnStarted As Boolean
    Dim strError As String
    Dim i As Long

    ' Check if we're displaying feats
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeFeats) Then Exit Sub
    'Get all valid feats
    enValid = ValidFeats()
    Select Case enValid
        Case veSkip: Exit Sub
        Case veEmpty: If menOutput <> oeFeats Then Exit Sub
    End Select
    ' Title
    If enValid = veErrors Then strError = "(Errors)"
    OutputTitle "Feats", , strError
    BlankLine
    If mblnReddit Then
        OutputText "|||"
        OutputText "|:--|:--|:--"
    End If
    ' Feat list mtypFeatOutput is loaded in ValidFeats
    With mtypFeatOutput
        If cfg.FeatOrderOutput = fooLevel Then
            For i = 1 To .Feats
                If mblnReddit Then
                    OutputFeatReddit i
                Else
                    OutputFeat i
                End If
            Next
            BlankLine
        Else
            'Show all feats if flag is set
            'This reads mtypFeatOutput, which is loaded
            Dim iEndChannel As Integer
            If cfg.FeatShowAll Then
                iEndChannel = fceGranted
            Else
                iEndChannel = fceDeity
            End If
            For enChannel = fceGeneral To iEndChannel
                lngCount = 0
                If enChannel <> fceGeneral Then
                    For i = 1 To .Feats
                        If .Feat(i).Channel = enChannel And Len(.Feat(i).FeatName) <> 0 Then Exit For
                    Next
                End If
                If i <= .Feats Then
                    If mblnReddit And blnStarted Then OutputText "|||"
                    blnStarted = True
                    For i = 1 To .Feats
                        If .Feat(i).Channel = enChannel Then
                            If mblnReddit Then OutputFeatReddit i Else OutputFeat i
                            lngCount = lngCount + 1
                        End If
                    Next
                    If lngCount <> 0 And mblnReddit = False Then BlankLine
                End If
            Next
        End If
    End With
    BlankLine
End Sub

Private Function OutputFeat(plngIndex As Long)
    Dim strLevel As String
    Dim strSource As String
    Dim blnExchange As Boolean
    Dim i As Long
    
    With mtypFeatOutput.Feat(plngIndex)
        blnExchange = (.Source = "Swap")
        strLevel = AlignText(CStr(.Level), 2, vbRightJustify)
        strSource = AlignText(.Source, 7, vbLeftJustify)
        If blnExchange Then
            OutputText strLevel & " ", False, , sceCourierBegin
            OutputText strSource, False, cfg.GetColor(cgeOutput, cveBlue)
            OutputText ": ", False, , sceCourierEnd
        Else
            OutputText strLevel & " " & strSource & ": ", False, , sceCourier
        End If
        If .Alternates = 0 Then
            If blnExchange Then
                OutputText .FeatName, False
                OutputText " replaces ", False, cfg.GetColor(cgeOutput, cveTextDim)
                OutputText .Replaces
            Else
                OutputText .FeatName
            End If
        Else
            OutputText .FeatName, False
            For i = 1 To .Alternates
                OutputText " OR ", False, cfg.GetColor(cgeOutput, cveTextDim)
                OutputText .Alternate(i), (i = .Alternates)
            Next
        End If
    End With
End Function

Private Function OutputFeatReddit(plngIndex As Long)
    Dim strSource As String
    Dim strFeat As String
    Dim blnExchange As Boolean
    Dim i As Long
    
    With mtypFeatOutput.Feat(plngIndex)
        blnExchange = (.Source = "Swap")
        strSource = .Level
        If Len(.Source) Then
            If blnExchange Then strSource = strSource & " **Swap**" Else strSource = strSource & " " & .Source
        End If
        strFeat = .FeatName
        If blnExchange Then
            strFeat = strFeat & " **replaces** " & .Replaces
        ElseIf .Alternates Then
            For i = 1 To .Alternates
                strFeat = strFeat & " **OR** " & .Alternate(i)
            Next
        End If
        OutputText "|" & strSource & "|" & strFeat & "|"
    End With
End Function

Public Function ValidFeats(Optional pblnImport As Boolean = False) As ValidEnum
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim blnEmpty As Boolean
    Dim blnErrors As Boolean
    Dim blnComplete As Boolean
    Dim i As Long
    
    If mblnDisplay And Not pblnImport Then
        Select Case menOutput
            Case oeAll, oeFeats
            Case Else: Exit Function
        End Select
    End If
    If build.Class(1) = ceAny Then Exit Function
    InitFeatList
    blnEmpty = True
    blnComplete = True
    mtypFeatOutput.Feats = 0
    'TODO MAX_FEATS
    ReDim mtypFeatOutput.Feat(MAX_FEATS)
    For i = 1 To Feat.Count
        ' Feat Output??
        Select Case AddFeatOutput(Feat.List(i))
            Case veSkip
            Case veEmpty
                blnComplete = False
            Case veErrors
                blnEmpty = False
                blnErrors = True
                blnComplete = False
            Case veComplete
                blnEmpty = False
        End Select
    Next
    If blnEmpty Then
        ValidFeats = veEmpty
    ElseIf blnErrors Then
        ValidFeats = veErrors
    ElseIf blnComplete Then
        ValidFeats = veComplete
    Else
        ValidFeats = veIncomplete
    End If
End Function

Private Function AddFeatOutput(ptypDetail As FeatDetailType) As ValidEnum
    Dim lngFeat As Long
    Dim strDisplay As String
    
    ' 2021.08.06 Show granted
    If Not cfg.FeatShowAll Then
        If ptypDetail.ActualType = bftGranted Then
            AddFeatOutput = veSkip
            Exit Function
        End If
    End If
    If ptypDetail.ErrorState Then
        AddFeatOutput = veErrors
    ElseIf ptypDetail.FeatID = 0 Then
        AddFeatOutput = veEmpty
    Else
        AddFeatOutput = veComplete
    End If
    With mtypFeatOutput
        If ptypDetail.ActualType = bftAlternate Then
            With .Feat(.Feats)
                .Alternates = .Alternates + 1
                ReDim Preserve .Alternate(1 To .Alternates)
                .Alternate(.Alternates) = ptypDetail.DisplayAlternate
            End With
        Else
            .Feats = .Feats + 1
            With .Feat(.Feats)
                .Level = ptypDetail.Level
                .FeatName = ptypDetail.Display
                .Source = ptypDetail.SourceOutput
                .Channel = ptypDetail.Channel
                If ptypDetail.ActualType = bftExchange And ptypDetail.ExchangeIndex <> 0 Then .Replaces = Feat.List(ptypDetail.ExchangeIndex).Display
            End With
        End If
    End With
End Function


' ************* SPELLS *************


Private Sub OutputSpells()
    Dim typClassSplit() As ClassSplitType
    Dim lngClasses As Long
    Dim lngCastingClasses As Long
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngLastLevel As Long
    Dim lngSlot As Long
    Dim strText As String
    Dim i As Long
    
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeSpells) Then Exit Sub
    Select Case ValidSpells(typClassSplit, lngClasses, lngCastingClasses)
        Case veSkip: Exit Sub
        Case veEmpty: If menOutput <> oeSpells Then Exit Sub
    End Select
    ' Title
    OutputTitle "Spells"
    BlankLine
    ' Reddit
    If mblnReddit Then
        OutputSpellsReddit typClassSplit, lngClasses
        Exit Sub
    ElseIf lngCastingClasses = 1 And mblnDisplay = False Then
        BlankLine
    End If
    ' Spell list
    For i = 0 To lngClasses - 1
        enClass = typClassSplit(i).ClassID
        lngLastLevel = build.Spell(enClass).MaxSpellLevel
        If lngLastLevel Then
            If lngCastingClasses > 1 Then OutputText db.Class(enClass).ClassName
            ListBegin True
            For lngLevel = 1 To lngLastLevel
                With build.Spell(enClass).Level(lngLevel)
                    ' Spell list for this class level
                    ListStartLine
                    strText = vbNullString
                    For lngSlot = 1 To .Slots
                        If Len(.Slot(lngSlot).Spell) = 0 And Not mblnDisplay Then
                            strText = strText & "<Any>"
                        Else
                            strText = strText & .Slot(lngSlot).Spell
                            If Len(.Slot(lngSlot).Spell) <> 0 Then strText = strText & " (" & .Slot(lngSlot).Level & ")"
                        End If
                        If lngSlot < .Slots Then strText = strText & ", "
                    Next
                    ' Output this line
                    OutputText strText, (lngLevel < lngLastLevel)
                End With
            Next
            ListEnd
            BlankLine
        End If
    Next
    If mblnDisplay Then BlankLine 2 Else BlankLine
End Sub

Private Sub OutputSpellsReddit(ptypClassSplit() As ClassSplitType, plngClasses As Long)
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngLastLevel As Long
    Dim lngSlot As Long
    Dim strText As String
    Dim blnBlankLine As Boolean
    Dim i As Long
    
    OutputText "|||"
    OutputText "|:--|:--|:--"
    For i = 0 To plngClasses - 1
        enClass = ptypClassSplit(i).ClassID
        lngLastLevel = build.Spell(enClass).MaxSpellLevel
        If lngLastLevel Then
            If blnBlankLine Then OutputText "|||" Else blnBlankLine = True
            For lngLevel = 1 To lngLastLevel
                With build.Spell(enClass).Level(lngLevel)
                    ' Spell list for this class level
                    strText = vbNullString
                    For lngSlot = 1 To .Slots
                        If Len(.Slot(lngSlot).Spell) = 0 And Not mblnDisplay Then
                            strText = strText & "<Any>"
                        Else
                            strText = strText & .Slot(lngSlot).Spell
                            If Len(.Slot(lngSlot).Spell) <> 0 Then strText = strText & " (" & .Slot(lngSlot).Level & ")"
                        End If
                        If lngSlot < .Slots Then strText = strText & ", "
                    Next
                    OutputText "|" & db.Class(enClass).ClassName & " " & lngLevel & "|" & strText & "|"
                End With
            Next
        End If
    Next
    BlankLine
End Sub

Private Function ValidSpells(ptypClassSplit() As ClassSplitType, plngClasses As Long, plngCastingClasses As Long) As ValidEnum
    Dim enClass As ClassEnum
    Dim lngLevel As Long
    Dim lngSlot As Long
    Dim blnEmpty As Boolean
    Dim blnComplete As Boolean
    Dim i As Long
    
    InitBuildSpells
    If build.CanCastSpell(1) = 0 Then Exit Function
    If mblnDisplay Then
        Select Case menOutput
            Case oeAll, oeSpells
            Case Else: Exit Function
        End Select
    End If
    blnComplete = True
    blnEmpty = True
    plngClasses = GetClassSplit(ptypClassSplit)
    For i = 0 To plngClasses - 1
        enClass = ptypClassSplit(i).ClassID
        For lngLevel = 1 To build.Spell(enClass).MaxSpellLevel
            For lngSlot = 1 To build.Spell(enClass).Level(lngLevel).Slots
                With build.Spell(enClass).Level(lngLevel).Slot(lngSlot)
                    If .SlotType = sseStandard Or .SlotType = sseFree Then
                        If Len(.Spell) = 0 Then blnComplete = False Else blnEmpty = False
                    End If
                End With
            Next
        Next
    Next
    plngCastingClasses = 0
    For i = 0 To plngClasses - 1
        With ptypClassSplit(i)
            If db.Class(.ClassID).CanCastSpell(1) <> 0 And .Levels >= db.Class(.ClassID).CanCastSpell(1) Then plngCastingClasses = plngCastingClasses + 1
        End With
    Next
    If blnEmpty Then
        ValidSpells = veEmpty
    ElseIf blnComplete Then
        ValidSpells = veComplete
    Else
        ValidSpells = veIncomplete
    End If
End Function

Private Function Mandatory(ByVal penClass As ClassEnum, pstrSpell As String) As Boolean
    Dim i As Long
    
    For i = 1 To db.Class(penClass).MandatorySpells
        If db.Class(penClass).MandatorySpell(i) = pstrSpell Then
            Mandatory = True
            Exit Function
        End If
    Next
End Function


' ************* ENHANCEMENTS *************
Private Sub OutputEnhancements()
    Dim enValid As ValidEnum
    Dim blnReverse As Boolean
    Dim frm As Form
    
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeEnhancements) Then
        Exit Sub
    End If
    enValid = ValidEnhancements()
    Select Case enValid
        Case veSkip: Exit Sub
        Case veEmpty: If build.Guide.Enhancements = 0 Then Exit Sub
    End Select
    If GetForm(frm, "frmEnhancements") Then
        blnReverse = (frm.CurrentTab = 2)
    End If
    ' Title
    If blnReverse Then
        If build.Guide.Enhancements > 0 Then OutputLevelingGuide True
        If enValid <> veEmpty Then OutputEnhancementTrees False, enValid
    Else
        If enValid <> veEmpty Then OutputEnhancementTrees True, enValid
        If build.Guide.Enhancements > 0 Then OutputLevelingGuide False
    End If
End Sub

Private Sub OutputEnhancementTrees(pblnFirst As Boolean, penValid As ValidEnum)
    Dim lngBuildTree As Long
    Dim lngTree As Long
    Dim lngTier As Long
    Dim strError As String
    Dim strDisplay As String
    Dim blnTiers As Boolean
    Dim typTreeOrder() As OutputTreeType
    Dim lngTreeOrder As Long
    Dim lngGuideTree As Long
    Dim lngColor As Long
    Dim i As Long
    
    ' Title
    strDisplay = GetSpentTreeText(penValid)
    If penValid = veErrors Then strError = "Errors"
    OutputTitle "Enhancements", strDisplay, strError
    ' Reddit
    If mblnReddit Then
        OutputEnhancementTreesReddit
        Exit Sub
    End If
    ' Build trees
    TreeOrder typTreeOrder
    For lngTreeOrder = 1 To build.Trees
        lngBuildTree = typTreeOrder(lngTreeOrder).ID
        If build.Tree(lngBuildTree).Abilities > 0 Then
            BlankLine
            'Guide
            lngGuideTree = FindGuideTree(lngBuildTree)
            If lngGuideTree > 0 Then
                If Guide.Tree(lngGuideTree).Spent > 0 Then
                    mblnColorOverride = True
                    ColorOpen Guide.Tree(lngGuideTree).Color
                End If
            End If
            OutputText build.Tree(lngBuildTree).TreeName & " (" & typTreeOrder(lngTreeOrder).AP & " AP)"
            lngTree = SeekTree(build.Tree(lngBuildTree).TreeName, peEnhancement)
            ' Cores
            strDisplay = vbNullString
            ListBegin False
            ListStartLine
            blnTiers = False
            For i = 1 To build.Tree(lngBuildTree).Abilities
                With build.Tree(lngBuildTree).Ability(i)
                    If .Tier = 0 Then
                        If Len(strDisplay) Then strDisplay = strDisplay & ", "
                        strDisplay = strDisplay & GetAbilityDisplay(db.Tree(lngTree), .Tier, .Ability, .Rank, .Selector)
                    Else
                        blnTiers = True
                    End If
                End With
            Next
            If Len(strDisplay) = 0 Then strDisplay = "(none)"
            OutputText strDisplay
            strDisplay = vbNullString
            If blnTiers Then
                ' Tiers
                ListBegin True
                lngTier = 1
                For i = 1 To build.Tree(lngBuildTree).Abilities
                    With build.Tree(lngBuildTree).Ability(i)
                        If .Tier > 0 Then
                            If .Tier > lngTier Then
                                If Len(strDisplay) Then strDisplay = Left$(strDisplay, Len(strDisplay) - 2) Else strDisplay = "(none)"
                                ListStartLine
                                OutputText strDisplay
                                lngTier = lngTier + 1
                                Do While .Tier > lngTier
                                    ListStartLine
                                    OutputText "(none)"
                                    lngTier = lngTier + 1
                                Loop
                                strDisplay = vbNullString
                            End If
                            strDisplay = strDisplay & GetAbilityDisplay(db.Tree(lngTree), .Tier, .Ability, .Rank, .Selector) & ", "
                        End If
                    End With
                Next
                If Len(strDisplay) Then strDisplay = Left$(strDisplay, Len(strDisplay) - 2) Else strDisplay = "(none)"
                ListStartLine
                OutputText strDisplay
                ListEnd
            End If
            ListEnd
            If lngGuideTree > 0 Then
                If Guide.Tree(lngGuideTree).Spent > 0 Then
                    ColorClose Guide.Tree(lngGuideTree).Color
                    mblnColorOverride = False
                End If
            End If
        End If
    Next
    If pblnFirst And build.Guide.Enhancements > 0 Then BlankLine 1 Else BlankLine 2
End Sub

Private Sub OutputEnhancementTreesReddit()
    Dim lngBuildTree As Long
    Dim lngTree As Long
    Dim lngTier As Long
    Dim strDisplay As String
    Dim blnTiers As Boolean
    Dim typTreeOrder() As OutputTreeType
    Dim lngTreeOrder As Long
    Dim lngGuideTree As Long
    Dim lngColor As Long
    Dim i As Long
    
    ' Build trees
    TreeOrder typTreeOrder
    For lngTreeOrder = 1 To build.Trees
        lngBuildTree = typTreeOrder(lngTreeOrder).ID
        If build.Tree(lngBuildTree).Abilities > 0 Then
            BlankLine
            lngGuideTree = FindGuideTree(lngBuildTree)
            OutputText build.Tree(lngBuildTree).TreeName & " (" & typTreeOrder(lngTreeOrder).AP & " AP)"
            BlankLine
            OutputText "|||"
            OutputText "|:--|:--|:--"
            lngTree = SeekTree(build.Tree(lngBuildTree).TreeName, peEnhancement)
            ' Cores
            strDisplay = vbNullString
            blnTiers = False
            For i = 1 To build.Tree(lngBuildTree).Abilities
                With build.Tree(lngBuildTree).Ability(i)
                    If .Tier = 0 Then
                        If Len(strDisplay) Then strDisplay = strDisplay & ", "
                        strDisplay = strDisplay & GetAbilityDisplay(db.Tree(lngTree), .Tier, .Ability, .Rank, .Selector)
                    Else
                        blnTiers = True
                    End If
                End With
            Next
            If Len(strDisplay) = 0 Then strDisplay = "(none)"
            OutputText "|Cores|" & strDisplay & "|"
            strDisplay = vbNullString
            If blnTiers Then
                ' Tiers
                lngTier = 1
                For i = 1 To build.Tree(lngBuildTree).Abilities
                    With build.Tree(lngBuildTree).Ability(i)
                        If .Tier > 0 Then
                            If .Tier > lngTier Then
                                If Len(strDisplay) Then strDisplay = Left$(strDisplay, Len(strDisplay) - 2) Else strDisplay = "(none)"
                                OutputText "|Tier " & lngTier & "|" & strDisplay & "|"
                                lngTier = lngTier + 1
                                Do While .Tier > lngTier
                                    OutputText "|Tier " & lngTier & "|(none)|"
                                    lngTier = lngTier + 1
                                Loop
                                strDisplay = vbNullString
                            End If
                            strDisplay = strDisplay & GetAbilityDisplay(db.Tree(lngTree), .Tier, .Ability, .Rank, .Selector) & ", "
                        End If
                    End With
                Next
                If Len(strDisplay) Then strDisplay = Left$(strDisplay, Len(strDisplay) - 2) Else strDisplay = "(none)"
                OutputText "|Tier " & lngTier & "|" & strDisplay & "|"
            End If
        End If
    Next
    BlankLine
End Sub

'This returns the text of Enh Spent and alters the penValid field if we're in error
' THis should be merged with frmEnhancements.GuideTotal and frmEnhancements.ShowSpentAll
Private Function GetSpentTreeText(penValid As ValidEnum) As String
    Dim lngSpentBase As Long
    Dim lngSpentRPLBonus As Long
    Dim lngSpentUniBonus As Long
    Dim lngMaxBase As Long
    Dim lngMaxRPLBonus As Long
    Dim lngMaxUniBonus As Long
    Dim strSpent As String
    Dim strMax As String
    Dim strDisplay As String
    
    'Display should be Spent/Max AP.  Long form is Spent+r+u/Max+r+u AP

    'retrieve each of the spent/maxes from the build tree object
    GetPointsSpentAndMax lngSpentBase, lngSpentRPLBonus, lngSpentUniBonus, lngMaxBase, lngMaxRPLBonus, lngMaxUniBonus
        
    'Display
    strDisplay = strDisplay & "Spent: " & lngSpentBase & " +" & lngSpentRPLBonus & "r +" & lngSpentUniBonus & "u / Max: "
    strDisplay = strDisplay & lngMaxBase & " +" & lngMaxRPLBonus & "r +" & lngMaxUniBonus & "u AP"
    
    'Figure out if we're in error
    If lngSpentBase > lngMaxBase Then
        penValid = veErrors
    End If
    
    'strSpent = lngSpentBase
    'If lngSpentBonus > 0 Then strSpent = strSpent & "+" & lngSpentBonus
    'strSpent = strSpent & " of "
    'strMax = lngMaxBase
    'If lngMaxBonus > 0 Then strMax = strMax & "+" & lngMaxBonus
    'If lngSpentBase = 80 And lngMaxBase = 80 And lngSpentBonus = lngMaxBonus Then strSpent = vbNullString
    
    'Return display
    GetSpentTreeText = "(" & strDisplay & ")"
End Function


'This returns the text of DSpent and alters the penValid field if we're in error
' THis should be merged with frmEnhancements.GuideTotal and frmEnhancements.ShowSpentAll
Private Function GetSpentDestinyText(penValid As ValidEnum) As String
    Dim lngSpentBase As Long
    Dim lngSpentPDPBonus As Long
    Dim lngMaxBase As Long
    Dim lngMaxPDPBonus As Long
    Dim strSpent As String
    Dim strMax As String
    Dim strDisplay As String
    
    'Display should be Spent/Max AP.  Long form is Spent+r+u/Max+r+u AP

    'retrieve each of the spent/maxes from the build tree object
    GetDestinyPointsSpentAndMax lngSpentBase, lngSpentPDPBonus, lngMaxBase, lngMaxPDPBonus
        
    'Display
    strDisplay = strDisplay & "Spent: " & (lngSpentBase + lngSpentPDPBonus) & " (" & lngSpentBase & " +" & lngSpentPDPBonus & "pdp) / Max: "
    strDisplay = strDisplay & (lngMaxBase + lngMaxPDPBonus) & " (" & lngMaxBase & " +" & lngMaxPDPBonus & "pdp) AP"
    
    'Figure out if we're in error
    If lngSpentBase > lngMaxBase Then
        penValid = veErrors
    End If
  
    
    'Return display
    GetSpentDestinyText = "(" & strDisplay & ")"
End Function


Private Function FindGuideTree(plngBuildTree As Long) As Long
    Dim i As Long
    
    If build.Guide.Enhancements > 0 Then
        For i = 1 To Guide.Trees
            If Guide.Tree(i).BuildTreeID = plngBuildTree Then
                FindGuideTree = i
                Exit For
            End If
        Next
    End If
End Function

Private Sub OutputLevelingGuide(pblnFirst As Boolean)
    Dim lngLevel As Long
    Dim blnGuideErrors As Boolean
    Dim lngTree As Long
    Dim lngTier As Long
    Dim blnIconic As Boolean
    Dim strError As String
    Dim i As Long
    
    For i = 1 To Guide.Enhancements
        If Guide.Enhancement(i).ErrorState Then
            blnGuideErrors = True
            Exit For
        End If
    Next
    If blnGuideErrors Then strError = "(Errors)"
    OutputTitle "Leveling Guide", , strError, False
    If mblnReddit Then
        OutputLevelingGuideReddit
        Exit Sub
    End If
    blnIconic = (db.Race(build.Race).Type = rteIconic)
    If Not blnIconic Then ListBegin True
    For lngLevel = FirstGuideLevel() To LastGuideLevel()
        If blnIconic Then
            If lngLevel > 14 Then OutputText lngLevel & ". ", False
        Else
            ListStartLine
        End If
        With mtypGuideOutput(lngLevel)
            If Len(.Reset) Then
                If blnIconic Then
                    If build.MaxLevels > 14 Then
                        If lngLevel <> 14 Then OutputText .Reset, mblnDisplay
                    Else
                        If lngLevel <> build.MaxLevels Then OutputText .Reset, mblnDisplay
                    End If
                Else
                    ' We need a new line only if displaying, or generating output that doesn't support list codes (plain text)
                    OutputText .Reset, mblnDisplay Or (mblnDisplay = False And (mblnCodes = False Or mblnLists = False))
                End If
                mblnColorOverride = True
                For lngTree = 1 To .BuildTrees
                    With .BuildTree(lngTree)
                        ColorOpen .Color
                        ListBegin False
                        ListStartLine
                        OutputText .TreeName & .Tier(0)
                        If .Tiers > 0 Then
                            ListBegin True
                            For lngTier = 1 To .Tiers
                                ListStartLine
                                OutputText .Tier(lngTier)
                            Next
                            ListEnd
                        End If
                        ListEnd
                        ColorClose .Color
                    End With
                Next
                mblnColorOverride = False
            ElseIf .Enhancements = 0 Then
                OutputText "(Bank AP)"
            Else
                For i = 1 To .Blocks
                    OutputText .Block(i).Display, (i = .Blocks), .Block(i).Color
                Next
            End If
        End With
    Next
    If Not blnIconic Then ListEnd
    If pblnFirst = False Or build.Trees = 0 Then BlankLine 2 Else BlankLine 1
End Sub

Private Sub OutputLevelingGuideReddit()
    Dim lngLevel As Long
    Dim lngTree As Long
    Dim lngTier As Long
    Dim blnIconic As Boolean
    Dim strDisplay As String
    Dim i As Long
    
    blnIconic = (db.Race(build.Race).Type = rteIconic)
    BlankLine
    OutputText "|||"
    OutputText "|--:|:--|:--"
    For lngLevel = FirstGuideLevel() To LastGuideLevel()
        With mtypGuideOutput(lngLevel)
            If Len(.Reset) Then
                If blnIconic Then
                    If build.MaxLevels > 14 Then
                        If lngLevel <> 14 Then OutputText "|" & lngLevel & "|" & .Reset & "|"
                    Else
                        If lngLevel <> build.MaxLevels Then OutputText "|" & lngLevel & "|" & .Reset & "|"
                    End If
                Else
                    OutputText "|" & lngLevel & "|" & .Reset & "|"
                End If
                For lngTree = 1 To .BuildTrees
                    With .BuildTree(lngTree)
                        OutputText "||**" & .TreeName & ":**" & Mid$(.Tier(0), 2) & "|"
                        For lngTier = 1 To .Tiers
                            OutputText "||_Tier " & lngTier & ":_ " & .Tier(lngTier) & "|"
                        Next
                    End With
                Next
            ElseIf .Enhancements = 0 Then
                OutputText "|" & lngLevel & "|**(Bank AP)**|"
            Else
                strDisplay = vbNullString
                For i = 1 To .Blocks
                    strDisplay = strDisplay & .Block(i).Display
                Next
                OutputText "|" & lngLevel & "|" & strDisplay & "|"
            End If
        End With
    Next
    BlankLine
End Sub

Private Function FirstGuideLevel() As Long
    If Guide.Enhancements = 0 Then Exit Function
    If db.Race(build.Race).Type = rteIconic Then
        If build.MaxLevels > 14 Then FirstGuideLevel = 14 Else FirstGuideLevel = build.MaxLevels
    Else
        FirstGuideLevel = 1
    End If
End Function

Private Function LastGuideLevel() As Long
    If Guide.Enhancements = 0 Then Exit Function
    LastGuideLevel = Guide.Enhancement(Guide.Enhancements).Level
End Function

Private Sub TreeOrder(ptypTreeOrder() As OutputTreeType)
    Dim typSwap As OutputTreeType
    Dim i As Long
    Dim j As Long
    
    ReDim ptypTreeOrder(1 To build.Trees)
    For i = 1 To build.Trees
        ptypTreeOrder(i).ID = i
        ptypTreeOrder(i).AP = QuickSpentInTree(i)
    Next
    For i = 2 To build.Trees
        typSwap = ptypTreeOrder(i)
        For j = i To 2 Step -1
            If typSwap.AP > ptypTreeOrder(j - 1).AP Then ptypTreeOrder(j) = ptypTreeOrder(j - 1) Else Exit For
        Next j
        ptypTreeOrder(j) = typSwap
    Next i
End Sub

Private Sub DestinyOrder(ptypTreeOrder() As OutputTreeType)
    Dim typSwap As OutputTreeType
    Dim i As Long
    Dim j As Long
    
    ReDim ptypTreeOrder(1 To build.Destinies)
    For i = 1 To build.Destinies
        ptypTreeOrder(i).ID = i
        ptypTreeOrder(i).AP = QuickSpentInDestiny(i)
    Next
    For i = 2 To build.Destinies
        typSwap = ptypTreeOrder(i)
        For j = i To 2 Step -1
            If typSwap.AP > ptypTreeOrder(j - 1).AP Then
                ptypTreeOrder(j) = ptypTreeOrder(j - 1)
            Else
                Exit For
            End If
        Next j
        ptypTreeOrder(j) = typSwap
    Next i
End Sub


Private Function ValidEnhancements() As ValidEnum
    Dim lngTree As Long
    Dim lngPoints As Long
    Dim blnEmpty As Boolean
    Dim blnError As Boolean
    Dim blnComplete As Boolean
    Dim i As Long
    
    If mblnDisplay And Not (menOutput = oeAll Or menOutput = oeEnhancements) Then
        ValidEnhancements = veSkip
        Exit Function
    End If
    blnEmpty = True
    For i = 1 To build.Trees
        lngTree = SeekTree(build.Tree(i).TreeName, peEnhancement)
        Select Case ValidTree(db.Tree(lngTree), build.Tree(i), lngPoints)
            Case veErrors
                blnError = True
                blnEmpty = False
            Case veComplete
                blnEmpty = False
        End Select
    Next
    'TODO do we need veIncomplete ???
    If blnEmpty Then
        ValidEnhancements = veEmpty
    ElseIf blnError Then
        ValidEnhancements = veErrors
    Else
        ValidEnhancements = veComplete
    End If
    InitGuideOutput
End Function

Private Function ValidTree(ptypTree As TreeType, ptypBuildTree As BuildTreeType, plngPoints As Long) As ValidEnum
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSpent() As Long
    Dim lngTotal As Long
    Dim blnError As Boolean
    Dim i As Long
    
    plngPoints = 0
    
    GetSpentInTree ptypTree, ptypBuildTree, lngSpent, lngTotal
    For i = 1 To ptypBuildTree.Abilities
        With ptypBuildTree.Ability(i)
            lngTier = .Tier
            lngAbility = .Ability
        End With
        If lngAbility = 0 Then
            blnError = True
        ElseIf CheckLevels(ptypBuildTree, lngTier, lngAbility) Then
            blnError = True
        Else
            If CheckAbilityErrors(ptypTree, ptypBuildTree, ptypBuildTree.Ability(i), lngSpent) Then blnError = True
        End If
        If blnError Then Exit For
    Next
    ' Points in tree
    If ptypBuildTree.Abilities = 0 Then
        ValidTree = veEmpty
    Else
        ValidTree = CheckSpentInTree(ptypTree, ptypBuildTree, plngPoints)
    End If
    If blnError Then ValidTree = veErrors
End Function



' Returns total points spent in tree, or 0 if "spent in tree" prereqs are violated
Private Function CheckSpentInTree(ptypTree As TreeType, ptypBuildTree As BuildTreeType, plngPoints As Long) As ValidEnum
    Dim lngSpent() As Long
    Dim lngMaxTier As Long
    Dim blnError As Boolean
    Dim i As Long
    
    If Not GetSpentInTree(ptypTree, ptypBuildTree, lngSpent, plngPoints) Then
        blnError = True
    Else
        With ptypBuildTree
            If .Abilities > 0 Then lngMaxTier = .Ability(.Abilities).Tier
        End With
        For i = lngMaxTier To 2 Step -1
            If lngSpent(i) > 0 Then
                If lngSpent(i - 1) < GetSpentReq(ptypTree.TreeType, i, 1) Then
                    blnError = True
                    Exit For
                End If
            End If
        Next
    End If
    If blnError Then
        CheckSpentInTree = veErrors
    ElseIf plngPoints = 0 Then
        CheckSpentInTree = veEmpty
    Else
        CheckSpentInTree = veComplete
    End If
End Function

Private Function GetAbilityDisplay(ptypTree As TreeType, ByVal plngTier As Long, ByVal plngAbility As Long, ByVal plngRanks As Long, ByVal plngSelector As Long) As String
    Dim strReturn As String
    
    If plngAbility = 0 Then
        GetAbilityDisplay = "Error"
        Exit Function
    End If
    With ptypTree.Tier(plngTier).Ability(plngAbility)
        If plngSelector Then
            If .SelectorOnly Then
                strReturn = .Selector(plngSelector).SelectorName
            Else
                strReturn = .Abbreviation & ": " & .Selector(plngSelector).SelectorName
            End If
            'TODO Fix display of Ability/Selector
            'strReturn = .Abbreviation & ": " & .Selector(plngSelector).SelectorName
        Else
            strReturn = .Abbreviation
        End If
        If .Ranks > 1 And plngRanks <> 0 Then strReturn = strReturn & " " & Left$("III", plngRanks)
    End With
    GetAbilityDisplay = strReturn
End Function

Private Sub InitGuideOutput()
    Dim lngLastLevel As Long
    Dim lngLevel As Long
    Dim lngColor As Long
    Dim blnNew As Boolean
    Dim lngResetLevel As Long
    Dim i As Long
    
    If build.Guide.Enhancements = 0 Then Exit Sub
    InitGuideOutputTrees
    lngLastLevel = Guide.Enhancement(Guide.Enhancements).Level
    ' Gather data
    ReDim mtypGuideOutput(1 To lngLastLevel)
    If db.Race(build.Race).Type = rteIconic Then
        If build.MaxLevels > 14 Then
            lngResetLevel = 14
            mtypGuideOutput(14).Reset = "Iconic spends 56 AP to start"
        Else
            lngResetLevel = build.MaxLevels
            mtypGuideOutput(build.MaxLevels).Reset = "Iconic spends all AP to start"
        End If
    End If
    For i = 1 To Guide.Enhancements
        Select Case Guide.Enhancement(i).Style
            Case geEnhancement
                If Guide.Enhancement(i).Level = lngResetLevel Then
                    ResetAbility i
                Else
                    With mtypGuideOutput(Guide.Enhancement(i).Level)
                        blnNew = True
                        If .Enhancements > 0 Then
                            If .Enhancement(.Enhancements).GuideTreeID = Guide.Enhancement(i).GuideTreeID Then
                                If .Enhancement(.Enhancements).Tier = Guide.Enhancement(i).Tier Then
                                    If .Enhancement(.Enhancements).Enhancement = Guide.Enhancement(i).Display Then
                                        blnNew = False
                                    End If
                                End If
                            End If
                        End If
                        If blnNew Then
                            .Enhancements = .Enhancements + 1
                            ReDim Preserve .Enhancement(1 To .Enhancements)
                            With .Enhancement(.Enhancements)
                                .GuideTreeID = Guide.Enhancement(i).GuideTreeID
                                .Tree = Guide.Tree(.GuideTreeID).Initial
                                .Tier = Guide.Enhancement(i).Tier
                                .Enhancement = Guide.Enhancement(i).Display
                                .Display = .Tree & .Tier & " " & Guide.Enhancement(i).Display & Guide.Enhancement(i).RankText
                                .Color = Guide.Tree(.GuideTreeID).Color
                            End With
                        Else
                            With .Enhancement(.Enhancements)
                                .Display = .Display & "," & Guide.Enhancement(i).RankText
                            End With
                        End If
                    End With
                End If
            Case geBankAP
                With mtypGuideOutput(Guide.Enhancement(i).Level)
                    .Enhancements = .Enhancements + 1
                    ReDim Preserve .Enhancement(1 To .Enhancements)
                    With .Enhancement(.Enhancements)
                        .Display = "(Bank " & Guide.Enhancement(i).Bank & " AP)"
                        .Color = -1
                    End With
                End With
            Case geResetTree, geResetAllTrees
                With Guide.Enhancement(i)
                    lngResetLevel = .Level
                    mtypGuideOutput(.Level).Reset = .Display
                End With
        End Select
    Next
    ' Process data
    For lngLevel = 1 To lngLastLevel
        With mtypGuideOutput(lngLevel)
            If Len(.Reset) = 0 Then
                lngColor = -1
                For i = 1 To .Enhancements
                    If i = 1 Or .Enhancement(i).Color <> lngColor Then
                        .Blocks = .Blocks + 1
                        ReDim Preserve .Block(1 To .Blocks)
                        lngColor = .Enhancement(i).Color
                        .Block(.Blocks).Color = lngColor
                    End If
                    .Block(.Blocks).Display = .Block(.Blocks).Display & .Enhancement(i).Display & "; "
                Next
                ' Trim out final semicolon
                If .Blocks > 0 Then .Block(.Blocks).Display = Left$(.Block(.Blocks).Display, Len(.Block(.Blocks).Display) - 2)
            Else
                For i = 1 To .BuildTrees
                    ResetBuildTree .BuildTree(i)
                Next
            End If
        End With
    Next
End Sub

' Add ability to guide level's buildtree
Private Sub ResetAbility(plngIndex As Long)
    Dim typNew As BuildAbilityType
    Dim lngGuideTree As Long
    Dim lngBuildTree As Long
    Dim lngLevel As Long
    Dim blnFound As Boolean
    Dim lngInsert As Long
    Dim strTreeName As String
    Dim lngColor As Long
    Dim i As Long
    
    ' Prep new ability
    With Guide.Enhancement(plngIndex)
        typNew.Tier = .Tier
        typNew.Ability = .Ability
        typNew.Selector = .Selector
        typNew.Rank = .Rank
        lngLevel = .Level
        lngGuideTree = .GuideTreeID
    End With
    With Guide.Tree(lngGuideTree)
        strTreeName = .TreeName
        lngColor = .Color
    End With
    With mtypGuideOutput(lngLevel)
        ' Find build tree
        For lngBuildTree = 1 To .BuildTrees
            If .BuildTree(lngBuildTree).GuideTreeID = lngGuideTree Then Exit For
        Next
        If lngBuildTree > .BuildTrees Then
            .BuildTrees = .BuildTrees + 1
            ReDim Preserve .BuildTree(1 To .BuildTrees)
            With .BuildTree(.BuildTrees)
                .GuideTreeID = lngGuideTree
                .TreeName = strTreeName
                .Color = lngColor
            End With
        End If
        With .BuildTree(lngBuildTree)
            ' Find ability if it already exists, identify insertion point if it doesn't
            For lngInsert = 1 To .Abilities
                If .Ability(lngInsert).Tier = typNew.Tier Then
                    If .Ability(lngInsert).Ability = typNew.Ability Then
                        blnFound = True
                        Exit For
                    ElseIf .Ability(lngInsert).Ability > typNew.Ability Then
                        Exit For
                    End If
                ElseIf .Ability(lngInsert).Tier > typNew.Tier Then
                    Exit For
                End If
            Next
            If blnFound Then
                If .Ability(lngInsert).Rank < typNew.Rank Then .Ability(lngInsert).Rank = typNew.Rank
            Else
                .Abilities = .Abilities + 1
                ReDim Preserve .Ability(1 To .Abilities)
                For i = .Abilities To lngInsert + 1 Step -1
                    .Ability(i) = .Ability(i - 1)
                Next
                .Ability(lngInsert) = typNew
            End If
        End With
    End With
End Sub

Private Sub ResetBuildTree(ptypTree As GuideOutputBuildTreeType)
    Dim lngTree As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRank As Long
    Dim strComma As String
    Dim i As Long
    
    If ptypTree.Abilities = 0 Or ptypTree.GuideTreeID = 0 Then Exit Sub
    lngTree = Guide.Tree(ptypTree.GuideTreeID).TreeID
    With ptypTree
        .Tiers = .Ability(.Abilities).Tier
        ReDim .Tier(.Tiers)
        For i = 1 To .Abilities
            With .Ability(i)
                lngTier = .Tier
                lngAbility = .Ability
                lngSelector = .Selector
                lngRank = .Rank
            End With
            If Len(.Tier(lngTier)) Then strComma = ", " Else strComma = vbNullString
            .Tier(lngTier) = .Tier(lngTier) & strComma & GetAbilityDisplay(db.Tree(lngTree), lngTier, lngAbility, lngRank, lngSelector)
        Next
        If Len(.Tier(0)) Then .Tier(0) = ": " & .Tier(0)
        For i = 1 To .Tiers
            If Len(.Tier(i)) = 0 Then .Tier(i) = "(none)"
        Next
    End With
End Sub

Private Sub InitGuideOutputTrees()
    Dim typTree() As OutputTreeType
    Dim typSwap As OutputTreeType
    Dim blnColor() As Boolean
    Dim i As Long
    Dim j As Long
    
    ' Get colors and initialize temp tree array
    If Guide.Trees > 0 Then ReDim typTree(1 To Guide.Trees)
    For i = 1 To Guide.Trees
        typTree(i).ID = i
        Guide.Tree(i).Color = -1
        Guide.Tree(i).Spent = 0
        With Guide.Tree(i)
            .Initial = db.Tree(.TreeID).Initial(0)
            If .TreeStyle = tseRace Then
                .Color = -1
            Else
                .Color = db.Tree(.TreeID).Color
            End If
        End With
    Next
    ' Count how many AP spent in each tree
    For i = 1 To Guide.Enhancements
        With Guide.Enhancement(i)
            If .GuideTreeID <> 0 And .Style = geEnhancement Then
                typTree(.GuideTreeID).AP = typTree(.GuideTreeID).AP + .Cost
                Guide.Tree(.GuideTreeID).Spent = Guide.Tree(.GuideTreeID).Spent + .Cost
            End If
        End With
    Next
    ' Sort trees by most AP spent first
    For i = 2 To Guide.Trees
        typSwap = typTree(i)
        For j = i To 2 Step -1
            If typSwap.AP > typTree(j - 1).AP Then typTree(j) = typTree(j - 1) Else Exit For
        Next j
        typTree(j) = typSwap
    Next i
    ' Commit colors
    ReDim blnColor(cveColorValues)
    For i = 1 To Guide.Trees
        If typTree(i).AP = 0 Then Exit For
        With Guide.Tree(typTree(i).ID)
            If .TreeStyle = tseRace Then .Color = cveLightGray Else .Color = FindColor(.Color, cveYellow, cveGreen, cveBlue, cveRed, cvePurple, cveOrange, blnColor)
            If .Color <> -1 Then .Color = cfg.GetColor(cgeOutput, .Color)
        End With
    Next
    ' Make abbreviations unique
    For i = 1 To Guide.Trees
        For j = i + 1 To Guide.Trees
            If Guide.Tree(typTree(i).ID).Initial = Guide.Tree(typTree(j).ID).Initial Then
                Guide.Tree(typTree(i).ID).Initial = db.Tree(Guide.Tree(typTree(i).ID).TreeID).Initial(1)
                Guide.Tree(typTree(j).ID).Initial = db.Tree(Guide.Tree(typTree(j).ID).TreeID).Initial(1)
                Exit For
            End If
        Next
    Next
End Sub

Private Function FindColor(penValue As ColorValueEnum, pen1 As ColorValueEnum, pen2 As ColorValueEnum, pen3 As ColorValueEnum, pen4 As ColorValueEnum, pen5 As ColorValueEnum, pen6 As ColorValueEnum, pblnTaken() As Boolean) As ColorValueEnum
    Dim enColor(6) As ColorValueEnum
    Dim i As Long
    
    FindColor = -1
    If penValue <> -1 Then
        enColor(0) = penValue
        enColor(1) = pen1
        enColor(2) = pen2
        enColor(3) = pen3
        enColor(4) = pen4
        enColor(5) = pen5
        enColor(6) = pen6
        For i = 0 To 6
            If enColor(i) <> -1 Then
                If Not pblnTaken(enColor(i)) Then
                    FindColor = enColor(i)
                    pblnTaken(enColor(i)) = True
                    Exit Function
                End If
            End If
        Next
    End If
End Function


' ************* DESTINY *************
Private Sub OutputDestinies()
    Dim enValid As ValidEnum
    Dim blnReverse As Boolean
    Dim frm As Form
    
    If Not mblnDisplay And Not (cfg.OutputSection = oeAll Or cfg.OutputSection = oeDestiny) Then
        Exit Sub
    End If
    enValid = ValidDestinies()
    Select Case enValid
        Case veSkip: Exit Sub
        Case veEmpty: Exit Sub
    End Select
    If GetForm(frm, "frmDestiny2") Then
        blnReverse = (frm.CurrentTab = 2)
    End If
    ' Title
    If blnReverse Then
        If enValid <> veEmpty Then OutputDestinyTrees False, enValid
    Else
        If enValid <> veEmpty Then OutputDestinyTrees True, enValid
    End If
End Sub

Private Sub OutputDestinyTrees(pblnFirst As Boolean, penValid As ValidEnum)
    Dim lngBuildDestiny As Long
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim strError As String
    Dim strDisplay As String
    Dim blnTiers As Boolean
    Dim typDestinyOrder() As OutputTreeType
    Dim lngDestinyOrder As Long
    Dim i As Long
   
    ' Title
    strDisplay = GetSpentDestinyText(penValid)
    If penValid = veErrors Then strError = "Errors"
    OutputTitle "Destinies", strDisplay, strError
    ' Reddit
    If mblnReddit Then
        OutputDestinyTreesReddit
        Exit Sub
    End If

    ' Build trees
    DestinyOrder typDestinyOrder
    For lngDestinyOrder = 1 To build.Destinies
        lngBuildDestiny = typDestinyOrder(lngDestinyOrder).ID
        If build.Destiny(lngBuildDestiny).Abilities > 0 Then
            BlankLine
            'Guide - TODO
            OutputText build.Destiny(lngBuildDestiny).TreeName & " (" & typDestinyOrder(lngDestinyOrder).AP & " AP)"
            lngDestiny = SeekTree(build.Destiny(lngBuildDestiny).TreeName, peDestiny)
            ' Cores
            strDisplay = vbNullString
            ListBegin False
            ListStartLine
            blnTiers = False
            'Output tier 0
            For i = 1 To build.Destiny(lngBuildDestiny).Abilities
                With build.Destiny(lngBuildDestiny).Ability(i)
                    If .Tier = 0 Then
                        If Len(strDisplay) Then strDisplay = strDisplay & ", "
                        strDisplay = strDisplay & GetAbilityDisplay(db.Destiny(lngDestiny), .Tier, .Ability, .Rank, .Selector)
                    Else
                        blnTiers = True
                    End If
                End With
            Next
            If Len(strDisplay) = 0 Then strDisplay = "(none)"
            OutputText strDisplay
            strDisplay = vbNullString
            'Output tiers 1-5
            If blnTiers Then
                ' Tiers
                ListBegin True
                lngTier = 1
                For i = 1 To build.Destiny(lngBuildDestiny).Abilities
                    With build.Destiny(lngBuildDestiny).Ability(i)
                        If .Tier > 0 Then
                            If .Tier > lngTier Then
                                If Len(strDisplay) Then
                                    strDisplay = Left$(strDisplay, Len(strDisplay) - 2)
                                Else
                                    strDisplay = "(none)"
                                End If
                                ListStartLine
                                OutputText strDisplay
                                lngTier = lngTier + 1
                                Do While .Tier > lngTier
                                    ListStartLine
                                    OutputText "(none)"
                                    lngTier = lngTier + 1
                                Loop
                                strDisplay = vbNullString
                            End If
                            strDisplay = strDisplay & GetAbilityDisplay(db.Destiny(lngDestiny), .Tier, .Ability, .Rank, .Selector) & ", "
                        End If
                    End With
                Next
                If Len(strDisplay) Then
                    strDisplay = Left$(strDisplay, Len(strDisplay) - 2)
                Else
                    strDisplay = "(none)"
                End If
                ListStartLine
                OutputText strDisplay
                ListEnd
            End If
            ListEnd
        End If
    Next
    If pblnFirst Then BlankLine 1 Else BlankLine 2
End Sub

Private Sub OutputDestinyTreesReddit()
    Dim lngBuildDestiny As Long
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim strDisplay As String
    Dim blnTiers As Boolean
    Dim typDestinyOrder() As OutputTreeType
    Dim lngDestinyOrder As Long
    Dim lngColor As Long
    Dim i As Long
    
    ' Build Destinys
    DestinyOrder typDestinyOrder
    For lngDestinyOrder = 1 To build.Destinies
        lngBuildDestiny = typDestinyOrder(lngDestinyOrder).ID
        If build.Destiny(lngBuildDestiny).Abilities > 0 Then
            BlankLine
            OutputText build.Destiny(lngBuildDestiny).TreeName & " (" & typDestinyOrder(lngDestinyOrder).AP & " AP)"
            BlankLine
            OutputText "|||"
            OutputText "|:--|:--|:--"
            lngDestiny = SeekTree(build.Destiny(lngBuildDestiny).TreeName, peDestiny)
            ' Cores
            strDisplay = vbNullString
            blnTiers = False
            For i = 1 To build.Destiny(lngBuildDestiny).Abilities
                With build.Destiny(lngBuildDestiny).Ability(i)
                    If .Tier = 0 Then
                        If Len(strDisplay) Then strDisplay = strDisplay & ", "
                        strDisplay = strDisplay & GetAbilityDisplay(db.Destiny(lngDestiny), .Tier, .Ability, .Rank, .Selector)
                    Else
                        blnTiers = True
                    End If
                End With
            Next
            If Len(strDisplay) = 0 Then strDisplay = "(none)"
            OutputText "|Cores|" & strDisplay & "|"
            strDisplay = vbNullString
            If blnTiers Then
                ' Tiers
                lngTier = 1
                For i = 1 To build.Destiny(lngBuildDestiny).Abilities
                    With build.Destiny(lngBuildDestiny).Ability(i)
                        If .Tier > 0 Then
                            If .Tier > lngTier Then
                                If Len(strDisplay) Then strDisplay = Left$(strDisplay, Len(strDisplay) - 2) Else strDisplay = "(none)"
                                OutputText "|Tier " & lngTier & "|" & strDisplay & "|"
                                lngTier = lngTier + 1
                                Do While .Tier > lngTier
                                    OutputText "|Tier " & lngTier & "|(none)|"
                                    lngTier = lngTier + 1
                                Loop
                                strDisplay = vbNullString
                            End If
                            strDisplay = strDisplay & GetAbilityDisplay(db.Destiny(lngDestiny), .Tier, .Ability, .Rank, .Selector) & ", "
                        End If
                    End With
                Next
                If Len(strDisplay) Then strDisplay = Left$(strDisplay, Len(strDisplay) - 2) Else strDisplay = "(none)"
                OutputText "|Tier " & lngTier & "|" & strDisplay & "|"
            End If
        End If
    Next
    BlankLine
End Sub


Private Function ValidDestinies() As ValidEnum
    Dim lngDestiny As Long
    Dim lngPoints As Long
    Dim blnEmpty As Boolean
    Dim blnError As Boolean
    Dim blnComplete As Boolean
    Dim i As Long
    
    If mblnDisplay And Not (menOutput = oeAll Or menOutput = oeDestiny) Then
        ValidDestinies = veSkip
        Exit Function
    End If
    blnEmpty = True
    For i = 1 To build.Destinies
        lngDestiny = SeekTree(build.Destiny(i).TreeName, peDestiny)
        Select Case ValidTree(db.Destiny(lngDestiny), build.Destiny(i), lngPoints)
            Case veErrors
                blnError = True
                blnEmpty = False
            Case veComplete
                blnEmpty = False
        End Select
    Next
    'TODO do we need veIncomplete ???
    If blnEmpty Then
        ValidDestinies = veEmpty
    ElseIf blnError Then
        ValidDestinies = veErrors
    Else
        ValidDestinies = veComplete
    End If
    Exit Function
        'Select Case plngPoints
        '    Case 0: enValid = veEmpty
        '    Case 1 To (MaxDestinyAP + build.PermDestinyPoints - 1): enValid = veIncomplete
        '    Case (MaxDestinyAP + build.PermDestinyPoints): enValid = veComplete
        '    Case Else: enValid = veErrors
        'End Select
End Function



' ************* OUTPUT *************


Private Sub ErrorFlag(pblnErrors As Boolean)
    If pblnErrors Then
        OutputText " (Errors)", True, cfg.GetColor(cgeOutput, cveTextError)
    Else
        BlankLine
    End If
End Sub

Private Sub OutputText(ByVal pstrText As String, Optional pblnNewLine As Boolean = True, Optional ByVal plngColor As Long = -1, Optional penSpecial As SpecialCodeEnum = sceNone, Optional pblnBold As Boolean, Optional pblnUnderline As Boolean)
    If mblnDisplay Then
        With frmMain.picBuild
            SetStyle pblnBold, pblnUnderline, plngColor, penSpecial
            GrowClient .TextWidth(pstrText), .TextHeight(pstrText)
            If pblnNewLine Then
                frmMain.picBuild.Print pstrText
            Else
                frmMain.picBuild.Print pstrText;
            End If
            ResetStyle pblnBold, pblnUnderline, plngColor, penSpecial
        End With
    Else
        If mblnCodes Then pstrText = InsertCodes(pstrText, pblnBold, pblnUnderline, plngColor, penSpecial)
        mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & pstrText
        If pblnNewLine Then
            mstrJoin(mlngCurrent) = RTrim$(mstrJoin(mlngCurrent))
            mlngCurrent = mlngCurrent + 1
        End If
    End If
End Sub

Private Sub OutputTitle(ByVal pstrText As String, Optional ByVal pstrExtra As String = vbNullString, Optional pstrError As String = vbNullString, Optional pblnUnderline As Boolean = True)
    Dim strText As String
    
    If mblnReddit Then
        strText = "**" & pstrText & "**"
        If Len(pstrExtra) Then strText = strText & " " & pstrExtra
        If Len(pstrError) Then strText = strText & " _" & pstrError & "_"
        OutputText strText
    Else
        OutputText pstrText, False, , , , pblnUnderline
        If Len(pstrExtra) Then OutputText " " & pstrExtra, False
        If Len(pstrError) Then OutputText " " & pstrError, False, cfg.GetColor(cgeOutput, cveTextError)
        BlankLine
    End If
End Sub

Private Sub ColorOpen(plngColor As Long)
    Dim lngColor As Long
    
    If Not mblnTextColors Then Exit Sub
    If mblnDisplay Then
        If plngColor = -1 Then lngColor = cfg.GetColor(cgeOutput, cveText) Else lngColor = plngColor
        frmMain.picBuild.ForeColor = lngColor
    Else
        If plngColor <> -1 And mblnColor Then mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & Replace(mstrColorOpen, "$", xp.ColorToHex(plngColor))
    End If
End Sub

Private Sub ColorClose(plngColor As Long)
    If Not mblnTextColors Then Exit Sub
    If mblnDisplay Then
        frmMain.picBuild.ForeColor = cfg.GetColor(cgeOutput, cveText)
    Else
        If plngColor <> -1 And mblnColor Then mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & mstrColorClose
    End If
End Sub

Private Sub GrowClient(plngWidth As Long, plngHeight As Long)
    With frmMain.picBuild
        ' Width
        If .CurrentX + plngWidth > mlngMaxWidth Then mlngMaxWidth = .CurrentX + plngWidth
        If mlngMaxWidth > .Width Then .Width = mlngMaxWidth
        ' Height
        If .CurrentY + plngHeight > mlngMaxHeight Then mlngMaxHeight = .CurrentY + plngHeight
        If mlngMaxHeight > .Height Then .Height = mlngMaxHeight
    End With
End Sub

Private Sub SetStyle(pblnBold As Boolean, pblnUnderline As Boolean, plngColor As Long, penSpecial As SpecialCodeEnum)
    With frmMain.picBuild
        If plngColor <> -1 Then
            If mblnTextColors Then .ForeColor = plngColor
        End If
        If penSpecial = sceCourier Or penSpecial = sceCourierBegin Then
            .FontName = mstrCourierFont
            .FontSize = msngCourierSize
        End If
        If pblnBold Then .FontBold = True
        If pblnUnderline Then .FontUnderline = True
    End With
End Sub

Private Sub ResetStyle(pblnBold As Boolean, pblnUnderline As Boolean, plngColor As Long, penSpecial As SpecialCodeEnum)
    With frmMain.picBuild
        If plngColor <> -1 Then
            If mblnTextColors Then .ForeColor = cfg.GetColor(cgeOutput, cveText)
        End If
        If penSpecial = sceCourier Or penSpecial = sceCourierEnd Then
            .FontName = mstrDefaultFont
            .FontSize = msngDefaultSize
        End If
        If pblnBold Then .FontBold = False
        If pblnUnderline Then .FontUnderline = False
    End With
End Sub

' In order for the AddDots() function to have the desired effect on the output, code tags have to press directly up against the text.
' This function strips out leading and trailing spaces, adds the code tags, then puts the leading and trailing spaces back in.
' "   Bold   " => "   [B]Bold[/B]   "
Private Function InsertCodes(pstrText As String, pblnBold As Boolean, pblnUnderline As Boolean, plngColor As Long, penSpecial As SpecialCodeEnum) As String
    Dim strPrefix As String
    Dim strSuffix As String
    Dim lngLeading As Long
    Dim lngTrailing As Long
    Dim strReturn As String
    
    If Not mblnCodes Then
        InsertCodes = pstrText
        Exit Function
    End If
    ' Strip out leading spaces
    strReturn = LTrim$(pstrText)
    lngLeading = Len(pstrText) - Len(strReturn)
    ' Strip out trailing spaces
    lngTrailing = Len(strReturn)
    strReturn = RTrim$(strReturn)
    lngTrailing = lngTrailing - Len(strReturn)
    ' Most codes can be skipped if we're just sending spaces
    If Len(strReturn) Then
        ' Prefix
        If plngColor <> -1 Then
            If mblnTextColors And mblnColor Then strPrefix = strPrefix & Replace(mstrColorOpen, "$", xp.ColorToHex(plngColor))
        End If
        If pblnBold And mblnBold Then strPrefix = strPrefix & mstrBoldOpen
        If pblnUnderline And mblnUnderline Then strPrefix = strPrefix & mstrUnderlineOpen
        ' Suffix
        If pblnUnderline And mblnUnderline Then strSuffix = strSuffix & mstrUnderlineClose
        If pblnBold And mblnBold Then strSuffix = strSuffix & mstrBoldClose
        If plngColor <> -1 Then
            If mblnTextColors And mblnColor Then strSuffix = strSuffix & mstrColorClose
        End If
    End If
    ' Add leading and trailing spaces back in
    strPrefix = Space$(lngLeading) & strPrefix
    strSuffix = strSuffix & Space$(lngTrailing)
    ' Courier tags must stay outside of the leading/trailing spaces to preserve proper formatting
    If mblnFixed And (penSpecial = sceCourier Or penSpecial = sceCourierBegin) Then strPrefix = mstrFixedOpen & strPrefix
    If mblnFixed And (penSpecial = sceCourier Or penSpecial = sceCourierEnd) Then strSuffix = strSuffix & mstrFixedClose
    ' Add code tags to our string
    InsertCodes = strPrefix & strReturn & strSuffix
End Function

Public Function AlignText(pstrText As String, plngWidth As Long, penAlign As AlignmentConstants) As String
    Dim strReturn As String
    Dim lngPadding As Long
    
    lngPadding = plngWidth - Len(pstrText)
    If Len(pstrText) > plngWidth Then
        If penAlign = vbRightJustify Then strReturn = Right$(pstrText, plngWidth) Else strReturn = Left$(pstrText, plngWidth)
    Else
        strReturn = Space$(plngWidth)
        If Len(pstrText) Then
            Select Case penAlign
                Case vbRightJustify: Mid$(strReturn, lngPadding + 1, Len(pstrText)) = pstrText
                Case vbCenter: Mid$(strReturn, lngPadding \ 2, Len(pstrText)) = pstrText
                Case vbLeftJustify: Mid$(strReturn, 1, Len(pstrText)) = pstrText
            End Select
        End If
    End If
    AlignText = strReturn
End Function

Private Sub BlankLine(Optional plngLines As Long = 1)
    Dim i As Long
    
    If mblnDisplay Then
        For i = 1 To plngLines
            frmMain.picBuild.Print
        Next
    Else
        mlngCurrent = mlngCurrent + plngLines
    End If
End Sub

'Start a display list with either a number or bullet.
Private Sub ListBegin(pblnNumbered As Boolean)
    With mtypList
        ' Push list onto stack
        .Lists = .Lists + 1
        ReDim Preserve .List(.Lists)
        .List(.Lists).Numbered = pblnNumbered
        With .List(.Lists)
            If mblnCodes Then
                If mlngCurrent > 1 Then
                    If Len(mstrJoin(mlngCurrent)) = 0 Then mlngCurrent = mlngCurrent - 1
                End If
                If .Numbered Then
                    mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & mstrNumberedOpen
                Else
                    mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & mstrBulletOpen
                End If
            End If
        End With
    End With
End Sub

Private Sub ListStartLine()
    Dim strPrefix As String
    Dim lngSpaces As Long
    Dim lngLevel As Long
    Dim lngLeft As Long
    
    With mtypList
        lngLevel = .Lists
        With .List(.Lists)
            .Count = .Count + 1
            If .Numbered Then strPrefix = .Count & ". " Else strPrefix = BulletPoint & " "
            lngSpaces = lngLevel * 6 - Len(strPrefix)
        End With
    End With
    If mblnDisplay Then
        frmMain.picBuild.CurrentX = lngLevel * mlngListIndent - frmMain.picBuild.TextWidth(strPrefix)
        If Not mblnColorOverride Then frmMain.picBuild.ForeColor = cfg.GetColor(cgeOutput, cveText)
        GrowClient 0, frmMain.picBuild.TextHeight(strPrefix)
        frmMain.picBuild.Print strPrefix;
    ElseIf mblnCodes And mblnLists Then
        mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & "[*]"
    Else
        mstrJoin(mlngCurrent) = Space$(lngSpaces) & strPrefix
    End If
End Sub

Private Sub ListEnd()
    Dim strClose As String
    
    With mtypList
        If mblnCodes And mblnLists Then
            If .List(.Lists).Numbered Then strClose = mstrNumberedClose Else strClose = mstrBulletClose
            If mlngCurrent > 1 Then
                If Len(mstrJoin(mlngCurrent)) = 0 Then mlngCurrent = mlngCurrent - 1
            End If
            mstrJoin(mlngCurrent) = mstrJoin(mlngCurrent) & strClose
        End If
        ' Pop list off stack
        .Lists = .Lists - 1
        If .Lists = 0 Then Erase .List Else ReDim Preserve .List(.Lists)
    End With
End Sub

Public Sub ResizeOutput()
    Dim lngOffsetX As Long
    Dim enHorizontal As FitEnum
    Dim enVertical As FitEnum
    Dim lngClientWidth As Long
    Dim lngClientHeight As Long
    Dim lngWidth As Long
    
    With frmMain
        lngOffsetX = .TextWidth(String(cfg.OutputMargin, 32))
        If mlngMaxWidth = 0 Or mlngMaxHeight = 0 Then
            .scrollHorizontal.Visible = False
            .scrollVertical.Visible = False
            Exit Sub
        End If
        ' Comically overcoded checks to only use scrollbars if absolutely needed
        enHorizontal = CheckFit(mlngMaxWidth, .ScaleWidth - lngOffsetX, .scrollVertical.Width)
        enVertical = CheckFit(mlngMaxHeight, .ScaleHeight, .scrollHorizontal.Height)
        Select Case enHorizontal
            Case feAlwaysFits: lngClientWidth = mlngMaxWidth
            Case feNeverFits: If enVertical = feAlwaysFits Then lngClientWidth = .ScaleWidth - lngOffsetX Else lngClientWidth = .ScaleWidth - .scrollVertical.Width - lngOffsetX
            Case feFitsWithoutScroll: If enVertical = feNeverFits Then lngClientWidth = .ScaleWidth - .scrollVertical.Width - lngOffsetX Else lngClientWidth = .ScaleWidth - lngOffsetX
        End Select
        Select Case enVertical
            Case feAlwaysFits: lngClientHeight = mlngMaxHeight
            Case feNeverFits: If enHorizontal = feAlwaysFits Then lngClientHeight = .ScaleHeight Else lngClientHeight = .ScaleHeight - .scrollHorizontal.Height
            Case feFitsWithoutScroll: If enHorizontal = feNeverFits Then lngClientHeight = .ScaleHeight - .scrollHorizontal.Height Else lngClientHeight = .ScaleHeight
        End Select
        ' End scrollbar check madness
        If lngClientWidth > 0 And lngClientHeight > 0 Then
            .picContainer.Move lngOffsetX, 0, lngClientWidth, lngClientHeight
            .picBuild.Move 0, 0, mlngMaxWidth, mlngMaxHeight
            With .scrollHorizontal
                .Visible = False
                If mlngMaxWidth > lngClientWidth Then
                    .Move 0, lngClientHeight, lngClientWidth + lngOffsetX, .Height
                    lngWidth = frmMain.ScaleX(mlngMaxWidth - lngClientWidth, vbTwips, vbPixels)
                    If lngWidth > 32767 Then lngWidth = 32767
                    .Max = lngWidth
                    .SmallChange = frmMain.ScaleX(frmMain.picBuild.TextWidth("XXXXXX"), vbTwips, vbPixels)
                    .LargeChange = frmMain.ScaleX(lngClientWidth, vbTwips, vbPixels)
                    .Value = 0
                    .Visible = True
                End If
            End With
            With .scrollVertical
                .Visible = False
                If mlngMaxHeight > lngClientHeight Then
                    .Move lngClientWidth + lngOffsetX, 0, .Width, lngClientHeight
                    .Max = frmMain.ScaleY(mlngMaxHeight - lngClientHeight, vbTwips, vbPixels)
                    .SmallChange = frmMain.ScaleY(frmMain.picBuild.TextHeight("X"), vbTwips, vbPixels)
                    .LargeChange = frmMain.ScaleY(lngClientHeight, vbTwips, vbPixels)
                    .Value = 0
                    .Visible = True
                End If
            End With
            .CanDrag = (.scrollHorizontal.Visible Or .scrollVertical.Visible)
        End If
    End With
End Sub

Private Function CheckFit(plngMax As Long, plngScale As Long, plngScroll As Long) As FitEnum
    If plngMax > plngScale Then
        CheckFit = feNeverFits
    ElseIf plngMax + plngScroll > plngScale Then
        CheckFit = feFitsWithoutScroll
    Else
        CheckFit = feAlwaysFits
    End If
End Function

' vBulletin inexplicably adds random spaces to very long strings of consecutive
' dots, seemingly without reason. To avoid this, only add dots to every other
' space.
'
' To make the dots appear smoother in the output window or when
' selecting text, AddDots() keeps the vertical dot columns aligned by
' only replacing odd numbered positions with dots. (1, 3, 5, etc...)
'
' Only issue is odd-length BBCodes, which don't use space in output
' but are part of the string variable. So the final Levelup28 line in the
' stats section will have mis-aligned dots because one of the columns
' will have bold tags denoting the preferred build points.
' [B][/B] = 7 characters
'
' This "problem" only occurs in the stats section for levelup28, and it's way too
' trivial to bother "fixing", especially since it would make generating output slower.
Private Sub AddDots(pstrText As String)
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim i As Long

    lngStart = 1
    Do
        ' Check for multiple spaces
        lngStart = InStr(lngStart, pstrText, "  ")
        If lngStart = 0 Then Exit Do ' None found; we're done
        ' Find out where the next non-space is
        For lngEnd = lngStart + 2 To Len(pstrText)
            If Mid$(pstrText, lngEnd, 1) <> " " Then Exit For
        Next
        ' Replace every other space with a dot
        For i = lngStart To lngEnd - 1
            If i Mod 2 = 1 Then Mid$(pstrText, i, 1) = "."
        Next
        ' Insert color tags around this blank area to hide dots
        If mblnCodes And mblnColor Then pstrText = Left$(pstrText, lngStart - 1) & mstrHideColor & Mid$(pstrText, lngStart, lngEnd - lngStart) & mstrColorClose & Mid$(pstrText, lngEnd)
        ' Start next iteration where this one ended
        lngStart = lngEnd
    Loop
    ' Final check: If leftmost character is a single space, that needs a dot as well
    If Left$(pstrText, 1) = " " Then
        Mid$(pstrText, 1, 1) = "."
        If mblnCodes And mblnColor Then pstrText = mstrHideColor & Left$(pstrText, 1) & mstrColorClose & Mid$(pstrText, 2)
    End If
    ' And also check if line starts with [courier] followed by a space (this became an issue with feat output)
    If mblnCodes And mblnFixed And Len(mstrFixedOpen) <> 0 Then
        If Left$(pstrText, Len(mstrFixedOpen) + 1) = mstrFixedOpen & " " Then pstrText = mstrFixedOpen & mstrHideColor & "." & mstrColorClose & Mid$(pstrText, Len(mstrFixedOpen) + 2)
    End If
End Sub
