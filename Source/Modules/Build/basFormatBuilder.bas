Attribute VB_Name = "basFormatBuilder"
Option Explicit

Private Type LevelFeatType
    Level As Long
    FeatName As String
    FeatType As String
End Type

Private Type LevelFeatsType
    Feat() As LevelFeatType
    Feats As Long
End Type

Private mstrGroup() As String
Private mlngGroups As Long

Private mtypFeat() As LevelFeatsType
Private mlngLevel As Long

Private mstrTag() As String
Private mlngTags As Long
Private mlngBuffer As Long

Private mstrXML As String
Private mstrTrail As String


' ************* ARRAY *************


Private Sub InitTags()
    mlngBuffer = 1023
    ReDim mstrTag(mlngBuffer)
    mlngTags = -1
    Erase mstrGroup
    mlngGroups = 0
End Sub

Private Sub OpenTag(pstrTag As String)
    mlngGroups = mlngGroups + 1
    ReDim Preserve mstrGroup(1 To mlngGroups)
    mstrGroup(mlngGroups) = pstrTag
    AddTag "<" & pstrTag & ">"
End Sub

Private Sub CloseTag()
    AddTag "</" & mstrGroup(mlngGroups) & ">"
    mlngGroups = mlngGroups - 1
    If mlngGroups = 0 Then Erase mstrGroup Else ReDim Preserve mstrGroup(1 To mlngGroups)
End Sub

Private Sub CloseTags()
    Do While mlngGroups > 0
        CloseTag
    Loop
End Sub

Private Sub TagValue(pstrTag As String, ByVal pstrValue As String)
    If Len(pstrValue) = 0 Then NullTag pstrTag Else AddTag "<" & pstrTag & ">" & pstrValue & "</" & pstrTag & ">", True
End Sub

Private Sub NullTag(pstrTag As String)
    AddTag "<" & pstrTag & "/>", True
End Sub

Private Sub AddTag(pstrTag As String, Optional pblnValue As Boolean = False)
    Dim lngSpaces As Long
    
    mlngTags = mlngTags + 1
    If mlngTags > mlngBuffer Then
        mlngBuffer = (mlngBuffer * 3) \ 2
        ReDim Preserve mstrTag(mlngBuffer)
    End If
    If pblnValue Then lngSpaces = mlngGroups Else lngSpaces = mlngGroups - 1
    If lngSpaces < 0 Then lngSpaces = 0
    lngSpaces = lngSpaces * 3
    mstrTag(mlngTags) = Space$(lngSpaces) & pstrTag
End Sub

Private Sub TrimTags()
    If mlngTags <> mlngBuffer Then ReDim Preserve mstrTag(mlngTags)
End Sub


' ************* EXPORT *************


Public Sub ExportFileBuilder()
    Dim strFile As String
    Dim strRaw As String
    Dim strValue As String
    Dim i As Long
    
    strFile = SaveAsDialog(peBuilder)
    If Len(strFile) = 0 Then Exit Sub
    InitTags
    OpenTag "DDOCharacterData"
    OpenTag "Character"
    TagValue "Name", build.BuildName
    TagValue "Alignment", BldGetAlignmentName(GetExportAlignment())
    If build.Race = reAny Then
        TagValue "Race", "Human"
    Else
        TagValue "Race", GetRaceName(build.Race)
    End If
    AddStatTags
    TagValue "TomeOfFate", 0
    OpenTag "SkillTomes"
    For i = seBalance To seUMD
        Select Case i
            Case seSpellcraft: TagValue "SpellCraft", build.SkillTome(i)
            Case seUMD: TagValue "UMD", build.SkillTome(i)
            Case Else: TagValue Replace(GetSkillName(i), " ", vbNullString), build.SkillTome(i)
        End Select
    Next
    CloseTag
    AddPastLives
    NullTag "ActiveStances"
    For i = 1 To 7
        If build.Levelups(i) = aeAny Then strValue = "Strength" Else strValue = GetStatName(build.Levelups(i))
        TagValue "Level" & i * 4, strValue
    Next
    For i = 0 To 2
        If build.BuildClass(i) = ceAny Then strValue = "Unknown" Else strValue = GetClassName(build.BuildClass(i))
        TagValue "Class" & i + 1, strValue
    Next
    OpenTag "SelectedEnhancementTrees"
    For i = 1 To 6
        TagValue "TreeName", "No Selection"
    Next
    CloseTag
    AddLevelTraining
    NullTag "ActiveEpicDestiny"
    TagValue "FatePoints", 5
    For i = 1 To 5
        OpenTag "TwistOfFate"
        TagValue "Tier", 0
        CloseTag
    Next
    TagValue "ActiveGear", "Standard"
    OpenTag "EquippedGear"
    TagValue "Name", "Standard"
    CloseTag
    CloseTags
    TrimTags
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    xp.File.SaveStringAs strFile, Join(mstrTag, vbNewLine)
    Erase mstrTag
End Sub

Private Function BldGetAlignmentName(ByVal penAlignment As AlignmentEnum) As String
    Select Case GetExportAlignment()
        Case aleNeutralGood: BldGetAlignmentName = "Neutral Good"
        Case aleLawfulNeutral: BldGetAlignmentName = "Lawful Neutral"
        Case aleLawfulGood: BldGetAlignmentName = "Lawful Good"
        Case aleChaoticNeutral: BldGetAlignmentName = "Chaotic Neutral"
        Case aleChaoticGood: BldGetAlignmentName = "Chaotic Good"
        Case Else: BldGetAlignmentName = "Neutral"
    End Select
End Function

Private Sub AddStatTags()
    Dim lngPoints As Long
    Dim strSkill As String
    Dim i As Long
    
    Select Case build.BuildPoints
        Case beAdventurer: lngPoints = 28
        Case beChampion: If build.Race = reDrow Then lngPoints = 28 Else lngPoints = 32
        Case beHero: If build.Race = reDrow Then lngPoints = 30 Else lngPoints = 34
        Case beLegend: If build.Race = reDrow Then lngPoints = 32 Else lngPoints = 36
    End Select
    OpenTag "AbilitySpend"
    TagValue "AvailableSpend", lngPoints
    For i = 1 To 6
        TagValue GetStatName(i, True) & "Spend", GetExportStatRaise(i)
    Next
    CloseTag
    TagValue "GuildLevel", 0
    For i = 1 To 6
        TagValue GetStatName(i, True) & "Tome", build.Tome(i)
    Next
End Sub

Private Sub AddPastLives()
    Dim lngLife() As Long
    Dim lngLives As Long
    Dim enClass As ClassEnum
    Dim i As Long
    
    lngLives = GetExportPastLives(lngLife)
    If lngLives = 0 Then
        NullTag "SpecialFeats"
    Else
        OpenTag "SpecialFeats"
        For enClass = 1 To ceClasses - 1
            For i = 1 To lngLife(enClass)
                OpenTag "TrainedFeat"
                TagValue "FeatName", "Past Life: " & db.Class(enClass).ClassName
                TagValue "Type", "Special"
                TagValue "LevelTrainedAt", 0
                CloseTag
            Next
        Next
        CloseTag
    End If
End Sub

Private Sub AddLevelTraining()
    Dim lngLevel As Long
    Dim enSkill As SkillsEnum
    Dim i As Long
    
    InitExportFeats
    For lngLevel = 1 To 20
        OpenTag "LevelTraining"
        If build.Class(lngLevel) = ceAny Then
            TagValue "SkillPointsAvailable", 0
            TagValue "SkillPointsSpent", 0
            NullTag "AutomaticFeats"
            NullTag "TrainedFeats"
        Else
            TagValue "Class", db.Class(build.Class(lngLevel)).ClassName
            TagValue "SkillPointsAvailable", Skill.Col(lngLevel).MaxPoints
            TagValue "SkillPointsSpent", Skill.Col(lngLevel).Points
            NullTag "AutomaticFeats"
            TrainedFeats lngLevel
            If Skill.Col(lngLevel).Points > 0 Then
                For enSkill = seBalance To seSkills - 1
                    For i = 1 To build.Skills(enSkill, lngLevel)
                        OpenTag "TrainedSkill"
                        TagValue "Skill", GetSkillName(enSkill)
                        CloseTag
                    Next
                Next
            End If
        End If
        CloseTag
    Next
    ' NEED TO DEAL WITH LEGENDARY HERE
    For lngLevel = 21 To 30  'Should be MaxLevel
        OpenTag "LevelTraining"
        TagValue "Class", "Epic"
        TagValue "SkillPointsAvailable", 0
        TagValue "SkillPointsSpent", 0
        NullTag "AutomaticFeats"
        TrainedFeats lngLevel
        CloseTag
    Next
    CloseExportFeats
End Sub

Private Sub TrainedFeats(plngLevel As Long)
    Dim i As Long
    
    With Export(plngLevel)
        If .BuilderFeats = 0 Then
            NullTag "TrainedFeats"
        Else
            OpenTag "TrainedFeats"
            For i = 1 To .Feats
                With .Feat(i)
                    If Len(.BuilderName) <> 0 And Len(.BuilderType) <> 0 Then
                        OpenTag "TrainedFeat"
                        TagValue "FeatName", .BuilderName
                        TagValue "Type", .BuilderType
                        TagValue "LevelTrainedAt", plngLevel - 1
                        CloseTag
                    End If
                End With
            Next
            CloseTag
        End If
    End With
End Sub


' ************* IMPORT *************


Public Sub ImportFileBuilder()
    Dim strFile As String
    
    strFile = OpenDialog(peBuilder)
    If Len(strFile) = 0 Then Exit Sub
    ClearBuild
    SetBuildDefaults
    ParseXML strFile
    BuildWasImported
End Sub

Private Sub ParseXML(pstrFile As String)
    InitTags
    mstrTrail = vbNullString
    mlngLevel = 0
    ReDim mtypFeat(1 To 30)
    mstrXML = xp.File.LoadToString(pstrFile)
    ParseTags 1, InStrRev(mstrXML, ">")
'    TrimTags
'    xp.File.SaveStringAs App.Path & "\Translate\Debug.txt", Join(mstrTag, vbNewLine)
    mstrXML = vbNullString
    Erase mstrTag
    ProcessFeats
    Erase mtypFeat
End Sub

' Recursive function
Private Sub ParseTags(ByVal plngLeft As Long, ByVal plngRight As Long)
    Dim lngOpenBegin As Long ' Position of "<" in open tag
    Dim lngOpenEnd As Long ' Position of ">" in open tag
    Dim lngCloseBegin As Long ' Position of "<" in close tag
    Dim lngCloseEnd As Long ' Position of ">" in close tag
    Dim strTag As String
    
    Do
        If Not FindText(lngOpenBegin, "<", plngLeft, plngRight) Then Exit Do
        If Not FindText(lngOpenEnd, ">", lngOpenBegin, plngRight) Then Exit Do
        If Mid$(mstrXML, lngOpenEnd - 1, 1) = "/" Then
            ' Null tag; skip past and continue
            lngCloseEnd = lngOpenEnd
        Else
            strTag = Mid$(mstrXML, lngOpenBegin + 1, lngOpenEnd - lngOpenBegin - 1)
            If Not FindText(lngCloseBegin, "</" & strTag & ">", lngOpenEnd, plngRight) Then Exit Do
            lngCloseEnd = lngCloseBegin + Len(strTag) + 2
            Select Case strTag
                ' Skip sections
                Case "AutomaticFeats"
                Case "EnhancementSpendInTree"
                Case "EpicDestinySpendInTree"
                Case "TwistOfFate"
                Case "TrainedSpell"
                Case Else
                    If FindText(0, "<", lngOpenEnd, lngCloseBegin - 1) Then
                        ' More tags inside this tag pair; recurse
                        Advance strTag
                        ParseTags lngOpenEnd + 1, lngCloseBegin - 1
                        Retreat strTag
                    Else
                        ProcessValue strTag, Mid$(mstrXML, lngOpenEnd + 1, lngCloseBegin - lngOpenEnd - 1)
                    End If
            End Select
        End If
        plngLeft = lngCloseEnd + 1
    Loop While plngLeft < plngRight
End Sub

Private Function FindText(plngReturn As Long, pstrText As String, plngLeft As Long, plngRight As Long) As Boolean
    plngReturn = InStr(plngLeft, mstrXML, pstrText)
    If plngReturn > 0 And plngReturn <= plngRight Then FindText = True
End Function

Private Sub Advance(pstrTag As String)
    Select Case pstrTag
        Case "DDOCharacterData", "Character"
        Case Else: mstrTrail = mstrTrail & "<" & pstrTag & ">"
    End Select
End Sub

Private Sub Retreat(pstrTag As String)
    Dim lngLen As Long
    
    lngLen = Len(mstrTrail) - Len(pstrTag) - 2
    If lngLen < 1 Then mstrTrail = vbNullString Else mstrTrail = Left$(mstrTrail, lngLen)
End Sub

Private Sub ProcessValue(pstrTag As String, pstrValue As String)
    Dim enSkill As SkillsEnum
    
    AddTag mstrTrail & "<" & pstrTag & ">" & pstrValue & "</" & pstrTag & ">"
    Select Case mstrTrail
        Case vbNullString
            ProcessOverview pstrTag, pstrValue
        Case "<AbilitySpend>"
            ProcessStats pstrTag, pstrValue
        Case "<SkillTomes>"
            enSkill = GetSkillID(pstrTag)
            If enSkill <> seAny Then build.SkillTome(enSkill) = val(pstrValue)
        Case Else
            If Left$(mstrTrail, 15) = "<LevelTraining>" Then LevelTraining pstrTag, pstrValue
    End Select
End Sub

Private Sub ProcessOverview(pstrTag As String, pstrValue As String)
    Dim lngLevelUp As Long
    Dim enStat As StatEnum
    Dim i As Long
    
    ' Levelups
    If Left$(pstrTag, 5) = "Level" Then
        lngLevelUp = val(Mid$(pstrTag, 6)) \ 4
        build.Levelups(lngLevelUp) = GetStatID(pstrValue)
        If lngLevelUp = 7 Then
            build.Levelups(0) = build.Levelups(1)
            For i = 2 To 7
                If build.Levelups(0) <> build.Levelups(i) Then
                    build.Levelups(0) = aeAny
                    Exit For
                End If
            Next
        End If
    ' Stat tomes
    ElseIf Right$(pstrTag, 4) = "Tome" Then
        enStat = GetStatID(Left$(pstrTag, 3))
        If enStat <> aeAny Then build.Tome(enStat) = val(pstrValue)
    ' Buildclass
    ElseIf Left$(pstrTag, 5) = "Class" Then
        i = val(Right$(pstrTag, 1))
        If i < 4 Then
            build.BuildClass(i - 1) = GetClassID(pstrValue)
        End If
    Else
    ' Miscellaneous
        Select Case pstrTag
            Case "Name": build.BuildName = pstrValue
            Case "Alignment": If pstrValue = "Neutral" Then build.Alignment = aleTrueNeutral Else build.Alignment = GetAlignmentID(pstrValue)
            Case "Race": build.Race = GetRaceID(pstrValue)
        End Select
    End If
End Sub

Private Sub ProcessStats(pstrTag As String, pstrValue As String)
    Dim enStat As StatEnum
    Dim enInclude As BuildPointsEnum
    Dim lngPoints As Long
    
    ' Build Points
    If pstrTag = "AvailableSpend" Then
        FilterPoints pstrValue
    Else ' Stat points spent
        enStat = GetStatID(Left$(pstrTag, 3))
        If enStat <> aeAny Then
            enInclude = build.BuildPoints
            lngPoints = GetImportStatRaise(pstrValue)
            build.StatPoints(enInclude, enStat) = lngPoints
            build.StatPoints(enInclude, 0) = build.StatPoints(enInclude, 0) + lngPoints
        End If
    End If
End Sub

' Build Points
Private Sub FilterPoints(ByVal plngValue As Long)
    Dim enBuildPoints As BuildPointsEnum
    Dim i As Long
    
    Select Case plngValue
        Case 28: enBuildPoints = beAdventurer
        Case 30: enBuildPoints = beChampion
        Case 32: If build.Race = reDrow Then enBuildPoints = beLegend Else enBuildPoints = beChampion
        Case 34: enBuildPoints = beHero
        Case 36: enBuildPoints = beLegend
    End Select
    build.BuildPoints = enBuildPoints
    For i = 0 To 3
        If i = enBuildPoints Then build.IncludePoints(i) = 1 Else build.IncludePoints(i) = 0
    Next
End Sub

Private Sub LevelTraining(pstrTag As String, pstrValue As String)
    Dim enSkill As SkillsEnum
    
    If pstrTag = "Class" Then
        mlngLevel = mlngLevel + 1
    ElseIf mlngLevel = 0 Then
        Exit Sub
    End If
    Select Case pstrTag
        Case "Class"
            If mlngLevel < 21 Then build.Class(mlngLevel) = GetClassID(pstrValue)
        Case "FeatName"
            With mtypFeat(mlngLevel)
                .Feats = .Feats + 1
                ReDim Preserve .Feat(1 To .Feats)
                .Feat(.Feats).FeatName = pstrValue
                .Feat(.Feats).Level = mlngLevel
            End With
        Case "Type"
            With mtypFeat(mlngLevel)
                If .Feats <> 0 Then .Feat(.Feats).FeatType = pstrValue
            End With
        Case "Skill"
            If mlngLevel > 0 And mlngLevel < 21 Then
                enSkill = GetSkillID(pstrValue)
                If enSkill <> seAny Then build.Skills(enSkill, mlngLevel) = build.Skills(enSkill, mlngLevel) + 1
            End If
    End Select
End Sub

Private Sub ProcessFeats()
    Dim lngLevel As Long
    Dim i As Long
    
    InitBuildFeats
    SortFeatMap peBuilder
    ' NEED TO DEAL WITH LEGENDARY HERE
    For lngLevel = 1 To 30  's/b MAXLEVEL
        For i = 1 To mtypFeat(lngLevel).Feats
            ProcessFeat mtypFeat(lngLevel).Feat(i)
        Next
    Next
End Sub

Private Sub ProcessFeat(ptypFeat As LevelFeatType)
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim strFeat As String
    Dim strSelector As String
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim lngPos As Long
    
    ' Type
    Select Case ptypFeat.FeatType
        Case "Standard", "EpicFeat", "EpicDestinyFeat": enType = bftStandard
        Case "Legendary": enType = bftLegend
        Case "FollowerOf", "ChildOf", "Deity", "BelovedOf", "DamageReduction": enType = bftDeity
        Case "HumanBonus", "Dilettante", "PDKBonus", "DragonbornRacial", "AasimarBond": enType = bftRace
        Case "ArtificerBonus": enType = GetClassType(ceArtificer)
        Case "DruidWildShape": enType = GetClassType(ceDruid)
        Case "FighterBonus": enType = GetClassType(ceFighter)
        Case "FavoredEnemy": enType = GetClassType(ceRanger)
        Case "RogueSpecialAbility": enType = GetClassType(ceRogue)
        Case "WarlockPact": enType = GetClassType(ceWarlock)
        Case "Metamagic": enType = GetClassType(ceWizard)
        Case "EnergyResistance", "FavoredSoulBattle", "FavoredSoulHeart": enType = GetClassType(ceFavoredSoul)
        Case "Domain": enType = GetClassType(ceCleric)
        Case "MonkBonus", "MonkPhilosophy", "MonkBonus6": enType = GetClassType(ceMonk)
        Case Else: enType = bftUnknown
    End Select
    If enType = bftUnknown Then Exit Sub
    ' Index
    For lngIndex = 1 To build.Feat(enType).Feats
        If build.Feat(enType).Feat(lngIndex).Level = ptypFeat.Level Then Exit For
    Next
    If lngIndex > build.Feat(enType).Feat(lngIndex).Level Then Exit Sub
    ' Feat
    lngFeat = SeekFeatMap(ptypFeat.FeatName)
    If lngFeat = 0 Then Exit Sub
    strFeat = db.FeatMap(lngFeat).Lite
    lngPos = InStr(strFeat, ":")
    If lngPos > 1 Then
        strSelector = LCase$(Trim$(Mid$(strFeat, lngPos + 1)))
        strFeat = Left$(strFeat, lngPos - 1)
    End If
    lngFeat = SeekFeat(strFeat)
    If lngFeat = 0 Then Exit Sub
    With db.Feat(lngFeat)
        If Len(strSelector) Then
            For lngSelector = .Selectors To 1 Step -1
                If LCase$(.Selector(lngSelector).SelectorName) = strSelector Then Exit For
            Next
        End If
    End With
    With build.Feat(enType).Feat(lngIndex)
        .FeatName = strFeat
        .Selector = lngSelector
    End With
End Sub

Private Function GetClassType(penClass As ClassEnum) As BuildFeatTypeEnum
    If build.BuildClass(0) = penClass Then
        GetClassType = bftClass1
    ElseIf build.BuildClass(1) = penClass Then
        GetClassType = bftClass2
    ElseIf build.BuildClass(2) = penClass Then
        GetClassType = bftClass3
    Else
        GetClassType = bftUnknown
    End If
End Function


' ************* RAW DATA *************


'Public Sub CollectFeats()
'    Dim strFile As String
'    Dim strRaw As String
'    Dim strFeat() As String
'    Dim strName() As String
'    Dim lngFeat As Long
'    Dim i As Long
'
'    ReDim strFeat(1023)
'    strFile = xp.Folder.UserDocs & "\My Games\DDO\DDO Builder\Feats.xml"
'    If Not xp.File.Exists(strFile) Then Exit Sub
'    strRaw = xp.File.LoadToString(strFile)
'    strName = Split(strRaw, "<Name>")
'    For i = 1 To UBound(strName)
'        If InStr(strName(i), "<Acquire>Train</Acquire>") Then
'            strFeat(lngFeat) = "Train: " & Left$(strName(i), InStr(strName(i), "</Name>") - 1)
'        Else
'            strFeat(lngFeat) = "Automatic: " & Left$(strName(i), InStr(strName(i), "</Name>") - 1)
'        End If
'        lngFeat = lngFeat + 1
'    Next
'    ReDim Preserve strFeat(lngFeat - 1)
'    strFile = App.Path & "\Translate\FeatsBuilder.txt"
'    If xp.File.Exists(strFile) Then xp.File.Delete strFile
'    xp.File.SaveStringAs strFile, Join(strFeat, vbNewLine)
'End Sub
