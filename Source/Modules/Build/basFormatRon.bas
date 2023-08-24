Attribute VB_Name = "basFormatRon"
Option Explicit

Private Enum RonRaceEnum
    rreHuman
    rreElf
    rreHalfling
    rreDwarf
    rreWarforged
    rreDrow
    rreHalfElf
    rreHalfOrc
    rreBladeforged
    rreMorninglord
    rrePDK
    rreShadarKai
    rreGnome
    rreDeepGnome
    rreDragonborn
    rreAasimar
    rreScourge
    rreWoodElf
End Enum

Private Enum RonClassEnum
    rceUnknown
    rceFighter
    rcePaladin
    rceBarbarian
    rceMonk
    rceRogue
    rceRanger
    rceCleric
    rceWizard
    rceSorcerer
    rceBard
    rceFavoredSoul
    rceArtificer
    rceDruid
    rceWarlock
End Enum

Private Enum RonAlignEnum
    raeLawfulGood
    raeLawfulNeutral
    raeNeutralGood
    raeNeutral
    raeChaoticGood
    raeChaoticNeutral
End Enum

Public Enum RonSlotEnum
    rseUnknown = 0
    rseStandard = 1
    rseHuman = 2
    rseFighter = 3
    rseWizard = 4
    rseFavoredEnemy = 5
    rseRogue = 6
    rseMonkBonus = 7
    rseMonkPath = 8
    rseDeity = 9
    rseFavoredSoul = 10
    rseDilettante = 15
    rseArtificer = 16
    rseDruid = 17
    rseDestiny = 18
    rseWarlockPact = 19
    rseLegend = 20
    rseDragonborn = 21
    rseDomain = 22
    rseBond = 23
End Enum

Private mstrLine() As String
Private mlngLines As Long
Private mlngBuffer As Long


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
    If mlngLines > 0 Then mlngLines = mlngLines - 1
    If mlngLines <> mlngBuffer Then ReDim Preserve mstrLine(mlngLines)
End Sub

Private Sub LoadArray(pstrFile As String)
    Dim strRaw As String
    
    strRaw = xp.File.LoadToString(pstrFile)
    mstrLine = Split(strRaw, vbNewLine)
    mlngLines = UBound(mstrLine)
End Sub


' ************* EXPORT *************


Public Sub ExportFileRon()
    Dim strFile As String
    Dim strArray() As String
    Dim i As Long
    Dim j As Long
    
    strFile = SaveAsDialog(peRon)
    If Len(strFile) = 0 Then Exit Sub
    InitArray
    AddLine "VERSION: 4.36.104;"
    AddLine "NAME: " & GetRonName() & ";"
    AddLine "RACE: " & GetRonRaceID() & ";"
    AddLine "SEX: 0;"
    AddLine "ALIGNMENT: " & GetRonAlignID() & ";"
    AddLine "CLASSRECORD: "
    For i = 1 To 20
        If build.Class(i) = ceAny Then
            AddLine "None,"
        Else
            AddLine GetClassName(build.Class(i)) & ","
        End If
    Next
    For i = 1 To 10
        If build.Class(1) = ceAny Then
            AddLine "None,"
        Else
            AddLine GetClassName(build.Class(1)) & ","
        End If
    Next
    AddLine ";"
    If build.BuildPoints = beChampion And build.Race <> reDrow Then
        AddLine "ABILITYFAVORBONUS: ;"
    End If
    AddLine AbilityRaiseLine()
    For i = 1 To MAX_LEVELUPS
        If build.Levelups(i) = aeAny Then
            AddLine "ABILITY" & i * 4 & ": 0;"
        Else
            AddLine "ABILITY" & i * 4 & ": " & build.Levelups(i) - 1 & ";"
        End If
            
    Next
    AddLine "TOMERAISE: "
    For i = 1 To 6
        AddLine build.Tome(i) & ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,"
    Next
    AddLine ";"
    AddLine "SKILLRAISE: "
    For i = 1 To 20
        ReDim strArray(21)
        For j = 0 To 20
            strArray(j) = build.Skills(j + 1, i)
        Next
        AddLine Join(strArray, ", ")
    Next
    AddLine ";"
    GetRonFeats
    AddLine "ENHANCEMENTTREELIST: "
    For i = 1 To 7
        AddLine "NoTree,"
    Next
    AddLine ";"
    AddLine "ENHANCEMENTLIST: "
    AddLine ";"
    AddLine "SPELLLIST: 0,"
    AddLine ";"
    AddLine "ITEMS: 0, "
    AddLine ";"
    AddLine "EQUIPPED: "
    AddLine "-1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, ;"
    AddLine "PASTLIFE: "
    AddLine GetRonPastLife()
    AddLine "ICONICPL: "
    AddLine "0, 0, 0, 0, 0, ;"
    AddLine "EPICPL: "
    AddLine "0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ;"
    AddLine "RACEPL: "
    AddLine "0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ;"
    TrimArray
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    xp.File.SaveStringAs strFile, Join(mstrLine, vbNewLine)
    Erase mstrLine
End Sub

Private Function GetRonName() As String
    Dim strFirst As String
    Dim strLast As String
    Dim lngPos As Long
    
    If Len(build.BuildName) = 0 Then
        GetRonName = " , "
    Else
        lngPos = InStr(build.BuildName, " ")
        If lngPos = 0 Then
            If Len(build.BuildName) > 12 Then strLast = build.BuildName Else strFirst = build.BuildName
        Else
            strFirst = Left$(build.BuildName, lngPos - 1)
            strLast = Mid$(build.BuildName, lngPos + 1)
        End If
        If Len(strFirst) > 12 Then strFirst = Left$(strFirst, 12)
        If Len(strLast) > 20 Then strLast = Left$(strLast, 20)
        GetRonName = strFirst & ", " & strLast
    End If
End Function

Private Function GetRonRaceID() As RonRaceEnum
    Select Case build.Race
        Case reDrow: GetRonRaceID = rreDrow
        Case reDwarf: GetRonRaceID = rreDwarf
        Case reElf: GetRonRaceID = rreElf
        Case reHalfling: GetRonRaceID = rreHalfling
        Case reHalfElf: GetRonRaceID = rreHalfElf
        Case reHalfOrc: GetRonRaceID = rreHalfOrc
        Case reWarforged: GetRonRaceID = rreWarforged
        Case reBladeforged: GetRonRaceID = rreBladeforged
        Case rePurpleDragonKnight: GetRonRaceID = rrePDK
        Case reMorninglord: GetRonRaceID = rreMorninglord
        Case reShadarKai: GetRonRaceID = rreShadarKai
        Case reGnome: GetRonRaceID = rreGnome
        Case reDeepGnome: GetRonRaceID = rreDeepGnome
        Case reDragonborn: GetRonRaceID = rreDragonborn
        Case reAasimar: GetRonRaceID = rreAasimar
        Case reScourge: GetRonRaceID = rreScourge
        Case reWoodElf: GetRonRaceID = rreWoodElf
        Case Else: GetRonRaceID = rreHuman
    End Select
End Function

Private Function GetRonAlignID() As RonAlignEnum
    Select Case GetExportAlignment()
        Case aleTrueNeutral: GetRonAlignID = raeNeutral
        Case aleNeutralGood: GetRonAlignID = raeNeutralGood
        Case aleLawfulNeutral: GetRonAlignID = raeLawfulNeutral
        Case aleLawfulGood: GetRonAlignID = raeLawfulGood
        Case aleChaoticNeutral: GetRonAlignID = raeChaoticNeutral
        Case aleChaoticGood: GetRonAlignID = raeChaoticGood
    End Select
End Function

Private Function AbilityRaiseLine() As String
    Dim strValue(1 To 6) As String
    Dim i As Long
    
    For i = 1 To 6
        strValue(i) = GetExportStatRaise(i)
    Next
    AbilityRaiseLine = "ABILITYRAISE: " & Join(strValue, ", ") & ";"
End Function

Private Sub GetRonFeats()
    Dim lngLevel As Long
    Dim lngFeats As Long
    Dim i As Long
    
    InitExportFeats
    For lngLevel = 1 To 30
        lngFeats = lngFeats + Export(lngLevel).RonFeats
    Next
    AddLine "FEATLIST: " & lngFeats & ","
    For lngLevel = 1 To 30
        For i = 1 To Export(lngLevel).Feats
            With Export(lngLevel).Feat(i)
                If Len(.RonName) <> 0 And .RonType <> rseUnknown Then AddLine .RonName & ", " & lngLevel & ", " & .RonType & ","
            End With
        Next
    Next
    AddLine ";"
    CloseExportFeats
End Sub

Private Function GetRonPastLife() As String
    Dim lngLife() As Long
    Dim lngLives As Long
    Dim enClass As ClassEnum
    Dim enRonClass As RonClassEnum
    Dim strLife() As String
    
    lngLives = GetExportPastLives(lngLife)
    ReDim strLife(1 To ceClasses)
    For enClass = 1 To ceClasses - 1
        enRonClass = GetRonClassID(enClass)
        strLife(enRonClass) = lngLife(enClass)
    Next
    GetRonPastLife = Join(strLife, ", ") & ";"
End Function

Private Function GetRonClassID(penClass As ClassEnum) As RonClassEnum
    Select Case penClass
        Case ceBarbarian: GetRonClassID = rceBarbarian
        Case ceBard: GetRonClassID = rceBard
        Case ceCleric: GetRonClassID = rceCleric
        Case ceFighter: GetRonClassID = rceFighter
        Case cePaladin: GetRonClassID = rcePaladin
        Case ceRanger: GetRonClassID = rceRanger
        Case ceRogue: GetRonClassID = rceRogue
        Case ceSorcerer: GetRonClassID = rceSorcerer
        Case ceWizard: GetRonClassID = rceWizard
        Case ceMonk: GetRonClassID = rceMonk
        Case ceFavoredSoul: GetRonClassID = rceFavoredSoul
        Case ceArtificer: GetRonClassID = rceArtificer
        Case ceDruid: GetRonClassID = rceDruid
        Case ceWarlock: GetRonClassID = rceWarlock
        Case Else: GetRonClassID = rceUnknown
    End Select
End Function


' ************* IMPORT *************


Public Sub ImportFileRon()
    Dim strFile As String
    
    strFile = OpenDialog(peRon)
    If Len(strFile) = 0 Then Exit Sub
    ClearBuild
    SetBuildDefaults
    LoadArray strFile
    build.BuildName = FindBuildName(strFile)
    build.Race = FindRace()
    build.Alignment = FindAlignment()
    IdentifyClasses
    IdentifyBuildPoints
    IdentifyLevelups
    IdentifyTomes
    IdentifySkills
    IdentifyFeats
    BuildWasImported
    Erase mstrLine
End Sub

Private Function FindValue(pstrField As String) As String
    Dim lngLine As Long
    Dim strReturn As String
    
    lngLine = FindLine(pstrField)
    If lngLine > mlngLines Then Exit Function
    strReturn = Trim$(Mid$(mstrLine(lngLine), Len(pstrField) + 1))
    If Right$(strReturn, 1) = ";" Then strReturn = Trim$(Left$(strReturn, Len(strReturn) - 1))
    FindValue = strReturn
End Function

Private Function FindLine(pstrField As String, Optional pblnBackwards As Boolean = False) As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngStep As Long
    Dim lngLen As Long
    Dim i As Long
    
    If pblnBackwards Then
        lngStart = mlngLines
        lngStep = -1
    Else
        lngEnd = mlngLines
        lngStep = 1
    End If
    lngLen = Len(pstrField)
    For i = lngStart To lngEnd Step lngStep
        If Left$(mstrLine(i), lngLen) = pstrField Then Exit For
    Next
    FindLine = i
End Function

Private Function FindBuildName(pstrFile As String) As String
    Dim strName As String
    
    strName = FindValue("NAME:")
    If InStr(strName, ",") Then strName = Replace(strName, ",", " ")
    strName = Trim$(strName)
    If InStr(strName, "  ") Then strName = Replace(strName, "  ", " ")
    If Len(strName) = 0 Then strName = GetNameFromFilespec(pstrFile)
    FindBuildName = strName
End Function

Private Function FindRace() As RaceEnum
    Dim strValue As String
    Dim enRonRace As RonRaceEnum
    
    strValue = FindValue("RACE:")
    enRonRace = val(strValue)
    Select Case enRonRace
        Case rreDrow: FindRace = reDrow
        Case rreDwarf: FindRace = reDwarf
        Case rreElf: FindRace = reElf
        Case rreHalfling: FindRace = reHalfling
        Case rreHalfElf: FindRace = reHalfElf
        Case rreHalfOrc: FindRace = reHalfOrc
        Case rreWarforged: FindRace = reWarforged
        Case rreBladeforged: FindRace = reBladeforged
        Case rrePDK: FindRace = rePurpleDragonKnight
        Case rreMorninglord: FindRace = reMorninglord
        Case rreShadarKai: FindRace = reShadarKai
        Case rreGnome: FindRace = reGnome
        Case rreDeepGnome: FindRace = reDeepGnome
        Case rreDragonborn: FindRace = reDragonborn
        Case rreAasimar: FindRace = reAasimar
        Case rreScourge: FindRace = reScourge
        Case rreWoodElf: FindRace = reWoodElf
        Case Else: FindRace = reHuman
    End Select
End Function

Private Function FindAlignment() As AlignmentEnum
    Dim strValue As String
    Dim enRonAlign As RonAlignEnum
    
    strValue = FindValue("ALIGNMENT:")
    enRonAlign = val(strValue)
    Select Case enRonAlign
        Case raeNeutral: FindAlignment = aleTrueNeutral
        Case raeNeutralGood: FindAlignment = aleNeutralGood
        Case raeLawfulNeutral: FindAlignment = aleLawfulNeutral
        Case raeLawfulGood: FindAlignment = aleLawfulGood
        Case raeChaoticNeutral: FindAlignment = aleChaoticNeutral
        Case raeChaoticGood: FindAlignment = aleChaoticGood
        Case Else: FindAlignment = aleAny
    End Select
End Function

Private Sub IdentifyClasses()
    Dim lngLine As Long
    Dim strClass As String
    Dim enClass As ClassEnum
    Dim i As Long
    Dim c As Long
    
    lngLine = FindLine("CLASSRECORD:")
    If lngLine > mlngLines - 20 Then Exit Sub
    For i = 1 To 20
        strClass = Trim$(mstrLine(lngLine + i))
        If Right$(strClass, 1) = "," Then strClass = Trim$(Left$(strClass, Len(strClass) - 1))
        enClass = GetClassID(strClass)
        If enClass = ceAny And build.BuildClass(0) <> ceAny Then enClass = build.BuildClass(0)
        For c = 0 To 2
            If build.BuildClass(c) = enClass Then Exit For
            If build.BuildClass(c) = ceAny Then
                build.BuildClass(c) = enClass
                Exit For
            End If
        Next
        build.Class(i) = enClass
    Next
    CalculateBAB
    InitBuildFeats
    InitBuildSpells
    InitBuildTrees
    InitLevelingGuide
End Sub

Private Sub IdentifyBuildPoints()
    Dim strList() As String
    Dim lngLine As Long
    Dim lngLives As Long
    Dim blnFavorBonus As Boolean
    Dim strValue As String
    Dim lngPoints As Long
    Dim i As Long
    
    lngLine = FindLine("PASTLIFE:", True)
    If lngLine >= 0 And lngLine < mlngLines Then
        strList = Split(mstrLine(lngLine + 1), ",")
        For i = 0 To UBound(strList)
            lngLives = lngLives + val(Trim$(strList(i)))
        Next
    End If
    lngLine = FindLine("ABILITYFAVORBONUS:")
    blnFavorBonus = (lngLine <= mlngLines)
    Select Case lngLives
        Case 0: If build.Race = reDrow Or blnFavorBonus = False Then build.BuildPoints = beAdventurer Else build.BuildPoints = beChampion
        Case 1: build.BuildPoints = beHero
        Case Else: build.BuildPoints = beLegend
    End Select
    For i = 0 To 3
        build.IncludePoints(i) = 0
    Next
    build.IncludePoints(build.BuildPoints) = 1
    strValue = FindValue("ABILITYRAISE:")
    strList = Split(strValue, ",")
    If UBound(strList) = 5 Then
        For i = 0 To 5
            lngPoints = GetImportStatRaise(val(Trim$(strList(i))))
            build.StatPoints(build.BuildPoints, i + 1) = lngPoints
            build.StatPoints(build.BuildPoints, 0) = build.StatPoints(build.BuildPoints, 0) + lngPoints
        Next
    End If
End Sub

Private Sub IdentifyLevelups()
    Dim strValue As String
    Dim enStat As StatEnum
    Dim blnMixed As Boolean
    Dim i As Long
    
    For i = 1 To MAX_LEVELUPS
        strValue = FindValue("ABILITY" & i * 4 & ":")
        enStat = val(strValue) + 1
        build.Levelups(i) = enStat
        If build.Levelups(0) = aeAny Then build.Levelups(0) = enStat
        If build.Levelups(0) <> enStat Then blnMixed = True
    Next
    If blnMixed Then build.Levelups(0) = aeAny
End Sub

Private Sub IdentifyTomes()
    Dim lngLine As Long
    Dim strLevel() As String
    Dim lngStat As Long
    Dim lngTome As Long
    Dim i As Long
    
    lngLine = FindLine("TOMERAISE:")
    If lngLine = 0 Or lngLine > mlngLines - 6 Then Exit Sub
    For lngStat = 1 To 6
        lngTome = 0
        strLevel = Split(mstrLine(lngLine + lngStat), ",")
        For i = 0 To UBound(strLevel)
            lngTome = lngTome + val(Trim$(strLevel(i)))
        Next
        If lngTome > tomes.Stat.Max Then lngTome = tomes.Stat.Max
        build.Tome(lngStat) = lngTome
    Next
End Sub

Private Sub IdentifySkills()
    Dim lngLine As Long
    Dim lngLevel As Long
    Dim strSkill() As String
    Dim i As Long
    
    lngLine = FindLine("SKILLRAISE:")
    If lngLine = 0 Or lngLine > mlngLines - 20 Then Exit Sub
    For lngLevel = 1 To 20
        strSkill = Split(mstrLine(lngLine + lngLevel), ",")
        If UBound(strSkill) >= 20 Then
            For i = 0 To 20
                build.Skills(i + 1, lngLevel) = val(Trim$(strSkill(i)))
            Next
        End If
    Next
End Sub

Private Sub IdentifyFeats()
    Dim lngLine As Long
    Dim strFeat As String
    Dim lngSelector As Long
    Dim enType As BuildFeatTypeEnum
    Dim lngSlot As Long
    
    lngLine = FindLine("FEATLIST:")
    If lngLine = 0 Then Exit Sub
    SortFeatMap peRon
    For lngLine = lngLine + 1 To mlngLines
        If InStr(mstrLine(lngLine), ";") Then Exit For
        If RonSplitFeat(mstrLine(lngLine), strFeat, lngSelector) Then
            If RonFindSlot(mstrLine(lngLine), enType, lngSlot) Then
                With build.Feat(enType).Feat(lngSlot)
                    .FeatName = strFeat
                    .Selector = lngSelector
                End With
            End If
        End If
    Next
End Sub

Private Function RonSplitFeat(pstrRaw As String, pstrFeat As String, plngSelector As Long) As Boolean
    Dim lngPos As Long
    Dim lngFeat As Long
    Dim strSelector As String
    Dim lngSelector As Long
    
    lngPos = InStr(pstrRaw, ",")
    If lngPos = 0 Then Exit Function
    pstrFeat = Left$(pstrRaw, lngPos - 1)
    lngFeat = SeekFeatMap(pstrFeat)
    If lngFeat = 0 Then Exit Function
    pstrFeat = db.FeatMap(lngFeat).Lite
    lngPos = InStr(pstrFeat, ":")
    If lngPos Then
        strSelector = Trim$(Mid$(pstrFeat, lngPos + 1))
        pstrFeat = Left$(pstrFeat, lngPos - 1)
    End If
    lngFeat = SeekFeat(pstrFeat)
    If lngFeat = 0 Then Exit Function
    If Len(strSelector) Then
        With db.Feat(lngFeat)
            For lngSelector = .Selectors To 1 Step -1
                If .Selector(lngSelector).SelectorName = strSelector Then Exit For
            Next
        End With
    End If
    plngSelector = lngSelector
    RonSplitFeat = True
End Function

Private Function RonFindSlot(pstrRaw As String, penType As BuildFeatTypeEnum, plngSlot As Long) As Boolean
    Dim strToken() As String
    Dim lngLevel As Long
    Dim enClass As ClassEnum
    Dim i As Long

    strToken = Split(pstrRaw, ",")
    If UBound(strToken) < 2 Then Exit Function
    Select Case val(Trim$(strToken(2)))
        Case rseStandard, rseDestiny: penType = bftStandard
        Case rseDeity: penType = bftDeity
        Case rseLegend: penType = bftLegend
        Case rseHuman, rseDilettante, rseDragonborn: penType = bftRace
        Case rseFighter: enClass = ceFighter
        Case rseWizard: enClass = ceWizard
        Case rseFavoredEnemy: enClass = ceRanger
        Case rseRogue: enClass = ceRogue
        Case rseMonkBonus, rseMonkPath: enClass = ceMonk
        Case rseFavoredSoul: enClass = ceFavoredSoul
        Case rseArtificer: enClass = ceArtificer
        Case rseDruid: enClass = ceDruid
        Case rseWarlockPact: enClass = ceWarlock
        Case rseDomain: enClass = ceCleric
        Case Else: penType = bftUnknown
    End Select
    If enClass <> ceAny Then
        Select Case CByte(enClass)
            Case build.BuildClass(0): penType = bftClass1
            Case build.BuildClass(1): penType = bftClass2
            Case build.BuildClass(2): penType = bftClass3
            Case Else: penType = bftUnknown
        End Select
    End If
    If penType = bftUnknown Then Exit Function
    lngLevel = val(Trim$(strToken(1)))
    With build.Feat(penType)
        For plngSlot = .Feats To 1 Step -1
            If .Feat(plngSlot).Level = lngLevel Then Exit For
        Next
    End With
    If plngSlot Then RonFindSlot = True
End Function


' ************* RAW DATA *************


'Private Type RonFeatType
'    Feat As String
'    Parent As String
'End Type
'
'Public Sub RonFeatList()
'    Dim strFile As String
'    Dim strRaw As String
'    Dim strLine() As String
'    Dim typRon() As RonFeatType
'    Dim typNew As RonFeatType
'    Dim typBlank As RonFeatType
'    Dim lngCurrent As Long
'    Dim i As Long
'
'    strFile = xp.Folder.UserDocs & "\My Games\DDO\Chargen\DataFiles\FeatsFile.txt"
'    If Not xp.File.Exists(strFile) Then
'        Debug.Print "FeatsFile.txt not found"
'        Exit Sub
'    End If
'    strRaw = xp.File.LoadToString(strFile)
'    strLine = Split(strRaw, vbNewLine)
'    ReDim typRon(UBound(strLine))
'    For i = 0 To UBound(strLine)
'        If Left$(strLine(i), 9) = "FEATNAME:" Then
'            typNew.Feat = Mid$(strLine(i), 11)
'        ElseIf Left$(strLine(i), 10) = "FFEATNAME:" Then
'            typNew.Feat = Mid$(strLine(i), 12)
'        End If
'        If Len(typNew.Feat) Then
'            If Right$(typNew.Feat, 1) = ";" Then typNew.Feat = Left$(typNew.Feat, Len(typNew.Feat) - 1)
'            If i > 0 Then
'                If Left$(strLine(i - 1), 14) = "PARENTHEADING:" Then
'                    typNew.Parent = Mid$(strLine(i - 1), 16)
'                    If Right$(typNew.Parent, 1) = ";" Then typNew.Parent = Left$(typNew.Parent, (Len(typNew.Parent) - 1))
'                End If
'            End If
'            typRon(lngCurrent) = typNew
'            typNew = typBlank
'            lngCurrent = lngCurrent + 1
'        End If
'    Next
'    ReDim strLine(lngCurrent - 1)
'    For i = 0 To lngCurrent - 1
'        If Len(typRon(i).Parent) Then strLine(i) = typRon(i).Parent & ": " & typRon(i).Feat Else strLine(i) = typRon(i).Feat
'    Next
'    strFile = DataPath() & "FeatsRon.txt"
'    If xp.File.Exists(strFile) Then xp.File.Delete strFile
'    xp.File.SaveStringAs strFile, Join(strLine, vbNewLine)
'End Sub
'
