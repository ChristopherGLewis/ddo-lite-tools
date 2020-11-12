Attribute VB_Name = "basDataLoad"
' Written by Ellis Dee
' Import routines to load all *.txt data files
Option Explicit

Public Sub InitData()
    Dim typBlank As DatabaseType
    
    db = typBlank
    LoadRaces
    LoadTemplates
    LoadSpells
    LoadClasses
    LoadFeats
    LoadFeatMap
    LoadEnhancements
    LoadDestinies
    LoadTomeData
    LoadNameChanges
    db.Loaded = True
End Sub

Public Function DataPath() As String
    DataPath = App.Path & "\Data\Builder\"
End Function

Private Function DataFile(pstrFile As String, pstrName As String) As Boolean
    log.Activity = actOpenFile
    log.LoadFile = pstrName
    Do
        If pstrName = "Templates.txt" Then
            pstrFile = UserTemplates()
            If xp.File.Exists(pstrFile) Then
                If Not UserTemplateOlder() Then Exit Do
            End If
        End If
        pstrFile = DataPath() & pstrName
        If Not xp.File.Exists(pstrFile) Then
            DataFile = True
            LogError
        End If
    Loop Until True
    log.Activity = actReadFile
End Function

Public Function UserTemplateOlder() As Boolean
    Dim dtmSystem As Date
    Dim dtmUser As Date
    
    If xp.File.Exists(UserTemplates()) Then
        dtmUser = FileDateTime(UserTemplates())
        If xp.File.Exists(SystemTemplates()) Then
            dtmSystem = FileDateTime(SystemTemplates())
            If dtmUser < dtmSystem Then UserTemplateOlder = True
        End If
    End If
End Function

Public Function SystemTemplates() As String
    SystemTemplates = DataPath() & "Templates.txt"
End Function

Public Function UserTemplates() As String
    UserTemplates = cfg.LitePath & "\Templates.txt"
End Function


' ************* RACES *************


Private Sub LoadRaces()
    Dim strFile As String
    Dim strRaw As String
    Dim strRace() As String
    Dim i As Long
    
    ReDim db.Race(reRaces - 1)
    If DataFile(strFile, "Races.txt") Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strRace = Split(strRaw, "RaceName: ")
    For i = 1 To UBound(strRace)
        If InStr(strRace(i), "Stats: ") Then LoadRace strRace(i) Else ErrorLoading strRace(i)
    Next
End Sub

Private Sub LoadRace(ByVal pstrRaw As String)
    Dim enRace As RaceEnum
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngID As Long
    Dim typNew As RaceType
    Dim lngPos As Long
    Dim strFeat As String
    Dim strSelector As Long
    Dim lngLevel As Long
    Dim blnError As Boolean
    Dim i As Long
    
    log.Activity = actFindRace
    CleanText pstrRaw  'Remove trailing CRLF
    strLine = Split(pstrRaw, vbNewLine) 'Make an array
    log.LoadItem = Trim$(strLine(0))
    enRace = GetRaceID(log.LoadItem)
    If enRace = reAny Then
        LogError
        Exit Sub
    End If
    log.Activity = actReadFile
    With typNew
        .RaceID = enRace
        .RaceName = Trim$(strLine(0))
        .Abbreviation = .RaceName
        log.LoadItem = .RaceName
        enRace = GetRaceID(.RaceName)
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.HasError = False
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "stats"
                        blnError = False
                        If lngListMax <> 5 Then
                            blnError = True
                        Else
                            For i = 0 To 5
                                .Stats(i + 1) = Val(strList(i))
                                Select Case .Stats(i + 1)
                                    Case 6, 8, 10
                                    Case Else: blnError = True
                                End Select
                            Next
                        End If
                        If blnError Then LogError
                    Case "abbreviation"
                        .Abbreviation = strItem
                    Case "type"
                        .Type = GetRaceTypeID(strItem)
                        If .Type = rteUnknown Then LoadError strLine(0) & " has invalid Type: " & strItem
                    Case "iconicclass"
                        .IconicClass = GetClassID(strItem)
                        If .IconicClass = ceAny Then LoadError strLine(0) & " has invalid IconicClass: " & strItem
                    Case "subrace"
                        .SubRace = GetRaceID(strItem)
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "bonusfeat"
                                    .BonusFeat = True
                                Case "bonusskill"
                                    .SkillPoints = 1
                                Case "listfirst"
                                    .ListFirst = True
                                Case Else
                                    LoadError strLine(0) & " has invalid Flag: " & strList(i)
                            End Select
                        Next
                    Case "trees"
                        .Trees = lngListMax + 1
                        ReDim .Tree(1 To .Trees)
                        For i = 0 To lngListMax
                            .Tree(i + 1) = strList(i)
                        Next
                    Case Else
                        If Left$(strField, 12) = "grantedfeats" Then
                            LoadGrantedFeats .GrantedFeat, .GrantedFeats, strField, strList, lngListMax
                        Else
                            LogError
                        End If
                End Select
            End If
        Next
    End With
    db.Race(enRace) = typNew
End Sub

Private Sub LoadGrantedFeats(ptypGrantedFeat() As PointerType, plngGrantedFeats As Long, ByVal pstrField As String, pstrList() As String, plngListMax As Long)
    Dim lngLevel As Long
    Dim lngStart As Long
    Dim i As Long
    
    lngLevel = Val(Mid$(pstrField, 13))
    If lngLevel < 1 Or lngLevel > MaxLevel Then
        LogError
        Exit Sub
    End If
    lngStart = plngGrantedFeats + 1
    plngGrantedFeats = plngGrantedFeats + plngListMax + 1
    ReDim Preserve ptypGrantedFeat(1 To plngGrantedFeats)
    For i = 0 To plngListMax
        With ptypGrantedFeat(lngStart + i)
            .Tier = lngLevel
            .Raw = "Feat: " & pstrList(i)
        End With
    Next
End Sub


' ************* CLASSES *************


Private Sub LoadClasses()
    Dim strFile As String
    Dim strRaw As String
    Dim strClass() As String
    Dim i As Long
    
    ReDim db.Class(ceClasses - 1)
    For i = 0 To 6
        db.Class(0).Alignment(i) = True
    Next
    If DataFile(strFile, "Classes.txt") Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strClass = Split(strRaw, "ClassName: ")
    For i = 1 To UBound(strClass)
        If InStr(strClass(i), "BAB: ") Then LoadClass strClass(i) Else ErrorLoading strClass(i)
    Next
End Sub

Private Sub LoadClass(ByVal pstrRaw As String)
    Dim enClass As ClassEnum
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngID As Long
    Dim lngClassLevel As Long
    Dim lngSpellLevel As Long
    Dim typNew As ClassType
    Dim lngLevel As Long
    Dim i As Long
    
    log.Activity = actFindClass
    typNew.Alignment(0) = True
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    log.LoadItem = Trim$(strLine(0))
    enClass = GetClassID(log.LoadItem)
    If enClass = ceAny Then
        LogError
        Exit Sub
    End If
    log.Activity = actReadFile
    With typNew
        log.Class = enClass
        .ClassID = enClass
        .ClassName = GetClassName(enClass)
        .Abbreviation = .ClassName
        ReDim .Initial(0)
        .Initial(0) = Left$(.ClassName, 1)
        log.LoadItem = .ClassName
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.HasError = False
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "abbreviation"
                        .Abbreviation = strItem
                    Case "initial"
                        If lngListMax = 3 Then .Initial = strList Else LogError
                    Case "color"
                        Select Case LCase$(strItem)
                            Case "red": .Color = cveRed
                            Case "green": .Color = cveGreen
                            Case "blue": .Color = cveBlue
                            Case "yellow": .Color = cveYellow
                            Case "purple": .Color = cvePurple
                            Case "orange": .Color = cveOrange
                            Case Else
                                LoadError strLine(0) & " has invalid color: " & strItem
                        End Select
                    Case "alignment"
                        For i = 0 To lngListMax
                            lngID = GetAlignmentID(strList(i))
                            If lngID <> aleAny Then
                                .Alignment(lngID) = True
                            Else
                                LoadError strLine(0) & " has invalid aligment: " & strList(i)
                            End If
                        Next
                    Case "bab"
                        Select Case strItem
                            Case "0.75": .BAB = beThreeQuarters
                            Case "0.5": .BAB = beHalf
                            Case "1": .BAB = beFull
                            Case Else: LoadError strLine(0) & " has invalid BAB: " & strItem
                        End Select
                    Case "skillpoints"
                        .SkillPoints = Val(strItem)
                        Select Case .SkillPoints
                            Case 2, 4, 6, 8
                            Case Else: LoadError strLine(0) & " has invalid skill points: " & strItem
                        End Select
                    Case "skills"
                        For i = 0 To lngListMax
                            lngID = GetSkillID(strList(i))
                            If lngID <> seAny Then .NativeSkills(lngID) = True Else LoadError strLine(0) & " has invalid skill: " & strList(i)
                        Next
                    Case "bonusfeat"
                        For i = 0 To lngListMax
                            lngLevel = Val(strList(i))
                            Select Case lngLevel
                                Case 1 To 20: .BonusFeat(lngLevel) = bfsClass
                                Case Else: LogError
                            End Select
                        Next
                    Case "classfeat"
                        For i = 0 To lngListMax
                            lngLevel = Val(strList(i))
                            Select Case lngLevel
                                Case 1 To 20: .BonusFeat(lngLevel) = bfsClassOnly
                                Case Else: LogError
                            End Select
                        Next
                    Case "trees"
                        .Trees = lngListMax + 1
                        ReDim .Tree(1 To .Trees)
                        For i = 0 To lngListMax
                            .Tree(i + 1) = strList(i)
                        Next
                    Case "maxspelllevel"
                        .MaxSpellLevel = Val(strItem)
                        Select Case .MaxSpellLevel
                            Case 4, 6, 9
                            Case Else: LogError
                        End Select
                        If .MaxSpellLevel > 0 Then
                            ReDim .SpellList(1 To .MaxSpellLevel)
                            ReDim .SpellSlots(1 To 20, 1 To .MaxSpellLevel)
                        End If
                    Case "healingspell"
                        .CanCastSpell(0) = Val(strItem)
                        If .CanCastSpell(0) < 1 Or .CanCastSpell(0) > 20 Then LogError
                    Case "freespells"
                        .FreeSpells = lngListMax + 1
                        ReDim .FreeSpell(1 To .FreeSpells)
                        For i = 0 To lngListMax
                            .FreeSpell(i + 1) = strList(i)
                        Next
                    Case "mandatoryspells"
                        .MandatorySpells = lngListMax + 1
                        ReDim .MandatorySpell(1 To .MandatorySpells)
                        For i = 0 To lngListMax
                            .MandatorySpell(i + 1) = strList(i)
                        Next
                    Case Else
                        If Left$(strField, 10) = "spellslots" Then
                            lngSpellLevel = Val(Mid$(strField, 11, 1))
                            If lngSpellLevel = 0 Then
                                ' Ignore header row
                            ElseIf lngSpellLevel > .MaxSpellLevel Then
                                LoadError strLine(0) & " has SpellSlots" & lngSpellLevel & " line but MaxSpellLevel is " & .MaxSpellLevel
                            ElseIf lngListMax <> 19 Then
                                LogError
                            Else
                                For i = 0 To lngListMax
                                    lngClassLevel = i + 1
                                    .SpellSlots(lngClassLevel, lngSpellLevel) = Val(strList(i))
                                    If .CanCastSpell(lngSpellLevel) = 0 And Val(strList(i)) <> 0 Then .CanCastSpell(lngSpellLevel) = lngClassLevel
                                Next
                            End If
                        ElseIf Left$(strField, 9) = "spelllist" Then
                            lngSpellLevel = Val(Mid$(strField, 10, 1))
                            If lngSpellLevel < 1 Then
                                LogError
                            ElseIf lngSpellLevel > .MaxSpellLevel Then
                                LoadError strLine(0) & " has SpellList" & lngSpellLevel & " line but MaxSpellLevel is " & .MaxSpellLevel
                            Else
                                With .SpellList(lngSpellLevel)
                                    .Spells = lngListMax + 1
                                    ReDim .Spell(1 To .Spells)
                                    For i = 0 To lngListMax
                                        .Spell(i + 1) = strList(i)
                                    Next
                                End With
                            End If
                        ElseIf Left$(strField, 12) = "grantedfeats" Then
                            LoadGrantedFeats .GrantedFeat, .GrantedFeats, strField, strList, lngListMax
                        ElseIf Left$(strField, 10) = "pactspells" Then
                            .Pacts = .Pacts + 1
                            ReDim Preserve .Pact(1 To .Pacts)
                            With .Pact(.Pacts)
                                .PactName = strList(0)
                                ReDim .Spells(1 To lngListMax)
                                For i = 1 To lngListMax
                                    .Spells(i) = strList(i)
                                Next
                            End With
                        Else
                            LogError
                        End If
                End Select
            End If
        Next
    End With
    db.Class(enClass) = typNew
End Sub


' ************* SPELLS *************


Private Sub LoadSpells()
    Dim strFile As String
    Dim strRaw As String
    Dim strSpell() As String
    Dim i As Long
    
    Erase db.Spell
    db.Spells = 0
    If DataFile(strFile, "Spells.txt") Then Exit Sub
    ' Allocate enough space that we never have to increase
    ReDim db.Spell(1023)
    strRaw = xp.File.LoadToString(strFile)
    strSpell = Split(strRaw, "SpellName: ")
    For i = 1 To UBound(strSpell)
        LoadSpell strSpell(i)
    Next
    With db
        If .Spells = 0 Then Erase .Spell Else ReDim Preserve .Spell(.Spells)
    End With
    SortSpells
End Sub

Private Sub LoadSpell(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngID As Long
    Dim typNew As SpellType
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    log.LoadItem = Trim$(strLine(0))
    With typNew
        .SpellName = log.LoadItem
        .Wiki = .SpellName
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.HasError = False
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "descrip"
                        .Descrip = strItem
                    Case "wikiname"
                        .Wiki = strItem
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "rare": .Rare = True
                                Case Else: LoadError strLine(0) & " has invalid Flag: " & strList(i)
                            End Select
                        Next
                    Case Else
                        LogError
                End Select
            End If
        Next
    End With
    With db
        .Spells = .Spells + 1
        .Spell(.Spells) = typNew
    End With
End Sub

Public Sub SortSpells()
    Dim i As Long
    Dim j As Long
    Dim typSwap As SpellType
    
    With db
        For i = 2 To db.Spells
            typSwap = db.Spell(i)
            For j = i To 2 Step -1
                If typSwap.SpellName < db.Spell(j - 1).SpellName Then db.Spell(j) = db.Spell(j - 1) Else Exit For
            Next j
            db.Spell(j) = typSwap
        Next
    End With
End Sub

' Simple binary search
Public Function SeekSpell(pstrSpellName As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    'Binary search into spell list - Spells has to be sorted.
    lngFirst = 1
    lngLast = db.Spells
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Spell(lngMid).SpellName < pstrSpellName Then
            lngFirst = lngMid + 1
        ElseIf db.Spell(lngMid).SpellName > pstrSpellName Then
            lngLast = lngMid - 1
        Else
            SeekSpell = lngMid
            Exit Function
        End If
    Loop
End Function

Public Sub SortSpellWiki()
    Dim i As Long
    Dim j As Long
    Dim typSwap As SpellType
    
    With db
        For i = 2 To db.Spells
            typSwap = db.Spell(i)
            For j = i To 2 Step -1
                If typSwap.Wiki < db.Spell(j - 1).Wiki Then db.Spell(j) = db.Spell(j - 1) Else Exit For
            Next j
            db.Spell(j) = typSwap
        Next
    End With
End Sub

' Simple binary search
Public Function SeekSpellWiki(pstrWiki As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Spells
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Spell(lngMid).Wiki < pstrWiki Then
            lngFirst = lngMid + 1
        ElseIf db.Spell(lngMid).Wiki > pstrWiki Then
            lngLast = lngMid - 1
        Else
            SeekSpellWiki = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* TEMPLATES *************


Public Sub LoadTemplates()
    Dim strFile As String
    Dim strRaw As String
    Dim strTemplate() As String
    Dim i As Long
    
    Erase db.Template
    db.Templates = 0
    If DataFile(strFile, "Templates.txt") Then Exit Sub
    ' Allocate enough space that we never have to increase
    ReDim db.Template(127)
    strRaw = xp.File.LoadToString(strFile)
    strTemplate = Split(strRaw, "Class: ")
    For i = 1 To UBound(strTemplate)
        LoadTemplate strTemplate(i)
    Next
    With db
        If .Templates = 0 Then Erase .Template Else ReDim Preserve .Template(.Templates)
    End With
End Sub

Private Sub LoadTemplate(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim enClass As ClassEnum
    Dim enStat As StatEnum
    Dim i As Long
    Dim typNew As TemplateType
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        log.LoadItem = strLine(0)
        .Class = GetClassID(strLine(0))
        If .Class = ceAny Then LogError
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.LoadLine = strLine(lngLine)
            If ParseTemplate(strLine(lngLine), strField, strItem) Then
                Select Case strField
                    Case "flags"
                        strList = Split(LCase$(strItem), ",")
                        For i = 0 To UBound(strList)
                            Select Case Trim$(strList(i))
                                Case "traps": .Trapping = True
                                Case "always": .Always = True
                                Case Else: LogError
                            End Select
                        Next
                    Case "caption"
                        .Caption = strItem
                        log.LoadItem = strLine(0) & ": " & strItem
                        If .Trapping Then log.LoadItem = log.LoadItem & " (Traps)"
                    Case "descrip"
                        .Descrip = strItem
                    Case "levelups"
                        .Levelups = GetStatID(strItem)
                        If .Levelups = aeAny Then LogError
                    Case "warning"
                        .Warning = strItem
                    Case "str", "dex", "con", "int", "wis", "cha"
                        enStat = GetStatID(strField)
                        For i = 0 To 4
                            .StatPoints(i, enStat) = Val(Mid$(strItem, (i * 3) + 1, 2))
                        Next
                    Case "pts"
                    Case Else
                        LogError
                End Select
            End If
        Next
    End With
    With db
        .Templates = .Templates + 1
        .Template(.Templates) = typNew
    End With
End Sub

Private Function ParseTemplate(ByVal pstrLine As String, pstrField As String, pstrItem As String) As Boolean
    Dim lngPos As Long
    Dim strValue As String
    Dim i As Long
    
    ' Prep
    pstrField = vbNullString
    pstrItem = vbNullString
    pstrLine = Trim$(pstrLine)
    If Len(pstrLine) = 0 Then Exit Function
    ' Field
    lngPos = InStr(pstrLine, ":")
    If lngPos = 0 Then
        LogError
        Exit Function
    End If
    ParseTemplate = True
    pstrField = LCase$(Trim$(Left$(pstrLine, lngPos - 1)))
    pstrItem = Mid$(pstrLine, lngPos + 2)
End Function


' ************* FEATS *************


Private Sub LoadFeats()
    Dim strFile As String
    Dim strRaw As String
    Dim strFeat() As String
    Dim i As Long
    
    Erase db.Feat
    db.Feats = 0
    If DataFile(strFile, "Feats.txt") Then Exit Sub
    ' Allocate enough space that we never have to increase
    ReDim db.Feat(511)
    strRaw = xp.File.LoadToString(strFile)
    strFeat = Split(strRaw, "FeatName: ")
    For i = 1 To UBound(strFeat)
        If InStr(strFeat(i), "Group: ") Then LoadFeat strFeat(i) Else ErrorLoading strFeat(i)
    Next
    With db
        If .Feats = 0 Then
            Erase .Feat
            Erase .FeatLookup
            Erase .FeatDisplay
        Else
            ReDim Preserve .Feat(.Feats)
            IndexFeatLookup
            IndexFeatDisplay
        End If
    End With
End Sub

Private Sub LoadFeat(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strFeat As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngGroupID As Long
    Dim typNew As FeatType
    Dim enRace As RaceEnum
    Dim enClass As ClassEnum
    Dim blnError As Boolean
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    strFeat = Trim$(strLine(0))
    With typNew
        .FeatName = strFeat
        .Abbreviation = strFeat
        .SortName = strFeat
        .Wiki = strFeat
        .Selectable = True
        ReDim .ClassBonus(ceClasses - 1)
        ReDim .RaceBonus(reRaces - 1)
        ReDim .Race(reRaces - 1)
        ReDim .Class(ceClasses - 1)
        ReDim .ClassLevel(ceClasses - 1)
        ReDim .Req(3)
        ReDim .Group(feFilters - 1)
        .Group(feAll) = True
        log.LoadItem = .FeatName
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.HasError = False
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                blnError = False
                Select Case strField
                    Case "abbreviation"
                        .Abbreviation = strItem
                    Case "sortname"
                        .SortName = strItem
                    Case "descrip"
                        .Descrip = strItem
                    Case "wikiname"
                        .Wiki = strItem
                    Case "group"
                        .Group(feAll) = True
                        For i = 0 To lngListMax
                            lngGroupID = GetGroupID(strList(i))
                            If lngGroupID = feAll Then
                                LoadError strFeat & " has invalid group: " & strList(i)
                            Else
                                Select Case lngGroupID
                                    Case feAll + 1 To feFilters - 1: .Group(lngGroupID) = True
                                End Select
                            End If
                        Next
                    Case "bab"
                        If lngValue < 1 Or lngValue > 20 Then LogError Else .BAB = lngValue
                    Case "repeat"
                        Select Case lngValue
                            ' Updated 2020.10.07 to allow for 2 ranks of Alchemical Studies: xxx
                            Case 1, 2, 3, 99: .Times = lngValue
                            Case Else: LogError
                        End Select
                    Case "stat"
                        .Stat = GetStatID(strItem)
                        If .Stat = aeAny Then LogError Else .StatValue = lngValue
                    Case "skill"
                        .Skill = GetSkillID(strItem)
                        If .Skill = seAny Then LogError Else .SkillValue = lngValue
                    Case "class"
                        .Class(0) = True
                        For i = 0 To lngListMax
                            If ParseClassLevel(strList(i), enClass, lngValue) Then
                                .Class(enClass) = True
                                .ClassLevel(enClass) = lngValue
                            Else
                                blnError = True
                            End If
                        Next
                    Case "grantedby"
                        enClass = GetClassID(strItem)
                        If enClass = ceAny Or lngValue < 1 Or lngValue > 20 Then
                            LogError
                        Else
                            .GrantedBy.Class = enClass
                            .GrantedBy.ClassLevels = lngValue
                        End If
                    Case "classbonuslevel"
                        enClass = GetClassID(strItem)
                        If enClass = ceAny Or lngValue < 1 Or lngValue > 20 Then
                            LogError
                        Else
                            .ClassBonusLevel.Class = enClass
                            .ClassBonusLevel.ClassLevels = lngValue
                        End If
                    Case "level"
                        If lngValue < 1 Or lngValue > MaxLevel Then LogError Else .Level = lngValue
                    Case "race"
                        .Race(0) = GetRaceReqID(strList(0))
                        If .Race(0) = rreAny Then
                            blnError = True
                        Else
                            For i = 1 To lngListMax
                                enRace = GetRaceID(strList(i))
                                If enRace = reAny Then blnError = True Else .Race(enRace) = 1
                            Next
                        End If
                    Case "cancastspell"
                        If lngValue < 0 Or lngValue > 9 Then
                            LogError
                        Else
                            .CanCastSpell = True
                            .CanCastSpellLevel = lngValue
                        End If
                    Case "racebonus"
                        .RaceBonus(0) = True
                        For i = 0 To lngListMax
                            enRace = GetRaceID(strList(i))
                            If enRace <> reAny Then .RaceBonus(enRace) = True Else LoadError strFeat & vbNewLine & "RaceBonus includes invalid race: " & strList(i)
                        Next
                    Case "classbonus"
                        .ClassBonus(0) = True
                        For i = 0 To lngListMax
                            enClass = GetClassID(strList(i))
                            If enClass = ceAny Then blnError = True Else .ClassBonus(enClass) = True
                        Next
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "selectoronly"
                                    .SelectorOnly = True
                                Case "deity"
                                    .Deity = True
                                Case "classonly"
                                    .ClassOnly = True
                                    ReDim .ClassOnlyClasses(ceClasses - 1)
                                Case "raceonly"
                                    .RaceOnly = True
                                Case "pastlife"
                                    .PastLife = True
                                Case "legend"
                                    .Legend = True
                                Case "unselectable"
                                    .Selectable = False
                                Case "skilltome"
                                    .SkillTome = True
                                Case "pact"
                                    .Pact = True
                                Case "domain"
                                    .Domain = True
                                Case Else
                                    blnError = True
                            End Select
                        Next
                    Case "classonlyclass"
                        For i = 0 To lngListMax
                            enClass = GetClassID(strList(i))
                            If enClass = ceAny Then blnError = True Else .ClassOnlyClasses(enClass) = True
                        Next
                    Case "classonlylevel"
                        For i = 0 To lngListMax
                            lngValue = Val(strList(i))
                            If lngValue < 1 Or lngValue > 20 Then blnError = True Else .ClassOnlyLevels(lngValue) = True
                        Next
                    Case "all", "one", "none"
                        With .Req(GetReqGroupID(strField))
                            .Reqs = lngListMax + 1
                            ReDim .Req(1 To .Reqs)
                            For i = 0 To lngListMax
                                .Req(i + 1).Raw = "Feat: " & strList(i)
                                .Req(i + 1).Style = peFeat
                            Next
                        End With
                    Case "selector"
                        If .SelectorStyle = sseNone Then .SelectorStyle = sseRoot
                        .Selectors = lngListMax + 1
                        ReDim .Selector(1 To .Selectors)
                        For i = 0 To lngListMax
                            ReDim .Selector(i + 1).Class(ceClasses - 1)
                            ReDim .Selector(i + 1).ClassLevel(ceClasses - 1)
                            .Selector(i + 1).SelectorName = strList(i)
                            If strList(i) = "All" Then .Selector(i + 1).All = True
                            .Selector(i + 1).ClassBonus = .ClassBonus
                            .Selector(i + 1).Race = .Race
                            .Selector(i + 1).Req = .Req
                            .Selector(i + 1).Skill = .Skill
                            .Selector(i + 1).SkillValue = .SkillValue
                            .Selector(i + 1).Stat = .Stat
                            .Selector(i + 1).StatValue = .StatValue
                        Next
                    Case "sharedselector"
                        .SelectorStyle = sseShared
                        .Parent.Raw = strItem
                    Case "selectorname"
                        log.Activity = actLoadSelector
                        strLine = Split(pstrRaw, "SelectorName: ")
                        For i = 1 To UBound(strLine)
                            LoadFeatSelector typNew, strLine(i)
                        Next
                        log.Activity = actReadFile
                        Exit For
                    Case Else
                        LogError
                End Select
                If blnError Then LogError
            End If
        Next
    End With
    SetFeatChannel typNew
    With db
        .Feats = .Feats + 1
        .Feat(.Feats) = typNew
        .Feat(.Feats).FeatIndex = .Feats
    End With
End Sub

Private Sub LoadFeatSelector(ptypFeat As FeatType, ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strSelector As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngSelector As Long
    Dim enRace As RaceEnum
    Dim enClass As ClassEnum
    Dim lngID As Long
    Dim blnError As Boolean
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    strSelector = Trim$(strLine(0))
    log.LoadSelector = strSelector
    If Len(strSelector) = 0 Then Exit Sub
    For lngSelector = 1 To ptypFeat.Selectors
        If ptypFeat.Selector(lngSelector).SelectorName = strSelector Then Exit For
    Next
    If lngSelector > ptypFeat.Selectors Then
        LogError
        Exit Sub
    End If
    With ptypFeat.Selector(lngSelector)
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                blnError = False
                Select Case strField
                    Case "wikiname"
                        .Wiki = strItem
                    Case "descrip"
                        .Descrip = strItem
                    Case "stat"
                        .Stat = GetStatID(strItem)
                        If .Stat = aeAny Then LogError Else .StatValue = lngValue
                    Case "skill"
                        .Skill = GetSkillID(strItem)
                        If .Skill = seAny Then LogError Else .SkillValue = lngValue
                    Case "race"
                        ReDim .Race(reRaces - 1)
                        .Race(0) = GetRaceReqID(strList(0))
                        If .Race(0) = rreAny Then
                            blnError = True
                        Else
                            For i = 1 To lngListMax
                                enRace = GetRaceID(strList(i))
                                If enRace = reAny Then blnError = True Else .Race(enRace) = 1
                            Next
                        End If
                    Case "classbonus"
                        ReDim .ClassBonus(ceClasses - 1)
                        For i = 0 To lngListMax
                            enClass = GetClassID(strList(i))
                            If enClass = ceAny Then blnError = True Else .ClassBonus(GetClassID(strList(i))) = True
                        Next
                    Case "class"
                        .Class(0) = True
                        For i = 0 To lngListMax
                            If ParseClassLevel(strList(i), enClass, lngValue) Then
                                .Class(enClass) = True
                                .ClassLevel(enClass) = lngValue
                            Else
                                blnError = True
                            End If
                        Next
                    Case "all", "one", "none"
                        With .Req(GetReqGroupID(strField))
                            .Reqs = lngListMax + 1
                            ReDim .Req(1 To lngListMax + 1)
                            For i = 0 To lngListMax
                                .Req(i + 1).Raw = "Feat: " & strList(i)
                                .Req(i + 1).Style = peFeat
                            Next
                        End With
                    Case "alignment"
                        .Alignment(0) = True
                        For i = 0 To lngListMax
                            lngID = GetAlignmentID(strList(i))
                            If lngID <> aleAny Then .Alignment(lngID) = True Else LogError
                        Next
                    Case "hide"
                        .Hide = True
                    Case "notclass"
                        .NotClass = GetClassID(strItem)
                        If .NotClass = ceAny Then LogError
                    Case Else
                        LogError
                End Select
                If blnError Then LogError
            End If
        Next
    End With
End Sub

Private Sub SetFeatChannel(ptypFeat As FeatType)
    With ptypFeat
        .Channel = fceGeneral
        If .Deity Then
            .Channel = fceDeity
        ElseIf .Domain Then
            .Channel = fceCleric
        ElseIf .Pact Then
            .Channel = fceWarlock
        ElseIf .RaceOnly Then
            .Channel = fceRacial
        ElseIf .ClassOnly Then
            If .ClassOnlyClasses(ceMonk) Then
                .Channel = fceMonk
            ElseIf .ClassOnlyClasses(ceRanger) Then
                .Channel = fceFavoredEnemy
            ElseIf .ClassOnlyClasses(ceRogue) Then
                .Channel = fceRogue
            ElseIf .ClassOnlyClasses(ceDruid) Then
                .Channel = fceWildShape
            ElseIf .ClassOnlyClasses(ceFavoredSoul) Then
                If .ClassOnlyLevels(5) Then .Channel = fceEnergy Else .Channel = fceFavoredSoul
            End If
        End If
    End With
End Sub

Private Sub IndexFeatLookup()
    Dim i As Long
    Dim j As Long
    Dim typSwap As FeatIndexType
    
    With db
        ' Create index
        ReDim .FeatLookup(1 To .Feats)
        For i = 1 To .Feats
            .FeatLookup(i).FeatIndex = i
            .FeatLookup(i).FeatName = .Feat(i).FeatName
        Next
        ' Sort index
        For i = 2 To db.Feats
            typSwap = db.FeatLookup(i)
            For j = i To 2 Step -1
                If typSwap.FeatName < db.FeatLookup(j - 1).FeatName Then db.FeatLookup(j) = db.FeatLookup(j - 1) Else Exit For
            Next j
            db.FeatLookup(j) = typSwap
        Next
    End With
End Sub

Public Sub IndexFeatDisplay()
    Dim i As Long
    Dim j As Long
    Dim typSwap As FeatIndexType
    
    With db
        ' Create index
        ReDim .FeatDisplay(1 To .Feats)
        For i = 1 To .Feats
            .FeatDisplay(i).FeatIndex = i
            If cfg.FeatOrder = foeAlphabetical Then .FeatDisplay(i).FeatName = .Feat(i).Abbreviation Else .FeatDisplay(i).FeatName = .Feat(i).SortName
        Next
        ' Sort index
        For i = 2 To db.Feats
            typSwap = db.FeatDisplay(i)
            For j = i To 2 Step -1
                If typSwap.FeatName < db.FeatDisplay(j - 1).FeatName Then db.FeatDisplay(j) = db.FeatDisplay(j - 1) Else Exit For
            Next j
            db.FeatDisplay(j) = typSwap
        Next
    End With
End Sub

' Simple binary search
Public Function SeekFeat(ByVal pstrFeatName As String, Optional pblnMatchCase As Boolean = True) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    If Len(pstrFeatName) = 0 Then Exit Function
    lngFirst = 1
    lngLast = db.Feats
    If pblnMatchCase Then
        Do While lngFirst <= lngLast
            lngMid = (lngFirst + lngLast) \ 2
            If db.FeatLookup(lngMid).FeatName < pstrFeatName Then
                lngFirst = lngMid + 1
            ElseIf db.FeatLookup(lngMid).FeatName > pstrFeatName Then
                lngLast = lngMid - 1
            Else
                SeekFeat = db.FeatLookup(lngMid).FeatIndex
                Exit Function
            End If
        Loop
    Else
        pstrFeatName = LCase$(pstrFeatName)
        Do While lngFirst <= lngLast
            lngMid = (lngFirst + lngLast) \ 2
            If LCase$(db.FeatLookup(lngMid).FeatName) < pstrFeatName Then
                lngFirst = lngMid + 1
            ElseIf LCase$(db.FeatLookup(lngMid).FeatName) > pstrFeatName Then
                lngLast = lngMid - 1
            Else
                SeekFeat = db.FeatLookup(lngMid).FeatIndex
                Exit Function
            End If
        Loop
    End If
End Function


' ************* FEAT MAP *************


Private Sub LoadFeatMap()
    Dim strFile As String
    Dim strRaw As String
    Dim strLine() As String
    Dim strToken() As String
    
    Erase db.FeatMap
    db.FeatMaps = 0
    db.FeatMapIndex = peUnknown
    If DataFile(strFile, "FeatMap.txt") Then Exit Sub
    log.Activity = actLoadFeatMap
    strRaw = xp.File.LoadToString(strFile)
    strLine = Split(strRaw, vbNewLine)
    db.FeatMaps = UBound(strLine)
    If db.FeatMaps < 1 Then
        LogError
        Exit Sub
    End If
    ReDim db.FeatMap(1 To db.FeatMaps)
    For log.LineNumber = 1 To db.FeatMaps
        log.LoadLine = strLine(log.LineNumber)
        strToken = Split(strLine(log.LineNumber), vbTab)
        If UBound(strToken) = 2 Then
            With db.FeatMap(log.LineNumber)
                .Lite = strToken(0)
                .Ron = strToken(1)
                .Builder = strToken(2)
            End With
        Else
            LogError
        End If
    Next
End Sub


' ************* ENHANCEMENTS & DESTINIES *************


' Enhancements and Destinies share the same data structure
Private Sub LoadEnhancements()
    Dim strFile As String
    Dim strRaw As String
    Dim strTree() As String
    Dim typNew As TreeType
    Dim typBlank As TreeType
    Dim i As Long
    
    Erase db.Tree
    db.Trees = 0
    If DataFile(strFile, "Enhancements.txt") Then Exit Sub
    log.Activity = actLoadTree
    ' Allocate enough space that we never have to increase
    ReDim db.Tree(127)
    strRaw = xp.File.LoadToString(strFile)
    strTree = Split(strRaw, "TreeName: ")
    For i = 1 To UBound(strTree)
        If InStr(strTree(i), "Type: ") Then
            typNew = typBlank
            If LoadTree(strTree(i), typNew) Then
                db.Trees = db.Trees + 1
                db.Tree(db.Trees) = typNew
            End If
        Else
            ErrorLoading strTree(i)
        End If
    Next
    With db
        If .Trees = 0 Then
            Erase .Tree
        Else
            ReDim Preserve .Tree(.Trees)
            SortTrees db.Tree, db.Trees
        End If
    End With
End Sub

Private Sub LoadDestinies()
    Dim strFile As String
    Dim strRaw As String
    Dim strDestiny() As String
    Dim typNew As TreeType
    Dim typBlank As TreeType
    Dim i As Long
    
    Erase db.Destiny
    db.Destinies = 0
    If DataFile(strFile, "Destinies.txt") Then Exit Sub
    log.Activity = actLoadTree
    ' Allocate enough space that we never have to increase
    ReDim db.Destiny(15)
    strRaw = xp.File.LoadToString(strFile)
    strDestiny = Split(strRaw, "DestinyName: ")
    For i = 1 To UBound(strDestiny)
        If InStr(strDestiny(i), "Stats: ") Then
            typNew = typBlank
            If LoadTree(strDestiny(i), typNew) Then
                db.Destinies = db.Destinies + 1
                db.Destiny(db.Destinies) = typNew
            End If
        Else
            ErrorLoading strDestiny(i)
        End If
    Next
    With db
        If .Destinies = 0 Then
            Erase .Destiny
        Else
            ReDim Preserve .Destiny(.Destinies)
            SortTrees db.Destiny, db.Destinies
        End If
    End With
End Sub

Private Function LoadTree(ByVal pstrRaw As String, ptypTree As TreeType) As Boolean
    Dim strAbility() As String
    Dim i As Long
    
    CleanText pstrRaw
    ' Split abilities now to isolate tree header
    strAbility = Split(pstrRaw, "AbilityName: ")
    ' Process header
    log.Tier = -1
    If LoadTreeHeader(strAbility(0), ptypTree) Then
        LoadError ptypTree.TreeName & " failed to load"
        Exit Function
    End If
    log.Tier = 0
     ' Process abilities
    For i = 1 To UBound(strAbility)
        LoadAbility strAbility(i), ptypTree
    Next
    ' Stats
    AddStats ptypTree
    ' Core prereqs
    AddCoreReqs ptypTree
    ' All good
    LoadTree = True
End Function

Private Function LoadTreeHeader(pstrRaw As String, ptypTree As TreeType) As Boolean
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngPos As Long
    Dim enStat As StatEnum
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With ptypTree
        .TreeName = strLine(0)
        .Abbreviation = .TreeName
        .Wiki = .TreeName
        ReDim .Initial(1)
        ReDim .Stats(6)
        log.LoadTree = .TreeName
        For lngLine = 1 To UBound(strLine)
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "abbreviation"
                        .Abbreviation = strItem
                    Case "wikiname"
                        .Wiki = strItem
                    Case "stats"
                        For i = 0 To lngListMax
                            enStat = GetStatID(strList(i))
                            If enStat = aeAny Then
                                LoadError .TreeName & " header has invalid stat: " & strList(i) & vbNewLine & strLine(lngLine)
                            Else
                                .Stats(enStat) = True
                                .Stats(0) = True
                            End If
                        Next
                    Case "lockout"
                        .Lockout = strItem
                    Case "initial"
                        Select Case lngListMax
                            Case 0: .Initial(0) = strList(0): .Initial(1) = strList(0)
                            Case 1: .Initial = strList
                            Case Else: LogError
                        End Select
                    Case "color"
                        Select Case LCase$(strItem)
                            Case "red": .Color = cveRed
                            Case "green": .Color = cveGreen
                            Case "blue": .Color = cveBlue
                            Case "yellow": .Color = cveYellow
                            Case "purple": .Color = cvePurple
                            Case "orange": .Color = cveOrange
                            Case Else: LogError
                        End Select
                    Case "type"
                        Select Case LCase$(strItem)
                            Case "race": InitTree ptypTree, 4, tseRace
                            Case "class": InitTree ptypTree, 5, tseClass
                            Case "raceclass": InitTree ptypTree, 5, tseRaceClass
                            Case "global": InitTree ptypTree, 5, tseGlobal
                            Case "destiny": InitTree ptypTree, 6, tseDestiny
                            Case Else
                                LoadError .TreeName & " header has invalid Type: " & strItem
                                LoadTreeHeader = True
                                Exit For
                        End Select
                    Case Else
                        LogError
                        Exit For
                End Select
            End If
        Next
    End With
End Function

Private Sub InitTree(ptypTree As TreeType, plngTiers As Long, penTreeStyle As TreeStyleEnum)
    Dim i As Long
    
    With ptypTree
        .TreeType = penTreeStyle
        .Tiers = plngTiers
        If penTreeStyle = tseDestiny Then .Wiki = .TreeName Else .Wiki = .Wiki & " enhancements"
        ReDim .Tier(.Tiers)
    End With
End Sub

Private Sub LoadAbility(ByVal pstrRaw As String, ptypTree As TreeType)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strAbility As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngPos As Long
    Dim typNew As AbilityType
    Dim lngTier As Long
    Dim lngReq As Long
    Dim lngGroupID As Long
    Dim lngRank As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    strAbility = Trim$(strLine(0))
    If Len(strAbility) = 0 Then Exit Sub
    With typNew
        .AbilityName = strAbility
        .Abbreviation = strAbility
        .Ranks = 1
        .Cost = 1
        log.LoadItem = .AbilityName
'        ReDim .Class(ceClasses - 1)
'        ReDim .ClassLevel(ceClasses - 1)
'        ReDim .Group(feFilters - 1)
'        .Group(feAll) = True
        ReDim .Req(3)
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "abbreviation"
                        .Abbreviation = strItem
                    Case "descrip"
                        .Descrip = strItem
                    Case "tier"
                        lngTier = lngValue
                        If Not ValidTier(ptypTree.TreeType, lngTier) Then lngTier = log.Tier
                        If log.Tier > lngTier Then LoadError log.LoadTree & " Tier " & lngTier & "(" & log.Tier & "?): " & strAbility & " not in Tier order"
                        log.Tier = lngTier
                    Case "ranks"
                        .Ranks = lngValue
                        If .Ranks < 1 Or .Ranks > 3 Then LogError
                    Case "cost"
                        .Cost = lngValue
                        If .Cost < 1 Or .Cost > 4 Then LogError
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "selectoronly": .SelectorOnly = True
                                Case Else: LogError
                            End Select
                        Next
                    Case "selector"
                        If .SelectorStyle = sseNone Then .SelectorStyle = sseRoot
                        .Selectors = lngListMax + 1
                        ReDim .Selector(1 To .Selectors)
                        For i = 0 To lngListMax
                            .Selector(i + 1).SelectorName = strList(i)
                            .Selector(i + 1).Cost = .Cost
                            .Selector(i + 1).Req = .Req
                        Next
                    Case "sharedselector"
                        .SelectorStyle = sseShared
                        .Parent.Raw = strItem
                    Case "parent"
                        .SelectorStyle = sseExclusive
                        .Parent.Raw = strItem
                    Case "siblings"
                        .Siblings = lngListMax + 1
                        If .Siblings Then
                            ReDim .Sibling(1 To lngListMax + 1)
                            For i = 0 To lngListMax
                                .Sibling(i + 1).Raw = strList(i)
                            Next
                        End If
                    Case "all", "one", "none"
                        With .Req(GetReqGroupID(strField))
                            .Reqs = lngListMax + 1
                            ReDim .Req(1 To .Reqs)
                            For i = 0 To lngListMax
                                .Req(i + 1).Raw = strList(i)
                                If Left$(strList(i), 5) = "Feat:" Then
                                    .Req(i + 1).Style = peFeat
                                ElseIf ptypTree.TreeType = tseDestiny Then
                                    .Req(i + 1).Style = peDestiny
                                Else
                                    .Req(i + 1).Style = peEnhancement
                                End If
                            Next
                        End With
                    Case "rank2all", "rank3all", "rank3none"
                        .RankReqs = True
                        InitRanks .Rank
                        lngRank = Val(Mid$(strField, 5, 1))
                        With .Rank(lngRank).Req(GetReqGroupID(Mid$(strField, 6)))
                            .Reqs = lngListMax + 1
                            ReDim .Req(1 To .Reqs)
                            For i = 0 To lngListMax
                                .Req(i + 1).Raw = strList(i)
                                If Left$(strList(i), 5) = "Feat:" Then
                                    .Req(i + 1).Style = peFeat
                                ElseIf ptypTree.TreeType = tseDestiny Then
                                    .Req(i + 1).Style = peDestiny
                                Else
                                    .Req(i + 1).Style = peEnhancement
                                End If
                            Next
                        End With
'                    Case "class"
'                        .Class(0) = True
'                        For i = 0 To lngListMax
'                            lngPos = InStrRev(strList(i), " ")
'                            If lngPos = 0 Then
'                                LogError
'                            Else
'                                strItem = Left$(strList(i), lngPos - 1)
'                                lngValue = GetClassID(strItem)
'                                If lngValue = 0 Then
'                                    LogError
'                                Else
'                                    .Class(lngValue) = True
'                                    .ClassLevel(lngValue) = Val(Mid$(strList(i), lngPos + 1))
'                                End If
'                            End If
'                        Next
'                    Case "rank2class"
'                        .RankReqs = True
'                        InitRanks .Rank
'                        With .Rank(2)
'                            .Class(0) = True
'                            For i = 0 To lngListMax
'                                lngPos = InStrRev(strList(i), " ")
'                                If lngPos = 0 Then
'                                    LogError
'                                Else
'                                    strItem = Left$(strList(i), lngPos - 1)
'                                    lngValue = GetClassID(strItem)
'                                    If lngValue = 0 Then
'                                        LogError
'                                    Else
'                                        .Class(lngValue) = True
'                                        .ClassLevel(lngValue) = Val(Mid$(strList(i), lngPos + 1))
'                                    End If
'                                End If
'                            Next
'                        End With
                    Case "selectorname"
                        strLine = Split(pstrRaw, "SelectorName: ")
                        For i = 1 To UBound(strLine)
                            LoadSelector typNew, strLine(i), ptypTree.TreeType
                        Next
                        Exit For
                    Case Else
                        LogError
'                        Exit For
                End Select
            End If
        Next
    End With
    With ptypTree.Tier(lngTier)
        .Abilities = .Abilities + 1
        ReDim Preserve .Ability(1 To .Abilities)
        .Ability(.Abilities) = typNew
    End With
End Sub

Private Function ValidTier(penType As TreeStyleEnum, plngTier As Long) As Boolean
    Dim lngMin As Long
    Dim lngMax As Long
    
    Select Case penType
        Case tseClass, tseRaceClass, tseGlobal
            lngMax = 5
        Case tseRace
            lngMax = 4
        Case tseDestiny
            lngMin = 1
            lngMax = 6
    End Select
    Select Case plngTier
        Case lngMin To lngMax: ValidTier = True
        Case Else: LogError
    End Select
End Function

Private Sub LoadSelector(ptypAbility As AbilityType, ByVal pstrRaw As String, penTreeStyle As TreeStyleEnum)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strSelector As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngSelector As Long
    Dim lngRank As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    strSelector = Trim$(strLine(0))
    If Len(strSelector) = 0 Then Exit Sub
    For lngSelector = 1 To ptypAbility.Selectors
        If ptypAbility.Selector(lngSelector).SelectorName = strSelector Then Exit For
    Next
    If lngSelector > ptypAbility.Selectors Then Exit Sub
    With ptypAbility.Selector(lngSelector)
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            log.LoadLine = strLine(lngLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "descrip"
                        ' 20200908
                        .Descrip = strItem
                    Case "cost"
                        .Cost = lngValue
                        If .Cost < 1 Or .Cost > 2 Then LogError
                    Case "all", "one", "none"
                        With .Req(GetReqGroupID(strField))
                            .Reqs = lngListMax + 1
                            ReDim .Req(1 To .Reqs)
                            For i = 0 To lngListMax
                                .Req(i + 1).Raw = strList(i)
                                If Left$(strList(i), 5) = "Feat:" Then
                                    .Req(i + 1).Style = peFeat
                                ElseIf penTreeStyle = tseDestiny Then
                                    .Req(i + 1).Style = peDestiny
                                Else
                                    .Req(i + 1).Style = peEnhancement
                                End If
                            Next
                        End With
                    Case "rank2all", "rank3all", "rank3none"
                        .RankReqs = True
                        InitRanks .Rank
                        lngRank = Val(Mid$(strField, 5, 1))
                        With .Rank(lngRank).Req(GetReqGroupID(Mid$(strField, 6)))
                            .Reqs = lngListMax + 1
                            ReDim .Req(1 To .Reqs)
                            For i = 0 To lngListMax
                                .Req(i + 1).Raw = strList(i)
                                If Left$(strList(i), 5) = "Feat:" Then
                                    .Req(i + 1).Style = peFeat
                                ElseIf penTreeStyle = tseDestiny Then
                                    .Req(i + 1).Style = peDestiny
                                Else
                                    .Req(i + 1).Style = peEnhancement
                                End If
                            Next
                        End With
                    Case Else
                        LogError
                        Exit For
                End Select
            End If
        Next
    End With
End Sub

Private Sub InitRanks(ptypRank() As RankType)
    Dim i As Long
    
    ReDim Preserve ptypRank(2 To 3)
    For i = 2 To 3
        With ptypRank(i)
            ReDim Preserve .Class(ceClasses - 1)
            ReDim Preserve .ClassLevel(ceClasses - 1)
            ReDim Preserve .Req(3)
        End With
    Next
End Sub

Private Sub AddStats(ptypTree As TreeType)
    Dim strStat() As String
    Dim lngStats As Long
    Dim lngSingle As Long ' Single stat ID
    Dim lngSelector As Long
    Dim lngTier As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strAbilityName As String
    Dim i As Long
    
    Select Case ptypTree.TreeType
        Case tseRace
            Exit Sub
        Case tseClass, tseRaceClass, tseGlobal
            lngStart = 3
            lngEnd = 4
        Case tseDestiny
            lngStart = 1
            lngEnd = 6
        Case Else
            LogError
            Exit Sub
    End Select
    ' Count stats (if only one stat, we don't need a selector)
    For i = 1 To 6
        If ptypTree.Stats(i) Then
            lngStats = lngStats + 1
            lngSingle = i
        End If
    Next
    Select Case lngStats
        Case 0: Exit Sub
        Case 1: strAbilityName = GetStatName(lngSingle)
        Case Else: strAbilityName = "Stat"
    End Select
    ' Add stats to end of tiers
    For lngTier = lngStart To lngEnd
        With ptypTree.Tier(lngTier)
            .Abilities = .Abilities + 1
            ReDim Preserve .Ability(1 To .Abilities)
            With .Ability(.Abilities)
                .AbilityName = strAbilityName
                .Abbreviation = .AbilityName
                .Descrip = "+1 to the selected ability score"
                .Cost = 2
                .Ranks = 1
                ' Add prereq to previous tier
                ReDim .Req(3)
                If lngTier > lngStart Then
                    With .Req(rgeAll)
                        .Reqs = 1
                        ReDim .Req(1 To .Reqs)
                        .Req(1).Raw = "Tier " & lngTier - 1 & ": " & strAbilityName
                    End With
                End If
                ' Add selectors
                If lngStats = 1 Then
                    .SelectorStyle = sseNone
                Else
                    .SelectorStyle = sseRoot
                    .SelectorOnly = True
                    .Selectors = lngStats
                    ReDim .Selector(1 To lngStats)
                    lngSelector = 1
                    For i = 1 To 6
                        If ptypTree.Stats(i) Then
                            .Selector(lngSelector).SelectorName = GetStatName(i)
                            .Selector(lngSelector).Cost = 2
                            .Selector(lngSelector).Req = .Req
                            lngSelector = lngSelector + 1
                        End If
                    Next
                End If
            End With
        End With
    Next
End Sub

Private Sub AddCoreReqs(ptypTree As TreeType)
    Dim i As Long
    
    With ptypTree.Tier(0)
        For i = 2 To .Abilities
            With .Ability(i).Req(rgeAll)
                .Reqs = .Reqs + 1
                ReDim Preserve .Req(1 To .Reqs)
                With .Req(.Reqs)
                    .Raw = "Tier 0: " & ptypTree.Tier(0).Ability(i - 1).AbilityName
                End With
            End With
        Next
    End With
End Sub

' Insertion sort
Private Sub SortTrees(ptypTree() As TreeType, plngTrees As Long)
    Dim i As Long
    Dim j As Long
    Dim typSwap As TreeType
    
    For i = 2 To plngTrees
        typSwap = ptypTree(i)
        For j = i To 2 Step -1
            If typSwap.TreeName < ptypTree(j - 1).TreeName Then ptypTree(j) = ptypTree(j - 1) Else Exit For
        Next j
        ptypTree(j) = typSwap
    Next i
    For i = 1 To plngTrees
        ptypTree(i).TreeID = i
    Next
End Sub

' Simple binary search
Public Function SeekTree(pstrTreeName As String, penTreeStyle As PointerEnum) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    If penTreeStyle = peDestiny Then
        lngFirst = 1
        lngLast = db.Destinies
        Do While lngFirst <= lngLast
            lngMid = (lngFirst + lngLast) \ 2
            If db.Destiny(lngMid).TreeName < pstrTreeName Then
                lngFirst = lngMid + 1
            ElseIf db.Destiny(lngMid).TreeName > pstrTreeName Then
                lngLast = lngMid - 1
            Else
                SeekTree = lngMid
                Exit Function
            End If
        Loop
    Else
        lngFirst = 1
        lngLast = db.Trees
        Do While lngFirst <= lngLast
            lngMid = (lngFirst + lngLast) \ 2
            If db.Tree(lngMid).TreeName < pstrTreeName Then
                lngFirst = lngMid + 1
            ElseIf db.Tree(lngMid).TreeName > pstrTreeName Then
                lngLast = lngMid - 1
            Else
                SeekTree = lngMid
                Exit Function
            End If
        Loop
    End If
End Function


' ************* NAME CHANGES *************


Private Sub LoadNameChanges()
    Dim strFile As String
    Dim strRaw As String
    Dim strNameChange() As String
    Dim i As Long
    
    Erase db.NameChange
    db.NameChanges = 0
    If DataFile(strFile, "NameChange.txt") Then Exit Sub
    ' Allocate enough space that we never have to increase
    ReDim db.NameChange(127)
    strRaw = xp.File.LoadToString(strFile)
    strNameChange = Split(strRaw, "NameChangeType: ")
    For i = 1 To UBound(strNameChange)
        LoadNameChange strNameChange(i)
    Next
    With db
        If .NameChanges = 0 Then Erase .NameChange Else ReDim Preserve .NameChange(.NameChanges)
    End With
    SortNameChanges
End Sub

Private Sub LoadNameChange(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim lngID As Long
    Dim typNew As NameChangeType
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .Type = strLine(0)
        log.LoadItem = .Type
        Select Case .Type
            Case "Feat"
                ' Process lines
                For lngLine = 1 To UBound(strLine)
                    log.HasError = False
                    log.LoadLine = strLine(lngLine)
                    If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                        Select Case strField
                            Case "old": .OldName = strItem
                            Case "new": .NewName = strItem
                            Case "assignselector": .AssignSelector = strItem
                            Case Else: LogError
                        End Select
                    End If
                Next
            Case Else
                LoadError "Invalid Type: " & strLine(0)
        End Select
    End With
    If Len(typNew.OldName) <> 0 And Len(typNew.NewName) <> 0 Then
        With db
            .NameChanges = .NameChanges + 1
            .NameChange(.NameChanges) = typNew
        End With
    End If
End Sub

Public Sub SortNameChanges()
    Dim i As Long
    Dim j As Long
    Dim typSwap As NameChangeType
    
    With db
        For i = 2 To db.NameChanges
            typSwap = db.NameChange(i)
            For j = i To 2 Step -1
                If CompareNameChange(typSwap, db.NameChange(j - 1)) = -1 Then db.NameChange(j) = db.NameChange(j - 1) Else Exit For
            Next j
            db.NameChange(j) = typSwap
        Next
    End With
End Sub

Private Function CompareNameChange(ptypLeft As NameChangeType, ptypRight As NameChangeType) As Long
    If ptypLeft.Type < ptypRight.Type Then
        CompareNameChange = -1
    ElseIf ptypLeft.Type > ptypRight.Type Then
        CompareNameChange = 1
    ElseIf ptypLeft.OldName < ptypRight.OldName Then
        CompareNameChange = -1
    ElseIf ptypLeft.OldName > ptypRight.OldName Then
        CompareNameChange = 1
    End If
End Function

' Simple binary search
Public Function SeekNameChange(pstrType As String, pstrOld As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    Dim typFind As NameChangeType
    
    typFind.Type = pstrType
    typFind.OldName = pstrOld
    lngFirst = 1
    lngLast = db.NameChanges
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        Select Case CompareNameChange(db.NameChange(lngMid), typFind)
            Case -1
                lngFirst = lngMid + 1
            Case 1
                lngLast = lngMid - 1
            Case Else
                SeekNameChange = lngMid
                Exit Function
        End Select
    Loop
End Function


' ************* GENERAL *************


Private Function ParseLine(ByVal pstrLine As String, pstrField As String, pstrItem As String, plngValue As Long, pstrList() As String, plngListMax As Long) As Boolean
    Dim lngPos As Long
    Dim strValue As String
    Dim i As Long
    
    ' Prep
    pstrField = vbNullString
    pstrItem = vbNullString
    plngValue = 0
    If plngListMax <> -1 Then Erase pstrList
    plngListMax = -1
    pstrLine = TrimAll(pstrLine)
    If Len(pstrLine) = 0 Then Exit Function
    ' Field
    lngPos = InStr(pstrLine, ":")
    If lngPos = 0 Then
        LogError
        Exit Function
    End If
    ParseLine = True
    pstrField = LCase$(TrimAll(Left$(pstrLine, lngPos - 1)))
    pstrLine = TrimAll(Mid$(pstrLine, lngPos + 1))
    ' Descriptions
    If Left$(pstrField, 4) = "wiki" Or pstrField = "descrip" Then
        pstrItem = pstrLine
        Exit Function
    End If
    ' List
    If InStr(pstrLine, ",") Then
        pstrList = Split(pstrLine, ",")
        plngListMax = UBound(pstrList)
        For i = 0 To plngListMax
            pstrList(i) = Trim$(pstrList(i))
        Next
        Exit Function
    End If
    ' Value
    pstrItem = pstrLine
    If InStr(pstrLine, " ") Then
        lngPos = InStrRev(pstrLine, " ")
        strValue = Mid$(pstrLine, lngPos + 1)
        If IsNumeric(strValue) Then
            pstrItem = Left$(pstrLine, lngPos - 1)
            plngValue = Val(strValue)
        End If
    Else
        ' If only a single value, and it's numeric, return it in Value also
        If IsNumeric(pstrItem) Then plngValue = Val(pstrItem)
    End If
    ' Return single item in list form as well
    plngListMax = 0
    ReDim pstrList(0)
    pstrList(0) = pstrLine
End Function

Private Function ParseClassLevel(pstrRaw As String, penClass As ClassEnum, plngLevel As Long) As Boolean
    Dim lngPos As Long
    
    lngPos = InStrRev(pstrRaw, " ")
    If lngPos = 0 Then Exit Function
    penClass = GetClassID(Left$(pstrRaw, lngPos - 1))
    If penClass = ceAny Then Exit Function
    plngLevel = Val(Mid$(pstrRaw, lngPos + 1))
    ParseClassLevel = (plngLevel >= 1 And plngLevel <= 20)
End Function
