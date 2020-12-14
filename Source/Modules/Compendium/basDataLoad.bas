Attribute VB_Name = "basDataLoad"
' Written by Ellis Dee
Option Explicit

Public Sub LoadDataFiles()
    Dim typBlank As DatabaseType
    
    db = typBlank
    DeleteOldFiles
    ' Static data
    LoadPatrons
    LoadAreas
    LoadPacks
    LoadQuests
    LoadSagas
    LoadChallenges
    LoadLinkLists
    LoadTables
    LoadTemplates
    ' Character data
    LoadCompendium
    ' If this is the first time running the program, match colors
    If Not cfg.RunBefore Then MatchColors True
End Sub

Private Sub DeleteOldFiles()
    DeleteOldFile "SagaQuests.txt"
End Sub

Private Sub DeleteOldFile(pstrFile As String)
    Dim strFile As String
    
    strFile = DataPath() & pstrFile
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
End Sub


' ************* PATRONS *************


Private Sub LoadPatrons()
    Dim strFile As String
    Dim strPatron() As String
    Dim lngPatrons As Long
    Dim i As Long
    
    strFile = DataPath() & "Patrons.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strPatron = Split(xp.File.LoadToString(strFile), "PatronName: ")
    lngPatrons = UBound(strPatron)
    If lngPatrons = 0 Then Exit Sub
    ReDim db.Patron(1 To lngPatrons + 1)
    For i = 1 To UBound(strPatron)
        LoadPatron strPatron(i)
    Next
    If UBound(db.Patron) <> db.Patrons Then ReDim Preserve db.Patron(1 To db.Patrons)
    SortPatrons
End Sub

Private Sub LoadPatron(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As PatronType
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Patron = Trim$(strLine(0))
    If Len(typNew.Patron) = 0 Then Exit Sub
    db.Patrons = db.Patrons + 1
    With typNew
        .Abbreviation = .Patron
        .Wiki = .Patron
        .Order = db.Patrons
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "abbreviation": .Abbreviation = strItem
                    Case "wiki": .Wiki = strItem
                End Select
            End If
        Next
    End With
    db.Patron(db.Patrons) = typNew
End Sub

' Insertion sort
Private Sub SortPatrons()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As PatronType
    
    iMin = 2
    iMax = db.Patrons
    For i = iMin To iMax
        typSwap = db.Patron(i)
        For j = i To iMin Step -1
            If typSwap.Patron < db.Patron(j - 1).Patron Then db.Patron(j) = db.Patron(j - 1) Else Exit For
        Next j
        db.Patron(j) = typSwap
    Next i
End Sub

Public Function SeekPatron(pstrPatron As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Patrons
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Patron(lngMid).Patron > pstrPatron Then
            lngLast = lngMid - 1
        ElseIf db.Patron(lngMid).Patron < pstrPatron Then
            lngFirst = lngMid + 1
        Else
            SeekPatron = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* AREAS *************


Private Sub LoadAreas()
    Dim strFile As String
    Dim strArea() As String
    Dim lngAreas As Long
    Dim i As Long
    
    strFile = DataPath() & "Wilderness.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strArea = Split(xp.File.LoadToString(strFile), "Area: ")
    lngAreas = UBound(strArea)
    If lngAreas = 0 Then Exit Sub
    ReDim db.Area(1 To lngAreas + 1)
    For i = 1 To UBound(strArea)
        LoadArea strArea(i)
    Next
    If UBound(db.Area) <> db.Areas Then ReDim Preserve db.Area(1 To db.Areas)
    SortAreas
End Sub

Private Sub LoadArea(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As AreaType
    Dim lngPack As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Area = Trim$(strLine(0))
    If Len(typNew.Area) = 0 Then Exit Sub
    With typNew
        .Wiki = .Area
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "wiki": .Wiki = strItem
                    Case "map": .Map = strItem
                    Case "low": .Lowest = lngValue
                    Case "high": .Highest = lngValue
                    Case "explorer": .Explorer = lngValue
                    Case "pack": .Pack = strItem
                    Case "link"
                        If lngListMax = 2 Then
                            .Links = .Links + 1
                            ReDim Preserve .Link(1 To .Links)
                            With .Link(.Links)
                                .FullName = strList(0)
                                .Abbreviation = strList(1)
                                .Target = strList(2)
                                .Style = mceLink
                            End With
                        End If
                End Select
            End If
        Next
    End With
    db.Areas = db.Areas + 1
    typNew.Order = db.Areas
    db.Area(db.Areas) = typNew
End Sub

' Insertion sort
Private Sub SortAreas()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As AreaType
    
    iMin = 2
    iMax = db.Areas
    For i = iMin To iMax
        typSwap = db.Area(i)
        For j = i To iMin Step -1
            If typSwap.Area < db.Area(j - 1).Area Then db.Area(j) = db.Area(j - 1) Else Exit For
        Next j
        db.Area(j) = typSwap
    Next i
End Sub

Public Function SeekArea(pstrArea As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Areas
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Area(lngMid).Area > pstrArea Then
            lngLast = lngMid - 1
        ElseIf db.Area(lngMid).Area < pstrArea Then
            lngFirst = lngMid + 1
        Else
            SeekArea = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* PACKS *************


Private Sub LoadPacks()
    Dim strFile As String
    Dim strPack() As String
    Dim lngPacks As Long
    Dim i As Long
    
    strFile = DataPath() & "Packs.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strPack = Split(xp.File.LoadToString(strFile), "PackName: ")
    lngPacks = UBound(strPack)
    If lngPacks = 0 Then Exit Sub
    ReDim db.Pack(1 To lngPacks + 1)
    For i = 1 To UBound(strPack)
        LoadPack strPack(i)
    Next
    If UBound(db.Pack) <> db.Packs Then ReDim Preserve db.Pack(1 To db.Packs)
    SortPacks
End Sub

Private Sub LoadPack(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As PackType
    Dim lngArea As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Pack = Trim$(strLine(0))
    If Len(typNew.Pack) = 0 Then Exit Sub
    With typNew
        .Abbreviation = .Pack
        .Wiki = .Pack
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "abbreviation"
                        .Abbreviation = strItem
                    Case "wiki"
                        .Wiki = strItem
                    Case "area"
                        lngArea = SeekArea(strItem)
                        If lngArea = 0 Then
                            Debug.Print "Error in Packs.txt: Area not found (" & strItem & ")"
                        Else
                            .Links = .Links + 1
                            ReDim Preserve .Link(1 To .Links)
                            With .Link(.Links)
                                .FullName = strItem
                                .Abbreviation = strItem
                                .Style = mceLink
                                .Target = MakeWiki(db.Area(lngArea).Wiki)
                            End With
                        End If
                    Case "link"
                        If lngListMax = 2 Then
                            .Links = .Links + 1
                            ReDim Preserve .Link(1 To .Links)
                            With .Link(.Links)
                                .FullName = strList(0)
                                .Abbreviation = .FullName
                                Select Case strList(1)
                                    Case "Link": .Style = mceLink
                                    Case "Shortcut": .Style = mceShortcut
                                End Select
                                .Target = strList(2)
                            End With
                        End If
                End Select
            End If
        Next
    End With
    db.Packs = db.Packs + 1
    typNew.Order = db.Packs
    db.Pack(db.Packs) = typNew
End Sub

' Insertion sort
Private Sub SortPacks()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As PackType
    
    iMin = 2
    iMax = db.Packs
    For i = iMin To iMax
        typSwap = db.Pack(i)
        For j = i To iMin Step -1
            If typSwap.Pack < db.Pack(j - 1).Pack Then db.Pack(j) = db.Pack(j - 1) Else Exit For
        Next j
        db.Pack(j) = typSwap
    Next i
End Sub

Public Function SeekPack(pstrPack As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Packs
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Pack(lngMid).Pack > pstrPack Then
            lngLast = lngMid - 1
        ElseIf db.Pack(lngMid).Pack < pstrPack Then
            lngFirst = lngMid + 1
        Else
            SeekPack = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* QUESTS *************


Private Sub LoadQuests()
    Dim strFile As String
    Dim strQuest() As String
    Dim lngQuests As Long
    Dim i As Long
    
    strFile = DataPath() & "Quests.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strQuest = Split(xp.File.LoadToString(strFile), "QuestName: ")
    lngQuests = UBound(strQuest)
    If lngQuests = 0 Then Exit Sub
    ReDim db.Quest(1 To lngQuests + 1)
    For i = 1 To UBound(strQuest)
        LoadQuest strQuest(i)
    Next
    If UBound(db.Quest) <> db.Quests Then ReDim Preserve db.Quest(1 To db.Quests)
    SortQuests
End Sub

Private Sub LoadQuest(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As QuestType
    Dim lngArea As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Quest = Trim$(strLine(0))
    If Len(typNew.Quest) = 0 Then Exit Sub
    With typNew
        .Style = qeQuest
        .Wiki = .Quest
        .ID = .Quest
        .SortName = MakeSortName(.Quest)
        .CompendiumName = .SortName
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "id": .ID = strItem
                    Case "wiki": .Wiki = strItem
                    Case "pack": .Pack = strItem
                    Case "patron": .Patron = strItem
                    Case "favor": .Favor = lngValue
                    Case "level"
                        .BaseLevel = lngValue
                        .GroupLevel = lngValue
                    Case "epic": .EpicLevel = lngValue
                    Case "style": .Style = GetQuestStyleID(strItem)
                    Case "grouplevel": .GroupLevel = lngValue
                    Case "sortname": .SortName = strItem
                    Case "compendium": .CompendiumName = strItem
                    Case "map"
                        lngArea = SeekArea(strItem)
                        If lngArea Then
                            .Links = .Links + 1
                            ReDim Preserve .Link(1 To .Links)
                            With .Link(.Links)
                                .FullName = "Map: " & db.Area(lngArea).Area
                                .Abbreviation = .FullName
                                .Style = mceLink
                                .Target = WikiImage(db.Area(lngArea).Map)
                            End With
                        Else
                            Debug.Print "Invalid line in Quests.txt" & pstrRaw
                        End If
                    Case "link"
                        If lngListMax = 2 Then
                            .Links = .Links + 1
                            ReDim Preserve .Link(1 To .Links)
                            With .Link(.Links)
                                .FullName = strList(0)
                                .Abbreviation = .FullName
                                Select Case strList(1)
                                    Case "Link": .Style = mceLink
                                    Case "Shortcut": .Style = mceShortcut
                                End Select
                                .Target = strList(2)
                            End With
                        End If
                End Select
            End If
        Next
    End With
    db.Quests = db.Quests + 1
    typNew.Order = db.Quests
    db.Quest(db.Quests) = typNew
End Sub

Private Function MakeSortName(pstrName As String) As String
    If Left$(pstrName, 4) = "The " Then MakeSortName = Mid$(pstrName, 5) Else MakeSortName = pstrName
End Function

' Omit plngLeft & plngRight; they are used internally during recursion
Private Sub SortQuests(Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim strMid As String
    Dim typSwap As QuestType
    
    If plngRight = 0 Then
        plngLeft = 1
        plngRight = db.Quests
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    strMid = db.Quest((plngLeft + plngRight) \ 2).ID
    Do
        Do While db.Quest(lngFirst).ID < strMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While strMid < db.Quest(lngLast).ID And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            typSwap = db.Quest(lngFirst)
            db.Quest(lngFirst) = db.Quest(lngLast)
            db.Quest(lngLast) = typSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then SortQuests plngLeft, lngLast
    If lngFirst < plngRight Then SortQuests lngFirst, plngRight
End Sub

Public Function SeekQuest(pstrID As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long

    lngFirst = 1
    lngLast = db.Quests
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Quest(lngMid).ID > pstrID Then
            lngLast = lngMid - 1
        ElseIf db.Quest(lngMid).ID < pstrID Then
            lngFirst = lngMid + 1
        Else
            SeekQuest = lngMid
            Exit Do
        End If
    Loop
End Function


' ************* SAGAS *************


Private Sub LoadSagas()
    Dim strFile As String
    Dim strSaga() As String
    Dim lngSagas As Long
    Dim lngHeroic As Long
    Dim lngEpic As Long
    Dim i As Long
    
    Erase db.Saga
    SagaOrder
    strFile = DataPath() & "Sagas.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strSaga = Split(xp.File.LoadToString(strFile), "SagaName: ")
    lngSagas = UBound(strSaga)
    If lngSagas = 0 Then Exit Sub
    ReDim Preserve db.Saga(1 To lngSagas + 1)
    For i = 1 To UBound(strSaga)
        LoadSaga strSaga(i), lngHeroic, lngEpic
    Next
    If UBound(db.Saga) <> db.Sagas Then ReDim Preserve db.Saga(1 To db.Sagas)
    SortSagas
    ConnectSagaQuests steHeroic
    ConnectSagaQuests steEpic
End Sub

Private Sub LoadSaga(ByVal pstrRaw As String, plngHeroic As Long, plngEpic As Long)
    Const Buffer As Long = 32
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As SagaType
    Dim lngQuest As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.SagaName = Trim$(strLine(0))
    If Len(typNew.SagaName) = 0 Then Exit Sub
    With typNew
        ReDim .Quest(1 To Buffer)
        .Wiki = .SagaName
        .Reward(0).Renown = 5000
        .Reward(1).Renown = 7500
        .Reward(2).Renown = 10000
        .Reward(3).Renown = 15000
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "abbreviation": .Abbreviation = strItem
                    Case "tier"
                        If LCase$(strItem) = "epic" Then
                            .Tier = steEpic
                            .Order = IncrementOrder(plngEpic)
                        Else
                            .Tier = steHeroic
                            .Order = IncrementOrder(plngHeroic)
                        End If
                    Case "npc"
                        .NPCs = .NPCs + 1
                        ReDim Preserve .NPC(1 To .NPCs)
                        .NPC(.NPCs) = strItem
                    Case "tome": .Tome = Val(strItem)
                    Case "astrals": .Astrals = lngValue
                    Case "xp"
                        If lngListMax = 3 Then
                            For i = 0 To 3
                                .Reward(i).xp = Val(strList(i)) * 1000
                            Next
                        End If
                    Case "quest"
                        lngQuest = SeekQuest(strItem)
                        If lngQuest Then
                            .Quests = .Quests + 1
                            .Quest(.Quests) = lngQuest
                        Else
                            Debug.Print "Quest not found in saga " & .SagaName & ": " & strItem
                        End If
                End Select
            End If
        Next
        If .Quests <> Buffer Then ReDim Preserve .Quest(1 To .Quests)
        .Reward(0).Points = .Quests
        .Reward(1).Points = Int(.Quests * 1.75)
        .Reward(2).Points = Int(.Quests * 2.5)
        .Reward(3).Points = .Quests * 3
    End With
    db.Sagas = db.Sagas + 1
    db.Saga(db.Sagas) = typNew
End Sub

Private Function IncrementOrder(plngOrder As Long) As Long
    plngOrder = plngOrder + 1
    IncrementOrder = plngOrder
End Function

Private Sub SagaOrder()
    SagaQuests "SagaOrderHeroic.txt", steHeroic
    SagaQuests "SagaOrderEpic.txt", steEpic
End Sub

Private Sub SagaQuests(pstrFile As String, penTier As SagaTierEnum)
    Dim strFile As String
    Dim strSystem As String
    Dim strUser As String
    
    strSystem = DataPath() & pstrFile
    strUser = DataPath() & "User" & pstrFile
    If xp.File.Exists(strUser) Then
        If xp.File.Exists(strSystem) Then
            If FileDateTime(strSystem) > FileDateTime(strUser) Then
                strFile = strSystem
                cfg.MessageAdd pstrFile & " is newer than User" & pstrFile & " so it was used instead."
            Else
                strFile = strUser
            End If
        Else
            strFile = strUser
        End If
    Else
        strFile = strSystem
    End If
    OrderSagaQuests strFile, penTier
End Sub

Private Sub OrderSagaQuests(pstrFile As String, penTier As SagaTierEnum)
    Dim strRaw As String
    Dim strLine() As String
    Dim lngQuest As Long
    Dim lngSagaGroup As Long
    Dim lngSagaQuest As Long
    Dim i As Long
    
    If Not xp.File.Exists(pstrFile) Then Exit Sub
    strRaw = xp.File.LoadToString(pstrFile)
    strLine = Split(strRaw, vbNewLine)
    If UBound(strLine) < 1 Then Exit Sub
    lngSagaGroup = 1
    With db
        For i = 0 To UBound(strLine)
            lngQuest = SeekQuest(strLine(i))
            If lngQuest = 0 Then
                lngSagaGroup = lngSagaGroup + 1
            Else
                lngSagaQuest = lngSagaQuest + 1
                With db.Quest(lngQuest)
                    .SagaGroup(penTier) = lngSagaGroup
                    .SagaOrder(penTier) = lngSagaQuest
                    ReDim .Saga(db.Sagas)
                End With
            End If
        Next
    End With
End Sub

' Insertion sort
Private Sub SortSagas()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As SagaType
    
    iMin = 2
    iMax = db.Sagas
    For i = iMin To iMax
        typSwap = db.Saga(i)
        For j = i To iMin Step -1
            If typSwap.SagaName < db.Saga(j - 1).SagaName Then db.Saga(j) = db.Saga(j - 1) Else Exit For
        Next j
        db.Saga(j) = typSwap
    Next i
End Sub

Public Function SeekSaga(pstrSaga As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Sagas
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Saga(lngMid).SagaName > pstrSaga Then
            lngLast = lngMid - 1
        ElseIf db.Saga(lngMid).SagaName < pstrSaga Then
            lngFirst = lngMid + 1
        Else
            SeekSaga = lngMid
            Exit Function
        End If
    Loop
End Function

Private Sub ConnectSagaQuests(penTier As SagaTierEnum)
    Dim lngQuest As Long
    Dim s As Long
    Dim q As Long
    
    For q = 1 To db.Quests
        With db.Quest(q)
            If .SagaGroup(penTier) Then
                ReDim Preserve .Saga(db.Sagas)
            End If
        End With
    Next
    For s = 1 To db.Sagas
        If db.Saga(s).Tier = penTier Then
            For q = 1 To db.Saga(s).Quests
                lngQuest = db.Saga(s).Quest(q)
                With db.Quest(lngQuest)
                    .Saga(s) = q
                End With
            Next
        End If
    Next
End Sub


' ************* TABLES *************


Private Sub LoadTables()
    Dim strFile As String
    Dim strTable() As String
    Dim lngTables As Long
    Dim i As Long
    
    strFile = DataPath() & "Tables.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strTable = Split(xp.File.LoadToString(strFile), "Table: ")
    lngTables = UBound(strTable)
    If lngTables = 0 Then Exit Sub
    ReDim db.Table(1 To lngTables + 1)
    For i = 1 To UBound(strTable)
        LoadTable strTable(i)
    Next
    If UBound(db.Table) <> db.Tables Then ReDim Preserve db.Table(1 To db.Tables)
    SortTables
End Sub

Private Sub LoadTable(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As TableType
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.TableID = Trim$(strLine(0))
    If Len(typNew.TableID) = 0 Then Exit Sub
    With typNew
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "title"
                        .Title = strItem
                    Case "columns"
                        .Columns = lngValue
                        ReDim .Column(1 To .Columns)
                    Case "rows"
                        ReDim .Row(1 To lngValue)
                    Case "group"
                        .Group = lngValue
                    Case "width"
                        If lngListMax = .Columns - 1 Then
                            For i = 1 To .Columns
                                .Column(i).Widest = strList(i - 1)
                            Next
                        Else
                            Debug.Print "Error reading Tables.txt: Width entry does not match Columns count (" & .TableID & ")"
                        End If
                    Case "style"
                        If lngListMax = .Columns - 1 Then
                            For i = 1 To .Columns
                                Select Case strList(i - 1)
                                    Case "TextLeft": .Column(i).Style = tcseTextLeft
                                    Case "TextRight": .Column(i).Style = tcseTextRight
                                    Case "TextCenter": .Column(i).Style = tcseTextCenter
                                End Select
                            Next
                        Else
                            Debug.Print "Error reading Tables.txt: Style entry does not match Columns count (" & .TableID & ")"
                        End If
                    Case "header"
                        If lngListMax = .Columns - 1 Then
                            .Headers = True
                            For i = 1 To .Columns
                                .Column(i).Header = strList(i - 1)
                            Next
                        Else
                            Debug.Print "Error reading Tables.txt: Header entry does not match Columns count (" & .TableID & ")"
                        End If
                    Case "row"
                        If lngListMax = .Columns - 1 Then
                            .Rows = .Rows + 1
                            With .Row(.Rows)
                                ReDim .Value(1 To typNew.Columns)
                                For i = 1 To typNew.Columns
                                    If typNew.Column(i).Style = tcseNumeric Then .Value(i) = Val(strList(i - 1)) Else .Value(i) = strList(i - 1)
                                Next
                            End With
                        Else
                            Debug.Print "Error reading Tables.txt: Row entry does not match Columns count (" & .TableID & ")"
                        End If
                    Case Else
                        Debug.Print "Invalid line in Tables.txt: " & strField & ": " & strItem
                End Select
            End If
        Next
    End With
    db.Tables = db.Tables + 1
    db.Table(db.Tables) = typNew
End Sub

' Insertion sort because there's only a handful (quicksort would be massive overkill)
Private Sub SortTables(Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As TableType
    
    iMin = 2
    iMax = db.Tables
    For i = iMin To iMax
        typSwap = db.Table(i)
        For j = i To iMin Step -1
            If typSwap.TableID < db.Table(j - 1).TableID Then db.Table(j) = db.Table(j - 1) Else Exit For
        Next j
        db.Table(j) = typSwap
    Next i
End Sub

Public Function SeekTable(pstrTableID As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long

    lngFirst = 1
    lngLast = db.Tables
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Table(lngMid).TableID > pstrTableID Then
            lngLast = lngMid - 1
        ElseIf db.Table(lngMid).TableID < pstrTableID Then
            lngFirst = lngMid + 1
        Else
            SeekTable = lngMid
            Exit Do
        End If
    Loop
End Function


' ************* CHALLENGES *************


Private Sub LoadChallenges()
    Dim strFile As String
    Dim strChallenge() As String
    Dim lngChallenges As Long
    Dim i As Long
    
    strFile = DataPath() & "Challenges.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strChallenge = Split(xp.File.LoadToString(strFile), "ChallengeName: ")
    lngChallenges = UBound(strChallenge)
    If lngChallenges = 0 Then Exit Sub
    ReDim db.Challenge(1 To lngChallenges + 1)
    For i = 1 To UBound(strChallenge)
        LoadChallenge strChallenge(i)
    Next
    If UBound(db.Challenge) <> db.Challenges Then ReDim Preserve db.Challenge(1 To db.Challenges)
    SortChallenges
End Sub

Private Sub LoadChallenge(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As ChallengeType
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Challenge = Trim$(strLine(0))
    If Len(typNew.Challenge) = 0 Then Exit Sub
    With typNew
        .ID = .Challenge
        .MaxStars = 6
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "id": .ID = strItem
                    Case "wiki": .Wiki = strItem
                    Case "group": .Group = strItem
                    Case "patron": .Patron = strItem
                    Case "levellow": .LevelLow = lngValue
                    Case "levelhigh": .LevelHigh = lngValue
                    Case "order": .GameOrder = lngValue
                    Case "maxstars": .MaxStars = lngValue
                End Select
            End If
        Next
    End With
    db.Challenges = db.Challenges + 1
    typNew.GroupOrder = db.Challenges
    db.Challenge(db.Challenges) = typNew
End Sub

' Omit plngLeft & plngRight; they are used internally during recursion
Private Sub SortChallenges(Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim strMid As String
    Dim typSwap As ChallengeType
    
    If plngRight = 0 Then
        plngLeft = 1
        plngRight = db.Challenges
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    strMid = db.Challenge((plngLeft + plngRight) \ 2).ID
    Do
        Do While db.Challenge(lngFirst).ID < strMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While strMid < db.Challenge(lngLast).ID And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            typSwap = db.Challenge(lngFirst)
            db.Challenge(lngFirst) = db.Challenge(lngLast)
            db.Challenge(lngLast) = typSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then SortChallenges plngLeft, lngLast
    If lngFirst < plngRight Then SortChallenges lngFirst, plngRight
End Sub

Public Function SeekChallenge(pstrID As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long

    lngFirst = 1
    lngLast = db.Challenges
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Challenge(lngMid).ID > pstrID Then
            lngLast = lngMid - 1
        ElseIf db.Challenge(lngMid).ID < pstrID Then
            lngFirst = lngMid + 1
        Else
            SeekChallenge = lngMid
            Exit Do
        End If
    Loop
End Function


' ************* LINKLIST TEMPLATES *************


Private Sub LoadTemplates()
    Dim strFile As String
    Dim strTemplate() As String
    Dim lngTemplates As Long
    Dim i As Long
    
    strFile = DataPath() & "LinkLists.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strTemplate = Split(xp.File.LoadToString(strFile), "LinkList: ")
    lngTemplates = UBound(strTemplate)
    If lngTemplates = 0 Then Exit Sub
    ReDim db.Template(1 To lngTemplates + 1)
    For i = 1 To UBound(strTemplate)
        LoadTemplate strTemplate(i)
    Next
    If UBound(db.Template) <> db.Templates Then ReDim Preserve db.Template(1 To db.Templates)
End Sub

Private Sub LoadTemplate(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As MenuType
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Title = Trim$(strLine(0))
    If Len(typNew.Title) = 0 Then Exit Sub
    With typNew
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "menu"
                        If lngListMax = 3 Then
                            .Commands = .Commands + 1
                            ReDim Preserve .Command(1 To .Commands)
                            With .Command(.Commands)
                                .Style = GetMenuStyleID(strList(0))
                                .Caption = strList(1)
                                .Target = strList(2)
                                .Param = strList(3)
                            End With
                        End If
                End Select
            End If
        Next
    End With
    db.Templates = db.Templates + 1
    db.Template(db.Templates) = typNew
End Sub


' ************* LINKLISTS *************


Public Sub LoadLinkLists()
    Dim strFile As String
    Dim strBackup As String
    Dim strRaw As String
    Dim strLinkList() As String
    Dim lngLinkLists As Long
    Dim i As Long
    
    Erase db.LinkList
    db.LinkLists = 0
    strFile = LinkListsFile()
    If xp.File.Exists(strFile) Then
        strRaw = xp.File.LoadToString(strFile)
        strLinkList = Split(strRaw, "LinkList: ")
        lngLinkLists = UBound(strLinkList)
        If lngLinkLists = 0 Then Exit Sub
        ReDim db.LinkList(1 To lngLinkLists + 1)
        For i = 1 To UBound(strLinkList)
            LoadLinkList strLinkList(i)
        Next
        If UBound(db.LinkList) <> db.LinkLists Then
            If db.LinkLists = 0 Then Erase db.LinkList Else ReDim Preserve db.LinkList(1 To db.LinkLists)
        End If
    End If
End Sub

Private Sub LoadLinkList(pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As MenuType
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Title = Trim$(strLine(0))
    If Len(typNew.Title) = 0 Then Exit Sub
    With typNew
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "left": .Left = lngValue
                    Case "top": .Top = lngValue
                    Case "deleted": If LCase$(strItem) = "true" Then .Deleted = True
                    Case "menu"
                        If lngListMax = 3 Then
                            .Commands = .Commands + 1
                            ReDim Preserve .Command(1 To .Commands)
                            With .Command(.Commands)
                                .Style = GetMenuStyleID(strList(0))
                                .Caption = strList(1)
                                .Target = strList(2)
                                .Param = strList(3)
                            End With
                        End If
                End Select
            End If
        Next
    End With
    db.LinkLists = db.LinkLists + 1
    db.LinkList(db.LinkLists) = typNew
End Sub


' ************* COMPENDIUM *************


Public Sub OpenCompendiumFile(pblnCreate As Boolean)
    Dim frm As Form
    
    db.Characters = 0
    Erase db.Character
    AllocateCharacters
    LoadCompendium pblnCreate
    CalculateFavor
    If GetForm(frm, "frmPatrons") Then frm.ReDrawForm
    If GetForm(frm, "frmChallenges") Then frm.DataFileChanged
    If GetForm(frm, "frmCharacter") Then frm.DataFileChanged
    If GetForm(frm, "frmSagas") Then frm.DataFileChanged
    frmCompendium.RedrawQuests
    Erase gblnDirtyFlag
    frmCompendium.SetCaption
End Sub

Public Sub ChangeCompendiumFile(pstrDataFile As String)
    Dim strBackup As String
    
    gblnDirtyFlag(dfeData) = True
    AutoSave
    strBackup = CompendiumBackupFile()
    If xp.File.Exists(strBackup) Then xp.File.Delete strBackup
    cfg.DataFile = pstrDataFile
    OpenCompendiumFile True
    BackupCompendium
End Sub

Public Sub LoadCompendium(Optional pblnCreate As Boolean = False)
    Dim strFile As String
    Dim strSection() As String
    Dim lngSections As Long
    Dim strName As String
    Dim lngPos As Long
    Dim i As Long
    Dim s As Long
    
    If Len(cfg.DataFile) = 0 Then Exit Sub
    strFile = CompendiumFile()
    If Not xp.File.Exists(strFile) Then
        If pblnCreate Then SaveCompendiumFile Else Exit Sub
    End If
    strSection = Split(xp.File.LoadToString(strFile), "SectionName: ")
    lngSections = UBound(strSection)
    If lngSections = 0 Then Exit Sub
    For i = 1 To UBound(strSection)
        lngPos = InStr(strSection(i), vbNewLine)
        If lngPos > 1 Then
            strName = LCase$(Trim$(Left$(strSection(i), lngPos - 1)))
            Select Case strName
                Case "settings": CompendiumSettings strSection(i)
                Case "characters": CompendiumCharacters strSection(i)
                Case "quests": CompendiumQuests strSection(i)
                Case "challenges": CompendiumChallenges strSection(i)
            End Select
        End If
    Next
    If cfg.CompendiumBackColor <> cfg.GetColor(cgeControls, cveBackground) Then MatchColors True
End Sub

Private Sub CompendiumSettings(pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    ' Process lines
    For lngLine = 1 To UBound(strLine)
        If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
            Select Case strField
                Case "backcolor": cfg.CompendiumBackColor = lngValue
            End Select
        End If
    Next
End Sub

Private Sub CompendiumCharacters(pstrRaw As String)
    Dim strFile As String
    Dim strCharacter() As String
    Dim lngCharacters As Long
    Dim i As Long
    
    strCharacter = Split(pstrRaw, "Character: ")
    lngCharacters = UBound(strCharacter)
    If lngCharacters = 0 Then Exit Sub
    ReDim db.Character(1 To lngCharacters + 1)
    db.Characters = 0
    For i = 1 To UBound(strCharacter)
        CompendiumCharacter strCharacter(i)
    Next
    If UBound(db.Character) <> db.Characters Then ReDim Preserve db.Character(1 To db.Characters)
    AllocateCharacters
End Sub

Private Sub AllocateCharacters()
    Dim i As Long
    
    For i = 1 To db.Quests
        With db.Quest(i)
            If db.Characters = 0 Then Erase .Progress Else ReDim .Progress(1 To db.Characters)
        End With
    Next
    For i = 1 To db.Challenges
        With db.Challenge(i)
            If db.Characters = 0 Then Erase .Stars Else ReDim .Stars(1 To db.Characters)
        End With
    Next
End Sub

Private Sub CompendiumCharacter(pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As CharacterType
    Dim lngSaga As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    typNew.Character = Trim$(strLine(0))
    If Len(typNew.Character) = 0 Then Exit Sub
    ' Allocate sagas
    InitCharacterSagas typNew
    With typNew
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "level": .Level = lngValue
                    Case "customcolor": .CustomColor = (LCase$(strItem) = "true")
                    Case "generatedcolor": .GeneratedColor = GetColorID(strItem)
                    Case "backcolor": .BackColor = lngValue
                    Case "dimcolor": .DimColor = lngValue
                    Case "leftclick": .LeftClick = strItem
                    Case "notes": If Len(.Notes) Then .Notes = .Notes & vbNewLine & strItem Else .Notes = strItem
                    Case "saga"
                        If lngListMax = 1 Then
                            lngSaga = SeekSaga(strList(0))
                            With .Saga(lngSaga)
                                For i = 1 To db.Saga(lngSaga).Quests
                                    .Progress(i) = GetProgressID(Mid$(strList(1), i, 1))
                                Next
                            End With
                        Else
                            Debug.Print "Invalid saga for " & .Character & ": " & strItem
                        End If
                    Case "menu"
                        If lngListMax = 3 Then
                            With .ContextMenu
                                .Commands = .Commands + 1
                                ReDim Preserve .Command(1 To .Commands)
                                With .Command(.Commands)
                                    .Style = GetMenuStyleID(strList(0))
                                    .Caption = strList(1)
                                    .Target = strList(2)
                                    .Param = strList(3)
                                End With
                            End With
                        Else
                            Debug.Print "Invalid menu command for " & .Character & ": " & strItem
                        End If
                    Case "tomestat"
                        For i = 1 To 6
                            .Tome.Stat(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, tomes.Stat.Max)
                        Next
                    Case "tomeskill"
                        For i = 1 To 21
                            .Tome.Skill(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, tomes.Skill.Max)
                        Next
                    Case "tomeracialap"
                        .Tome.RacialAP = LimitValue(Val(strItem), 0, tomes.RacialAPMax)
                    Case "tomefate"
                        .Tome.Fate = LimitValue(Val(strItem), 0, tomes.FateMax)
                    Case "tomepower"
                        For i = 1 To 3
                            .Tome.Power(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, tomes.PowerMax)
                        Next
                    Case "tomerr"
                        For i = 1 To 2
                            .Tome.RR(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, tomes.RRMax)
                        Next
                    Case "tomeheroicxp"
                        If strItem = "Lesser" Or strItem = "Greater" Then .Tome.HerociXP = strItem
                    Case "tomeepicxp"
                        If strItem = "Lesser" Or strItem = "Greater" Then .Tome.EpicXP = strItem
                    Case "pastlifeclass"
                        For i = 1 To maxClassCount
                            .PastLife.Class(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                    Case "pastliferace"
                        For i = 1 To maxRaceCount
                            .PastLife.Racial(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                    Case "pastlifeiconic"
                        For i = 1 To maxIconicCount
                            .PastLife.Iconic(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                    Case "pastlifeepic"
                        For i = 1 To maxEpicCount
                            .PastLife.Epic(i) = LimitValue(Val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                End Select
            End If
        Next
    End With
    db.Characters = db.Characters + 1
    db.Character(db.Characters) = typNew
End Sub

Private Function LimitValue(plngValue As Long, plngMin As Long, plngMax As Long) As Long
    If plngValue < plngMin Then
        LimitValue = plngMin
    ElseIf plngValue > plngMax Then
        LimitValue = plngMax
    Else
        LimitValue = plngValue
    End If
End Function

Public Sub InitCharacterSagas(ptypCharacter As CharacterType)
    Dim i As Long
    
    Erase ptypCharacter.Saga
    If db.Sagas Then
        ReDim ptypCharacter.Saga(db.Sagas)
        For i = 1 To db.Sagas
            ReDim ptypCharacter.Saga(i).Progress(1 To db.Saga(i).Quests)
        Next
    End If
End Sub

Private Sub CompendiumQuests(pstrRaw As String)
    Dim strQuest() As String
    Dim strToken() As String
    Dim lngQuest As Long
    Dim q As Long
    Dim c As Long
    
    strQuest = Split(pstrRaw, vbNewLine)
    For q = 1 To UBound(strQuest)
        strToken = Split(strQuest(q), vbTab)
        If UBound(strToken) = 2 Then
            lngQuest = SeekQuest(strToken(0))
            If lngQuest Then
                With db.Quest(lngQuest)
                    For c = 1 To db.Characters
                        .Progress(c) = GetProgressID(Mid$(strToken(1), c, 1))
                    Next
                    .Skipped = (InStr(strToken(2), "k") <> 0)
                End With
            End If
        End If
    Next
End Sub

Private Sub CompendiumChallenges(pstrRaw As String)
    Dim strChallenge() As String
    Dim strToken() As String
    Dim lngChallenge As Long
    Dim i As Long
    Dim c As Long
    
    strChallenge = Split(pstrRaw, vbNewLine)
    For i = 1 To UBound(strChallenge)
        strToken = Split(strChallenge(i), vbTab)
        If UBound(strToken) = 1 Then
            lngChallenge = SeekChallenge(strToken(0))
            If lngChallenge Then
                With db.Challenge(lngChallenge)
                    For c = 1 To db.Characters
                        .Stars(c) = Val(Mid$(strToken(1), c, 1))
                    Next
                End With
            End If
        End If
    Next
End Sub


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
    pstrLine = Trim$(pstrLine)
    If Len(pstrLine) = 0 Then Exit Function
    If Left$(pstrLine, 1) = ";" Then Exit Function
    ' Field
    lngPos = InStr(pstrLine, ":")
    If lngPos = 0 Then Exit Function
    ParseLine = True
    pstrField = LCase$(Trim$(Left$(pstrLine, lngPos - 1)))
    pstrLine = Trim$(Mid$(pstrLine, lngPos + 1))
    ' Descriptions
    If Left$(pstrField, 4) = "wiki" Or pstrField = "notes" Then
        pstrItem = pstrLine
        Exit Function
    End If
    ' List
    If InStr(pstrLine, vbTab) Then
        pstrList = Split(pstrLine, vbTab)
        plngListMax = UBound(pstrList)
'        For i = 0 To plngListMax
'            pstrList(i) = Trim$(pstrList(i))
'        Next
        Exit Function
    End If
    ' Value
    pstrItem = pstrLine
    If Left$(pstrField, 4) <> "tome" And Left$(pstrField, 8) <> "pastlife" Then
        If IsNumeric(pstrItem) Then plngValue = Val(pstrItem)
    End If
    ' Return single item in list form as well
    plngListMax = 0
    ReDim pstrList(0)
    pstrList(0) = pstrLine
End Function

