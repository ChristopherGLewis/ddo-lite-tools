Attribute VB_Name = "basFormat"
' Format modules handle the various file formats:
'   basFormat.bas               General helper routines shared by multiple formats
'   basFormatLite.bas          Save/Load native Character Builder Lite text files
'   basFormatRon.bas          Import/Export Ron's Character Planner build files
'   basFormatBuilder.bas      Import/Export DDO Builder xml files
Option Explicit

Public Type ExportFeatType
    Type As BuildFeatTypeEnum
    Level As Long
    Class As ClassEnum
    ClassLevel As Long
    LiteName As String
    RonName As String
    RonType As RonSlotEnum
    BuilderName As String
    BuilderType As String
End Type

Public Type ExportFeatsType
    Feat() As ExportFeatType
    Feats As Long
    RonFeats As Long
    BuilderFeats As Long
End Type

Public Export() As ExportFeatsType


' ************* ALIGNMENT *************


Public Function GetExportAlignment() As AlignmentEnum
    Dim blnFilterAlign(6) As Boolean
    Dim lngClass As Long
    Dim enClass As ClassEnum
    Dim enType As BuildFeatTypeEnum
    Dim lngAlign As Long
    Dim lngFeat As Long
    Dim i As Long
    
    If build.Alignment <> aleAny Then
        GetExportAlignment = build.Alignment
        Exit Function
    End If
    ' If no alignment is specified, return a valid possible alignment
    For lngAlign = 0 To 6
        blnFilterAlign(lngAlign) = True
    Next
    ' Apply class restrictions
    For lngClass = 0 To 2
        enClass = build.BuildClass(lngClass)
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) And Not db.Class(enClass).Alignment(lngAlign) Then blnFilterAlign(lngAlign) = False
        Next
    Next
    ' Apply class feat restrictions (Warlock Pact and Cleric Domain)
    For enType = bftClass1 To bftClass3
        With build.Feat(enType)
            For i = 1 To .Feats
                With .Feat(i)
                    If .Selector <> 0 Then
                        If .FeatName = "Pact" Or .FeatName = "Domain" Then
                            lngFeat = SeekFeat(.FeatName)
                            If lngFeat Then
                                With db.Feat(lngFeat).Selector(.Selector)
                                    For lngAlign = 1 To 6
                                        If blnFilterAlign(lngAlign) And Not .Alignment(lngAlign) Then blnFilterAlign(lngAlign) = False
                                    Next
                                End With
                            End If
                        End If
                    End If
                End With
            Next
        End With
    Next
    ' Choose alignments in order of most to least preferable
    If blnFilterAlign(aleTrueNeutral) = True Then
        GetExportAlignment = aleTrueNeutral
    ElseIf blnFilterAlign(aleLawfulNeutral) = True Then
        GetExportAlignment = aleLawfulNeutral
    ElseIf blnFilterAlign(aleChaoticNeutral) = True Then
        GetExportAlignment = aleChaoticNeutral
    ElseIf blnFilterAlign(aleNeutralGood) = True Then
        GetExportAlignment = aleNeutralGood
    ElseIf blnFilterAlign(aleLawfulGood) = True Then
        GetExportAlignment = aleLawfulGood
    ElseIf blnFilterAlign(aleChaoticGood) = True Then
        GetExportAlignment = aleChaoticGood
    Else ' No valid alignment (great old one with monk levels?) so just return true neutral
        GetExportAlignment = aleTrueNeutral
    End If
End Function


' ************* STAT POINTS *************


Public Function GetExportStatRaise(penStat As StatEnum) As Long
    Dim lngPoints As Long
    
    lngPoints = build.StatPoints(build.BuildPoints, penStat)
    Select Case lngPoints
        Case 8: GetExportStatRaise = 7
        Case 10: GetExportStatRaise = 8
        Case 13: GetExportStatRaise = 9
        Case 16: GetExportStatRaise = 10
        Case Else: GetExportStatRaise = lngPoints
    End Select
End Function

Public Function GetImportStatRaise(ByVal plngPoints As Long) As Long
    Select Case plngPoints
        Case 7: GetImportStatRaise = 8
        Case 8: GetImportStatRaise = 10
        Case 9: GetImportStatRaise = 13
        Case 10: GetImportStatRaise = 16
        Case Else: GetImportStatRaise = plngPoints
    End Select
End Function


' ************* PAST LIVES *************


Public Function GetExportPastLives(plngLives() As Long) As Long
    Dim lngLives As Long
    Dim lngSelector As Long
    Dim lngPastLife As Long
    Dim enClass As ClassEnum
    Dim i As Long
    Dim j As Long
    
    ReDim plngLives(1 To ceClasses - 1)
    Select Case build.BuildPoints
        Case beAdventurer, beChampion
        Case beHero, beLegend
            ' For hero and legend builds...
            lngPastLife = SeekFeat("Past Life")
            For i = 1 To build.Feat(bftStandard).Feats
                Select Case build.Feat(bftStandard).Feat(i).FeatName
                    ' ...add a past life for every active past life feat taken
                    Case "Past Life"
                        lngSelector = build.Feat(bftStandard).Feat(i).Selector
                        If lngPastLife Then
                            If lngSelector > 0 And lngSelector <= db.Feat(lngPastLife).Selectors Then
                                enClass = GetClassID(db.Feat(lngPastLife).Selector(lngSelector).SelectorName)
                                If enClass <> ceAny Then plngLives(enClass) = 1
                            End If
                        End If
                    ' ...completionists get 1 life on every class
                    Case "Completionist"
                        For j = 1 To ceClasses - 1
                            plngLives(j) = 1
                        Next
                End Select
            Next
            For i = 1 To ceClasses - 1
                lngLives = lngLives + plngLives(i)
            Next
    End Select
    ' At this point, enClass will be set to the last class that had a past life feat, or ceAny if we couldn't find any past life feats
    Select Case build.BuildPoints
        Case beHero
            ' Hero builds must have exactly one past life...
            If lngLives <> 1 Then
                For i = 0 To 2
                    If build.BuildClass(i) <> ceAny Then
                        If plngLives(build.BuildClass(i)) = 1 Then
                            enClass = build.BuildClass(i)
                            Exit For
                        End If
                    End If
                Next
                If enClass = ceAny Then enClass = build.BuildClass(0)
                If enClass = ceAny Then enClass = ceFighter
                ReDim plngLives(1 To ceClasses - 1)
                plngLives(enClass) = 1
            End If
            lngLives = 1
        Case beLegend
            ' Legend builds must have at least two past lives. If we're short, make up the difference by adding primary class lives
            If lngLives < 2 Then
                enClass = build.BuildClass(0)
                If enClass = ceAny Then enClass = ceFighter
                plngLives(enClass) = plngLives(enClass) + 2 - lngLives
                lngLives = 2
            End If
    End Select
    GetExportPastLives = lngLives
End Function


' ************* FEATS *************


Public Sub InitExportFeats()
    Dim lngLevel As Long
    Dim i As Long
    
    SortFeatMap peLite
    ReDim Export(1 To 30)
    TransposeFeats
    For lngLevel = 1 To 30
        For i = 1 To Export(lngLevel).Feats
            IdentifyChannels Export(lngLevel).Feat(i)
            With Export(lngLevel)
                If Len(.Feat(i).RonName) <> 0 And .Feat(i).RonType <> rseUnknown Then .RonFeats = .RonFeats + 1
                If Len(.Feat(i).BuilderName) <> 0 And Len(.Feat(i).BuilderType) <> 0 Then .BuilderFeats = .BuilderFeats + 1
            End With
        Next
    Next
End Sub

Private Sub TransposeFeats()
    Dim lngIndex As Long
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim strLite As String
    Dim i As Long
    
    For i = 1 To Feat.Count
        Select Case Feat.List(i).ActualType
            Case bftGranted, bftAlternate, bftExchange, bftUnknown
            Case Else
                ' Start by getting the FeatID and Selector, which can be supplied by an exchange feat for this slot
                lngIndex = Feat.List(i).ExchangeIndex
                If lngIndex = 0 Then
                    lngIndex = i
                ElseIf Feat.List(lngIndex).FeatID = 0 Then
                    lngIndex = i
                End If
                lngFeat = Feat.List(lngIndex).FeatID
                lngSelector = Feat.List(lngIndex).Selector
                ' Validate selector
                With db.Feat(lngFeat)
                    If lngSelector = 0 And .Selectors <> 0 Then
                        lngFeat = 0
                    ElseIf lngSelector <> 0 And .Selectors = 0 Then
                        lngFeat = 0
                    ElseIf lngSelector > .Selectors Then
                        lngFeat = 0
                    End If
                End With
                ' All other relevant info comes from the actual feat slot, which is Feat.List(i)
                If lngFeat Then
                    strLite = db.Feat(lngFeat).FeatName
                    If lngSelector <> 0 Then strLite = strLite & ": " & db.Feat(lngFeat).Selector(lngSelector).SelectorName
                    lngIndex = SeekFeatMap(strLite)
                    If lngIndex Then
                        With Export(Feat.List(i).Level)
                            .Feats = .Feats + 1
                            ReDim Preserve .Feat(1 To .Feats)
                            With .Feat(.Feats)
                                .Type = Feat.List(i).ActualType
                                .Level = Feat.List(i).Level
                                .Class = Feat.List(i).Class
                                .ClassLevel = Feat.List(i).ClassLevel
                                .LiteName = strLite
                                .BuilderName = db.FeatMap(lngIndex).Builder
                                .RonName = db.FeatMap(lngIndex).Ron
                            End With
                        End With
                    End If
                End If
        End Select
    Next
End Sub

Private Sub IdentifyChannels(ptypFeat As ExportFeatType)
    With ptypFeat
        Select Case .Type
            Case bftStandard
                Select Case .Level
                    Case 1, 3, 6, 9, 12, 15, 18
                        .BuilderType = "Standard"
                        .RonType = rseStandard
                    Case 21, 24, 27, 30
                        .BuilderType = "EpicFeat"
                        .RonType = rseStandard
                    Case 26, 28, 29
                        .BuilderType = "EpicDestinyFeat"
                        .RonType = rseDestiny
                End Select
            Case bftLegend
                .BuilderType = "Legendary"
                .RonType = rseLegend
            Case bftDeity
                Select Case .ClassLevel
                    Case 1: .BuilderType = "FollowerOf"
                    Case 3: .BuilderType = "ChildOf"
                    Case 6: .BuilderType = "Deity"
                    Case 12: .BuilderType = "BelovedOf"
                    Case 20: .BuilderType = "DamageReduction"
                End Select
                .RonType = rseDeity
            Case bftRace
                Select Case build.Race
                    Case reHuman
                        .BuilderType = "HumanBonus"
                        .RonType = rseHuman
                    Case reHalfElf
                        .BuilderType = "Dilettante"
                        .RonType = rseDilettante
                    Case rePurpleDragonKnight
                        .BuilderType = "PDKBonus"
                        .RonType = rseHuman
                    Case reDragonborn
                        .BuilderType = "DragonbornRacial"
                        .RonType = rseDragonborn
                    Case reAasimar
                        .BuilderType = "AasimarBond"
                        .RonType = rseBond
                End Select
            Case bftClass1, bftClass2, bftClass3
                Select Case .Class
                    Case ceArtificer
                        .BuilderType = "ArtificerBonus"
                        .RonType = rseArtificer
                    Case ceDruid
                        .BuilderType = "DruidWildShape"
                        .RonType = rseDruid
                    Case ceFighter
                        .BuilderType = "FighterBonus"
                        .RonType = rseFighter
                    Case ceRanger
                        .BuilderType = "FavoredEnemy"
                        .RonType = rseFavoredEnemy
                    Case ceRogue
                        .BuilderType = "RogueSpecialAbility"
                        .RonType = rseRogue
                    Case ceWarlock
                        .BuilderType = "WarlockPact"
                        .RonType = rseWarlockPact
                    Case ceWizard
                        .BuilderType = "Metamagic"
                        .RonType = rseWizard
                    Case ceFavoredSoul
                        Select Case .ClassLevel
                            Case 5, 10, 15: .BuilderType = "EnergyResistance"
                            Case 2: .BuilderType = "FavoredSoulBattle"
                            Case 7: .BuilderType = "FavoredSoulHeart"
                        End Select
                        .RonType = rseFavoredSoul
                    Case ceCleric
                        .BuilderType = "Domain"
                        .RonType = rseDomain
                    Case ceMonk
                        Select Case .ClassLevel
                            Case 1, 2
                                .BuilderType = "MonkBonus"
                                .RonType = rseMonkBonus
                            Case 3
                                .BuilderType = "MonkPhilosophy"
                                .RonType = rseMonkPath
                            Case 6
                                .BuilderType = "MonkBonus6"
                                .RonType = rseMonkBonus
                        End Select
                End Select
        End Select
    End With
End Sub

Public Sub CloseExportFeats()
    Erase Export
End Sub

' Shellsort
Public Sub SortFeatMap(penIndex As PlannerEnum)
    Dim lngHold As Long
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As FeatMapType
    
    If db.FeatMapIndex = penIndex Then Exit Sub
    db.FeatMapIndex = penIndex
    If db.FeatMapIndex = peUnknown Then Exit Sub
    iMin = 1
    iMax = db.FeatMaps
    lngGap = iMin
    Do
        lngGap = 3 * lngGap + 1
    Loop Until lngGap > iMax
    Do
        lngGap = lngGap \ 3
        For i = lngGap + iMin To iMax
            typSwap = db.FeatMap(i)
            lngHold = i
            Do While CompareFeatMap(db.FeatMap(lngHold - lngGap), typSwap) = 1
                db.FeatMap(lngHold) = db.FeatMap(lngHold - lngGap)
                lngHold = lngHold - lngGap
                If lngHold < iMin + lngGap Then Exit Do
            Loop
            db.FeatMap(lngHold) = typSwap
        Next i
    Loop Until lngGap = 1
End Sub

' Simple binary search
Public Function SeekFeatMap(pstrDisplay As String) As Long
    Dim typSearch As FeatMapType
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    Select Case db.FeatMapIndex
        Case peLite: typSearch.Lite = pstrDisplay
        Case peRon: typSearch.Ron = pstrDisplay
        Case peBuilder: typSearch.Builder = pstrDisplay
    End Select
    lngFirst = 1
    lngLast = db.FeatMaps
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        Select Case CompareFeatMap(db.FeatMap(lngMid), typSearch)
            Case -1
                lngFirst = lngMid + 1
            Case 1
                lngLast = lngMid - 1
            Case 0
                SeekFeatMap = lngMid
                Exit Function
        End Select
    Loop
End Function

Private Function CompareFeatMap(ptypLeft As FeatMapType, ptypRight As FeatMapType) As Long
    Select Case db.FeatMapIndex
        Case peLite
            If ptypLeft.Lite > ptypRight.Lite Then
                CompareFeatMap = 1
            ElseIf ptypLeft.Lite < ptypRight.Lite Then
                CompareFeatMap = -1
            Else
                CompareFeatMap = 0
            End If
        Case peRon
            If ptypLeft.Ron > ptypRight.Ron Then
                CompareFeatMap = 1
            ElseIf ptypLeft.Ron < ptypRight.Ron Then
                CompareFeatMap = -1
            Else
                CompareFeatMap = 0
            End If
        Case peBuilder
            If ptypLeft.Builder > ptypRight.Builder Then
                CompareFeatMap = 1
            ElseIf ptypLeft.Builder < ptypRight.Builder Then
                CompareFeatMap = -1
            Else
                CompareFeatMap = 0
            End If
    End Select
End Function

