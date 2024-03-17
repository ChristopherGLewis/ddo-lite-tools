Attribute VB_Name = "basData"
Option Explicit

Public Const MAX_CRAFTING_LEVEL = 32


Public Sub InitData()
    Dim typBlank As DatabaseType
    
    db = typBlank
    LoadGeneral
    LoadFarms
    LoadSchools
    LoadScaling
    LoadMaterials
    LoadShards
    IndexShards
    LoadItems
    LoadItemCombo
    LoadItemChoices
    LoadAugments
    LoadAugmentItems
    LoadAugmentVendors
    LoadRituals
    InitHelp
    CalculateValues
    DeleteOldFiles
End Sub

Public Function DataPath() As String
    DataPath = App.Path & "\Data\Cannith\"
End Function

Private Sub DeleteOldFiles()
    DeleteOldFile "Supply.txt"
    DeleteOldFile "Vendors.txt"
End Sub

Private Sub DeleteOldFile(pstrFile As String)
    Dim strFile As String
    
    strFile = DataPath() & pstrFile
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
End Sub


' ************* GENERAL *************


Private Sub LoadGeneral()
    Dim strFile As String
    Dim strRaw As String
    Dim strLine() As String
    Dim strField As String
    Dim strItem As String
    Dim lngPos As Long
    Dim i As Long
    
    SetGeneralDefaults
    strFile = DataPath() & "General.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strLine = Split(strRaw, vbNewLine)
    For i = 0 To UBound(strLine)
        If Len(strLine(i)) <> 0 And Left$(strLine(i), 1) <> "'" Then
            lngPos = InStr(strLine(i), ":")
            If lngPos <> 0 Then
                strField = LCase$(Left$(strLine(i), lngPos - 1))
                strItem = Trim$(Mid$(strLine(i), lngPos + 1))
                Select Case strField
                    Case "frequency": SetGeneral db.Frequency, strItem
                    Case "any": SetGeneral db.Backpack, strItem
                    Case "essencerate": db.EssenceRate = val(strItem)
                    Case "demand": SetDemandStyle LCase$(strItem)
                    Case "demandvalue": SetGeneral db.DemandValue, LCase$(strItem)
                    Case "demandweights": SetGeneral db.DemandWeight, strItem
                    Case Else: Debug.Print "Invalid line in General.txt" & vbNewLine & strLine(i) & vbNewLine
                End Select
            End If
        End If
    Next
End Sub

Private Sub SetGeneralDefaults()
    db.EssenceRate = 700 / 3600
    ReDim db.Frequency(1 To 3)
    db.Frequency(feCommon) = 0.7
    db.Frequency(feUncommon) = 0.2
    db.Frequency(feRare) = 0.05
    ReDim db.Backpack(1 To 4)
    db.Backpack(seArcane) = 0.22
    db.Backpack(seCultural) = 0.5
    db.Backpack(seLore) = 0.2
    db.Backpack(seNatural) = 0.08
    ReDim db.DemandValue(1 To 6)
    db.DemandValue(deUniversal) = 10
    db.DemandValue(deCommon) = 8
    db.DemandValue(deUncommon) = 6
    db.DemandValue(deNiche) = 4
    db.DemandValue(deObsolete) = 2
    db.DemandValue(deRarelyUsed) = 1
    ReDim db.DemandWeight(1 To 6)
    db.DemandWeight(deUniversal) = 99
    db.DemandWeight(deCommon) = 5
    db.DemandWeight(deUncommon) = 3
    db.DemandWeight(deNiche) = 2
    db.DemandWeight(deObsolete) = 1
    db.DemandWeight(deRarelyUsed) = 1
    db.DemandStyle = dseTop
    db.DemandTop = 5
End Sub

Private Sub SetGeneral(pvarArray As Variant, pstrList As String)
    Dim strToken() As String
    Dim i As Long
    
    If InStr(pstrList, " ") Then pstrList = Replace(pstrList, " ", vbNullString)
    strToken = Split(pstrList, ",")
    For i = 0 To UBound(strToken)
        pvarArray(i + 1) = val(strToken(i))
    Next
End Sub

Private Sub SetDemandStyle(pstrStyle As String)
    Select Case pstrStyle
        Case "raw"
            db.DemandStyle = dseRaw
        Case "weighted"
            db.DemandStyle = dseWeighted
        Case Else
            If Left$(pstrStyle, 3) = "top" Then
                db.DemandTop = val(Trim$(Mid$(pstrStyle, 4)))
                If db.DemandTop = 0 Then Debug.Print "Invalid demand style in General.txt:" & vbNewLine & pstrStyle
            Else
                Debug.Print "Invalid demand style in General.txt:" & vbNewLine & pstrStyle
            End If
    End Select
End Sub


' ************* MATERIALS *************


Private Sub LoadMaterials()
    Dim strItem() As String
    Dim i As Long
    
    Erase db.Material
    db.Materials = 0
    If SplitFile("Collectables.txt", "Collectable: ", strItem) Then
        ReDim Preserve db.Material(1 To db.Materials + UBound(strItem) + 1)
        For i = 1 To UBound(strItem)
            If InStr(strItem(i), "School: ") Then LoadMaterialCollectable strItem(i)
        Next
    End If
    If SplitFile("SoulGems.txt", "SoulGem: ", strItem) Then
        ReDim Preserve db.Material(1 To db.Materials + UBound(strItem) + 1)
        For i = 1 To UBound(strItem)
            LoadMaterialOther strItem(i), meSoulGem
        Next
    End If
    If SplitFile("Misc.txt", "MiscName: ", strItem) Then
        ReDim Preserve db.Material(1 To db.Materials + UBound(strItem) + 1)
        For i = 1 To UBound(strItem)
            LoadMaterialOther strItem(i), meMisc
        Next
    End If
    If UBound(db.Material) <> db.Materials Then ReDim Preserve db.Material(1 To db.Materials)
    SortMaterials 1, db.Materials
End Sub

Private Function SplitFile(pstrFileName As String, pstrDelimiter As String, pstrArray() As String) As Boolean
    Dim strFile As String
    Dim strRaw As String
    
    Erase pstrArray
    strFile = DataPath() & pstrFileName
    If Not xp.File.Exists(strFile) Then Exit Function
    strRaw = xp.File.LoadToString(strFile)
    If InStr(strRaw, pstrDelimiter) = 0 Then Exit Function
    pstrArray = Split(strRaw, pstrDelimiter)
    SplitFile = True
End Function

Private Sub LoadMaterialCollectable(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As MaterialType
    Dim lngPos As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .MatType = meCollectable
        .Material = Trim$(strLine(0))
        .Plural = DefaultPlural(.Material)
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "school": .School = GetSchoolID(strItem)
                    Case "tier": .Tier = lngValue
                    Case "frequency": .Frequency = GetFrequencyValue(strItem)
                    Case "plural": .Plural = strItem
                    Case "flags": If strItem = "Eberron" Then .Eberron = True
                    Case "value": .Override = lngValue
                    Case Else: Debug.Print "Invalid field in Collectables.txt" & vbNewLine & .Material & vbNewLine & strLine(lngLine) & vbNewLine
                End Select
            End If
        Next
    End With
    With db
        .Materials = .Materials + 1
        .Material(.Materials) = typNew
    End With
    ' Update pool counters
    With typNew
        With db.School(.School).Tier(.Tier).Freq(.Frequency)
            .World(weEberron).Pool = .World(weEberron).Pool + 1
            If Not typNew.Eberron Then .World(weRealms).Pool = .World(weRealms).Pool + 1
        End With
    End With
End Sub

Private Sub LoadMaterialOther(pstrRaw As String, penMatType As MaterialEnum)
    Dim strMaterial As String
    Dim strNotes As String
    Dim lngPos As Long
    
    CleanText pstrRaw
    lngPos = InStr(pstrRaw, vbNewLine)
    If lngPos = 0 Then
        strMaterial = pstrRaw
    Else
        strMaterial = Left$(pstrRaw, lngPos - 1)
        strNotes = Mid$(pstrRaw, lngPos + 2)
    End If
    If Len(strMaterial) Then
        db.Materials = db.Materials + 1
        With db.Material(db.Materials)
            .MatType = penMatType
            If penMatType = meSoulGem Then .Material = "Soul Gem: " & strMaterial Else .Material = strMaterial
            If .Material = "Adamantine Ore" Then .Plural = .Material Else .Plural = DefaultPlural(.Material)
            .Notes = strNotes
        End With
    End If
End Sub

Private Function DefaultPlural(pstrSingular As String) As String
    Dim lngPos As Long
    
    If Left$(pstrSingular, 10) = "Soul Gem: " Then
        DefaultPlural = "Soul Gems: " & Mid$(pstrSingular, 11)
    ElseIf Left$(pstrSingular, 6) = "Tome: " Then
        DefaultPlural = "Tomes: " & Mid$(pstrSingular, 7)
    Else
        lngPos = InStr(pstrSingular, " of ")
        If lngPos Then
            DefaultPlural = Left$(pstrSingular, lngPos - 1) & "s" & Mid$(pstrSingular, lngPos)
        Else
            DefaultPlural = pstrSingular & "s"
        End If
    End If
End Function

Public Function Pluralized(ptypIngredient As IngredientType) As String
    Dim lngMaterial As Long
    
    With ptypIngredient
        If .Count = 1 Then
            Pluralized = .Material
        Else
            lngMaterial = SeekMaterial(.Material)
            If lngMaterial Then
                Pluralized = db.Material(lngMaterial).Plural
            Else
                Pluralized = DefaultPlural(.Material)
            End If
        End If
    End With
End Function

' Quicksort
' Omit plngLeft & plngRight; they are used internally during recursion
Public Sub SortMaterials(Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim strMid As Variant
    Dim typSwap As MaterialType
    
    lngFirst = plngLeft
    lngLast = plngRight
    strMid = SortTerm((plngLeft + plngRight) \ 2)
    Do
        Do While SortTerm(lngFirst) < strMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While strMid < SortTerm(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            typSwap = db.Material(lngFirst)
            db.Material(lngFirst) = db.Material(lngLast)
            db.Material(lngLast) = typSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then SortMaterials plngLeft, lngLast
    If lngFirst < plngRight Then SortMaterials lngFirst, plngRight
End Sub

Private Function SortTerm(plngIndex As Long) As String
    SortTerm = SearchTerm(db.Material(plngIndex).Material)
End Function

' Ignore the leading quote in 'Wavecrasher' Cargo Manifest
Private Function SearchTerm(pstrSearch As String) As String
    If Left$(pstrSearch, 1) = "'" Then SearchTerm = Mid$(pstrSearch, 2) Else SearchTerm = pstrSearch
End Function

' Simple binary search
Public Function SeekMaterial(ByVal pstrSearch As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    Dim strCompare As String
    
    pstrSearch = SearchTerm(pstrSearch)
    lngFirst = 1
    lngLast = db.Materials
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If SortTerm(lngMid) < pstrSearch Then
            lngFirst = lngMid + 1
        ElseIf SortTerm(lngMid) > pstrSearch Then
            lngLast = lngMid - 1
        Else
            SeekMaterial = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* FARMS *************


Private Sub LoadFarms()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    
    db.Farms = 0
    Erase db.Farm
    strFile = DataPath() & "Farms.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "Farm: ")
    If UBound(strItem) < 1 Then Exit Sub
    ReDim db.Farm(1 To UBound(strItem))
    For i = 1 To UBound(strItem)
        If InStr(strItem(i), "Notes: ") Then LoadFarm strItem(i)
    Next
    With db
         If UBound(.Farm) <> .Farms Then ReDim Preserve .Farm(1 To .Farms)
    End With
    SortFarms
End Sub

Private Sub LoadFarm(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As FarmType
    Dim lngPos As Long
    Dim enSchool As SchoolEnum
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        enSchool = seUnknown
        .Farm = Trim$(strLine(0))
        .Wiki = MakeWiki(.Farm)
        ' Process lines
        For lngLine = 0 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "any": .Any = lngValue
                    Case "arcane": .Arcane = lngValue
                    Case "lore": .Lore = lngValue
                    Case "natural": .Natural = lngValue
                    Case "wiki": .Wiki = MakeWiki(strItem)
                    Case "flags"
                        Select Case strItem
                            Case "Realms": .Realms = True
                            Case "Treasure Bag": .TreasureBag = True
                            Case Else: Debug.Print "Invalid field in Farms.txt" & vbNewLine & .Farm & vbNewLine & strLine(lngLine) & vbNewLine
                        End Select
                    Case "need": .Need = strItem
                    Case "fight": .Fight = strItem
                    Case "notes": If Len(.Notes) = 0 Then .Notes = strItem Else .Notes = .Notes & vbNewLine & strItem
                    Case "time": .Seconds = TimeToSeconds(strItem)
                    Case "video": .Video = strItem
                    Case Else: Debug.Print "Invalid field in Farms.txt" & vbNewLine & .Farm & vbNewLine & strLine(lngLine) & vbNewLine
                End Select
            End If
        Next
    End With
    With db
        .Farms = .Farms + 1
        .Farm(.Farms) = typNew
    End With
End Sub

' Gnome sort, because list is assumed to already be sorted, or maybe 1 or 2 out of place
' Gnome sort is optimally efficient for this scenario, slightly ahead of Insertion sort
Public Sub SortFarms()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As FarmType
    
    With db
        iMin = 2
        iMax = .Farms
        i = iMin
        j = i + 1
        Do While i <= iMax
            If .Farm(i).Farm < .Farm(i - 1).Farm Then
                typSwap = .Farm(i)
                .Farm(i) = .Farm(i - 1)
                .Farm(i - 1) = typSwap
                If i > iMin Then i = i - 1
            Else
                i = j
                j = j + 1
            End If
        Loop
    End With
End Sub

' Simple binary search
Public Function SeekFarm(pstrFarm As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Farms
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Farm(lngMid).Farm < pstrFarm Then
            lngFirst = lngMid + 1
        ElseIf db.Farm(lngMid).Farm > pstrFarm Then
            lngLast = lngMid - 1
        Else
            SeekFarm = lngMid
            Exit Function
        End If
    Loop
End Function

' Limited functionality; only supports 0:00 to 9:59 because I'm lazy
Private Function TimeToSeconds(pstrTime As String) As Long
    TimeToSeconds = (val(Left$(pstrTime, 1)) * 60) + val(Right$(pstrTime, 2))
End Function

Public Function SecondsToTime(plngSeconds As Long) As String
    SecondsToTime = plngSeconds \ 60 & ":" & Format(plngSeconds Mod 60, "00")
End Function


' ************* SCHOOLS *************


Private Sub LoadSchools()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    Dim enSchool As SchoolEnum
    Dim lngTier As Long
    Dim strFarm As String
    Dim strDifficulty As String
    
    Erase db.School
    strFile = DataPath() & "Schools.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, vbNewLine)
    If UBound(strItem) < 1 Then Exit Sub
    For i = 0 To UBound(strItem)
        If ParseSchool(strItem(i), enSchool, lngTier, strFarm, strDifficulty) Then
            If SeekFarm(strFarm) = 0 Then
                Debug.Print "Invalid Farm in Schools.txt: " & strFarm
            Else
                With db.School(enSchool).Tier(lngTier)
                    .TierFarms = .TierFarms + 1
                    ReDim Preserve .TierFarm(1 To .TierFarms)
                    With .TierFarm(.TierFarms)
                        .Farm = strFarm
                        .Difficulty = strDifficulty
                    End With
                End With
            End If
        End If
    Next
End Sub

Private Function ParseSchool(ByVal pstrRaw As String, penSchool As SchoolEnum, plngTier As Long, pstrFarm As String, pstrDifficulty As String) As Boolean
    Dim lngPos As Long
    
    CleanText pstrRaw
    lngPos = InStr(pstrRaw, ": ")
    If lngPos = 0 Then Exit Function
    penSchool = GetSchoolID(Left$(pstrRaw, lngPos - 3))
    plngTier = val(Mid$(pstrRaw, lngPos - 1, 1))
    pstrFarm = Mid$(pstrRaw, lngPos + 2)
    lngPos = InStrRev(pstrFarm, ": ")
    If lngPos = 0 Then
        pstrDifficulty = vbNullString
    Else
        pstrDifficulty = Mid$(pstrFarm, lngPos + 2)
        pstrFarm = Left$(pstrFarm, lngPos - 1)
    End If
    If penSchool <> seUnknown And plngTier >= 1 And plngTier <= 6 Then
        ParseSchool = True
    End If
End Function


' ************* SCALING *************


Private Sub LoadScaling()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim strScale() As String
    Dim i As Long
    Dim j As Long

    Erase db.Scaling
    db.Scales = 0
    strFile = DataPath() & "Scaling.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, vbNewLine)
    If UBound(strItem) < 1 Then Exit Sub
    With db
        ReDim .Scaling(1 To UBound(strItem))
        For i = 1 To UBound(strItem)
            strScale = Split(strItem(i), vbTab)
            If UBound(strScale) = 35 Then
                If Len(Trim$(strScale(1))) <> 0 Then
                    .Scales = .Scales + 1
                    .Scaling(.Scales).Order = .Scales
                    With .Scaling(.Scales)
                        .Group = strScale(0)
                        .ScaleName = strScale(1)
                        ReDim .Table(1 To MAX_CRAFTING_LEVEL)
                        For j = 1 To MAX_CRAFTING_LEVEL
                            .Table(j) = strScale(j + 1)
                        Next
                    End With
                End If
            End If
        Next
        If UBound(.Scaling) <> .Scales Then ReDim Preserve .Scaling(1 To .Scales)
    End With
    SortScaling
End Sub

' Comb sort, because list is completely shuffled to start with. Quicksort or Shellsort would be
' slightly more efficient, but comb sort is a touch more elegant for dropping a udt into the code
Private Sub SortScaling()
    Const ShrinkFactor = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As ScalingType
    Dim blnSwapped As Boolean
    
    iMin = 1
    iMax = db.Scales
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If db.Scaling(i).ScaleName > db.Scaling(i + lngGap).ScaleName Then
                typSwap = db.Scaling(i)
                db.Scaling(i) = db.Scaling(i + lngGap)
                db.Scaling(i + lngGap) = typSwap
                blnSwapped = True
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

' Simple binary search
Public Function SeekScaling(pstrScaleName As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Scales
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Scaling(lngMid).ScaleName < pstrScaleName Then
            lngFirst = lngMid + 1
        ElseIf db.Scaling(lngMid).ScaleName > pstrScaleName Then
            lngLast = lngMid - 1
        Else
            SeekScaling = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* SHARDS *************


Private Sub LoadShards()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    
    db.Shards = 0
    Erase db.Shard
    strFile = DataPath() & "Shards.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "ShardName: ")
    If UBound(strItem) < 1 Then Exit Sub
    ReDim db.Shard(1 To UBound(strItem))
    For i = 1 To UBound(strItem)
        If InStr(strItem(i), "Group: ") Then LoadShard strItem(i)
    Next
    With db
         If UBound(.Shard) <> .Shards Then ReDim Preserve .Shard(1 To .Shards)
    End With
    SortShards
End Sub

Private Sub LoadShard(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As ShardType
    Dim lngPos As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .ShardName = Trim$(strLine(0))
        .ScaleName = .ShardName
        If Left$(.ShardName, 11) = "Insightful " Then
            .Abbreviation = "Ins. " & Mid$(.ShardName, 12)
            .GridName = .Abbreviation
        Else
            .Abbreviation = .ShardName
            .GridName = .ShardName
        End If
        If Right$(.GridName, 11) = " Resistance" Then .GridName = Left$(.GridName, Len(.GridName) - 4)
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "group"
                        .Group = strItem
                        AddGroup strItem
                    Case "scaling"
                        .ScaleName = strItem
                        If SeekScaling(strItem) = 0 Then Debug.Print "Invalid scale in Shards.txt" & vbNewLine & .ShardName & vbNewLine & strLine(lngLine) & vbNewLine
                    Case "abbreviation"
                        .Abbreviation = strItem
                        If .GridName = .ShardName Then .GridName = .Abbreviation
                    Case "shortname": .ShortName = strItem
                    Case "gridname": .GridName = strItem
                    Case "ml": .ML = lngValue
                    Case "prefix": ParseSlots .Prefix, strList, lngListMax
                    Case "suffix": ParseSlots .Suffix, strList, lngListMax
                    Case "extra": ParseSlots .Extra, strList, lngListMax
                    Case "boundlevel": .Bound.Level = lngValue
                    Case "boundessences": .Bound.Essences = lngValue
                    Case "boundingredients": ParseIngredients .Bound, strList, lngListMax
                    Case "unboundlevel": .Unbound.Level = lngValue
                    Case "unboundessences": .Unbound.Essences = lngValue
                    Case "unboundingredients": ParseIngredients .Unbound, strList, lngListMax
                    Case "demand": .Demand = GetDemandValue(strItem)
                    Case "notes": .Notes = strItem
                    Case "descrip": .Descrip = strItem
                    Case "warning": .Warning = strItem
                    Case Else: Debug.Print "Invalid field in Shards.txt" & vbNewLine & .ShardName & vbNewLine & strLine(lngLine) & vbNewLine
                End Select
            End If
        Next
        If Len(.ShortName) = 0 Then .ShortName = AutoShortName(.GridName)
    End With
    With db
        .Shards = .Shards + 1
        .Shard(.Shards) = typNew
    End With
End Sub

Private Function AutoShortName(pstrGridName As String) As String
    Dim strShort As String
    
    strShort = pstrGridName
    If Left$(strShort, 5) = "Ins. " Then strShort = "i" & Mid$(strShort, 6)
    If Left$(strShort, 10) = "Efficient " Then strShort = "Eff " & Mid$(strShort, 11)
    If Right$(strShort, 7) = " Effect" Then strShort = Left$(strShort, Len(strShort) - 7)
    If Right$(strShort, 11) = " Absorption" Then strShort = Left$(strShort, Len(strShort) - 11) & " Absorb"
    AutoShortName = strShort
End Function

' Do a binary search to find the group, if not found, insert in proper sorted position
Private Sub AddGroup(pstrGroup As String)
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    Dim i As Long
    
    lngFirst = 1
    lngLast = db.Groups
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Group(lngMid) > pstrGroup Then
            lngLast = lngMid - 1
        ElseIf db.Group(lngMid) < pstrGroup Then
            lngFirst = lngMid + 1
        Else
            Exit Sub
        End If
    Loop
    With db
        .Groups = .Groups + 1
        ReDim Preserve .Group(1 To .Groups)
        For i = .Groups To lngFirst + 1 Step -1
            .Group(i) = .Group(i - 1)
        Next
        .Group(lngFirst) = pstrGroup
    End With
End Sub

Private Sub ParseSlots(pblnSlot() As Boolean, pstrList() As String, plngMax As Long)
    Dim lngIndex As Long
    Dim i As Long
    
    For i = 0 To plngMax
        lngIndex = GetGearID(pstrList(i))
        If lngIndex = -1 Then
            Debug.Print "Gear slot not found: " & pstrList(i) & vbNewLine
        Else
            pblnSlot(lngIndex) = True
        End If
    Next
End Sub

Private Sub ParseIngredients(ptypRecipe As RecipeType, pstrList() As String, plngMax As Long)
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strMaterial As String
    Dim i As Long
    
    With ptypRecipe
        .Ingredients = plngMax + 1
        ReDim .Ingredient(1 To .Ingredients)
        For i = 1 To .Ingredients
            lngPos = InStr(pstrList(i - 1), " ")
            If lngPos = 0 Then
                Debug.Print "Error parsing: " & pstrList(i - 1) & vbNewLine
            Else
                lngCount = val(Left$(pstrList(i - 1), lngPos - 1))
                strMaterial = Mid$(pstrList(i - 1), lngPos + 1)
                If lngCount = 0 Or Len(strMaterial) = 0 Then
                    Debug.Print "Error parsing: " & pstrList(i - 1) & vbNewLine
                Else
                    .Ingredient(i).Count = lngCount
                    .Ingredient(i).Material = strMaterial
                    If SeekMaterial(strMaterial) = 0 Then Debug.Print "Collectable not found: " & strMaterial & vbNewLine
                End If
            End If
        Next
    End With
End Sub

' Gnome sort, because list is assumed to already be sorted, or maybe 1 or 2 out of place
' Gnome sort is optimally efficient for this scenario, slightly ahead of Insertion sort
Public Sub SortShards()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim blnSorted As Boolean
    Dim typSwap As ShardType
    
    blnSorted = True
    With db
        iMin = 2
        iMax = .Shards
        i = iMin
        j = i + 1
        Do While i <= iMax
            If .Shard(i).ShardName < .Shard(i - 1).ShardName Then
                typSwap = .Shard(i)
                .Shard(i) = .Shard(i - 1)
                .Shard(i - 1) = typSwap
                blnSorted = False
                If i > iMin Then i = i - 1
            Else
                i = j
                j = j + 1
            End If
        Loop
    End With
    If Not blnSorted Then Debug.Print "Shards.txt isn't sorted"
End Sub

' Simple binary search
Public Function SeekShard(pstrShardName As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Shards
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Shard(lngMid).ShardName < pstrShardName Then
            lngFirst = lngMid + 1
        ElseIf db.Shard(lngMid).ShardName > pstrShardName Then
            lngLast = lngMid - 1
        Else
            SeekShard = lngMid
            Exit Function
        End If
    Loop
End Function

' Create index for Group+ShardName
' (Otherwise, sorted listbox would put "Spellpower: Whatever" before "Spell: Whatever")
Private Sub IndexShards()
    Dim i As Long
    
    ReDim db.ShardIndex(1 To db.Shards)
    For i = 1 To db.Shards
        db.ShardIndex(i) = i
    Next
    CombsortShardIndex
End Sub

' Combsort, because this was faster to implement than quicksort, which on first pass was buggy
Public Sub CombsortShardIndex()
    Const ShrinkFactor = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim lngSwap As Long
    Dim blnSwapped As Boolean
    
    iMin = 1
    iMax = db.Shards
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If CompareIndex(i, i + lngGap) = 1 Then
                lngSwap = db.ShardIndex(i)
                db.ShardIndex(i) = db.ShardIndex(i + lngGap)
                db.ShardIndex(i + lngGap) = lngSwap
                blnSwapped = True
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

Private Function CompareIndex(plngLeft As Long, plngRight As Long) As Long
    Dim lngLeft As Long
    Dim lngRight As Long
    
    lngLeft = db.ShardIndex(plngLeft)
    lngRight = db.ShardIndex(plngRight)
    If db.Shard(lngLeft).Group < db.Shard(lngRight).Group Then
        CompareIndex = -1
    ElseIf db.Shard(lngLeft).Group > db.Shard(lngRight).Group Then
        CompareIndex = 1
    ElseIf db.Shard(lngLeft).ShardName < db.Shard(lngRight).ShardName Then
        CompareIndex = -1
    ElseIf db.Shard(lngLeft).ShardName > db.Shard(lngRight).ShardName Then
        CompareIndex = 1
    End If
End Function


' ************* RECIPES *************


Public Sub AggregateRecipe(ptypRecipe As RecipeType, ptypAdd As RecipeType)
    Dim r As Long
    Dim a As Long
    
    With ptypRecipe
        .Essences = .Essences + ptypAdd.Essences
        If .Level < ptypAdd.Level Then .Level = ptypAdd.Level
        For a = 1 To ptypAdd.Ingredients
            For r = 1 To .Ingredients
                If .Ingredient(r).Material = ptypAdd.Ingredient(a).Material Then
                    .Ingredient(r).Count = .Ingredient(r).Count + ptypAdd.Ingredient(a).Count
                    Exit For
                End If
            Next
            If r > .Ingredients Then
                .Ingredients = .Ingredients + 1
                ReDim Preserve .Ingredient(1 To .Ingredients)
                .Ingredient(.Ingredients) = ptypAdd.Ingredient(a)
            End If
        Next
    End With
End Sub

Public Sub AggregateRecipeML(ptypRecipe As RecipeType, plngML As Long, pblnBound As Boolean)
    Dim lngLevel As Long
    Dim lngEssences As Long
    
    lngLevel = MLShardLevel(plngML, pblnBound)
    lngEssences = MLShardEssences(plngML, pblnBound)
    With ptypRecipe
        If .Level < lngLevel Then .Level = lngLevel
        .Essences = .Essences + lngEssences
    End With
End Sub

' Insertion sort (want stable algorithm to preserve alphabetical order within same # of ingedients)
Public Sub AggregateRecipeSort(ptypRecipe As RecipeType)
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typHold As IngredientType
    
    GetFrequencies ptypRecipe
    With ptypRecipe
        iMin = 2
        iMax = .Ingredients
        For i = iMin To iMax
            typHold = .Ingredient(i)
            For j = i To iMin Step -1
                If AggregateSortCompare(typHold, .Ingredient(j - 1)) Then .Ingredient(j) = .Ingredient(j - 1) Else Exit For
'                If typHold.Count > .Ingredient(j - 1).Count Then .Ingredient(j) = .Ingredient(j - 1) Else Exit For
            Next j
            .Ingredient(j) = typHold
        Next i
        ' Now move purifieds to the end
        For i = 1 To iMax
            If .Ingredient(i).Material = "Purified Eberron Dragonshard" Then Exit For
        Next
        If i < iMax Then
            typHold = .Ingredient(i)
            For j = i To iMax - 1
                .Ingredient(j) = .Ingredient(j + 1)
            Next
            .Ingredient(iMax) = typHold
        End If
    End With
End Sub

Private Sub GetFrequencies(ptypRecipe As RecipeType)
    Dim lngMaterial As Long
    Dim i As Long
    
    With ptypRecipe
        For i = 1 To .Ingredients
            With .Ingredient(i)
                lngMaterial = SeekMaterial(.Material)
                If lngMaterial Then
                    Select Case db.Material(lngMaterial).MatType
                        Case meCollectable: .Frequency = db.Material(lngMaterial).Frequency
                        Case meSoulGem: .Frequency = 10
                        Case meMisc: .Frequency = 11
                    End Select
                End If
            End With
        Next
    End With
End Sub

Private Function AggregateSortCompare(ptypHold As IngredientType, ptypIngredient As IngredientType) As Boolean
    If ptypHold.Frequency < ptypIngredient.Frequency Then
        AggregateSortCompare = True
    ElseIf ptypHold.Frequency = ptypIngredient.Frequency Then
        If SearchTerm(ptypHold.Material) < SearchTerm(ptypIngredient.Material) Then AggregateSortCompare = True
    End If
End Function

Public Function MLShardEssences(plngML As Long, pblnBound As Boolean) As Long
    If pblnBound Then
        MLShardEssences = plngML * 10
    Else
        If plngML = 1 Then MLShardEssences = 100 Else MLShardEssences = 2 * (plngML * 10 + 50)
    End If
End Function

Public Function MLShardLevel(plngML As Long, pblnBound As Boolean) As Long
    If pblnBound Then
        If plngML = 1 Then MLShardLevel = 1 Else MLShardLevel = plngML * 10
    Else
        If plngML = 1 Then MLShardLevel = 50 Else MLShardLevel = plngML * 10 + 50
    End If
End Function

Public Sub AddRecipeToInfo(ptypRecipe As RecipeType, pinfo As userInfo, Optional pstrIndent As String)
    Dim strDisplay As String
    Dim lngPadding As Long
    Dim i As Long
    
    With ptypRecipe
        lngPadding = 1
        For i = 1 To .Ingredients
            Select Case .Ingredient(i).Count
                Case Is > 999: If lngPadding < 4 And .Ingredient(i).Material <> "Purified Eberron Dragonshard" Then lngPadding = 4
                Case Is > 99: If lngPadding < 3 And .Ingredient(i).Material <> "Purified Eberron Dragonshard" Then lngPadding = 3
                Case Is > 9: If lngPadding < 2 Then lngPadding = 2
            End Select
        Next
        If .Essences > 0 Then pinfo.AddText FormatEssences(.Essences) & " Essences", 0, pstrIndent
        For i = 1 To .Ingredients
            strDisplay = Pluralized(.Ingredient(i))
            With .Ingredient(i)
                pinfo.AddText vbNullString
                pinfo.AddNumber .Count, lngPadding, pstrIndent
                pinfo.AddLink strDisplay, lseMaterial, .Material, 0, False
            End With
        Next
    End With
End Sub

Public Function BindingText(pblnBound As Boolean) As String
    If pblnBound Then BindingText = "Bound" Else BindingText = "Unbound"
End Function

Public Function FormatEssences(plngEssences As Long) As String
    If plngEssences < 10000 Then FormatEssences = plngEssences Else FormatEssences = Format(plngEssences, "#,##0")
End Function


' ************* ITEMS *************


Private Sub LoadItems()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    
    db.Items = 0
    Erase db.Item
    strFile = DataPath() & "Items.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "Item: ")
    If UBound(strItem) < 1 Then Exit Sub
    ReDim db.Item(1 To UBound(strItem))
    For i = 1 To UBound(strItem)
        LoadItem strItem(i)
    Next
    With db
         If UBound(.Item) <> .Items Then ReDim Preserve .Item(1 To .Items)
    End With
    SortItems
End Sub

Private Sub LoadItem(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As ItemType
    Dim lngPos As Long
    Dim enSchool As SchoolEnum
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .ItemName = Trim$(strLine(0))
        .Image = .ItemName
        ' Process lines
        For lngLine = 0 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "resourceid"
                        .ResourceID = strItem
                    Case "type"
                        .ItemType = GetItemTypeID(strItem)
                    Case "style"
                        .ItemStyle = GetItemStyleID(strItem)
                        If .ItemStyle = iseMelee2H Then .TwoHand = True
                    Case "twohand"
                        .TwoHand = True
                    Case "scaling"
                        .Scales = True
                        If lngListMax = ItemScales Then .Scaling = strList Else Debug.Print "Invalid Scaling in Items.txt: " & strLine(lngLine)
                    Case "slots"
                        .SlotStyles = lngListMax + 1
                        ReDim .SlotStyle(1 To .SlotStyles)
                        For i = 1 To .SlotStyles
                            Select Case LCase$(strList(i - 1))
                                Case "red": .SlotStyle(i) = isseRed
                                Case "blue": .SlotStyle(i) = isseBlue
                                Case "green": .SlotStyle(i) = isseGreen
                                Case "yellow": .SlotStyle(i) = isseYellow
                                Case "colorless": .SlotStyle(i) = isseColorless
                                Case "dual": .SlotStyle(i) = isseDual
                                Case "none"
                                    .SlotStyles = 0
                                    Erase .SlotStyle
                                Case Else
                                    Debug.Print "Unknown slots for " & .ItemName & " in Items.txt: " & strList(i - 1)
                            End Select
                        Next
                    Case Else
                        Debug.Print "Unknown line in Items.txt: " & strLine(lngLine)
                End Select
            End If
        Next
    End With
    With db
        .Items = .Items + 1
        .Item(.Items) = typNew
    End With
End Sub

' Gnome sort, because list is assumed to already be sorted, or maybe 1 or 2 out of place
' Gnome sort is optimally efficient for this scenario, slightly ahead of Insertion sort
Public Sub SortItems()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As ItemType
    
    With db
        iMin = 2
        iMax = .Items
        i = iMin
        j = i + 1
        Do While i <= iMax
            If .Item(i).ItemName < .Item(i - 1).ItemName Then
                typSwap = .Item(i)
                .Item(i) = .Item(i - 1)
                .Item(i - 1) = typSwap
                If i > iMin Then i = i - 1
            Else
                i = j
                j = j + 1
            End If
        Loop
    End With
End Sub

' Simple binary search
Public Function SeekItem(pstrItem As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Items
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Item(lngMid).ItemName < pstrItem Then
            lngFirst = lngMid + 1
        ElseIf db.Item(lngMid).ItemName > pstrItem Then
            lngLast = lngMid - 1
        Else
            SeekItem = lngMid
            Exit Function
        End If
    Loop
End Function

Public Function GetArmorMaterial(pstrItemName As String) As ArmorMaterialEnum
    Dim lngIndex As Long

    lngIndex = SeekItem(pstrItemName)
    If lngIndex = 0 Then Exit Function
    Select Case db.Item(lngIndex).ItemStyle
        Case iseMetalArmor: GetArmorMaterial = ameMetal
        Case iseLeatherArmor: GetArmorMaterial = ameLeather
        Case iseClothArmor: GetArmorMaterial = ameCloth
        Case iseDocent: GetArmorMaterial = ameDocent
    End Select
End Function

Public Sub GetMainhandInfo(pstrItemName As String, penMainHand As MainHandEnum, pblnTwoHand As Boolean)
    Dim lngIndex As Long
    
    lngIndex = SeekItem(pstrItemName)
    If lngIndex = 0 Then Exit Sub
    Select Case db.Item(lngIndex).ItemStyle
        Case iseRange, iseThrower: penMainHand = mheRange
        Case Else: penMainHand = mheMelee
    End Select
    pblnTwoHand = db.Item(lngIndex).TwoHand
End Sub

Public Function GetOffhandStyle(pstrItemName As String) As OffHandEnum
    Dim lngIndex As Long
    
    lngIndex = SeekItem(pstrItemName)
    If lngIndex = 0 Then Exit Function
    Select Case db.Item(lngIndex).ItemStyle
        Case iseMelee1H: GetOffhandStyle = oheMelee
        Case iseShield: GetOffhandStyle = oheShield
        Case iseOrb: GetOffhandStyle = oheOrb
        Case iseRunearm: GetOffhandStyle = oheRunearm
        Case iseEmpty: GetOffhandStyle = oheEmpty
    End Select
End Function


' ************* COMBO ITEMS *************


Private Sub LoadItemCombo()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    
    strFile = DataPath() & "UserItemCombo.txt"
    If Not xp.File.Exists(strFile) Then strFile = DataPath() & "ItemCombo.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "Item: ")
    If UBound(strItem) < 1 Then Exit Sub
    For i = 1 To UBound(strItem)
        LoadCombo strItem(i)
    Next
End Sub

Private Sub LoadCombo(pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim strCombo As String
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    strCombo = Trim$(strLine(0))
    ' Process lines
    For lngLine = 0 To UBound(strLine)
        If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
            Select Case strField
                Case "parent": AddComboItem strCombo, strItem
            End Select
        End If
    Next
End Sub

Private Sub AddComboItem(pstrCombo As String, pstrParent As String)
    Dim typNew As ItemType
    Dim lngIndex As Long
    Dim i As Long
    
    lngIndex = SeekItem(pstrCombo)
    If lngIndex Then Exit Sub
    lngIndex = SeekItem(pstrParent)
    If lngIndex = 0 Then Exit Sub
    typNew = db.Item(lngIndex)
    typNew.ItemName = pstrCombo
    db.Items = db.Items + 1
    ReDim Preserve db.Item(1 To db.Items)
    For i = db.Items To 2 Step -1
        If db.Item(i - 1).ItemName < typNew.ItemName Then Exit For
        db.Item(i) = db.Item(i - 1)
    Next
    db.Item(i) = typNew
End Sub


' ************* ITEM CHOICES *************


Private Sub LoadItemChoices()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim lngPos As Long
    Dim i As Long
    
    strFile = DataPath() & "UserItemChoices.txt"
    If Not xp.File.Exists(strFile) Then strFile = DataPath() & "ItemChoices.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "Style: ")
    If UBound(strItem) < 1 Then Exit Sub
    For i = 1 To UBound(strItem)
        lngPos = InStr(strItem(i), vbNewLine)
        If lngPos Then
            Select Case LCase$(Left$(strItem(i), lngPos - 1))
                Case "melee1h": LoadChoices strItem(i), db.Melee1H
                Case "melee2h": LoadChoices strItem(i), db.Melee2H
                Case "range": LoadChoices strItem(i), db.Range
            End Select
        End If
    Next
End Sub

Private Sub LoadChoices(pstrRaw As String, ptypChoice As ChoiceType)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typBlank As ChoiceType
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    ptypChoice = typBlank
    With ptypChoice
        ReDim .List(1 To UBound(strLine) + 1)
        ' Process lines
        For lngLine = 0 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "default"
                        .Default.Choice = strItem
                    Case "choice"
                        .Count = .Count + 1
                        If lngListMax = 0 Then
                            .List(.Count).Choice = strItem
                        Else
                            .List(.Count).Choice = strList(0)
                            .List(.Count).OffhandPair = strList(1)
                        End If
                End Select
            End If
        Next
        If .Count = 0 Then
            Erase .List
        ElseIf UBound(.List) <> .Count Then
            ReDim Preserve .List(1 To .Count)
        End If
        For i = 1 To .Count
            If .Default.Choice = .List(i).Choice Then
                .Default.OffhandPair = .List(i).OffhandPair
                Exit For
            End If
        Next
        If i > .Count Then .Default = .List(1)
    End With
End Sub


' ************* RITUALS *************


Private Sub LoadRituals()
    Dim strFile As String
    Dim strRaw As String
    Dim strRitual() As String
    Dim i As Long
    
    db.Rituals = 0
    Erase db.Ritual
    strFile = DataPath() & "Eldritch.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strRitual = Split(strRaw, "Ritual: ")
    If UBound(strRitual) < 1 Then Exit Sub
    ReDim db.Ritual(1 To UBound(strRitual))
    For i = 1 To UBound(strRitual)
        LoadRitual strRitual(i)
    Next
    With db
         If UBound(.Ritual) <> .Rituals Then ReDim Preserve .Ritual(1 To .Rituals)
    End With
    SortRituals
End Sub

Private Sub LoadRitual(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As RitualType
    Dim lngPos As Long
    Dim enSchool As SchoolEnum
    Dim lngID As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .RitualName = Trim$(strLine(0))
        ' Process lines
        For lngLine = 0 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "descrip"
                        .Descrip = strItem
                    Case "type"
                        For i = 0 To lngListMax
                            lngID = GetItemTypeID(strList(i))
                            If lngID Then
                                .ItemType(0) = True
                                .ItemType(lngID) = True
                            End If
                        Next
                    Case "style"
                        For i = 0 To lngListMax
                            lngID = GetItemStyleID(strList(i))
                            If lngID Then
                                .ItemStyle(0) = True
                                .ItemStyle(lngID) = True
                            End If
                        Next
                    Case "ingredients"
                        ParseIngredients .Recipe, strList, lngListMax
                    Case Else
                        Debug.Print "Unknown line in Rituals.txt: " & strLine(lngLine)
                End Select
            End If
        Next
    End With
    With db
        .Rituals = .Rituals + 1
        .Ritual(.Rituals) = typNew
    End With
End Sub

' Gnome sort, because list is assumed to already be sorted, or maybe 1 or 2 out of place
' Gnome sort is optimally efficient for this scenario, slightly ahead of Insertion sort
Public Sub SortRituals()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As RitualType
    
    With db
        iMin = 2
        iMax = .Rituals
        i = iMin
        j = i + 1
        Do While i <= iMax
            If .Ritual(i).RitualName < .Ritual(i - 1).RitualName Then
                typSwap = .Ritual(i)
                .Ritual(i) = .Ritual(i - 1)
                .Ritual(i - 1) = typSwap
                If i > iMin Then i = i - 1
            Else
                i = j
                j = j + 1
            End If
        Loop
    End With
End Sub

' Simple binary search
Public Function SeekRitual(pstrRitual As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.Rituals
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.Ritual(lngMid).RitualName < pstrRitual Then
            lngFirst = lngMid + 1
        ElseIf db.Ritual(lngMid).RitualName > pstrRitual Then
            lngLast = lngMid - 1
        Else
            SeekRitual = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* AUGMENTS *************


Private Sub LoadAugments()
    Dim strFile As String
    Dim strRaw As String
    Dim strAugment() As String
    Dim strScale() As String
    Dim a As Long
    Dim s As Long
    
    Erase db.Augment
    db.Augments = 0
    strFile = DataPath() & "Augments.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strAugment = Split(strRaw, "Augment: ")
    db.Augments = UBound(strAugment)
    If db.Augments < 1 Then Exit Sub
    ReDim db.Augment(1 To db.Augments)
    For a = 1 To db.Augments
        strScale = Split(strAugment(a), "ML: ")
        ParseAugment db.Augment(a), strScale(0)
        With db.Augment(a)
            .Scalings = UBound(strScale)
            If .Scalings Then ReDim .Scaling(1 To .Scalings)
            For s = 1 To .Scalings
                ParseAugmentScale .Scaling(s), .Variations, strScale(s)
            Next
        End With
    Next
End Sub

Private Sub ParseAugment(ptypAugment As AugmentType, pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With ptypAugment
        .AugmentName = Trim$(strLine(0))
        .Variations = 1
        ReDim .Variation(1 To 1)
        .Variation(1) = .AugmentName
        ReDim .Wiki(1 To 1)
        .Wiki(1) = .AugmentName
        ReDim .StoreMissing(1 To 1)
        ' Process lines
        For lngLine = 0 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "color"
                        .Color = GetAugmentColorID(strItem)
                    Case "resourceid"
                        .ResourceID = strItem
                    Case "variants"
                        .Variations = lngListMax + 1
                        ReDim .Variation(1 To .Variations)
                        ReDim .Wiki(1 To .Variations)
                        For i = 1 To .Variations
                            .Variation(i) = strList(i - 1)
                            .Wiki(i) = strList(i - 1)
                        Next
                        ReDim .StoreMissing(1 To .Variations)
                    Case "descrip"
                        If Len(.Descrip) Then .Descrip = .Descrip & vbNewLine
                        .Descrip = .Descrip & strItem
                    Case "links"
                        For i = 1 To .Variations
                            .Wiki(i) = strList(i - 1)
                        Next
                    Case "storemissing"
                        For i = 1 To .Variations
                            If LCase$(strList(i - 1)) = "true" Then .StoreMissing(i) = True
                        Next
                    Case "notes"
                        If Len(.Notes) Then .Notes = .Notes & vbNewLine
                        .Notes = .Notes & strItem
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "static": .Static = True
                                Case "prefixnotvalue": .PrefixNotValue = True
                                Case "named": .Named = True
                                Case Else
                                    Debug.Print "Unknown flag for Augment " & .AugmentName & ": " & strList(i)
                            End Select
                        Next
                    Case Else
                        Debug.Print "Unknown line for Augment " & .AugmentName & ": " & strLine(lngLine)
                End Select
            End If
        Next
    End With
End Sub

Private Sub ParseAugmentScale(ptypScale As AugmentScaleType, plngVariants As Long, pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With ptypScale
        .ML = val(strLine(0))
        ReDim .Prefix(1 To plngVariants)
        ' Process lines
        For lngLine = 0 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "value": .Value = strItem
                    Case "store": .Store = lngValue
                    Case "remnants": .Remnants = lngValue
                    Case "vendors"
                        .Vendors = lngListMax + 1
                        ReDim .Vendor(1 To .Vendors)
                        For i = 1 To .Vendors
                            .Vendor(i) = GetAugmentVendorID(strList(i - 1))
                        Next
                    Case "prefix"
                        If lngListMax = 0 Then
                            For i = 1 To plngVariants
                                .Prefix(i) = strItem
                            Next
                        Else
                            For i = 1 To plngVariants
                                If i <= lngListMax + 1 Then .Prefix(i) = strList(i - 1)
                            Next
                        End If
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "remnantsbonusdays": .RemnantsBonusDays = True
                                Case Else:
                                  Debug.Print "Unknown flag for Scale " & .ML & ": " & strList(i)
                            End Select
                        Next
                    Case Else
                        Debug.Print "Unknown line for Scale " & .ML & ": " & strLine(lngLine)
                End Select
            End If
        Next
    End With
End Sub

Public Function AugmentFullName(plngAugment As Long, plngVariant As Long, plngScale As Long) As String
    Dim strPrefix As String
    Dim strVariant As String
    Dim strReturn As String
    Dim strValue As String
    
    If plngAugment = 0 Or plngVariant = 0 Or plngScale = 0 Then Exit Function
    With db.Augment(plngAugment)
        strPrefix = .Scaling(plngScale).Prefix(plngVariant)
        strVariant = .Variation(plngVariant)
        If Not .PrefixNotValue Then strValue = .Scaling(plngScale).Value
    End With
    If Len(strPrefix) Then strReturn = strPrefix & " "
    strReturn = strReturn & strVariant
    If Len(strValue) Then strReturn = strReturn & " " & strValue
    AugmentFullName = strReturn
End Function

Public Function AugmentScaledName(plngAugment As Long, plngVariant As Long, plngML As Long) As String
    Dim lngScale As Long
    
    With db.Augment(plngAugment)
        For lngScale = .Scalings To 1 Step -1
            If .Scaling(lngScale).ML <= plngML Then Exit For
        Next
    End With
    If lngScale = 0 Then lngScale = 1
    AugmentScaledName = AugmentFullName(plngAugment, plngVariant, lngScale)
End Function


' ************* AUGMENT ITEMS *************


Private Sub LoadAugmentItems()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    
    db.AugmentItems = 0
    Erase db.AugmentItem
    strFile = DataPath() & "AugmentItems.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "Augment: ")
    If UBound(strItem) < 1 Then Exit Sub
    ReDim db.AugmentItem(1 To UBound(strItem))
    For i = 1 To UBound(strItem)
        If InStr(strItem(i), "Item: ") Then LoadAugmentItem strItem(i)
    Next
    With db
         If UBound(.AugmentItem) <> .AugmentItems Then ReDim Preserve .AugmentItem(1 To .AugmentItems)
    End With
    SortAugmentItems
End Sub

Private Sub LoadAugmentItem(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As AugmentItemType
    Dim lngPos As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .Augment = Trim$(strLine(0))
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "item"
                        .Items = .Items + 1
                        ReDim Preserve .Item(1 To .Items)
                        .Item(.Items) = strItem
                    Case Else: Debug.Print "Invalid field in AugmentItems.txt" & vbNewLine & .Augment & vbNewLine & strLine(lngLine) & vbNewLine
                End Select
            End If
        Next
    End With
    With db
        .AugmentItems = .AugmentItems + 1
        .AugmentItem(.AugmentItems) = typNew
    End With
End Sub

' Gnome sort, because list is assumed to already be sorted, or maybe 1 or 2 out of place
' Gnome sort is optimally efficient for this scenario, slightly ahead of Insertion sort
Public Sub SortAugmentItems()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As AugmentItemType
    
    With db
        iMin = 2
        iMax = .AugmentItems
        i = iMin
        j = i + 1
        Do While i <= iMax
            If .AugmentItem(i).Augment < .AugmentItem(i - 1).Augment Then
                typSwap = .AugmentItem(i)
                .AugmentItem(i) = .AugmentItem(i - 1)
                .AugmentItem(i - 1) = typSwap
                If i > iMin Then i = i - 1
            Else
                i = j
                j = j + 1
            End If
        Loop
    End With
End Sub

' Simple binary search
Public Function SeekAugmentItem(pstrAugment As String) As Long
    Dim lngFirst As Long
    Dim lngMid As Long
    Dim lngLast As Long
    
    lngFirst = 1
    lngLast = db.AugmentItems
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        If db.AugmentItem(lngMid).Augment < pstrAugment Then
            lngFirst = lngMid + 1
        ElseIf db.AugmentItem(lngMid).Augment > pstrAugment Then
            lngLast = lngMid - 1
        Else
            SeekAugmentItem = lngMid
            Exit Function
        End If
    Loop
End Function


' ************* AUGMENT VENDORS *************


Private Sub LoadAugmentVendors()
    Dim strFile As String
    Dim strRaw As String
    Dim strItem() As String
    Dim i As Long
    
    db.AugmentVendors = 0
    Erase db.AugmentVendor
    strFile = DataPath() & "AugmentVendors.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strItem = Split(strRaw, "Vendor: ")
    If UBound(strItem) < 1 Then Exit Sub
    ReDim db.AugmentVendor(1 To UBound(strItem))
    For i = 1 To UBound(strItem)
        If InStr(strItem(i), "Type: ") Then LoadAugmentVendor strItem(i)
    Next
    With db
         If UBound(.AugmentVendor) <> .AugmentVendors Then ReDim Preserve .AugmentVendor(1 To .AugmentVendors)
    End With
End Sub

Private Sub LoadAugmentVendor(ByVal pstrRaw As String)
    Dim strLine() As String
    Dim lngLine As Long
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim typNew As AugmentVendorType
    Dim lngPos As Long
    Dim i As Long
    
    CleanText pstrRaw
    strLine = Split(pstrRaw, vbNewLine)
    With typNew
        .Vendor = Trim$(strLine(0))
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
                    Case "type": .Style = GetAugmentVendorID(strItem)
                    Case "color": .Color = GetAugmentColorID(strItem)
                    Case "ml": .ML = lngValue
                    Case "fast": .Fast = strItem
                    Case "cost": .Cost = strItem
                    Case "location": .Location = strItem
                    Case "flags"
                        For i = 0 To lngListMax
                            Select Case LCase$(strList(i))
                                Case "anylevel": .AnyLevel = True
                                Case Else: Debug.Print "Unknown flag for Vendor " & .Vendor & ": " & strList(i)
                            End Select
                        Next
                    Case Else: Debug.Print "Invalid field in AugmentVendors.txt" & vbNewLine & .Vendor & vbNewLine & strLine(lngLine) & vbNewLine
                End Select
            End If
        Next
    End With
    With db
        .AugmentVendors = .AugmentVendors + 1
        .AugmentVendor(.AugmentVendors) = typNew
    End With
End Sub


' ************* GENERAL *************


' Remove leading and trailing spaces and linefeeds
Public Function CleanText(pstrText As String)
    pstrText = Trim$(pstrText)
    Do While Left$(pstrText, 2) = vbNewLine
        pstrText = Trim$(Mid$(pstrText, 3))
    Loop
    Do While Right$(pstrText, 2) = vbNewLine
        pstrText = Trim$(Left$(pstrText, Len(pstrText) - 2))
    Loop
End Function

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
    If Len(pstrLine) = 0 Or Left$(pstrLine, 1) = ";" Then Exit Function
    ' Field
    lngPos = InStr(pstrLine, ":")
    If lngPos = 0 Then Exit Function
    ParseLine = True
    pstrField = LCase$(Trim$(Left$(pstrLine, lngPos - 1)))
    pstrLine = Mid$(pstrLine, lngPos + 2)
    If pstrField <> "notes" Then pstrLine = Trim$(pstrLine)
    ' Descriptions
    If Left$(pstrField, 4) = "wiki" Or pstrField = "descrip" Or pstrField = "notes" Then
        pstrItem = pstrLine
        Exit Function
    End If
    ' List
    If InStr(pstrLine, ";") Then
        pstrList = Split(pstrLine, ";")
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
            plngValue = val(strValue)
        End If
    Else
        ' If only a single value, and it's numeric, return it in Value also
        If IsNumeric(pstrItem) And InStr(pstrItem, "d") = 0 Then plngValue = val(pstrItem)
    End If
    ' Return single item in list form as well
    plngListMax = 0
    ReDim pstrList(0)
    pstrList(0) = pstrLine
End Function

Public Function GetGearName(penGear As GearEnum, Optional pblnDescrip As Boolean = False) As String
    Dim strReturn As String
    
    Select Case penGear
        Case geHelmet: strReturn = "Helmet"
        Case geGoggles: strReturn = "Goggles"
        Case geNecklace: strReturn = "Necklace"
        Case geCloak: strReturn = "Cloak"
        Case geBracers: strReturn = "Bracers"
        Case geGloves: strReturn = "Gloves"
        Case geBelt: strReturn = "Belt"
        Case geBoots: strReturn = "Boots"
        Case geRing: strReturn = "Ring"
        Case geTrinket: strReturn = "Trinket"
        Case ge2hMelee: If pblnDescrip Then strReturn = "Two-Handed Melee Weapon" Else strReturn = "2H Melee"
        Case geHandwraps: strReturn = "Handwraps"
        Case ge1hMelee: If pblnDescrip Then strReturn = "One-Handed Melee Weapon" Else strReturn = "1H Melee"
        Case geShield: strReturn = "Shield"
        Case geRange: If pblnDescrip Then strReturn = "Bow, Crossbow, or Thrower" Else strReturn = "Range"
        Case geRunearm: strReturn = "Runearm"
        Case geOrb: strReturn = "Orb"
        Case geMetalArmor: If pblnDescrip Then strReturn = "Chainmail, Breastplate, or Plate Armor" Else strReturn = "Metal Armor"
        Case geLeatherArmor: If pblnDescrip Then strReturn = "Leather or Hide Armor" Else strReturn = "Leather Armor"
        Case geClothArmor: If pblnDescrip Then strReturn = "Robe or Outfit" Else strReturn = "Cloth Armor"
        Case geDocent: strReturn = "Docent"
    End Select
    GetGearName = strReturn
End Function

Public Function GetGearID(ByVal pstrGear As String) As GearEnum
    Select Case LCase$(pstrGear)
        Case "helmet": GetGearID = geHelmet
        Case "goggles": GetGearID = geGoggles
        Case "necklace": GetGearID = geNecklace
        Case "cloak": GetGearID = geCloak
        Case "bracers": GetGearID = geBracers
        Case "gloves": GetGearID = geGloves
        Case "belt": GetGearID = geBelt
        Case "boots": GetGearID = geBoots
        Case "ring": GetGearID = geRing
        Case "trinket": GetGearID = geTrinket
        Case "2h melee": GetGearID = ge2hMelee
        Case "handwraps": GetGearID = geHandwraps
        Case "1h melee": GetGearID = ge1hMelee
        Case "shield": GetGearID = geShield
        Case "range": GetGearID = geRange
        Case "runearm": GetGearID = geRunearm
        Case "orb": GetGearID = geOrb
        Case "metal armor": GetGearID = geMetalArmor
        Case "leather armor": GetGearID = geLeatherArmor
        Case "cloth armor": GetGearID = geClothArmor
        Case "docent": GetGearID = geDocent
        Case Else: GetGearID = geUnknown
    End Select
End Function

Public Function GetSchoolID(pstrSchool As String) As SchoolEnum
    Select Case LCase$(pstrSchool)
        Case "any": GetSchoolID = seAny
        Case "arcane": GetSchoolID = seArcane
        Case "cultural": GetSchoolID = seCultural
        Case "lore": GetSchoolID = seLore
        Case "natural": GetSchoolID = seNatural
        Case Else: GetSchoolID = seUnknown
    End Select
End Function

Public Function GetSchoolName(penSchool As SchoolEnum) As String
    Select Case penSchool
        Case seAny: GetSchoolName = "Any"
        Case seArcane: GetSchoolName = "Arcane"
        Case seCultural: GetSchoolName = "Cultural"
        Case seLore: GetSchoolName = "Lore"
        Case seNatural: GetSchoolName = "Natural"
    End Select
End Function

Public Function GetGroupID(pstrGroup As String) As Long
    Dim i As Long
    
    For i = 1 To db.Groups
        If pstrGroup = db.Group(i) Then
            GetGroupID = i
            Exit For
        End If
    Next
End Function

Private Function GetGroupName(plngGroupID As Long) As String
    GetGroupName = db.Group(plngGroupID)
End Function

Public Function GetDemandValue(pstrDemand As String) As DemandEnum
    Select Case pstrDemand
        Case "Universal": GetDemandValue = deUniversal
        Case "Common": GetDemandValue = deCommon
        Case "Uncommon": GetDemandValue = deUncommon
        Case "Niche": GetDemandValue = deNiche
        Case "Obsolete": GetDemandValue = deObsolete
        Case "Rarely Used": GetDemandValue = deRarelyUsed
        Case Else: Debug.Print "Invalid Demand Value: " & pstrDemand
    End Select
End Function

Public Function GetDemandName(penDemand As DemandEnum) As String
    Select Case penDemand
        Case deUniversal: GetDemandName = "Universal"
        Case deCommon: GetDemandName = "Common"
        Case deUncommon: GetDemandName = "Uncommon"
        Case deNiche: GetDemandName = "Niche"
        Case deObsolete: GetDemandName = "Obsolete"
        Case deRarelyUsed: GetDemandName = "Rarely Used"
    End Select
End Function

Private Function GetFrequencyValue(pstrFrequency As String) As FrequencyEnum
    Select Case pstrFrequency
        Case "Common": GetFrequencyValue = feCommon
        Case "Uncommon": GetFrequencyValue = feUncommon
        Case "Rare": GetFrequencyValue = feRare
        Case Else: Debug.Print "Invalid Frequency Value: " & pstrFrequency
    End Select
End Function

Public Function GetFrequencyText(plngFreq As Long) As String
    Select Case plngFreq
        Case 1: GetFrequencyText = "Common"
        Case 2: GetFrequencyText = "Uncommon"
        Case 3: GetFrequencyText = "Rare"
    End Select
End Function

Public Function GetItemTypeID(pstrType As String) As ItemTypeEnum
    Select Case pstrType
        Case "Weapon": GetItemTypeID = iteWeapon
        Case "Armor": GetItemTypeID = iteArmor
        Case "Shield": GetItemTypeID = iteShield
        Case "Orb": GetItemTypeID = iteOrb
        Case "Runearm": GetItemTypeID = iteRunearm
        Case "Accessory": GetItemTypeID = iteAccessory
        Case "Empty": GetItemTypeID = iteEmpty
        Case Else: Debug.Print "ItemType not found: " & pstrType
    End Select
End Function

Public Function GetItemTypeName(penItemType As ItemTypeEnum) As String
    Select Case penItemType
        Case iteWeapon: GetItemTypeName = "Weapon"
        Case iteArmor: GetItemTypeName = "Armor"
        Case iteShield: GetItemTypeName = "Shield"
        Case iteOrb: GetItemTypeName = "Orb"
        Case iteRunearm: GetItemTypeName = "Runearm"
        Case iteAccessory: GetItemTypeName = "Accessory"
        Case iteEmpty: GetItemTypeName = "Empty"
        Case Else: GetItemTypeName = "Unknown"
    End Select
End Function

Public Function GetItemStyleID(pstrStyle As String) As ItemStyleEnum
    Select Case pstrStyle
        Case "Melee2H": GetItemStyleID = iseMelee2H
        Case "Melee1H": GetItemStyleID = iseMelee1H
        Case "Range": GetItemStyleID = iseRange
        Case "Thrower": GetItemStyleID = iseThrower
        Case "Metal Armor": GetItemStyleID = iseMetalArmor
        Case "Leather Armor": GetItemStyleID = iseLeatherArmor
        Case "Cloth Armor": GetItemStyleID = iseClothArmor
        Case "Docent": GetItemStyleID = iseDocent
        Case "Shield": GetItemStyleID = iseShield
        Case "Orb": GetItemStyleID = iseOrb
        Case "Runearm": GetItemStyleID = iseRunearm
        Case "Clothing": GetItemStyleID = iseClothing
        Case "Jewelry": GetItemStyleID = iseJewelry
        Case "Empty": GetItemStyleID = iseEmpty
        Case Else: Debug.Print "ItemStyle not found: " & pstrStyle
    End Select
End Function

Public Function GetItemStyleName(penStyle As ItemStyleEnum) As String
    Select Case penStyle
        Case iseMelee2H: GetItemStyleName = "Melee2H"
        Case iseMelee1H: GetItemStyleName = "Melee1H"
        Case iseRange: GetItemStyleName = "Range"
        Case iseThrower: GetItemStyleName = "Thrower"
        Case iseMetalArmor: GetItemStyleName = "Metal Armor"
        Case iseLeatherArmor: GetItemStyleName = "Leather Armor"
        Case iseClothArmor: GetItemStyleName = "Cloth Armor"
        Case iseDocent: GetItemStyleName = "Docent"
        Case iseShield: GetItemStyleName = "Shield"
        Case iseOrb: GetItemStyleName = "Orb"
        Case iseRunearm: GetItemStyleName = "Runearm"
        Case iseClothing: GetItemStyleName = "Belt, Boots, Cloak, Gloves, Helmet"
        Case iseJewelry: GetItemStyleName = "Bracers, Goggles, Necklace, Ring, Trinket"
        Case iseEmpty: GetItemStyleName = "Empty"
        Case Else: GetItemStyleName = "Unknown"
    End Select
End Function

Public Function GetAugmentColorName(penColor As AugmentColorEnum, Optional pblnAbbreviate As Boolean = False) As String
    Dim strReturn As String
    
    Select Case penColor
        Case aceRed: strReturn = "Red"
        Case aceOrange: strReturn = "Orange"
        Case acePurple: strReturn = "Purple"
        Case aceBlue: strReturn = "Blue"
        Case aceGreen: strReturn = "Green"
        Case aceYellow: strReturn = "Yellow"
        Case aceColorless: If pblnAbbreviate Then strReturn = "Clear" Else strReturn = "Colorless"
    End Select
    GetAugmentColorName = strReturn
End Function

Public Function GetAugmentColorID(pstrColor As String) As AugmentColorEnum
    Dim enReturn As AugmentColorEnum
    
    Select Case LCase$(pstrColor)
        Case "clear", "colorless": enReturn = aceColorless
        Case "yellow": enReturn = aceYellow
        Case "blue": enReturn = aceBlue
        Case "red": enReturn = aceRed
        Case "orange": enReturn = aceOrange
        Case "purple": enReturn = acePurple
        Case "green": enReturn = aceGreen
    End Select
    GetAugmentColorID = enReturn
End Function

Public Function GetAugmentColorValue(penColor As AugmentColorEnum) As ColorValueEnum
    Dim enReturn As ColorValueEnum
    
    Select Case penColor
        Case aceRed: enReturn = cveRed
        Case aceOrange: enReturn = cveOrange
        Case acePurple: enReturn = cvePurple
        Case aceBlue: enReturn = cveBlue
        Case aceGreen: enReturn = cveGreen
        Case aceYellow: enReturn = cveYellow
        Case aceColorless: enReturn = cveLightGray
    End Select
    GetAugmentColorValue = enReturn
End Function

Public Function GetAugmentVendorID(pstrAugmentVendor As String) As AugmentVendorEnum
    Dim enReturn As AugmentVendorEnum
    
    Select Case pstrAugmentVendor
        Case "Collector": enReturn = aveCollector
        Case "Gianthold": enReturn = aveGianthold
        Case "Lahar": enReturn = aveLahar
        Case "Lahar5Greater": enReturn = aveLahar5
        Case Else: Debug.Print "AugmentVendor not recognized: " & pstrAugmentVendor
    End Select
    GetAugmentVendorID = enReturn
End Function

Public Function GetAugmentVendorName(penAugmentVendor As AugmentVendorEnum) As String
    Dim strReturn As String
    
    Select Case penAugmentVendor
        Case aveCollector: strReturn = "Collector"
        Case aveGianthold: strReturn = "Gianthold"
        Case aveLahar: strReturn = "Lahar"
        Case aveLahar5: strReturn = "Lahar5Greater"
    End Select
    GetAugmentVendorName = strReturn
End Function
