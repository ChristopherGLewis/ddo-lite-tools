Attribute VB_Name = "basValue"
Option Explicit

Public Type MaterialFarmType
    Farm As FarmType
    Difficulty As String
    Dispensers As Double
    Rate As Long ' Seconds to farm 1
End Type

Private Type DemandType
    Universal As String
    Common As String
    Uncommon As String
    Niche As String
    Obsolete As String
    RarelyUsed As String
End Type


' ************* ENTRYPOINT *************


Public Sub CalculateValues()
    Dim i As Long
    
    CalculateSupply
    CalculateDemand
    For i = 1 To db.Materials
        With db.Material(i)
            If .MatType = meCollectable And .Frequency <> feRare Then
                If .Override Then .Value = .Override Else .Value = Nearest(.Supply * .Demand * db.EssenceRate, 25)
            End If
        End With
    Next
End Sub

Private Function Nearest(pdblValue As Double, plngIncrement As Long) As Long
    Nearest = ((Int(pdblValue + 0.5) + (plngIncrement \ 2)) \ plngIncrement) * plngIncrement
End Function


' ************* SUPPLY *************


Private Sub CalculateSupply()
    Dim lngSchool As Long
    Dim lngTier As Long
    Dim lngFreq As Long
    Dim lngMat As Long
    Dim lngSeconds As Long
    
    For lngSchool = 1 To 4
        For lngTier = 1 To 6
            FindFastest lngSchool, db.School(lngSchool).Tier(lngTier)
        Next
    Next
    For lngMat = 1 To db.Materials
        With db.Material(lngMat)
            If .MatType = meCollectable And .Frequency <> feRare Then
                lngSeconds = db.School(.School).Tier(.Tier).Freq(.Frequency).World(weEberron).Fastest
                If Not .Eberron Then
                    With db.School(.School).Tier(.Tier).Freq(.Frequency).World(weRealms)
                        If lngSeconds = 0 Or (lngSeconds > .Fastest And .Fastest <> 0) Then lngSeconds = .Fastest
                    End With
                End If
                .Supply = lngSeconds
            End If
        End With
    Next
End Sub

Private Sub FindFastest(plngSchool As Long, ptypTier As TierType)
    Dim lngFarm As Long
    Dim lngSeconds As Long
    Dim dblPulls As Double
    Dim lngFreq As Long
    Dim enWorld As WorldEnum
    Dim i As Long
    
    For i = 1 To ptypTier.TierFarms
        lngFarm = SeekFarm(ptypTier.TierFarm(i).Farm)
        If lngFarm Then
            dblPulls = CountPulls(db.Farm(lngFarm), plngSchool)
            If dblPulls <> 0 Then
                If db.Farm(lngFarm).Realms Then enWorld = weRealms Else enWorld = weEberron
                For lngFreq = 1 To 2
                    With ptypTier.Freq(lngFreq).World(enWorld)
                        lngSeconds = Int(db.Farm(lngFarm).Seconds / (dblPulls * db.Frequency(lngFreq) / .Pool) + 0.5)
                        If .Fastest = 0 Or .Fastest > lngSeconds Then .Fastest = lngSeconds
                    End With
                Next
            End If
        End If
    Next
End Sub

Private Function CountPulls(ptypFarm As FarmType, penSchool As SchoolEnum) As Double
    Dim dblPulls As Double
    
    Select Case penSchool
        Case seArcane: dblPulls = ptypFarm.Arcane
        Case seLore: dblPulls = ptypFarm.Lore
        Case seNatural: dblPulls = ptypFarm.Natural
    End Select
    CountPulls = dblPulls + (ptypFarm.Any * db.Backpack(penSchool))
End Function


' ************* DEMAND *************


Public Sub CalculateDemand()
    CountDemandValues
    Select Case db.DemandStyle
        Case dseRaw: CalculateDemandRaw
        Case dseTop: CalculateDemandTop
        Case dseWeighted: CalculateDemandWeighted
    End Select
End Sub

Private Sub CountDemandValues()
    Dim lngShard As Long
    Dim lngMat As Long
    Dim i As Long
    
    For i = 1 To db.Materials
        With db.Material(i)
            .Demand = 0
            Erase .DemandMatrix
        End With
    Next
    For lngShard = 1 To db.Shards
        With db.Shard(lngShard)
            For i = 1 To .Bound.Ingredients
                lngMat = SeekMaterial(.Bound.Ingredient(i).Material)
                If lngMat Then
                    If db.Material(lngMat).MatType = meCollectable Then
                        db.Material(lngMat).DemandMatrix(.Demand) = db.Material(lngMat).DemandMatrix(.Demand) + 1
                    End If
                End If
            Next
        End With
    Next
End Sub

Private Sub CalculateDemandRaw()
    Dim lngMaterial As Long
    Dim i As Long
    
    For lngMaterial = 1 To db.Materials
        With db.Material(lngMaterial)
            If .MatType = meCollectable Then
                .Demand = 0
                For i = 1 To 6
                    .Demand = .Demand + (db.DemandValue(i) * .DemandMatrix(i))
                Next
            End If
        End With
    Next
End Sub

Private Sub CalculateDemandTop()
    Dim lngMaterial As Long
    Dim lngApply As Long
    Dim lngRemain As Long
    Dim i As Long
    
    For lngMaterial = 1 To db.Materials
        With db.Material(lngMaterial)
            If .MatType = meCollectable Then
                .Demand = 0
                lngRemain = db.DemandTop
                For i = 1 To 6
                    lngApply = .DemandMatrix(i)
                    If lngApply > lngRemain Then lngApply = lngRemain
                    .Demand = .Demand + (db.DemandValue(i) * lngApply)
                    lngRemain = lngRemain - lngApply
                    If lngRemain = 0 Then Exit For
                Next
            End If
        End With
    Next
End Sub

Private Sub CalculateDemandWeighted()
    Dim lngMaterial As Long
    Dim lngApply As Long
    Dim lngRemain As Long
    Dim i As Long
    
    For lngMaterial = 1 To db.Materials
        With db.Material(lngMaterial)
            If .MatType = meCollectable Then
                .Demand = 0
                For i = 1 To 6
                    lngApply = .DemandMatrix(i)
                    If lngApply > db.DemandWeight(i) Then lngApply = db.DemandWeight(i)
                    .Demand = .Demand + (db.DemandValue(i) * lngApply)
                Next
            End If
        End With
    Next
End Sub


' ************* FARMS *************


Public Sub GetFarms(ptypMaterial As MaterialType, ptypFarm() As MaterialFarmType, plngFarms As Long)
    Dim enSchool As SchoolEnum
    Dim i As Long
    
    plngFarms = 0
    Erase ptypFarm
    With ptypMaterial
        If .School = seCultural Then enSchool = seAny Else enSchool = .School
        enSchool = .School
        With db.School(enSchool).Tier(.Tier)
            For i = 1 To .TierFarms
                AddFarm .TierFarm(i), ptypMaterial, .Freq(ptypMaterial.Frequency), ptypFarm, plngFarms
            Next
        End With
    End With
End Sub

Private Sub AddFarm(ptypTierFarm As TierFarmType, ptypMaterial As MaterialType, ptypPool As PoolType, ptypFarm() As MaterialFarmType, plngFarms As Long)
    Dim lngFarm As Long
    Dim dblDispensers As Double
    Dim dblPool As Double
    Dim dblSeconds As Double
    Dim dblRate As Double
    Dim typNew As MaterialFarmType
    Dim i As Long
    
    lngFarm = SeekFarm(ptypTierFarm.Farm)
    If lngFarm = 0 Then Exit Sub
    ' Eberron check
    If ptypMaterial.Eberron And db.Farm(lngFarm).Realms Then Exit Sub
    ' Count dispensers
    With db.Farm(lngFarm)
        dblSeconds = .Seconds
        Select Case ptypMaterial.School
            Case seArcane: dblDispensers = .Arcane
            Case seLore: dblDispensers = .Lore
            Case seNatural: dblDispensers = .Natural
        End Select
        dblDispensers = dblDispensers + (.Any * db.Backpack(ptypMaterial.School))
        ' Get pool
        If .Realms Then dblPool = ptypPool.World(weRealms).Pool Else dblPool = ptypPool.World(weEberron).Pool
    End With
    ' Calculate seconds to farm up one collectable
    If dblSeconds <> 0 And dblDispensers <> 0 And dblPool <> 0 Then dblRate = dblSeconds / (dblDispensers * db.Frequency(ptypMaterial.Frequency) / dblPool)
    ' Load data into new tierfarm
    typNew.Farm = db.Farm(lngFarm)
    typNew.Difficulty = ptypTierFarm.Difficulty
    typNew.Dispensers = dblDispensers
    typNew.Rate = Int(dblRate + 0.5)
    ' Add this new tierfarm to the list in its proper sorted position
    plngFarms = plngFarms + 1
    ReDim Preserve ptypFarm(1 To plngFarms)
    ptypFarm(plngFarms) = typNew
    If typNew.Rate = 0 And typNew.Farm.TreasureBag = False Then Exit Sub
    For i = plngFarms To 2 Step -1
        If Not ptypFarm(i - 1).Farm.TreasureBag Then
            If ptypFarm(i - 1).Rate = 0 Or ptypFarm(i - 1).Rate > ptypFarm(i).Rate Or (ptypFarm(i - 1).Rate = ptypFarm(i).Rate And ptypFarm(i - 1).Dispensers < ptypFarm(i).Dispensers) Or (ptypFarm(i - 1).Farm.TreasureBag = False And ptypFarm(i).Farm.TreasureBag = True) Then
                typNew = ptypFarm(i - 1)
                ptypFarm(i - 1) = ptypFarm(i)
                ptypFarm(i) = typNew
            End If
        End If
    Next
End Sub


' ************* DEBUGGING *************


Public Sub VerifyDemand()
    Dim typDemand() As DemandType
    Dim lngGroup As Long
    Dim strLine() As String
    Dim strDelimiter As String
    Dim i As Long
    
    ReDim typDemand(1 To db.Groups)
    For i = 1 To db.Shards
        lngGroup = GetGroupID(db.Shard(i).Group)
        Select Case db.Shard(i).Demand
            Case deUniversal: VerifyDemandAdd typDemand(lngGroup).Universal, db.Shard(i).ShardName
            Case deCommon: VerifyDemandAdd typDemand(lngGroup).Common, db.Shard(i).ShardName
            Case deUncommon: VerifyDemandAdd typDemand(lngGroup).Uncommon, db.Shard(i).ShardName
            Case deNiche: VerifyDemandAdd typDemand(lngGroup).Niche, db.Shard(i).ShardName
            Case deObsolete: VerifyDemandAdd typDemand(lngGroup).Obsolete, db.Shard(i).ShardName
            Case deRarelyUsed: VerifyDemandAdd typDemand(lngGroup).RarelyUsed, db.Shard(i).ShardName
            Case Else: Debug.Print "Invalid Demand for " & db.Shard(i).ShardName
        End Select
    Next
    strDelimiter = ","
    ReDim strLine(db.Groups)
    strLine(0) = "Group,Universal,Common,Uncommon,Niche,Obsolete,Rarely Used"
    For i = 1 To db.Groups
        With typDemand(i)
            strLine(i) = db.Group(i) & strDelimiter & .Universal & strDelimiter & .Common & strDelimiter & .Uncommon & strDelimiter & .Niche & strDelimiter & .Obsolete & strDelimiter & .RarelyUsed
        End With
    Next
    xp.File.SaveStringAs App.Path & "\Demand.csv", Join(strLine, vbNewLine)
End Sub

Private Sub VerifyDemandAdd(pstrText As String, pstrAdd As String)
    If Len(pstrText) = 0 Then
        pstrText = Chr(34) & pstrAdd & Chr(34)
    Else
        pstrText = Left$(pstrText, Len(pstrText) - 1) & ", " & pstrAdd & Chr(34)
    End If
End Sub

Public Sub OutputValues()
    Dim strFile As String
    Dim strRaw As String
    Dim i As Long
    
    For i = 1 To db.Materials
        With db.Material(i)
            If .MatType = meCollectable And .Frequency <> feRare Then
                strRaw = strRaw & .Value & "," & .Supply & "," & .Demand & "," & .Material & vbNewLine
            End If
        End With
    Next
    strFile = App.Path & "\Values.csv"
    xp.File.SaveStringAs strFile, strRaw
End Sub
