Attribute VB_Name = "basAnalysis"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private mlngProgressMax As Long
Private mlngProgressTerms As Long
Private mlngProgressTerm() As Long

Private mlngValid As Long
Private mlngFailedOn As Long
Private mblnFinished As Boolean
Private mlngCopyMemoryLength As Long

Private mlngCount As Long


' ************* INITIALIZE *************


Public Sub InitProcessing(gs As GearsetType, anal As AnalysisType)
    Dim typBlank As AnalysisType
    Dim i As Long
    
    StopwatchStart
    mlngCount = 0
    mlngValid = 0
    mlngFailedOn = 0
    mblnFinished = False
    anal = typBlank
    MapItemSpots gs
    CreateItemSpotKey gs, anal
    CreateEffectSpotKey gs, anal
    InitProgress gs, anal
    ' Reset counters
    For i = 1 To gs.Effects
        anal.Effect(i).Current = 0
        ReDim anal.Effect(i).ValidSpot(SpotKeyMax)
    Next
    anal.ItemSpot = anal.ItemSpotKey
    With anal
        mlngCopyMemoryLength = (UBound(.ItemSpotKey) + 1) * LenB(.ItemSpotKey(0))
    End With
End Sub

' Create dynamic base number system for charting progress
Private Sub InitProgress(gs As GearsetType, anal As AnalysisType)
    Dim i As Long
    
    mlngProgressMax = 1
    For mlngProgressTerms = 1 To gs.Effects
        mlngProgressMax = mlngProgressMax * (anal.Effect(mlngProgressTerms).Spots + 1)
        If mlngProgressMax > 1000 Then Exit For
    Next
    If mlngProgressTerms > gs.Effects Then mlngProgressTerms = gs.Effects
    ReDim mlngProgressTerm(1 To mlngProgressTerms)
    mlngProgressTerm(mlngProgressTerms) = 1
    For i = mlngProgressTerms - 1 To 1 Step -1
        mlngProgressTerm(i) = mlngProgressTerm(i + 1) * (anal.Effect(i + 1).Spots + 1)
    Next
End Sub


' ************* MAPPING *************


' Create mapping that points every ITEM in our gearset directly to its proper position in ItemSpotKey
Public Sub MapItemSpots(gs As GearsetType)
    gs.Item(seHelmet).SpotKey = spotHelmet
    gs.Item(seGoggles).SpotKey = spotGoggles
    gs.Item(seNecklace).SpotKey = spotNecklace
    gs.Item(seCloak).SpotKey = spotCloak
    gs.Item(seBracers).SpotKey = spotBracers
    gs.Item(seGloves).SpotKey = spotGloves
    gs.Item(seBelt).SpotKey = spotBelt
    gs.Item(seBoots).SpotKey = spotBoots
    gs.Item(seRing1).SpotKey = spotRing
    gs.Item(seRing2).SpotKey = spotRing
    gs.Item(seTrinket).SpotKey = spotTrinket
    gs.Item(seArmor).SpotKey = spotArmor
    If gs.MainHand = mheRange Then
        gs.Item(seMainHand).SpotKey = spotRange
    ElseIf gs.TwoHanded Then
        gs.Item(seMainHand).SpotKey = spotMelee2H
    Else
        gs.Item(seMainHand).SpotKey = spotMelee1H
    End If
    Select Case gs.OffHand
        Case oheMelee: gs.Item(seOffHand).SpotKey = spotMelee1H
        Case oheShield: gs.Item(seOffHand).SpotKey = spotShield
        Case oheOrb: gs.Item(seOffHand).SpotKey = spotOrb
        Case oheRunearm: gs.Item(seOffHand).SpotKey = spotRunearm
    End Select
End Sub

' Count up number of slots of each type in gearset. For example, if the gearset has 2 rings,
' we put 2 in ring prefix, ring suffix, and (if ML10+) ring extra
Public Sub CreateItemSpotKey(gs As GearsetType, anal As AnalysisType)
    Dim lngSpots As Long
    Dim lngIndex As Long
    Dim i As Long
    Dim s As Long
    
    ReDim anal.ItemSpotKey(SpotKeyMax)
    If gs.BaseLevel < 10 Then lngSpots = 1 Else lngSpots = 2
    For i = 0 To seSlotCount - 1
        If gs.Item(i).Crafted Then
            For s = 0 To lngSpots
                lngIndex = gs.Item(i).SpotKey + s
                anal.ItemSpotKey(lngIndex) = anal.ItemSpotKey(lngIndex) + 1
            Next
        End If
    Next
End Sub

' Doesn't need to be fast (only runs once)
Private Sub CreateEffectSpotKey(gs As GearsetType, anal As AnalysisType)
    Dim lngEffect As Long
    Dim enGear As GearEnum
    
    ReDim anal.Effect(gs.Effects)
    For lngEffect = 1 To gs.Effects
        anal.Effect(lngEffect).Spots = -1
        With db.Shard(gs.Effect(lngEffect))
            For enGear = geHelmet To geRuneArm
                If enGear <> geHandwraps Then ' Ignore handwraps (they're a true duplicate of 2hMelee)
                    SetEffectSpot anal.Effect(lngEffect), aePrefix, .Prefix, enGear, anal
                    SetEffectSpot anal.Effect(lngEffect), aeSuffix, .Suffix, enGear, anal
                    SetEffectSpot anal.Effect(lngEffect), aeExtra, .Extra, enGear, anal
                End If
            Next
            ' Handle armor separately, depending on type chosen
            Select Case gs.Armor
                Case ameMetal: enGear = geMetalArmor
                Case ameLeather: enGear = geLeatherArmor
                Case ameCloth: enGear = geClothArmor
                Case ameDocent: enGear = geDocent
            End Select
            SetEffectSpot anal.Effect(lngEffect), aePrefix, .Prefix, enGear, anal
            SetEffectSpot anal.Effect(lngEffect), aeSuffix, .Suffix, enGear, anal
            SetEffectSpot anal.Effect(lngEffect), aeExtra, .Extra, enGear, anal
        End With
    Next
End Sub

Private Sub SetEffectSpot(ptypEffect As EffectSpotType, penAffix As AffixEnum, pblnSlot() As Boolean, penGear As GearEnum, anal As AnalysisType)
    Dim lngSpot As Long
    
    If Not pblnSlot(penGear) Then Exit Sub
    lngSpot = MapEffectSpot(penGear) + penAffix
    If anal.ItemSpotKey(lngSpot) = 0 Then Exit Sub
    With ptypEffect
        .Spots = .Spots + 1
        ReDim Preserve .Spot(.Spots)
        .Spot(.Spots) = lngSpot
    End With
End Sub

' Create mapping that points every EFFECT in our gearset directly to its proper position in ItemSpotKey
Public Function MapEffectSpot(penGear As GearEnum) As SpotEnum
    Select Case penGear
        Case geHelmet: MapEffectSpot = spotHelmet
        Case geGoggles: MapEffectSpot = spotGoggles
        Case geNecklace: MapEffectSpot = spotNecklace
        Case geCloak: MapEffectSpot = spotCloak
        Case geBracers: MapEffectSpot = spotBracers
        Case geGloves: MapEffectSpot = spotGloves
        Case geBelt: MapEffectSpot = spotBelt
        Case geBoots: MapEffectSpot = spotBoots
        Case geRing: MapEffectSpot = spotRing
        Case geTrinket: MapEffectSpot = spotTrinket
        Case ge2hMelee, geHandwraps: MapEffectSpot = spotMelee2H
        Case ge1hMelee: MapEffectSpot = spotMelee1H
        Case geShield: MapEffectSpot = spotShield
        Case geRange: MapEffectSpot = spotRange
        Case geRuneArm: MapEffectSpot = spotRunearm
        Case geOrb: MapEffectSpot = spotOrb
        Case geMetalArmor, geLeatherArmor, geClothArmor, geDocent: MapEffectSpot = spotArmor
    End Select
End Function


' ************* ANALYZE *************


' For some reason, DoEvents appears to not do anything, so I had to look for
' an alternative to allow users to hit the cancel button during processing.
' So I rewrote the code to process in 1-second chunks, then stop and wait for
' the next call to resume where it left off.
' Then I set up a timer with minimum interval to call the "Process a Chunk"
' function, resuming where the last chunk finished.
' Optimization testing specs: 30 effects that fit in 30 slots, check 159.3 million combinations, find 126 valid combinations
Public Sub ProcessChunk(gs As GearsetType, anal As AnalysisType)
    Dim lngEffect As Long
    Dim lngSpot As Long
    Dim lngStop As Long
'    Dim lngCurrent As Long
'    Dim lngEffects As Long
    
    ' Storing gs.Effects to a variable ahead of time actually slows things down for some reason
'    lngEffects = gs.Effects
    lngStop = Int(StopwatchStop()) + 1
    Do
        ' Much faster than natively copying array (anal.ItemSpot = anal.ItemSpotKey)
        CopyMemory anal.ItemSpot(0), anal.ItemSpotKey(0), mlngCopyMemoryLength
        ' Check if current permutation is valid
        For lngEffect = 1 To gs.Effects
            ' lngSpot isn't just for readability; removing it and referencing udt chain directly more than doubles the execution time (!)
            lngSpot = anal.Effect(lngEffect).Spot(anal.Effect(lngEffect).Current)
            If anal.ItemSpot(lngSpot) <> 0 Then
                anal.ItemSpot(lngSpot) = anal.ItemSpot(lngSpot) - 1
            Else
                ' This permutation failed on lngEffect; increment effect counters
                If lngEffect > mlngFailedOn Then mlngFailedOn = lngEffect
                ' Using an evil GoTo here removes 159.3 million unecessary comparisons
                GoTo IncrementCounters
            End If
        Next
        mlngValid = mlngValid + 1
        ' This combination is valid, so mark the position of every effect in this combination as valid
        For lngEffect = 1 To gs.Effects
            ' Test run results:
            ' 1:17-1:20 using a single line: anal.Effect(lngEffect).ValidSpot(anal.Effect(lngEffect).Spot(anal.Effect(lngEffect).Current)) = True
            anal.Effect(lngEffect).ValidSpot(anal.Effect(lngEffect).Spot(anal.Effect(lngEffect).Current)) = True
            ' 1:19 storing indexes to variables then referencing the indexes
'            lngCurrent = anal.Effect(lngEffect).Current
'            lngSpot = anal.Effect(lngEffect).Spot(lngCurrent)
'            anal.Effect(lngEffect).ValidSpot(lngSpot) = True
            ' 1:22 using With anal.Effect(lngEffect)
'            With anal.Effect(lngEffect)
'                .ValidSpot(.Spot(.Current)) = True
'            End With
        Next
        lngEffect = gs.Effects
IncrementCounters:
        ' Increment counters
        ' NOTE: This loop  looks like it could start from gs.Effects, but it cannot. The short-circuited value of
        ' lngEffect from the Goto line above when lngEffect fails is what makes this whole thing "fast"
        ' Also note that on a success, incrementing begins at the least significant "digit"
        ' On failure, incrementing begins at the failed effect because every less significant "digit" is already 0 by definition
        ' (Failures fail on their first try, meaning the moment it is incremented. Which happens when every "digit" after it is zero.)
        For lngEffect = lngEffect To 1 Step -1
            If anal.Effect(lngEffect).Current < anal.Effect(lngEffect).Spots Then
                anal.Effect(lngEffect).Current = anal.Effect(lngEffect).Current + 1
                Exit For
            Else
                anal.Effect(lngEffect).Current = 0
            End If
        Next
        mlngCount = mlngCount + 1
        ' Inlining StopwatchStop() had no measurable effect
        If StopwatchStop() > lngStop Then Exit Sub
    Loop While lngEffect <> 0
    mblnFinished = True
End Sub

Public Function GetProgress(gs As GearsetType, anal As AnalysisType) As Double
    Dim lngProgress As Long
    Dim i As Long
    
    If mblnFinished Then
        lngProgress = mlngProgressMax
    Else
        For i = 1 To mlngProgressTerms
            lngProgress = lngProgress + anal.Effect(i).Current * mlngProgressTerm(i)
        Next
    End If
    GetProgress = CDbl(lngProgress / mlngProgressMax)
End Function

Public Function ProcessingFinished() As Boolean
    ProcessingFinished = mblnFinished
End Function

Public Function GetValid() As Long
    GetValid = mlngValid
End Function

' Index in gs.Effect()
Public Function GetFailedOn() As Long
    GetFailedOn = mlngFailedOn
End Function

Public Function GetCombinations()
    GetCombinations = mlngCount
End Function

Public Function FormattedCombinations() As String
    If mlngCount < 10000 Then
        FormattedCombinations = Format(mlngCount, "#,##0")
    ElseIf mlngCount < 1000000 Then
        FormattedCombinations = Format(mlngCount / 1000, "#,##0") & "k"
    ElseIf mlngCount < 10000000 Then
        FormattedCombinations = Format(mlngCount / 1000000, "0.00") & " million"
    ElseIf mlngCount < 1000000000 Then
        FormattedCombinations = Format(mlngCount / 1000000, "0.0") & " million"
    Else
        FormattedCombinations = Format(mlngCount / 1000000000, "0.000") & " billion"
    End If
End Function


' ************* ORIGINAL ALGORITHM *************


Public Sub InitAnalysis(gs As GearsetType, anal As AnalysisType)
    Dim typBlank As AnalysisType
    
    anal = typBlank
    MapItemSpots gs
    CreateItemSpotKey gs, anal
    CreateEffectSpotKey gs, anal
End Sub

' Unoptimized version of original algorithm (starting point)
Public Function Analyze(gs As GearsetType, anal As AnalysisType) As String
    Dim lngEffect As Long
    Dim lngCurrent As Long
    Dim lngSpot As Long
    Dim lngValid As Long
    Dim lngFailedOn As Long
    Dim i As Long
    
    ' Reset counters
    For lngEffect = 1 To gs.Effects
        anal.Effect(lngEffect).Current = 0
    Next
    Do
        anal.ItemSpot = anal.ItemSpotKey
        ' Check if current permutation is valid
        For lngEffect = 1 To gs.Effects
            lngSpot = anal.Effect(lngEffect).Spot(anal.Effect(lngEffect).Current)
            If anal.ItemSpot(lngSpot) <> 0 Then
                anal.ItemSpot(lngSpot) = anal.ItemSpot(lngSpot) - 1
            Else
                ' This permutation failed on lngEffect; increment effect counters
                If lngEffect > lngFailedOn Then lngFailedOn = lngEffect
                Exit For
            End If
        Next
        ' Valid permutation?
        If lngEffect > gs.Effects Then
            lngValid = lngValid + 1
            lngEffect = gs.Effects
            ' Add code here to save this valid combination
        End If
        ' Increment counters
        For lngEffect = lngEffect To 1 Step -1
            If anal.Effect(lngEffect).Current < anal.Effect(lngEffect).Spots Then
                anal.Effect(lngEffect).Current = anal.Effect(lngEffect).Current + 1
                Exit For
            Else
                anal.Effect(lngEffect).Current = 0
            End If
        Next
    Loop While lngEffect
    If lngValid Then
        Analyze = lngValid & " valid combinations"
    Else
        Analyze = "Failed on " & db.Shard(gs.Effect(lngFailedOn)).ShardName
    End If
End Function

