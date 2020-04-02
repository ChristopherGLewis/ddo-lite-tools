Attribute VB_Name = "basQuickMatch"
Option Explicit

Private Const QMSpotKeyMax As Long = 41
Private Const SeedValue As Double = 10

' A "Spot" is defined as a Gear Slot + Affix position
Private Enum QMSpotEnum
    qmUnused = -1
    qmHelmet = 0
    qmGoggles = 3
    qmNecklace = 6
    qmCloak = 9
    qmBracers = 12
    qmGloves = 15
    qmBelt = 18
    qmBoots = 21
    qmRing1 = 24
    qmRing2 = 27
    qmTrinket = 30
    qmArmor = 33
    qmMainHand = 36
    qmOffHand = 39
End Enum

Private Type QuickMatchEffectType
    Slot As Long
    BestSlot As Long ' Slot value for iteration with best results
    Include As Boolean
    Possible As Long
    Value As Double
    Spots As Long ' Number of TRUE Spotkeys
    Spot() As Long ' List of TRUE SpotKeys
    SpotKey() As Boolean ' SpotKey(QMSpotKey) = True if Effect goes in that spot
End Type

Private Type QuickMatchType
    SlotValue() As Double
    SlotAvailable() As Boolean
    Effects As Long
    Effect() As QuickMatchEffectType
    Failed As Long
    Iterations As Long
End Type

Dim qm As QuickMatchType


' ************* METHODS *************


Public Sub QuickMatchCheck(gs As GearsetType, plngIterations As Long, pdblTime As Double, plngFailed As Long)
    QMatch gs
    pdblTime = StopwatchStop()
    plngIterations = qm.Iterations
    plngFailed = qm.Failed
End Sub

Public Sub QuickMatchCommit(gs As GearsetType)
    QMatch gs
    CommitResults gs
End Sub


' ************* ENTRY POINT *************


' This is the meta-function that repeats the actual function until success or 0.1 seconds have elapsed
Private Sub QMatch(gs As GearsetType)
    Dim lngFailed As Long
    
    StopwatchStart
    InitQuickMatch gs
    Do
        ResetEffects
        ResetAvailableSpots gs
        qm.Iterations = qm.Iterations + 1
        lngFailed = QuickMatch()
        If lngFailed = 0 Then
            qm.Failed = 0
            Exit Do
        ElseIf qm.Failed < lngFailed Then
            UpdateFailed lngFailed
        End If
    Loop While StopwatchStop() < 0.1
End Sub

Private Sub InitQuickMatch(gs As GearsetType)
    Dim typBlank As QuickMatchType
    Dim lngEffect As Long
    Dim lngShard As Long
    Dim enSlot As SlotEnum
    Dim enGear As GearEnum
    Dim enAffix As AffixEnum
    Dim enLastAffix As AffixEnum
    Dim enSpot As QMSpotEnum
    
    qm = typBlank ' Quickly and easily zero out ALL data
    If gs.BaseLevel < 10 Then enLastAffix = aeSuffix Else enLastAffix = aeExtra
    ' Init Effects
    qm.Effects = gs.Effects
    ReDim qm.Effect(qm.Effects)
    ' Create two different lookups detailing which spots each effect can go
    For lngEffect = 1 To gs.Effects
        With qm.Effect(lngEffect)
            ReDim .SpotKey(QMSpotKeyMax)
            .Spots = -1
            ReDim .Spot(8) ' Allocate a reasonable amount of spots in advance
        End With
        lngShard = gs.Effect(lngEffect)
        For enSlot = seHelmet To seSlotCount - 1
            If gs.Item(enSlot).Crafted Then
                enGear = gs.Item(enSlot).Gear
                For enAffix = aePrefix To enLastAffix
                    enSpot = enSlot * 3 + enAffix
                    Select Case enAffix
                        Case aePrefix: InitEffectQM qm.Effect(lngEffect), enSpot, db.Shard(lngShard).Prefix(enGear)
                        Case aeSuffix: InitEffectQM qm.Effect(lngEffect), enSpot, db.Shard(lngShard).Suffix(enGear)
                        Case aeExtra: InitEffectQM qm.Effect(lngEffect), enSpot, db.Shard(lngShard).Extra(enGear)
                    End Select
                Next
            End If
        Next
        ' Reclaim unused space
        With qm.Effect(lngEffect)
            If .Spots = -1 Then
                Erase .Spot
            ElseIf UBound(.Spot) <> .Spots Then
                ReDim Preserve .Spot(.Spots)
            End If
        End With
    Next
End Sub

' For each valid spot for an effect:
' 1) Set that spot as TRUE in the effect's SpotKey(spot)
' 2) Add that spot to the effect's Spots list
Private Sub InitEffectQM(ptypEffect As QuickMatchEffectType, penSpot As QMSpotEnum, pblnValid As Boolean)
    If Not pblnValid Then Exit Sub
    With ptypEffect
        .SpotKey(penSpot) = True
        .Spots = .Spots + 1
        ' Generally speaking, the initial allocation of 8 (9 total spots counting 0) should be plenty for most effects
        ' If we happen to go over, add a few more spots at a time
        If .Spots > UBound(.Spot) Then ReDim Preserve .Spot(.Spots + 4)
        .Spot(.Spots) = penSpot
    End With
End Sub

' Save this failed run's results in case it ends up being the best fail
Private Sub UpdateFailed(plngEffect As Long)
    Dim i As Long
    
    qm.Failed = plngEffect
    For i = 1 To qm.Effects
        qm.Effect(i).BestSlot = qm.Effect(i).Slot
    Next
End Sub

Private Sub CommitResults(gs As GearsetType)
    Dim lngSlot As Long
    Dim enSlot As SlotEnum
    Dim enAffix As AffixEnum
    Dim lngValue As Long
    Dim lngEffect As Long
    
    ' Zero out all slot info
    For enSlot = 0 To seSlotCount - 1
        For enAffix = aePrefix To aeExtra
            gs.Item(enSlot).Effect(enAffix) = 0
        Next
    Next
    For lngEffect = 1 To qm.Effects
        If qm.Failed Then lngSlot = qm.Effect(lngEffect).BestSlot Else lngSlot = qm.Effect(lngEffect).Slot
        If lngSlot <> -1 Then
            enSlot = lngSlot \ 3
            enAffix = lngSlot Mod 3
            gs.Item(enSlot).Effect(enAffix) = gs.Effect(lngEffect)
        End If
    Next
End Sub


' ************* PREP *************


' Zero out effect counters
Private Sub ResetEffects()
    Dim i As Long
    
    For i = 1 To qm.Effects
        With qm.Effect(i)
            .Include = False
            .Possible = 0
            .Slot = -1
            .Value = 0
        End With
    Next
End Sub

' Reset available spots to defaults
Private Sub ResetAvailableSpots(gs As GearsetType)
    Dim enSlot As SlotEnum
    Dim enGear As GearEnum
    Dim enAffix As AffixEnum
    Dim enLastAffix As AffixEnum
    Dim enSpot As QMSpotEnum
    
    If gs.BaseLevel < 10 Then enLastAffix = aeSuffix Else enLastAffix = aeExtra
    ReDim qm.SlotAvailable(QMSpotKeyMax)
    For enSlot = seHelmet To seSlotCount - 1
        If gs.Item(enSlot).Crafted Then
            enGear = gs.Item(enSlot).Gear
            For enAffix = aePrefix To enLastAffix
                enSpot = enSlot * 3 + enAffix
                qm.SlotAvailable(enSpot) = True
            Next
        End If
    Next
End Sub


' ************* ALGORITHM *************


Private Function QuickMatch() As Long
    Dim lngEffect As Long
    
    lngEffect = 1
    Do
        If Not QuickMatchMainLoop(lngEffect) Then Exit Do
    Loop While lngEffect
    QuickMatch = lngEffect
End Function

Private Function QuickMatchMainLoop(plngEffect As Long) As Boolean
    Dim lngSpotList() As Long ' List of spots the current effect can still go
    Dim lngSpotListMax As Long ' UBound(lngSpotList) or -1 if list is empty
    Dim lngSpot As Long ' Chosen Spot
    Dim lngEffect As Long ' Chosen Effect
    
    ' Find all the spots where the current effect can still be placed
    If Not CreateSpotList(plngEffect, lngSpotList, lngSpotListMax) Then Exit Function
    ' For effects that can be slotted anywhere in the current spot list, calculate values
    CalculateEffectValues plngEffect, lngSpotList, lngSpotListMax
    CalculateSlotValues plngEffect, lngSpotList, lngSpotListMax
    ' Get the lowest spot value
    lngSpot = ChooseSpot(lngSpotList, lngSpotListMax)
    ' Get the highest effect value
    lngEffect = ChooseEffect(lngSpot, plngEffect)
    ' Slot effect into spot
    ApplyEffectToSpot lngEffect, lngSpot
    ' Increment plngEffect to the next unslotted effect
    NextEffect plngEffect
    ' Success
    QuickMatchMainLoop = True
End Function

' Find all the spots where the current effect can still be placed
Private Function CreateSpotList(plngEffect As Long, plngSpotList() As Long, plngSpotListMax As Long) As Boolean
    Dim i As Long
    
    Erase plngSpotList
    plngSpotListMax = -1
    For i = 0 To QMSpotKeyMax
        ' If effect goes in ItemSpot and ItemSpot isn't filled yet...
        If qm.Effect(plngEffect).SpotKey(i) And qm.SlotAvailable(i) Then
            plngSpotListMax = plngSpotListMax + 1
            ReDim Preserve plngSpotList(plngSpotListMax)
            plngSpotList(plngSpotListMax) = i
        End If
    Next
    CreateSpotList = (plngSpotListMax <> -1)
End Function

Private Sub CalculateEffectValues(plngEffect As Long, plngSpotList() As Long, plngSpotListMax As Long)
    Dim lngEffect As Long
    Dim lngSpot As Long
    Dim i As Long
    
    ' Find all effects that can go anywhere in the spot list
    For lngEffect = plngEffect To qm.Effects
        With qm.Effect(lngEffect)
            .Possible = 0
            ' Ignore effects that have already been slotted
            If .Slot = -1 Then
                ' Can this effect go anywhere in the spot list?
                For i = 0 To plngSpotListMax
                    If .SpotKey(plngSpotList(i)) Then
                        ' Add up all slots this effect can still go in
                        For lngSpot = 0 To .Spots
                            If qm.SlotAvailable(.Spot(lngSpot)) Then .Possible = .Possible + 1
                        Next
                        ' We finished tallying this effect; move on to the next effect
                        Exit For
                    End If
                Next
                ' Calculate this effect's value
                If .Possible > 0 Then .Value = SeedValue / .Possible
            End If
        End With
    Next
End Sub

Private Sub CalculateSlotValues(plngEffect As Long, plngSpotList() As Long, plngSpotListMax As Long)
    Dim lngSpot As Long
    Dim lngEffect As Long
    
    ' Zero out spot value totals
    ReDim qm.SlotValue(QMSpotKeyMax)
    ' Add up values for all spots in the spot list
    For lngSpot = 0 To plngSpotListMax
        For lngEffect = plngEffect To qm.Effects
            ' If effect hasn't been slotted yet and can be slotted in this spot...
            If qm.Effect(lngEffect).Slot = -1 And qm.Effect(lngEffect).SpotKey(plngSpotList(lngSpot)) = True Then
                ' ...add effect's value to spot's value
                qm.SlotValue(plngSpotList(lngSpot)) = qm.SlotValue(plngSpotList(lngSpot)) + qm.Effect(lngEffect).Value
            End If
        Next
    Next
End Sub

' Choose a spot to fill (must have lowest value, or be tied for lowest value)
Private Function ChooseSpot(plngSpotList() As Long, plngSpotListMax As Long) As Long
    Dim lngIndex() As Long ' List of spots tied for lowest value
    Dim lngCount As Long ' Count of spots tied for lowest value
    Dim dblLowest As Double
    Dim i As Long
    
    ' Start with first value
    dblLowest = qm.SlotValue(plngSpotList(0))
    ReDim lngIndex(0) ' Initializes to zeroes, so no need to actually put a 0 there
    lngCount = 1
    ' Loop remaining values, starting with the second value
    For i = 1 To plngSpotListMax
        If dblLowest > qm.SlotValue(plngSpotList(i)) Then
            ' This new value is lower than previous low, so reset list
            dblLowest = qm.SlotValue(plngSpotList(i))
            ReDim lngIndex(0)
            lngIndex(0) = i
            lngCount = 1
        ElseIf dblLowest = qm.SlotValue(plngSpotList(i)) Then
            ' Tied; add this index to list
            ReDim Preserve lngIndex(lngCount)
            lngIndex(lngCount) = i
            lngCount = lngCount + 1
        End If
    Next
    ' Randomly choose a value from the list of ties
    ChooseSpot = plngSpotList(lngIndex(Int(lngCount * Rnd)))
End Function

' Choose an effect to slot (must have highest value, or be tied for highest value)
Private Function ChooseEffect(plngSpot As Long, plngEffect As Long) As Long
    Dim lngIndex() As Long ' List of effects tied for highest value
    Dim lngCount As Long ' Count of effects tied for highest value
    Dim dblHighest As Double
    Dim i As Long
    
    For i = plngEffect To qm.Effects
        ' If effect hasn't been slotted and can be slotted in plngSpot...
        If qm.Effect(i).Slot = -1 And qm.Effect(i).SpotKey(plngSpot) = True Then
            If dblHighest < qm.Effect(i).Value Then
                ' This new value is higher than previous high, so reset list
                dblHighest = qm.Effect(i).Value
                ReDim lngIndex(0)
                lngIndex(0) = i
                lngCount = 1
            ElseIf dblHighest = qm.Effect(i).Value Then
                ' Tied; add this index to list
                ReDim Preserve lngIndex(lngCount)
                lngIndex(lngCount) = i
                lngCount = lngCount + 1
            End If
        End If
    Next
    ' Randomly choose a value from the list of ties
    ChooseEffect = lngIndex(Int(lngCount * Rnd))
End Function

Private Sub ApplyEffectToSpot(plngEffect As Long, plngSpot As Long)
    qm.SlotAvailable(plngSpot) = False
    With qm.Effect(plngEffect)
        .Slot = plngSpot
        .Value = 0
        .Possible = 0
        .Include = False
    End With
End Sub

Private Sub NextEffect(plngEffect As Long)
    ' Increment plngEffect to next unslotted effect
    For plngEffect = plngEffect To qm.Effects
        If qm.Effect(plngEffect).Slot = -1 Then Exit Sub
    Next
    ' No more unslotted effects, so clear plngEffect
    plngEffect = 0
End Sub


' ************* ORIGINAL IMPLEMENTATION *************


' Original code before refactoring
Private Function QuickMatchOld(plngEffect As Long) As Boolean
    Dim lngEffect As Long
    Dim lngSpot As Long
    Dim lngItemSpots As Long
    Dim lngItemSpot() As Long
    Dim dblValue As Double
    Dim lngCount As Long
    Dim lngIndex() As Long
    Dim i As Long

    ' Reset effect values
    For lngEffect = 1 To qm.Effects
        With qm.Effect(lngEffect)
            .Include = False
            .Possible = 0
            .Value = 0
            .Slot = 0
        End With
    Next
    ReDim qm.SlotValue(QMSpotKeyMax)
    ' Find all the spots where the current effect can be placed
    lngItemSpots = -1
    For i = 0 To QMSpotKeyMax
        ' If effect goes in ItemSpot and ItemSpot isn't filled yet...
        If qm.Effect(plngEffect).SpotKey(i) And qm.SlotAvailable(i) Then
            lngItemSpots = lngItemSpots + 1
            ReDim Preserve lngItemSpot(lngItemSpots)
            lngItemSpot(lngItemSpots) = i
        End If
    Next
    ' If current effect can't be slotted anywhere, fail and return
    If lngItemSpots = -1 Then
        QuickMatchOld = True
        Exit Function
    End If
    ' Find all effects that can go in any slot the current effect can go
    For lngEffect = plngEffect To qm.Effects
        With qm.Effect(lngEffect)
            ' Ignore effects that have already been slotted
            If .Slot = -1 Then
                ' Can this effect go in any slot the current effect can go?
                For i = 0 To lngItemSpots
                    If .SpotKey(lngItemSpot(i)) Then
                        ' Add up all slots this effect can still go in
                        For lngSpot = 0 To .Spots
                            If qm.SlotAvailable(.Spot(lngSpot)) Then .Possible = .Possible + 1
                        Next
                        ' We finished tallying this effect; move on to the next effect
                        Exit For
                    End If
                Next
                ' Calculate this effect's value
                If .Possible > 0 Then .Value = SeedValue / .Possible
            End If
        End With
    Next
    ' Add up values for all slots the current effect can go
    For i = 0 To lngItemSpots
        For lngEffect = plngEffect To qm.Effects
            If qm.Effect(lngEffect).Slot = -1 And qm.Effect(lngEffect).SpotKey(lngItemSpot(i)) = True Then qm.SlotValue(lngItemSpot(i)) = qm.SlotValue(lngItemSpot(i)) + qm.Effect(lngEffect).Value
        Next
    Next
    ' Find the smallest slot value
    ReDim dblValueList(0)
    dblValue = qm.SlotValue(lngItemSpot(0))
    ReDim lngIndex(0)
    lngCount = 1
    For i = 1 To lngItemSpots
        If dblValue > qm.SlotValue(lngItemSpot(i)) Then
            dblValue = qm.SlotValue(lngItemSpot(i))
            ReDim lngIndex(0)
            lngIndex(0) = i
            lngCount = 1
        ElseIf dblValue = qm.SlotValue(lngItemSpot(i)) Then
            ReDim Preserve lngIndex(lngCount)
            lngIndex(lngCount) = i
            lngCount = lngCount + 1
        End If
    Next
    ' If there's a tie, randomly choose among the tied
    lngSpot = lngItemSpot(lngIndex(Int(lngCount * Rnd)))
    ' For this chosen spot, which effects contributed the largest value?
    Erase lngIndex
    lngCount = 0
    dblValue = 0
    For lngEffect = plngEffect To qm.Effects
        If qm.Effect(lngEffect).Slot = -1 And qm.Effect(lngEffect).SpotKey(lngSpot) = True Then
            If dblValue < qm.Effect(lngEffect).Value Then
                dblValue = qm.Effect(lngEffect).Value
                ReDim lngIndex(0)
                lngIndex(0) = lngEffect
                lngCount = 1
            ElseIf dblValue = qm.Effect(lngEffect).Value Then
                ReDim Preserve lngIndex(lngCount)
                lngIndex(lngCount) = lngEffect
                lngCount = lngCount + 1
            End If
        End If
    Next
    ' If there's a tie, randomly choose among the tied
    lngEffect = lngIndex(Int(lngCount * Rnd))
    ' Apply this effect to this slot
    qm.SlotAvailable(lngSpot) = False
    qm.Effect(lngEffect).Slot = lngSpot
    qm.Effect(lngEffect).Value = 0
    qm.Effect(lngEffect).Possible = 0
    qm.Effect(lngEffect).Include = False
    ' Increment plngEffect to next unslotted effect
    For plngEffect = plngEffect To qm.Effects
        If qm.Effect(plngEffect).Slot = -1 Then Exit Function
    Next
    ' No more unslotted effects, so clear plngEffect
    plngEffect = 0
End Function

