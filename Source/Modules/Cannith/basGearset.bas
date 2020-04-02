Attribute VB_Name = "basGearset"
Option Explicit


' ************* GEARSET *************


Public Sub AddEffect(ptypGearset As GearsetType, pstrEffect As String)
    Dim lngShard As Long
    
    lngShard = SeekShard(pstrEffect)
    If lngShard = 0 Then
        Debug.Print "Shard not found: " & pstrEffect
    Else
        With ptypGearset
            .Effects = .Effects + 1
            ReDim Preserve .Effect(1 To .Effects)
            .Effect(.Effects) = lngShard
        End With
    End If
End Sub

Public Function CheckForErrors(gs As GearsetType, pinfo As userInfo) As Boolean
    Dim lngShard() As Long
    Dim blnItem(GearMax) As Boolean
    Dim i As Long
    Dim j As Long

    MapSlotsToGear gs
    For i = 0 To seSlotCount - 1
        If gs.Item(i).Crafted Then blnItem(gs.Item(i).Gear) = True
    Next
    If gs.Item(seArmor).Crafted Then
        For i = 1 To gs.Effects
            With db.Shard(gs.Effect(i))
                If Len(.Warning) Then pinfo.AddText "Note: " & .Warning, 2
                If .ML > gs.BaseLevel Then
                    pinfo.AddError "Error: " & .Abbreviation & " is ML" & .ML, 2
                    CheckForErrors = True
                End If
            End With
        Next
    End If
    For i = 1 To gs.Effects
        If Not EffectCanBeSlotted(blnItem, db.Shard(gs.Effect(i)).Prefix) Then
            If Not EffectCanBeSlotted(blnItem, db.Shard(gs.Effect(i)).Suffix) Then
                If Not EffectCanBeSlotted(blnItem, db.Shard(gs.Effect(i)).Extra) Then
                    pinfo.AddError "Error: " & db.Shard(gs.Effect(i)).Abbreviation & " can't be slotted", 2
                    CheckForErrors = True
                End If
            End If
        End If
    Next
End Function

Private Function EffectCanBeSlotted(pblnItem() As Boolean, pblnSlot() As Boolean) As Boolean
    Dim i As Long
    
    For i = 0 To GearMax
        If pblnSlot(i) Then
            If pblnItem(i) Then
                EffectCanBeSlotted = True
                Exit Function
            End If
        End If
    Next
End Function

Public Function MapSlotsToGear(gs As GearsetType)
    MapSlotToGear gs, seHelmet, geHelmet
    MapSlotToGear gs, seGoggles, geGoggles
    MapSlotToGear gs, seNecklace, geNecklace
    MapSlotToGear gs, seCloak, geCloak
    MapSlotToGear gs, seBracers, geBracers
    MapSlotToGear gs, seGloves, geGloves
    MapSlotToGear gs, seBelt, geBelt
    MapSlotToGear gs, seBoots, geBoots
    MapSlotToGear gs, seRing1, geRing
    MapSlotToGear gs, seRing2, geRing
    MapSlotToGear gs, seTrinket, geTrinket
    Select Case gs.Armor
        Case ameMetal: MapSlotToGear gs, seArmor, geMetalArmor
        Case ameLeather: MapSlotToGear gs, seArmor, geLeatherArmor
        Case ameCloth: MapSlotToGear gs, seArmor, geClothArmor
        Case ameDocent: MapSlotToGear gs, seArmor, geDocent
    End Select
    Select Case gs.Mainhand
        Case mheMelee
            If gs.TwoHanded Then
                MapSlotToGear gs, seMainHand, ge2hMelee
            Else
                MapSlotToGear gs, seMainHand, ge1hMelee
            End If
        Case mheRange
            MapSlotToGear gs, seMainHand, geRange
    End Select
    Select Case gs.Offhand
        Case oheMelee: MapSlotToGear gs, seOffHand, ge1hMelee
        Case oheShield: MapSlotToGear gs, seOffHand, geShield
        Case oheOrb: MapSlotToGear gs, seOffHand, geOrb
        Case oheRunearm: MapSlotToGear gs, seOffHand, geRunearm
    End Select
End Function

Private Sub MapSlotToGear(gs As GearsetType, penSlot As SlotEnum, penGear As GearEnum)
    ' I can't figure out what the original purpose was for mapping named items to geUnknown
    ' Leaving this here in case it was actually useful for something and I need to go back to it
    'If gs.Item(penSlot).Crafted Then gs.Item(penSlot).Gear = penGear Else gs.Item(penSlot).Gear = geUnknown
    gs.Item(penSlot).Gear = penGear
End Sub


' ************* AUGMENTS *************


' Can't pass udts to/from user controls, so these function convert between the udt array and strings
Public Function GearsetAugmentToString(ptypAugSlot() As AugmentSlotType) As String
    Dim lngCount As Long
    Dim strReturn As String
    Dim i As Long
    
    For i = 1 To 7
        With ptypAugSlot(i)
            If .Exists Then lngCount = lngCount + 1
            If lngCount < 4 Then
                strReturn = strReturn & "|" & .Exists & ";" & .Augment & ";" & .Variation & ";" & .Scaling & ";" & .Done
            Else
                strReturn = strReturn & "|" & .Exists & ";0;0;0" & .Done
            End If
        End With
    Next
    GearsetAugmentToString = strReturn
End Function

Public Function StringToGearsetAugment(ptypAugSlot() As AugmentSlotType, pstrRaw As String) As Boolean
    Dim lngCount As Long
    Dim strSlot() As String
    Dim strAugment() As String
    Dim blnExists As Boolean
    Dim i As Long
    
    Erase ptypAugSlot
    strSlot = Split(pstrRaw, "|")
    If UBound(strSlot) <> 7 Then
        Debug.Print "Invalid GearsetAugment String: " & pstrRaw
        Exit Function
    End If
    For i = 1 To 7
        strAugment = Split(strSlot(i), ";")
        If UBound(strAugment) <> 4 Then
            Debug.Print "Invalid GearsetAugment String: " & pstrRaw
        Else
            blnExists = (strAugment(0) = "True")
            With ptypAugSlot(i)
                If blnExists Then lngCount = lngCount + 1
                If lngCount < 4 Then
                    .Exists = blnExists
                    .Augment = Val(strAugment(1))
                    .Variation = Val(strAugment(2))
                    .Scaling = Val(strAugment(3))
                    .Done = (strAugment(4) = "True")
                End If
            End With
        End If
    Next
    StringToGearsetAugment = True
End Function

Public Function ScaledAugmentName(ptypAugment As AugmentSlotType, penColor As AugmentColorEnum, plngML As Long, pblnError As Boolean) As String
    Dim strReturn As String
    Dim i As Long
    
    pblnError = False
    With ptypAugment
        If .Augment = 0 Or .Variation = 0 Then
            strReturn = "Empty " & GetAugmentColorName(penColor) & " Slot"
        Else
            For .Scaling = db.Augment(.Augment).Scalings To 1 Step -1
                If db.Augment(.Augment).Scaling(.Scaling).ML <= plngML Then Exit For
            Next
            If .Scaling = 0 Then
                strReturn = db.Augment(.Augment).Variation(.Variation)
                pblnError = True
            Else
                strReturn = AugmentFullName(.Augment, .Variation, .Scaling)
            End If
        End If
    End With
    ScaledAugmentName = strReturn
End Function


' ************* GRID *************


Public Sub InitGrid(grid As GridType, gs As GearsetType, anal As AnalysisType, pblnFiltered As Boolean)
    Dim typBlank As GridType
    Dim lngScale As Long
    Dim i As Long
    
    grid = typBlank
    InitSlots grid, gs
    InitRows grid, gs
    FindAllEffectSlots grid, gs
    FindAllSlotEffects grid, gs
    If gs.Analyzed And pblnFiltered Then
        ApplyAnalysis gs, anal, grid
        ActivateCells grid, vleFiltered
    Else
        ActivateCells grid, vleAll
    End If
    ExplicitSelections grid, gs
    ImplicitSelections grid, gs
    grid.Initialized = True
End Sub

Private Sub InitSlots(grid As GridType, gs As GearsetType)
    Dim blnAdd As Boolean
    Dim lngCols As Long
    Dim lngCol As Long
    Dim i As Long
    
    If gs.BaseLevel < 10 Then lngCols = 2 Else lngCols = 3
    For i = 0 To seSlotCount - 1
        blnAdd = gs.Item(i).Crafted
        If blnAdd = True And i = seOffHand Then blnAdd = Not gs.TwoHanded
        If blnAdd Then
            grid.Slots = grid.Slots + 1
            ReDim Preserve grid.Slot(1 To grid.Slots)
            grid.Slot(grid.Slots).GearsetSlot = i
            grid.Slot(grid.Slots).Gear = GetSlotGear(i, gs)
            grid.Cols = grid.Cols + lngCols
            ReDim Preserve grid.Col(1 To grid.Cols)
            For lngCol = 1 To lngCols
                With grid.Col(grid.Cols - (lngCols - lngCol))
                    .Slot = grid.Slots
                    .Item = i
                    Select Case lngCol
                        Case 1
                            .Affix = aePrefix
                            .LeftThick = True
                        Case 2
                            .Affix = aeSuffix
                            .RightThick = (lngCols = 2)
                        Case 3
                            .Affix = aeExtra
                            .RightThick = True
                    End Select
                End With
            Next
        End If
    Next
End Sub

Public Function GetSlotGear(penSlot As SlotEnum, gs As GearsetType) As GearEnum
    Dim enGear As GearEnum
    
    Select Case penSlot
        Case seHelmet: enGear = geHelmet
        Case seGoggles: enGear = geGoggles
        Case seNecklace: enGear = geNecklace
        Case seCloak: enGear = geCloak
        Case seBracers: enGear = geBracers
        Case seGloves: enGear = geGloves
        Case seBelt: enGear = geBelt
        Case seBoots: enGear = geBoots
        Case seRing1, seRing2: enGear = geRing
        Case seTrinket: enGear = geTrinket
        Case seArmor
            Select Case gs.Armor
                Case ameMetal: enGear = geMetalArmor
                Case ameLeather: enGear = geLeatherArmor
                Case ameCloth: enGear = geClothArmor
                Case ameDocent: enGear = geDocent
            End Select
        Case seMainHand
            Select Case gs.Mainhand
                Case mheMelee
                    If gs.TwoHanded Then
                        If gs.Item(seMainHand).ItemStyle = "Handwraps" Then enGear = geHandwraps Else enGear = ge2hMelee
                    Else
                        enGear = ge1hMelee
                    End If
                Case mheRange
                    enGear = geRange
            End Select
        Case seOffHand
            Select Case gs.Offhand
                Case oheMelee: enGear = ge1hMelee
                Case oheShield: enGear = geShield
                Case oheOrb: enGear = geOrb
                Case oheRunearm: enGear = geRunearm
            End Select
    End Select
    GetSlotGear = enGear
End Function

Private Sub InitRows(grid As GridType, gs As GearsetType)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngScale As Long
    Dim lngItem As Long
    Dim enAffix As AffixEnum
    
    Erase grid.Row
    grid.Rows = gs.Effects
    If grid.Rows = 0 Then Exit Sub
    ReDim grid.Row(1 To grid.Rows)
    For lngRow = 1 To grid.Rows
        With grid.Row(lngRow)
            .Effect = lngRow
            .Shard = gs.Effect(lngRow)
            .Caption = db.Shard(.Shard).GridName
            lngScale = SeekScaling(db.Shard(.Shard).ScaleName)
            If lngScale Then
                .ScaleOrder = db.Scaling(lngScale).Order
                .ScaleGroup = db.Scaling(lngScale).Group
            End If
            For lngCol = 1 To grid.Cols
                lngItem = grid.Col(lngCol).Item
                enAffix = grid.Col(lngCol).Affix
                If gs.Item(lngItem).Effect(enAffix) = gs.Effect(.Effect) Then
                    .ColSelected = lngCol
                    Exit For
                End If
            Next
        End With
    Next
    SortGridRows grid
    For lngRow = 1 To grid.Rows
        If lngRow = 1 Then
            grid.Row(lngRow).TopThick = True
        ElseIf grid.Row(lngRow).ScaleGroup <> grid.Row(lngRow - 1).ScaleGroup Then
            grid.Row(lngRow).TopThick = True
        End If
        If lngRow = grid.Rows Then
            grid.Row(lngRow).BottomThick = True
        ElseIf grid.Row(lngRow).ScaleGroup <> grid.Row(lngRow + 1).ScaleGroup Then
            grid.Row(lngRow).BottomThick = True
        End If
    Next
End Sub

Private Sub SortGridRows(grid As GridType)
    Const ShrinkFactor = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As RowType
    Dim blnSwapped As Boolean
    
    iMin = 1
    iMax = grid.Rows
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If CompareGridRows(grid.Row(i), grid.Row(i + lngGap)) = 1 Then
                typSwap = grid.Row(i)
                grid.Row(i) = grid.Row(i + lngGap)
                grid.Row(i + lngGap) = typSwap
                blnSwapped = True
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

Private Function CompareGridRows(ptypLeft As RowType, ptypRight As RowType) As Long
    If ptypLeft.ScaleOrder < ptypRight.ScaleOrder Then
        CompareGridRows = -1
    ElseIf ptypLeft.ScaleOrder > ptypRight.ScaleOrder Then
        CompareGridRows = 1
    ElseIf ptypLeft.Caption < ptypRight.Caption Then
        CompareGridRows = -1
    ElseIf ptypLeft.Caption > ptypRight.Caption Then
        CompareGridRows = 1
    End If
End Function

Public Sub FindAllEffectSlots(grid As GridType, gs As GearsetType)
    Dim lngShard As Long
    Dim lngRow As Long
    
    For lngRow = 1 To grid.Rows
        lngShard = grid.Row(lngRow).Shard
'        If db.Shard(lngShard).GridName = "Healing Amp" Then Stop
        FindAffixSlot grid, lngRow, db.Shard(lngShard).Prefix, aePrefix
        FindAffixSlot grid, lngRow, db.Shard(lngShard).Suffix, aeSuffix
        FindAffixSlot grid, lngRow, db.Shard(lngShard).Extra, aeExtra
    Next
End Sub

Private Sub FindAffixSlot(grid As GridType, plngRow As Long, pblnSlot() As Boolean, penAffix As AffixEnum)
    Dim enGear As GearEnum
    Dim lngCol As Long
    
    For enGear = 0 To geGearCount - 1
'        If enGear = geLeatherArmor Then Stop
        If pblnSlot(enGear) Then
            For lngCol = 1 To grid.Cols
                If grid.Slot(grid.Col(lngCol).Slot).Gear = enGear Then
                    If grid.Col(lngCol).Affix = penAffix Then
                        With grid.Row(plngRow).Spot(vleAll)
                            .Count = .Count + 1
                            ReDim Preserve .Value(1 To .Count)
                            .Value(.Count) = lngCol
                        End With
                    End If
                End If
            Next
        End If
    Next
End Sub

Public Sub FindAllSlotEffects(grid As GridType, gs As GearsetType)
    Dim lngShard As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngItem As Long
    Dim enAffix As AffixEnum
    Dim i As Long
    
    For lngCol = 1 To grid.Cols
        lngItem = grid.Slot(grid.Col(lngCol).Slot).Gear
        enAffix = grid.Col(lngCol).Affix
        For lngRow = 1 To grid.Rows
            lngShard = grid.Row(lngRow).Shard
            Select Case enAffix
                Case aePrefix: FindSlotAffix grid, lngRow, lngCol, db.Shard(lngShard).Prefix(lngItem)
                Case aeSuffix: FindSlotAffix grid, lngRow, lngCol, db.Shard(lngShard).Suffix(lngItem)
                Case aeExtra: FindSlotAffix grid, lngRow, lngCol, db.Shard(lngShard).Extra(lngItem)
            End Select
        Next
    Next
End Sub

Private Sub FindSlotAffix(grid As GridType, plngRow As Long, plngCol As Long, pblnValid As Boolean)
    If Not pblnValid Then Exit Sub
    With grid.Col(plngCol).Effect(vleAll)
        .Count = .Count + 1
        ReDim Preserve .Value(1 To .Count)
        .Value(.Count) = plngRow
    End With
End Sub

Public Sub ApplyAnalysis(gs As GearsetType, anal As AnalysisType, grid As GridType)
    Dim lngEffect As Long
    Dim enSpot As SpotEnum
    Dim lngRow As Long
    Dim lngCol As Long
    Dim enGear As GearEnum
    Dim enAffix As AffixEnum
    
    For lngEffect = 1 To gs.Effects
        For enSpot = 0 To SpotKeyMax
            If anal.Effect(lngEffect).ValidSpot(enSpot) Then
                ' Find the row based on lngEffect
                For lngRow = 1 To grid.Rows
                    If grid.Row(lngRow).Effect = lngEffect Then Exit For
                Next
                ' Find the col based on enSpot (can be multiple hits)
                enGear = GetGearFromSpot(enSpot, gs)
                enAffix = enSpot Mod 3
                For lngCol = 1 To grid.Cols
                    If grid.Slot(grid.Col(lngCol).Slot).Gear = enGear And grid.Col(lngCol).Affix = enAffix Then AddMatch lngRow, lngCol, grid
                Next
            End If
        Next
    Next
End Sub

Private Sub AddMatch(plngRow As Long, plngCol As Long, grid As GridType)
    With grid.Row(plngRow).Spot(vleFiltered)
        .Count = .Count + 1
        ReDim Preserve .Value(1 To .Count)
        .Value(.Count) = plngCol
    End With
    With grid.Col(plngCol).Effect(vleFiltered)
        .Count = .Count + 1
        ReDim Preserve .Value(1 To .Count)
        .Value(.Count) = plngRow
    End With
End Sub

Private Function GetGearFromSpot(ByVal penSpot As SpotEnum, gs As GearsetType) As GearEnum
    penSpot = (penSpot \ 3) * 3
    Select Case penSpot
        Case spotHelmet: GetGearFromSpot = geHelmet
        Case spotGoggles: GetGearFromSpot = geGoggles
        Case spotNecklace: GetGearFromSpot = geNecklace
        Case spotCloak: GetGearFromSpot = geCloak
        Case spotBracers: GetGearFromSpot = geBracers
        Case spotGloves: GetGearFromSpot = geGloves
        Case spotBelt: GetGearFromSpot = geBelt
        Case spotBoots: GetGearFromSpot = geBoots
        Case spotRing: GetGearFromSpot = geRing
        Case spotTrinket: GetGearFromSpot = geTrinket
        Case spotMelee2H: GetGearFromSpot = ge2hMelee
        Case spotMelee1H: GetGearFromSpot = ge1hMelee
        Case spotShield: GetGearFromSpot = geShield
        Case spotRange: GetGearFromSpot = geRange
        Case spotRunearm: GetGearFromSpot = geRunearm
        Case spotOrb: GetGearFromSpot = geOrb
        Case spotArmor
            Select Case gs.Armor
                Case ameMetal: GetGearFromSpot = geMetalArmor
                Case ameLeather: GetGearFromSpot = geLeatherArmor
                Case ameCloth: GetGearFromSpot = geClothArmor
                Case ameDocent: GetGearFromSpot = geDocent
            End Select
    End Select
End Function

Private Sub ActivateCells(grid As GridType, penValues As ValueListEnum)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    If grid.Rows = 0 Or grid.Cols = 0 Then
        Erase grid.Cell
        Exit Sub
    End If
    ReDim grid.Cell(1 To grid.Rows, 1 To grid.Cols)
    For lngRow = 1 To grid.Rows
        grid.Row(lngRow).Spot(vleActive) = grid.Row(lngRow).Spot(penValues)
    Next
    For lngCol = 1 To grid.Cols
        grid.Col(lngCol).Effect(vleActive) = grid.Col(lngCol).Effect(penValues)
    Next
    ' Activation pass could be done with either Rows or Cols
    For lngRow = 1 To grid.Rows
        With grid.Row(lngRow).Spot(vleActive)
            For i = 1 To .Count
                grid.Cell(lngRow, .Value(i)).Active = True
            Next
        End With
    Next
End Sub

' Apply user selections
Private Sub ExplicitSelections(grid As GridType, gs As GearsetType)
    Dim enSlot As SlotEnum
    Dim enAffix As AffixEnum
    Dim enLastAffix As AffixEnum
    Dim lngEffect As Long
    Dim lngRow As Long
    Dim lngCol As Long
    
    If gs.BaseLevel < 10 Then enLastAffix = aeSuffix Else enLastAffix = aeExtra
    For enSlot = 0 To seSlotCount - 1
        If gs.Item(enSlot).Crafted Then
            For enAffix = aePrefix To enLastAffix
                lngEffect = gs.Item(enSlot).Effect(enAffix)
                If lngEffect Then
                    lngRow = FindRow(grid, lngEffect)
                    lngCol = FindCol(grid, enSlot, enAffix)
                    If lngRow <> 0 And lngCol <> 0 Then
                        grid.Row(lngRow).ColSelected = lngCol
                        grid.Col(lngCol).RowSelected = lngRow
                        grid.Cell(lngRow, lngCol).Selected = True
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Function FindRow(grid As GridType, plngEffect As Long) As Long
    Dim lngRow As Long
    
    For lngRow = 1 To grid.Rows
        If grid.Row(lngRow).Shard = plngEffect Then
            FindRow = lngRow
            Exit Function
        End If
    Next
End Function

Private Function FindCol(grid As GridType, penSlot As SlotEnum, penAffix As AffixEnum) As Long
    Dim lngCol As Long
    
    For lngCol = 1 To grid.Cols
        If grid.Slot(grid.Col(lngCol).Slot).GearsetSlot = penSlot And grid.Col(lngCol).Affix = penAffix Then
            FindCol = lngCol
            Exit Function
        End If
    Next
End Function

' If only one choice for a column or row, select it
Public Sub ImplicitSelections(grid As GridType, gs As GearsetType)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngValues As Long
    Dim lngValue As Long
    Dim lngChanges As Long
    Dim i As Long
    
    Do
        lngChanges = 0
        ' Rows with only one available column
        For lngRow = 1 To grid.Rows
            lngValues = 0
            If grid.Row(lngRow).ColSelected = 0 Then
                For i = 1 To grid.Row(lngRow).Spot(vleActive).Count
                    lngCol = grid.Row(lngRow).Spot(vleActive).Value(i)
                    If grid.Cell(lngRow, lngCol).Active And grid.Col(lngCol).RowSelected = 0 Then
                        lngValues = lngValues + 1
                        lngValue = lngCol
                    End If
                Next
                If lngValues = 1 Then
                    grid.Cell(lngRow, lngValue).Selected = True
                    grid.Row(lngRow).ColSelected = lngValue
                    grid.Col(lngValue).RowSelected = lngRow
                    GridSelection gs, grid, lngRow, lngValue, True
                    lngChanges = lngChanges + 1
                End If
            End If
        Next
        ' Columns with only one available row
        For lngCol = 1 To grid.Cols
            lngValues = 0
            If grid.Col(lngCol).RowSelected = 0 Then
                For i = 1 To grid.Col(lngCol).Effect(vleActive).Count
                    lngRow = grid.Col(lngCol).Effect(vleActive).Value(i)
                    If grid.Cell(lngRow, lngCol).Active And grid.Row(lngRow).ColSelected = 0 Then
                        lngValues = lngValues + 1
                        lngValue = lngRow
                    End If
                Next
                If lngValues = 1 Then
                    grid.Cell(lngValue, lngCol).Selected = True
                    grid.Row(lngValue).ColSelected = lngCol
                    grid.Col(lngCol).RowSelected = lngValue
                    GridSelection gs, grid, lngValue, lngCol, True
                    lngChanges = lngChanges + 1
                End If
            End If
        Next
    Loop Until lngChanges = 0
End Sub

Public Sub GridSelection(gs As GearsetType, grid As GridType, plngRow As Long, plngCol As Long, blnSelected As Boolean)
    Dim enSlot As SlotEnum
    Dim enAffix As AffixEnum
    Dim lngEffect As Long
    
    enSlot = grid.Slot(grid.Col(plngCol).Slot).GearsetSlot
    enAffix = grid.Col(plngCol).Affix
    If blnSelected Then lngEffect = grid.Row(plngRow).Shard
    gs.Item(enSlot).Effect(enAffix) = lngEffect
End Sub

Public Sub ClearSelections(gs As GearsetType)
    Dim enSlot As SlotEnum
    Dim enAffix As AffixEnum
    
    For enSlot = 0 To seSlotCount - 1
        For enAffix = aePrefix To aeExtra
            gs.Item(enSlot).Effect(enAffix) = 0
        Next
    Next
End Sub


' ************* SAVE/LOAD *************


Public Sub LoadItemList(pstrFile As String, gs As GearsetType)
    Dim strRaw As String
    Dim strSlot() As String
    
    If Not xp.File.Exists(pstrFile) Then Exit Sub
    strRaw = xp.File.LoadToString(pstrFile)
    ParseItemList gs, strRaw
End Sub

Public Sub SaveItemList(pstrFile As String, gs As GearsetType)
    Dim strArray() As String
    Dim lngLine As Long
    
    ReDim strArray(63)
    AddItemList gs, strArray, lngLine, False
    If UBound(strArray) <> lngLine Then ReDim Preserve strArray(lngLine)
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
    xp.File.SaveStringAs pstrFile, Join(strArray, vbNewLine)
End Sub

Public Function LoadGearset(pstrFile As String, gs As GearsetType, anal As AnalysisType) As Boolean
    Dim typBlankGearset As GearsetType
    Dim typBlankAnalysis As AnalysisType
    Dim strRaw As String
    Dim strSection() As String
    Dim lngPos As Long
    Dim i As Long
    
    If Not xp.File.Exists(pstrFile) Then Exit Function
    anal = typBlankAnalysis
    gs = typBlankGearset
    ReDim gs.Item(seSlotCount - 1)
    strRaw = xp.File.LoadToString(pstrFile)
    strSection = Split(strRaw, "Section: ")
    For i = 1 To UBound(strSection)
        lngPos = InStr(strSection(i), vbNewLine)
        If lngPos Then
            Select Case Left$(strSection(i), lngPos - 1)
                Case "General": LoadGearsetGeneral strSection(i), gs
                Case "Slots": LoadGearsetSlots strSection(i), gs
                Case "Effects": LoadGearsetEffects strSection(i), gs, anal
            End Select
        End If
    Next
    LoadGearset = True
End Function

Private Sub LoadGearsetGeneral(pstrRaw As String, gs As GearsetType)
    Dim strLine() As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim i As Long
    
    strLine = Split(pstrRaw, vbNewLine)
    For i = 1 To UBound(strLine)
        If ParseLine(strLine(i), strField, strItem, lngValue, strList, lngListMax) Then
            Select Case LCase$(strField)
                Case "baselevel"
                    gs.BaseLevel = lngValue
                Case "analyzed"
                    gs.Analyzed = (strItem = "True")
                Case "notes"
                    If Len(gs.Notes) Then gs.Notes = gs.Notes & vbNewLine
                    gs.Notes = gs.Notes & strItem
            End Select
        End If
    Next
End Sub

Private Sub LoadGearsetSlots(pstrRaw As String, gs As GearsetType)
    Dim strSlot() As String
    Dim lngSlot As Long
    Dim enSlot As SlotEnum
    Dim lngLine As Long
    Dim strLine() As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim blnHandwraps As Boolean
    Dim i As Long
    
    strSlot = Split(pstrRaw, "Slot:")
    For lngSlot = 1 To UBound(strSlot)
        enSlot = GetSlotID(Trim$(Left$(strSlot(lngSlot), InStr(strSlot(lngSlot), vbNewLine) - 1)))
        If enSlot <> seUnknown Then
            With gs.Item(enSlot)
                strLine = Split(strSlot(lngSlot), vbNewLine)
                For lngLine = 1 To UBound(strLine)
                    If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                        Select Case LCase$(strField)
                            Case "crafted": .Crafted = (strItem = "True")
                            Case "named": .Named = strItem
                            Case "armorstyle": gs.Armor = GetArmorID(strItem)
                            Case "twohanded": gs.TwoHanded = (strItem = "True")
                            Case "handwraps": blnHandwraps = (strItem = "True")
                            Case "mainhandstyle": gs.Mainhand = GetMainHandID(strItem)
                            Case "offhandstyle": gs.Offhand = GetOffHandID(strItem)
                            Case "itemname": If SeekItem(strItem) <> 0 Then .ItemStyle = strItem
                            Case "ml": .ML = lngValue
                            Case "prefix": .Effect(aePrefix) = SeekShard(strItem)
                            Case "suffix": .Effect(aeSuffix) = SeekShard(strItem)
                            Case "extra": .Effect(aeExtra) = SeekShard(strItem)
                            Case "eldritch": .EldritchRitual = SeekRitual(strItem)
                            Case "done"
                                For i = 0 To lngListMax
                                    Select Case LCase$(strList(i))
                                        Case "ml": .MLDone = True
                                        Case "prefix": .EffectDone(0) = True
                                        Case "suffix": .EffectDone(1) = True
                                        Case "extra": .EffectDone(2) = True
                                        Case "red": .Augment(aceRed).Done = True
                                        Case "orange": .Augment(aceOrange).Done = True
                                        Case "purple": .Augment(acePurple).Done = True
                                        Case "blue": .Augment(aceBlue).Done = True
                                        Case "green": .Augment(aceGreen).Done = True
                                        Case "yellow": .Augment(aceYellow).Done = True
                                        Case "colorless": .Augment(aceColorless).Done = True
                                        Case "eldritch": .EldritchDone = True
                                    End Select
                                Next
                            Case Else
                                ParseGearsetAugment gs.Item(enSlot).Augment, strField, strItem
                        End Select
                    End If
                Next
            End With
        End If
    Next
    DefaultItemNames gs, blnHandwraps
    MapSlotsToGear gs
End Sub

Private Sub ParseGearsetAugment(ptypAugment() As AugmentSlotType, pstrField As String, pstrVariant As String)
    Dim enColor As AugmentColorEnum
    Dim a As Long
    Dim v As Long
    
    enColor = GetAugmentColorID(pstrField)
    If enColor = aceAny Then Exit Sub
    ptypAugment(enColor).Exists = True
    If pstrVariant = "Empty" Then Exit Sub
    For a = 1 To db.Augments
        For v = 1 To db.Augment(a).Variations
            If db.Augment(a).Variation(v) = pstrVariant Then
                ptypAugment(enColor).Augment = a
                ptypAugment(enColor).Variation = v
            End If
        Next
    Next
End Sub

' When loading older style gearset, make intelligent guesses as to what specific weapons they were going for
Private Sub DefaultItemNames(gs As GearsetType, pblnHandwraps)
    Dim strArmor As String
    Dim strMainhand As String
    Dim strOffhand As String
    Dim blnMainhand As Boolean
    Dim blnOffhand As Boolean
    
    ' If we already have ItemStyles for armor, mainhand and offhand then no need to guess
    If Len(gs.Item(seArmor).ItemStyle) <> 0 And Len(gs.Item(seMainHand).ItemStyle) <> 0 And Len(gs.Item(seOffHand).ItemStyle) <> 0 Then Exit Sub
    ' Armor
    If Len(gs.Item(seArmor).ItemStyle) = 0 Then
        Select Case gs.Armor
            Case ameMetal: strArmor = "Chainmail"
            Case ameLeather: strArmor = "Leather"
            Case ameCloth: strArmor = "Outfit"
            Case ameDocent: strArmor = "Docent"
        End Select
    End If
    If pblnHandwraps Then
        strMainhand = "Handwraps"
        strOffhand = "Empty"
    Else
        blnMainhand = gs.Item(seMainHand).Crafted = True Or Len(gs.Item(seMainHand).Named) <> 0
        blnOffhand = gs.Item(seOffHand).Crafted = True Or Len(gs.Item(seOffHand).Named) <> 0
        ' Offhand
        If blnMainhand And Not blnOffhand Then
            strOffhand = "Empty"
        Else
            Select Case gs.Offhand
                Case oheMelee: strOffhand = "Short Sword"
                Case oheShield: If gs.Mainhand = mheMelee Then strOffhand = "Heavy Shield" Else strOffhand = "Buckler"
                Case oheOrb: strOffhand = "Orb"
                Case oheRunearm: strOffhand = "Runearm"
                Case oheEmpty: strOffhand = "Empty"
            End Select
        End If
        ' Mainhand
        Select Case gs.Mainhand
            Case mheMelee
                If gs.TwoHanded Then
                    Select Case gs.Armor
                        Case ameMetal, ameDocent: strMainhand = "Great Axe"
                        Case ameLeather, ameCloth: strMainhand = "Quarterstaff"
                    End Select
                    strOffhand = "Empty"
                Else
                    Select Case strOffhand
                        Case "Orb": strMainhand = "Scepter"
                        Case "Empty": strMainhand = "Rapier"
                        Case Else: strMainhand = db.Melee1H.Default.Choice
                    End Select
                End If
            Case mheRange
                If gs.TwoHanded Then
                    strMainhand = "Bow"
                    strOffhand = "Empty"
                Else
                    Select Case strOffhand
                        Case "Short Sword": strMainhand = "Shuriken"
                        Case "Orb", "Buckler": strMainhand = "Throwing Dagger"
                        Case Else: strMainhand = "Crossbow"
                    End Select
                End If
        End Select
    End If
    ' Final armor check now that weapons are finalized
    If blnMainhand = False And blnOffhand = False And gs.Armor = ameMetal And Len(gs.Item(seOffHand).ItemStyle) = 0 Then
        If gs.TwoHanded Then
            strMainhand = "Great Axe"
            strOffhand = "Empty"
        Else
            strMainhand = db.Melee1H.Default.Choice
            strOffhand = "Heavy Shield"
            gs.Offhand = oheShield
        End If
        strArmor = "Full Plate"
    ElseIf strMainhand = "Scepter" And strArmor = "Outfit" Then
        strArmor = "Robe"
    ElseIf gs.Armor = ameMetal And (gs.TwoHanded = True Or strOffhand = "Heavy Shield") Then
        strArmor = "Half Plate"
    End If
    ' Use our guesses if item names are otherwise blank
    With gs.Item(seMainHand)
        If Len(.ItemStyle) = 0 Then .ItemStyle = strMainhand
    End With
    With gs.Item(seOffHand)
        If Len(.ItemStyle) = 0 Then
            .ItemStyle = strOffhand
            If strOffhand = "Empty" And gs.Offhand <> oheEmpty Then gs.Offhand = oheEmpty
        End If
    End With
    With gs.Item(seArmor)
        If Len(.ItemStyle) = 0 Then .ItemStyle = strArmor
    End With
End Sub

Private Sub LoadGearsetEffects(pstrRaw As String, gs As GearsetType, anal As AnalysisType)
    Dim strEffect() As String
    Dim lngEffect As Long
    Dim enSlot As SlotEnum
    Dim strLine() As String
    Dim strField As String
    Dim strItem As String
    Dim lngValue As Long
    Dim strList() As String
    Dim lngListMax As Long
    Dim strCurrent As String
    Dim lngCurrent As Long
    Dim lngSpot As Long
    Dim enSpot As SpotEnum
    Dim lngPos As Long
    Dim i As Long
    
    strEffect = Split(pstrRaw, "Effect:")
    For lngEffect = 1 To UBound(strEffect)
        strCurrent = strEffect(lngEffect)
        lngPos = InStr(strEffect(lngEffect), vbNewLine) - 1
        If lngPos > 1 Then strCurrent = Left$(strCurrent, lngPos)
        strCurrent = Trim$(strCurrent)
        lngCurrent = SeekShard(strCurrent)
        If lngCurrent Then
            gs.Effects = gs.Effects + 1
            ReDim Preserve gs.Effect(1 To gs.Effects)
            gs.Effect(gs.Effects) = lngCurrent
            strLine = Split(strEffect(lngEffect), vbNewLine)
            For i = 1 To UBound(strLine)
                If ParseLine(strLine(i), strField, strItem, lngValue, strList, lngListMax) Then
                    Select Case LCase$(strField)
                        Case "validspots"
                            ReDim Preserve anal.Effect(gs.Effects)
                            ReDim anal.Effect(gs.Effects).ValidSpot(SpotKeyMax)
                            For lngSpot = 0 To lngListMax
                                enSpot = GetSpotID(strList(lngSpot))
                                If enSpot <> spotUnused Then anal.Effect(gs.Effects).ValidSpot(enSpot) = True
                            Next
                    End Select
                End If
            Next
        End If
    Next
End Sub

Public Sub SaveGearset(pstrFile As String, gs As GearsetType, anal As AnalysisType)
    Dim strArray() As String
    Dim strNotes() As String
    Dim lngLine As Long
    Dim i As Long
    
    ReDim strArray(255)
    ' General
    AddLine "Section: General", strArray, lngLine, 2
    AddLine "BaseLevel: " & gs.BaseLevel, strArray, lngLine
    AddLine "Analyzed: " & gs.Analyzed, strArray, lngLine
    If Len(gs.Notes) Then
        strNotes = Split(gs.Notes, vbNewLine)
        For i = 0 To UBound(strNotes)
            AddLine "Notes: " & strNotes(i), strArray, lngLine
        Next
    End If
    ' Slots
    AddLine vbNullString, strArray, lngLine, 2
    AddLine "Section: Slots", strArray, lngLine, 2
    AddItemList gs, strArray, lngLine, True
    ' Effects
    AddLine vbNullString, strArray, lngLine, 2
    AddLine "Section: Effects", strArray, lngLine, 2
    For i = 1 To gs.Effects
        AddLine "Effect: " & db.Shard(gs.Effect(i)).ShardName, strArray, lngLine
        If gs.Analyzed Then AddLine "ValidSpots: " & CreateSpotList(anal.Effect(i).ValidSpot), strArray, lngLine, 2
    Next
    ' Finish
    If UBound(strArray) <> lngLine Then ReDim Preserve strArray(lngLine)
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
    xp.File.SaveStringAs pstrFile, Join(strArray, vbNewLine)
End Sub

Private Function CreateSpotList(pblnValid() As Boolean) As String
    Dim strSpot() As String
    Dim lngSpots As Long
    Dim i As Long
    
    For i = 0 To SpotKeyMax
        If pblnValid(i) Then
            ReDim Preserve strSpot(lngSpots)
            strSpot(lngSpots) = GetSpotName(i)
            lngSpots = lngSpots + 1
        End If
    Next
    If lngSpots Then CreateSpotList = Join(strSpot, ", ")
End Function

Private Sub AddItemList(gs As GearsetType, pstrArray() As String, plngLine As Long, pblnGearset As Boolean)
    Dim enSlot As SlotEnum
    Dim lngCount As Long
    Dim strAugment As String
    Dim strColor As String
    Dim strAugmentDone As String
    Dim strDone As String
    Dim i As Long
    
    For enSlot = 0 To seSlotCount - 1
        With gs.Item(enSlot)
            AddLine "Slot: " & GetSlotName(enSlot), pstrArray, plngLine
            AddLine "Crafted: " & .Crafted, pstrArray, plngLine
            If Len(.Named) Then AddLine "Named: " & .Named, pstrArray, plngLine
            Select Case enSlot
                Case seArmor
                    AddLine "ArmorStyle: " & GetArmorName(gs.Armor), pstrArray, plngLine
                    AddLine "ItemName: " & .ItemStyle, pstrArray, plngLine
                Case seMainHand
                    AddLine "MainHandStyle: " & GetMainHandName(gs.Mainhand), pstrArray, plngLine
                    AddLine "TwoHanded: " & gs.TwoHanded, pstrArray, plngLine
                    AddLine "ItemName: " & .ItemStyle, pstrArray, plngLine
                Case seOffHand
                    AddLine "OffHandStyle: " & GetOffHandName(gs.Offhand), pstrArray, plngLine
                    AddLine "ItemName: " & .ItemStyle, pstrArray, plngLine
            End Select
            lngCount = 0
            For i = 1 To 7
                strAugmentDone = vbNullString
                With .Augment(i)
                    If .Exists = True And lngCount < 3 Then
                        lngCount = lngCount + 1
                        strColor = GetAugmentColorName(i)
                        If .Augment > 0 And .Variation > 0 Then
                            strAugment = db.Augment(.Augment).Variation(.Variation)
                            If .Done Then strAugmentDone = strAugmentDone & ", " & strColor
                        Else
                            strAugment = "Empty"
                        End If
                        AddLine strColor & ": " & strAugment, pstrArray, plngLine
                    End If
                End With
            Next
            If .EldritchRitual Then AddLine "Eldritch: " & db.Ritual(.EldritchRitual).RitualName, pstrArray, plngLine
            If pblnGearset Then
                If .ML <> 0 Then AddLine "ML: " & .ML, pstrArray, plngLine
                If .Effect(aePrefix) <> 0 Then AddLine "Prefix: " & db.Shard(.Effect(aePrefix)).ShardName, pstrArray, plngLine
                If .Effect(aeSuffix) <> 0 Then AddLine "Suffix: " & db.Shard(.Effect(aeSuffix)).ShardName, pstrArray, plngLine
                If .Effect(aeExtra) <> 0 Then AddLine "Extra: " & db.Shard(.Effect(aeExtra)).ShardName, pstrArray, plngLine
                strDone = vbNullString
                If .MLDone Then strDone = strDone & ", ML"
                If .EffectDone(0) Then strDone = strDone & ", Prefix"
                If .EffectDone(1) Then strDone = strDone & ", Suffix"
                If .EffectDone(2) Then strDone = strDone & ", Extra"
                strDone = strDone & strAugmentDone
                If .EldritchDone Then strDone = strDone & ", Eldritch"
                If Len(strDone) Then AddLine "Done: " & Mid$(strDone, 3), pstrArray, plngLine ' Chop off leading ", "
            End If
            AddLine vbNullString, pstrArray, plngLine
        End With
    Next
End Sub

Private Sub AddLine(pstrLine As String, pstrArray() As String, plngIndex As Long, Optional plngBlankLines As Long = 1)
    Dim lngMax As Long
    
    lngMax = UBound(pstrArray)
    If plngIndex > lngMax Then
        lngMax = (lngMax * 3) \ 2
        ReDim Preserve pstrArray(lngMax)
    End If
    pstrArray(plngIndex) = pstrLine
    plngIndex = plngIndex + plngBlankLines
End Sub

Private Sub ParseItemList(gs As GearsetType, pstrRaw As String)
    Dim enSlot As SlotEnum
    Dim strSlot() As String
    Dim strLine() As String
    Dim lngPos As Long
    Dim strField As String
    Dim strValue As String
    Dim blnHandwraps As Boolean
    Dim i As Long
    Dim j As Long
    
    ReDim gs.Item(seSlotCount - 1)
    gs.Armor = ameMetal
    gs.TwoHanded = False
    gs.Mainhand = mheMelee
    gs.Offhand = oheMelee
    strSlot = Split(pstrRaw, "Slot: ")
    For i = 1 To UBound(strSlot)
        strLine = Split(strSlot(i), vbNewLine)
        enSlot = GetSlotID(strLine(0))
        If enSlot <> seUnknown Then
            With gs.Item(enSlot)
                For j = 0 To UBound(strLine)
                    lngPos = InStr(strLine(j), ": ")
                    If lngPos Then
                        strField = LCase$(Left$(strLine(j), lngPos - 1))
                        strValue = Mid$(strLine(j), lngPos + 2)
                        Select Case strField
                            Case "crafted": .Crafted = (strValue = "True")
                            Case "named": .Named = strValue
                            Case "armorstyle": gs.Armor = GetArmorID(strValue)
                            Case "twohanded": gs.TwoHanded = (strValue = "True")
                            Case "handwraps": blnHandwraps = (strValue = "True")
                            Case "mainhandstyle": gs.Mainhand = GetMainHandID(strValue)
                            Case "offhandstyle": gs.Offhand = GetOffHandID(strValue)
                            Case "itemname": If SeekItem(strValue) <> 0 Then .ItemStyle = strValue
                            Case "eldritch": .EldritchRitual = SeekRitual(strValue)
                            Case Else: ParseGearsetAugment gs.Item(enSlot).Augment, strField, strValue
                        End Select
                    End If
                Next
            End With
        End If
    Next
    DefaultItemNames gs, blnHandwraps
End Sub

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
    ' Field
    lngPos = InStr(pstrLine, ":")
    If lngPos = 0 Then Exit Function
    ParseLine = True
    pstrField = LCase$(Trim$(Left$(pstrLine, lngPos - 1)))
    pstrLine = Trim$(Mid$(pstrLine, lngPos + 1))
    ' List
    If InStr(pstrLine, ",") And pstrField <> "named" And pstrField <> "notes" Then
        pstrList = Split(pstrLine, ",")
        plngListMax = UBound(pstrList)
        For i = 0 To plngListMax
            pstrList(i) = Trim$(pstrList(i))
        Next
        Exit Function
    Else
        pstrItem = pstrLine
    End If
    ' Value
    If IsNumeric(pstrItem) Then plngValue = Val(pstrItem)
    ' Return single item in list form as well
    plngListMax = 0
    ReDim pstrList(0)
    pstrList(0) = pstrItem
End Function

Public Function GetSlotName(penSlot As SlotEnum) As String
    Select Case penSlot
        Case seHelmet: GetSlotName = "Helmet"
        Case seGoggles: GetSlotName = "Goggles"
        Case seNecklace: GetSlotName = "Necklace"
        Case seCloak: GetSlotName = "Cloak"
        Case seBracers: GetSlotName = "Bracers"
        Case seGloves: GetSlotName = "Gloves"
        Case seBelt: GetSlotName = "Belt"
        Case seBoots: GetSlotName = "Boots"
        Case seRing1: GetSlotName = "Ring1"
        Case seRing2: GetSlotName = "Ring2"
        Case seTrinket: GetSlotName = "Trinket"
        Case seArmor: GetSlotName = "Armor"
        Case seMainHand: GetSlotName = "Mainhand"
        Case seOffHand: GetSlotName = "Offhand"
    End Select
End Function

Private Function GetSlotID(pstrSlot As String) As SlotEnum
    Select Case LCase$(pstrSlot)
        Case "helmet", "helm": GetSlotID = seHelmet
        Case "goggles": GetSlotID = seGoggles
        Case "necklace": GetSlotID = seNecklace
        Case "cloak": GetSlotID = seCloak
        Case "bracers": GetSlotID = seBracers
        Case "gloves": GetSlotID = seGloves
        Case "belt": GetSlotID = seBelt
        Case "boots": GetSlotID = seBoots
        Case "ring1": GetSlotID = seRing1
        Case "ring2": GetSlotID = seRing2
        Case "trinket": GetSlotID = seTrinket
        Case "armor": GetSlotID = seArmor
        Case "mainhand": GetSlotID = seMainHand
        Case "offhand": GetSlotID = seOffHand
        Case Else: GetSlotID = seUnknown
    End Select
End Function

Private Function GetArmorName(penArmor As ArmorMaterialEnum) As String
    Select Case penArmor
        Case ameMetal: GetArmorName = "Metal"
        Case ameLeather: GetArmorName = "Leather"
        Case ameCloth: GetArmorName = "Cloth"
        Case ameDocent: GetArmorName = "Docent"
    End Select
End Function

Private Function GetArmorID(pstrArmor As String) As ArmorMaterialEnum
    Select Case LCase$(pstrArmor)
        Case "metal": GetArmorID = ameMetal
        Case "leather": GetArmorID = ameLeather
        Case "cloth": GetArmorID = ameCloth
        Case "docent": GetArmorID = ameDocent
    End Select
End Function

Private Function GetMainHandName(penMainHand As MainHandEnum) As String
    Select Case penMainHand
        Case mheMelee: GetMainHandName = "Melee"
        Case mheRange: GetMainHandName = "Range"
    End Select
End Function

Private Function GetMainHandID(pstrMainHand As String) As MainHandEnum
    Select Case LCase$(pstrMainHand)
        Case "melee": GetMainHandID = mheMelee
        Case "range": GetMainHandID = mheRange
    End Select
End Function

Private Function GetOffHandName(penOffHand As OffHandEnum) As String
    Select Case penOffHand
        Case oheMelee: GetOffHandName = "Melee"
        Case oheShield: GetOffHandName = "Shield"
        Case oheOrb: GetOffHandName = "Orb"
        Case oheRunearm: GetOffHandName = "Runearm"
        Case oheEmpty: GetOffHandName = "Empty"
    End Select
End Function

Private Function GetOffHandID(pstrOffHand As String) As OffHandEnum
    Select Case LCase$(pstrOffHand)
        Case "melee": GetOffHandID = oheMelee
        Case "shield": GetOffHandID = oheShield
        Case "orb": GetOffHandID = oheOrb
        Case "runearm": GetOffHandID = oheRunearm
        Case "empty": GetOffHandID = oheEmpty
    End Select
End Function

Private Function GetSpotName(penSpot As SpotEnum) As String
    Dim enSpot As SpotEnum
    Dim enAffix As AffixEnum
    Dim strGear As String
    Dim strAffix As String

    If penSpot = spotUnused Then Exit Function
    enAffix = penSpot Mod 3
    enSpot = penSpot - enAffix
    Select Case enSpot
        Case spotHelmet: strGear = "Helmet"
        Case spotGoggles: strGear = "Goggles"
        Case spotNecklace: strGear = "Necklace"
        Case spotCloak: strGear = "Cloak"
        Case spotBracers: strGear = "Bracers"
        Case spotGloves: strGear = "Gloves"
        Case spotBelt: strGear = "Belt"
        Case spotBoots: strGear = "Boots"
        Case spotRing: strGear = "Ring"
        Case spotTrinket: strGear = "Trinket"
        Case spotArmor: strGear = "Armor"
        Case spotMelee2H: strGear = "Melee2H"
        Case spotMelee1H: strGear = "Melee1H"
        Case spotShield: strGear = "Shield"
        Case spotRange: strGear = "Range"
        Case spotRunearm: strGear = "Runearm"
        Case spotOrb: strGear = "Orb"
    End Select
    Select Case enAffix
        Case aePrefix: strAffix = "Prefix"
        Case aeSuffix: strAffix = "Suffix"
        Case aeExtra: strAffix = "Extra"
    End Select
    GetSpotName = strGear & " " & strAffix
End Function

Private Function GetSpotID(pstrSpot As String) As SpotEnum
    Dim enSpot As SpotEnum
    Dim enAffix As AffixEnum
    Dim strToken() As String
    
    If Len(pstrSpot) = 0 Then
        GetSpotID = spotUnused
        Exit Function
    End If
    strToken = Split(pstrSpot, " ")
    Select Case strToken(0)
        Case "Helmet": enSpot = spotHelmet
        Case "Goggles": enSpot = spotGoggles
        Case "Necklace": enSpot = spotNecklace
        Case "Cloak": enSpot = spotCloak
        Case "Bracers": enSpot = spotBracers
        Case "Gloves": enSpot = spotGloves
        Case "Belt": enSpot = spotBelt
        Case "Boots": enSpot = spotBoots
        Case "Ring": enSpot = spotRing
        Case "Trinket": enSpot = spotTrinket
        Case "Armor": enSpot = spotArmor
        Case "Melee2H": enSpot = spotMelee2H
        Case "Melee1H": enSpot = spotMelee1H
        Case "Shield": enSpot = spotShield
        Case "Range": enSpot = spotRange
        Case "Runearm": enSpot = spotRunearm
        Case "Orb": enSpot = spotOrb
    End Select
    If UBound(strToken) = 1 Then
        Select Case strToken(1)
            Case "Prefix": enAffix = aePrefix
            Case "Suffix": enAffix = aeSuffix
            Case "Extra": enAffix = aeExtra
        End Select
    End If
    GetSpotID = enSpot + enAffix
End Function
