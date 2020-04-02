Attribute VB_Name = "basDeprecate"
' I don't think I'm technically using the word "Deprecate" properly
' But it's the closest word to what this is trying to do
Option Explicit

Public Type DeprecateFeatType
    Type As BuildFeatTypeEnum
    Index As Long
    FeatName As String
    SelectorName As String
    Selector As Long
    Child As BuildFeatTypeEnum
    ChildIndex As Long
    Current As String ' Display text as created now
    Old As String ' Display text as stored in build file
End Type

Public Type DeprecateAbilityType
    PointerType As PointerEnum
    Raw As String
    TreeID As Long
    Ability As Long
    Selector As Long
    Tier As Long
    Rank As Long
    TreeName As String
    AbilityName As String
    SelectorName As String
    GuideTreeID As Long
    GuideTreeDisplay As String
    Parsed As Boolean
    Deprecated As Boolean
    Binary As Boolean
End Type

Public Type DeprecateType
    Deprecated As Boolean
    SaveDeprecations As Boolean
    Feats As Long
    Feat() As DeprecateFeatType
    NameChange() As Long
    NameChanges As Long
    WarSoul As Boolean
    DivineMight As Boolean
    Enhancements As Long
    Enhancement() As DeprecateAbilityType
    Trees As Long
    Tree() As String
    LevelingGuide As DeprecateAbilityType
    BinaryDestiny As String
    Destiny() As DeprecateAbilityType
    Destinies As Long
    Twists As Long
    Twist() As DeprecateAbilityType
    BinaryTwists As Long
    BinaryTwist() As TwistType
End Type

Public gtypDeprecate As DeprecateType


' ************* DEPRECATE *************


Public Sub DeprecateInit(Optional pblnSaveDeprecations As Boolean = False)
    Dim typBlank As DeprecateType
    
    gtypDeprecate = typBlank
    gtypDeprecate.SaveDeprecations = pblnSaveDeprecations
End Sub

Public Sub DeprecateNameChange(plngIndex As Long)
    Dim blnDone
    Dim i As Long
    
    With gtypDeprecate
        For i = 1 To .NameChanges
            If .NameChange(i) = plngIndex Then blnDone = True
        Next
        If Not blnDone Then
            .NameChanges = .NameChanges + 1
            ReDim Preserve .NameChange(1 To .NameChanges)
            .NameChange(.NameChanges) = plngIndex
            .Deprecated = True
        End If
    End With
End Sub

Public Function DeprecateGetFeatID(pstrFeat As String) As Long
    Dim lngNameChange As Long
    Dim lngFeat As Long
    
    lngFeat = SeekFeat(pstrFeat, False)
    If lngFeat = 0 Then
        lngNameChange = SeekNameChange("Feat", pstrFeat)
        If lngNameChange Then lngFeat = SeekFeat(db.NameChange(lngNameChange).NewName)
        If lngFeat Then DeprecateNameChange lngNameChange
    End If
    DeprecateGetFeatID = lngFeat
End Function

Public Sub ShowDeprecated()
    Dim frm As Form
    
    If gtypDeprecate.Deprecated Then
        OpenForm "frmDeprecate"
    ElseIf GetForm(frm, "frmDeprecate") Then
        Unload frmDeprecate
    End If
End Sub


' ************* BINARY *************


Public Sub DeprecateBinary()
    DeprecateInit True
    DeprecateBinaryFeats
    WarpriestToWarSoul
    DeprecateBinaryEnhancements
    DeprecateBinaryLevelingGuide
    DeprecateBinaryDestiny
    DeprecateBinaryTwists
End Sub

Private Sub DeprecateBinaryFeats()
    Dim enType As BuildFeatTypeEnum
    Dim lngFeat As Long
    Dim blnDeprecate As Boolean
    Dim i As Long
    
    For enType = bftStandard To bftFeatTypes - 1
        With build.Feat(enType)
            For i = 1 To .Feats
                With .Feat(i)
                    If Len(.FeatName) > 0 And .Level <= build.MaxLevels Then
                        lngFeat = DeprecateGetFeatID(.FeatName)
                        If lngFeat = 0 Then
                            blnDeprecate = DeprecateBinaryFeat(enType, i)
                        ElseIf (db.Feat(lngFeat).Selectors > 0 And .Selector = 0) Or db.Feat(lngFeat).Selectors < .Selector Then
                            blnDeprecate = DeprecateBinaryFeat(enType, i)
                        Else
                            blnDeprecate = False
                        End If
                        If blnDeprecate Then
                            .FeatName = vbNullString
                            .Selector = 0
                        Else
                            .FeatName = db.Feat(lngFeat).FeatName
                        End If
                    End If
                End With
            Next
        End With
    Next
End Sub

Private Function DeprecateBinaryFeat(penType As BuildFeatTypeEnum, plngIndex As Long) As Boolean
    Dim typNew As DeprecateFeatType
    
    With build.Feat(penType).Feat(plngIndex)
        typNew.Type = penType
        typNew.Index = plngIndex
        typNew.FeatName = .FeatName
        typNew.Selector = .Selector
        typNew.Old = .FeatName
        If .Selector > 0 Then typNew.Old = typNew.Old & ": Selector " & .Selector
    End With
    With gtypDeprecate
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats) = typNew
        .Deprecated = True
    End With
    DeprecateBinaryFeat = True
End Function

' This is only required for binary builds, because the change to War Soul happened in the same update that
' introduced text builds. So it's not possible for a FVS text build to point to Warpriest instead of War Soul.
Private Sub WarpriestToWarSoul()
    Dim lngTree As Long
    Dim blnDone As Boolean
    Dim i As Long
    
    ' Enhancements
    For lngTree = 1 To build.Trees
        If build.Tree(lngTree).Source = ceFavoredSoul And build.Tree(lngTree).TreeName = "Warpriest" Then
            build.Tree(lngTree).TreeName = "War Soul"
            gtypDeprecate.WarSoul = True
            For i = 1 To build.Tree(lngTree).Abilities
                With build.Tree(lngTree).Ability(i)
                    If .Tier = 1 And .Ability = 1 Then
                        If .Selector = 0 Then
                            .Selector = 1
                            gtypDeprecate.DivineMight = True
                        End If
                        blnDone = True
                    End If
                End With
                If blnDone Then Exit For
            Next
            Exit For
        End If
    Next
    ' Leveling Guide
    blnDone = False
    For lngTree = 1 To build.Guide.Trees
        If build.Guide.Tree(lngTree).Class = ceFavoredSoul And build.Guide.Tree(lngTree).TreeName = "Warpriest" Then
            build.Guide.Tree(lngTree).TreeName = "War Soul"
            gtypDeprecate.WarSoul = True
            For i = 1 To build.Guide.Enhancements
                With build.Guide.Enhancement(i)
                    If .ID = lngTree And .Tier = 1 And .Ability = 1 Then
                        If .Selector = 0 Then
                            .Selector = 1
                            gtypDeprecate.DivineMight = True
                        End If
                        blnDone = True
                    End If
                End With
            Next
        End If
    Next
    With gtypDeprecate
        If .WarSoul Or .DivineMight Then .Deprecated = True
    End With
End Sub

Private Sub DeprecateBinaryEnhancements()
    Dim lngBuildTree As Long
    Dim strTree As String
    Dim lngTree As Long
    Dim blnDeprecate As Boolean
    Dim i As Long
    
    For lngBuildTree = build.Trees To 1 Step -1
        blnDeprecate = False
        With build.Tree(lngBuildTree)
            strTree = .TreeName
            lngTree = SeekTree(strTree, peEnhancement)
            If lngTree = 0 Then
                blnDeprecate = True
            ElseIf DeprecateBinaryTree(build.Tree(lngBuildTree), db.Tree(lngTree)) Then
                blnDeprecate = True
            End If
        End With
        If blnDeprecate Then
            With gtypDeprecate
                .Trees = .Trees + 1
                ReDim Preserve .Tree(1 To .Trees)
                .Tree(.Trees) = strTree
                .Deprecated = True
            End With
            With build
                For i = lngBuildTree To build.Trees - 1
                    .Tree(i) = .Tree(i + 1)
                Next
                .Trees = .Trees - 1
                If .Trees = 0 Then Erase .Tree Else ReDim Preserve .Tree(1 To .Trees)
            End With
        End If
    Next
End Sub

Private Sub DeprecateBinaryDestiny()
    Dim lngDestiny As Long
    Dim blnDeprecate As Boolean
    Dim i As Long
    
    If Len(build.Destiny.TreeName) = 0 Then Exit Sub
    lngDestiny = SeekTree(build.Destiny.TreeName, peDestiny)
    If lngDestiny = 0 Then
        blnDeprecate = True
    ElseIf DeprecateBinaryTree(build.Destiny, db.Destiny(lngDestiny)) Then
        blnDeprecate = True
    End If
    If blnDeprecate Then
        gtypDeprecate.BinaryDestiny = build.Destiny.TreeName
        build.Destiny.TreeName = vbNullString
        Erase build.Destiny.Ability
        build.Destiny.Abilities = 0
        gtypDeprecate.Deprecated = True
    End If
End Sub

' Return TRUE if tree needs to be deprecated
Private Function DeprecateBinaryTree(ptypBuildTree As BuildTreeType, ptypTree As TreeType) As Boolean
    Dim i As Long
    
    For i = 1 To ptypBuildTree.Abilities
        If Not ValidAbility(ptypTree, ptypBuildTree.Ability(i)) Then
            DeprecateBinaryTree = True
            Exit Function
        End If
    Next
End Function

' Return TRUE if ability is valid
Private Function ValidAbility(ptypTree As TreeType, ptypAbility As BuildAbilityType) As Boolean
On Error GoTo ValidAbilityErr
    If ptypTree.Tiers < ptypAbility.Tier Then Exit Function
    If ptypTree.Tier(ptypAbility.Tier).Abilities < ptypAbility.Ability Then Exit Function
    If ptypTree.Tier(ptypAbility.Tier).Ability(ptypAbility.Ability).Ranks < ptypAbility.Rank Then Exit Function
    If ptypTree.Tier(ptypAbility.Tier).Ability(ptypAbility.Ability).Selectors < ptypAbility.Selector Then Exit Function
    If ptypTree.Tier(ptypAbility.Tier).Ability(ptypAbility.Ability).Selectors > 0 And ptypAbility.Selector = 0 Then Exit Function
    ValidAbility = True
    
ValidAbilityExit:
    Exit Function
    
ValidAbilityErr:
    Resume ValidAbilityExit
End Function

Private Sub DeprecateBinaryLevelingGuide()
    Dim lngGuideTree As Long
    Dim lngTree As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRank As Long
    Dim strTree As String
    Dim blnDeprecate As Boolean
    Dim typBuild As BuildGuideType
    Dim typGuide As GuideType
    Dim typAbility As DeprecateAbilityType
    Dim i As Long
    
    For i = 1 To build.Guide.Enhancements
        With build.Guide.Enhancement(i)
            lngGuideTree = .ID
            lngTier = .Tier
            lngAbility = .Ability
            lngSelector = .Selector
            lngRank = .Rank
        End With
        If lngGuideTree > 0 And lngGuideTree < 100 Then
            blnDeprecate = True
            strTree = "Tree " & lngGuideTree
            If lngGuideTree > build.Guide.Trees Then Exit For
            strTree = build.Guide.Tree(lngGuideTree).TreeName
            lngTree = SeekTree(strTree, peEnhancement)
            If lngTree = 0 Then Exit For
            If lngTier > db.Tree(lngTree).Tiers Then Exit For
            If lngAbility > db.Tree(lngTree).Tier(lngTier).Abilities Then Exit For
            If lngRank > db.Tree(lngTree).Tier(lngTier).Ability(lngAbility).Ranks Then Exit For
            If lngSelector = 0 Then
                If db.Tree(lngTree).Tier(lngTier).Ability(lngAbility).SelectorStyle <> sseNone Then Exit For
            Else
                If lngSelector > db.Tree(lngTree).Tier(lngTier).Ability(lngAbility).Selectors Then Exit For
            End If
            blnDeprecate = False
        End If
    Next
    If blnDeprecate Then
        With gtypDeprecate.LevelingGuide
            .Deprecated = True
            .Binary = True
            .GuideTreeDisplay = strTree
            .TreeName = strTree
            .Tier = lngTier
            .AbilityName = "Ability " & lngAbility
            .SelectorName = "Selector " & lngSelector
            .Rank = lngRank
        End With
        gtypDeprecate.Deprecated = True
        i = i - 1
'        If i = 0 Then
            build.Guide = typBuild
            Guide = typGuide
'        Else
'            build.Guide.Enhancements = i
'            ReDim Preserve build.Guide.Enhancement(1 To i)
'            Guide.Enhancements = i
'            ReDim Preserve Guide.Enhancement(i)
'        End If
    End If
End Sub

Private Sub DeprecateBinaryTwists()
    Dim typBlank As TwistType
    Dim lngDestiny As Long
    Dim blnError As Boolean
    Dim i As Long
    
    For i = build.Twists To 1 Step -1
        With build.Twist(i)
            blnError = False
            If .Tier = 0 Or .Ability = 0 Or Len(.DestinyName) = 0 Then
                ' Do nothing
            ElseIf .Tier > 4 Then
                blnError = True
            Else
                lngDestiny = SeekTree(.DestinyName, peDestiny)
                If lngDestiny = 0 Then
                    blnError = True
                ElseIf .Ability > db.Destiny(lngDestiny).Tier(.Tier).Abilities Then
                    blnError = True
                ElseIf .Selector > db.Destiny(lngDestiny).Tier(.Tier).Ability(.Ability).Selectors Then
                    blnError = True
                ElseIf .Selector = 0 And db.Destiny(lngDestiny).Tier(.Tier).Ability(.Ability).Selectors > 0 Then
                    blnError = True
                End If
            End If
        End With
        If blnError Then
            With gtypDeprecate
                .BinaryTwists = .BinaryTwists + 1
                ReDim Preserve .BinaryTwist(1 To .BinaryTwists)
                .BinaryTwist(.BinaryTwists) = build.Twist(i)
                .Deprecated = True
            End With
            RemoveTwist i
        End If
    Next
End Sub

Private Sub RemoveTwist(plngTwist As Long)
    Dim i As Long
    
    With build
        For i = plngTwist To build.Twists - 1
            .Twist(i) = .Twist(i + 1)
        Next
        .Twists = .Twists - 1
        If .Twists = 0 Then Erase .Twist Else ReDim Preserve .Twist(1 To .Twists)
    End With
End Sub


' ************* FEATS *************


Public Sub DeprecateFeat(penType As BuildFeatTypeEnum, plngIndex As Long, pstrFeat As String, pstrSelector As String, penChild As BuildFeatTypeEnum, plngChildIndex As Long)
    Dim typNew As DeprecateFeatType
    
    With build.Feat(penType).Feat(plngIndex)
        typNew.Type = penType
        typNew.Index = plngIndex
        typNew.FeatName = pstrFeat
        typNew.SelectorName = pstrSelector
        typNew.Child = penChild
        typNew.ChildIndex = plngChildIndex
    End With
    With gtypDeprecate
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats) = typNew
        .Deprecated = True
    End With
End Sub

Public Sub DeprecateFeatSelector(penType As BuildFeatTypeEnum, plngIndex As Long, pstrDescrip As String)
    Dim typNew As DeprecateFeatType
    
    With build.Feat(penType).Feat(plngIndex)
        typNew.Type = penType
        typNew.Index = plngIndex
        typNew.Old = pstrDescrip
    End With
    With gtypDeprecate
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats) = typNew
        .Deprecated = True
    End With
End Sub


' ************* ABILITIES *************


Public Sub DeprecateAbility(ptypAbility As DeprecateAbilityType)
    With gtypDeprecate
        .Deprecated = True
        Select Case ptypAbility.PointerType
            Case peEnhancement
                .Enhancements = .Enhancements + 1
                ReDim Preserve .Enhancement(1 To .Enhancements)
                .Enhancement(.Enhancements) = ptypAbility
            Case peDestiny
                .Destinies = .Destinies + 1
                ReDim Preserve .Destiny(1 To .Destinies)
                .Destiny(.Destinies) = ptypAbility
        End Select
    End With
End Sub

Public Sub DeprecateLevelingGuide(ptypAbility As DeprecateAbilityType)
    With gtypDeprecate
        .LevelingGuide = ptypAbility
        .Deprecated = True
    End With
End Sub

Public Sub DeprecateTwist(ptypAbility As DeprecateAbilityType)
    With gtypDeprecate
        .Twists = .Twists + 1
        ReDim Preserve .Twist(1 To .Twists)
        .Twist(.Twists) = ptypAbility
        .Deprecated = True
    End With
End Sub
