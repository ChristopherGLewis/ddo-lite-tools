Attribute VB_Name = "basProcess"
' Written by Ellis Dee
' Propagate selector lists, and convert references to feats/selectors/trees/abilities from text to "pointers"
' Also monitor data integrity, generating errors for any invalid pointers
Option Explicit


Public Sub ProcessData()
    ProcessRacialTrees
    ProcessClassTrees
    ProcessFeatSelectors
    ProcessAbilitySelectors
    ProcessPointers
    ProcessSpells
    ProcessTemplates
    ProcessFeatMap
    ProcessTreeLockouts
End Sub


' ************* SPELLS *************


Private Sub ProcessSpells()
    Dim i As Long
    
    log.Activity = actProcessClassSpells
    log.LoadFile = "Classes.txt"
    For log.Class = 1 To ceClasses - 1
        log.Level = 0
        With db.Class(log.Class)
            log.LoadSpellType = "Mandatory Spell"
            ProcessSpellList .MandatorySpell, .MandatorySpells
            log.LoadSpellType = "Free Spell"
            ProcessSpellList .FreeSpell, .FreeSpells
            log.LoadSpellType = "Pact Spell"
            For i = 1 To .Pacts
                With .Pact(i)
                    ProcessSpellList .Spells, UBound(.Spells)
                End With
            Next
            For log.Level = 1 To .MaxSpellLevel
                With .SpellList(log.Level)
                    ProcessSpellList .Spell, .Spells
                End With
            Next
        End With
    Next
End Sub

Private Sub ProcessSpellList(pstrSpell() As String, ByVal plngLast As Long)
    Dim i As Long
    
    For i = 1 To plngLast
        log.LoadSpell = pstrSpell(i)
        If SeekSpell(log.LoadSpell) = 0 Then LogError
    Next
End Sub


' ************* FEATMAP *************

Private Sub ProcessFeatMap()
    Dim lngPos As Long
    Dim strFeat As String
    Dim strSelector As String
    Dim i As Long
    
    log.LoadFile = "FeatMap.txt"
    log.Activity = actProcessFeatMap
    For i = 1 To db.FeatMaps
        log.LoadItem = db.FeatMap(i).Lite
        If Len(log.LoadItem) Then
            lngPos = InStr(log.LoadItem, ": ")
            If lngPos = 0 Then
                log.LoadSelector = vbNullString
            Else
                log.LoadSelector = Mid$(log.LoadItem, lngPos + 2)
                log.LoadItem = Left$(log.LoadItem, lngPos - 1)
            End If
            log.Feat = SeekFeat(log.LoadItem)
            If log.Feat = 0 Then
                LogError
            ElseIf Len(log.LoadSelector) Then
                For log.Selector = 1 To db.Feat(log.Feat).Selectors
                    If db.Feat(log.Feat).Selector(log.Selector).SelectorName = log.LoadSelector Then Exit For
                Next
                If log.Selector > db.Feat(log.Feat).Selectors Then LogError
            End If
        End If
    Next
End Sub


' ************* TEMPLATES *************


Public Sub ProcessTemplates()
    Dim lngPoints As Long
    Dim enStat As StatEnum
    
    log.LoadFile = "Templates.txt"
    For log.Template = 1 To db.Templates
        For log.Points = 0 To 4
            log.Activity = actTemplateStats
            log.Total = 0
            For log.Stat = 1 To 6
                lngPoints = db.Template(log.Template).StatPoints(log.Points, log.Stat)
                Select Case lngPoints
                    Case 0 To 6, 8, 10, 13, 16
                    Case Else: LogError
                End Select
                log.Total = log.Total + lngPoints
            Next
            log.Activity = actTemplatePoints
            If log.Total <> 28 + (log.Points * 2) Then LogError
        Next
    Next
End Sub


' ************* TREES *************


Private Sub ProcessRacialTrees()
    Dim lngTree As Long
    
    log.Activity = actProcessRacialTrees
    For log.Race = 1 To reRaces - 1
        ' Race tree
        log.Tree = 0
        lngTree = SeekTree(GetRaceName(log.Race), peEnhancement)
        If lngTree = 0 Then
            LogError
        Else
            If db.Tree(lngTree).TreeType <> tseRace Then LogError
        End If
        ' RaceClass tree
        For log.Tree = 1 To db.Race(log.Race).Trees
            lngTree = SeekTree(db.Race(log.Race).Tree(log.Tree), peEnhancement)
            If lngTree = 0 Then
                LogError
            Else
                If db.Tree(lngTree).TreeType <> tseRaceClass Then LogError
            End If
        Next
    Next
End Sub

Private Sub ProcessClassTrees()
    Dim lngTree As Long
    
    log.Activity = actProcessClassTrees
    For log.Class = 1 To ceClasses - 1
        For log.Tree = 1 To db.Class(log.Class).Trees
            lngTree = SeekTree(db.Class(log.Class).Tree(log.Tree), peEnhancement)
            If lngTree = 0 Then
                LogError
            Else
                If db.Tree(lngTree).TreeType <> tseClass Then LogError
            End If
        Next
    Next
End Sub

Private Sub ProcessTreeLockouts()
    Dim lngTree As Long
    Dim i As Long
    
    log.Activity = actProcessTreeLockouts
    For log.Tree = 1 To db.Trees
        If Len(db.Tree(log.Tree).Lockout) Then
            lngTree = SeekTree(db.Tree(log.Tree).Lockout, peEnhancement)
            If lngTree = 0 Then
                LogError
            Else
                Select Case db.Tree(lngTree).TreeType
                    Case tseRace, tseGlobal, tseDestiny: LogError
                End Select
            End If
        End If
    Next
End Sub


' ************* SELECTORS *************


Private Sub ProcessFeatSelectors()
    Dim lngSelector As Long
    Dim lngParent As Long
    
    log.Activity = actProcessFeatSelectors
    log.Style = peFeat
    For log.Feat = 1 To db.Feats
        log.HasError = False
        With db.Feat(log.Feat)
            If .SelectorStyle = sseShared Or .SelectorStyle = sseExclusive Then
                lngParent = log.Feat
                Do
                    lngParent = SeekFeat(db.Feat(lngParent).Parent.Raw)
                    If lngParent = 0 Then
                        LogError
                        Exit Do
                    End If
                    If .Parent.Feat = 0 Then .Parent.Feat = lngParent
                Loop While db.Feat(lngParent).Selectors = 0
                If Not log.HasError Then
                    .Parent.Style = peFeat
                    If .Selectors = 0 Then
                        .Selectors = db.Feat(lngParent).Selectors
                        ReDim .Selector(1 To .Selectors)
                        For lngSelector = 1 To .Selectors
                            .Selector(lngSelector).SelectorName = db.Feat(lngParent).Selector(lngSelector).SelectorName
                            .Selector(lngSelector).ClassBonus = .ClassBonus
                            .Selector(lngSelector).Race = .Race
                            .Selector(lngSelector).Req = .Req
                            .Selector(lngSelector).Skill = .Skill
                            .Selector(lngSelector).SkillValue = .SkillValue
                            .Selector(lngSelector).Stat = .Stat
                            .Selector(lngSelector).StatValue = .StatValue
                            ReDim .Selector(lngSelector).Class(ceClasses - 1)
                            ReDim .Selector(lngSelector).ClassLevel(ceClasses - 1)
                        Next
                    End If
                End If
            End If
        End With
    Next
End Sub

Private Sub ProcessAbilitySelectors()
    log.Activity = actProcessEnhancementSelectors
    log.Style = peEnhancement
    For log.Tree = 1 To db.Trees
        ProcessTreeSelectors db.Tree(log.Tree)
    Next
    log.Activity = actProcessDestinySelectors
    log.Style = peDestiny
    For log.Tree = 1 To db.Destinies
        ProcessTreeSelectors db.Destiny(log.Tree)
    Next
End Sub

Private Sub ProcessTreeSelectors(ptypTree As TreeType)
    Dim lngSelector As Long
    Dim lngParent As Long
    
    For log.Tier = 0 To ptypTree.Tiers
        For log.Ability = 1 To ptypTree.Tier(log.Tier).Abilities
            log.HasError = False
            With ptypTree.Tier(log.Tier).Ability(log.Ability)
                Select Case .SelectorStyle
                    Case sseShared, sseExclusive
                        If .Selectors = 0 Then
                            If Left$(.Parent.Raw, 5) = "Feat:" Then
                                log.ptr = .Parent
                                log.Style = peFeat
                                ' Feats (eg: Magister [School] Specialist)
                                lngParent = SeekFeat(Mid$(.Parent.Raw, 7))
                                If lngParent = 0 Then
                                    LogError
                                Else
                                    .Parent.Style = peFeat
                                    .Parent.Feat = lngParent
                                    .Selectors = db.Feat(lngParent).Selectors
                                    ReDim .Selector(1 To .Selectors)
                                    For lngSelector = 1 To .Selectors
                                        .Selector(lngSelector).SelectorName = db.Feat(lngParent).Selector(lngSelector).SelectorName
                                        .Selector(lngSelector).Req = db.Feat(lngParent).Selector(lngSelector).Req
                                        .Selector(lngSelector).Cost = .Cost
                                    Next
                                End If
                            Else
                                log.ptr = .Parent
                                If ptypTree.TreeType = tseDestiny Then log.Style = peDestiny Else log.Style = peEnhancement
                                ' Enhancements / Destinies
                                ProcessTreeSelectorParent ptypTree, .Parent
                                If .Parent.Ability = 0 Then
                                    LogError
                                Else
                                    .Selectors = ptypTree.Tier(.Parent.Tier).Ability(.Parent.Ability).Selectors
                                    If .Selectors > 0 Then
                                        ReDim .Selector(1 To .Selectors)
                                        For lngSelector = 1 To .Selectors
                                            .Selector(lngSelector).SelectorName = ptypTree.Tier(.Parent.Tier).Ability(.Parent.Ability).Selector(lngSelector).SelectorName
                                            .Selector(lngSelector).Req = .Req
                                            .Selector(lngSelector).Cost = .Cost
                                        Next
                                    End If
                                End If
                            End If
                        End If
                End Select
            End With
        Next
    Next
End Sub

' Parent selectors in trees can only come from within the same tree (or feats, of course)
Private Sub ProcessTreeSelectorParent(ptypTree As TreeType, ptypPointer As PointerType)
    Dim strRaw As String
    Dim lngPos As Long
    Dim i As Long
    
    log.HasError = False
    If Left$(ptypPointer.Raw, 5) <> "Tier " Then LogError
    If Not log.HasError Then
        strRaw = Mid$(ptypPointer.Raw, 6)
        lngPos = InStr(strRaw, ":")
        If lngPos = 0 Then LogError
    End If
    If Not log.HasError Then
        With ptypPointer
            .Style = log.Style
            .Tree = log.Tree
            .Tier = val(Left$(strRaw, lngPos - 1))
            strRaw = Mid$(strRaw, lngPos + 2)
            With ptypTree.Tier(.Tier)
                For i = 1 To .Abilities
                    If .Ability(i).AbilityName = strRaw Then Exit For
                Next
                If i > .Abilities Then
                    LogError
                Else
                    ptypPointer.Ability = i
                End If
            End With
        End With
    End If
End Sub


' ************* POINTERS *************


Private Sub ProcessPointers()
    Dim i As Long
    
    log.Style = peFeat
    ' Races
    log.Activity = actProcessRaceGrantedFeats
    For log.Race = 1 To reRaces - 1
        For log.Feat = 1 To db.Race(log.Race).GrantedFeats
            ProcessPointer db.Race(log.Race).GrantedFeat(log.Feat)
        Next
    Next
    ' Classes
    log.Activity = actProcessClassGrantedFeats
    For log.Class = 1 To ceClasses - 1
        For log.Feat = 1 To db.Class(log.Class).GrantedFeats
            ProcessPointer db.Class(log.Class).GrantedFeat(log.Feat)
        Next
    Next
    ' Feats
    log.Activity = actProcessFeatReqs
    For log.Feat = 1 To db.Feats
        ProcessReqs db.Feat(log.Feat).Req
        For log.Selector = 1 To db.Feat(log.Feat).Selectors
            ProcessReqs db.Feat(log.Feat).Selector(log.Selector).Req
        Next
        log.Selector = 0
    Next
    ' Trees
    log.Activity = actProcessEnhancementReqs
    log.Style = peEnhancement
    For log.Tree = 1 To db.Trees
        ProcessTree db.Tree(log.Tree)
    Next
    ' Destinies
    log.Activity = actProcessDestinyReqs
    log.Style = peDestiny
    For log.Tree = 1 To db.Destinies
        ProcessTree db.Destiny(log.Tree)
    Next
End Sub

Private Sub ProcessPointer(ptypPointer As PointerType)
    Dim strField As String
    Dim strValue As String
    Dim lngTier As Long
    Dim strData As String
    Dim strSelector As String
    Dim lngPos As Long
    
    If Len(ptypPointer.Raw) = 0 Or ptypPointer.Ability <> 0 Or ptypPointer.Feat <> 0 Then Exit Sub
    log.HasError = False
    log.ptr = ptypPointer
    ' Split on first ":" strField = Left, strData = Right
    lngPos = InStr(ptypPointer.Raw, ": ")
    strField = Left$(ptypPointer.Raw, lngPos - 1)
    strData = Mid$(ptypPointer.Raw, lngPos + 2)
    ProcessRank ptypPointer, strData
        
    'Requirements can be FEAT
    If strField = "Feat" Then
        ProcessPointerFeat ptypPointer
        Exit Sub
    End If
    
    'or TIER
    ' Get tier from rightmost word in strField
    If strField = "Tier" Then
        lngPos = InStr(strField, "Tier ")
        strValue = Mid$(strField, lngPos + 5)  'Tier Number
        ptypPointer.Tier = val(strValue)
        If lngPos = 1 Then
            ptypPointer.Tree = log.Tree
            ptypPointer.Style = log.Style
        Else
            strField = Left$(strField, lngPos - 2)
            ' Pointing to a foreign tree
            ptypPointer.Tree = SeekTree(strField, ptypPointer.Style)
            If ptypPointer.Tree = 0 Then
                LogError
                Exit Sub
            End If
        End If
        If ptypPointer.Style = peEnhancement Then
            FindAbility db.Tree(ptypPointer.Tree), ptypPointer, strData
        Else
            FindAbility db.Destiny(ptypPointer.Tree), ptypPointer, strData
        End If
    End If
    If strField = "Class" Then
    End If
End Sub

Private Sub ProcessRank(ptypPointer As PointerType, pstrData As String)
    If Len(pstrData) < 8 Then Exit Sub
    If Mid$(pstrData, Len(pstrData) - 6, 6) <> " Rank " Then Exit Sub
    ptypPointer.Rank = val(Right$(pstrData, 1))
    pstrData = Left$(pstrData, Len(pstrData) - 7)
End Sub

Private Sub ProcessTree(ptypTree As TreeType)
    Dim lngStart As Long
    Dim i As Long
    
    If log.Style = peDestiny Then lngStart = 1
    With ptypTree
        For log.Tier = lngStart To .Tiers
            For log.Ability = 1 To .Tier(log.Tier).Abilities
                With .Tier(log.Tier).Ability(log.Ability)
                    ProcessPointer .Parent
                    For i = 1 To .Siblings
                        ProcessPointer .Sibling(i)
                    Next
                    If .RankReqs Then
                        For log.Rank = 2 To 3
                            ProcessReqs .Rank(log.Rank).Req
                        Next
                    End If
                    ProcessReqs .Req
                    For log.Selector = 1 To .Selectors
                        ProcessReqs .Selector(log.Selector).Req
                        If .Selector(log.Selector).RankReqs Then
                            For log.Rank = 2 To 3
                                ProcessReqs .Selector(log.Selector).Rank(log.Rank).Req
                            Next
                        End If
                    Next
                    log.Selector = 0
                End With
            Next
        Next
    End With
End Sub

Private Sub ProcessReqs(ptypReqList() As ReqListType)
    For log.ReqGroup = rgeAll To rgeNone
        With ptypReqList(log.ReqGroup)
            For log.Req = 1 To .Reqs
                ProcessPointer .Req(log.Req)
            Next
        End With
    Next
End Sub

Private Sub ProcessPointerFeat(ptypPointer As PointerType)
    Dim strRaw As String
    Dim strFeat As String
    Dim strSelector As String
    Dim lngPos As Long
    Dim blnError As Boolean
    Dim i As Long
    
    log.HasError = False
    log.Level = ptypPointer.Tier
    ptypPointer.Style = peFeat
    If Left$(ptypPointer.Raw, 6) = "Feat: " Then strRaw = Mid$(ptypPointer.Raw, 7) Else strRaw = ptypPointer.Raw
    lngPos = InStr(strRaw, ": ")
    If lngPos = 0 Then
        ptypPointer.Selector = 0
        ptypPointer.Feat = SeekFeat(strRaw)
        If ptypPointer.Feat = 0 Then LogError
    Else
        strFeat = Left$(strRaw, lngPos - 1)
        strSelector = Mid$(strRaw, lngPos + 2)
        ptypPointer.Feat = SeekFeat(strFeat)
        If ptypPointer.Feat = 0 Then
            LogError
        Else
            With db.Feat(ptypPointer.Feat)
                For i = 1 To .Selectors
                    If .Selector(i).SelectorName = strSelector Then
                        ptypPointer.Selector = i
                        Exit For
                    End If
                Next
                If ptypPointer.Selector = 0 Then LogError
            End With
        End If
    End If
End Sub

Private Sub FindAbility(ptypTree As TreeType, ptypPointer As PointerType, pstrRaw As String)
    Dim strAbility As String
    Dim strSelector As String
    Dim lngMax As Long
    Dim lngPos As Long
    Dim lngRank As Long
    Dim i As Long
    Dim strSPC As String
    Dim strSPC2 As String
    strSPC = "Spell Critical Chance: "
    strSPC2 = "Spell Critical: "
    
    ' Deal with SPC
    If (Left(pstrRaw, Len(strSPC)) = strSPC) Or (Left(pstrRaw, Len(strSPC2)) = strSPC2) Then
        strAbility = pstrRaw
    Else
        'Parse the string for ':' -
        lngPos = InStr(pstrRaw, ": ")
        If lngPos Then
            strAbility = Left$(pstrRaw, lngPos - 1)
            strSelector = Mid$(pstrRaw, lngPos + 2)
        Else
            strAbility = pstrRaw
        End If
    End If
    With ptypPointer
        With ptypTree.Tier(.Tier)
            lngMax = .Abilities
            For i = 1 To lngMax
                If .Ability(i).AbilityName = strAbility Then Exit For
            Next
        End With
    End With
    If i > lngMax Then
        LogError
        Exit Sub
    End If
    ptypPointer.Ability = i
    If Len(strSelector) = 0 Then
        ptypPointer.Selector = 0
        Exit Sub
    End If
    With ptypPointer
        With ptypTree.Tier(.Tier).Ability(.Ability)
            lngMax = .Selectors
            For i = 1 To lngMax
                If .Selector(i).SelectorName = strSelector Then Exit For
            Next
        End With
    End With
    If i > lngMax Then
        AddError ptypTree.TreeName & " Tier " & ptypPointer.Tier & " selector not found: " & pstrRaw
    Else
        ptypPointer.Selector = i
    End If
End Sub
