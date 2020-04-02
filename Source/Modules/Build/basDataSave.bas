Attribute VB_Name = "basDataSave"
Option Explicit

Public Sub SaveSpellFile()
    Dim strFile As String
    Dim strOld As String
    Dim strRaw() As String
    Dim i As Long
    
    If db.Spells = 0 Then Exit Sub
    SortSpells
    strFile = DataPath() & "Spells.txt"
    strOld = DataPath() & "Spells.old"
    If xp.File.Exists(strOld) Then xp.File.Delete strOld
    xp.File.Rename strFile, strOld
    ReDim strRaw(1 To db.Spells)
    For i = 1 To db.Spells
        With db.Spell(i)
            strRaw(i) = "SpellName: " & .SpellName & vbNewLine
            If .Wiki <> .SpellName Then strRaw(i) = strRaw(i) & "WikiName: " & .Wiki & vbNewLine
            If .Rare Then strRaw(i) = strRaw(i) & "Flags: Rare" & vbNewLine
            If Len(.Descrip) Then strRaw(i) = strRaw(i) & "Descrip: " & .Descrip & vbNewLine
        End With
    Next
    xp.File.SaveStringAs strFile, Join(strRaw, vbNewLine)
End Sub

Public Sub SaveEnhancementsFile()
    Dim strFile As String
    Dim strContents As String
    Dim strRaw() As String
    Dim strStat As String
    Dim lngStat As Long
    Dim strWiki As String
    Dim strInitials As String
    Dim strColor As String
    Dim i As Long
    
    If db.Trees = 0 Then Exit Sub
    strFile = DataPath() & "Enhancements.txt"
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    ReDim strRaw(1 To db.Trees)
    For i = 1 To db.Trees
        With db.Tree(i)
            strRaw(i) = "TreeName: " & .TreeName & vbNewLine
            If .Abbreviation <> .TreeName Then strRaw(i) = strRaw(i) & "Abbreviation: " & .Abbreviation & vbNewLine
            strWiki = .Wiki
            If Right$(strWiki, 13) = " enhancements" Then strWiki = Left$(strWiki, Len(strWiki) - 13)
            If strWiki <> .TreeName Then strRaw(i) = strRaw(i) & "WikiName: " & strWiki & vbNewLine
            strRaw(i) = strRaw(i) & "Type: " & GetTreeStyleName(.TreeType) & vbNewLine
            If .Initial(0) = .Initial(1) Then strInitials = .Initial(0) Else strInitials = .Initial(0) & ", " & .Initial(1)
            strRaw(i) = strRaw(i) & "Initial: " & strInitials & vbNewLine
            Select Case .Color
                Case cveRed: strColor = "Red"
                Case cveBlue: strColor = "Blue"
                Case cveYellow: strColor = "Yellow"
                Case cveGreen: strColor = "Green"
                Case cveOrange: strColor = "Orange"
                Case cvePurple: strColor = "Purple"
                Case Else: strColor = vbNullString
            End Select
            If Len(strColor) Then strRaw(i) = strRaw(i) & "Color: " & strColor & vbNewLine
            If .Stats(0) Then
                strStat = "Stats: "
                For lngStat = 1 To 6
                    If .Stats(lngStat) Then strStat = strStat & GetStatName(lngStat) & ", "
                Next
                strRaw(i) = strRaw(i) & Left$(strStat, Len(strStat) - 2) & vbNewLine
            End If
            If Len(.Lockout) Then strRaw(i) = strRaw(i) & "Lockout: " & .Lockout & vbNewLine
            strRaw(i) = strRaw(i) & vbNewLine & SaveAbilities(db.Tree(i))
        End With
    Next
    strContents = Join(strRaw, vbNewLine)
    Do While Right$(strContents, 2) = vbNewLine
        strContents = Left$(strContents, Len(strContents) - 2)
    Loop
    xp.File.SaveStringAs strFile, strContents
End Sub

Public Sub SaveDestinyFile()
    Dim strFile As String
    Dim strContents As String
    Dim strRaw() As String
    Dim strStat As String
    Dim lngStat As Long
    Dim i As Long
    
    If db.Destinies = 0 Then Exit Sub
    strFile = DataPath() & "Destinies.txt"
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    ReDim strRaw(1 To db.Destinies)
    For i = 1 To db.Destinies
        With db.Destiny(i)
            strRaw(i) = "DestinyName: " & .TreeName & vbNewLine
            If .Abbreviation <> .TreeName Then strRaw(i) = strRaw(i) & "Abbreviation: " & .Abbreviation & vbNewLine
            strRaw(i) = strRaw(i) & "Type: Destiny" & vbNewLine
            strStat = "Stats: "
            For lngStat = 1 To 6
                If .Stats(lngStat) Then strStat = strStat & GetStatName(lngStat) & ", "
            Next
            strRaw(i) = strRaw(i) & Left$(strStat, Len(strStat) - 2) & vbNewLine
            strRaw(i) = strRaw(i) & vbNewLine & SaveAbilities(db.Destiny(i))
        End With
    Next
    strContents = Join(strRaw, vbNewLine)
    Do While Right$(strContents, 2) = vbNewLine
        strContents = Left$(strContents, Len(strContents) - 2)
    Loop
    xp.File.SaveStringAs strFile, strContents
End Sub

Private Function SaveAbilities(ptypTree As TreeType) As String
    Dim lngFirstTier As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim strRaw() As String
    Dim strGroup As String
    Dim lngGroup As Long
    Dim strReqs As String
    Dim strRankReqs As String
    Dim i As Long
    Dim r As Long
    
    ReDim strRaw(1 To 32)
    If ptypTree.TreeType = tseDestiny Then lngFirstTier = 1
    With ptypTree
        For lngTier = lngFirstTier To .Tiers
            With .Tier(lngTier)
                For lngAbility = 1 To GetLastAbility(lngTier, .Abilities, ptypTree.TreeType, ptypTree.Stats(0))
                    With .Ability(lngAbility)
                        i = i + 1
                        ' Name(s) and description
                        strRaw(i) = "AbilityName: " & .AbilityName & vbNewLine
                        If .Abbreviation <> .AbilityName Then strRaw(i) = strRaw(i) & "Abbreviation: " & .Abbreviation & vbNewLine
                        strRaw(i) = strRaw(i) & "Descrip: " & .Descrip & vbNewLine
                        ' Base info
                        strRaw(i) = strRaw(i) & "Tier: " & lngTier & vbNewLine
                        If .Ranks <> 1 Then strRaw(i) = strRaw(i) & "Ranks: " & .Ranks & vbNewLine
                        If .Cost <> 1 Then strRaw(i) = strRaw(i) & "Cost: " & .Cost & vbNewLine
                        ' All/One/None
                        strReqs = GetTreeReqs(ptypTree.TreeID, ptypTree.Tier(lngTier).Ability(lngAbility).Req, lngTier)
                        If Len(strReqs) Then strRaw(i) = strRaw(i) & strReqs
'                        ' Class
'                        If .Class(0) Then
'                            strRaw(i) = strRaw(i) & GetClassReqs(.Class, .ClassLevel) & vbNewLine
'                        End If
                        ' Rank reqs
                        If .RankReqs Then
                            strRankReqs = GetRankReqs(ptypTree.TreeID, ptypTree.Tier(lngTier).Ability(lngAbility).Rank, lngTier)
                            If Len(strRankReqs) Then strRaw(i) = strRaw(i) & strRankReqs
                        End If
                        ' Selectors
                        If .SelectorOnly Then strRaw(i) = strRaw(i) & "Flags: SelectorOnly" & vbNewLine
                        If .SelectorStyle <> sseNone Then strRaw(i) = strRaw(i) & GetTreeSelectors(ptypTree.TreeID, ptypTree.Tier(lngTier).Ability(lngAbility))
                    End With
                Next
            End With
        Next
    End With
    If i = 0 Then Exit Function
    ReDim Preserve strRaw(1 To i)
    SaveAbilities = Join(strRaw, vbNewLine)
End Function

Private Function GetClassReqs(pblnClass() As Boolean, plngClassLevel() As Long) As String
    Dim strReturn As String
    Dim i As Long
    
    For i = 1 To ceClasses - 1
        If pblnClass(i) Then
            If Len(strReturn) Then strReturn = ", " & strReturn
            strReturn = strReturn & db.Class(i).ClassName & " " & plngClassLevel(i)
        End If
    Next
    GetClassReqs = "Class: " & strReturn
End Function

Private Function GetLastAbility(plngTier As Long, plngAbilities As Long, penTreeType As TreeStyleEnum, pblnStats As Boolean) As Long
    Select Case penTreeType
        Case tseRace: GetLastAbility = plngAbilities
        Case tseDestiny: GetLastAbility = plngAbilities - 1
        Case Else: If pblnStats And (plngTier = 3 Or plngTier = 4) Then GetLastAbility = plngAbilities - 1 Else GetLastAbility = plngAbilities
    End Select
End Function

Private Function GetTreeReqs(plngSource As Long, ptypReq() As ReqListType, plngTier As Long) As String
    Dim strReturn As String
    Dim i As Long
    
    For i = 1 To 3
        strReturn = strReturn & GetTreeReq(plngSource, ptypReq(i), i, plngTier)
    Next
    GetTreeReqs = strReturn
End Function

Private Function GetRankReqs(plngSource As Long, ptypRank() As RankType, plngTier As Long) As String
    Dim strReturn As String
    Dim strReq As String
    Dim i As Long
    Dim r As Long
    
    For r = 2 To 3
        If ptypRank(r).Class(0) Then
            strReturn = strReturn & "Rank" & r & GetClassReqs(ptypRank(r).Class, ptypRank(r).ClassLevel) & vbNewLine
        End If
        For i = 1 To 3
            strReq = GetTreeReq(plngSource, ptypRank(r).Req(i), i, plngTier)
            If Len(strReq) Then strReturn = strReturn & "Rank" & r & strReq
        Next
    Next
    GetRankReqs = strReturn
End Function

Private Function GetTreeReq(plngSource As Long, ptypReq As ReqListType, plngReqGroupID As Long, plngTier As Long) As String
    Dim strReturn As String
    Dim blnReqs As Boolean
    Dim i As Long
    
    If ptypReq.Reqs = 0 Then Exit Function
    strReturn = GetReqGroupName(plngReqGroupID) & ": "
    For i = 1 To ptypReq.Reqs
        If Not (plngReqGroupID = 1 And plngTier = 0 And ptypReq.Req(i).Tree = plngSource And ptypReq.Req(i).Tier = 0) Then
            strReturn = strReturn & ResolvePointer(ptypReq.Req(i), plngSource) & ", "
            blnReqs = True
        End If
    Next
    If blnReqs Then
        Mid$(strReturn, Len(strReturn) - 1, 2) = vbNewLine
        GetTreeReq = strReturn
    End If
End Function

Private Function ResolvePointer(ptypPointer As PointerType, plngSource As Long) As String
    Dim strSource As String
    Dim strName As String
    Dim strSelector As String
    Dim strRank As String
    
    With ptypPointer
        Select Case .Style
            Case peFeat
                If plngSource Then strSource = "Feat: "
                strName = db.Feat(.Feat).FeatName
                If .Selector Then strSelector = ": " & db.Feat(.Feat).Selector(.Selector).SelectorName
            Case peEnhancement
                If .Tree <> plngSource Then strSource = db.Tree(.Tree).TreeName & " "
                strName = "Tier " & .Tier & ": " & db.Tree(.Tree).Tier(.Tier).Ability(.Ability).AbilityName
                If .Selector Then strSelector = ": " & db.Tree(.Tree).Tier(.Tier).Ability(.Ability).Selector(.Selector).SelectorName
                If .Rank <> 0 Then strRank = " Rank " & .Rank
            Case peDestiny
                If .Tree <> plngSource Then strSource = db.Destiny(.Tree).TreeName & " "
                strName = "Tier " & .Tier & ": " & db.Destiny(.Tree).Tier(.Tier).Ability(.Ability).AbilityName
                If .Selector Then strSelector = ": " & db.Destiny(.Tree).Tier(.Tier).Ability(.Ability).Selector(.Selector).SelectorName
                If .Rank <> 0 Then strRank = " Rank " & .Rank
        End Select
    End With
    ResolvePointer = strSource & strName & strSelector & strRank
End Function

Private Function GetTreeSelectors(plngSource As Long, ptypAbility As AbilityType) As String
    Dim blnSelectorList As Boolean
    Dim strRoot As String
    Dim strSiblings As String
    Dim strSelectorList As String
    Dim strUnique As String
    Dim i As Long
    
    With ptypAbility
        Select Case .SelectorStyle
            Case sseRoot
                blnSelectorList = True
            Case sseExclusive
                strRoot = "Parent: "
                For i = 1 To .Siblings
                    strSiblings = strSiblings & ResolvePointer(.Sibling(i), plngSource) & ", "
                Next
                If Len(strSiblings) Then
                    Mid$(strSiblings, Len(strSiblings) - 1, 2) = vbNewLine
                    strSiblings = "Siblings: " & strSiblings
                End If
            Case sseShared
                strRoot = "SharedSelector: "
        End Select
    End With
    If Len(strRoot) Then strRoot = strRoot & ResolvePointer(ptypAbility.Parent, plngSource) & vbNewLine
    strUnique = GetTreeUniqueSelectors(ptypAbility, plngSource)
    If Len(strUnique) Then blnSelectorList = True
    strSelectorList = GetTreeSelectorList(ptypAbility, blnSelectorList)
    GetTreeSelectors = strRoot & strSiblings & strSelectorList & strUnique
End Function

Private Function GetTreeSelectorList(ptypAbility As AbilityType, pblnSelectorList As Boolean) As String
    Dim strSelector() As String
    Dim i As Long
    
    With ptypAbility
        ReDim strSelector(1 To .Selectors)
        For i = 1 To .Selectors
            strSelector(i) = .Selector(i).SelectorName
        Next
    End With
    If Not pblnSelectorList Then
        With ptypAbility.Parent
            Select Case ptypAbility.Parent.Style
                Case peFeat
                    For i = 1 To db.Feat(.Feat).Selectors
                        If db.Feat(.Feat).Selector(i).SelectorName <> strSelector(i) Then
                            pblnSelectorList = True
                            Exit For
                        End If
                    Next
                Case peEnhancement
                    For i = 1 To db.Tree(.Tree).Tier(.Tier).Ability(.Ability).Selectors
                        If db.Tree(.Tree).Tier(.Tier).Ability(.Ability).Selector(i).SelectorName <> strSelector(i) Then
                            pblnSelectorList = True
                            Exit For
                        End If
                    Next
                Case peDestiny
                    For i = 1 To db.Destiny(.Tree).Tier(.Tier).Ability(.Ability).Selectors
                        If db.Destiny(.Tree).Tier(.Tier).Ability(.Ability).Selector(i).SelectorName <> strSelector(i) Then
                            pblnSelectorList = True
                            Exit For
                        End If
                    Next
            End Select
        End With
    End If
    If pblnSelectorList Then GetTreeSelectorList = "Selector: " & Join(strSelector, ", ") & vbNewLine
End Function

Private Function GetTreeUniqueSelectors(ptypAbility As AbilityType, plngSource As Long) As String
    Dim strReturn As String
    Dim strUnique As String
    Dim i As Long
    
    With ptypAbility
        For i = 1 To .Selectors
            strUnique = GetTreeUniqueSelectorReqs(ptypAbility, .Selector(i), plngSource)
            If Len(strUnique) Then strReturn = strReturn & strUnique
        Next
    End With
    GetTreeUniqueSelectors = strReturn
End Function

Private Function GetTreeUniqueSelectorReqs(ptypAbility As AbilityType, ptypSelector As SelectorType, plngSource As Long) As String
    Dim strReturn As String
    Dim i As Long
    
    If ptypSelector.Cost <> ptypAbility.Cost Then strReturn = strReturn & "Cost: " & ptypSelector.Cost & vbNewLine
    If Not CompareReqs(ptypSelector.Req, ptypAbility.Req) Then strReturn = strReturn & GetTreeReqs(plngSource, ptypSelector.Req, -1)
    If ptypSelector.RankReqs Then strReturn = strReturn & GetRankReqs(plngSource, ptypSelector.Rank, -1)
    If Len(strReturn) Then GetTreeUniqueSelectorReqs = "SelectorName: " & ptypSelector.SelectorName & vbNewLine & strReturn
End Function

Private Function CompareReqs(ptypReq1() As ReqListType, ptypReq2() As ReqListType) As Boolean
    Dim i As Long
    
    For i = 1 To 3
        If Not CompareReq(ptypReq1(i), ptypReq2(i)) Then Exit Function
    Next
    CompareReqs = True
End Function

Private Function CompareReq(ptypReq1 As ReqListType, ptypReq2 As ReqListType) As Boolean
    Dim i As Long
    
    If ptypReq1.Reqs <> ptypReq2.Reqs Then Exit Function
    For i = 1 To ptypReq1.Reqs
        If ptypReq1.Req(i).Style <> ptypReq2.Req(i).Style Then Exit Function
        If ptypReq1.Req(i).Tier <> ptypReq2.Req(i).Tier Then Exit Function
        If ptypReq1.Req(i).Ability <> ptypReq2.Req(i).Ability Then Exit Function
        If ptypReq1.Req(i).Selector <> ptypReq2.Req(i).Selector Then Exit Function
        If ptypReq1.Req(i).Tree <> ptypReq2.Req(i).Tree Then Exit Function
        If ptypReq1.Req(i).Feat <> ptypReq2.Req(i).Feat Then Exit Function
    Next
    CompareReq = True
End Function
