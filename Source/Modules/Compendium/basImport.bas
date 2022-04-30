Attribute VB_Name = "basImport"
' Written by Ellis Dee
'
' This is all duplicate code, which is a definite bad practice.
' I was concerned about bugs in the load routines if I tried to incorporate this
' directly into basDataLoad, so I just copied what I needed here and
' called it good enough.
Option Explicit
    
Public idb As DatabaseType


' ************* IMPORT COMPENDIUM *************


Public Sub ImportCompendium(pstrFile As String)
    Dim strSection() As String
    Dim lngSections As Long
    Dim strName As String
    Dim lngPos As Long
    Dim i As Long
    Dim s As Long
    
    If Not xp.File.Exists(pstrFile) Then Exit Sub
    ClearImportData
    idb = db
    strSection = Split(xp.File.LoadToString(pstrFile), "SectionName: ")
    lngSections = UBound(strSection)
    If lngSections = 0 Then Exit Sub
    For i = 1 To UBound(strSection)
        lngPos = InStr(strSection(i), vbNewLine)
        If lngPos > 1 Then
            strName = LCase$(Trim$(Left$(strSection(i), lngPos - 1)))
            Select Case strName
                Case "characters": CompendiumCharacters strSection(i)
                Case "quests": CompendiumQuests strSection(i)
                Case "challenges": CompendiumChallenges strSection(i)
            End Select
        End If
    Next
End Sub

Public Sub ClearImportData()
    Dim typBlank As DatabaseType
    
    idb = typBlank
End Sub

Private Sub CompendiumCharacters(pstrRaw As String)
    Dim strFile As String
    Dim strCharacter() As String
    Dim lngCharacters As Long
    Dim i As Long
    
    strCharacter = Split(pstrRaw, "Character: ")
    lngCharacters = UBound(strCharacter)
    If lngCharacters = 0 Then Exit Sub
    ReDim idb.Character(1 To lngCharacters + 1)
    idb.Characters = 0
    For i = 1 To UBound(strCharacter)
        CompendiumCharacter strCharacter(i)
    Next
    If UBound(idb.Character) <> idb.Characters Then ReDim Preserve idb.Character(1 To idb.Characters)
    AllocateCharacters
End Sub

Private Sub AllocateCharacters()
    Dim i As Long
    
    For i = 1 To idb.Quests
        With idb.Quest(i)
            If idb.Characters = 0 Then Erase .Progress Else ReDim .Progress(1 To idb.Characters)
        End With
    Next
    For i = 1 To idb.Challenges
        With idb.Challenge(i)
            If idb.Characters = 0 Then Erase .Stars Else ReDim .Stars(1 To idb.Characters)
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
    InitImportSagas typNew
    With typNew
        ' Process lines
        For lngLine = 1 To UBound(strLine)
            If ParseLine(strLine(lngLine), strField, strItem, lngValue, strList, lngListMax) Then
                Select Case strField
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
                                For i = 1 To idb.Saga(lngSaga).Quests
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
                            .Tome.Stat(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, tomes.Stat.Max)
                        Next
                    Case "tomeskill"
                        For i = 1 To 21
                            .Tome.Skill(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, tomes.Skill.Max)
                        Next
                    Case "tomeracialap"
                        .Tome.RacialAP = LimitValue(val(strItem), 0, tomes.RacialAPMax)
                    Case "tomeuniversalap"
                        .Tome.UniversalAP = LimitValue(val(strItem), 0, tomes.UniversalAPMax)
                    Case "tomefate"
                        .Tome.Fate = LimitValue(val(strItem), 0, tomes.FateTomeMax)
                    Case "tomedestiny"
                        .Tome.Destiny = LimitValue(val(strItem), 0, tomes.DestinyTomeMax)
                    Case "tomepower"
                        For i = 1 To 3
                            .Tome.Power(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, tomes.PowerMax)
                        Next
                    Case "tomerr"
                        For i = 1 To 2
                            .Tome.RR(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, tomes.RRMax)
                        Next
                    Case "tomeheroicxp"
                        If strItem = "Lesser" Or strItem = "Greater" Then .Tome.HeroicXP = strItem
                    Case "tomeepicxp"
                        If strItem = "Lesser" Or strItem = "Greater" Then .Tome.EpicXP = strItem
                    Case "pastlifeclass"
                        For i = 1 To 14
                            .PastLife.Class(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                    Case "pastliferace"
                        For i = 1 To 11
                            .PastLife.Racial(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                    Case "pastlifeiconic"
                        For i = 1 To 6
                            .PastLife.Iconic(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                    Case "pastlifeepic"
                        For i = 1 To 12
                            .PastLife.Epic(i) = LimitValue(val(Mid$(strItem, i, 1)), 0, 3)
                        Next
                End Select
            End If
        Next
    End With
    idb.Characters = idb.Characters + 1
    idb.Character(idb.Characters) = typNew
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

Private Sub InitImportSagas(ptypCharacter As CharacterType)
    Dim i As Long
    
    Erase ptypCharacter.Saga
    If idb.Sagas Then
        ReDim ptypCharacter.Saga(idb.Sagas)
        For i = 1 To idb.Sagas
            ReDim ptypCharacter.Saga(i).Progress(1 To idb.Saga(i).Quests)
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
                With idb.Quest(lngQuest)
                    For c = 1 To idb.Characters
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
                With idb.Challenge(lngChallenge)
                    For c = 1 To idb.Characters
                        .Stars(c) = val(Mid$(strToken(1), c, 1))
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
        If IsNumeric(pstrItem) Then plngValue = val(pstrItem)
    End If
    ' Return single item in list form as well
    plngListMax = 0
    ReDim pstrList(0)
    pstrList(0) = pstrLine
End Function


