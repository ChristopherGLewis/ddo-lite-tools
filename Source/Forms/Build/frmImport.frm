VERSION 5.00
Begin VB.Form frmImport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import"
   ClientHeight    =   7770
   ClientLeft      =   30
   ClientTop       =   405
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12225
   Begin VB.TextBox txtImport 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2148
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3420
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.Label lnkNav 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   0
      Left            =   11520
      TabIndex        =   3
      Tag             =   "nav"
      Top             =   84
      Width           =   432
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Paste from Clipboard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   1
      Left            =   264
      TabIndex        =   2
      Tag             =   "nav"
      Top             =   84
      Width           =   2064
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Translate Build"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   2
      Left            =   2892
      TabIndex        =   1
      Tag             =   "nav"
      Top             =   84
      Width           =   1452
   End
   Begin VB.Shape shpBorder 
      Height          =   720
      Left            =   780
      Top             =   1920
      Visible         =   0   'False
      Width           =   1656
   End
   Begin VB.Shape shpNav 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   384
      Left            =   0
      Tag             =   "nav"
      Top             =   0
      Width           =   12216
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Type GuessType
    BuildPoints As BuildPointsEnum
    Include As Boolean
    Skills As ValidEnum
    Feats As ValidEnum
End Type

Private Enum SectionEnum
    seHeader
    seLevelOrder
    seStats
    seSkills
    seFeats
    seSpells
    seEnhancements
    seLevelingGuide
    seDestiny
    seTwists
    seNoChange
End Enum

Private Enum StatColumnEnum
    scAdventurer
    scChampion
    scHero
    scLegend
    scTome
    scLevelup
End Enum

Private Enum ProcessEnum
    peBegin
    peProcess
    peEnd
End Enum

Private mtypFeatOutput As FeatOutputType

Private mblnOverride As Boolean

Private mlngHeroic As Long
Private mlngEpic As Long
Private mbln20orMax As Boolean
Private mblnMax As Boolean

Private mlngColumns As Long
Private menColumn() As StatColumnEnum
Private mlngColPos() As Long
Private mblnColumns As Boolean

Private menProcess As ProcessEnum

Private menClassSpell As ClassEnum
Private mlngSpellLevel As Long

Private mlngTree As Long
Private mlngTier As Long

Private mlngDestiny As Long
Private mlngDestinyTier As Long

Private mtypGuess() As GuessType
Private mlngGuesses As Long


' ************* INITIALIZE *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.Configure Me
End Sub

Private Sub Form_Activate()
    ActivateForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
    If mblnOverride Then mblnOverride = False Else frmMain.tmrOutput.Enabled = True
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    If cfg.GetColor(cgeNavigation, cveBackground) = cfg.GetColor(cgeControls, cveBackground) Then lngTop = (Me.shpNav.Height * 3) \ 2 Else lngTop = Me.shpNav.Height
    lngHeight = Me.ScaleHeight - lngTop
    With Me.shpBorder
        .Move 0, lngTop, Me.ScaleWidth, lngHeight
        Me.txtImport.Move PixelX, .Top + PixelY, .Width - PixelX * 2, .Height - PixelY * 2
        .Visible = True
    End With
    Me.txtImport.Visible = True
End Sub


' ************* NAVIGATION *************


Private Sub lnkNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strFile As String
    
    xp.SetMouseCursor mcHand
    If Button <> vbLeftButton Then Exit Sub
    Select Case Me.lnkNav(Index).Caption
        Case "Paste from Clipboard"
            Me.txtImport.Text = vbNullString
            If Clipboard.GetFormat(vbCFText) Then Me.txtImport.Text = Clipboard.GetText
        Case "Translate Build"
            BeginTranslate
        Case "Help"
            ShowHelp "Import"
    End Select
End Sub


' ************* TRANSLATE *************


Private Sub BeginTranslate()
On Error GoTo BeginTranslateErr
    If Len(Me.txtImport.Text) = 0 Then
        Notice "Nothing to translate."
        Exit Sub
    End If
    If CheckDirty() Then Exit Sub
    DoEvents
    xp.Mouse = msSystemWait
    InitBuild
    Translate
    FinishBuild
    CascadeChanges cceAll
    ShowBuild
    xp.Mouse = msSystemWait
    
BeginTranslateExit:
    Exit Sub
    
BeginTranslateErr:
    Notice "Invalid or Incomplete build text. Be sure you copied the entire build into the clipboard."
    BuildClose
    Resume BeginTranslateExit
End Sub

Private Sub InitBuild()
    ClearBuild
    SetBuildDefaults
    Erase build.IncludePoints
    mbln20orMax = False
    mblnMax = False
    mlngColumns = 0
    mblnColumns = False
    build.BuildName = "Imported Build"
    menClassSpell = ceAny
    mlngSpellLevel = 0
    mlngTree = 0
    mlngTier = 0
End Sub

Private Sub Translate()
    Dim strLine() As String
    Dim enSection As SectionEnum
    Dim i As Long
    
    strLine = Split(Me.txtImport.Text, vbNewLine)
    TrimLeadingSpaces strLine
    For i = 0 To UBound(strLine)
        Select Case enSection
            Case seStats: CleanStatsLine strLine(i)
            Case seFeats: strLine(i) = Replace(strLine(i), ".", " ")
            Case Else: strLine(i) = Replace(Trim$(strLine(i)), ".", " ")
        End Select
        If Len(strLine(i)) Then
            If Not ParseSection(strLine(i), enSection) Then
                Select Case enSection
                    Case seHeader: ParseHeader strLine(i)
                    Case seLevelOrder: ParseLevelOrder strLine(i)
                    Case seStats: ParseStats strLine(i)
                    Case seSkills: ParseSkills strLine(i)
                    Case seFeats: ParseFeats strLine(i)
                    Case seSpells: ParseSpells strLine(i)
                    Case seEnhancements: ParseEnhancements strLine(i)
                    Case seLevelingGuide
                    Case seDestiny: ParseDestiny strLine(i)
                End Select
            End If
        End If
    Next
End Sub

' Remove universal leading spaces (when entire build is padded with leading spaces on every line)
Private Sub TrimLeadingSpaces(pstrLine() As String)
    Dim lngSpaces As Long
    Dim i As Long
    Dim j As Long
    
    lngSpaces = 9999
    For i = 1 To UBound(pstrLine)
        If Len(pstrLine(i)) Then
            For j = 1 To Len(pstrLine(i))
                If Mid$(pstrLine(i), j, 1) <> " " Then Exit For
            Next
            If j = 1 Then Exit Sub
            If lngSpaces > j Then lngSpaces = j
        End If
    Next
    For i = 1 To UBound(pstrLine)
        pstrLine(i) = Mid$(pstrLine(i), lngSpaces)
    Next
End Sub

Private Function CleanStatsLine(pstrRaw As String)
    If InStr(pstrRaw, " Strength") Then
        pstrRaw = LTrim$(pstrRaw)
    ElseIf InStr(pstrRaw, " Dexterity") Then
        pstrRaw = LTrim$(pstrRaw)
    ElseIf InStr(pstrRaw, " Constitution") Then
        pstrRaw = LTrim$(pstrRaw)
    ElseIf InStr(pstrRaw, " Intelligence") Then
        pstrRaw = LTrim$(pstrRaw)
    ElseIf InStr(pstrRaw, " Wisdom") Then
        pstrRaw = LTrim$(pstrRaw)
    ElseIf InStr(pstrRaw, " Charisma") Then
        pstrRaw = LTrim$(pstrRaw)
    End If
    pstrRaw = Replace(pstrRaw, ".", " ")
End Function

' Returns TRUE if current line is a section header
Private Function ParseSection(pstrText As String, penSection As SectionEnum) As Boolean
    Dim strText As String
    Dim enNew As SectionEnum
    
    strText = Trim$(pstrText)
    ' Prep for next section
    If Left$(strText, 11) = "Level Order" Then
        enNew = seLevelOrder
    ElseIf Left$(strText, 5) = "Stats" Then
        enNew = seStats
    ElseIf Len(strText) = 6 And Left$(strText, 6) = "Skills" Then
        enNew = seSkills
    ElseIf Left$(strText, 5) = "Feats" Then
        enNew = seFeats
        InitBuildFeats
    ElseIf Len(strText) = 6 And Left$(strText, 6) = "Spells" Then
        InitBuildSpells
        enNew = seSpells
    ElseIf Left$(strText, 12) = "Enhancements" Then
        enNew = seEnhancements
    ElseIf Left$(strText, 14) = "Leveling Guide" Then
        enNew = seLevelingGuide
    ElseIf Left$(strText, 7) = "Destiny" Then
        enNew = seDestiny
    ElseIf Left$(strText, 14) = "Twists of Fate" Then
        enNew = seTwists
    Else
        Exit Function
    End If
    ' Cleanup for previous section, which is now finished
    Select Case penSection
        Case seHeader
            CalculateBAB
        Case seFeats
            InitBuildFeats
            If mbln20orMax And Not mblnMax Then build.MaxLevels = 20
    End Select
    menProcess = peBegin
    penSection = enNew
    ' Last minute initializations
    Select Case penSection
        Case seSpells
            For menClassSpell = ceBarbarian To ceClasses - 1
                If build.Spell(menClassSpell).MaxSpellLevel > 0 Then Exit For
            Next
            If menClassSpell = ceClasses Then menClassSpell = ceAny
        Case seDestiny
            mlngTree = 0
            mlngTier = 1
    End Select
    ParseSection = True
End Function

Private Sub FinishBuild()
    Dim enAllLevelups As StatEnum
    Dim i As Long
    
    ' All Levelups
    build.Levelups(0) = build.Levelups(1)
    For i = 2 To 7
        If build.Levelups(i) <> build.Levelups(0) Then
            build.Levelups(0) = aeAny
            Exit For
        End If
    Next
    GuessBuildPoints
End Sub

Private Sub ShowBuild()
    SetBuildOpen True
    SetDirty
    EnableMenus
    GenerateOutput oeAll
End Sub


' ************* HEADER *************


Private Sub ParseHeader(ByVal pstrText As String)
    If ParseClassSplit(pstrText) Then Exit Sub
    If ParseAlignRace(pstrText) Then Exit Sub
    build.BuildName = pstrText
End Sub

Private Function ParseClassSplit(ByVal pstrText As String) As Boolean
    Dim lngPos As Long
    Dim enClass As ClassEnum
    Dim strLevel As String
    Dim strLevels() As String
    Dim strClasses() As String
    Dim i As Long
    
    mlngHeroic = 0
    mlngEpic = 0
    ' Epic levels?
    lngPos = InStr(pstrText, ", Epic ")
    If lngPos > 0 Then
        strLevel = Mid$(pstrText, InStrRev(pstrText, " ") + 1)
        If Not IsNumeric(strLevel) Then Exit Function
        mlngEpic = val(strLevel)
        pstrText = Left$(pstrText, lngPos - 1)
    End If
    ' Pure class? ("Class ##")
    For enClass = 1 To ceClasses - 1
        If Left$(pstrText, Len(db.Class(enClass).ClassName)) = db.Class(enClass).ClassName Then
            lngPos = InStrRev(pstrText, " ")
            If lngPos = 0 Then Exit Function
            strLevel = Mid$(pstrText, lngPos + 1)
            If Not IsNumeric(strLevel) Then Exit Function
            mlngHeroic = val(strLevel)
            build.BuildClass(0) = enClass
            SetMaxLevel
            For i = 1 To 20
                build.Class(i) = enClass
            Next
            CalculateBAB
            ParseClassSplit = True
            Exit Function
        End If
    Next
    ' Multiclass? (#/#/# Class1/Class2/Class3)
    If InStr(pstrText, "/") = 0 Then Exit Function
    lngPos = InStr(pstrText, " ")
    If lngPos = 0 Then Exit Function
    strLevels = Split(Left$(pstrText, lngPos - 1), "/")
    strClasses = Split(Mid(pstrText, lngPos + 1), "/")
    If UBound(strLevels) < 1 Or UBound(strLevels) <> UBound(strClasses) Then Exit Function
    mlngHeroic = 0
    For i = 0 To UBound(strLevels)
        If Not IsNumeric(strLevels(i)) Then Exit For
        mlngHeroic = mlngHeroic + val(strLevels(i))
        enClass = GetClassID(strClasses(i))
        If enClass = ceAny Then Exit For
        build.BuildClass(i) = enClass
    Next
    If mlngHeroic = 0 Or enClass = ceAny Then
        Erase build.BuildClass
        Exit Function
    End If
    SetMaxLevel
    ParseClassSplit = True
End Function

Private Sub SetMaxLevel()
    If mlngHeroic = 20 And mlngEpic = 0 Then
        build.MaxLevels = MaxLevel
        mbln20orMax = True
    Else
        build.MaxLevels = mlngHeroic + mlngEpic
    End If
End Sub

Private Function ParseAlignRace(ByVal pstrText As String) As Boolean
    Dim strAlignment As String
    Dim strRace As String
    Dim i As Long
    
    For i = 1 To 6
        strAlignment = GetAlignmentName(i)
        If Left$(pstrText, Len(strAlignment)) = strAlignment Then
            build.Alignment = i
            Exit For
        End If
    Next
    For i = 1 To reRaces - 1
        strRace = GetRaceName(i)
        If Right$(pstrText, Len(strRace)) = strRace Then
            build.Race = i
            Exit For
        End If
    Next
    If build.Alignment <> aleAny Or build.Race <> reAny Then ParseAlignRace = True
End Function


' ************* LEVEL ORDER *************


Private Sub ParseLevelOrder(pstrRaw As String)
    ParseLevelOrderEntry pstrRaw, 1, 18
    ParseLevelOrderEntry pstrRaw, 19, 37
    ParseLevelOrderEntry pstrRaw, 38, 56
    ParseLevelOrderEntry pstrRaw, 57, 76
    CalculateBAB
End Sub

Private Sub ParseLevelOrderEntry(pstrRaw As String, plngStart As Long, ByVal plngEnd As Long)
    Dim strLevel As String
    Dim lngLevel As Long
    Dim strClass As String
    Dim enClass As ClassEnum
    
    ' Level
    If Len(pstrRaw) < plngStart - 2 Then Exit Sub
    strLevel = Trim$(Mid$(pstrRaw, plngStart, 2))
    If Not IsNumeric(strLevel) Then Exit Sub
    lngLevel = val(strLevel)
    If lngLevel < 1 Or lngLevel > 20 Then Exit Sub
    ' Class
    If Len(pstrRaw) < plngEnd Then plngEnd = Len(pstrRaw)
    If plngEnd < plngStart + 5 Then Exit Sub
    strClass = Trim$(Mid$(pstrRaw, plngStart + 3, plngEnd - plngStart - 2))
    enClass = GetClassID(strClass)
    If enClass = ceAny Then Exit Sub
    build.Class(lngLevel) = enClass
End Sub


' ************* STATS *************


Private Sub ParseStats(ByVal pstrRaw As String)
    Dim strColumn() As String
    Dim blnInclude As Boolean
    Dim enStat As StatEnum
    Dim lngStat As Long
    Dim strLevelup() As String
    Dim lngLevel As Long
    Dim strTome As String
    Dim i As Long
    
    pstrRaw = RTrim$(pstrRaw)
    If InStr(pstrRaw, "-") Then Exit Sub
    If InStr(pstrRaw, ": ") Then pstrRaw = Replace(pstrRaw, ": ", ":")
    StripSpaces pstrRaw, strColumn
    If mlngColumns = 0 Then
        mlngColumns = UBound(strColumn)
        ReDim menColumn(mlngColumns)
        For i = 1 To UBound(strColumn)
            blnInclude = True
            Select Case strColumn(i)
                Case "28pt": menColumn(i) = scAdventurer
                Case "30pt": menColumn(i) = scHero
                Case "32pt": If build.Race = reDrow Then menColumn(i) = scLegend Else menColumn(i) = scChampion
                Case "34pt": menColumn(i) = scHero
                Case "36pt": menColumn(i) = scLegend
                Case "Tome"
                    menColumn(i) = scTome
                    blnInclude = False
                Case "Level"
                    menColumn(i) = scLevelup
                    blnInclude = False
                Case Else
                    blnInclude = False
            End Select
            If blnInclude Then build.IncludePoints(menColumn(i)) = 1
        Next
    Else
        enStat = GetStatID(strColumn(0))
        For i = 1 To UBound(strColumn)
            ' First check can't be column based since Level 28 levelup is in a different column than the rest of the levelups
            If InStr(strColumn(i), ":") Then
                strLevelup = Split(strColumn(i), ":")
                If UBound(strLevelup) <> 1 Then Exit Sub
                If Not IsNumeric(strLevelup(0)) Then Exit Sub
                lngLevel = strLevelup(0)
                Select Case lngLevel
                    Case 4, 8, 12, 16, 20, 24, 28
                    Case Else: Exit Sub
                End Select
                enStat = GetStatID(strLevelup(1))
                If enStat = aeAny Then Exit Sub
                build.Levelups(lngLevel \ 4) = enStat
            ElseIf menColumn(i) = scTome Then
                If Len(strColumn(i)) <> 2 Then Exit Sub
                If Left$(strColumn(i), 1) <> "+" Then Exit Sub
                strTome = Right$(strColumn(i), 1)
                If Not IsNumeric(strTome) Then Exit Sub
                build.Tome(enStat) = val(strTome)
            Else
                If enStat = aeAny Then Exit Sub
                If Not IsNumeric(strColumn(i)) Then Exit Sub
                lngStat = val(strColumn(i)) - db.Race(build.Race).Stats(enStat)
                Select Case lngStat
                    Case 7: lngStat = 8
                    Case 8: lngStat = 10
                    Case 9: lngStat = 13
                    Case 10: lngStat = 16
                End Select
                build.StatPoints(menColumn(i), enStat) = lngStat
                build.StatPoints(menColumn(i), 0) = build.StatPoints(menColumn(i), 0) + lngStat
            End If
        Next
    End If
End Sub

Private Sub GuessBuildPoints()
    Dim i As Long
    
    Erase mtypGuess
    mlngGuesses = 0
    If AddGuess(beAdventurer) Then Exit Sub
    If AddGuess(beChampion) Then Exit Sub
    If AddGuess(beHero) Then Exit Sub
    If AddGuess(beLegend) Then Exit Sub
    ' We don't have a perfect distribution; look for the best fit
    If BestGuess(veComplete, veIncomplete) Then Exit Sub
    If BestGuess(veIncomplete, veComplete) Then Exit Sub
    If BestGuess(veIncomplete, veIncomplete) Then Exit Sub
    If BestGuess(veComplete, veErrors) Then Exit Sub
    If BestGuess(veErrors, veComplete) Then Exit Sub
    If BestGuess(veIncomplete, veErrors) Then Exit Sub
    If BestGuess(veErrors, veIncomplete) Then Exit Sub
    ' No good options; give up and just take the first possibility
    If mlngGuesses > 0 Then SetBuildPoints 1
End Sub

Private Function AddGuess(penBuildPoints As BuildPointsEnum) As Boolean
    Dim blnClassRow As Boolean
    
    If build.IncludePoints(penBuildPoints) = 0 Then Exit Function
    mlngGuesses = mlngGuesses + 1
    ReDim Preserve mtypGuess(1 To mlngGuesses)
    With mtypGuess(mlngGuesses)
        .BuildPoints = penBuildPoints
        SetBuildPoints mlngGuesses
        .Skills = ValidSkills(blnClassRow, True)
        If .Skills = veEmpty Then .Skills = veComplete
        .Feats = ValidFeats(True)
        If .Feats = veEmpty Then .Feats = veComplete
        If .Skills = veComplete And .Feats = veComplete Then AddGuess = SetBuildPoints(mlngGuesses)
    End With
End Function

Private Function BestGuess(penSkills As ValidEnum, penFeats As ValidEnum) As Boolean
    Dim i As Long
    
    For i = 1 To mlngGuesses
        If mtypGuess(mlngGuesses).Skills = penSkills And mtypGuess(mlngGuesses).Feats = penFeats Then
            BestGuess = SetBuildPoints(i)
            Exit For
        End If
    Next
End Function

Private Function SetBuildPoints(plngGuess As Long) As Boolean
    With mtypGuess(plngGuess)
        build.BuildPoints = .BuildPoints
    End With
    SetBuildPoints = True
End Function


' ************* SKILLS *************


Private Sub ParseSkills(ByVal pstrRaw As String)
    Dim enSkill As SkillsEnum
    Dim strSkill As String
    Dim lngLevel As Long
    Dim strRanks As String
    Dim lngRanks As Long
    Dim lngPoints As Long
    Dim lngPos As Long
    Dim i As Long
    
    If Left$(LTrim$(pstrRaw), 3) = "---" Then
        If menProcess = peBegin Then menProcess = peProcess Else menProcess = peEnd
        Exit Sub
    End If
    If menProcess <> peProcess Then Exit Sub
    For enSkill = seBalance To seUMD
        strSkill = GetSkillName(enSkill, True)
        If Left$(pstrRaw, Len(strSkill)) = strSkill Then Exit For
    Next
    If enSkill > seUMD Then Exit Sub
    For lngLevel = 1 To 20
        lngPos = (lngLevel * 3) + 7
        If Len(pstrRaw) < lngPos + 2 Then Exit For
        strRanks = Trim$(Mid$(pstrRaw, lngPos, 3))
        If Len(strRanks) Then ParseRanks enSkill, lngLevel, strRanks
    Next
End Sub

Private Sub ParseRanks(penSkill As SkillsEnum, plngLevel As Long, ByVal pstrRanks As String)
    Dim enClass As ClassEnum
    Dim blnNative As Boolean
    Dim blnHalf As Boolean
    Dim lngPoints As Long
    
    enClass = build.Class(plngLevel)
    If enClass = ceAny Then Exit Sub
    blnNative = db.Class(enClass).NativeSkills(penSkill)
    If Right$(pstrRanks, 1) = "½" Then
        blnHalf = True
        If Len(pstrRanks) = 1 Then pstrRanks = "0" Else pstrRanks = Left$(pstrRanks, Len(pstrRanks) - 1)
    End If
    If Not IsNumeric(pstrRanks) Then Exit Sub
    lngPoints = val(pstrRanks)
    If Not blnNative Then lngPoints = lngPoints * 2
    If blnHalf Then lngPoints = lngPoints + 1
    build.Skills(penSkill, plngLevel) = lngPoints
End Sub


' ************* FEATS *************


Private Sub ParseFeats(ByVal pstrRaw As String)
    Dim strLevel As String
    Dim lngLevel As Long
    Dim strSource As String
    Dim strDisplay As String
    Dim blnFound As Boolean
    Dim lngIndex As Long
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim strOR() As String
    Dim lngAlternate As Long
    Dim i As Long
    
    ' Parse line into level and display
    strLevel = Trim$(Left$(pstrRaw, 2))
    If Not IsNumeric(strLevel) Then Exit Sub
    lngLevel = val(strLevel)
    If Mid$(pstrRaw, 11, 2) <> ": " Then Exit Sub
    strSource = Trim$(Mid$(pstrRaw, 4, 7))
    strDisplay = Mid$(pstrRaw, 13)
    ' Handle feat swaps
    If strSource = "Swap" Then
        AddExchangeFeat strDisplay, lngLevel
    Else
        ' Find this slot
        For lngIndex = 1 To Feat.Count
            Select Case Feat.List(lngIndex).ActualType
                Case bftGranted, bftAlternate, bftExchange
                Case Else
                    If Feat.List(lngIndex).SourceOutput = strSource And Feat.List(lngIndex).Level = lngLevel Then
                        blnFound = True
                        Exit For
                    End If
            End Select
        Next
        If Not blnFound Then Exit Sub
        If InStr(strDisplay, " OR ") Then
            strOR = Split(strDisplay, " OR ")
            If ParseFeat(strOR(0), lngFeat, lngSelector) Then
                ApplyFeat lngIndex, lngFeat, lngSelector
                For i = 1 To UBound(strOR)
                    If Not ParseFeat(strOR(i), lngAlternate, lngSelector) Then ParseFeat db.Feat(lngFeat).FeatName & ": " & strOR(i), lngAlternate, lngSelector
                    If lngAlternate Then AddAlternateFeat lngIndex, lngAlternate, lngSelector
                Next
            End If
        Else
            ParseFeat strDisplay, lngFeat, lngSelector
            ApplyFeat lngIndex, lngFeat, lngSelector
        End If
    End If
    If mbln20orMax And lngLevel > 20 Then mblnMax = True
End Sub

' Split a feat display into FeatName and Selector#
Private Function ParseFeat(pstrDisplay As String, plngFeat As Long, plngSelector As Long) As Boolean
    Dim strFeat As String
    Dim strSelector As String
    Dim lngPos As Long
    Dim i As Long
    
    plngSelector = 0
    plngFeat = SeekFeat(pstrDisplay)
    If plngFeat = 0 Then
        lngPos = InStr(pstrDisplay, ": ")
        If lngPos Then
            ' Selector feat
            strFeat = Left$(pstrDisplay, lngPos - 1)
            strSelector = Mid$(pstrDisplay, lngPos + 2)
            plngFeat = SeekFeat(strFeat)
            If plngFeat = 0 Then
                plngSelector = 0
                Exit Function
            End If
        Else
            ' SelectorOnly feat
            For i = 1 To db.Feats
                With db.Feat(i)
                    If .SelectorOnly Then
                        For plngSelector = 1 To .Selectors
                            If .Selector(plngSelector).SelectorName = pstrDisplay Then Exit For
                        Next
                        If plngSelector > .Selectors Then
                            plngSelector = 0
                        Else
                            plngFeat = i
                            strSelector = pstrDisplay
                        End If
                    End If
                End With
                If plngFeat Then Exit For
            Next
        End If
    End If
    ' Locate selector
    If plngSelector = 0 And Len(strSelector) <> 0 Then
        With db.Feat(plngFeat)
            For plngSelector = 1 To .Selectors
                If .Selector(plngSelector).SelectorName = strSelector Then Exit For
            Next
            If plngSelector > .Selectors Then plngSelector = 0
        End With
    End If
    ParseFeat = (plngFeat <> 0)
End Function

Private Sub ApplyFeat(plngIndex As Long, plngFeat As Long, plngSelector As Long)
    If plngIndex = 0 Then Exit Sub
    With Feat.List(plngIndex)
        With build.Feat(.ActualType).Feat(.Index)
            .FeatName = db.Feat(plngFeat).FeatName
            .Selector = plngSelector
        End With
        .Display = GetFeatDisplay(plngFeat, plngSelector, False, False)
        .FeatID = plngFeat
        .FeatName = db.Feat(plngFeat).FeatName
        .Selector = plngSelector
    End With
    ' Refresh feats list (super inefficient calling this for every added feat, but who cares; I'm in a hurry)
    InitBuildFeats
End Sub

Private Sub AddAlternateFeat(plngIndex As Long, plngFeat As Long, plngSelector As Long)
    Dim typNew As BuildFeatType
    Dim i As Long
    
    If plngFeat = 0 Then Exit Sub
    With Feat.List(plngIndex)
        typNew = build.Feat(.ActualType).Feat(.Index)
        typNew.Type = bftAlternate
        typNew.ChildType = .ActualType
        typNew.Child = .Index
        typNew.FeatName = db.Feat(plngFeat).FeatName
        typNew.Selector = plngSelector
    End With
    With build.Feat(bftAlternate)
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats) = typNew
    End With
End Sub

Private Sub AddExchangeFeat(pstrRaw As String, plngLevel As Long)
    Dim typNew As BuildFeatType
    Dim strRaw() As String
    Dim lngOldFeat As Long
    Dim lngOldSelector As Long
    Dim lngNewFeat As Long
    Dim lngNewSelector As Long
    Dim i As Long
    
    If Len(pstrRaw) = 0 Then Exit Sub
    strRaw = Split(pstrRaw, " replaces ")
    If UBound(strRaw) <> 1 Then Exit Sub
    If Not ParseFeat(strRaw(0), lngNewFeat, lngNewSelector) Then Exit Sub
    If Not ParseFeat(strRaw(1), lngOldFeat, lngOldSelector) Then Exit Sub
    For i = 1 To Feat.Count
        If Feat.List(i).FeatID = lngOldFeat And Feat.List(i).Selector = lngOldSelector Then Exit For
    Next
    If i > Feat.Count Then Exit Sub
    With Feat.List(i)
        typNew.Type = bftExchange
        typNew.ChildType = .ActualType
        typNew.Child = .Index
        typNew.FeatName = db.Feat(lngNewFeat).FeatName
        typNew.Selector = lngNewSelector
        typNew.Level = plngLevel
    End With
    With build.Feat(bftExchange)
        .Feats = .Feats + 1
        ReDim Preserve .Feat(1 To .Feats)
        .Feat(.Feats) = typNew
    End With
End Sub


' ************* SPELLS *************


Private Sub ParseSpells(pstrRaw As String)
    Dim strText As String
    Dim enClass As ClassEnum
    Dim enSlotType As SpellSlotEnum
    Dim strSpell() As String
    Dim strNew As String
    Dim lngSlots As Long
    Dim i As Long

    strText = StripBrowserFormatting(pstrRaw)
    If Len(strText) = 0 Then Exit Sub
    enClass = GetClassID(strText)
    If enClass <> ceAny Then
        menClassSpell = enClass
        mlngSpellLevel = 0
    ElseIf menClassSpell <> ceAny Then
        ' Spell list
        mlngSpellLevel = mlngSpellLevel + 1
        strSpell = Split(strText, ",")
        With build.Spell(menClassSpell).Level(mlngSpellLevel)
            lngSlots = .Slots
            i = 1
            Do While i <= lngSlots
                With .Slot(i)
                    enSlotType = .SlotType
                    strNew = Trim$(strSpell(i - 1))
                    If strNew = "<Any>" Then strNew = vbNullString
                    If Len(strNew) Then
                        Select Case enSlotType
                            Case sseStandard, sseFree
                                If CheckFree(menClassSpell, strNew) Then
                                    AddFreeSpellSlot menClassSpell, mlngSpellLevel
                                    lngSlots = lngSlots + 1
                                End If
                                .Spell = strNew
                            Case sseClericCure, sseWarlockPact
                                ' Ignore mandatory slots
                        End Select
                    End If
                End With
                i = i + 1
            Loop
        End With
    End If
End Sub


' ************* ENHANCEMENTS *************


Private Sub ParseEnhancements(pstrRaw As String)
    Dim strText As String
    Dim lngTree As Long
    Dim strList() As String
    Dim lngPos As Long
    Dim i As Long
    
    strText = StripBrowserFormatting(pstrRaw)
    If Len(strText) = 0 Then Exit Sub
    If Right$(strText, 4) = " AP)" Then
        lngPos = InStrRev(strText, " (")
        If lngPos = 0 Then Exit Sub
        strText = Left$(strText, lngPos - 1)
    End If
    lngTree = SeekTree(strText, peEnhancement)
    If lngTree <> 0 Then
        mlngTree = lngTree
        mlngTier = 0
        AddBuildTree strText
    ElseIf mlngTree <> 0 Then
        strList = Split(strText, ",")
        For i = 0 To UBound(strList)
            ParseAbility strList(i), db.Tree(mlngTree).Tier(mlngTier), build.Tree(build.Trees)
        Next
        If mlngTier = 5 Then build.Tier5 = db.Tree(mlngTree).TreeName
        mlngTier = mlngTier + 1
    End If
End Sub

Private Sub AddBuildTree(pstrRaw As String)
    Dim lngLevels As Long
    Dim typClassSplit() As ClassSplitType
    Dim lngClass As Long
    Dim enClass As ClassEnum
    Dim blnFound As Boolean
    Dim i As Long
    
    ' Class tree info
    If db.Tree(mlngTree).TreeType = tseClass Then
        For lngClass = 0 To GetClassSplit(typClassSplit) - 1
            enClass = typClassSplit(lngClass).ClassID
            For i = 1 To db.Class(enClass).Trees
                If db.Class(enClass).Tree(i) = db.Tree(mlngTree).TreeName Then
                    blnFound = True
                    Exit For
                End If
            Next
            If blnFound Then Exit For
        Next
        If blnFound Then
            For i = 1 To HeroicLevels()
                If build.Class(i) = enClass Then lngLevels = lngLevels + 1
            Next
        End If
    End If
    ' Add this buildtree
    With build
        .Trees = .Trees + 1
        ReDim Preserve .Tree(1 To .Trees)
        With .Tree(.Trees)
            .TreeName = db.Tree(mlngTree).TreeName
            .TreeType = db.Tree(mlngTree).TreeType
            .Source = enClass
            .ClassLevels = lngLevels
        End With
    End With
End Sub

Private Sub ParseAbility(pstrRaw As String, ptypTier As TierType, ptypBuildTree As BuildTreeType)
    Dim strAbility As String
    Dim strSelector As String
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long
    
    ' Two conflicts to worry about are that some ability names contain ":", some end with " I", " II" or " III"
    ' (Normally a ":" denotes a selector, and ending in " I", " II" or " III" denotes ranks)
    ' First check for a perfect match of raw text
    strAbility = pstrRaw
    strSelector = vbNullString
    lngRanks = 1
    If FindAbility(ptypTier, strAbility, strSelector, lngAbility, lngSelector) Then
        AddAbility ptypBuildTree, lngAbility, lngSelector, lngRanks
        Exit Sub
    End If
    ' If no match, see if we have rank info
    If GetRanks(strAbility, lngRanks) Then
        If FindAbility(ptypTier, strAbility, strSelector, lngAbility, lngSelector) Then
            AddAbility ptypBuildTree, lngAbility, lngSelector, lngRanks
            Exit Sub
        End If
    End If
    ' Still no match; see if we have a selector
    If GetSelector(strAbility, strSelector) Then
        If FindAbility(ptypTier, strAbility, strSelector, lngAbility, lngSelector) Then
            AddAbility ptypBuildTree, lngAbility, lngSelector, lngRanks
            Exit Sub
        End If
    End If
    ' No match
End Sub

Private Function FindAbility(ptypTier As TierType, pstrAbility As String, pstrSelector As String, plngAbility As Long, plngSelector As Long) As Boolean
    Dim blnFound As Boolean
    Dim i As Long
    Dim s As Long
    
    plngAbility = 0
    plngSelector = 0
    ' Perfect match?
    For i = 1 To ptypTier.Abilities
        If ptypTier.Ability(i).Abbreviation = pstrAbility Then
            plngAbility = i
            If Len(pstrSelector) Then
                For s = 1 To ptypTier.Ability(i).Selectors
                    If ptypTier.Ability(i).Selector(s).SelectorName = pstrSelector Then
                        plngSelector = s
                        Exit For
                    End If
                Next
            End If
            FindAbility = True
            Exit Function
        End If
    Next
    ' Selector Only?
    For i = 1 To ptypTier.Abilities
        For s = 1 To ptypTier.Ability(i).Selectors
            If ptypTier.Ability(i).Selector(s).SelectorName = pstrAbility Then
                plngAbility = i
                plngSelector = s
                FindAbility = True
                Exit Function
            End If
        Next
    Next
End Function

Private Sub AddAbility(ptypBuildTree As BuildTreeType, plngAbility As Long, plngSelector As Long, plngRanks As Long)
    With ptypBuildTree
        .Abilities = .Abilities + 1
        ReDim Preserve .Ability(1 To .Abilities)
        With .Ability(.Abilities)
            .Tier = mlngTier
            .Ability = plngAbility
            .Selector = plngSelector
            .Rank = plngRanks
        End With
    End With
End Sub

Private Function GetRanks(pstrAbility As String, plngRanks As Long) As Boolean
    If Right$(pstrAbility, 2) = " I" Then
        pstrAbility = Left$(pstrAbility, Len(pstrAbility) - 2)
    ElseIf Right$(pstrAbility, 3) = " II" Then
        plngRanks = 2
        pstrAbility = Left$(pstrAbility, Len(pstrAbility) - 3)
    ElseIf Right$(pstrAbility, 4) = " III" Then
        plngRanks = 3
        pstrAbility = Left$(pstrAbility, Len(pstrAbility) - 4)
    Else
        Exit Function
    End If
    GetRanks = True
End Function

Private Function GetSelector(pstrAbility As String, pstrSelector As String) As Boolean
    Dim lngPos As Long
    
    lngPos = InStr(pstrAbility, ":")
    If lngPos = 0 Then Exit Function
    pstrSelector = Mid$(pstrAbility, lngPos + 2)
    pstrAbility = Left$(pstrAbility, lngPos - 1)
    GetSelector = True
End Function


' ************* DESTINY *************


Private Sub ParseDestiny(ByVal pstrRaw As String)
    Dim strText As String
    Dim lngDestiny As Long
    Dim strList() As String
    Dim lngPos As Long
    Dim i As Long
    
    strText = StripBrowserFormatting(pstrRaw)
    If Len(strText) = 0 Then Exit Sub
    If Right$(strText, 4) = " AP)" Then
        lngPos = InStrRev(strText, " (")
        If lngPos = 0 Then Exit Sub
        strText = Left$(strText, lngPos - 1)
    End If
    ''Find our destiny id
    lngDestiny = SeekTree(strText, peDestiny)
    If lngDestiny <> 0 Then
        mlngDestiny = lngDestiny
        mlngDestinyTier = 0
        AddBuildDestiny strText
    ElseIf mlngDestiny <> 0 Then
        strList = Split(strText, ",")
        For i = 0 To UBound(strList)
            'TODO Fix
            ParseAbility strList(i), db.Destiny(mlngDestiny).Tier(mlngDestinyTier), build.Destiny(mlngDestiny)
        Next
        mlngDestinyTier = mlngDestinyTier + 1
    End If
End Sub

Private Sub AddBuildDestiny(pstrRaw As String)
    ' Add this build Destiny
    With build
        .Destinies = .Destinies + 1
        ReDim Preserve .Destiny(1 To .Destinies)
        With .Destiny(.Destinies)
            .TreeName = db.Destiny(mlngDestiny).TreeName
            .TreeType = db.Destiny(mlngDestiny).TreeType
            .Source = 0 ' not used for destiny
            .ClassLevels = 0 ' not used for destiny
        End With
    End With
End Sub




' ************* GENERAL *************


Private Function StripBrowserFormatting(pstrRaw As String) As String
    Dim strReturn As String
    
    ' Strip leading and trailing spaces
    strReturn = Trim$(pstrRaw)
    If Len(strReturn) < 3 Then Exit Function
    ' If there's two spaces in a row, that's from a "." inside a name that got stripped out. Put it back.
    If InStr(strReturn, "  ") Then strReturn = Replace(strReturn, "  ", ". ")
    ' Remove spaces after commas
    If InStr(strReturn, ", ") Then strReturn = Replace(strReturn, ", ", ",")
    ' Strip out Internet Explorer bullet points
    If Left$(strReturn, 1) = "•" Then strReturn = Mid$(strReturn, 2)
    If IsNumeric(Left$(strReturn, 1)) And Mid$(strReturn, 2, 1) = " " Then strReturn = Mid$(strReturn, 3)
    ' All done
    StripBrowserFormatting = Trim$(strReturn)
End Function

Private Sub StripSpaces(ByVal pstrRaw As String, pstrArray() As String)
    Do While InStr(pstrRaw, "  ")
        pstrRaw = Replace(pstrRaw, "  ", " ")
    Loop
    pstrArray = Split(pstrRaw, " ")
End Sub
