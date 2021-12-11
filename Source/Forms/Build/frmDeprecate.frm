VERSION 5.00
Begin VB.Form frmDeprecate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Messages"
   ClientHeight    =   7755
   ClientLeft      =   30
   ClientTop       =   390
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
   Icon            =   "frmDeprecate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin CharacterBuilderLite.userCheckBox usrchkGuess 
      Height          =   252
      Left            =   4380
      TabIndex        =   6
      Top             =   3720
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   450
      Caption         =   "Best Guess"
   End
   Begin VB.CheckBox chkSelector 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Choose"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1332
   End
   Begin CharacterBuilderLite.userDetails usrDetails 
      Height          =   3264
      Left            =   420
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4020
      Width           =   3612
      _ExtentX        =   6376
      _ExtentY        =   5768
   End
   Begin VB.CheckBox chkChoose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Choose"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1332
   End
   Begin VB.ListBox lstMessage 
      Appearance      =   0  'Flat
      Height          =   2340
      ItemData        =   "frmDeprecate.frx":000C
      Left            =   420
      List            =   "frmDeprecate.frx":000E
      TabIndex        =   2
      Top             =   900
      Width           =   11412
   End
   Begin VB.ListBox lstGuess 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   3180
      ItemData        =   "frmDeprecate.frx":0010
      Left            =   4320
      List            =   "frmDeprecate.frx":0012
      TabIndex        =   8
      Top             =   4020
      Width           =   3612
   End
   Begin VB.CheckBox chkIgnore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ignore"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1332
   End
   Begin CharacterBuilderLite.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   688
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin VB.ListBox lstSelector 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   3180
      ItemData        =   "frmDeprecate.frx":0014
      Left            =   8220
      List            =   "frmDeprecate.frx":0016
      TabIndex        =   11
      Top             =   4020
      Width           =   3612
   End
   Begin CharacterBuilderLite.userDetails usrInfo 
      Height          =   3264
      Left            =   8220
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4020
      Width           =   3612
      _ExtentX        =   6376
      _ExtentY        =   5768
   End
   Begin VB.Label lblGuess 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Selectors"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   4380
      TabIndex        =   13
      Top             =   3744
      Visible         =   0   'False
      Width           =   3552
   End
   Begin VB.Label lblDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Details"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   480
      TabIndex        =   3
      Top             =   3744
      Width           =   3552
   End
   Begin VB.Label lblSelectors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Selectors"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   8280
      TabIndex        =   9
      Top             =   3744
      Width           =   3552
   End
   Begin VB.Label lblIssues 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Build file contains the following issues. Click an issue to view details."
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   11352
   End
End
Attribute VB_Name = "frmDeprecate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum DeprecateGroupEnum
    dgeMain = 0
    dgeNameChange = 1000
    dgeFeat = 2000
    dgeEnhancement = 3000
    dgeGuide = 4000
    dgeDestiny = 5000
    dgeTwist = 6000
    dgeWarSoul = 7000
    dgeBinaryTree = 8000
    dgeBinaryDestiny = 9000
    dgeBinaryTwist = 10000
End Enum

Private menGroup As DeprecateGroupEnum

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    menGroup = dgeMain
    cfg.Configure Me
    ActivateForm
    ShowMessages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
End Sub

Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Messages"
    End Select
End Sub


' ************* SHOW MESSAGES *************


Private Sub ShowMessages()
    ListboxClear Me.lstMessage
    Me.usrDetails.Clear
    ListboxClear Me.lstGuess
    ListboxClear Me.lstSelector
    WriteLabel
    ShowIssues
    ShowNotices
    If Me.lstMessage.ListCount > 0 Then Me.lstMessage.ListIndex = 0
    EnableControls
End Sub

Private Sub WriteLabel()
    Dim lngIssues As Long
    Dim lngNotices As Long
    Dim strIssues As String
    Dim strNotices As String
    Dim strMessage As String
    
    CountMessages lngIssues, lngNotices
    If lngIssues = 1 Then strIssues = "1 Issue" Else strIssues = lngIssues & " Issues"
    If lngIssues Then strMessage = strIssues
    If lngNotices = 1 Then strNotices = "1 Message" Else strNotices = lngNotices & " Messages"
    If lngNotices Then
        If Len(strMessage) Then strMessage = strMessage & " and "
        strMessage = strMessage & strNotices
    End If
    If Len(strMessage) = 0 Then strMessage = "No issues"
    Me.lblIssues.Caption = strMessage
End Sub

Private Function CountMessages(Optional plngIssues As Long, Optional plngNotices As Long) As Long
    plngIssues = 0
    plngNotices = 0
    With gtypDeprecate
        ' Feats
        plngIssues = .Feats
        plngNotices = .NameChanges
        ' War Soul
        If .WarSoul Then plngNotices = plngNotices + 1
        If .DivineMight Then plngIssues = plngIssues + 1
        ' Enhancements
        plngNotices = plngNotices + .Trees + .Enhancements
        ' Leveling Guide
        If .LevelingGuide.Deprecated Then plngNotices = plngNotices + 1
        ' Destiny
        plngNotices = plngNotices + .Destinies
    End With
    CountMessages = plngIssues + plngNotices
End Function


' ************* SHOW ISSUES *************


Private Sub ShowIssues()
    Dim i As Long
    
    With gtypDeprecate
        ' Feats
        For i = 1 To .Feats
            ListboxAddItem Me.lstMessage, "Feats: " & FeatIssue(.Feat(i)), dgeFeat + i
        Next
        ' War Soul: Divine Might
        If .DivineMight Then ListboxAddItem Me.lstMessage, "Enhancements: War Soul Tier 1: Divine Might defaulted to Strength", dgeWarSoul + 2
    End With
End Sub

Private Function FeatIssue(ptypFeat As DeprecateFeatType) As String
    Dim strType As String
    Dim lngLevel As Long
    Dim strReturn As String
    
    With ptypFeat
        If .Index <= build.Feat(.Type).Feats Then
            lngLevel = build.Feat(.Type).Feat(.Index).Level
            strType = GetFeatTypeName(.Type, lngLevel)
        End If
        strReturn = .FeatName
        If Len(.SelectorName) Then strReturn = strReturn & ": " & .SelectorName
        strReturn = strReturn & " not found (taken as a"
        If InStr("aeiou", LCase$(Left$(strType, 1))) Then strReturn = strReturn & "n"
        strReturn = strReturn & " " & strType & " feat at level " & lngLevel & ")"
    End With
    FeatIssue = strReturn
End Function

Private Function GetFeatTypeName(ByVal penType As BuildFeatTypeEnum, ByVal plngLevel As Long) As String
    Dim strReturn As String
    
    Select Case penType
        Case bftStandard
            Select Case plngLevel
                Case 1, 3, 6, 9, 12, 15, 18: strReturn = "Heroic"
                Case 21, 24, 27, 30: strReturn = "Epic"
                Case 26, 28, 29: strReturn = "Destiny"
                Case Else: strReturn = "Standard"
            End Select
        Case bftLegend
            strReturn = "Legend"
        Case bftRace
            strReturn = "Racial"
        Case bftClass1, bftClass2, bftClass3
            If plngLevel > 0 And plngLevel < 21 Then
                strReturn = GetClassName(build.Class(plngLevel))
            Else
                strReturn = "Class"
            End If
        Case bftDeity
            strReturn = "Deity"
        Case bftAlternate
            strReturn = "Alternate"
        Case bftExchange
            strReturn = "Exchange"
    End Select
    GetFeatTypeName = strReturn
End Function


' ************* SHOW NOTICES *************


Private Sub ShowNotices()
    Dim strMessage As String
    Dim i As Long
    
    With gtypDeprecate
        ' Feats
        For i = 1 To .NameChanges
            ListboxAddItem Me.lstMessage, "Feats: " & NameChangeNotice(.NameChange(i)), dgeNameChange + i
        Next
        ' Warpriest => War Soul
        If .WarSoul Then ListboxAddItem Me.lstMessage, "Enhancements: Warpriest tree renamed to War Soul", dgeWarSoul + 1
        ' Enhancements
        For i = 1 To .Trees
            ListboxAddItem Me.lstMessage, "Enhancements: " & .Tree(i) & " was reset", dgeBinaryTree + i
        Next
        For i = 1 To .Enhancements
            ListboxAddItem Me.lstMessage, "Enhancements: " & AbilityNotice(.Enhancement(i)), dgeEnhancement + i
        Next
        ' Leveling guide
        If .LevelingGuide.Deprecated Then ListboxAddItem Me.lstMessage, "Leveling Guide: " & AbilityNotice(.LevelingGuide), dgeGuide + 1
        
        ' Destiny  TODO
        'If Len(.Destiny) Then ListboxAddItem Me.lstMessage, "Destiny: " & .Destiny & " was reset", dgeBinaryDestiny + 1
        'For i = 1 To .Destinies
        '    ListboxAddItem Me.lstMessage, "Destiny: " & AbilityNotice(.Destiny(i)), dgeDestiny + i
        'Next
    End With
End Sub

Private Function NameChangeNotice(plngIndex As Long) As String
    Dim strReturn As String
    
    With db.NameChange(plngIndex)
        strReturn = Chr(34) & .OldName & Chr(34) & " renamed to " & Chr(34) & .NewName & Chr(34)
    End With
    NameChangeNotice = strReturn
End Function

Private Function AbilityNotice(ptypAbility As DeprecateAbilityType) As String
    Dim strReturn As String
    
    With ptypAbility
        strReturn = .TreeName & " Tier " & .Tier & ": " & .AbilityName
        If Len(.SelectorName) Then strReturn = strReturn & ": " & .SelectorName
    End With
    AbilityNotice = strReturn & " not found"
End Function

Private Function BinaryTwistNotice(ptypTwist As TwistType) As String
    Dim strReturn As String
    
    With ptypTwist
        strReturn = .DestinyName & " Tier " & .Tier & ": Ability " & .Ability
        If .Selector <> 0 Then strReturn = strReturn & ": Selector " & .Selector
    End With
    BinaryTwistNotice = strReturn & " not found"
End Function

Private Sub lstMessage_Click()
    Dim enGroup As DeprecateGroupEnum
    Dim lngIndex As Long
    Dim blnPassive As Boolean
    Dim strCaption As String
    Dim strMessage As String
    
    GetMessage enGroup, lngIndex
    Me.usrDetails.Clear
    Me.usrInfo.Clear
    blnPassive = True
    strCaption = "OK"
    Select Case enGroup
        ' Feats
        Case dgeFeat
            ShowFeatDetails lngIndex
            ListFeats
            strCaption = "Ignore"
            blnPassive = False
        Case dgeNameChange
            Me.usrDetails.AddText "Name changes are automatic."
            Me.usrDetails.AddText "Nothing to resolve."
        ' Enhancements
        Case dgeEnhancement
            ShowAbilityDetails gtypDeprecate.Enhancement(lngIndex), "Enhancement not found or invalid:"
        Case dgeBinaryTree
            Me.usrDetails.AddText gtypDeprecate.Tree(lngIndex) & " was reset due to an invalid enhancement."
        Case dgeWarSoul
            Select Case lngIndex
                Case 1, 3
                    Me.usrDetails.AddText "Name changes are automatic."
                    Me.usrDetails.AddText "Nothing to resolve."
                Case 2, 4
                    ShowDivineMightSelectors
                    strCaption = "Ignore"
                    blnPassive = False
            End Select
        ' Leveling Guide
        Case dgeGuide
            If gtypDeprecate.LevelingGuide.Binary Then
                strMessage = "Leveling Guide reset due to invalid ability:"
            Else
                strMessage = "Leveling Guide truncated at invalid ability:"
            End If
            ShowAbilityDetails gtypDeprecate.LevelingGuide, strMessage
        ' Destiny
        Case dgeDestiny
            ShowAbilityDetails gtypDeprecate.DestinyType(lngIndex), "Destiny ability not found or invalid:"
    End Select
    Me.chkIgnore.Caption = strCaption
    ShowControls enGroup, blnPassive
    EnableControls
    Me.usrDetails.Refresh
End Sub

Private Sub GetMessage(penGroup As DeprecateGroupEnum, plngIndex As Long)
    Dim lngItemData As Long
    
    lngItemData = ListboxGetValue(Me.lstMessage)
    penGroup = (lngItemData \ 1000) * 1000
    plngIndex = lngItemData Mod 1000
End Sub

Private Sub ShowAbilityDetails(ptypAbility As DeprecateAbilityType, Optional pstrHeader As String)
    If Len(pstrHeader) Then
        Me.usrDetails.AddText pstrHeader
        Me.usrDetails.AddText vbNullString
    End If
    With ptypAbility
        Me.usrDetails.AddText "Tree: " & .TreeName
        If Len(.GuideTreeDisplay) Then Me.usrDetails.AddText "TreeID: " & .GuideTreeDisplay
        Me.usrDetails.AddText "Tier: " & .Tier
        Me.usrDetails.AddText "Ability: " & .AbilityName
        If Len(.SelectorName) Then
            Me.usrDetails.AddText "Selector: " & .SelectorName
        Else
            Me.usrDetails.AddText "Selector: [None]"
        End If
        Me.usrDetails.AddText "Rank: " & .Rank
    End With
End Sub

Private Sub ShowControls(penGroup As DeprecateGroupEnum, pblnPassive As Boolean)
    Dim blnFeat As Boolean
    Dim blnWarSoul As Boolean
    
    If pblnPassive Then
        Me.lstGuess.Clear
        Me.lstSelector.Clear
    End If
    Select Case penGroup
        Case dgeFeat
            blnFeat = True
            Me.lblSelectors.Caption = "Selectors"
        Case dgeNameChange
            blnFeat = True
        Case dgeWarSoul
            blnWarSoul = True
            Me.lblGuess.Caption = "Selectors"
            Me.lblSelectors.Caption = "Description"
    End Select
    Me.chkChoose.Visible = Not pblnPassive
    Me.usrchkGuess.Visible = blnFeat And Not pblnPassive
    Me.lblGuess.Visible = blnWarSoul And Not pblnPassive
    Me.chkSelector.Visible = blnFeat And Not pblnPassive
    Me.lblSelectors.Visible = Not pblnPassive
    Me.lstSelector.Visible = blnFeat
    If Me.lstGuess.ListCount > 0 Then Me.lstGuess.ListIndex = 0
End Sub

Private Sub EnableControls()
    Dim lngIssue As Long
    Dim blnEnabled As Boolean
    
    lngIssue = ListboxGetValue(Me.lstMessage)
    blnEnabled = lngIssue
    
    EnableLabel Me.lblDetails, blnEnabled
    Me.chkIgnore.Enabled = blnEnabled
    
    blnEnabled = (Me.lstGuess.ListCount > 0)
    Me.usrchkGuess.Enabled = blnEnabled
    Me.chkChoose.Enabled = (Me.lstGuess.ListIndex <> -1)
    Me.lstGuess.Enabled = blnEnabled
    
    blnEnabled = (Me.lstSelector.ListIndex <> -1)
    Me.chkSelector.Enabled = blnEnabled
    
    blnEnabled = (Me.lstSelector.ListCount > 0)
    EnableLabel Me.lblSelectors, blnEnabled Or (Me.lblSelectors.Caption <> "Selectors")
    Me.lstSelector.Enabled = blnEnabled
End Sub

Private Sub EnableLabel(plbl As Label, pblnEnabled As Boolean)
    If pblnEnabled Then plbl.ForeColor = cfg.GetColor(cgeWorkspace, cveText) Else plbl.ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
End Sub


' ************* FEATS *************


Private Sub ListFeats()
    Dim enGroup As DeprecateGroupEnum
    Dim lngIndex As Long
    Dim i As Long
    
    ListboxClear Me.lstGuess
    If Me.usrchkGuess.Value Then
        GetMessage enGroup, lngIndex
        FindSimilarFeats gtypDeprecate.Feat(lngIndex).FeatName
    Else
        For i = 1 To db.Feats
            With db.Feat(i)
                If .Selectable Then ListboxAddItem Me.lstGuess, .FeatName, i
            End With
        Next
        Me.lstGuess.ListIndex = -1
    End If
End Sub

Private Function FindSimilarFeats(ByVal pstrFeat As String)
    Dim lngDistance() As Long
    Dim lngPos As Long
    Dim i As Long
    
    ReDim lngDistance(1 To 2, 1 To db.Feats)
    lngPos = InStr(pstrFeat, ":")
    If lngPos Then pstrFeat = Left$(pstrFeat, lngPos - 1)
    For i = 1 To db.Feats
        lngDistance(1, i) = i
        lngDistance(2, i) = LevenshteinDistance(pstrFeat, db.Feat(i).FeatName)
    Next
    QuickSort2 lngDistance, 1, 2
    For i = 1 To 10
        If lngDistance(2, i) > lngDistance(2, 1) * 4 Then Exit For
        ListboxAddItem Me.lstGuess, db.Feat(lngDistance(1, i)).FeatName, lngDistance(1, i)
    Next
    Me.lstGuess.ListIndex = 0
End Function

Private Sub ShowFeatDetails(plngIndex As Long)
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim strType As String
    Dim lngFeat As Long
    Dim typFeat As BuildFeatType
    
    enType = gtypDeprecate.Feat(plngIndex).Type
    lngIndex = gtypDeprecate.Feat(plngIndex).Index
    If build.Feat(enType).Feats < lngIndex Then Exit Sub
    typFeat = build.Feat(enType).Feat(lngIndex)
    With gtypDeprecate.Feat(plngIndex)
        lngFeat = SeekFeat(.FeatName, False)
        If lngFeat Then Me.usrDetails.AddText "Selector not found" Else Me.usrDetails.AddText "Feat not found"
        Me.usrDetails.AddText vbNullString
        Me.usrDetails.AddText "Feat: " & .FeatName
        If lngFeat Then
            If db.Feat(lngFeat).Selectors > 0 Then
                If Len(.SelectorName) Then
                    Me.usrDetails.AddText "Selector: " & .SelectorName
                ElseIf .Selector <> 0 Then
                    Me.usrDetails.AddText "Selector: Selector #" & .Selector
                Else
                    Me.usrDetails.AddText "Selector: Unknown"
                End If
                Me.usrDetails.AddText vbNullString
            End If
        End If
        strType = GetFeatTypeName(enType, typFeat.Level)
        Me.usrDetails.AddText "Type: " & strType
        Me.usrDetails.AddText "Level: " & typFeat.Level
        Select Case .Type
            Case bftClass1, bftClass2, bftClass3
                Me.usrDetails.AddText strType & " Level: " & typFeat.ClassLevel
            Case bftAlternate, bftExchange
                If build.Feat(.Child).Feats >= .ChildIndex Then
                    With build.Feat(.Child).Feat(.ChildIndex)
                        Me.usrDetails.AddText vbNullString
                        Me.usrDetails.AddText "Parent Slot"
                        strType = GetFeatTypeName(.Type, .Level)
                        Me.usrDetails.AddText "Type: " & strType
                        Me.usrDetails.AddText "Level: " & .Level
                        Select Case .Type
                            Case bftClass1, bftClass2, bftClass3: Me.usrDetails.AddText strType & " Level: " & .ClassLevel
                        End Select
                        Me.usrDetails.AddText "Feat: " & .FeatName
                        lngFeat = SeekFeat(.FeatName)
                        If lngFeat > 0 And .Selector > 0 Then Me.usrDetails.AddText "Selector: " & db.Feat(lngFeat).Selector(.Selector).SelectorName
                    End With
                End If
        End Select
    End With
End Sub

Private Sub ChooseFeat(plngIndex As Long)
    Dim strMessage As String
    Dim i As Long
    
    If Me.lstGuess.ListIndex = -1 Then
        strMessage = "Choose a feat."
    ElseIf Me.lstSelector.ListCount > 0 And Me.lstSelector.ListIndex = -1 Then
        strMessage = "Choose a selector."
    End If
    If Len(strMessage) Then
        Notice strMessage
        Exit Sub
    End If
    With gtypDeprecate.Feat(plngIndex)
        If build.Feat(.Type).Feats >= .Index Then
            With build.Feat(.Type).Feat(.Index)
                .FeatName = Me.lstGuess.Text
                .Selector = Me.lstSelector.ListIndex + 1
            End With
        End If
    End With
    UpdateBuild cceFeat
End Sub


' ************* WAR SOUL *************


Private Sub ShowDivineMightSelectors()
    Me.usrDetails.AddText "Choose a Divine Might version."
    ListboxClear Me.lstGuess
    ListboxAddItem Me.lstGuess, "Divine Might", 1
    ListboxAddItem Me.lstGuess, "Divine Presence", 2
    ListboxAddItem Me.lstGuess, "Divine Will", 3
End Sub

Private Sub ChooseDivineMight(plngIndex As Long)
    Dim lngTree As Long
    Dim lngSelector As Long
    Dim i As Long
    
    lngSelector = ListboxGetValue(Me.lstGuess)
    If lngSelector < 1 Or lngSelector > 3 Then
        Notice "Invalid selector"
        Exit Sub
    End If
    ' Enhancements
    For lngTree = 1 To build.Trees
        If build.Tree(lngTree).TreeName = "War Soul" Then
            For i = 1 To build.Tree(lngTree).Abilities
                With build.Tree(lngTree).Ability(i)
                    If .Tier = 1 And .Ability = 1 Then .Selector = lngSelector
                End With
            Next
        End If
    Next
    ' Leveling Guide
    For lngTree = 1 To build.Guide.Trees
        If build.Guide.Tree(lngTree).TreeName = "War Soul" Then Exit For
    Next
    If lngTree > build.Guide.Trees Then
        Notice "Guide tree not found."
        Exit Sub
    End If
    For i = 1 To build.Guide.Enhancements
        With build.Guide.Enhancement(i)
            If .ID = lngTree And .Tier = 1 And .Ability = 1 Then .Selector = lngSelector
        End With
    Next
    gtypDeprecate.DivineMight = False
    UpdateBuild cceEnhancements
End Sub


' ************* RESOLVE *************


Private Sub usrchkGuess_UserChange()
    ListFeats
End Sub

Private Sub lstGuess_Click()
    Dim enGroup As DeprecateGroupEnum
    Dim lngIndex As Long
    
    GetMessage enGroup, lngIndex
    Select Case enGroup
        Case dgeFeat: ClickFeat
        Case dgeWarSoul: ClickDivineMight
    End Select
    EnableControls
End Sub

Private Sub ClickFeat()
    Dim lngFeat As Long
    Dim i As Long
    
    Me.chkChoose.Enabled = (Me.lstGuess.ListIndex <> -1)
    ListboxClear Me.lstSelector
    lngFeat = ListboxGetValue(Me.lstGuess)
    If lngFeat Then
        For i = 1 To db.Feat(lngFeat).Selectors
            ListboxAddItem Me.lstSelector, db.Feat(lngFeat).Selector(i).SelectorName, i
        Next
    End If
    Me.chkChoose.Enabled = (Me.lstGuess.ListIndex <> -1)
End Sub

Private Sub ClickDivineMight()
    Me.lblSelectors.Caption = Me.lstGuess.Text
    Me.usrInfo.Clear
    Me.usrInfo.AddText "Cost: 21/18/15 spell points"
    Me.usrInfo.AddText "Duration: 30/60/120 second"
    Me.usrInfo.AddText "Cooldown: 20 seconds"
    Me.usrInfo.AddText vbNullString
    Select Case ListboxGetValue(Me.lstGuess)
        Case 0: Me.usrInfo.Clear
        Case 1: Me.usrInfo.AddText "You gain an Insight bonus to Strength equal to your Charisma Modifier."
        Case 2: Me.usrInfo.AddText "You gain an Insight bonus to melee damage and the DC of tactical feats equal to 1/2 of your Charisma modifier."
        Case 3: Me.usrInfo.AddText "You gain an Insight bonus to melee damage and the DC of tactical feats equal to 1/2 of your Wisdom modifier."
    End Select
    Me.usrInfo.Refresh
End Sub

Private Sub lstSelector_Click()
    Me.chkSelector.Enabled = (Me.lstSelector.ListIndex <> -1)
End Sub

Private Sub chkIgnore_Click()
    If UncheckButton(Me.chkIgnore, mblnOverride) Then Exit Sub
    ResolveMessage
    ShowMessages
End Sub

Private Sub chkChoose_Click()
    Dim enGroup As DeprecateGroupEnum
    Dim lngIndex As Long
    
    If UncheckButton(Me.chkChoose, mblnOverride) Then Exit Sub
    GetMessage enGroup, lngIndex
    Select Case enGroup
        Case dgeFeat: ChooseFeat lngIndex
        Case dgeWarSoul: ChooseDivineMight lngIndex
    End Select
End Sub

Private Sub chkSelector_Click()
    Dim enGroup As DeprecateGroupEnum
    Dim lngIndex As Long
    
    If UncheckButton(Me.chkSelector, mblnOverride) Then Exit Sub
    GetMessage enGroup, lngIndex
    Select Case enGroup
        Case dgeFeat: ChooseFeat lngIndex
    End Select
End Sub

Private Sub UpdateBuild(penCascade As CascadeChangeEnum)
    CascadeChanges penCascade
    ResolveMessage
    ShowMessages
    SetDirty True
    SetAppCaption
End Sub

Private Sub ResolveMessage()
    Dim enGroup As DeprecateGroupEnum
    Dim lngIndex As Long
    Dim typBlank As DeprecateAbilityType
    Dim i As Long

    GetMessage enGroup, lngIndex
    With gtypDeprecate
        Select Case enGroup
            ' Feats
            Case dgeFeat
                For i = lngIndex To .Feats - 1
                    .Feat(i) = .Feat(i + 1)
                Next
                .Feats = .Feats - 1
                If .Feats = 0 Then Erase .Feat Else ReDim Preserve .Feat(1 To .Feats)
            Case dgeNameChange
                For i = lngIndex To .NameChanges - 1
                    .NameChange(i) = .NameChange(i + 1)
                Next
                .NameChanges = .NameChanges - 1
                If .NameChanges = 0 Then Erase .NameChange Else ReDim Preserve .NameChange(1 To .NameChanges)
            ' Enhancements
            Case dgeEnhancement
                For i = lngIndex To .Enhancements - 1
                    .Enhancement(i) = .Enhancement(i + 1)
                Next
                .Enhancements = .Enhancements - 1
                If .Enhancements = 0 Then Erase .Enhancement Else ReDim Preserve .Enhancement(1 To .Enhancements)
            Case dgeBinaryTree
                For i = lngIndex To .Trees - 1
                    .Tree(i) = .Tree(i + 1)
                Next
                .Trees = .Trees - 1
                If .Trees = 0 Then Erase .Tree Else ReDim Preserve .Tree(1 To .Trees)
            Case dgeWarSoul
                Select Case lngIndex
                    Case 1: .WarSoul = False
                    Case 2: .DivineMight = False
                End Select
            Case dgeGuide
                .LevelingGuide = typBlank
            ' Destiny
            Case dgeDestiny
                For i = lngIndex To .Destinies - 1
                    .Destiny(i) = .Destiny(i + 1)
                Next
                .Destinies = .Destinies - 1
                If .Destinies = 0 Then Erase .Destiny Else ReDim Preserve .Destiny(1 To .Destinies)
        End Select
    End With
    Me.usrDetails.Clear
    ListboxClear Me.lstGuess
    ListboxClear Me.lstSelector
    If CountMessages() = 0 Then
        gtypDeprecate.Deprecated = False
        frmMain.UpdateToolsMenu
        Unload Me
    End If
End Sub


' ************* LEVENSHTEIN DISTANCE *************


' This is not my code. Within minutes of googling for how to compare similar strings,
' I found the following page on stackoverflow:
' https://stackoverflow.com/questions/5859561/getting-the-closest-string-match
' The answer provided by Alain surprisingly included actual VB6 code. While
' his answer was far more involved, just this first function is sufficient for my needs.
' This is his function as posted, unmodified in any way. All comments are his.
' - Ellis

'Calculate the Levenshtein Distance between two strings (the number of insertions,
'deletions, and substitutions needed to transform the first string into the second)
Public Function LevenshteinDistance(ByRef S1 As String, ByVal S2 As String) As Long
    Dim L1 As Long, L2 As Long, D() As Long 'Length of input strings and distance matrix
    Dim i As Long, j As Long, Cost As Long 'loop counters and cost of substitution for current letter
    Dim cI As Long, cD As Long, cS As Long 'cost of next Insertion, Deletion and Substitution
    L1 = Len(S1): L2 = Len(S2)
    ReDim D(0 To L1, 0 To L2)
    For i = 0 To L1: D(i, 0) = i: Next i
    For j = 0 To L2: D(0, j) = j: Next j

    For j = 1 To L2
        For i = 1 To L1
            Cost = Abs(StrComp(Mid$(S1, i, 1), Mid$(S2, j, 1), vbTextCompare))
            cI = D(i - 1, j) + 1
            cD = D(i, j - 1) + 1
            cS = D(i - 1, j - 1) + Cost
            If cI <= cD Then 'Insertion or Substitution
                If cI <= cS Then D(i, j) = cI Else D(i, j) = cS
            Else 'Deletion or Substitution
                If cD <= cS Then D(i, j) = cD Else D(i, j) = cS
            End If
        Next i
    Next j
    LevenshteinDistance = D(L1, L2)
End Function


' ************* SORTING *************


' Sort a 2-dimensional array on either dimension
' Omit plngLeft & plngRight; they are used internally during recursion
' Sample usage to sort on column 4
' Dim MyArray(1 to 1000, 1 to 5) As Long
' QuickSort2 MyArray, 2, 4
' Dim MyArray(1 to 5, 1 to 1000) As Long
' QuickSort2 MyArray, 1, 4
Private Sub QuickSort2(ByRef pvarArray As Variant, plngDim As Long, plngCol As Long, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    Dim c As Long
    Dim cMin As Long
    Dim cMax As Long
    
    cMin = LBound(pvarArray, plngDim)
    cMax = UBound(pvarArray, plngDim)
    Select Case plngDim
        Case 1
            If plngRight = 0 Then
                plngLeft = LBound(pvarArray, 2)
                plngRight = UBound(pvarArray, 2)
            End If
            lngFirst = plngLeft
            lngLast = plngRight
            varMid = pvarArray(plngCol, (plngLeft + plngRight) \ 2)
            Do
                Do While pvarArray(plngCol, lngFirst) < varMid And lngFirst < plngRight
                    lngFirst = lngFirst + 1
                Loop
                Do While varMid < pvarArray(plngCol, lngLast) And lngLast > plngLeft
                    lngLast = lngLast - 1
                Loop
                If lngFirst <= lngLast Then
                    For c = cMin To cMax
                        varSwap = pvarArray(c, lngFirst)
                        pvarArray(c, lngFirst) = pvarArray(c, lngLast)
                        pvarArray(c, lngLast) = varSwap
                    Next
                    lngFirst = lngFirst + 1
                    lngLast = lngLast - 1
                End If
            Loop Until lngFirst > lngLast
            If plngLeft < lngLast Then QuickSort2 pvarArray, plngDim, plngCol, plngLeft, lngLast
            If lngFirst < plngRight Then QuickSort2 pvarArray, plngDim, plngCol, lngFirst, plngRight
        Case 2
            If plngRight = 0 Then
                plngLeft = LBound(pvarArray, 1)
                plngRight = UBound(pvarArray, 1)
            End If
            lngFirst = plngLeft
            lngLast = plngRight
            varMid = pvarArray((plngLeft + plngRight) \ 2, plngCol)
            Do
                Do While pvarArray(lngFirst, plngCol) < varMid And lngFirst < plngRight
                    lngFirst = lngFirst + 1
                Loop
                Do While varMid < pvarArray(lngLast, plngCol) And lngLast > plngLeft
                    lngLast = lngLast - 1
                Loop
                If lngFirst <= lngLast Then
                    For c = cMin To cMax
                        varSwap = pvarArray(lngFirst, c)
                        pvarArray(lngFirst, c) = pvarArray(lngLast, c)
                        pvarArray(lngLast, c) = varSwap
                    Next
                    lngFirst = lngFirst + 1
                    lngLast = lngLast - 1
                End If
            Loop Until lngFirst > lngLast
            If plngLeft < lngLast Then QuickSort2 pvarArray, plngDim, plngCol, plngLeft, lngLast
            If lngFirst < plngRight Then QuickSort2 pvarArray, plngDim, plngCol, lngFirst, plngRight
    End Select
End Sub
