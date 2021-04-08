VERSION 5.00
Begin VB.Form frmDestiny 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Destiny"
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
   Icon            =   "frmDestiny.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12225
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
      LeftLinks       =   "Destiny;Twists of Fate"
      RightLinks      =   "Help"
   End
   Begin CharacterBuilderLite.userList usrList 
      Height          =   6492
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   4932
      _ExtentX        =   8705
      _ExtentY        =   11456
   End
   Begin VB.ComboBox cboDestiny 
      Height          =   312
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   3492
   End
   Begin VB.ListBox lstAbility 
      Appearance      =   0  'Flat
      Height          =   3012
      IntegralHeight  =   0   'False
      ItemData        =   "frmDestiny.frx":000C
      Left            =   5520
      List            =   "frmDestiny.frx":000E
      TabIndex        =   8
      Top             =   1560
      Width           =   3492
   End
   Begin VB.ListBox lstSub 
      Appearance      =   0  'Flat
      Height          =   3012
      IntegralHeight  =   0   'False
      ItemData        =   "frmDestiny.frx":0010
      Left            =   9240
      List            =   "frmDestiny.frx":0012
      TabIndex        =   10
      Top             =   1560
      Width           =   2652
   End
   Begin CharacterBuilderLite.userCheckBox usrchkShowAll 
      Height          =   252
      Left            =   7560
      TabIndex        =   7
      Top             =   1296
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   450
      Value           =   0   'False
      Caption         =   "Show All"
      CheckPosition   =   1
   End
   Begin CharacterBuilderLite.userDetails usrDetails 
      Height          =   2040
      Left            =   5520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4932
      Width           =   6372
      _ExtentX        =   9128
      _ExtentY        =   3598
   End
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7380
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "< Enhancements"
   End
   Begin CharacterBuilderLite.userSpinner usrspnDestinyAP 
      Height          =   300
      Left            =   10000
      TabIndex        =   17
      Top             =   7020
      Width           =   972
      _ExtentX        =   1720
      _ExtentY        =   529
      Min             =   0
      Value           =   0
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Label lblTome 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Tome"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   11205
      TabIndex        =   18
      Top             =   7020
      Width           =   510
   End
   Begin VB.Label lblSpent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fate Points: 15"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   2400
      TabIndex        =   3
      Top             =   7020
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0 / 24 AP"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   480
      TabIndex        =   2
      Top             =   7020
      Width           =   1600
   End
   Begin VB.Label lblProg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "12 AP spent in tree"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   7680
      TabIndex        =   15
      Top             =   7020
      Visible         =   0   'False
      Width           =   1704
   End
   Begin VB.Label lblRanks 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ranks: 3"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   4240
      TabIndex        =   13
      Top             =   7020
      Visible         =   0   'False
      Width           =   804
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cost: 2 AP per rank"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   5480
      TabIndex        =   14
      Top             =   7020
      Visible         =   0   'False
      Width           =   1776
   End
   Begin VB.Label lblDestiny 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Destiny"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   5520
      TabIndex        =   4
      Top             =   600
      Width           =   3444
   End
   Begin VB.Label lblAbilities 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Abilities"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5520
      TabIndex        =   6
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label lblSelectors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Selectors"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   9240
      TabIndex        =   9
      Top             =   1320
      Width           =   2652
   End
   Begin VB.Label lblDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Details"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5520
      TabIndex        =   11
      Top             =   4680
      Width           =   6372
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "Clear this slot"
         Index           =   0
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Clear all slots"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmDestiny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private mblnTwists As Boolean

Private mblnMouse As Boolean
Private mlngSourceIndex As Long
Private menDragState As DragEnum
Private mblnDragComplete As Boolean
Private msngDownX As Single
Private msngDownY As Single

Private mblnOverride As Boolean

' Max Destiny. 24 + Tomes
Const MaxDestinyAP = 24

' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.Configure Me
    mblnTwists = False
    LoadData
    ShowAvailable False
    ComboListHeight Me.cboDestiny, 20
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Activate()
    ActivateForm oeDestiny
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    UnloadForm Me, mblnOverride
End Sub

Public Sub Cascade()
    ShowAbilities
    ShowAvailable True
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    If IsOver(Me.usrDetails.hwnd, Xpos, Ypos) Then Me.usrDetails.Scroll lngValue
    
    If IsOver(Me.usrspnDestinyAP.hwnd, Xpos, Ypos) Then Me.usrspnDestinyAP.WheelScroll lngValue
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Destiny": TabClick False
        Case "Twists of Fate": TabClick True
        Case "Help": ShowHelp "Destiny"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "< Enhancements": OpenForm "frmEnhancements"
        Case "Gear >": OpenForm "frmGear"
    End Select
    mblnOverride = True
    Unload Me
End Sub

Private Sub TabClick(pblnTwists As Boolean)
    xp.LockWindow Me.hwnd
    NoSelection
    mblnOverride = True
    mblnTwists = pblnTwists
    Me.usrList.Clear
    LoadData
    ShowAbilities
    ShowAvailable False
    mblnOverride = False
    xp.UnlockWindow
    SaveBackup
End Sub


' ************* INITIALIZE *************


Private Sub LoadData()
    Dim lngRows As Long
    
    mlngSourceIndex = 0
    Me.usrDetails.Clear
    With Me.usrList
        If mblnTwists Then
            lngRows = MaxTwistSlots()
            .DefineDimensions 4, 3, 2
            .DefineColumn 1, vbCenter, "Tier", "Tier"
            .DefineColumn 2, vbCenter, "Ability"
            .DefineColumn 3, vbRightJustify, "Fate Points", " 20"
        Else
            lngRows = build.Destiny.Abilities
            If lngRows = 0 Then lngRows = 1
            .DefineDimensions lngRows, 3, 2
            .DefineColumn 1, vbCenter, "Tier", "Tier"
            .DefineColumn 2, vbCenter, "Ability"
            .DefineColumn 3, vbCenter, "AP", " 20"
        End If
        .Refresh
    End With
    
    Me.usrspnDestinyAP.Max = tomes.DestinyMax
    Me.usrspnDestinyAP.Value = build.DestinyTome
    PopulateCombo
    If Not mblnTwists Then ComboSetText Me.cboDestiny, build.Destiny.TreeName
End Sub

Private Sub PopulateCombo()
    Dim strFirst As String
    Dim i As Long
    
    ListboxClear Me.lstSub
    ListboxClear Me.lstAbility
    ComboClear Me.cboDestiny
    If mblnTwists Then strFirst = "All Destinies"
    ComboAddItem Me.cboDestiny, strFirst, 0
    For i = 1 To db.Destinies
        If Not (mblnTwists And build.Destiny.TreeName = db.Destiny(i).TreeName) Then ComboAddItem Me.cboDestiny, db.Destiny(i).TreeName, i
    Next
    If Len(build.Destiny.TreeName) And Not mblnTwists Then ComboSetText Me.cboDestiny, build.Destiny.TreeName
End Sub


' ************* DISPLAY *************


Private Sub ShowAbilities()
    If mblnTwists Then ShowTwists dsDefault Else ShowDestinyAbilities
End Sub

Private Sub ShowDestinyAbilities()
    Dim lngDestiny As Long
    Dim strCaption As String
    Dim lngCost As Long
    Dim lngRanks As Long
    Dim lngMaxRanks As Long
    Dim lngTotal As Long
    Dim lngSpent() As Long
    Dim lngColor As Long
    Dim i As Long
    
    If build.Destiny.Abilities Then
        Me.usrList.Rows = build.Destiny.Abilities
    Else
        Me.usrList.Rows = 1
        SetSlot 1, vbNullString, vbNullString, vbNullString, 0, 0
        Me.usrList.SetError 1, False
        Me.usrList.SetDropState 1, dsDefault
    End If
    lngDestiny = ComboGetValue(Me.cboDestiny)
    If lngDestiny > 0 Then
        GetSpentInTree db.Destiny(lngDestiny), build.Destiny, lngSpent, lngTotal
        For i = 1 To build.Destiny.Abilities
            If build.Destiny.Ability(i).Ability = 0 Then
                SetSlot i, vbNullString, vbNullString, vbNullString, 0, 0
            Else
                GetSlotInfo db.Destiny(lngDestiny), build.Destiny.Ability(i), strCaption, lngCost, lngRanks, lngMaxRanks
                SetSlot i, build.Destiny.Ability(i).Tier, lngCost, strCaption, lngRanks, lngMaxRanks
                If mblnTwists Then Me.usrList.SetError i, False Else Me.usrList.SetError i, CheckErrors(db.Destiny(lngDestiny), build.Destiny.Ability(i), lngSpent)
            End If
        Next
    End If
    'Set bottom label visibility
    Me.lblTotal.Visible = True
    Me.usrspnDestinyAP.Visible = True
    If lngTotal > (MaxDestinyAP + Me.usrspnDestinyAP.Value) Then
        lngColor = cfg.GetColor(cgeWorkspace, cveTextError)
    Else
        lngColor = cfg.GetColor(cgeWorkspace, cveText)
    End If
    Me.lblTotal.Caption = lngTotal & " / " & (MaxDestinyAP + Me.usrspnDestinyAP.Value) & "AP"
    Me.lblTotal.ForeColor = lngColor
    Me.lblSpent.Visible = False
End Sub

Private Sub ShowTwists(penDropState As DropStateEnum, Optional ByVal plngSource As Long = 0)
    Dim lngFate As Long
    Dim lngTotal As Long
    Dim lngRows As Long
    Dim enDropState As DropStateEnum
    Dim i As Long
    
    If plngSource = 0 Then lngRows = build.Twists + 1 Else lngRows = build.Twists
    If lngRows < 1 Then lngRows = 1
    If lngRows > MaxTwistSlots() Then lngRows = MaxTwistSlots()
    Me.usrList.Rows = lngRows
    For i = 1 To lngRows
        Me.usrList.SetError i, False
        If i = plngSource Then enDropState = dsDefault Else enDropState = penDropState
        Me.usrList.SetDropState i, enDropState
        If i > build.Twists Then
            Me.usrList.SetSlot i, vbNullString, 0, 0
            Me.usrList.SetText i, 1, vbNullString
            Me.usrList.SetText i, 3, vbNullString
        Else
            Me.usrList.SetSlot i, GetTwistDisplayName(i), 0, 0
            Me.usrList.SetError i, CheckTwistSlot(i)
            lngFate = CalculateFatePoints(i, build.Twist(i).Tier)
            lngTotal = lngTotal + lngFate
            Me.usrList.SetText i, 1, build.Twist(i).Tier
            Me.usrList.SetText i, 3, lngFate
        End If
    Next
    'Set bottom label visibility
    Me.lblTotal.Visible = False
    Me.usrspnDestinyAP.Visible = False
    With Me.lblSpent
        .Caption = "Fate Points: " & lngTotal
        If lngTotal > MaxFatePoints() Then .ForeColor = cfg.GetColor(cgeWorkspace, cveTextError) Else .ForeColor = cfg.GetColor(cgeWorkspace, cveText)
        .Visible = True
    End With
End Sub

' Light up valid drop locations during drag operations
Private Sub ShowDropSlots()
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim enDropState As DropStateEnum
    Dim lngSpent() As Long
    Dim typCheck As BuildAbilityType
    Dim i As Long
    
    lngDestiny = ComboGetValue(Me.cboDestiny)
    GetSpentInTree db.Destiny(lngDestiny), build.Destiny, lngSpent, 0
    With build.Destiny
        For i = 1 To .Abilities
            If .Ability(i).Ability <> 0 Then
                enDropState = dsDefault
            Else
                GetUserChoices lngDestiny, lngTier, lngAbility, lngSelector
                With typCheck
                    .Tier = lngTier
                    .Ability = lngAbility
                    .Selector = lngSelector
                    .Rank = 1
                End With
                If CheckErrors(db.Destiny(lngDestiny), typCheck, lngSpent) Then enDropState = dsCanDropError Else enDropState = dsCanDrop
            End If
            Me.usrList.SetDropState i, enDropState
        Next
    End With
End Sub

' Returns TRUE if errors found
Private Function CheckErrors(ptypTree As TreeType, ptypAbility As BuildAbilityType, plngSpent() As Long) As Boolean
    CheckErrors = CheckAbilityErrors(ptypTree, build.Destiny, ptypAbility, plngSpent)
End Function

Private Sub GetSlotInfo(ptypTree As TreeType, ptypAbility As BuildAbilityType, pstrCaption As String, plngCost As Long, plngRanks, plngMaxRanks)
    With ptypAbility
        plngRanks = .Rank
        With ptypTree.Tier(.Tier).Ability(.Ability)
            If ptypAbility.Selector = 0 Then
                pstrCaption = .Abbreviation
                plngCost = .Cost
            Else
                pstrCaption = .Selector(ptypAbility.Selector).SelectorName
                If Not .SelectorOnly Then pstrCaption = .Abbreviation & ": " & pstrCaption
                plngCost = .Selector(ptypAbility.Selector).Cost
            End If
            If plngRanks <> 0 Then plngCost = plngCost * plngRanks
            plngMaxRanks = .Ranks
        End With
    End With
End Sub

Private Sub SetSlot(plngSlot As Long, ByVal pstrTier As String, ByVal pstrCost As String, pstrCaption As String, plngRanks As Long, plngMaxRanks As Long)
    With Me.usrList
        .SetText plngSlot, 1, pstrTier
        .SetText plngSlot, 3, pstrCost
        .SetSlot plngSlot, pstrCaption, plngRanks, plngMaxRanks
    End With
End Sub


' ************* GENERAL *************


Private Function GetUserChoices(plngDestiny As Long, plngTier As Long, plngAbility As Long, plngSelector As Long) As Boolean
    Dim lngItemData As Long
    
    plngTier = 0
    plngAbility = 0
    If plngSelector <> -1 Then plngSelector = 0
    If Me.lstAbility.ListIndex = -1 Then Exit Function
    lngItemData = Me.lstAbility.ItemData(Me.lstAbility.ListIndex)
    SplitAbilityID lngItemData, plngTier, plngAbility, plngDestiny
    If plngSelector <> -1 And db.Destiny(plngDestiny).Tier(plngTier).Ability(plngAbility).SelectorStyle <> sseNone Then
        If Me.lstSub.ListIndex = -1 Then Exit Function
        plngSelector = Me.lstSub.ItemData(Me.lstSub.ListIndex)
    End If
    GetUserChoices = True
End Function

Private Sub StartDrag()
    mlngSourceIndex = 0
    If mblnTwists Then
        ShowTwists dsCanDrop
    Else
        If Not AddAbility(True) Then Exit Sub
        ShowDropSlots
    End If
    If Me.lstSub.ListIndex = -1 Then ListboxDrag Me.lstAbility Else ListboxDrag Me.lstSub
End Sub

Private Sub ListboxDrag(plst As ListBox)
    plst.OLEDropMode = vbOLEDropManual
    plst.OLEDrag
End Sub

Private Function AddAbility(Optional pblnBlank As Boolean = False) As Boolean
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long
    Dim lngInsert As Long
    Dim typBlank As BuildAbilityType
    Dim i As Long
    
    If Not GetUserChoices(lngDestiny, lngTier, lngAbility, lngSelector) Then Exit Function
    lngInsert = GetInsertionPoint(build.Destiny, lngTier, lngAbility)
    If lngAbility Then lngRanks = db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility).Ranks
    With build.Destiny
        .Abilities = .Abilities + 1
        ReDim Preserve .Ability(1 To .Abilities)
        For i = .Abilities To lngInsert + 1 Step -1
            .Ability(i) = .Ability(i - 1)
        Next
        If pblnBlank Then
            .Ability(lngInsert) = typBlank
            .Ability(lngInsert).Tier = 0
        Else
            With .Ability(lngInsert)
                .Tier = lngTier
                .Ability = lngAbility
                .Selector = lngSelector
                .Rank = lngRanks
            End With
        End If
    End With
    ShowAbilities
    Me.usrList.Active = lngInsert
    Me.usrList.ForceVisible lngInsert
    AddAbility = True
End Function

Private Sub DropDestiny(Index As Integer)
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long
    
    If GetUserChoices(lngDestiny, lngTier, lngAbility, lngSelector) Then
        lngRanks = db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility).Ranks
        With build.Destiny.Ability(Index)
            .Tier = lngTier
            .Ability = lngAbility
            .Selector = lngSelector
            .Rank = lngRanks
        End With
    End If
End Sub

Private Sub DropTwist(Index As Integer)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngDestiny As Long
    
    If Index > MaxTwistSlots() Then Exit Sub
    If GetUserChoices(lngDestiny, lngTier, lngAbility, lngSelector) Then
        With build
            If Index > .Twists Then
                .Twists = .Twists + 1
                ReDim Preserve .Twist(1 To .Twists)
            End If
            With .Twist(Index)
                .DestinyName = db.Destiny(lngDestiny).TreeName
                .Tier = lngTier
                .Ability = lngAbility
                .Selector = lngSelector
            End With
        End With
    End If
End Sub


' ************* SLOTS *************


Private Sub usrList_SlotClick(Index As Integer, Button As Integer)
    Dim typBlank As TwistType
    
    With Me.lstSub
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstAbility
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.SetFocus
    If Me.usrList.Rows = 0 Then
        Exit Sub
    ElseIf Len(Me.usrList.GetCaption(1)) = 0 Then
        Exit Sub
    ElseIf Me.usrList.Selected = Index Then
        If Button = vbRightButton Then ContextMenu Index Else NoSelection
    Else
        Select Case Button
            Case vbLeftButton
                Me.usrList.Selected = Index
                If mblnTwists Then
                    If Index > build.Twists Then
                        ClearDetails True
                    Else
                        With build.Twist(Index)
                            ShowDetails SeekTree(.DestinyName, peDestiny), .Tier, .Ability, .Selector, Index
                        End With
                    End If
                Else
                    With build.Destiny.Ability(Index)
                        ShowDetails ComboGetValue(Me.cboDestiny), .Tier, .Ability, .Selector, Index
                    End With
                End If
            Case vbRightButton
                ContextMenu Index
        End Select
    End If
End Sub

Private Sub ContextMenu(Index As Integer)
    If Me.usrList.Selected <> Index Then Me.usrList.Selected = Index
    Me.usrList.Active = Index
    If mblnTwists Then
        Me.mnuContext(0).Caption = "Clear this twist"
        Me.mnuContext(1).Caption = "Clear all twists"
    Else
        Me.mnuContext(0).Caption = "Clear this ability"
        Me.mnuContext(1).Caption = "Reset destiny"
    End If
    PopupMenu Me.mnuMain(0)
End Sub

Private Sub mnuContext_Click(Index As Integer)
    Select Case Index
        Case 0 ' Clear slot
            If Me.usrList.Selected <> 0 Then ClearSlot Me.usrList.Selected
        Case 1 ' Clear all
            If mblnTwists Then
                If Not Ask("Clear all twists?") Then Exit Sub
                Me.usrList.Selected = 0
                Me.usrList.Active = 0
                Erase build.Twist
                build.Twists = 0
            Else
                If Not Ask("Reset " & Me.cboDestiny.Text & "?") Then Exit Sub
                Me.usrList.Selected = 0
                Me.usrList.Active = 0
                build.Destiny.TreeName = vbNullString
                build.Destiny.Abilities = 0
                Erase build.Destiny.Ability
                build.Destiny.TreeType = tseDestiny
            End If
    End Select
    ShowAbilities
    ShowAvailable False
    SetDirty
    If Index = 1 Then SaveBackup
End Sub

Private Sub usrList_SlotDblClick(Index As Integer)
    ClearSlot Index
End Sub

Private Sub ClearSlot(Index As Integer)
    If mblnTwists Then
        If RemoveTwist(Index) Then Exit Sub
    Else
        If build.Destiny.Abilities >= Index Then build.Destiny.Ability(Index).Ability = 0
        RemoveBlanks build.Destiny
    End If
    ShowAbilities
    ShowAvailable True
    NoSelection
    SetDirty
End Sub

Private Sub usrList_RequestDrag(Index As Integer, Allow As Boolean)
    If mblnTwists Then
        If Index > build.Twists Then Exit Sub
        ShowTwists dsCanDrop, Index
    End If
    mlngSourceIndex = Index
    Me.lstAbility.OLEDropMode = vbOLEDropManual
    Allow = True
End Sub

Private Sub usrList_OLEDragDrop(Index As Integer, Data As DataObject)
'    mblnDragComplete = True
    If mblnTwists Then
        If Data.GetData(vbCFText) = "List" Then
            DropTwist Index
        Else
            SwapTwists Index, Data.GetData(vbCFText)
        End If
    Else
        DropDestiny Index
        Me.usrList.SetDropState Index, dsDefault
    End If
    ShowAbilities
    Me.usrList.Selected = Index
    Me.usrList.SetFocus
    ShowAvailable True
    mlngSourceIndex = 0
    SetDirty
End Sub

Private Sub usrList_OLECompleteDrag(Index As Integer, Effect As Long)
    If mblnTwists Then
        ShowAbilities
        ShowAvailable True
        SetDirty
    End If
End Sub

Private Sub usrList_RankChange(Index As Integer, Ranks As Long)
    If Me.usrList.Selected <> Index Then Me.usrList.Selected = Index
    build.Destiny.Ability(Index).Rank = Ranks
    ShowAbilities
    ShowAvailable True
    With build.Destiny.Ability(Index)
        ShowDetails ComboGetValue(Me.cboDestiny), .Tier, .Ability, .Selector, Index
    End With
    SetDirty
End Sub

Private Sub SwapTwists(ByVal plngOne As Long, ByVal plngTwo As Long)
    Dim typSwap As TwistType
    
    With build
        typSwap = .Twist(plngOne)
        .Twist(plngOne) = .Twist(plngTwo)
        .Twist(plngTwo) = typSwap
    End With
End Sub

Private Function RemoveTwist(Index As Integer) As Boolean
    Dim i As Long
    
    If Index > build.Twists Then
        RemoveTwist = True
        Exit Function
    End If
    With build
        For i = Index To .Twists - 1
            .Twist(i) = .Twist(i + 1)
        Next
        .Twists = .Twists - 1
        If .Twists = 0 Then Erase .Twist Else ReDim Preserve .Twist(1 To .Twists)
    End With
End Function


' ************* ABILITIES *************


Private Sub lstAbility_Click()
    If mblnMouse Then mblnMouse = False Else ListAbilityClick
End Sub

Private Sub lstAbility_DblClick()
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    
    If Me.lstAbility.ListIndex = -1 Or Me.lstSub.ListCount > 0 Then Exit Sub
    If mblnTwists Then
        If build.Twists >= MaxTwistSlots() Then Exit Sub
        DropTwist build.Twists + 1
        ShowAbilities
        Me.usrList.Active = build.Twists
    Else
        If Not AddAbility() Then Exit Sub
    End If
    ShowAvailable True
    Me.usrList.SetFocus
    SetDirty
End Sub

Private Sub lstAbility_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long

    If Button <> vbLeftButton Or Not GetUserChoices(lngDestiny, lngTier, lngAbility, -1) Then Exit Sub
    Me.usrList.Selected = 0
    mblnMouse = ListAbilityClick()
    If mblnMouse Then menDragState = dragMouseDown
End Sub

Private Sub lstAbility_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            Me.lstAbility.OLEDropMode = vbOLEDropNone
            StartDrag
        End If
    End If
End Sub

Private Sub lstAbility_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
    ListAbilityClick
End Sub

' Show Details and Selectors, and return TRUE if we can drag this ability (ie: it has no selectors)
Private Function ListAbilityClick() As Boolean
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    
    ListboxClear Me.lstSub
    If Me.lstAbility.ListIndex = -1 Then Exit Function
    GetUserChoices lngDestiny, lngTier, lngAbility, -1
    ShowDetails lngDestiny, lngTier, lngAbility, 0, 0
    If db.Destiny(lngDestiny).Tier(lngTier).Ability(lngAbility).SelectorStyle <> sseNone Then ShowSelectors lngDestiny, lngTier, lngAbility Else ListAbilityClick = True
End Function

Private Sub lstAbility_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List"
End Sub

Private Sub lstAbility_OLECompleteDrag(Effect As Long)
    If Not mblnDragComplete Then
        If Not mblnTwists Then
            RemoveBlanks build.Destiny
            ShowDropSlots
        End If
        Me.usrList.Active = 0
        ShowAbilities
    End If
    mblnDragComplete = False
End Sub

Private Sub lstAbility_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngSourceIndex Then
        ClearSlot CInt(mlngSourceIndex)
        ShowAvailable True
        SetDirty
    End If
End Sub


' ************* SELECTORS *************


Private Sub lstSub_Click()
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    If mblnMouse Then
        mblnMouse = False
    Else
        GetUserChoices lngDestiny, lngTier, lngAbility, lngSelector
        ShowDetails lngDestiny, lngTier, lngAbility, lngSelector, 0
    End If
End Sub

Private Sub lstSub_DblClick()
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    If Me.lstSub.ListIndex = -1 Then Exit Sub
    If mblnTwists Then
        If build.Twists >= MaxTwistSlots() Then Exit Sub
        DropTwist build.Twists + 1
        ShowAbilities
        Me.usrList.Active = build.Twists
    Else
        If Not AddAbility() Then Exit Sub
    End If
    ShowAvailable True
    Me.usrList.SetFocus
    SetDirty
End Sub
    
Private Sub lstSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngDestiny As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    If Not GetUserChoices(lngDestiny, lngTier, lngAbility, lngSelector) Then Exit Sub
    Me.usrList.Selected = 0
    ShowDetails lngDestiny, lngTier, lngAbility, lngSelector, 0
    mblnMouse = True
    menDragState = dragMouseDown
End Sub

Private Sub lstSub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            Me.lstAbility.OLEDropMode = vbOLEDropNone
            StartDrag
        End If
    End If
End Sub

Private Sub lstSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub lstSub_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List"
End Sub

Private Sub lstSub_OLECompleteDrag(Effect As Long)
    If Not mblnDragComplete Then
        If Not mblnTwists Then
            RemoveBlanks build.Destiny
            ShowDropSlots
        End If
        Me.usrList.Active = 0
        ShowAbilities
    End If
    mblnDragComplete = False
End Sub


' ************* FILTERS *************


Private Sub usrchkShowAll_UserChange()
    ShowAvailable False
End Sub

Private Sub cboDestiny_Click()
    Dim strDestiny As String
    Dim strQuestion As String
    
    ClearDetails True
    If Not mblnTwists Then
        With Me.cboDestiny
            If .ListIndex <> -1 Then strDestiny = .List(.ListIndex)
        End With
        If build.Destiny.TreeName <> strDestiny Then
            If mblnOverride Then
                ComboSetText Me.cboDestiny, build.Destiny.TreeName
                Exit Sub
            ElseIf build.Destiny.Abilities Then
                If Len(strDestiny) Then strQuestion = "Change to " & strDestiny & "?" Else strQuestion = "Clear destiny?"
                If Not Ask(strQuestion) Then
                    ComboSetText Me.cboDestiny, build.Destiny.TreeName
                    Exit Sub
                End If
            End If
            build.Destiny.TreeName = strDestiny
            build.Destiny.Abilities = 0
            Erase build.Destiny.Ability
            build.Destiny.TreeType = tseDestiny
            SetDirty
        End If
    End If
    ShowAbilities
    ShowAvailable False
    SaveBackup
End Sub

Private Sub ShowAvailable(pblnPreserveTopIndex As Boolean)
    Dim lngTopIndex As Long
    Dim lngExclude As Long
    Dim lngMaxTier As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long

    If pblnPreserveTopIndex Then lngTopIndex = Me.lstAbility.TopIndex
    ListboxClear Me.lstSub
    ListboxClear Me.lstAbility
    iMin = ComboGetValue(Me.cboDestiny)
    If iMin < 1 And Not mblnTwists Then Exit Sub
    iMax = iMin
    If mblnTwists Then
        lngMaxTier = 4
        If iMin < 1 Then
            iMin = 1
            iMax = db.Destinies
            For lngExclude = 1 To db.Destinies
                If db.Destiny(lngExclude).TreeName = build.Destiny.TreeName Then Exit For
            Next
        End If
    Else
        lngMaxTier = 6
    End If
    For i = iMin To iMax
        If i <> lngExclude Then ShowAvailableDestiny i, lngMaxTier
    Next
    If pblnPreserveTopIndex Then
        With Me.lstAbility
            If lngTopIndex > .ListCount - 1 Then lngTopIndex = .ListCount - 1
            If lngTopIndex <> -1 Then .TopIndex = lngTopIndex
        End With
    End If
End Sub

Private Sub ShowAvailableDestiny(plngDestiny As Long, plngMaxTier As Long)
    Dim lngTier As Long
    Dim typCheck As BuildAbilityType
    Dim lngSpent() As Long
    Dim i As Long
    
    If Not mblnTwists Then GetSpentInTree db.Destiny(plngDestiny), build.Destiny, lngSpent, 0
    For lngTier = 1 To plngMaxTier
        typCheck.Tier = lngTier
        With db.Destiny(plngDestiny).Tier(lngTier)
            For i = 1 To .Abilities
                Do
                    If mblnTwists Then
                        If TwistTaken(db.Destiny(plngDestiny).TreeName, lngTier, i) Then Exit Do
                        If Not Me.usrchkShowAll.Value Then
                            If CheckTwistErrors(.Ability(i)) Then Exit Do
                        End If
                    Else
                        If AbilityTaken(build.Destiny, lngTier, i) Then Exit Do
                        typCheck.Ability = i
                        If Not Me.usrchkShowAll.Value Then
                            If CheckErrors(db.Destiny(plngDestiny), typCheck, lngSpent) Then Exit Do
                        End If
                    End If
                    ListboxAddItem Me.lstAbility, lngTier & ": " & .Ability(i).Abbreviation, CreateAbilityID(lngTier, i, plngDestiny)
                Loop Until True
            Next
        End With
    Next
End Sub

Private Function TwistTaken(pstrDestiny As String, plngTier As Long, plngAbility As Long) As Boolean
    Dim i As Long
    
    For i = 1 To build.Twists
        If build.Twist(i).DestinyName = pstrDestiny And build.Twist(i).Tier = plngTier And build.Twist(i).Ability = plngAbility Then
            TwistTaken = True
            Exit For
        End If
    Next
End Function

Private Sub ShowSelectors(plngDestiny As Long, plngTier As Long, plngAbility As Long)
    Dim blnSelector() As Boolean
    Dim i As Long

    ListboxClear Me.lstSub
    With db.Destiny(plngDestiny).Tier(plngTier).Ability(plngAbility)
        If mblnTwists Then
            For i = 1 To .Selectors
                ListboxAddItem Me.lstSub, .Selector(i).SelectorName, i
            Next
        Else
            GetSelectors db.Destiny(plngDestiny), plngTier, plngAbility, blnSelector
            For i = 1 To .Selectors
                If blnSelector(i) Then ListboxAddItem Me.lstSub, .Selector(i).SelectorName, i
            Next
        End If
    End With
End Sub


' ************* DETAILS *************


Private Sub ShowDetails(ByVal plngDestiny As Long, ByVal plngTier As Long, ByVal plngAbility As Long, ByVal plngSelector As Long, ByVal plngIndex As Long)
    Dim lngCost As Long
    Dim enReq As ReqGroupEnum
    Dim lngSpent() As Long
    Dim lngTotal As Long
    Dim lngProg As Long
    Dim i As Long
    
    ClearDetails False
    With db.Destiny(plngDestiny).Tier(plngTier).Ability(plngAbility)
        ' Caption
        If mblnTwists Then
            Me.lblDetails.Caption = db.Destiny(plngDestiny).Abbreviation & " Tier " & plngTier & ": " & .AbilityName
        Else
            Me.lblDetails.Caption = "Tier " & plngTier & ": " & .AbilityName
        End If
        ' Description
        If Len(.Descrip) Then Me.usrDetails.AddDescrip .Descrip, MakeWiki(db.Destiny(plngDestiny).Wiki) & TierLink(plngTier)
'        ' Class
'        If .Class(0) Then
'            Me.usrDetails.AddText "Requires class:"
'            For i = 1 To ceClasses - 1
'                If .Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
'            Next
'        End If
        ' Reqs
        For enReq = rgeAll To rgeNone
            If plngSelector = 0 Then ShowDetailsReqs .Req(enReq), enReq, 0 Else ShowDetailsReqs .Selector(plngSelector).Req(enReq), enReq, 0
        Next
        ' Rank reqs
        If plngSelector = 0 Then ShowRankReqs .RankReqs, .Rank Else ShowRankReqs .Selector(plngSelector).RankReqs, .Selector(plngSelector).Rank
        ' Error?
        If plngIndex <> 0 Then
            gstrError = vbNullString
            If mblnTwists Then
                If CheckTwistErrors(db.Destiny(plngDestiny).Tier(plngTier).Ability(plngAbility)) Then Me.usrDetails.AddErrorText "Error: " & gstrError
            Else
                GetSpentInTree db.Destiny(plngDestiny), build.Destiny, lngSpent, lngTotal
                If CheckErrors(db.Destiny(plngDestiny), build.Destiny.Ability(plngIndex), lngSpent) Then Me.usrDetails.AddErrorText "Error: " & gstrError
            End If
        End If
        ' Finished adding details
        Me.usrDetails.Refresh
        ' Ranks
        Me.lblRanks.Caption = "Ranks: " & .Ranks
        Me.lblRanks.Visible = Not mblnTwists
        ' Cost
        Me.lblCost.Caption = CostDescrip(db.Destiny(plngDestiny).Tier(plngTier).Ability(plngAbility), plngSelector)
        Me.lblCost.Visible = Not mblnTwists
        ' Spent in tree
        lngProg = GetSpentReq(tseDestiny, plngTier, plngAbility)
        If lngProg Then
            Me.lblProg.Caption = lngProg & " AP spent in tree"
            Me.lblProg.Visible = Not mblnTwists
        Else
            Me.lblProg.Visible = False
        End If
    End With
End Sub

Private Sub ShowDetailsReqs(ptypReqList As ReqListType, penGroup As ReqGroupEnum, plngRank As Long)
    Dim lngDestiny As Long
    Dim strText As String
    Dim i As Long
    
    If ptypReqList.Reqs = 0 Then Exit Sub
    lngDestiny = ComboGetValue(Me.cboDestiny)
    If plngRank < 2 Then strText = "Requires " Else strText = "Rank " & plngRank & " requires "
    strText = strText & LCase$(GetReqGroupName(penGroup)) & " of:"
    If plngRank = 0 Then
        Me.usrDetails.AddText "Requires " & LCase$(GetReqGroupName(penGroup)) & " of:"
    Else
        Me.usrDetails.AddText "Rank " & plngRank & " requires " & LCase$(GetReqGroupName(penGroup)) & " of:"
    End If
    For i = 1 To ptypReqList.Reqs
        Me.usrDetails.AddText " - " & PointerDisplay(ptypReqList.Req(i), True, lngDestiny)
    Next
End Sub

Private Sub ShowRankReqs(pblnRankReqs As Boolean, ptypRank() As RankType)
    Dim lngDestiny As Long
    Dim enReq As ReqGroupEnum
    Dim lngRank As Long
    Dim i As Long
    
    If Not pblnRankReqs Then Exit Sub
    lngDestiny = ComboGetValue(Me.cboDestiny)
    For lngRank = 2 To 3
        With ptypRank(lngRank)
            ' Class
            If .Class(0) Then
                Me.usrDetails.AddText "Rank " & lngRank & " requires Class:"
                For i = 1 To ceClasses - 1
                    If .Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
                Next
            End If
            ' Reqs
            For enReq = rgeAll To rgeNone
                ShowDetailsReqs .Req(enReq), enReq, lngRank
            Next
        End With
    Next
End Sub

Private Sub ClearDetails(pblnClearLabel As Boolean)
    If pblnClearLabel Then Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
    Me.lblRanks.Visible = False
    Me.lblCost.Visible = False
    Me.lblProg.Visible = False
End Sub

Private Sub NoSelection()
    With Me.lstSub
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstAbility
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.Selected = 0
    Me.usrList.SetFocus
    ClearDetails True
End Sub

Private Sub Form_Click()
'    NoSelection
End Sub

Private Sub usrspnDestinyAP_Change()
    Dim lngLives As Long
    
    If mblnOverride Then Exit Sub
    build.DestinyTome = Me.usrspnDestinyAP.Value
    If build.DestinyTome > tomes.DestinyMax Then lngLives = build.DestinyTome - tomes.DestinyMax
    InitGuideEnhancements
    'ShowSpentAll Me.lblSpentAll
    ShowDestinyAbilities
    'Me.lblTotal.Caption = lngTotal & " / " & (MaxDestinyAP + Me.usrspnDestinyAP.Value) & "AP"

    SetDirty
End Sub
