VERSION 5.00
Begin VB.Form frmFeats 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Feats"
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
   Icon            =   "frmFeats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12225
   Begin CharacterBuilderLite.userCheckBox usrchkChannels 
      Height          =   252
      Left            =   5520
      TabIndex        =   13
      Tag             =   "nav"
      Top             =   7440
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   450
      Value           =   0   'False
      Caption         =   "Channels"
      Bold            =   -1  'True
   End
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7380
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "< Skills"
      RightLinks      =   "Next >"
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
      LeftLinks       =   "Selected;Granted"
      RightLinks      =   "Help"
   End
   Begin CharacterBuilderLite.userDetails usrDetails 
      Height          =   2232
      Left            =   6720
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4860
      Width           =   2772
      _ExtentX        =   4895
      _ExtentY        =   3942
   End
   Begin VB.ListBox lstGroup 
      Appearance      =   0  'Flat
      Height          =   2232
      IntegralHeight  =   0   'False
      ItemData        =   "frmFeats.frx":000C
      Left            =   9720
      List            =   "frmFeats.frx":002E
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   4860
      Width           =   2172
   End
   Begin VB.ListBox lstSub 
      Appearance      =   0  'Flat
      Height          =   3390
      Left            =   9720
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   2172
   End
   Begin VB.ListBox lstFeat 
      Appearance      =   0  'Flat
      Height          =   3390
      ItemData        =   "frmFeats.frx":008F
      Left            =   6720
      List            =   "frmFeats.frx":0091
      TabIndex        =   4
      Top             =   960
      Width           =   2772
   End
   Begin CharacterBuilderLite.userCheckBox usrchkShowAll 
      Height          =   252
      Left            =   8160
      TabIndex        =   3
      Top             =   696
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   450
      Value           =   0   'False
      Caption         =   "Show All"
      CheckPosition   =   1
   End
   Begin CharacterBuilderLite.userList usrList 
      Height          =   6492
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   6132
      _ExtentX        =   10821
      _ExtentY        =   11456
   End
   Begin VB.Label lblFeats 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Feats"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   6720
      TabIndex        =   2
      Top             =   720
      Width           =   492
   End
   Begin VB.Label lblExpand 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   9168
      TabIndex        =   8
      Top             =   4620
      Width           =   324
   End
   Begin VB.Label lblRestore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   11568
      TabIndex        =   11
      Top             =   4620
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.Label lblSelectors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Selectors"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   9720
      TabIndex        =   5
      Top             =   720
      Width           =   2172
   End
   Begin VB.Label lblGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Filter"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   9720
      TabIndex        =   10
      Top             =   4620
      Width           =   1812
   End
   Begin VB.Label lblDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Details"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6720
      TabIndex        =   7
      Top             =   4620
      Width           =   2448
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuFeat 
         Caption         =   "Clear This Feat"
         Index           =   0
      End
      Begin VB.Menu mnuFeat 
         Caption         =   "Clear All Feats"
         Index           =   1
      End
      Begin VB.Menu mnuFeat 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFeat 
         Caption         =   "Add Alternate"
         Index           =   3
      End
      Begin VB.Menu mnuFeat 
         Caption         =   "Exchange Feat"
         Index           =   4
      End
      Begin VB.Menu mnuFeat 
         Caption         =   "Delete This Slot"
         Index           =   5
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmFeats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

'Private mlngTab As Long
Private mlngActive As Long
Private menChannel As FeatChannelEnum
Private menActive As FeatChannelEnum

Private mlngAnchor As Long ' Used to preserve scroll position of feats list
Private mlngOffset As Long

Private mblnMouse As Boolean
Private mlngSourceIndex As Long
Private menDragState As DragEnum
Private mblnDragComplete As Boolean
Private msngDownX As Single
Private msngDownY As Single

' Extra help needed to smooth out filter behavior when dragging
Private mlngSlotFilter As Long ' active slot filter when drag from a listbox began
Private mblnDropped As Boolean ' dragging from a listbox resulted in actually slotting a feat

Private mblnNoFocus As Boolean
Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.Configure Me
    Cascade
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Activate()
    ActivateForm oeFeats
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    UnloadForm Me, mblnOverride
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.usrList.hwnd, Xpos, Ypos): Me.usrList.Scroll lngValue
        Case IsOver(Me.usrDetails.hwnd, Xpos, Ypos): Me.usrDetails.Scroll lngValue
    End Select
End Sub

Public Sub Cascade()
    NoFocus True
    menChannel = fceUnknown
    LoadData
    ChangeTab LoadChannels()
    mlngSourceIndex = 0
    mlngActive = 0
    Me.usrchkChannels.Visible = HasChannels()
    ClearDetails
    NoFocus False
End Sub

Private Sub NoFocus(pblnNoFocus As Boolean)
    mblnNoFocus = pblnNoFocus
    Me.usrDetails.NoFocus = pblnNoFocus
    Me.usrList.NoFocus = pblnNoFocus
End Sub

Public Sub OrderChanged()
    ShowAvailable
End Sub

Private Sub FeatsChanged()
    CascadeChanges cceFeat
    SetDirty
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Feats"
        Case Else: ChangeTab GetFeatChannelID(pstrCaption)
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "< Skills"
            If Not OpenForm("frmSkills") Then Exit Sub
        Case "Spells >"
            If Not OpenForm("frmSpells") Then Exit Sub
        Case "Enhancements >"
            If Not OpenForm("frmEnhancements") Then Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub

Private Sub ChangeTab(penChannel As Long)
    Dim blnVisible As Boolean
    Dim i As Long
    
    If menChannel = penChannel Then Exit Sub
    If Not xp.DebugMode Then xp.LockWindow Me.hwnd
    ClearDetails
    blnVisible = (penChannel <> fceGranted)
    Me.lblFeats.Visible = blnVisible
    Me.usrchkShowAll.Visible = (penChannel = fceSelected Or penChannel = fceGeneral)
    Me.lblSelectors.Visible = blnVisible
    Me.lstFeat.Visible = blnVisible
    Me.lstSub.Visible = blnVisible
    If penChannel = fceGranted Then
        DetailsExpand
    ElseIf menChannel = fceGranted Then
        DetailsRestore
    End If
    menChannel = penChannel
    Me.lblRestore.Visible = False
    If menChannel = fceGranted Then
        ShowGrantedFeats
    Else
        InitSelectedFeatsList
    End If
    Me.Active = 0
    Me.usrHeader.SyncTab GetFeatChannelName(menChannel, build.Race)
    If Not xp.DebugMode Then xp.UnlockWindow
    SaveBackup
End Sub

Private Sub ShowChannelErrors()
    Dim blnChannelError() As Boolean
    Dim i As Long
    
    If Not cfg.FeatChannels Then Exit Sub
    ReDim blnChannelError(fceChannels)
    For i = 1 To Feat.Count
        With Feat.List(i)
            If .ErrorState Then blnChannelError(.Channel) = True
        End With
    Next
    For i = fceGeneral To fceChannels - 1
        Me.usrHeader.SetError GetFeatChannelName(i, build.Race), blnChannelError(i)
    Next
End Sub

Private Function HasChannels() As Boolean
    Dim i As Long

    For i = fceGeneral + 1 To fceGranted - 1
        If Feat.ChannelCount(i) > 0 Then
            HasChannels = True
            Exit Function
        End If
    Next
End Function


' ************* INITIALIZE *************


Private Sub LoadData()
    Dim i As Long
    
    mblnOverride = True
    Me.usrchkChannels.Value = cfg.FeatChannels
    'Navigation
    If build.CanCastSpell(1) <> 0 Then
        Me.usrFooter.RightLinks = "Spells >"
    Else
        Me.usrFooter.RightLinks = "Enhancements >"
    End If
    'Load the feat selection grouping list
    Me.lstGroup.Clear
    'Loop through all groups - feFilters is always last
    For i = 0 To feFilters - 1
        ListboxAddItem Me.lstGroup, GetFeatGroupName(i), i
    Next
    Me.lstGroup.ListIndex = feAll
    Me.lstGroup.Selected(feAll) = True
    mblnOverride = False
End Sub

Private Function LoadChannels() As FeatChannelEnum
    Dim strLinks As String
    Dim enChannel As FeatChannelEnum
    Dim i As Long
    
    mblnOverride = True
    Me.usrchkChannels.Value = cfg.FeatChannels
    If cfg.FeatChannels Then
        For i = fceGeneral To fceGranted - 1
            If Feat.ChannelCount(i) > 0 Then strLinks = strLinks & GetFeatChannelName(i, build.Race) & ";"
        Next
        If strLinks = "General;" Then
            strLinks = "Selected;Granted"
            LoadChannels = fceSelected
        Else
            strLinks = strLinks & "Granted"
            LoadChannels = fceGeneral
        End If
    Else
        strLinks = "Selected;Granted"
        LoadChannels = fceSelected
    End If
    Me.usrHeader.LeftLinks = strLinks
    ShowChannelErrors
    mblnOverride = False
End Function


' ************* DISPLAY *************


Private Sub ShowGrantedFeats()
    Dim lngFeat As Long
    Dim i As Long
    
    With Me.usrList
        .Clear
        .DefineDimensions Feat.ChannelCount(menChannel), 3, 2
        'TODO New Class - check max name
        .DefineColumn 1, vbLeftJustify, "Source", MAX_CLASS_NAME
        .DefineColumn 2, vbCenter, "Granted Feat"
        .DefineColumn 3, vbRightJustify, "Level", "20"
        .Refresh
    End With
    For i = 1 To Feat.Count
        With Feat.List(i)
            If .ActualType = bftGranted Then
                Me.usrList.SetText .ChannelSlot(menChannel), 1, .SourceForm
                Me.usrList.SetSlot .ChannelSlot(menChannel), .Display
                Me.usrList.SetText .ChannelSlot(menChannel), 3, .Level
                Me.usrList.SetItemData .ChannelSlot(menChannel), i
            End If
        End With
    Next
End Sub

Private Sub InitSelectedFeatsList()
    Dim i As Long
    
    With Me.usrList
        .Clear
        .DefineDimensions Feat.ChannelCount(menChannel), 4, 3
        .DefineColumn 1, vbLeftJustify, "Source", "Legendary"
        .DefineColumn 2, vbRightJustify, "Level", "Level"
        .DefineColumn 3, vbCenter, "Feat"
        .DefineColumn 4, vbCenter, "BAB", "BAB"
        .Refresh
    End With
    For i = 1 To Feat.Count
        With Feat.List(i)
            If .Channel = menChannel Or (menChannel = fceSelected And .ActualType <> bftGranted) Then
                Me.usrList.SetText .ChannelSlot(menChannel), 1, .SourceForm
                Me.usrList.SetText .ChannelSlot(menChannel), 2, .Level & "  "
                Me.usrList.SetText .ChannelSlot(menChannel), 4, build.BAB(.Level)
                Me.usrList.SetItemData .ChannelSlot(menChannel), i
                Me.usrList.SetDropState .ChannelSlot(menChannel), dsDefault
            End If
        End With
    Next
    ShowSelectedFeats
End Sub

Private Sub ShowSelectedFeats()
    Dim i As Long
    
    For i = 1 To Feat.Count
        With Feat.List(i)
            If .Channel = menChannel Or (menChannel = fceSelected And .ActualType <> bftGranted) Then
                Me.usrList.SetSlot .ChannelSlot(menChannel), .Display
                Me.usrList.SetError .ChannelSlot(menChannel), .ErrorState
                Me.usrList.SetDropState .ChannelSlot(menChannel), dsDefault
            End If
        End With
    Next
End Sub

Private Sub usrchkChannels_UserChange()
    cfg.FeatChannels = Me.usrchkChannels.Value
    ChangeTab LoadChannels()
End Sub


' ************* GENERAL *************


Public Property Get Active() As Long
    Active = mlngActive
End Property

Public Property Let Active(ByVal plngActive As Long)
    mlngActive = plngActive
    Me.usrList.Selected = mlngActive
    Me.usrList.Active = mlngActive
    If mlngActive = 0 Then
        Me.lblDetails.Caption = "Details"
        Me.usrDetails.Clear
    Else
        With Feat.List(SlotIndex(mlngActive))
            ShowDetails .FeatID, .Selector
        End With
    End If
    If menChannel <> fceGranted Then ShowAvailable
    If Me.Visible And Not mblnNoFocus Then Me.usrList.SetFocus
End Property

Private Function SlotIndex(plngSlot As Long) As Long
    If plngSlot Then SlotIndex = Me.usrList.GetItemData(plngSlot)
End Function

Private Function GetUserChoices(plngFeat As Long, plngSelector As Long) As Boolean
    plngSelector = 0
    plngFeat = ListboxGetValue(Me.lstFeat)
    If plngFeat = 0 Then Exit Function
    plngSelector = ListboxGetValue(Me.lstSub)
    If plngSelector <> 0 Or Me.lstSub.ListCount = 0 Then GetUserChoices = True
End Function

Private Sub DoubleClick()
    Dim enDropState As DropStateEnum
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim lngSlot As Long
    Dim lngSlotIndex As Long
    Dim lngIndex As Long
    Dim enType As BuildFeatTypeEnum
    Dim typCheck As BuildFeatType
    Dim blnNoScroll As Boolean
    
    If Not GetUserChoices(lngFeat, lngSelector) Then Exit Sub
    SaveScrollPosition
    If mlngActive >= 1 And mlngActive <= Me.usrList.Rows Then
        lngSlot = mlngActive
    Else
        For lngSlot = 1 To Me.usrList.Rows
            lngSlotIndex = SlotIndex(lngSlot)
            If Feat.List(lngSlotIndex).FeatID = 0 Then
                enType = Feat.List(lngSlotIndex).ActualType
                lngIndex = Feat.List(lngSlotIndex).Index
                typCheck = build.Feat(enType).Feat(lngIndex)
                typCheck.Selector = lngSelector
                Select Case CheckFeatSlot(db.Feat(lngFeat), typCheck)
                    Case dsCanDrop, dsCanDropError: Exit For
                End Select
            End If
        Next
        If lngSlot > Me.usrList.Rows Then Exit Sub
        blnNoScroll = True
    End If
    SetSlot lngSlot, lngFeat, lngSelector
    CheckSlotErrors
    ShowChannelErrors
    ShowSelectedFeats
    mlngActive = 0
    Me.usrList.Selected = 0
    Me.usrList.Active = lngSlot
    Me.usrList.ForceVisible lngSlot
    SaveScrollPosition
    ShowAvailable
    RestoreScrollPosition
    FeatsChanged
End Sub

Private Sub CompleteDrag()
    mlngSourceIndex = 0
    ShowSelectedFeats
    If mlngSlotFilter And Not mblnDropped Then
        Me.usrList.SetDropState mlngSlotFilter, dsCanDrop
    Else
        mblnDropped = False
        mlngSlotFilter = 0
    End If
End Sub

Private Sub GetSlotInfo(ByVal plngSlot As Long, penType As BuildFeatTypeEnum, plngIndex As Long)
    With Feat.List(SlotIndex(plngSlot))
        penType = .ActualType
        plngIndex = .Index
    End With
End Sub

Private Sub ShowDropSlots()
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim lngSlot As Long
    Dim lngIndex As Long
    Dim enType As BuildFeatTypeEnum
    Dim typDrag As BuildFeatType
    
    If mlngSourceIndex = 0 Then
        If Not GetUserChoices(lngFeat, lngSelector) Then Exit Sub
    Else
        lngFeat = Feat.List(SlotIndex(mlngSourceIndex)).FeatID
        If lngFeat = 0 Then Exit Sub
        lngSelector = Feat.List(SlotIndex(mlngSourceIndex)).Selector
    End If
    For lngSlot = 1 To Me.usrList.Rows
        With Feat.List(SlotIndex(lngSlot))
            enType = .ActualType
            lngIndex = .Index
            typDrag = build.Feat(enType).Feat(lngIndex)
            typDrag.Selector = lngSelector
            Me.usrList.SetDropState lngSlot, CheckFeatSlot(db.Feat(lngFeat), typDrag)
        End With
    Next
End Sub
'
'Private Function SlotError(plngSlot As Long) As Boolean
'    Dim enType As BuildFeatTypeEnum
'    Dim lngIndex As Long
'    Dim lngFeat As Long
'    Dim lngLevel As Long
'    Dim lngSelector As Long
'    Dim blnValid As Boolean
'
'    With Feat.List(SlotIndex(plngSlot))
'        lngFeat = .FeatID
'        lngIndex = .Index
'        enType = .ActualType
'    End With
'    With build.Feat(enType).Feat(lngIndex)
'        lngLevel = .Level
'        lngSelector = .Selector
'    End With
'    If CheckFeatSlot(db.Feat(lngFeat), build.Feat(enType).Feat(lngIndex)) <> dsCanDrop Then SlotError = True
'End Function


' ************* SLOTS *************


Private Sub usrList_SlotClick(Index As Integer, Button As Integer)
    Select Case Button
        Case vbLeftButton
            If Me.Active = Index Then
                Me.Active = 0
            Else
                Me.Active = Index
                ShowRelated
            End If
        Case vbRightButton
            If menChannel <> fceGranted Then ContextMenu Index
    End Select
End Sub

Private Sub usrList_SlotDblClick(Index As Integer)
    If menChannel <> fceGranted Then ClearSlot Index
End Sub

Public Sub ShowRelated()
    Dim lngSlotIndex As Long
    Dim enHomeType As BuildFeatTypeEnum
    Dim lngHomeIndex As Long
    Dim enAwayType As BuildFeatTypeEnum
    Dim lngAwayIndex As Long
    Dim i As Long
    
    lngSlotIndex = SlotIndex(mlngActive)
    If lngSlotIndex = 0 Then Exit Sub
    With Feat.List(lngSlotIndex)
        enHomeType = .ActualType
        lngHomeIndex = .Index
        enAwayType = .EffectiveType
        lngAwayIndex = .ParentIndex
    End With
    For i = 1 To Feat.Count
        With Feat.List(i)
            If .ActualType <> bftGranted Then
                If i <> lngSlotIndex Then
                    If .EffectiveType = enHomeType And .ParentIndex = lngHomeIndex Then
                        Me.usrList.SetDropState .ChannelSlot(menChannel), dsSubordinate
                    ElseIf .ActualType = enAwayType And .Index = lngAwayIndex Then
                        Me.usrList.SetDropState .ChannelSlot(menChannel), dsSubordinate
                    End If
                End If
            End If
        End With
    Next
End Sub

Private Sub ContextMenu(ByVal plngSlot As Long)
    Dim lngSlotIndex As Long
    Dim blnSpecial As Boolean
    
    Me.Active = plngSlot
    ShowRelated
    lngSlotIndex = SlotIndex(plngSlot)
    If lngSlotIndex = 0 Then Exit Sub
    Select Case Feat.List(lngSlotIndex).ActualType
        Case bftAlternate, bftExchange
            Me.mnuFeat(3).Visible = False
            Me.mnuFeat(4).Visible = False
            Me.mnuFeat(5).Visible = True
        Case Else
            Me.mnuFeat(3).Visible = True
            Me.mnuFeat(4).Visible = (Feat.List(lngSlotIndex).Level < build.MaxLevels And Feat.List(lngSlotIndex).ExchangeIndex = 0)
            Me.mnuFeat(5).Visible = False
    End Select
    PopupMenu Me.mnuMain(0)
End Sub

Private Sub mnuFeat_Click(Index As Integer)
    If Me.Active = 0 Then Exit Sub
    Select Case Me.mnuFeat(Index).Caption
        Case "Clear This Feat": ClearSlot Me.Active
        Case "Clear All Feats": ClearAllSlots
        Case "Add Alternate": AddAlternate
        Case "Exchange Feat": ExchangeFeat
        Case "Delete This Slot": DeleteSlot
    End Select
End Sub

Private Sub usrList_RequestDrag(Index As Integer, Allow As Boolean)
    If menChannel <> fceGranted And mlngSourceIndex = 0 Then
        Me.Active = Index
        mlngSourceIndex = Index
        Allow = True
        Me.lstFeat.OLEDropMode = vbOLEDropManual
        ShowDropSlots
    End If
End Sub

Private Sub usrList_OLEDragDrop(Index As Integer, Data As DataObject)
    If menChannel = fceGranted Then Exit Sub
    If mlngSourceIndex = 0 Then
        SlotDrop Index
        mblnDropped = True
    ElseIf mlngSourceIndex <> Index Then
        SwapSlots mlngSourceIndex, Index
    End If
End Sub

Private Sub usrList_OLECompleteDrag(Index As Integer, Effect As Long)
    CompleteDrag
End Sub

Private Sub SetSlot(plngSlot As Long, plngFeat As Long, plngSelector As Long)
    Dim lngSlotIndex As Long
    
    If plngSlot = 0 Then Exit Sub
    lngSlotIndex = Me.usrList.GetItemData(plngSlot)
    If lngSlotIndex = 0 Then Exit Sub
    With Feat.List(lngSlotIndex)
        .FeatID = plngFeat
        .Selector = plngSelector
        If plngFeat = 0 Then .FeatName = vbNullString Else .FeatName = db.Feat(plngFeat).FeatName
        With build.Feat(.ActualType).Feat(.Index)
            .FeatName = db.Feat(plngFeat).FeatName
            .Selector = plngSelector
        End With
    End With
    GetDisplayNames Feat.List(lngSlotIndex)
End Sub

Private Sub ClearSlot(ByVal plngSlot As Long)
    Dim lngSlotIndex As Long
    Dim lngItemData As Long
    
    If plngSlot = 0 Then Exit Sub
    SetSlot plngSlot, 0, 0
    CheckSlotErrors
    ShowChannelErrors
    ShowSelectedFeats
    Me.Active = 0
    ShowAvailable
    FeatsChanged
End Sub

Private Sub ClearAllSlots()
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim blnEmpty As Boolean
    Dim blnErrors As Boolean
    Dim blnComplete As Boolean
    
    If Not Ask("Clear all feats?") Then Exit Sub
    xp.LockWindow Me.hwnd
    Erase build.Feat(bftAlternate).Feat
    build.Feat(bftAlternate).Feats = 0
    Erase build.Feat(bftExchange).Feat
    build.Feat(bftExchange).Feats = 0
    For enType = bftStandard To bftDeity
        With build.Feat(enType)
            For lngIndex = 1 To .Feats
                With .Feat(lngIndex)
                    .FeatName = vbNullString
                    .Selector = 0
                End With
            Next
        End With
    Next
    InitBuildFeats
    InitSelectedFeatsList
    Me.Active = 0
    ShowAvailable
    xp.UnlockWindow
    FeatsChanged
End Sub

Private Sub AddAlternate()
    Dim lngParent As Long
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    
    If mlngActive = 0 Then Exit Sub
    lngParent = SlotIndex(mlngActive)
    enType = Feat.List(lngParent).ActualType
    lngIndex = Feat.List(lngParent).Index
    FeatListChange Feat.List(AddSpecialFeat(bftAlternate, enType, lngIndex)).ChannelSlot(menChannel), False
End Sub

Private Sub ExchangeFeat()
    Dim lngParent As Long
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim lngNewIndex As Long
    
    lngParent = SlotIndex(Me.Active)
    glngLevel = Feat.List(lngParent).Level
    frmLevel.Show vbModal, Me
    If glngLevel = 0 Then Exit Sub
    lngParent = SlotIndex(mlngActive)
    enType = Feat.List(lngParent).ActualType
    lngIndex = Feat.List(lngParent).Index
    lngNewIndex = AddSpecialFeat(bftExchange, enType, lngIndex)
    FeatListChange Feat.List(lngNewIndex).ChannelSlot(menChannel), True
    If SlotLocked(lngNewIndex) Then
        ClearDetails
        Me.lblDetails.Caption = "Fred says:"
        Me.usrDetails.AddErrorText gstrError
        Me.usrDetails.Refresh
    End If
    gstrError = vbNullString
End Sub

Private Sub DeleteSlot()
    Dim lngSlotIndex As Long
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim i As Long
    
    lngSlotIndex = SlotIndex(Me.Active)
    If lngSlotIndex = 0 Then Exit Sub
    With Feat.List(lngSlotIndex)
        enType = .ActualType
        lngIndex = .Index
    End With
    DeleteFeatSlot enType, lngIndex
    FeatListChange 0, False
    ShowChannelErrors
End Sub

Private Sub FeatListChange(plngForceVisibleIndex As Long, pblnShowRelated As Boolean)
    Dim lngScroll As Long
    
    xp.LockWindow Me.hwnd
    lngScroll = Me.usrList.ScrollPosition
    InitFeatList
    InitSelectedFeatsList
    Me.usrList.ScrollPosition = lngScroll
    If plngForceVisibleIndex Then
        Me.usrList.ForceVisible plngForceVisibleIndex
        Me.Active = plngForceVisibleIndex
        If pblnShowRelated Then ShowRelated
'        Me.usrList.Active = plngForceVisibleIndex
    Else
        Me.Active = 0
    End If
    xp.UnlockWindow
    FeatsChanged
End Sub

Private Sub SlotDrop(ByVal plngSlot As Long)
    Dim lngFeat As Long
    Dim lngSelector As Long
    
    SaveScrollPosition
    Select Case Me.usrList.GetDropState(plngSlot)
        Case dsCanDrop, dsCanDropError
        Case Else: Exit Sub
    End Select
    If Not GetUserChoices(lngFeat, lngSelector) Then Exit Sub
    SetSlot plngSlot, lngFeat, lngSelector
    CheckSlotErrors
    ShowChannelErrors
    mlngActive = 0
    ShowAvailable
    If Me.lstSub.ListCount > 1 Then
        ListboxSetValue Me.lstFeat, lngFeat
    Else
        If Not mblnNoFocus Then Me.lstFeat.SetFocus
    End If
    RestoreScrollPosition
    FeatsChanged
End Sub

Private Sub SwapSlots(ByVal plngFrom As Long, ByVal plngTo As Long)
    Dim lngFrom As Long
    Dim lngTo As Long
    Dim lngFeat As Long
    Dim lngSelector As Long
    
    lngFrom = Me.usrList.GetItemData(plngFrom)
    lngFeat = Feat.List(lngFrom).FeatID
    lngSelector = Feat.List(lngFrom).Selector
    lngTo = Me.usrList.GetItemData(plngTo)
    SetSlot plngFrom, Feat.List(lngTo).FeatID, Feat.List(lngTo).Selector
    SetSlot plngTo, lngFeat, lngSelector
    CheckSlotErrors
    ShowChannelErrors
    mlngActive = 0
    ShowAvailable
    FeatsChanged
End Sub


' ************* FEAT LIST *************


Private Sub lstFeat_Click()
    If mblnMouse Then mblnMouse = False Else ListFeatClick
End Sub

Private Sub lstFeat_DblClick()
    DoubleClick
End Sub

Private Sub lstFeat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngFeat As Long
    
    lngFeat = ListboxGetValue(Me.lstFeat)
    If Button <> vbLeftButton Or lngFeat = 0 Then Exit Sub
    Me.usrList.Selected = Me.Active
    mblnMouse = ListFeatClick()
    If mblnMouse Then menDragState = dragMouseDown
End Sub

Private Sub lstFeat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            Me.lstFeat.OLEDropMode = vbOLEDropNone
            If ListFeatClick() Then
                mlngSlotFilter = Me.usrList.Selected
                ShowDropSlots
                Me.lstFeat.OLEDropMode = vbOLEDropManual
                Me.lstFeat.OLEDrag
            End If
        End If
    End If
End Sub

Private Sub lstFeat_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
    ListFeatClick
End Sub

Private Function ListFeatClick() As Boolean
    Dim lngFeat As Long
    
    ListboxClear Me.lstSub
    If Me.lstFeat.ListIndex = -1 Then Exit Function
    lngFeat = Me.lstFeat.ItemData(Me.lstFeat.ListIndex)
    ShowDetails lngFeat, 0
    If db.Feat(lngFeat).SelectorStyle <> sseNone Then ShowSelectors Else ListFeatClick = True
End Function

Private Sub lstFeat_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData Me.lstFeat.Text
End Sub

Private Sub lstFeat_OLECompleteDrag(Effect As Long)
    CompleteDrag
End Sub

Private Sub SaveScrollPosition()
    With Me.lstFeat
        If .ListIndex < 1 Then
            mlngAnchor = -1
            mlngOffset = -1
        Else
            mlngAnchor = .ItemData(.ListIndex - 1)
            mlngOffset = .ListIndex - .TopIndex
        End If
    End With
End Sub

Private Sub RestoreScrollPosition()
    Dim lngTopIndex As Long
    Dim i As Long
    
    If Me.lstFeat.ListCount = 0 Then Exit Sub
    With Me.lstFeat
        If mlngAnchor = -1 Then
            .TopIndex = 0
        Else
            For i = 0 To .ListCount - 1
                If .ItemData(i) = mlngAnchor Then
                    lngTopIndex = i - mlngOffset + 1
                    If lngTopIndex < 0 Then lngTopIndex = 0
                    .TopIndex = lngTopIndex
                End If
            Next
        End If
    End With
    mlngAnchor = -1
    mlngOffset = -1
End Sub

Private Sub lstFeat_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngSourceIndex Then
        ClearSlot CInt(mlngSourceIndex)
        ShowAvailable
        FeatsChanged
    End If
End Sub

Private Sub ShowSelectors()
    Dim lngFeat As Long
    Dim blnSelector() As Boolean
    Dim i As Long
    
    ListboxClear Me.lstSub
    lngFeat = ListboxGetValue(Me.lstFeat)
    ValidSelectors db.Feat(lngFeat), build.MaxLevels, blnSelector
    For i = 1 To db.Feat(lngFeat).Selectors
        If blnSelector(i) Then
            'Hide this selector - currently only needed for composite shortbows (no idea why)
            'and Divination, which has no spell focus feat.
            If Not db.Feat(lngFeat).Selector(i).Hide Then
                ListboxAddItem Me.lstSub, db.Feat(lngFeat).Selector(i).SelectorName, i
            End If
        End If
    Next
End Sub

' Given a feat, populate list of all selectors that are still a valid choice
Private Sub ValidSelectors(ptypFeat As FeatType, ByVal plngLevel As Long, pblnSelector() As Boolean)
    Dim typTaken() As FeatTakenType
    Dim enReqGroup As ReqGroupEnum
    Dim blnAlternate As Boolean
    Dim i As Long
    
    If Me.usrList.Selected > 0 Then blnAlternate = (Feat.List(SlotIndex(Me.usrList.Selected)).ActualType = bftAlternate)
    With ptypFeat
        ' Initialize list based on which selectors have already been chosen
        IdentifyTakenFeats typTaken, plngLevel
        ReDim pblnSelector(1 To .Selectors)
        For i = 1 To .Selectors
            pblnSelector(i) = Not .Selector(i).All
        Next
        If .Times <> 99 Then
            For i = 1 To .Selectors
                If typTaken(.FeatIndex).Selector(i) Then pblnSelector(i) = False
            Next
            If .SelectorStyle = sseShared And Not blnAlternate Then
                For i = 1 To .Selectors
                    If Not typTaken(.Parent.Feat).Selector(i) Then pblnSelector(i) = False
                Next
            End If
        End If
        ' Now remove selectors that fail selector reqs
        If Not (Me.usrchkShowAll.Value = vbChecked Or .Times = 99 Or blnAlternate) Then
            For i = 1 To ptypFeat.Selectors
                If pblnSelector(i) Then
                    With .Selector(i)
                        Do
                            pblnSelector(i) = False
                            If RaceRestricted(.Race) Then Exit Do
                            If .Class(0) Then
                                If Not CheckClassLevels(plngLevel, .Class, .ClassLevel) Then Exit Do
                            End If
                            If .Stat <> aeAny Then
                                If .StatValue > CalculateStat(.Stat, plngLevel) Then Exit Do
                            End If
                            If .Skill <> seAny Then
                                If .SkillValue > CalculateSkill(.Skill, plngLevel, ptypFeat.SkillTome) Then Exit Do
                            End If
                            For enReqGroup = rgeAll To rgeNone
                                If CheckFeatReq(.Req(enReqGroup), enReqGroup, plngLevel, typTaken) Then Exit Do
                            Next
                            pblnSelector(i) = True
                        Loop Until True
                    End With
                End If
            Next
        End If
    End With
End Sub


' ************* SELECTORS *************


Private Sub lstSub_Click()
    Dim lngFeat As Long
    Dim lngSelector As Long
    
    If Not GetUserChoices(lngFeat, lngSelector) Then Exit Sub
    ShowDetails lngFeat, lngSelector
End Sub

Private Sub lstSub_DblClick()
    DoubleClick
End Sub

Private Sub lstSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTopIndex As Long
    Dim lngFeat As Long
    Dim lngSelector As Long
    
    If Not GetUserChoices(lngFeat, lngSelector) Then Exit Sub
    ShowDetails lngFeat, lngSelector
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
            Me.lstFeat.OLEDropMode = vbOLEDropNone
            mlngSlotFilter = Me.usrList.Selected
            ShowDropSlots
            Me.lstSub.OLEDropMode = 1
            Me.lstSub.OLEDrag
        End If
    End If
End Sub

Private Sub lstSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub lstSub_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData Me.lstSub
End Sub

Private Sub lstSub_OLECompleteDrag(Effect As Long)
    CompleteDrag
End Sub


' ************* FILTERS *************


Private Sub ShowAvailable()
    Dim typTaken() As FeatTakenType
    Dim lngFeat As Long
    Dim enType As BuildFeatTypeEnum
    Dim lngIndex As Long
    Dim strPrefix As String
    Dim lngSlotIndex As Long
    Dim i As Long
    
    ' Clear the feat list
    ListboxClear Me.lstSub
    ListboxClear Me.lstFeat
    FeatListTitle
    ' Create list of taken feats now to speed up loop
    IdentifyTakenFeats typTaken, build.MaxLevels
    ' Add all feats that meet prereqs
    For i = 1 To db.Feats
        lngFeat = db.FeatDisplay(i).FeatIndex
        Do
            ' Wrong channel?
            If menChannel <> fceSelected And db.Feat(lngFeat).Channel <> menChannel Then Exit Do
            ' User filter?
            If Not FilterFeatGroup(lngFeat) Then Exit Do
            ' Already taken? (Selectors normally allowed to be taken once for each selector)
            If db.Feat(lngFeat).SelectorStyle = sseNone Then
                If typTaken(lngFeat).Times <> 0 Then
                    If db.Feat(lngFeat).Times = 0 Or typTaken(lngFeat).Times >= db.Feat(lngFeat).Times Then Exit Do
                End If
            Else
                If typTaken(lngFeat).Times <> 0 Then
                    If db.Feat(lngFeat).Times <> 0 And typTaken(lngFeat).Times >= db.Feat(lngFeat).Times Then Exit Do
                End If
                ' Still have selectors available?
                If NoValidSelectors(typTaken, lngFeat) Then Exit Do
            End If
            If Not Me.usrchkShowAll.Value Then
                ' Apply slot filter
                If Me.Active <> 0 Then
                    lngSlotIndex = SlotIndex(Me.Active)
                    If lngSlotIndex = 0 Then
                        Exit Do
                    Else
                        enType = Feat.List(lngSlotIndex).ActualType
                        lngIndex = Feat.List(lngSlotIndex).Index
                        If CheckFeatSlot(db.Feat(lngFeat), build.Feat(enType).Feat(lngIndex), False) = dsDefault Then Exit Do
                    End If
                Else
                    If Not FilterFeat(lngFeat) Then Exit Do
                End If
            End If
            ' This feat meets all qualifications, so add it to list
            If db.Feat(lngFeat).SortName <> db.Feat(lngFeat).FeatName And IsNumeric(Right$(db.Feat(lngFeat).SortName, 1)) And cfg.FeatOrder = foeGroupRelated Then
                strPrefix = "  "
            Else
                strPrefix = vbNullString
            End If
            ListboxAddItem Me.lstFeat, strPrefix & db.Feat(lngFeat).Abbreviation, lngFeat
        Loop Until True
    Next
End Sub

Private Sub FeatListTitle()
    Dim strTitle As String
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    If menChannel = fceGranted Then Exit Sub
    If Me.Active = 0 Then
        strTitle = "Feats"
    Else
        strTitle = Feat.List(Me.usrList.GetItemData(Me.Active)).SourceFilter
    End If
    With Me.lblFeats
        .Caption = strTitle
        lngLeft = .Left + .Width
        lngWidth = Me.lstFeat.Width - .Width
    End With
    With Me.usrchkShowAll
'        .Visible = False
        .Move lngLeft, .Top, lngWidth
'        .Visible = True
    End With
End Sub

Private Function NoValidSelectors(ptypTaken() As FeatTakenType, plngFeat As Long) As Boolean
    Dim blnSelector() As Boolean
    Dim lngLevel As Long
    Dim lngCount As Long
    Dim i As Long
    
    If Me.Active = 0 Then
        lngLevel = build.MaxLevels
    Else
        lngLevel = Feat.List(Me.usrList.GetItemData(Me.Active)).Level
    End If
    ValidSelectors db.Feat(plngFeat), lngLevel, blnSelector
    For i = 1 To db.Feat(plngFeat).Selectors
        If blnSelector(i) Then lngCount = lngCount + 1
    Next
    NoValidSelectors = (lngCount = 0)
End Function

Private Sub usrchkShowAll_UserChange()
    ShowAvailable
End Sub

Private Sub lstGroup_Click()
    If mblnOverride Then Exit Sub
    ShowAvailable
End Sub

' Can this feat ever be taken? (race restricted; class levels never taken; spellcasting never achieved; stat, skill or level never reached)
Private Function FilterFeat(plngFeat As Long) As Boolean
    Dim lngClassLevels() As Long
    Dim blnClass As Boolean
    Dim i As Long
    
    With db.Feat(plngFeat)
        Do
            ' Race
            If RaceRestricted(.Race) Then Exit Do
            ' Class levels
            If .Class(0) Then
                ReDim lngClassLevels(ceClasses - 1)
                For i = 1 To HeroicLevels()
                    lngClassLevels(build.Class(i)) = lngClassLevels(build.Class(i)) + 1
                Next
                For i = 1 To ceClasses - 1
                    If .Class(i) Then
                        If lngClassLevels(i) >= .ClassLevel(i) Then
                            blnClass = True
                            Exit For
                        End If
                    End If
                Next
                If Not blnClass Then Exit Do
            End If
            ' CanCastSpell
            If .CanCastSpell Then
                If build.CanCastSpell(.CanCastSpellLevel) = 0 Then Exit Do
            End If
            ' Stat
            If .Stat <> aeAny Then
                If .StatValue > CalculateStat(.Stat, build.MaxLevels) Then Exit Do
            End If
            ' Skill
            If .Skill <> aeAny Then
                If .SkillValue > CalculateSkill(.Skill, build.MaxLevels, .SkillTome) Then Exit Do
            End If
            ' Past lives
            If .PastLife And build.BuildPoints < beHero Then Exit Do
            If .Legend And build.BuildPoints < beLegend Then Exit Do
            ' Max Level not high enough
            If .Level > build.MaxLevels Then Exit Do
            FilterFeat = True
        Loop Until True
    End With
End Function

' Support multiple selections - "AND" version
Private Function FilterFeatGroup(plngFeat As Long) As Boolean
    Dim enGroup As FilterEnum
    Dim i As Long
    
    FilterFeatGroup = True
    ' Group
    With Me.lstGroup
        For i = 0 To .ListCount - 1
            If .Selected(i) And Not db.Feat(plngFeat).Group(i) Then
                FilterFeatGroup = False
                Exit For
            End If
        Next
    End With
End Function
'
'' Support multiple selections - "OR" version
'Private Function FilterFeatGroup(plngFeat As Long) As Boolean
'    Dim enGroup As FilterEnum
'    Dim i As Long
'
'    ' Group
'    With Me.lstGroup
'        For i = 0 To .ListCount - 1
'            If .Selected(i) And db.Feat(plngFeat).Group(i) Then
'                FilterFeatGroup = True
'                Exit For
'            End If
'        Next
'    End With
'End Function


' ************* DETAILS *************


Private Sub ClearDetails()
    Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
End Sub

Private Sub ShowDetails(plngFeat As Long, plngSelector As Long)
    Dim lngFeat As Long
    Dim lngSelector As Long
    Dim strText As String
    Dim strDescrip As String
    Dim strWiki As String
    Dim lngErrorIndex As Long
    Dim i As Long
    
    Me.usrDetails.Clear
    If plngFeat = 0 Then
        If Me.Active And Me.lstFeat.ListIndex = -1 Then
            lngErrorIndex = SlotIndex(Me.Active)
            If Feat.List(lngErrorIndex).ErrorState Then
                Me.lblDetails.Caption = "Fred says:"
                Me.usrDetails.AddErrorText Feat.List(lngErrorIndex).ErrorText
                Me.usrDetails.Refresh
                Exit Sub
            End If
        End If
        Me.lblDetails.Caption = "Details"
        Exit Sub
    Else
        Me.lblDetails.Caption = GetFeatDisplay(plngFeat, plngSelector, True, False)
    End If
    With db.Feat(plngFeat)
        If plngSelector <> 0 Then
            strDescrip = .Selector(plngSelector).Descrip
            strWiki = .Selector(plngSelector).Wiki
        End If
        If Len(strDescrip) = 0 Then strDescrip = .Descrip
        If Len(strWiki) = 0 Then strWiki = .Wiki
        If Len(strDescrip) Then Me.usrDetails.AddDescrip strDescrip, MakeWiki(strWiki)
        If HasRequirements(plngFeat, plngSelector) Then Me.usrDetails.AddText "Requirements:"
        ' Level
        If .Level Then Me.usrDetails.AddText " - " & .Level & " character levels"
        ' CanCastSpell
        If .CanCastSpell Then
            If .CanCastSpellLevel = 0 Then
                Me.usrDetails.AddText " - can cast healing spells"
            Else
                Me.usrDetails.AddText " - can cast level " & .CanCastSpellLevel & " spells"
            End If
        End If
        ' BAB
        If .BAB <> 0 Then Me.usrDetails.AddText " - BAB: " & .BAB
        ' Stat
        If plngSelector = 0 Then
            If .StatValue Then Me.usrDetails.AddText " - " & GetStatName(.Stat) & " " & .StatValue
        Else
            If .Selector(plngSelector).StatValue Then Me.usrDetails.AddText " - " & GetStatName(.Selector(plngSelector).Stat) & " " & .Selector(plngSelector).StatValue
        End If
        ' Skill
        If plngSelector = 0 Then
            If .SkillValue Then Me.usrDetails.AddText " - " & GetSkillName(.Skill) & " " & .SkillValue
        Else
            If .Selector(plngSelector).SkillValue Then Me.usrDetails.AddText " - " & GetSkillName(.Selector(plngSelector).Skill) & " " & .Selector(plngSelector).SkillValue
        End If
        ' Past lives
        If .Legend Then Me.usrDetails.AddText " - Legend build"
        If .PastLife Then Me.usrDetails.AddText " - Hero or Legend build"
        ' Race
        If plngSelector = 0 Then ShowDetailsRace .Race Else ShowDetailsRace .Selector(plngSelector).Race
        ' RaceOnly
        If .RaceOnly Then Me.usrDetails.AddText "Only as racial bonus feat"
        ' ClassOnly
        If .ClassOnly Then
            Me.usrDetails.AddText "Only as bonus feat for:"
            For i = 1 To ceClasses - 1
                If .ClassOnlyClasses(i) Then Me.usrDetails.AddText " - " & GetClassName(i)
            Next
            Me.usrDetails.AddText "Only on class level(s):"
            strText = " - "
            For i = 1 To 20
                If .ClassOnlyLevels(i) Then strText = strText & i & ", "
            Next
            Me.usrDetails.AddText Left$(strText, Len(strText) - 2)
        ' Class
        ElseIf .Class(0) Then
            Me.usrDetails.AddText "Class required:"
            For i = 1 To ceClasses - 1
                If .Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
            Next
        End If
        If plngSelector <> 0 Then
            If .Selector(plngSelector).Class(0) Then
                Me.usrDetails.AddText "Class required:"
                For i = 1 To ceClasses - 1
                    If .Selector(plngSelector).Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .Selector(plngSelector).ClassLevel(i)
                Next
            End If
        End If
        ' Alignment
        If plngSelector <> 0 Then
            If .Selector(plngSelector).Alignment(0) Then
                Me.usrDetails.AddText "Alignment Required: "
                For i = 1 To 6
                    If .Selector(plngSelector).Alignment(i) Then Me.usrDetails.AddText " - " & GetAlignmentName(i)
                Next
            End If
        End If
'        ' Class Bonus Level
'        If .ClassBonusLevel(0) = -1 Then
'            Me.usrDetails.AddText "Bonus Feat Class Levels:"
'            For i = 1 To ceClasses - 1
'                If .ClassBonusLevel(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassBonusLevel(i)
'            Next
'        End If
        ' FeatReqs
        If plngSelector = 0 Then ShowDetailsReqs .Req Else ShowDetailsReqs .Selector(plngSelector).Req
        ' Shared selector
        If .SelectorStyle = sseShared Then
            Me.usrDetails.AddText "Same selector as:"
            Me.usrDetails.AddText " - " & PointerDisplay(.Parent, True)
        End If
    End With
    If Me.Active And Me.lstFeat.ListIndex = -1 Then
        With Feat.List(SlotIndex(Me.Active))
            If .ErrorState Then Me.usrDetails.AddErrorText "Error: " & .ErrorText
        End With
    End If
    Me.usrDetails.Refresh
End Sub

Private Function HasRequirements(plngFeat As Long, plngSelector As Long) As Boolean
    Dim blnNoReqs As Boolean
    
    With db.Feat(plngFeat)
        Do
            If .Level Then Exit Do
            If .CanCastSpell Then Exit Do
            If .BAB <> 0 Then Exit Do
            If plngSelector = 0 Then
                If .StatValue Then Exit Do
                If .SkillValue Then Exit Do
            Else
                If .Selector(plngSelector).StatValue Then Exit Do
                If .Selector(plngSelector).SkillValue Then Exit Do
            End If
            If .PastLife Or .Legend Then Exit Do
            blnNoReqs = True
        Loop Until True
    End With
    HasRequirements = Not blnNoReqs
End Function

Private Sub ShowDetailsRace(plngRace() As Long)
    Dim enRaceReq As RaceReqEnum
    Dim lngCount As Long
    Dim strText As String
    Dim strPrefix As String
    Dim strSuffix As String
    Dim i As Long
    
    enRaceReq = plngRace(0)
    If enRaceReq = rreAny Then Exit Sub
    For i = 1 To reRaces - 1
        lngCount = lngCount + plngRace(i)
    Next
    Select Case enRaceReq
        Case rreRequired
            strText = "Race required:"
        Case rreNotAllowed
            strText = "Race restricted:"
            strPrefix = "Not "
        Case rreStandard
            Me.usrDetails.AddText "Race restricted:"
            strText = " - Not Iconic"
            strPrefix = "Not "
        Case rreIconic
            Me.usrDetails.AddText "Race restricted:"
            strText = " - Iconic"
            strPrefix = "Not "
    End Select
    Me.usrDetails.AddText strText
    For i = 1 To reRaces - 1
        If plngRace(i) = 1 Then Me.usrDetails.AddText " - " & strPrefix & GetRaceName(i) & strSuffix
    Next
End Sub

Private Sub ShowDetailsReqs(ptypReqList() As ReqListType)
    Dim lngReqGroup As Long
    Dim i As Long
    
    For lngReqGroup = 1 To 3
        With ptypReqList(lngReqGroup)
            If .Reqs Then
                Me.usrDetails.AddText "Requires " & LCase$(GetReqGroupName(lngReqGroup)) & " of:"
                For i = 1 To .Reqs
                    Me.usrDetails.AddText " - " & PointerDisplay(.Req(i), True)
                Next
            End If
        End With
    Next
End Sub

Public Sub DetailsExpand()
    xp.SetMouseCursor mcHand
    Me.lblExpand.Visible = False
    Me.lblGroup.Visible = False
    Me.lstGroup.Visible = False
    Me.lblRestore.Visible = True
    Me.lblDetails.Width = Me.lblRestore.Left - Me.lblDetails.Left
    Me.usrDetails.Width = Me.lstGroup.Left + Me.lstGroup.Width - Me.usrDetails.Left
End Sub

Private Sub DetailsRestore()
    xp.SetMouseCursor mcHand
    Me.lblDetails.Width = Me.lblExpand.Left - Me.lblDetails.Left
    Me.lblExpand.Visible = True
    Me.usrDetails.Width = Me.lstFeat.Width
    Me.lblRestore.Visible = False
    Me.lblGroup.Visible = True
    Me.lstGroup.Visible = True
End Sub

Private Sub lblDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If menChannel <> fceGranted Then xp.SetMouseCursor mcHand
End Sub

Private Sub lblDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If menChannel = fceGranted Then Exit Sub
    If Me.lblExpand.Visible Then DetailsExpand Else DetailsRestore
End Sub

Private Sub lblExpand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lblExpand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DetailsExpand
End Sub

Private Sub lblRestore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lblRestore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DetailsRestore
End Sub

Private Sub NoSelection()
    With Me.lstSub
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstFeat
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.Active = 0
    FeatListTitle
    Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
    On Error Resume Next
    If Not mblnNoFocus Then Me.usrList.SetFocus
End Sub

Private Sub Form_Click()
'    NoSelection
End Sub
