VERSION 5.00
Begin VB.Form frmSpells 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spells"
   ClientHeight    =   7764
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpells.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7764
   ScaleWidth      =   12216
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7380
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   677
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "< Feats"
      RightLinks      =   "Enhancements >"
   End
   Begin CharacterBuilderLite.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12216
      _ExtentX        =   21548
      _ExtentY        =   677
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin CharacterBuilderLite.userList usrList 
      Height          =   3012
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   6132
      _ExtentX        =   9758
      _ExtentY        =   5313
   End
   Begin VB.ComboBox cboClass 
      Height          =   312
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   1812
   End
   Begin VB.ListBox lstSpell 
      Appearance      =   0  'Flat
      Height          =   5556
      IntegralHeight  =   0   'False
      ItemData        =   "frmSpells.frx":000C
      Left            =   7560
      List            =   "frmSpells.frx":000E
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   1200
      Width           =   3732
   End
   Begin CharacterBuilderLite.userDetails usrDetails 
      Height          =   2412
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6132
      _ExtentX        =   10816
      _ExtentY        =   4255
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   312
      Left            =   10308
      TabIndex        =   6
      Top             =   840
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   550
      Appearance3D    =   -1  'True
      Max             =   9
      Value           =   1
      StepLarge       =   2
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Label lblRare 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "* Rare scrolls are indented"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   7560
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   3732
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Level:"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   9672
      TabIndex        =   5
      Top             =   876
      Width           =   540
   End
   Begin VB.Label lblDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Details"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   960
      TabIndex        =   2
      Top             =   4080
      Width           =   6132
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuSpell 
         Caption         =   "Clear this spell"
         Index           =   0
      End
      Begin VB.Menu mnuSpell 
         Caption         =   "Clear [class] spells"
         Index           =   1
      End
      Begin VB.Menu mnuSpell 
         Caption         =   "Clear all spells"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private menClass As ClassEnum
Private mlngSpellLevel As Long

Private mblnMouse As Boolean
Private mlngSourceIndex As Long
Private menDragState As DragEnum
Private msngDownX As Single
Private msngDownY As Single

Private mlngLeaveHighlighted As Long

Private mlngSlot As Long

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    InitBuildSpells
    Cascade
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Activate()
    ActivateForm oeSpells
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    UnloadForm Me, mblnOverride
End Sub

Public Sub Cascade()
    mlngLeaveHighlighted = 0
    menDragState = dragNormal
    mblnOverride = False
    LoadData Me.cboClass.Text, Me.usrSpinner.Value
    InitSlots
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.usrSpinner.hwnd, Xpos, Ypos): Me.usrSpinner.WheelScroll lngValue
        Case IsOver(Me.usrDetails.hwnd, Xpos, Ypos): Me.usrDetails.Scroll lngValue
    End Select
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Spells"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "< Feats": If Not OpenForm("frmFeats") Then Exit Sub
        Case "Enhancements >": If Not OpenForm("frmEnhancements") Then Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.usrList.Active = 0
End Sub


' ************* INITIALIZE *************


Private Sub LoadData(Optional pstrClass As String, Optional plngLevel As Long)
    Dim typClassSplit() As ClassSplitType
    Dim blnLockWindow As Boolean
    Dim i As Long
    
    blnLockWindow = Me.Visible
    If blnLockWindow Then xp.LockWindow Me.hwnd
    ' Initialize slots
    mlngSourceIndex = 0
    Me.usrList.DefineDimensions 1, 2, 2
    Me.usrList.DefineColumn 1, vbCenter, "Acquired", "Pact @ 19"
    Me.usrList.DefineColumn 2, vbCenter, "Spells"
    Me.usrList.Refresh
    ' Class(es)
    mblnOverride = True
    ComboClear Me.cboClass
    For i = 0 To GetClassSplit(typClassSplit) - 1
        With typClassSplit(i)
            If db.Class(.ClassID).CanCastSpell(1) Then
                If db.Class(.ClassID).SpellSlots(.Levels, 1) Then
                    ComboAddItem Me.cboClass, .ClassName, .ClassID
                End If
            End If
        End With
    Next
    ' Set class to previous value or first
    With Me.cboClass
        For i = 0 To .ListCount - 1
            If .List(i) = pstrClass Then Exit For
        Next
        If i >= .ListCount Then
            .ListIndex = 0
            plngLevel = 1
        Else
            .ListIndex = i
        End If
        menClass = .ItemData(.ListIndex)
        .Visible = (.ListCount > 1)
    End With
    With Me.usrSpinner
        .Max = build.Spell(menClass).MaxSpellLevel
        If plngLevel > .Max Then plngLevel = .Max
        mlngSpellLevel = plngLevel
        .Value = mlngSpellLevel
    End With
    mblnOverride = False
    InitSlots
    ShowAvailable False
    Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
    If blnLockWindow Then xp.UnlockWindow
End Sub

Private Sub InitSlots()
    Dim i As Long
    
    Me.usrList.Selected = 0
    With build.Spell(menClass).Level(mlngSpellLevel)
        If Me.usrList.Rows <> .Slots Then Me.usrList.Rows = .Slots
        For i = 1 To .Slots
            Me.usrList.SetText i, 1, GetLevelText(i)
            Me.usrList.SetSlot i, .Slot(i).Spell
        Next
    End With
End Sub


' ************* DISPLAY *************


Private Sub ShowAvailable(pblnFocus As Boolean)
    Dim lngIndex As Long
    Dim strSpell As String
    Dim blnRare As Boolean
    Dim i As Long
    
    If Me.Visible And pblnFocus Then Me.usrList.SetFocus
    lngIndex = Me.lstSpell.ListIndex
    If lngIndex <> -1 Then Me.lstSpell.Selected(lngIndex) = False
    ListboxClear Me.lstSpell
    With db.Class(menClass).SpellList(mlngSpellLevel)
        For i = 1 To .Spells
            If Not SpellTaken(.Spell(i)) Then
                strSpell = .Spell(i)
                If MakeSpellName(strSpell) Then blnRare = True
                ListboxAddItem Me.lstSpell, strSpell, i
            End If
        Next
    End With
    If blnRare Then
        If menClass = ceWizard And Me.usrSpinner.Value > 7 Then
            Me.lblRare.Caption = "* All level " & Me.usrSpinner.Value & " Wizard scrolls are rare"
        Else
            Me.lblRare.Caption = "* Rare scrolls are indented"
        End If
    End If
    Me.lblRare.Visible = blnRare
    If lngIndex > Me.lstSpell.ListCount - 1 Then lngIndex = Me.lstSpell.ListCount - 1
    Me.lstSpell.ListIndex = lngIndex
    If lngIndex <> -1 Then Me.lstSpell.Selected(lngIndex) = False
End Sub

Private Function SpellTaken(pstrSpell As String) As Boolean
    Dim i As Long
    
    With build.Spell(menClass).Level(mlngSpellLevel)
        For i = 1 To .Slots
            If .Slot(i).Spell = pstrSpell Then
                SpellTaken = True
                Exit For
            End If
        Next
    End With
End Function

Private Sub ShowDropSlots()
    Dim blnSourceFree As Boolean
    Dim blnDestFree As Boolean
    Dim i As Long
    
    With build.Spell(menClass).Level(mlngSpellLevel)
        If mlngSourceIndex = 0 Then blnSourceFree = CheckFree(menClass, Me.lstSpell.Text) Else blnSourceFree = (.Slot(mlngSourceIndex).SlotType = sseFree)
        If mlngSourceIndex = 0 And blnSourceFree Then AddFreeSlot
        For i = 1 To .Slots
            With .Slot(i)
                blnDestFree = (.SlotType = sseFree)
                If .SlotType <> sseClericCure And .SlotType <> sseWarlockPact Then
                    If blnSourceFree = blnDestFree And i <> mlngSourceIndex Then
                        Me.usrList.SetDropState i, dsCanDrop
                    Else
                        Me.usrList.SetDropState i, dsDefault
                    End If
                End If
            End With
        Next
    End With
End Sub

Private Sub AddFreeSlot()
    AddFreeSpellSlot menClass, mlngSpellLevel
    ShowSlots
End Sub

Private Sub RemoveFreeBlank()
    RemoveEmptyFreeSlots menClass, mlngSpellLevel
    ShowSlots
End Sub

Private Sub ShowSlots()
    Dim i As Long
    
    With build.Spell(menClass).Level(mlngSpellLevel)
        If Me.usrList.Rows <> .Slots Then Me.usrList.Rows = .Slots
        For i = 1 To .Slots
            Me.usrList.SetText i, 1, GetLevelText(i)
            Me.usrList.SetSlot i, .Slot(i).Spell
        Next
    End With
End Sub

Private Function GetLevelText(plngSlot As Long)
    With build.Spell(menClass).Level(mlngSpellLevel).Slot(plngSlot)
        If .SlotType = sseWarlockPact Then
            GetLevelText = "Pact @ " & .Level
        Else
            GetLevelText = "Level " & .Level
        End If
    End With
End Function


' ************* GENERAL *************


Private Sub SetSpellSlot(ByVal plngSlot As Long, pstrSpell As String)
    With build.Spell(menClass).Level(mlngSpellLevel)
        .Slot(plngSlot).Spell = pstrSpell
        Me.usrList.SetSlot plngSlot, pstrSpell
    End With
    ShowAvailable True
    SetDirty
    Me.usrList.Active = plngSlot
End Sub

Private Sub CompleteDrag()
    mlngSourceIndex = 0
    RemoveFreeBlank
    Me.usrList.Selected = mlngLeaveHighlighted
    mlngLeaveHighlighted = 0
    Me.usrList.SetFocus
End Sub


' ************* SLOTS *************


Private Sub usrList_SlotClick(Index As Integer, Button As Integer)
    With Me.lstSpell
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.SetFocus
    If Button = vbRightButton Then
        ContextMenu Index
    ElseIf Me.usrList.Selected = Index Then
        NoSelection
    Else
        Me.usrList.Selected = Index
        ShowDetails Me.usrList.GetCaption(CLng(Index))
    End If
End Sub

Private Sub ContextMenu(Index As Integer)
    Me.usrList.Selected = Index
    Me.usrList.Active = Index
    ShowDetails Me.usrList.GetCaption(CLng(Index))
    With build.Spell(menClass).Level(mlngSpellLevel).Slot(Index)
        Me.mnuSpell(0).Enabled = (.SlotType = sseStandard Or .SlotType = sseFree)
    End With
    Me.mnuSpell(1).Caption = "Clear " & db.Class(menClass).ClassName & " spells"
    Me.mnuSpell(1).Visible = (Me.cboClass.ListCount > 1)
    PopupMenu Me.mnuMain(0)
End Sub

Private Sub usrList_SlotDblClick(Index As Integer)
    ClearSlot Index
End Sub

Private Sub mnuSpell_Click(Index As Integer)
    Select Case Me.mnuSpell(Index).Caption
        Case "Clear this spell"
            Select Case build.Spell(menClass).Level(mlngSpellLevel).Slot(Me.usrList.Selected).SlotType
                Case sseClericCure, sseWarlockPact
                Case Else: ClearSlot Me.usrList.Selected
            End Select
        Case "Clear all spells"
            If Me.cboClass.ListCount > 1 Then
                If Not Ask("Clear all spells for all classes and levels?") Then Exit Sub
                ClearAllSpells
                SetDirty
            Else
                If Not Ask("Clear all spells for all levels?") Then Exit Sub
                ClearClassSpells menClass
                SetDirty
            End If
        Case Else
                If Not Ask("Clear all " & db.Class(menClass).ClassName & " spells for all levels?") Then Exit Sub
                ClearClassSpells menClass
                SetDirty
    End Select
End Sub

Private Sub ClearSlot(ByVal plngSlot As Integer)
    With build.Spell(menClass).Level(mlngSpellLevel).Slot(plngSlot)
        If .SlotType = sseStandard Or .SlotType = sseFree Then
            .Spell = vbNullString
            Me.usrList.SetSlot plngSlot, vbNullString
        End If
    End With
    ' This is outside the "With" block because it redims .Slot(), which is locked by the With block
    If build.Spell(menClass).Level(mlngSpellLevel).Slot(plngSlot).SlotType = sseFree Then RemoveFreeBlank
    ShowAvailable True
    ClearDetails
    SetDirty
End Sub

Private Sub ClearClassSpells(penClass As ClassEnum)
    Dim lngLevel As Long
    Dim lngSlot As Long
    
    For lngLevel = 1 To build.Spell(penClass).MaxSpellLevel
        For lngSlot = 1 To build.Spell(penClass).Level(lngLevel).Slots
            Select Case build.Spell(penClass).Level(lngLevel).Slot(lngSlot).SlotType
                Case sseStandard, sseFree
                    build.Spell(penClass).Level(lngLevel).Slot(lngSlot).Spell = vbNullString
            End Select
        Next
        RemoveEmptyFreeSlots penClass, lngLevel
    Next
    If penClass = menClass Then
        ShowSlots
        ShowAvailable True
        Me.usrList.Selected = 0
    End If
End Sub

Private Sub ClearAllSpells()
    Dim enClass As ClassEnum
    
    For enClass = 1 To ceClasses - 1
        If build.Spell(enClass).MaxSpellLevel > 0 Then ClearClassSpells enClass
    Next
End Sub

Private Sub usrList_RequestDrag(Index As Integer, Allow As Boolean)
    Me.usrList.Selected = 0
    Select Case build.Spell(menClass).Level(mlngSpellLevel).Slot(Index).SlotType
        Case sseStandard, sseFree
            mlngLeaveHighlighted = Index
            mlngSourceIndex = Index
            ShowDropSlots
            Allow = True
            ShowDetails build.Spell(menClass).Level(mlngSpellLevel).Slot(Index).Spell
    End Select
End Sub

Private Sub usrList_OLEDragDrop(Index As Integer, Data As DataObject)
    Dim strSpell As String
    Dim strSwap As String
    
    If mlngSourceIndex = 0 Then
        strSpell = Me.lstSpell.Text
        SetSpellSlot Index, strSpell
        ShowDetails strSpell
    Else
        With build.Spell(menClass).Level(mlngSpellLevel)
            strSwap = .Slot(Index).Spell
            .Slot(Index).Spell = .Slot(mlngSourceIndex).Spell
            .Slot(mlngSourceIndex).Spell = strSwap
            Me.usrList.SetSlot Index, .Slot(Index).Spell
            Me.usrList.SetSlot mlngSourceIndex, .Slot(mlngSourceIndex).Spell
            Me.usrList.Selected = Index
        End With
        ShowAvailable True
        SetDirty
    End If
    mlngLeaveHighlighted = Index
End Sub

Private Sub usrList_OLECompleteDrag(Index As Integer, Effect As Long)
    CompleteDrag
End Sub


' ************* RARE SCROLLS *************


Private Function MakeSpellName(pstrSpell As String) As Boolean
    Dim lngSpell As Long
    
    If Not (menClass = ceWizard Or menClass = ceArtificer) Then Exit Function
    lngSpell = SeekSpell(pstrSpell)
    If lngSpell = 0 Then Exit Function
    If Not db.Spell(lngSpell).Rare Then Exit Function
    pstrSpell = "  " & pstrSpell & "*"
    MakeSpellName = True
End Function

Private Function CleanSpellName(pstrSpell As String) As String
    Dim strReturn As String
    
    strReturn = Trim$(pstrSpell)
    If Right$(strReturn, 1) = "*" Then strReturn = Left$(strReturn, Len(strReturn) - 1)
    CleanSpellName = strReturn
End Function


' ************* SPELL LIST *************


Private Sub lstSpell_Click()
    If mblnMouse Then mblnMouse = False Else ShowDetails Me.lstSpell.Text
End Sub

Private Sub lstSpell_DblClick()
    ChooseSpell
End Sub

Private Sub ChooseSpell()
    Dim strSpell As String
    Dim lngSlot As Long
    Dim i As Long
    
    strSpell = CleanSpellName(Me.lstSpell.Text)
    If CheckFree(menClass, strSpell) Then
        AddFreeSpellSlot menClass, mlngSpellLevel
        ShowSlots
        lngSlot = 1
    Else
        With build.Spell(menClass).Level(mlngSpellLevel)
            For i = 1 To .Slots
                If Len(.Slot(i).Spell) = 0 Then
                    lngSlot = i
                    Exit For
                End If
            Next
        End With
    End If
    If lngSlot Then
        SetSpellSlot lngSlot, strSpell
        Me.usrList.Selected = lngSlot
        ShowDetails strSpell
    End If
End Sub

Private Sub lstSpell_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.lstSpell.ListIndex = -1 Or Button <> vbLeftButton Then Exit Sub
    Me.usrList.Selected = 0
    Me.usrList.Active = 0
    mblnMouse = True
    ShowDetails Me.lstSpell.Text
    menDragState = dragMouseDown
End Sub

Private Sub lstSpell_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            mlngSourceIndex = 0
            ShowDropSlots
            Me.lstSpell.OLEDrag
        End If
    End If
End Sub

Private Sub lstSpell_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub lstSpell_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List"
End Sub

Private Sub lstSpell_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngSourceIndex = 0 Then Exit Sub
    ClearSlot mlngSourceIndex
    mlngSourceIndex = 0
    ShowAvailable True
    Me.usrList.Selected = mlngSourceIndex
    SetDirty
End Sub

Private Sub lstSpell_OLECompleteDrag(Effect As Long)
    CompleteDrag
End Sub

Private Sub lstSpell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyReturn: ChooseSpell
    End Select
End Sub


' ************* FILTERS *************


Private Sub cboClass_Click()
    If mblnOverride Then Exit Sub
    menClass = ComboGetValue(Me.cboClass)
    PopulateSpellLevels
    SaveBackup
End Sub

Private Sub PopulateSpellLevels()
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    With Me.usrSpinner
        .Value = 1
        .Max = build.Spell(menClass).MaxSpellLevel
    End With
    SpellLevelChange
End Sub

Private Sub usrSpinner_Change()
    If mblnOverride Then Exit Sub
    SpellLevelChange
    SaveBackup
End Sub

Private Sub SpellLevelChange()
    If mblnOverride Then Exit Sub
    mlngSpellLevel = Me.usrSpinner.Value
    InitSlots
    ShowAvailable False
    Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
End Sub


' ************* DETAILS *************


Private Sub ShowDetails(ByVal pstrSpell As String)
    Dim lngIndex As Long
    Dim strSpell As String
    Dim strLink As String
    
    pstrSpell = CleanSpellName(pstrSpell)
    Me.usrDetails.Clear
    If Len(pstrSpell) = 0 Then
        Me.lblDetails.Caption = "Details"
        Exit Sub
    Else
        Me.lblDetails.Caption = pstrSpell
        lngIndex = SeekSpell(pstrSpell)
        If lngIndex = 0 Then
            Me.usrDetails.AddText "Spell not found."
        Else
            With db.Spell(lngIndex)
                Me.usrDetails.AddDescrip .Descrip, MakeWiki(.Wiki)
            End With
        End If
    End If
    Me.usrDetails.Refresh
End Sub

Private Sub ClearDetails()
    With Me.lstSpell
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.Selected = 0
    Me.usrList.SetFocus
    Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
End Sub

Private Sub NoSelection()
    With Me.lstSpell
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.Selected = 0
    Me.usrList.SetFocus
    ClearDetails
End Sub
