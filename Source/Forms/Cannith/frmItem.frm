VERSION 5.00
Begin VB.Form frmItem 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cannith Crafting Item Planner"
   ClientHeight    =   9024
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13548
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9024
   ScaleWidth      =   13548
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8640
      Width           =   13548
      _ExtentX        =   23897
      _ExtentY        =   677
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "New Item;New Gearset;Load Gearset"
      RightLinks      =   "Effects;Materials;Augments;Scaling"
   End
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13548
      _ExtentX        =   23897
      _ExtentY        =   677
      Spacing         =   264
      BorderColor     =   -2147483640
      LeftLinks       =   "Effects;Review"
      RightLinks      =   "Help"
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8112
      Index           =   0
      Left            =   120
      ScaleHeight     =   8112
      ScaleWidth      =   13332
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   13332
      Begin CannithCrafting.userInfo usrinfoWeapon 
         Height          =   1812
         Left            =   300
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   6180
         Width           =   2412
         _ExtentX        =   4255
         _ExtentY        =   3196
         TitleSize       =   2
         TitleIcon       =   0   'False
         CanScroll       =   0   'False
      End
      Begin CannithCrafting.userInfo usrinfoPrefix 
         Height          =   2352
         Left            =   3060
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5640
         Width           =   3072
         _ExtentX        =   5419
         _ExtentY        =   4149
         TitleSize       =   1
         TitleText       =   "Level:"
         CanScroll       =   0   'False
      End
      Begin VB.ListBox lstGear 
         Appearance      =   0  'Flat
         Height          =   4560
         ItemData        =   "frmItem.frx":08CA
         Left            =   300
         List            =   "frmItem.frx":090D
         TabIndex        =   24
         Top             =   960
         Width           =   2412
      End
      Begin VB.ListBox lstPrefix 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   3060
         TabIndex        =   26
         Top             =   960
         Width           =   3072
      End
      Begin VB.ListBox lstSuffix 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   6420
         TabIndex        =   28
         Top             =   960
         Width           =   3072
      End
      Begin VB.ListBox lstExtra 
         Appearance      =   0  'Flat
         Height          =   4560
         Left            =   9780
         TabIndex        =   30
         Top             =   960
         Width           =   3252
      End
      Begin CannithCrafting.userSpinner usrspnML 
         Height          =   312
         Left            =   780
         TabIndex        =   32
         Top             =   5700
         Width           =   852
         _ExtentX        =   1503
         _ExtentY        =   550
         Max             =   34
         Value           =   34
         StepLarge       =   3
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   0
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CannithCrafting.userCheckBox usrchkBound 
         Height          =   252
         Left            =   1800
         TabIndex        =   33
         Top             =   5712
         Width           =   1152
         _ExtentX        =   2032
         _ExtentY        =   445
         Caption         =   "Bound"
      End
      Begin CannithCrafting.userInfo usrinfoSuffix 
         Height          =   2352
         Left            =   6420
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   5640
         Width           =   3072
         _ExtentX        =   5419
         _ExtentY        =   4149
         TitleSize       =   1
         TitleText       =   "Level:"
         CanScroll       =   0   'False
      End
      Begin CannithCrafting.userInfo usrinfoExtra 
         Height          =   2352
         Left            =   9780
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5640
         Width           =   3252
         _ExtentX        =   5736
         _ExtentY        =   4149
         TitleSize       =   1
         TitleText       =   "Level:"
         CanScroll       =   0   'False
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   1
         Left            =   1140
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   3
         Left            =   2220
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   4
         Left            =   2760
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   5
         Left            =   3300
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   6
         Left            =   3840
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   7
         Left            =   4380
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   8
         Left            =   4920
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   9
         Left            =   5460
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   10
         Left            =   6300
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   11
         Left            =   6840
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   12
         Left            =   7380
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   13
         Left            =   7920
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   14
         Left            =   8760
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   15
         Left            =   9300
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   16
         Left            =   9840
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   17
         Left            =   10680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   18
         Left            =   11220
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   20
         Left            =   12300
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   19
         Left            =   11760
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin CannithCrafting.userIcon usrIcon 
         Height          =   408
         Index           =   0
         Left            =   600
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   408
         _ExtentX        =   720
         _ExtentY        =   720
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "ML"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   5736
         Width           =   252
      End
      Begin VB.Label lblPrefix 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prefix"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3060
         TabIndex        =   25
         Top             =   660
         Width           =   3072
      End
      Begin VB.Label lblSuffix 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Suffix"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6420
         TabIndex        =   27
         Top             =   660
         Width           =   3072
      End
      Begin VB.Label lblExtra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Extra"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   9780
         TabIndex        =   29
         Top             =   660
         Width           =   3252
      End
      Begin VB.Label lblGearSlot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Gear Slot"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   300
         TabIndex        =   23
         Top             =   660
         Width           =   2412
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7872
      Index           =   1
      Left            =   120
      ScaleHeight     =   7872
      ScaleWidth      =   13332
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   13332
      Begin CannithCrafting.userCheckBox usrchkBoundRecipe 
         Height          =   252
         Left            =   5700
         TabIndex        =   41
         Top             =   360
         Width           =   1212
         _ExtentX        =   3196
         _ExtentY        =   445
         Caption         =   "Bound"
         CheckPosition   =   1
      End
      Begin CannithCrafting.userInfo usrinfoReview 
         Height          =   5952
         Left            =   540
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   300
         Width           =   6372
         _ExtentX        =   11240
         _ExtentY        =   10499
         TitleText       =   "Gear Slot"
      End
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSpellPowerGroup As String
Private menGear As GearEnum
Private mblnOverride As Boolean
Private mlngTab As Long

Private mtypItem As OpenItemType


' ************* FORM *************


Private Sub Form_Load()
    Me.Tag = Me.Caption
    menGear = geUnknown
    mblnOverride = False
    cfg.RefreshColors Me
    InitControls
    ShowDetails
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    CloseApp
End Sub

Private Sub InitControls()
    Dim strML As String
    Dim i As Long
    
    Me.usrinfoPrefix.Clear
    mblnOverride = True
    Me.lstGear.Clear
    For i = 0 To geGearCount - 1
        Me.lstGear.AddItem GetGearName(i, False)
        Me.usrIcon(i).Init i, uiseToggle, (i <> ge2hMelee)
    Next
    mblnOverride = False
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    If IsOver(Me.usrspnML.hwnd, Xpos, Ypos) Then Me.usrspnML.WheelScroll Rotation
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    If pstrCaption = "Help" Then
        ShowHelp "Items"
        Exit Sub
    End If
    Select Case pstrCaption
        Case "Effects"
            ShowDetails
            ShowTab 0
        Case "Augments"
            ShowTab 1
        Case "Stone of Change"
            ShowTab 2
        Case "Review"
            ShowIngredients
            ShowTab 1
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    Dim frm As Form
    
    If pstrCaption = "New Item" Then
        If Me.lstGear.ListIndex = -1 Then
            MsgBox "This is already a new item.", vbInformation, "Notice:"
        Else
            Set frm = New frmItem
            frm.usrspnML.Value = Me.usrspnML.Value
            frm.Show
            Set frm = Nothing
        End If
    Else
        If FooterClick(pstrCaption) Then
            If Me.lstGear.ListIndex = -1 Then Unload Me
        End If
    End If
End Sub

Private Sub ShowTab(plngTab As Long)
    Dim i As Long
    
    If mlngTab = plngTab Then Exit Sub
    mlngTab = plngTab
    For i = 0 To Me.picTab.UBound
        Me.picTab(i).Visible = (i = mlngTab)
    Next
End Sub


' ************* GEARSET *************


Public Sub OpenGearsetItem()
    Dim typBlank As OpenItemType
    
    mtypItem = gtypOpenItem
    gtypOpenItem = typBlank
    With mtypItem
        SetItemType .Mainhand
        SetItemType .Offhand
        SetItemType .Armor
        Me.usrspnML.Value = .ML
        SelectGear .Gear
        ListboxSetValue Me.lstPrefix, .Prefix
        ListboxSetValue Me.lstSuffix, .Suffix
        ListboxSetValue Me.lstExtra, .Extra
    End With
End Sub

Private Sub SetItemType(pstrItem As String)
    Dim lngIndex As Long
    Dim enGear As GearEnum
    
    If Len(pstrItem) = 0 Then Exit Sub
    lngIndex = SeekItem(pstrItem)
    If lngIndex = 0 Then Exit Sub
    Select Case db.Item(lngIndex).ItemStyle
        Case iseMelee2H: If pstrItem = "Handwraps" Then enGear = geHandwraps Else enGear = ge2hMelee
        Case iseMelee1H: enGear = ge1hMelee
        Case iseShield: enGear = geShield
        Case iseRange, iseThrower: enGear = geRange
        Case iseMetalArmor: enGear = geMetalArmor
        Case iseLeatherArmor: enGear = geLeatherArmor
        Case iseClothArmor: enGear = geClothArmor
        Case Else: Exit Sub
    End Select
    Me.usrIcon(enGear).IconName = db.Item(lngIndex).ItemName
End Sub



' ************* SLOT *************


Private Sub lstGear_Click()
    If mblnOverride Then Exit Sub
    SelectGear Me.lstGear.ListIndex
End Sub

Private Sub usrIcon_ActiveChange(Index As Integer, pblnActive As Boolean)
    If mblnOverride Then Exit Sub
    GearClick Index
End Sub

Private Sub usrIcon_StyleChange(Index As Integer, pstrStyle As String)
    ShowFormIcon
End Sub

Private Sub GearClick(Index As Integer)
    xp.SetMouseCursor mcHand
    If mblnOverride Then Exit Sub
    If menGear = Index Then SelectGear geUnknown Else SelectGear Index
    ShowDetails
End Sub
'
'Public Property Get Gear() As GearEnum
'    Gear = menGear
'End Property
'
'Public Property Let Gear(ByVal penGear As GearEnum)
'    SelectGear penGear
'End Property
'
'Public Function GetEffect(penAffix As AffixEnum, plngShard As Long) As Long
'    Select Case penAffix
'        Case aePrefix: GetEffect = ListboxGetValue(Me.lstPrefix)
'        Case aeSuffix: GetEffect = ListboxGetValue(Me.lstSuffix)
'        Case aeExtra: GetEffect = ListboxGetValue(Me.lstExtra)
'    End Select
'End Function
'
'Public Sub SetEffect(penAffix As AffixEnum, plngShard As Long)
'    Select Case penAffix
'        Case aePrefix: ListboxSetValue Me.lstPrefix, plngShard
'        Case aeSuffix: ListboxSetValue Me.lstSuffix, plngShard
'        Case aeExtra: ListboxSetValue Me.lstExtra, plngShard
'    End Select
'End Sub

Private Sub SelectGear(ByVal penGear As GearEnum)
    Dim lngPrefix As Long
    Dim lngSuffix As Long
    Dim lngExtra As Long
    Dim lngPrefixTop As Long
    Dim lngSuffixTop As Long
    Dim lngExtraTop As Long
    
    With Me.lstPrefix
        If .ListIndex <> -1 Then lngPrefix = .ItemData(.ListIndex)
        lngPrefixTop = .ListIndex - .TopIndex
    End With
    With Me.lstSuffix
        If .ListIndex <> -1 Then lngSuffix = .ItemData(.ListIndex)
        lngSuffixTop = .ListIndex - .TopIndex
    End With
    With Me.lstExtra
        If .ListIndex <> -1 Then
            lngExtra = .ItemData(.ListIndex)
            lngExtraTop = .ListIndex - .TopIndex
        End If
    End With
    mblnOverride = True
    Me.lstGear.ListIndex = penGear
    mblnOverride = False
    ShowPicture False
    menGear = penGear
    ShowPicture True
    PopulateLists
    If menGear = geUnknown Then Me.Caption = Me.Tag Else Me.Caption = GetGearName(menGear)
    mblnOverride = True
    RememberChoice Me.lstPrefix, lngPrefix, lngPrefixTop
    RememberChoice Me.lstSuffix, lngSuffix, lngSuffixTop
    RememberChoice Me.lstExtra, lngExtra, lngExtraTop
    mblnOverride = False
    If penGear = geUnknown Then Me.picTab(0).SetFocus '~ throws runtime error if you add a crafted item that was named, go immediately to the slotting screen and click that item icon
    ShowDetails
End Sub

Private Sub RememberChoice(plst As ListBox, plngItemData As Long, plngTopIndex As Long)
    Dim i As Long
    
    If plngItemData = 0 Then Exit Sub
    For i = 0 To plst.ListCount - 1
        If plst.ItemData(i) = plngItemData Then
            plst.ListIndex = i
            If i > plngTopIndex And plngTopIndex > 0 Then plst.TopIndex = i - plngTopIndex
        End If
    Next
End Sub

Private Sub ShowPicture(pblnSelected As Boolean)
    If menGear <> geUnknown Then
        mblnOverride = True
        Me.usrIcon(menGear).Active = pblnSelected
        mblnOverride = False
    End If
    ShowFormIcon
End Sub

Private Sub ShowFormIcon()
    SetFormIcon Me, GetItemResource(GetIconName())
End Sub

Private Function GetIconName() As String
    If menGear = geUnknown Then GetIconName = "Trinket" Else GetIconName = Me.usrIcon(menGear).IconName
End Function

Private Sub PopulateLists()
    Dim i As Long
    
    mblnOverride = True
    ListboxClear Me.lstPrefix
    ListboxClear Me.lstSuffix
    ListboxClear Me.lstExtra
    If menGear <> geUnknown Then
        For i = 1 To db.Shards
            With db.Shard(db.ShardIndex(i))
                If .Prefix(menGear) Then ListboxAdd Me.lstPrefix, .Group & ": " & .Abbreviation, db.ShardIndex(i)
                If .Suffix(menGear) Then ListboxAdd Me.lstSuffix, .Group & ": " & .Abbreviation, db.ShardIndex(i)
                If .Extra(menGear) Then ListboxAdd Me.lstExtra, .Group & ": " & .Abbreviation, db.ShardIndex(i)
            End With
        Next
        ListboxAdd Me.lstPrefix, "<None>", 0
        ListboxAdd Me.lstSuffix, "<None>", 0
        ListboxAdd Me.lstExtra, "<None>", 0
    End If
    mblnOverride = False
End Sub

Private Sub ListboxClear(plst As ListBox)
    plst.ListIndex = -1
    plst.Clear
End Sub

Private Sub ListboxAdd(plst As ListBox, pstrText As String, plngItemData As Long)
    plst.AddItem pstrText
    plst.ItemData(plst.NewIndex) = plngItemData
End Sub


' ************* EFFECTS *************


Private Sub lstPrefix_Click()
    If mblnOverride Then Exit Sub
    If Me.lstPrefix.ListIndex = Me.lstPrefix.ListCount - 1 Then ClearChoice Me.lstPrefix Else ShowDetails
End Sub

Private Sub lstPrefix_DblClick()
    OpenShardDetail Me.lstPrefix
End Sub

Private Sub lstSuffix_Click()
    If mblnOverride Then Exit Sub
    If Me.lstSuffix.ListIndex = Me.lstSuffix.ListCount - 1 Then ClearChoice Me.lstSuffix Else ShowDetails
End Sub

Private Sub lstSuffix_DblClick()
    OpenShardDetail Me.lstSuffix
End Sub

Private Sub lstExtra_Click()
    If mblnOverride Then Exit Sub
    If Me.lstExtra.ListIndex = Me.lstExtra.ListCount - 1 Then ClearChoice Me.lstExtra Else ShowDetails
End Sub

Private Sub lstExtra_DblClick()
    OpenShardDetail Me.lstExtra
End Sub

Private Sub OpenShardDetail(plst As ListBox)
    If plst.ListIndex = -1 Then Exit Sub
    OpenShard db.Shard(plst.ItemData(plst.ListIndex)).ShardName
End Sub

Private Sub ClearChoice(plst As ListBox)
    plst.ListIndex = -1
    plst.TopIndex = 0
    On Error Resume Next
    Me.picTab(0).SetFocus
End Sub

Private Sub usrspnML_Change()
    ShowDetails
End Sub

Private Sub usrchkBound_UserChange()
    Me.usrchkBoundRecipe.Value = Me.usrchkBound.Value
    ShowDetails
End Sub

Private Sub ShowDetails()
    Dim lngML As Long
    
    Me.usrinfoWeapon.Clear
    lngML = Me.usrspnML.Value
    Me.usrinfoWeapon.AddText BindingText(Me.usrchkBound.Value) & " Level: " & MLShardLevel(lngML, Me.usrchkBound.Value)
    Me.usrinfoWeapon.AddText "Essences: " & MLShardEssences(lngML, Me.usrchkBound.Value), 2
    Select Case menGear
        Case geMetalArmor, geLeatherArmor, geClothArmor, geDocent
            Me.usrinfoWeapon.AddText "+" & GetScale("Enhancement Bonus", lngML) & " Armor Bonus"
        Case geDocent
            Me.usrinfoWeapon.AddText "+" & GetScale("Enhancement Bonus", lngML) & " Docent"
        Case ge1hMelee, ge2hMelee, geHandwraps, geRange
            Me.usrinfoWeapon.AddText "+" & GetScale("Enhancement Bonus", lngML) & " Enhancement Bonus"
            Me.usrinfoWeapon.AddText GetScale("Base Damage", lngML) & " Base Damage"
            If HasSpellpower() Then Me.usrinfoWeapon.AddText "+" & GetScale("Implement Bonus", lngML) & " Implement Bonus"
        Case geShield
            Me.usrinfoWeapon.AddText "+" & GetScale("Enhancement Bonus", lngML) & " Shield Bonus"
        Case geOrb
            Me.usrinfoWeapon.AddText "+" & GetScale("Enhancement Bonus", lngML) & " Orb Bonus"
            If HasSpellpower() Then Me.usrinfoWeapon.AddText "+" & GetScale("Implement Bonus", lngML) & " Implement Bonus"
    End Select
    ShowChoice Me.lstPrefix, Me.usrinfoPrefix
    ShowChoice Me.lstSuffix, Me.usrinfoSuffix
    ShowChoice Me.lstExtra, Me.usrinfoExtra, True
End Sub

Private Function HasSpellpower() As Boolean
    HasSpellpower = True
    If IsSpellpower(Me.lstPrefix) Then Exit Function
    If IsSpellpower(Me.lstSuffix) Then Exit Function
    If IsSpellpower(Me.lstExtra) Then Exit Function
    HasSpellpower = False
End Function

Private Function IsSpellpower(plst As ListBox) As Boolean
    If plst.ListIndex = -1 Then Exit Function
    If Len(mstrSpellPowerGroup) = 0 Then FindSpellpowerGroup
    With db.Shard(plst.ItemData(plst.ListIndex))
        If .Group = mstrSpellPowerGroup And Me.usrspnML.Value >= .ML Then IsSpellpower = True
    End With
End Function

' Should probably be a binary search, but only ever runs once per form instance
' Combustion is the 28th shard in list as of writing this, so, meh.
Private Sub FindSpellpowerGroup()
    Dim i As Long
    
    For i = 1 To db.Shards
        If db.Shard(i).ShardName = "Combustion" Then
            mstrSpellPowerGroup = db.Shard(i).Group
            Exit For
        End If
    Next
End Sub

Private Function GetScale(pstrScaleName As String, plngML As Long)
    Dim lngIndex As Long
    
    lngIndex = SeekScaling(pstrScaleName)
    If lngIndex Then GetScale = db.Scaling(lngIndex).Table(plngML)
End Function

Private Sub ShowChoice(plst As ListBox, pinfo As userInfo, Optional pblnExtra As Boolean = False)
    Dim lngScale As Long
    
    pinfo.Clear
    If plst.ListIndex = -1 Then Exit Sub
    If plst.ItemData(plst.ListIndex) = 0 Then Exit Sub
    If pblnExtra And Me.usrspnML.Value < 10 Then
        pinfo.AddError "Extra slots are only available"
        pinfo.AddError "at ML10 or higher"
    Else
        With db.Shard(plst.ItemData(plst.ListIndex))
            If .ML > Me.usrspnML.Value Then
                pinfo.AddError .Abbreviation & " is ML" & .ML
            Else
                If .ScaleName <> "None" Then lngScale = SeekScaling(.ScaleName)
                If lngScale Then
                    pinfo.TitleText = "Level " & Me.usrspnML.Value & ":"
                    pinfo.AddLink .Abbreviation, lseShard, .ShardName, 0
                    pinfo.AddText " " & db.Scaling(lngScale).Table(Me.usrspnML.Value)
                    pinfo.AddText Scaling(lngScale, plst.ItemData(plst.ListIndex)), 2
                Else
                    pinfo.TitleText = vbNullString
                    pinfo.AddLink .Abbreviation, lseShard, .ShardName, 3
                End If
                If Me.usrchkBound.Value Then
                    pinfo.AddText "Bound Level " & .Bound.Level
                    AddRecipeToInfo .Bound, pinfo
                Else
                    pinfo.AddText "Unbound Level " & .Unbound.Level
                    AddRecipeToInfo .Unbound, pinfo
                End If
            End If
        End With
    End If
End Sub

Private Function Scaling(plngScale As Long, plngShard As Long) As String
    Dim lngML As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim iMin As Long
    Dim i As Long
    
    lngML = Me.usrspnML.Value
    lngFirst = lngML
    lngLast = lngML
    iMin = db.Shard(plngShard).ML
    If iMin < 1 Then iMin = 1
    With db.Scaling(plngScale)
        ' Do this in two different loops because scales may not be filled in (and it's also more efficient)
        ' So if scale is "? ? ? 3 ? ? 4 ? ? ? 5 ? ?" and actual is a ? between 3 and 4, lngFirst and lngLast should be inside 3 and 4
        For i = lngML - 1 To iMin Step -1
            If .Table(i) <> .Table(lngML) Then Exit For
            lngFirst = i
        Next
        For i = lngML + 1 To 34
            If .Table(i) <> .Table(lngML) Then Exit For
            lngLast = i
        Next
        If lngFirst <> lngML Or lngLast <> lngML Then
            Scaling = "Scale: ML" & lngFirst & " to ML" & lngLast
        Else
            Scaling = "Scale: ML" & lngML
        End If
    End With
End Function


' ************* INGREDIENTS *************


Private Sub usrchkBoundRecipe_UserChange()
    Me.usrchkBound.Value = Me.usrchkBoundRecipe.Value
    ShowIngredients
End Sub

Private Sub ShowIngredients()
    Const Indent As String = "   "
    Dim typRecipe As RecipeType
    Dim lngPrefix As Long
    Dim lngSuffix As Long
    Dim lngExtra As Long
    Dim i As Long
    
    Me.usrinfoReview.Clear
    If menGear = geUnknown Then Exit Sub
    GatherItemRecipe typRecipe, lngPrefix, lngSuffix, lngExtra
    With Me.usrinfoReview
        .SetIcon GetItemResource(GetIconName())
        .TitleText = GetIconName()
        .AddText "Minimum Level: " & Me.usrspnML.Value, 2
        ShowSlot "Prefix:", lngPrefix
        ShowSlot "Suffix:", lngSuffix
        ShowSlot "Extra:", lngExtra
        .AddText vbNullString
        If Me.usrchkBoundRecipe.Value Then .AddText "Bound Crafting:", 0 Else .AddText "Unbound Crafting:", 0
        .AddClipboard ClipboardText(typRecipe, lngPrefix, lngSuffix, lngExtra)
        .AddText "Crafting Level " & typRecipe.Level, 1, Indent
        .AddText typRecipe.Essences & " Essences", 1, Indent
        For i = 1 To typRecipe.Ingredients
            .AddNumber typRecipe.Ingredient(i).Count, 2, Indent
'            .AddText typRecipe.Ingredient(i).Count & " ", 0, Indent
            Me.usrinfoReview.AddLink Pluralized(typRecipe.Ingredient(i)), lseMaterial, typRecipe.Ingredient(i).Material
        Next
    End With
End Sub

Private Sub GatherItemRecipe(ptypRecipe As RecipeType, plngPrefix As Long, plngSuffix As Long, plngExtra As Long)
    With ptypRecipe
        .Level = MLShardLevel(Me.usrspnML.Value, Me.usrchkBoundRecipe.Value)
        .Essences = MLShardEssences(Me.usrspnML.Value, Me.usrchkBoundRecipe.Value)
    End With
    plngPrefix = GetShardIndex(Me.lstPrefix)
    plngSuffix = GetShardIndex(Me.lstSuffix)
    If Me.usrspnML.Value >= 10 Then plngExtra = GetShardIndex(Me.lstExtra)
    If Me.usrchkBoundRecipe.Value Then
        If plngPrefix Then AggregateRecipe ptypRecipe, db.Shard(plngPrefix).Bound
        If plngSuffix Then AggregateRecipe ptypRecipe, db.Shard(plngSuffix).Bound
        If plngExtra Then AggregateRecipe ptypRecipe, db.Shard(plngExtra).Bound
    Else
        If plngPrefix Then AggregateRecipe ptypRecipe, db.Shard(plngPrefix).Unbound
        If plngSuffix Then AggregateRecipe ptypRecipe, db.Shard(plngSuffix).Unbound
        If plngExtra Then AggregateRecipe ptypRecipe, db.Shard(plngExtra).Unbound
    End If
    AggregateRecipeSort ptypRecipe
End Sub

Private Function GetShardIndex(plst As ListBox) As Long
    Dim lngShard As Long
    
    If plst.ListIndex = -1 Then Exit Function
    lngShard = plst.ItemData(plst.ListIndex)
    If db.Shard(lngShard).ML <= Me.usrspnML.Value Then GetShardIndex = lngShard
End Function

Private Sub ShowSlot(pstrSlot As String, plngShard As Long)
    Dim lngScale As Long
    Dim strScale As String
    
    If plngShard = 0 Then Exit Sub
    Me.usrinfoReview.AddText pstrSlot, 0
    With db.Shard(plngShard)
        Me.usrinfoReview.AddLink .ShardName, lseShard, .ShardName, 0
        If .ScaleName <> "None" Then
            lngScale = SeekScaling(.ScaleName)
            strScale = " " & db.Scaling(lngScale).Table(Me.usrspnML.Value)
        End If
        Me.usrinfoReview.AddText strScale
    End With
End Sub

Private Function ClipboardText(ptypRecipe As RecipeType, plngPrefix As Long, plngSuffix As Long, plngExtra As Long) As String
    Dim strReturn As String
    Dim strShards As String
    Dim i As Long
    
    If Me.usrchkBoundRecipe.Value Then strReturn = "Bound" Else strReturn = "Unbound"
    strReturn = strReturn & " Shards: ML" & Me.usrspnML.Value
    If plngPrefix Then strReturn = strReturn & ", " & db.Shard(plngPrefix).ShardName
    If plngSuffix Then strReturn = strReturn & ", " & db.Shard(plngSuffix).ShardName
    If plngExtra Then strReturn = strReturn & ", " & db.Shard(plngExtra).ShardName
    strReturn = strReturn & vbNewLine
    strReturn = strReturn & "Highest Crafting Level: " & ptypRecipe.Level & vbNewLine
    strReturn = strReturn & ptypRecipe.Essences & " Essences" & vbNewLine
    With ptypRecipe
        For i = 1 To .Ingredients
            strReturn = strReturn & .Ingredient(i).Count & " " & Pluralized(.Ingredient(i)) & vbNewLine
        Next
    End With
    ClipboardText = strReturn
End Function
