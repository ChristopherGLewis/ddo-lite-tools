VERSION 5.00
Begin VB.UserControl userGearSlots 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2880
   ScaleWidth      =   7740
   Begin CannithCrafting.userGearSlot usrSlot 
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   677
   End
End
Attribute VB_Name = "userGearSlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event CraftedChange(ByVal Slot As SlotEnum, ByVal Crafted As Boolean)
Public Event ItemTypeChange(ByVal Slot As SlotEnum, ByVal ItemType As String, HideOffhand As Boolean)
Public Event NamedItemChange(ByVal Slot As SlotEnum, ByVal NamedItem As String)
Public Event AugmentClick(ByVal Slot As SlotEnum, Left As Long, Top As Long, Right As Long, Bottom As Long)
Public Event EldritchChange(ByVal Slot As SlotEnum, ByVal Ritual As Long)
Public Event Tooltip(TooltipText As String, Left As Long, Top As Long, Height As Long)

Private mlngHeight As Long
Private mlngMarginY As Long

Private mlngTipIndex As Long
Private mblnTwoHand As Boolean
Private mstrTipText As String

Private mblnOverride As Boolean

Public Sub Init()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim i As Long
    
    With UserControl
        .usrSlot(0).SetHeight
        GetHeight
        lngLeft = 0
        lngTop = 0
        lngWidth = .ScaleWidth
        For i = 0 To seSlotCount - 1
            If i > .usrSlot.UBound Then Load .usrSlot(i)
            .usrSlot(i).Move lngLeft, lngTop, lngWidth, mlngHeight
            .usrSlot(i).Visible = True
            lngTop = lngTop + mlngHeight + mlngMarginY
        Next
        mblnOverride = True
        .Height = lngTop - mlngMarginY
        mblnOverride = False
        .usrSlot(0).Init 0, "Helmet", vbNullString
        .usrSlot(1).Init 1, "Goggles", vbNullString
        .usrSlot(2).Init 2, "Necklace", vbNullString
        .usrSlot(3).Init 3, "Cloak", vbNullString
        .usrSlot(4).Init 4, "Bracers", vbNullString
        .usrSlot(5).Init 5, "Gloves", vbNullString
        .usrSlot(6).Init 6, "Belt", vbNullString
        .usrSlot(7).Init 7, "Boots", vbNullString
        .usrSlot(8).Init 8, "Ring", vbNullString
        .usrSlot(9).Init 9, "Ring", vbNullString
        .usrSlot(10).Init 10, "Trinket", vbNullString
        .usrSlot(11).Init 11, "Full Plate", vbNullString
        .usrSlot(12).Init 12, db.Melee1H.Default.Choice, vbNullString
        .usrSlot(13).Init 13, "Heavy Shield", vbNullString
    End With
End Sub

Public Sub RefreshColors()
    Dim i As Long
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        For i = 0 To .usrSlot.UBound
            .usrSlot(i).RefreshColors
        Next
    End With
End Sub

Private Sub GetHeight()
    With UserControl
        mlngMarginY = .ScaleY(3, vbPixels, vbTwips)
        mlngHeight = .usrSlot(0).Height
    End With
End Sub

Private Sub UserControl_Resize()
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    With UserControl
        For i = 0 To .usrSlot.UBound
            .usrSlot(i).Width = .ScaleWidth
        Next
    End With
    mblnOverride = True
    GetHeight
    UserControl.Height = seSlotCount * (mlngHeight + mlngMarginY) - mlngMarginY
    mblnOverride = False
End Sub

Public Sub GetCoords(plngColumn As Long, plngLeft As Long, plngRight As Long)
    UserControl.usrSlot(0).GetCoords plngColumn, plngLeft, plngRight
End Sub

Private Sub usrSlot_Tooltip(Index As Integer, TooltipText As String, Left As Long, Top As Long, Height As Long)
    If Len(TooltipText) = 0 Then
        ClearTip
    Else
        If mstrTipText = TooltipText And mlngTipIndex <> Index Then ClearTip
        mstrTipText = TooltipText
        mlngTipIndex = Index
        With UserControl.usrSlot(Index)
            RaiseEvent Tooltip(TooltipText, .Left + Left, .Top + Top, Height)
        End With
    End If
End Sub

Private Sub ClearTip()
    RaiseEvent Tooltip(vbNullString, 0, 0, 0)
End Sub


' ************* CRAFTED *************


Public Function GetCrafted(penSlot As SlotEnum) As Boolean
    GetCrafted = UserControl.usrSlot(penSlot).Crafted
End Function

Public Sub SetCrafted(penSlot As SlotEnum, pblnCrafted As Boolean)
    UserControl.usrSlot(penSlot).Crafted = pblnCrafted
End Sub

Private Sub usrSlot_CraftedChange(Index As Integer, Crafted As Boolean)
    RaiseEvent CraftedChange(Index, Crafted)
End Sub


' ************* ITEM STYLE *************


Public Function GetItemStyle(penSlot As SlotEnum) As String
    GetItemStyle = UserControl.usrSlot(penSlot).ItemType
End Function

Public Sub SetItemStyle(penSlot As SlotEnum, pstrStyle As String)
    Dim lngIndex As Long
    
    UserControl.usrSlot(penSlot).ItemType = pstrStyle
    Select Case penSlot
        Case seMainHand
            lngIndex = SeekItem(pstrStyle)
            If lngIndex = 0 Then Exit Sub
            mblnTwoHand = db.Item(lngIndex).TwoHand
            HideOffhand
        Case seOffhand
            HideOffhand
    End Select
End Sub

Private Sub usrSlot_ItemTypeChange(Index As Integer, ItemType As String)
    RaiseEvent ItemTypeChange(Index, ItemType, mblnTwoHand)
    HideOffhand
End Sub

Private Sub HideOffhand()
    With UserControl.usrSlot(seOffhand)
        If mblnTwoHand Then
            If .ItemType <> "Empty" Then
                .ItemType = "Empty"
                RaiseEvent ItemTypeChange(seOffhand, "Empty", True)
            End If
        End If
        UserControl.usrSlot(seOffhand).Visible = Not mblnTwoHand
    End With
End Sub

Private Sub usrSlot_OffhandPairChange(Index As Integer, OffhandType As String)
    UserControl.usrSlot(seOffhand).ItemType = OffhandType
    RaiseEvent ItemTypeChange(seOffhand, OffhandType, False)
End Sub


' ************* NAMED ITEM / EFFECTS *************


Public Function GetNamedItem(penSlot As SlotEnum) As String
    GetNamedItem = UserControl.usrSlot(penSlot).NamedItem
End Function

Public Sub SetNamedItem(penSlot As SlotEnum, pstrNamedItem As String)
    UserControl.usrSlot(penSlot).NamedItem = pstrNamedItem
End Sub

Private Sub usrSlot_NamedItemChange(Index As Integer, NamedItem As String)
    RaiseEvent NamedItemChange(Index, NamedItem)
End Sub

Public Sub SetEffects(penSlot As SlotEnum, pstrPrefix As String, pstrSuffix As String, pstrExtra As String)
    UserControl.usrSlot(penSlot).SetEffects pstrPrefix, pstrSuffix, pstrExtra
End Sub


' ************* AUGMENT SLOTS *************


Public Sub SetAugmentSlots(penSlot As SlotEnum, pstrText As String)
    UserControl.usrSlot(penSlot).SetAugmentSlots pstrText
End Sub

Private Sub usrSlot_AugmentClick(Index As Integer, Left As Long, Right As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    
    With UserControl.usrSlot(Index)
        lngLeft = .Left + Left
        lngTop = .Top
        lngRight = .Left + Right
        lngBottom = .Top + .Height
    End With
    RaiseEvent AugmentClick(Index, lngLeft, lngTop, lngRight, lngBottom)
End Sub

Public Sub SimulateAugmentClick(penSlot As SlotEnum)
    UserControl.usrSlot(penSlot).SimulateAugmentClick
End Sub


' ************* ELDRITCH RITUAL *************


Public Function GetEldritchRitual(penSlot As SlotEnum) As Long
    GetEldritchRitual = UserControl.usrSlot(penSlot).Ritual
End Function

Public Sub SetEldritchRitual(penSlot As SlotEnum, plngRitual As Long)
    UserControl.usrSlot(penSlot).Ritual = plngRitual
End Sub

Private Sub usrSlot_EldritchChange(Index As Integer, Ritual As Long)
    RaiseEvent EldritchChange(Index, Ritual)
End Sub
