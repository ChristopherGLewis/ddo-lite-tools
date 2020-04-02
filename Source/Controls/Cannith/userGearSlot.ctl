VERSION 5.00
Begin VB.UserControl userGearSlot 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   384
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
   ScaleHeight     =   384
   ScaleWidth      =   7740
   Begin VB.PictureBox picSwaps 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   6900
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   384
   End
   Begin VB.PictureBox picEldritch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   6240
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   384
   End
   Begin VB.PictureBox picAugments 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   4800
      ScaleHeight     =   384
      ScaleWidth      =   1200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1200
   End
   Begin VB.PictureBox picGear 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   360
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   384
   End
   Begin CannithCrafting.userCheckBox usrchkCrafted 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   677
      Caption         =   ""
   End
   Begin VB.PictureBox picEffects 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   276
      Left            =   960
      ScaleHeight     =   252
      ScaleWidth      =   3576
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.TextBox txtNamed 
      Appearance      =   0  'Flat
      Height          =   276
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   3600
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Armor"
      Index           =   0
      Begin VB.Menu mnuArmor 
         Caption         =   "Full Plate"
         Index           =   0
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Half Plate"
         Index           =   1
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Breastplate"
         Index           =   3
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Scalemail"
         Index           =   4
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Hide"
         Index           =   5
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Chainmail"
         Index           =   7
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Leather"
         Index           =   8
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Robe"
         Index           =   10
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Outfit"
         Index           =   11
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuArmor 
         Caption         =   "Docent"
         Index           =   13
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "MainHand"
      Index           =   1
      Begin VB.Menu mnuMainHand 
         Caption         =   "Melee 1H"
         Index           =   0
         Begin VB.Menu mnuMelee1H 
            Caption         =   "Item"
            Index           =   0
         End
      End
      Begin VB.Menu mnuMainHand 
         Caption         =   "Melee 2H"
         Index           =   1
         Begin VB.Menu mnuMelee2H 
            Caption         =   "Item"
            Index           =   0
         End
      End
      Begin VB.Menu mnuMainHand 
         Caption         =   "Ranged"
         Index           =   2
         Begin VB.Menu mnuRange 
            Caption         =   "Item"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "OffHand"
      Index           =   2
      Begin VB.Menu mnuOffhand 
         Caption         =   "Melee"
         Index           =   0
         Begin VB.Menu mnuTWF 
            Caption         =   "Item"
            Index           =   0
         End
      End
      Begin VB.Menu mnuOffhand 
         Caption         =   "Shield"
         Index           =   1
         Begin VB.Menu mnuShield 
            Caption         =   "Tower Shield"
            Index           =   0
         End
         Begin VB.Menu mnuShield 
            Caption         =   "Heavy Shield"
            Index           =   1
         End
         Begin VB.Menu mnuShield 
            Caption         =   "Light Shield"
            Index           =   2
         End
         Begin VB.Menu mnuShield 
            Caption         =   "Buckler"
            Index           =   3
         End
      End
      Begin VB.Menu mnuOffhand 
         Caption         =   "Orb"
         Index           =   2
      End
      Begin VB.Menu mnuOffhand 
         Caption         =   "Runearm"
         Index           =   3
      End
      Begin VB.Menu mnuOffhand 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuOffhand 
         Caption         =   "Empty"
         Index           =   5
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Eldritch"
      Index           =   4
      Begin VB.Menu mnuEldritch 
         Caption         =   "None"
         Index           =   0
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Alchemical Armor"
         Index           =   2
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Alchemical Shield"
         Index           =   3
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Force Damage"
         Index           =   4
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Force Critical"
         Index           =   5
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Resistance"
         Index           =   6
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Adamantine Ritual I"
         Index           =   8
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Adamantine Ritual II"
         Index           =   9
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Adamantine Ritual III"
         Index           =   10
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Adamantine Ritual IV"
         Index           =   11
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "Adamantine Ritual V"
         Index           =   12
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuEldritch 
         Caption         =   "View All Recipes"
         Index           =   14
      End
   End
End
Attribute VB_Name = "userGearSlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ItemTypeChange(ItemType As String)
Public Event OffhandPairChange(OffhandType As String)
Public Event CraftedChange(Crafted As Boolean)
Public Event NamedItemChange(NamedItem As String)
Public Event AugmentSlotChange(Color As AugmentColorEnum, Value As Boolean)
Public Event AugmentClick(Left As Long, Right As Long)
Public Event EldritchChange(Ritual As Long)
Public Event Tooltip(TooltipText As String, Left As Long, Top As Long, Height As Long)

Private mstrItemType As String
Private mstrNamed As String
Private mtypAugment(1 To 7) As AugmentSlotType
Private mstrRitual As String

Private menSlot As SlotEnum
Private mstrPrefix As String
Private mstrSuffix As String
Private mstrExtra As String

Private mblnOverride As Boolean


' ************* USERCONTROL *************


Public Sub SetHeight()
    UserControl.Height = UserControl.picGear.Height
End Sub

Public Sub Init(penSlot As SlotEnum, pstrItemType As String, pstrNamed As String)
    menSlot = penSlot
    InitMenus
    mstrNamed = pstrNamed
    UserControl.picGear.TooltipText = GetSlotName(menSlot)
    DefaultItem pstrItemType
    UpdateStatus
    mstrRitual = vbNullString
    RefreshColors
    SetImages
End Sub

Private Sub UserControl_Resize()
    Dim lngMarginX As Long
    Dim lngWidth As Long
    
    With UserControl
        .Height = .picGear.Height
        lngMarginX = .ScaleX(12, vbPixels, vbTwips)
        ' Gear Icon
        .picGear.Left = .usrchkCrafted.BoxWidth + lngMarginX
        ' Textbox
        lngWidth = .ScaleWidth - .usrchkCrafted.BoxWidth - .picGear.Width - .picAugments.Width - .picEldritch.Width - .picSwaps.Width - (lngMarginX * 5)
        If lngWidth < .picGear.Width Then lngWidth = .picGear.Width
        .txtNamed.Left = .picGear.Left + .picGear.Width + lngMarginX
        .txtNamed.Width = lngWidth
        .txtNamed.Top = (.picGear.Height - .txtNamed.Height) \ 2
        With .txtNamed
            UserControl.picEffects.Move .Left, .Top, .Width, .Height
        End With
        ' Augments Slots
        .picAugments.Left = .txtNamed.Left + .txtNamed.Width + lngMarginX
        ' Eldritch Ritual
        .picEldritch.Left = .picAugments.Left + .picAugments.Width + lngMarginX
        ' Swaps
        .picSwaps.Left = .picEldritch.Left + .picEldritch.Width + lngMarginX
    End With
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearTip
End Sub

Private Sub ClearTip()
    RaiseEvent Tooltip(vbNullString, 0, 0, 0)
End Sub

Public Sub RefreshColors()
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .usrchkCrafted.RefreshColors cgeWorkspace
        cfg.ApplyColors .txtNamed, cgeControls
        cfg.ApplyColors .picEffects, cgeControls
    End With
End Sub

Private Sub DefaultItem(pstrItem As String)
    Dim strItem As String
    
    If Len(pstrItem) Then
        strItem = pstrItem
    Else
        Select Case menSlot
            Case seHelmet: strItem = "Helmet"
            Case seGoggles: strItem = "Goggles"
            Case seNecklace: strItem = "Necklace"
            Case seCloak: strItem = "Cloak"
            Case seBracers: strItem = "Bracers"
            Case seGloves: strItem = "Gloves"
            Case seBelt: strItem = "Belt"
            Case seBoots: strItem = "Boots"
            Case seRing1: strItem = "Ring"
            Case seRing2: strItem = "Ring"
            Case seTrinket: strItem = "Trinket"
            Case seArmor: strItem = "Full Plate"
            Case seMainHand: strItem = db.Melee1H.Default.Choice
            Case seOffHand: strItem = "Heavy Shield"
        End Select
    End If
    Me.ItemType = strItem
End Sub

Private Sub InitMenus()
    Select Case menSlot
        Case seMainHand
            InitMenu UserControl.mnuMelee1H, db.Melee1H
            InitMenu UserControl.mnuMelee2H, db.Melee2H
            InitMenu UserControl.mnuRange, db.Range
        Case seOffHand
            InitMenu UserControl.mnuTWF, db.Melee1H
    End Select
End Sub

Private Sub InitMenu(pobj As Object, ptypChoice As ChoiceType)
    Dim i As Long
    
    With ptypChoice
        For i = 1 To .Count
            Load pobj(i)
            pobj(i).Checked = False
            pobj(i).Enabled = True
            pobj(i).Caption = .List(i).Choice
            pobj(i).Tag = .List(i).OffhandPair
        Next
        If .Count Then pobj(0).Visible = False
    End With
End Sub


' ************* DRAWING *************


Private Sub SetImages()
    ShowItem
    ShowSlot UserControl.picSwaps, ssePaperDoll
    ShowAugments
    ShowEldritch
End Sub

Private Sub ShowItem()
    Dim strID As String
    
    strID = GetItemResource(mstrItemType)
    With UserControl.picGear
        .PaintPicture LoadResPicture(strID, vbResIcon), 0, 0, .Width, .Height
    End With
End Sub

Private Sub ShowSlot(ppic As PictureBox, penStyle As SlotStyleEnum)
    Dim strID As String
    
    strID = GetSlotResource(menSlot)
    ppic.PaintPicture LoadResPicture(strID, vbResBitmap), 0, 0, ppic.Width, ppic.Height
End Sub

Private Sub ShowEldritch()
    Dim strID As String
    
    With UserControl.picEldritch
        If Len(mstrRitual) Then strID = "MSCELDRITCH" Else strID = "MSCBARTER"
        .PaintPicture LoadResPicture(strID, vbResBitmap), 0, 0, .Width, .Height
    End With
End Sub

Public Sub GetCoords(plngColumn As Long, plngLeft As Long, plngRight As Long)
    With UserControl
        Select Case plngColumn
            Case 1
                plngLeft = .usrchkCrafted.Left
                plngRight = .picGear.Left + .picGear.Width
            Case 2
                plngLeft = .txtNamed.Left
                plngRight = .txtNamed.Left + .txtNamed.Width
            Case 3
                plngLeft = .picAugments.Left
                plngRight = .picAugments.Left + .picAugments.Width
            Case 4
                plngLeft = .picEldritch.Left
                plngRight = .picEldritch.Left + .picEldritch.Width
            Case 5
                plngLeft = .picSwaps.Left
                plngRight = .picSwaps.Left + .picSwaps.Width
        End Select
    End With
End Sub


' ************* CRAFTED *************


Public Property Get Crafted() As Boolean
    Crafted = UserControl.usrchkCrafted.Value
End Property

Public Property Let Crafted(ByVal pblnCrafted As Boolean)
    UserControl.usrchkCrafted.Value = pblnCrafted
    UpdateStatus
    SetTextboxFocus
End Property

Private Sub usrchkCrafted_UserChange()
    UpdateStatus
    SetTextboxFocus
    RaiseEvent CraftedChange(UserControl.usrchkCrafted.Value)
End Sub

' Update textbox, and also toggle crafted to false if switching to Empty
Private Sub UpdateStatus()
    Dim blnCrafted As Boolean
    Dim blnEmpty As Boolean
    Dim blnEnabled As Boolean
    Dim blnVisible As Boolean
    Dim enColor As ColorGroupEnum
    Dim strText As String
    
    mblnOverride = True
    blnEmpty = (mstrItemType = "Empty")
    blnCrafted = UserControl.usrchkCrafted.Value And Not blnEmpty
    blnEnabled = Not (blnCrafted Or blnEmpty)
    blnVisible = Not blnEmpty
    If blnEnabled And Len(mstrNamed) <> 0 Then enColor = cgeDropSlots Else enColor = cgeControls
    Select Case True
        Case blnEmpty: strText = vbNullString
        Case blnCrafted: CraftedText
        Case Else: strText = mstrNamed
    End Select
    With UserControl.usrchkCrafted
        If .Value <> blnCrafted Then
            .Value = blnCrafted
            RaiseEvent CraftedChange(blnCrafted)
        End If
        .Enabled = (Not blnEmpty)
    End With
    cfg.ApplyColors UserControl.txtNamed, enColor
    If Not blnCrafted Then
        UserControl.picEffects.Visible = False
        With UserControl.txtNamed
            .Text = strText
            .Enabled = blnEnabled
            .Visible = True
        End With
    End If
    With UserControl
        If .usrchkCrafted.Visible <> blnVisible Then
            .usrchkCrafted.Visible = blnVisible
            If Not blnVisible Then
                .txtNamed.Visible = False
                .picEffects.Visible = False
            End If
            .picAugments.Visible = blnVisible
            .picEldritch.Visible = blnVisible
            .picSwaps.Visible = blnVisible
        End If
    End With
    mblnOverride = False
End Sub

Private Sub CraftedText()
    Dim strExtra As String
    Dim strText As String
    
    With UserControl
        .txtNamed.Visible = False
        With .picEffects
            .Cls
            .CurrentX = .TextWidth(" ") \ 3
            .CurrentY = (.ScaleHeight - .TextHeight("Q")) \ 3
        End With
        If Len(mstrPrefix & mstrSuffix & mstrExtra) = 0 Then
            PrintEffect "cannith crafted"
        Else
            PrintEffect mstrPrefix
            PrintEffect "of "
            PrintEffect mstrSuffix
            If Len(mstrExtra) Then
                PrintEffect "w/"
                PrintEffect mstrExtra
            End If
        End If
        .picEffects.Visible = True
    End With
End Sub

Private Function PrintEffect(ByVal pstrText As String) As String
    Dim enColor As ColorValueEnum
    
    If Len(pstrText) = 0 Then pstrText = "(nothing) "
    If pstrText = "of " Or pstrText = "w/" Or pstrText = "(nothing) " Or pstrText = "cannith crafted" Then
        enColor = cveTextDim
    Else
        pstrText = pstrText & " "
        enColor = cveText
    End If
    UserControl.picEffects.ForeColor = cfg.GetColor(cgeControls, enColor)
    UserControl.picEffects.Print pstrText;
End Function

Private Sub picGear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    ClearTip
End Sub

Private Sub picGear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    If Button = vbLeftButton Then
        If Not UserControl.usrchkCrafted.Enabled Then Exit Sub
        Me.Crafted = Not Me.Crafted
        RaiseEvent CraftedChange(Me.Crafted)
    Else
        ShowItemTypeMenu
    End If
End Sub

Private Sub picGear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub picGear_DblClick()
    xp.SetMouseCursor mcHand
    If Not UserControl.usrchkCrafted.Enabled Then Exit Sub
    Me.Crafted = Not Me.Crafted
    RaiseEvent CraftedChange(Me.Crafted)
End Sub


' ************* ITEM TYPE *************


Public Property Get ItemType() As String
    ItemType = mstrItemType
End Property

Public Property Let ItemType(ByVal pstrItemType As String)
    mstrItemType = pstrItemType
    Select Case menSlot
        Case seArmor, seMainHand, seOffHand: UserControl.picGear.TooltipText = GetSlotName(menSlot) & ": " & mstrItemType
    End Select
    UpdateStatus
    ShowItem
End Property

Private Sub ShowItemTypeMenu()
    Dim i As Long
    
    ClearTip
    Select Case menSlot
        Case seArmor
            For i = 0 To UserControl.mnuArmor.UBound
                UserControl.mnuArmor(i).Checked = (UserControl.mnuArmor(i).Caption = mstrItemType)
            Next
            PopupMenu UserControl.mnuContext(0)
        Case seMainHand
            For i = 0 To UserControl.mnuMelee2H.UBound
                UserControl.mnuMelee2H(i).Checked = (UserControl.mnuMelee2H(i).Caption = mstrItemType)
            Next
            For i = 0 To UserControl.mnuMelee1H.UBound
                UserControl.mnuMelee1H(i).Checked = (UserControl.mnuMelee1H(i).Caption = mstrItemType)
            Next
            For i = 0 To UserControl.mnuRange.UBound
                UserControl.mnuRange(i).Checked = (UserControl.mnuRange(i).Caption = mstrItemType)
            Next
            PopupMenu UserControl.mnuContext(1)
        Case seOffHand
            For i = 0 To UserControl.mnuOffhand.UBound
                UserControl.mnuOffhand(i).Checked = (UserControl.mnuOffhand(i).Caption = mstrItemType)
            Next
            For i = 0 To UserControl.mnuTWF.UBound
                UserControl.mnuTWF(i).Checked = (UserControl.mnuTWF(i).Caption = mstrItemType)
            Next
            For i = 0 To UserControl.mnuShield.UBound
                UserControl.mnuShield(i).Checked = (UserControl.mnuShield(i).Caption = mstrItemType)
            Next
            PopupMenu UserControl.mnuContext(2)
    End Select
End Sub

Private Sub mnuMelee1H_Click(Index As Integer)
    ItemTypePairChange UserControl.mnuMelee1H(Index)
End Sub

Private Sub mnuMelee2H_Click(Index As Integer)
    ItemTypePairChange UserControl.mnuMelee2H(Index)
End Sub

Private Sub mnuRange_Click(Index As Integer)
    ItemTypePairChange UserControl.mnuRange(Index)
End Sub

Private Sub mnuArmor_Click(Index As Integer)
    ItemTypeUserChange UserControl.mnuArmor(Index).Caption
End Sub

Private Sub mnuShield_Click(Index As Integer)
    ItemTypeUserChange UserControl.mnuShield(Index).Caption
End Sub

Private Sub mnuTWF_Click(Index As Integer)
    ItemTypeUserChange UserControl.mnuTWF(Index).Caption
End Sub

' Opening submenus sends a click with the parent menu's index
' Ignore submenus (Melee, Shield) but handle actual choices (Orb, Runearm, Empty)
Private Sub mnuOffhand_Click(Index As Integer)
    Dim strCaption As String
    
    strCaption = UserControl.mnuOffhand(Index).Caption
    Select Case strCaption
        Case "Melee", "Shield": Exit Sub
    End Select
    ItemTypeUserChange strCaption
End Sub

Private Sub ItemTypeUserChange(pstrItemType As String)
    Me.ItemType = pstrItemType
    UpdateStatus
    If UserControl.usrchkCrafted.Value = False Then SetTextboxFocus
    RaiseEvent ItemTypeChange(mstrItemType)
End Sub

Private Sub ItemTypePairChange(pmnu As Menu)
    Dim lngIndex As Long
    
    ItemTypeUserChange pmnu.Caption
    If menSlot = seMainHand Then
        If Len(pmnu.Tag) Then
            RaiseEvent OffhandPairChange(pmnu.Tag)
        Else
            lngIndex = SeekItem(pmnu.Caption)
            If lngIndex Then
                If db.Item(lngIndex).TwoHand Then RaiseEvent OffhandPairChange("Empty")
            End If
        End If
    End If
End Sub


' ************* NAMED ITEM / EFFECTS *************


Public Property Get NamedItem() As String
    NamedItem = mstrNamed
End Property

Public Property Let NamedItem(ByVal pstrNamedItem As String)
    mstrNamed = pstrNamedItem
    If Not UserControl.usrchkCrafted.Value Then
        mblnOverride = True
        UserControl.txtNamed.Text = mstrNamed
        mblnOverride = False
    End If
    UpdateStatus
End Property

Private Sub SetTextboxFocus()
    If UserControl.txtNamed.Visible = False Or UserControl.txtNamed.Enabled = False Then Exit Sub
    On Error Resume Next
    UserControl.txtNamed.SetFocus
End Sub

Private Sub txtNamed_GotFocus()
    cfg.ApplyColors UserControl.txtNamed, cgeDropSlots
    With UserControl.txtNamed
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNamed_LostFocus()
    Dim enColor As ColorGroupEnum
    
    If UserControl.usrchkCrafted.Value = False And Len(UserControl.txtNamed.Text) > 0 Then enColor = cgeDropSlots Else enColor = cgeControls
    cfg.ApplyColors UserControl.txtNamed, enColor
End Sub

Private Sub txtNamed_Change()
    If mblnOverride Then Exit Sub
    mstrNamed = UserControl.txtNamed.Text
    RaiseEvent NamedItemChange(mstrNamed)
End Sub

Private Sub txtNamed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearTip
End Sub

Private Sub picEffects_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearTip
End Sub

Public Sub SetEffects(pstrPrefix As String, pstrSuffix As String, pstrExtra As String)
    mstrPrefix = pstrPrefix
    mstrSuffix = pstrSuffix
    mstrExtra = pstrExtra
    UpdateStatus
End Sub


' ************* AUGMENTS *************


Public Sub SetAugmentSlots(pstrText As String)
    StringToGearsetAugment mtypAugment, pstrText
    ShowAugments
End Sub

Private Sub picAugments_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    AugmentTooltip X
End Sub

Private Sub AugmentTooltip(X As Single)
    Dim lngWidth As Long
    Dim enColor As AugmentColorEnum
    Dim lngSlot As Long
    Dim lngCount As Long
    Dim strTip As String
    
    lngWidth = UserControl.picGear.Width
    lngSlot = X \ lngWidth + 1
    For enColor = 1 To 7
        If mtypAugment(enColor).Exists Then
            lngCount = lngCount + 1
            If lngCount = lngSlot Then Exit For
        End If
    Next
    If lngCount <> lngSlot Then
        ClearTip
    Else
        With mtypAugment(enColor)
            If .Augment = 0 Or .Variation = 0 Then
                strTip = "Empty " & GetAugmentColorName(enColor) & " Slot"
            Else
                strTip = db.Augment(.Augment).Variation(.Variation)
            End If
        End With
        With UserControl.picAugments
            RaiseEvent Tooltip(strTip, .Left + (lngSlot * lngWidth), .Top, .Height)
        End With
    End If
End Sub

Private Sub picAugments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SimulateAugmentClick
End Sub

Public Sub SimulateAugmentClick()
    Dim lngLeft As Long
    Dim lngRight As Long
    
    xp.SetMouseCursor mcHand
    With UserControl.picAugments
        lngLeft = .Left
        lngRight = .Left + .Width
    End With
    RaiseEvent AugmentClick(lngLeft, lngRight)
End Sub

Private Sub picAugments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub ShowAugments()
    Dim strID As String
    Dim lngCount As Long
    Dim i As Long
    
    With UserControl.picAugments
        strID = "AUGBACK"
        .PaintPicture LoadResPicture(strID, vbResBitmap), 0, 0, .Width, .Height
    End With
    For i = 1 To 7
        If mtypAugment(i).Exists Then
            lngCount = lngCount + 1
            DrawAugment i, lngCount
        End If
    Next
End Sub

Private Sub DrawAugment(penColor As AugmentColorEnum, plngSlot As Long)
    Dim strID As String
    Dim strColor As String
    Dim lngLeft As Long
    Dim lngTotalWidth As Long
    Dim lngWidth As Long
    
    If plngSlot < 1 Or plngSlot > 3 Then Exit Sub
    lngTotalWidth = UserControl.picAugments.Width
    lngWidth = UserControl.picGear.Width
    Select Case plngSlot
        Case 1: lngLeft = 0
        Case 2: lngLeft = lngWidth + (lngTotalWidth - (lngWidth * 3)) \ 2
        Case 3: lngLeft = lngTotalWidth - lngWidth
    End Select
    strColor = GetAugmentColorName(penColor, True)
    With UserControl.picAugments
        strID = "AUG" & UCase$(strColor)
        .PaintPicture LoadResPicture(strID, vbResIcon), lngLeft, 0, lngWidth, .Height
    End With
End Sub


' ************* ELDRITCH *************


Private Sub picEldritch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    With UserControl.picEldritch
        RaiseEvent Tooltip(mstrRitual, .Left + .Width, .Top, .Height)
    End With
End Sub

Private Sub picEldritch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    ShowEldritchMenu
End Sub

Private Sub picEldritch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub ShowEldritchMenu()
    Dim strName As String
    Dim lngRitual As Long
    Dim lngItem As Long
    Dim lngType As Long
    Dim lngStyle As Long
    Dim blnVisible As Boolean
    Dim blnSeparator As Boolean
    Dim lngEnum As Long
    Dim i As Long
    
    ClearTip
    With UserControl.mnuEldritch(0)
        strName = mstrItemType
        If strName = "Ring" Then strName = GetSlotName(menSlot)
        If Len(mstrRitual) = 0 Then
            .Caption = "Ritual for " & strName
            .Enabled = False
        Else
            .Caption = "Clear Ritual for " & strName
            .Enabled = True
        End If
    End With
    lngItem = SeekItem(mstrItemType)
    If lngItem Then
        With db.Item(lngItem)
            lngType = .ItemType
            lngStyle = .ItemStyle
        End With
    End If
    For i = 1 To UserControl.mnuEldritch.UBound
        With UserControl.mnuEldritch(i)
            lngRitual = SeekRitual(.Caption)
            If lngRitual <> 0 Then
                With db.Ritual(lngRitual)
                    If .ItemType(0) = True And lngType <> 0 Then
                        blnVisible = .ItemType(lngType)
                    ElseIf .ItemStyle(0) = True And lngStyle <> 0 Then
                        blnVisible = .ItemStyle(lngStyle)
                    Else
                        blnVisible = True
                    End If
                End With
                .Visible = blnVisible
                If blnVisible And Left$(.Caption, 17) <> "Adamantine Ritual" Then blnSeparator = True
            End If
            .Checked = (.Caption = mstrRitual)
        End With
    Next
    UserControl.mnuEldritch(1).Visible = blnSeparator
    PopupMenu UserControl.mnuContext(4)
End Sub

Private Sub mnuEldritch_Click(Index As Integer)
    Select Case Index
        Case 0: Ritual = 0
        Case 2 To 12: Ritual = SeekRitual(UserControl.mnuEldritch(Index).Caption)
        Case 14: OpenForm "frmEldritch"
    End Select
End Sub

Public Property Get Ritual() As Long
    Ritual = SeekRitual(mstrRitual)
End Property

Public Property Let Ritual(ByVal plngRitual As Long)
    Dim strRitual As String
    
    If plngRitual Then strRitual = db.Ritual(plngRitual).RitualName Else strRitual = vbNullString
    If mstrRitual = strRitual Then Exit Property
    mstrRitual = strRitual
    ShowEldritch
    RaiseEvent EldritchChange(plngRitual)
End Property


' ************* SWAPS *************


Private Sub picSwaps_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcNo
End Sub

Private Sub picSwaps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcNo
End Sub

Private Sub picSwaps_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcNo
End Sub

Private Sub picSwaps_DblClick()
    xp.SetMouseCursor mcNo
End Sub

