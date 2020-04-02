VERSION 5.00
Begin VB.UserControl userIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   408
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   408
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   408
   ScaleWidth      =   408
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Left            =   0
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   408
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Index           =   0
      Begin VB.Menu mnuItem 
         Caption         =   "Item Style"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuItem 
         Caption         =   "-"
         Index           =   1
      End
   End
End
Attribute VB_Name = "userIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum UserIconStyleEnum
    uiseToggle
    uiseLink
End Enum

Public Event Click()
Public Event ActiveChange(pblnActive As Boolean)
Public Event StyleChange(pstrStyle As String)

Private menGear As GearEnum
Private mlngLastItem As Long
Private mlngSelected As Long
Private mblnActive As Boolean
Private mblnAllowHandwraps As Boolean

Private mblnInitialized As Boolean
Private mblnAllowMenu As Boolean
Private menStyle As UserIconStyleEnum


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mblnAllowMenu = True
    UserControl.pic.Enabled = True
    menStyle = uiseToggle
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AllowMenu", mblnAllowMenu, True
    PropBag.WriteProperty "Enabled", UserControl.pic.Enabled, True
    PropBag.WriteProperty "Style", menStyle, uiseToggle
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mblnAllowMenu = PropBag.ReadProperty("AllowMenu", True)
    UserControl.pic.Enabled = PropBag.ReadProperty("Enabled", True)
    menStyle = PropBag.ReadProperty("Style", uiseToggle)
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .pic.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub

Public Sub RefreshColors()
    With UserControl
        .BackColor = cfg.GetColor(cgeControls, cveBackground)
        .Cls
        .pic.BackColor = cfg.GetColor(cgeControls, cveBackground)
        .Cls
    End With
    DrawIcon
End Sub

Public Sub Init(penGear As GearEnum, penStyle As UserIconStyleEnum, pblnAllowHandwraps As Boolean)
    UserControl.pic.BackColor = RGB(30, 30, 30)
    mblnActive = False
    menGear = penGear
    menStyle = penStyle
    mblnAllowHandwraps = pblnAllowHandwraps
    Select Case menGear
        Case geHelmet: InitItems "Helmet"
        Case geGoggles: InitItems "Goggles"
        Case geNecklace: InitItems "Necklace"
        Case geCloak: InitItems "Cloak"
        Case geBracers: InitItems "Bracers"
        Case geGloves: InitItems "Gloves"
        Case geBelt: InitItems "Belt"
        Case geBoots: InitItems "Boots"
        Case geRing: InitItems "Ring"
        Case geTrinket: InitItems "Trinket"
        Case ge2hMelee: InitItemsList "Two-Hand Weapon", db.Melee2H
        Case ge1hMelee: InitItemsList "One-Hand Weapon", db.Melee1H
        Case geHandwraps: InitItems "Handwraps"
        Case geRange: InitItemsList "Range Weapon", db.Range
        Case geShield: InitItems "Shields,Tower Shield,Heavy Shield,Light Shield,Buckler", "Heavy Shield", "Shield"
        Case geOrb: InitItems "Orb"
        Case geRunearm: InitItems "Runearm"
        Case geMetalArmor: InitItems "Metal Armor,Full Plate,Half Plate,-,Breastplate,Scalemail,-,Chainmail", "Full Plate"
        Case geLeatherArmor: InitItems "Natural Armor,Hide,Leather", "Leather", "Leather or Hide Armor"
        Case geClothArmor: InitItems "Cloth Armor,Outfit,Robe", "Outfit"
        Case geDocent: InitItems "Docent"
        Case Else
    End Select
    mblnInitialized = True
    DrawIcon
End Sub

Private Sub InitItems(pstrItemList As String, Optional pstrCurrent As String, Optional ByVal pstrTooltip As String = vbNullString)
    Dim strSplit() As String
    Dim i As Long

    mlngSelected = 0
    strSplit = Split(pstrItemList, ",")
    SetMenuItem strSplit(0), 0
    If menStyle = uiseLink Then
        pstrTooltip = vbNullString
    ElseIf Len(pstrTooltip) = 0 Then
        pstrTooltip = strSplit(0)
    End If
    UserControl.pic.TooltipText = pstrTooltip
    For i = 1 To UBound(strSplit)
        SetMenuItem strSplit(i), i + 1
        If strSplit(i) = pstrCurrent Then mlngSelected = i + 1
    Next
End Sub

Private Sub InitItemsList(pstrGroup As String, ptypChoice As ChoiceType)
    Dim i As Long
    
    mlngSelected = 0
    SetMenuItem pstrGroup & "s", 0
    UserControl.pic.TooltipText = pstrGroup
    For i = 1 To ptypChoice.Count
        With ptypChoice
            SetMenuItem .List(i).Choice, i + 1
            If .List(i).Choice = .Default.Choice Then mlngSelected = i + 1
        End With
    Next
End Sub

Private Sub SetMenuItem(pstrItem As String, plngIndex As Long)
    Dim strResourceID As String
    Dim lngItem As Long
    
    If pstrItem = "Handwraps" And Not mblnAllowHandwraps Then Exit Sub
    lngItem = SeekItem(pstrItem)
    If lngItem Then strResourceID = db.Item(lngItem).ResourceID Else strResourceID = "UNKNOWN"
    If plngIndex > UserControl.mnuItem.UBound Then Load UserControl.mnuItem(plngIndex)
    With UserControl.mnuItem(plngIndex)
        .Enabled = (plngIndex > 0)
        .Checked = False
        .Caption = pstrItem
        .Tag = strResourceID
        .Visible = True
    End With
    mlngLastItem = plngIndex
End Sub

Private Sub DrawIcon()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngPixelX As Long
    Dim lngPixelY As Long
    Dim strResourceID As String
    Dim lngColor As Long
    
    If Not mblnInitialized Then Exit Sub
    strResourceID = UserControl.mnuItem(mlngSelected).Tag
    lngPixelX = Screen.TwipsPerPixelX
    lngPixelY = Screen.TwipsPerPixelY
    With UserControl
        .pic.Move 0, 0, .ScaleWidth, .ScaleHeight
        With .pic
            lngWidth = .ScaleWidth
            lngHeight = .ScaleHeight
            .Cls
            On Error Resume Next
            .PaintPicture LoadResPicture(strResourceID, vbResIcon), lngPixelX, lngPixelY, lngWidth - (lngPixelX * 2), lngHeight - (lngPixelY * 2)
            If Err.Number Then .PaintPicture LoadResPicture("UNKNOWN", vbResIcon), lngPixelX, lngPixelY, lngWidth - (lngPixelX * 2), lngHeight - (lngPixelY * 2)
            On Error GoTo 0
        End With
    End With
    If Not Me.Enabled Then
        GrayScale UserControl.pic
        Exit Sub
    End If
    If Not Me.Active Then Exit Sub
    lngColor = RGB(60, 160, 60)
    ' Top and Bottom
    lngLeft = lngPixelX * 2
    lngRight = lngWidth - lngPixelX * 2
    lngTop = 0
    lngBottom = lngHeight - lngPixelY
    UserControl.pic.Line (lngLeft, lngTop)-(lngRight, lngTop), lngColor
    UserControl.pic.Line (lngLeft, lngBottom)-(lngRight, lngBottom), lngColor
    ' Left and Right
    lngLeft = 0
    lngRight = lngWidth - lngPixelX
    lngTop = lngPixelY * 2
    lngBottom = lngHeight - lngPixelY * 2
    UserControl.pic.Line (lngLeft, lngTop)-(lngLeft, lngBottom), lngColor
    UserControl.pic.Line (lngRight, lngTop)-(lngRight, lngBottom), lngColor
    ' Draw boxes
    lngLeft = lngPixelX
    lngRight = lngWidth - lngPixelX * 2
    lngTop = lngPixelX
    lngBottom = lngHeight - lngPixelY * 2
    DrawBox lngLeft, lngTop, lngRight, lngBottom, RGB(165, 240, 165), lngPixelX, lngPixelY
    DrawBox lngLeft, lngTop, lngRight, lngBottom, RGB(250, 255, 250), lngPixelX, lngPixelY
    DrawBox lngLeft, lngTop, lngRight, lngBottom, RGB(160, 225, 160), lngPixelX, lngPixelY
    DrawBox lngLeft, lngTop, lngRight, lngBottom, RGB(70, 185, 90), lngPixelX, lngPixelY
End Sub

Private Sub DrawBox(plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long, plngColor As Long, plngPixelX As Long, plngPixelY As Long)
    UserControl.pic.Line (plngLeft, plngTop)-(plngRight, plngBottom), plngColor, B
    ' Shrink next box one pixel on all sides
    plngLeft = plngLeft + plngPixelX
    plngTop = plngTop + plngPixelY
    plngRight = plngRight - plngPixelX
    plngBottom = plngBottom - plngPixelY
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    If Button = vbRightButton And mlngLastItem And mblnAllowMenu Then
        Active = True
        ShowPopup
    ElseIf menStyle = uiseToggle Then
        Active = Not Active
    Else
        RaiseEvent Click
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    If menStyle = uiseLink And Not Me.Active Then Me.Active = True
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub pic_DblClick()
    xp.SetMouseCursor mcHand
    Select Case menStyle
        Case uiseToggle: Me.Active = Not Me.Active
        Case uiseLink: RaiseEvent Click
    End Select
End Sub

Private Sub ShowPopup()
    Dim i As Long
    
    With UserControl
        For i = 2 To mlngLastItem
            .mnuItem(i).Checked = (mlngSelected = i)
        Next
    End With
    PopupMenu UserControl.mnuContext(0)
End Sub

Private Sub mnuItem_Click(Index As Integer)
    mlngSelected = Index
    DrawIcon
    RaiseEvent StyleChange(UserControl.mnuItem(Index).Caption)
End Sub

Public Property Get Active() As Boolean
    Active = mblnActive
End Property

Public Property Let Active(ByVal pblnActive As Boolean)
    If mblnActive = pblnActive Then Exit Property
    mblnActive = pblnActive
    RaiseEvent ActiveChange(mblnActive)
    DrawIcon
End Property

Public Property Get IconName() As String
    IconName = UserControl.mnuItem(mlngSelected).Caption
End Property

Public Property Let IconName(ByVal pstrIcon As String)
    Dim i As Long
    
    If Len(pstrIcon) = 0 Then Exit Property
    With UserControl
        For mlngSelected = 2 To .mnuItem.UBound
            If .mnuItem(mlngSelected).Caption = pstrIcon Then Exit For
        Next
        If mlngSelected > .mnuItem.UBound Then mlngSelected = 0
    End With
    DrawIcon
End Property

Public Property Get AllowMenu() As Boolean
    AllowMenu = mblnAllowMenu
End Property

Public Property Let AllowMenu(ByVal pblnAllowMenu As Boolean)
    mblnAllowMenu = pblnAllowMenu
    PropertyChanged "AllowMenu"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.pic.Enabled
End Property

Public Property Let Enabled(ByVal pblnEnabled As Boolean)
    UserControl.pic.Enabled = pblnEnabled
    PropertyChanged "Enabled"
    DrawIcon
End Property

Public Property Get Style() As UserIconStyleEnum
    Style = menStyle
End Property

Public Property Let Style(ByVal penStyle As UserIconStyleEnum)
    menStyle = penStyle
    PropertyChanged "Style"
End Property
