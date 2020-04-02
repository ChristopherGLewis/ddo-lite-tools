VERSION 5.00
Begin VB.UserControl userAugment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   4308
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2808
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4308
   ScaleWidth      =   2808
   Begin VB.ListBox lstAugment 
      Appearance      =   0  'Flat
      Height          =   3696
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2808
   End
   Begin CannithCrafting.userTab usrTabLow 
      Height          =   312
      Left            =   0
      TabIndex        =   1
      Top             =   300
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   550
      Captions        =   "All,Colorless,Yellow"
   End
   Begin CannithCrafting.userTab usrTabHigh 
      Height          =   312
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2652
      _ExtentX        =   4678
      _ExtentY        =   550
      Captions        =   "Blue,Red,Named"
      ActiveTab       =   0
   End
End
Attribute VB_Name = "userAugment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event AugmentSelected(Augment As Long, Variation As Long, ML As Long)
Public Event AugmentSlotted(Augment As Long, Variation As Long, GearSlot As SlotEnum, AugmentSlot As AugmentColorEnum)

Private Enum LevelEnum
    leTop
    leVariants
    leScale
End Enum

Private Enum DirectionEnum
    deDown
    deUp
End Enum

Private Type ListboxStateType
    TopIndex As Long
    Selected As Long
End Type

Private menSlotColor As AugmentColorEnum
Private menGearSlot As SlotEnum
Private mblnVariantOnly As Boolean

Private menTab As AugmentColorEnum ' Any = All, Orange = Named
Private menLevel As LevelEnum
Private mtypState(2) As ListboxStateType

Private mlngAugment As Long
Private mlngVariant As Long
Private mlngScale As Long

Private mblnInitialized As Boolean
Private mblnOverride As Boolean

Private mlngHighTop As Long
Private mlngLowTop As Long
Private mlngListOffset As Long


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    menSlotColor = aceAny
    mblnVariantOnly = False
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "SlotColor", menSlotColor, aceAny
    PropBag.WriteProperty "VariantOnly", mblnVariantOnly, False
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    menSlotColor = PropBag.ReadProperty("SlotColor", aceAny)
    mblnVariantOnly = PropBag.ReadProperty("VariantOnly", False)
End Sub


' ************* USERCONTROL *************


Private Sub UserControl_Resize()
    If mblnOverride Then Exit Sub
    With UserControl
        .lstAugment.Width = .ScaleWidth
        .lstAugment.Height = .ScaleHeight - .lstAugment.Top
        mblnOverride = True
        .Height = .lstAugment.Top + .lstAugment.Height
        mblnOverride = False
    End With
End Sub

Public Sub Init(penSlotColor As AugmentColorEnum)
    RefreshColors
    With UserControl
        mlngHighTop = .usrTabHigh.Top
        mlngLowTop = .usrTabLow.Top
        mlngListOffset = .lstAugment.Top - mlngLowTop
    End With
    menTab = aceAny
    Me.SlotColor = penSlotColor
    ShowTopLevel
End Sub

Public Sub RefreshColors()
    Dim i As Long
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        SetTabColors .usrTabLow
        SetTabColors .usrTabHigh
        cfg.ApplyColors .lstAugment, cgeControls
    End With
End Sub

Public Sub ClearSelection()
    mblnOverride = True
    UserControl.lstAugment.ListIndex = -1
    mblnOverride = False
End Sub

Private Sub SetTabColors(ptab As userTab)
    ptab.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    ptab.TabInactiveColor = cfg.GetColor(cgeControls, cveBackground)
    ptab.TabActiveColor = cfg.GetColor(cgeDropSlots, cveBackground)
    ptab.TextActiveColor = cfg.GetColor(cgeDropSlots, cveText)
    ptab.TextInactiveColor = cfg.GetColor(cgeControls, cveText)
End Sub

Public Sub SetSelected(plngAugment As Long, plngVariant As Long, plngScale As Long)
    Dim strTab As String
    
    Erase mtypState
    mlngAugment = plngAugment
    mlngVariant = plngVariant
    mlngScale = plngScale
    mblnOverride = True
    If mlngAugment Then
        If db.Augment(mlngAugment).Named Then
            menTab = aceOrange
            strTab = "Named"
        Else
            menTab = db.Augment(mlngAugment).Color
            Select Case menTab
                Case aceOrange, acePurple, aceGreen: strTab = "Named"
                Case Else: strTab = GetAugmentColorName(menTab)
            End Select
        End If
        UserControl.usrTabHigh.ActiveTab = strTab
        UserControl.usrTabLow.ActiveTab = strTab
        If db.Augment(mlngAugment).Scalings > 1 And Not mblnVariantOnly Then
            menLevel = leScale
            ShowScaling
            HighlightChoice mlngScale
        ElseIf db.Augment(mlngAugment).Variations > 1 Then
            menLevel = leVariants
            ShowVariants
            HighlightChoice mlngVariant
        Else
            menLevel = leTop
            ShowTopLevel False
            HighlightChoice mlngAugment
        End If
    End If
    mblnOverride = False
End Sub

Private Sub HighlightChoice(plngItemData As Long)
    mblnOverride = True
    ListboxSetValue UserControl.lstAugment, plngItemData
End Sub

Public Property Get SlotColor() As AugmentColorEnum
    SlotColor = menSlotColor
End Property

Public Property Let SlotColor(ByVal SlotColor As AugmentColorEnum)
    menSlotColor = SlotColor
    FilterTabs
    menTab = aceAny
    ShowTopLevel
End Property

Public Property Let GearSlot(GearSlot As SlotEnum)
    menGearSlot = GearSlot
End Property

Public Property Get VariantOnly() As Boolean
    VariantOnly = mblnVariantOnly
End Property

Public Property Let VariantOnly(pblnVariantOnly As Boolean)
    mblnVariantOnly = pblnVariantOnly
    PropertyChanged "VariantOnly"
End Property

Private Sub FilterTabs()
    Dim strLow As String
    Dim strHigh As String
    Dim lngTabs As Long
    
    lngTabs = 2
    Select Case menSlotColor
        Case aceAny
            strHigh = "Blue,Red,Named"
            strLow = "All,Colorless,Yellow"
        Case aceRed
            strHigh = "Named"
            strLow = "All,Colorless,Red"
        Case aceOrange
            strHigh = "Red,Named"
            strLow = "All,Colorless,Yellow"
        Case acePurple
            strHigh = "Red,Named"
            strLow = "All,Colorless,Blue"
        Case aceBlue
            strHigh = "Named"
            strLow = "All,Colorless,Blue"
        Case aceGreen
            strHigh = "Blue,Named"
            strLow = "All,Colorless,Yellow"
        Case aceYellow
            strHigh = "Named"
            strLow = "All,Colorless,Yellow"
        Case aceColorless
            strHigh = ""
            strLow = "Colorless,Named"
            lngTabs = 1
    End Select
    mblnOverride = True
    With UserControl
        .usrTabLow.Captions = strLow
        .usrTabLow.ActiveTab = "All"
        Select Case lngTabs
            Case 1
                .usrTabHigh.Visible = False
                .usrTabLow.Top = mlngHighTop
            Case 2
                .usrTabLow.Top = mlngLowTop
                .usrTabHigh.Captions = strHigh
                .usrTabHigh.ActiveTab = vbNullString
                .usrTabHigh.Visible = True
            End Select
            .lstAugment.Top = .usrTabLow.Top + mlngListOffset
            .lstAugment.Height = .ScaleHeight - .lstAugment.Top
    End With
    mblnOverride = False
End Sub


' ************* TABS *************


Private Sub usrTabLow_Click(pstrCaption As String)
    If mblnOverride Then Exit Sub
    mblnOverride = True
    UserControl.usrTabHigh.ActiveTab = vbNullString
    mblnOverride = False
    ChangeTab pstrCaption
End Sub

Private Sub usrTabHigh_Click(pstrCaption As String)
    If mblnOverride Then Exit Sub
    mblnOverride = True
    UserControl.usrTabLow.ActiveTab = vbNullString
    mblnOverride = False
    ChangeTab pstrCaption
End Sub

Private Sub ChangeTab(pstrCaption As String)
    Erase mtypState
    Select Case pstrCaption
        Case "All": menTab = aceAny
        Case "Colorless": menTab = aceColorless
        Case "Yellow": menTab = aceYellow
        Case "Blue": menTab = aceBlue
        Case "Red": menTab = aceRed
        Case "Named": menTab = aceOrange
    End Select
    ShowTopLevel
End Sub


' ************* LIST *************


Private Sub ShowTopLevel(Optional pblnReset As Boolean = True)
    Dim blnAdd As Boolean
    Dim enColor As AugmentColorEnum
    Dim i As Long
    
    If pblnReset Then
        mlngAugment = 0
        mlngVariant = 0
        mlngScale = 0
    End If
    mblnOverride = True
    menLevel = leTop
    ListboxClear UserControl.lstAugment
    blnAdd = True
    For i = 1 To db.Augments
        enColor = db.Augment(i).Color
        blnAdd = False
        ' Filer based on slot color
        If menSlotColor = aceAny Or enColor = menSlotColor Or enColor = aceColorless Then
            blnAdd = True
        ElseIf menSlotColor = aceGreen Then
            If enColor = aceYellow Or enColor = aceBlue Then blnAdd = True
        ElseIf menSlotColor = aceOrange Then
            If enColor = aceYellow Or enColor = aceRed Then blnAdd = True
        ElseIf menSlotColor = acePurple Then
            If enColor = aceRed Or enColor = aceBlue Then blnAdd = True
        End If
        ' Filter based on active tab
        If blnAdd Then
            Select Case menTab
                Case aceAny: blnAdd = True
                Case aceOrange: blnAdd = db.Augment(i).Named
                Case Else: blnAdd = (enColor = menTab)
            End Select
        End If
        If blnAdd Then ListboxAddItem UserControl.lstAugment, db.Augment(i).AugmentName, i
    Next
    mblnOverride = False
End Sub

Private Sub lstAugment_Click()
    Dim lngItemData As Long
    Dim strItem As String
    
    If UserControl.lstAugment.ListIndex = -1 Or mblnOverride = True Then Exit Sub
    lngItemData = ListboxGetValue(UserControl.lstAugment)
    With UserControl.lstAugment
        strItem = .Text
        mtypState(menLevel).TopIndex = .TopIndex
        mtypState(menLevel).Selected = .ListIndex
    End With
    Select Case menLevel
        Case leTop
            mlngAugment = lngItemData
            mlngVariant = 0
            mlngScale = 0
            ShowVariants
        Case leVariants
            If lngItemData = 0 Then
                ShowTopLevel
                MoveFocus
            Else
                mlngVariant = lngItemData
                ShowScaling
            End If
        Case leScale
            If lngItemData = 0 Then
                If db.Augment(mlngAugment).Variations = 1 Then ShowTopLevel Else ShowVariants
                MoveFocus
            Else
                mlngScale = lngItemData
                RaiseEvent AugmentSelected(mlngAugment, mlngVariant, mlngScale)
            End If
    End Select
End Sub

Private Sub MoveFocus()
    mblnOverride = True
    With UserControl.lstAugment
        .TopIndex = mtypState(menLevel).TopIndex
        .ListIndex = mtypState(menLevel).Selected
        .ListIndex = -1
    End With
    mblnOverride = False
End Sub

Private Sub ShowVariants()
    Dim i As Long
    
    If db.Augment(mlngAugment).Variations = 1 Then
        mlngVariant = 1
        mlngScale = 0
        ShowScaling
        Exit Sub
    End If
    mblnOverride = True
    menLevel = leVariants
    ListboxClear UserControl.lstAugment
    UserControl.lstAugment.AddItem "<Up One Level>", 0
    With db.Augment(mlngAugment)
        For i = 1 To .Variations
            If db.Augment(mlngAugment).Scalings = 1 Then
                ListboxAddItem UserControl.lstAugment, AugmentFullName(mlngAugment, i, 1), i
            Else
                ListboxAddItem UserControl.lstAugment, .Variation(i), i
            End If
        Next
    End With
    mblnOverride = False
End Sub

Private Sub ShowScaling()
    Dim strText As String
    Dim i As Long
    
    If db.Augment(mlngAugment).Scalings = 1 Or mblnVariantOnly = True Then
        mlngScale = 1
        RaiseEvent AugmentSelected(mlngAugment, mlngVariant, mlngScale)
        RaiseEvent AugmentSlotted(mlngAugment, mlngVariant, menGearSlot, menSlotColor)
        Exit Sub
    End If
    mblnOverride = True
    menLevel = leScale
    ListboxClear UserControl.lstAugment
    UserControl.lstAugment.AddItem "<Up One Level>"
    With db.Augment(mlngAugment)
        For i = 1 To .Scalings
            ListboxAddItem UserControl.lstAugment, CreateCaption(i), i
        Next
    End With
    mblnOverride = False
End Sub

Private Function CreateCaption(plngScale As Long) As String
    Dim strText As String
    Dim strValue As String
    Dim lngTotalWidth As Long
    Dim lngValueWidth As Long
    Dim lngWidth As Long
    
    With db.Augment(mlngAugment)
        With .Scaling(plngScale)
            strText = "ML" & .ML & ": "
            If Len(.Prefix(mlngVariant)) Then strText = strText & Trim$(.Prefix(mlngVariant)) & " "
        End With
        strText = strText & .Variation(mlngVariant)
        If Not .PrefixNotValue Then strValue = .Scaling(plngScale).Value
    End With
    With UserControl
        lngTotalWidth = .lstAugment.Width - .TextWidth("  ")
        If Len(strText) Then lngValueWidth = .TextWidth(" " & strValue)
        lngWidth = lngTotalWidth - lngValueWidth
        If .TextWidth(strText) <= lngWidth Then
            strValue = " " & strValue
        Else
            lngWidth = lngWidth + .TextWidth(" ")
            Do While .TextWidth(strText & "...") > lngWidth
                strText = Trim$(Left$(strText, Len(strText) - 1))
            Loop
            strText = Trim$(strText) & "..."
        End If
    End With
    strText = strText & strValue
    CreateCaption = strText
End Function
