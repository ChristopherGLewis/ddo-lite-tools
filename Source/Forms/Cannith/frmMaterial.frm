VERSION 5.00
Begin VB.Form frmMaterial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials"
   ClientHeight    =   7740
   ClientLeft      =   36
   ClientTop       =   384
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
   Icon            =   "frmMaterial.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   13548
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userInfo usrinfoShards 
      Height          =   6312
      Left            =   10320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   2952
      _ExtentX        =   5207
      _ExtentY        =   11134
      TitleSize       =   1
      TitleIcon       =   0   'False
      TitleText       =   "Used in recipes for:"
   End
   Begin VB.Frame fraCommands 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1212
      Left            =   300
      TabIndex        =   4
      Top             =   6000
      Width           =   3672
      Begin CannithCrafting.userCheckBox usrchkType 
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1872
         _ExtentX        =   3302
         _ExtentY        =   445
         Caption         =   "Collectables"
      End
      Begin CannithCrafting.userCheckBox usrchkType 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   900
         Width           =   1872
         _ExtentX        =   3302
         _ExtentY        =   445
         Value           =   0   'False
         Caption         =   "All Materials"
      End
      Begin CannithCrafting.userCheckBox usrchkType 
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1872
         _ExtentX        =   3302
         _ExtentY        =   445
         Value           =   0   'False
         Caption         =   "Miscellaneous"
      End
      Begin CannithCrafting.userCheckBox usrchkType 
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1872
         _ExtentX        =   3302
         _ExtentY        =   445
         Value           =   0   'False
         Caption         =   "Soul Gems"
      End
      Begin VB.Label lnkSchool 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Natural"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   216
         Index           =   4
         Left            =   2556
         TabIndex        =   12
         Top             =   900
         Width           =   732
      End
      Begin VB.Label lnkSchool 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Lore"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   216
         Index           =   3
         Left            =   2688
         TabIndex        =   11
         Top             =   600
         Width           =   468
      End
      Begin VB.Label lnkSchool 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cultural"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   216
         Index           =   2
         Left            =   2532
         TabIndex        =   10
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lnkSchool 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Arcane"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   216
         Index           =   1
         Left            =   2568
         TabIndex        =   9
         Top             =   0
         Width           =   708
      End
   End
   Begin CannithCrafting.userInfo usrInfo 
      Height          =   6492
      Left            =   4140
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   6012
      _ExtentX        =   10605
      _ExtentY        =   11451
   End
   Begin CannithCrafting.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7356
      Width           =   13548
      _ExtentX        =   23897
      _ExtentY        =   677
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "New Item;New Gearset;Load Gearset"
      RightLinks      =   "Effects;Augments;Scaling"
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   1260
      TabIndex        =   2
      Top             =   600
      Width           =   2412
   End
   Begin VB.ListBox lstMaterial 
      Appearance      =   0  'Flat
      Height          =   4776
      Left            =   300
      TabIndex        =   3
      Top             =   1020
      Width           =   3672
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
      RightLinks      =   "Help"
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Search:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   624
      Width           =   852
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mtypFarm() As MaterialFarmType
Private mlngFarms As Long

Private menType As MaterialEnum
Private mstrMaterial As String

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Initialize()
    menType = meCollectable
End Sub

Private Sub Form_Load()
    cfg.RefreshColors Me
    ShowMatches
    If Not xp.DebugMode Then
        Call WheelHook(Me.hwnd) ' Hooked to catch wheel
        Call WheelHook(Me.lstMaterial.hwnd) ' Hooked to ignore wheel (because listboxes have built-in wheel support)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then
        Call WheelUnHook(Me.hwnd)
        Call WheelUnHook(Me.lstMaterial.hwnd)
    End If
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim blnHandled As Boolean
    Dim blnOver As Boolean
    Dim lngValue As Long
    Dim ctl As Control
    
    If Rotation < 0 Then lngValue = 3 Else lngValue = -3
    Select Case True
        Case IsOver(Me.usrInfo.hwnd, Xpos, Ypos): Me.usrInfo.Scroll lngValue
        Case IsOver(Me.usrinfoShards.hwnd, Xpos, Ypos): Me.usrinfoShards.Scroll lngValue
        Case IsOver(Me.lstMaterial.hwnd, Xpos, Ypos): Me.lstMaterial.SetFocus
    End Select
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": If menType = meSoulGem Then ShowHelp "Soul Gems" Else ShowHelp "Collectables"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    FooterClick pstrCaption
End Sub

Private Sub lnkSchool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkSchool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkSchool_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    OpenSchool CLng(Index)
End Sub


' ************* OPEN TO MATERIAL *************


Public Property Get Material() As String
    Material = mstrMaterial
End Property

Public Property Let Material(ByVal pstrMaterial As String)
    Dim lngType As Long
    Dim lngTopIndex As Long
    Dim blnFound As Boolean
    Dim lngIndex As Long
    Dim i As Long
    
    lngIndex = SeekMaterial(pstrMaterial)
    If lngIndex Then Me.MaterialType = db.Material(lngIndex).MatType
    If Me.MaterialType = meSoulGem Then pstrMaterial = Mid(pstrMaterial, 11)
    With Me.lstMaterial
        For i = 0 To .ListCount
            If .List(i) = pstrMaterial Then
                blnFound = True
                Exit For
            End If
        Next
        If blnFound Then
            .ListIndex = i
            lngTopIndex = i - (.Height \ Me.TextHeight("Q")) \ 2
            If lngTopIndex < 0 Then lngTopIndex = 0
            .TopIndex = lngTopIndex
        End If
    End With
End Property

Public Property Get MaterialType() As MaterialEnum
    MaterialType = menType
End Property

Public Property Let MaterialType(ByVal penMaterial As MaterialEnum)
    Dim blnChanged As Boolean
    Dim i As Long
    
    blnChanged = (menType <> penMaterial)
    menType = penMaterial
    For i = 0 To 3
        Me.usrchkType(i).Value = (i = menType)
    Next
    If blnChanged Then ShowMatches
End Property


' ************* SEARCH *************


Private Sub txtSearch_GotFocus()
    With Me.txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyReturn
            If Me.lstMaterial.ListCount > 0 Then
                Me.lstMaterial.ListIndex = 0
                Me.lstMaterial.SetFocus
            End If
    End Select
End Sub

Private Sub txtSearch_Change()
    ShowMatches
End Sub

Private Sub usrchkType_UserChange(Index As Integer)
    Me.MaterialType = Index
End Sub

Private Sub ShowMatches()
    Me.lstMaterial.ListIndex = -1
    Me.lstMaterial.Clear
    SetCaption
    If Me.MaterialType = meSoulGem Then
        ShowSoulGems False
        ShowSoulGems True
    Else
        ShowMaterials
    End If
End Sub

Private Sub ShowMaterials()
    Dim blnMatch As Boolean
    Dim i As Long

    For i = 1 To db.Materials
        blnMatch = True
        If Me.MaterialType <> meUnknown Then
            If db.Material(i).MatType <> Me.MaterialType Then blnMatch = False
        End If
        If blnMatch And Len(Me.txtSearch.Text) <> 0 Then
            If InStr(db.Material(i).Material, Me.txtSearch.Text) = 0 Then blnMatch = False
        End If
        If blnMatch Then
            With Me.lstMaterial
                .AddItem db.Material(i).Material
                .ItemData(.NewIndex) = i
            End With
        End If
    Next
End Sub

Private Sub ShowSoulGems(pblnStrong As Boolean)
    Dim blnMatch As Boolean
    Dim i As Long

    For i = 1 To db.Materials
        blnMatch = True
        If db.Material(i).MatType <> meSoulGem Then
            blnMatch = False
        ElseIf pblnStrong Then
            If Left$(db.Material(i).Material, 16) <> "Soul Gem: Strong" Then blnMatch = False
        Else
            If Left$(db.Material(i).Material, 16) = "Soul Gem: Strong" Then blnMatch = False
        End If
        If blnMatch And Len(Me.txtSearch.Text) <> 0 Then
            If InStr(db.Material(i).Material, Me.txtSearch.Text) = 0 Then blnMatch = False
        End If
        If blnMatch Then
            With Me.lstMaterial
                .AddItem Mid$(db.Material(i).Material, 11)
                .ItemData(.NewIndex) = i
            End With
        End If
    Next
End Sub


' ************* DETAILS *************


Private Sub usrInfo_Click(strLink As String)
    Select Case strLink
        Case "Chart": frmValueChart.Show
        Case "Details": OpenValue mstrMaterial
        Case Else: MsgBox "Sorry, feature under construction...", vbInformation, strLink
    End Select
End Sub

Private Sub lstMaterial_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Me.txtSearch.SetFocus
        Case vbKeyUp
            If Me.lstMaterial.ListIndex = 0 Then Me.txtSearch.SetFocus
    End Select
End Sub

Private Sub lstMaterial_Click()
    ShowDetails
End Sub

Private Sub ShowDetails()
    Dim lngType As Long
    Dim lngIndex As Long
    
    Me.usrinfoShards.Visible = False
    Me.usrinfoShards.ClearContents
    Me.usrInfo.Visible = False
    Me.usrInfo.Clear
    SetCaption
    If Me.lstMaterial.ListIndex = -1 Then Exit Sub
    mstrMaterial = db.Material(Me.lstMaterial.ItemData(Me.lstMaterial.ListIndex)).Material
    Me.usrInfo.TitleText = mstrMaterial
    With Me.lstMaterial
        lngIndex = .ItemData(.ListIndex)
    End With
    Select Case db.Material(lngIndex).MatType
        Case meCollectable: ShowCollectable lngIndex
        Case meSoulGem: ShowSoulGem lngIndex
        Case meMisc: ShowMisc lngIndex
    End Select
    ShowShards
    Me.usrInfo.Visible = True
    Me.usrinfoShards.Visible = True
End Sub

Private Sub SetCaption()
    If Me.lstMaterial.ListIndex = -1 Then
        Select Case menType
            Case meCollectable: Me.Caption = "Collectables"
            Case meSoulGem: Me.Caption = "Soul Gems"
            Case Else: Me.Caption = "Materials"
        End Select
    Else
        Me.Caption = Me.lstMaterial.Text
    End If
    ShowIcons
End Sub

Private Sub ShowIcons()
    Dim lngIndex As Long
    Dim strID As String
    
    With Me.lstMaterial
        If .ListIndex <> -1 Then lngIndex = .ItemData(.ListIndex)
    End With
    If lngIndex = 0 Then
        Select Case menType
            Case meSoulGem: SetFormIcon Me, GetGeneralResource(gieSoulGems)
            Case meMisc: SetFormIcon Me, GetGeneralResource(gieIngredients)
            Case Else: SetFormIcon Me, GetGeneralResource(gieCollectables)
        End Select
    Else
        strID = GetMaterialResource(lngIndex)
        Me.usrInfo.SetIcon strID
        SetFormIcon Me, strID
    End If
End Sub

Private Sub ShowCollectable(plngIndex As Long)
    Dim strSchool As String
    
    GetFarms db.Material(plngIndex), mtypFarm, mlngFarms
    With db.Material(plngIndex)
        strSchool = GetSchoolName(.School)
        Me.usrInfo.AddText "Tier " & .Tier, 0
        Me.usrInfo.AddLink strSchool, lseSchool, strSchool, 0
        Me.usrInfo.AddText " (" & GetFrequencyText(.Frequency) & ")"
        If .Frequency = feRare Then
            Me.usrInfo.AddText "Value calculations do not apply to Rare collectables", 2
        Else
            Me.usrInfo.AddText "Value: " & .Value & " Essences (", 0
            Me.usrInfo.BackupOneSpace
            Me.usrInfo.AddLink "Details", lseCommand, "Details", 0
            Me.usrInfo.AddText " /", 0
            Me.usrInfo.AddLink "Chart", lseCommand, "Chart", 0
            Me.usrInfo.AddText ")", 2
        End If
        If xp.DebugMode Then
            With db.School(.School).Tier(.Tier).Freq(.Frequency)
                Me.usrInfo.AddText "Eberron: " & SecondsToTime(.World(weEberron).Fastest)
                Me.usrInfo.AddText "Realms: " & SecondsToTime(.World(weRealms).Fastest), 2
            End With
            Me.usrInfo.AddText "Supply: " & .Supply & " (" & SecondsToTime(.Supply) & ")"
            Me.usrInfo.AddText "Demand: " & .Demand, 2
        End If
        If .Eberron Then Me.usrInfo.AddError "Only found in Eberron, not in the Forgotten Realms.", 2
        If ShowSchoolFarms(.School, .Tier) Then
            Me.usrInfo.AddText vbNullString
            Me.usrInfo.AddText "Farming times collected with an", 0
            Me.usrInfo.AddLink "Epic Challenge Runner", lseURL, "https://www.ddo.com/forums/showthread.php/482405-Epic-Challenge-Runner", 0
            Me.usrInfo.AddText ". Your times may vary."
        End If
        Me.usrInfo.AddText vbNullString
        Me.usrInfo.AddText "See also", 0
        Me.usrInfo.AddLink "this forum thread", lseURL, "https://www.ddo.com/forums/showthread.php/478811-This-is-How-To-Farm-the-New-Collectables-System-Efficiently", 0
        Me.usrInfo.AddText " for more farming data."
    End With
End Sub

Private Function ShowSchoolFarms(penSchool As SchoolEnum, plngTier As Long) As Boolean
    Dim blnBold As Boolean
    Dim blnBagNotice As Boolean
    Dim blnFarmNotice As Boolean
    Dim strBagNotice As String
    Dim i As Long
    
    For i = 1 To mlngFarms
        With mtypFarm(i)
            If .Farm.TreasureBag Then
                If Not blnBagNotice Then
                    Select Case plngTier
                        Case 2: strBagNotice = "Undead tend to drop funerary tokens and necromantic gems, gnolls and goblinoids tend to drop blades of the dark six."
                        Case 3: strBagNotice = "Demons tend to drop planar crystals and spoor, trolls and giants tend to drop amulets."
                    End Select
                    If Len(strBagNotice) Then Me.usrInfo.AddText "Note: Treasure Bags tend to be themed. " & strBagNotice, 2
                    Me.usrInfo.AddText "Mobs that drop Tier " & plngTier & " Treasure Bags:", 2
                    blnBagNotice = True
                End If
                Me.usrInfo.AddLink .Farm.Farm, lseURL, .Farm.Wiki
                AddCultureFarm .Farm.Notes
            Else
                If Not blnFarmNotice Then
                    If penSchool = seCultural Then
                        Me.usrInfo.AddText "Tier " & plngTier & " (Any) farms: (" & Int(db.Backpack(seCultural) * 100 + 0.5) & "% chance to drop Cultural)", 2
                    Else
                        Me.usrInfo.AddText "Good places to farm:", 2
                    End If
                End If
                blnFarmNotice = True
                Me.usrInfo.AddLink .Farm.Farm, lseURL, .Farm.Wiki, 0
                Me.usrInfo.AddText " (" & .Difficulty & ")", 0
                If Len(.Farm.Video) Then Me.usrInfo.AddLink "Video", lseURL, .Farm.Video Else: Me.usrInfo.AddText vbNullString
                ShowFarmStats .Farm, penSchool, .Rate
                blnBold = False
                If Len(.Farm.Need) Then
                    Me.usrInfo.AddTextFormatted "Need:", True, False, False, -1, 0
                    Me.usrInfo.AddText .Farm.Need
                    blnBold = True
                End If
                If Len(.Farm.Fight) Then
                    Me.usrInfo.AddTextFormatted "Fight:", True, False, False, -1, 0
                    Me.usrInfo.AddText .Farm.Fight
                    blnBold = True
                End If
                If Len(.Farm.Notes) Then
                    If blnBold Then Me.usrInfo.AddTextFormatted "Notes:", True, False, False, -1, 0
                    Me.usrInfo.AddText .Farm.Notes
                End If
                Me.usrInfo.AddText vbNullString
            End If
        End With
    Next
    ShowSchoolFarms = blnFarmNotice
End Function

Private Sub ShowFarmStats(ptypFarm As FarmType, penSchool As SchoolEnum, plngRate As Long)
    Dim strTime As String
    Dim strNodes As String
    
    Select Case penSchool
        Case seArcane
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seLore
            NodeCount strNodes, ptypFarm, seNatural
        Case seCultural, seAny
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seLore
            NodeCount strNodes, ptypFarm, seNatural
        Case seLore
            NodeCount strNodes, ptypFarm, seLore
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seNatural
        Case seNatural
            NodeCount strNodes, ptypFarm, seNatural
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seLore
    End Select
    If ptypFarm.Seconds > 0 Then strTime = " in " & SecondsToTime(ptypFarm.Seconds)
    Me.usrInfo.AddText strNodes & strTime
    If plngRate Then Me.usrInfo.AddText SecondsToTime(plngRate) & " to farm one " & mstrMaterial
End Sub

Private Sub NodeCount(pstrNodes As String, ptypFarm As FarmType, penSchool As SchoolEnum)
    Dim strNew As String
    
    Select Case penSchool
        Case seAny: If ptypFarm.Any > 0 Then strNew = ptypFarm.Any & " Any"
        Case seArcane: If ptypFarm.Arcane > 0 Then strNew = ptypFarm.Arcane & " Arcane"
        Case seLore: If ptypFarm.Lore > 0 Then strNew = ptypFarm.Lore & " Lore"
        Case seNatural: If ptypFarm.Natural > 0 Then strNew = ptypFarm.Natural & " Natural"
    End Select
    If Len(strNew) = 0 Then Exit Sub
    If Len(pstrNodes) Then pstrNodes = pstrNodes & ", "
    pstrNodes = pstrNodes & strNew
End Sub

Private Sub AddCultureFarm(pstrNotes As String)
    Dim strNote() As String
    Dim strToken() As String
    Dim strWiki As String
    Dim lngPos As Long
    Dim strLink As String
    Dim i As Long
    
    strNote = Split(pstrNotes, vbNewLine)
    For i = 0 To UBound(strNote)
        strToken = Split(strNote(i), "|")
        If UBound(strToken) <> 2 Then ReDim Preserve strToken(2)
        If Len(strToken(1)) = 0 Then strToken(1) = MakeWiki(strToken(0)) Else strToken(1) = MakeWiki(strToken(1))
        Me.usrInfo.AddText "- " & strToken(0), 0
        Me.usrInfo.AddLinkParentheses "link", lseURL, strToken(1)
        If Len(strToken(2)) Then Me.usrInfo.AddText strToken(2), 1, "- "
    Next
    Me.usrInfo.AddText vbNullString
End Sub

Private Sub ShowSoulGem(plngIndex As Long)
    Dim strHitDice As String
    
    If InStr(db.Material(plngIndex).Material, "Strong") Then strHitDice = "30" Else strHitDice = "20"
    Me.usrInfo.AddText "Use the " & strHitDice & " Hit Dice version of the spell.", 2
    Me.usrInfo.AddText "Good farming locations include:", 2
    Me.usrInfo.AddText db.Material(plngIndex).Notes, 0
    Me.usrInfo.AddText vbNullString, 2
    Me.usrInfo.AddText "For more information, see the {url=" & MakeWiki("Soul Gems") & "}Soul_Gems article on ddowiki."
End Sub

Private Sub ShowMisc(plngIndex As Long)
    Me.usrInfo.AddText db.Material(plngIndex).Notes
End Sub


' ************* SHARDS *************


Private Sub ShowShards()
    Dim strText As String
    Dim blnBound As Boolean
    Dim blnUnbound As Boolean
    Dim strQualifier As String
    Dim lngLines As Long
    Dim i As Long
    
    If Me.lstMaterial.ListIndex = -1 Then Exit Sub
    strText = db.Material(Me.lstMaterial.ItemData(Me.lstMaterial.ListIndex)).Material
    If strText = "Purified Eberron Dragonshard" Then xp.Mouse = msWait
    For i = 1 To db.Shards
        With db.Shard(i)
            blnBound = CheckRecipe(.Bound, strText)
            blnUnbound = CheckRecipe(.Unbound, strText)
        End With
        If blnBound Or blnUnbound Then
            lngLines = 0
            If Not blnBound Then
                strQualifier = " (Unbound)"
            ElseIf Not blnUnbound Then
                strQualifier = " (Bound)"
            Else
                strQualifier = vbNullString
                lngLines = 1
            End If
            With Me.usrinfoShards
                .AddLink db.Shard(i).Abbreviation, lseShard, db.Shard(i).ShardName, lngLines, False
                If Len(strQualifier) Then .AddText strQualifier
            End With
        End If
    Next
    For i = 1 To db.Rituals
        If CheckRecipe(db.Ritual(i).Recipe, strText) Then
            With Me.usrinfoShards
                .AddLink "Eldritch " & db.Ritual(i).RitualName, lseForm, "frmEldritch", 1, False
            End With
        End If
    Next
    If strText = "Purified Eberron Dragonshard" Then xp.Mouse = msNormal
End Sub

Private Function CheckRecipe(ptypRecipe As RecipeType, pstrText As String) As Boolean
    Dim i As Long
    
    With ptypRecipe
        For i = 1 To .Ingredients
            If .Ingredient(i).Material = pstrText Then
                CheckRecipe = True
                Exit For
            End If
        Next
    End With
End Function
