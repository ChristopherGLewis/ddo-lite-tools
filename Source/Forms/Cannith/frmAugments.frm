VERSION 5.00
Begin VB.Form frmAugments 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Augments"
   ClientHeight    =   9024
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
   Icon            =   "frmAugments.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9024
   ScaleWidth      =   13548
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkScaling 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Back to Scaling"
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7380
      Visible         =   0   'False
      Width           =   1872
   End
   Begin CannithCrafting.userInfo usrHelp 
      Height          =   1872
      Left            =   240
      TabIndex        =   17
      Tag             =   "Scaling"
      Top             =   6540
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   3302
      TitleSize       =   2
      CanScroll       =   0   'False
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3240
      ScaleHeight     =   252
      ScaleWidth      =   9972
      TabIndex        =   11
      Top             =   480
      Width           =   9972
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Level"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   -108
         TabIndex        =   16
         Top             =   0
         Width           =   468
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Colorless"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   1140
         TabIndex        =   15
         Top             =   0
         Width           =   816
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Yellow"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   2580
         TabIndex        =   14
         Top             =   0
         Width           =   564
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Blue"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   3
         Left            =   4140
         TabIndex        =   13
         Top             =   0
         Width           =   372
      End
      Begin VB.Label lblHeader 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Red"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   4
         Left            =   5400
         TabIndex        =   12
         Top             =   0
         Width           =   336
      End
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   7812
      Left            =   13260
      TabIndex        =   4
      Top             =   780
      Width           =   252
   End
   Begin CannithCrafting.userAugment usrAugment 
      Height          =   5808
      Left            =   240
      TabIndex        =   2
      Top             =   540
      Width           =   2832
      _ExtentX        =   4995
      _ExtentY        =   10245
   End
   Begin CannithCrafting.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8640
      Width           =   13548
      _ExtentX        =   23897
      _ExtentY        =   677
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "New Item;New Gearset;Load Gearset"
      RightLinks      =   "Effects;Materials;Scaling"
   End
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   13548
      _ExtentX        =   23897
      _ExtentY        =   677
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin CannithCrafting.userInfo usrDetail 
      Height          =   7692
      Left            =   3240
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   6852
      _ExtentX        =   12086
      _ExtentY        =   13568
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7812
      Left            =   3240
      ScaleHeight     =   7812
      ScaleWidth      =   9972
      TabIndex        =   3
      Top             =   780
      Width           =   9972
      Begin VB.PictureBox picScale 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1392
         Left            =   0
         ScaleHeight     =   1392
         ScaleWidth      =   7092
         TabIndex        =   5
         Tag             =   "ctl"
         Top             =   0
         Width           =   7092
         Begin VB.Label lnkRed 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Red"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   5580
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   336
         End
         Begin VB.Label lnkBlue 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Blue"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   4352
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.Label lnkYellow 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Yellow"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   2932
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   564
         End
         Begin VB.Label lblML 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ML"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Tag             =   "ctl"
            Top             =   0
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Label lnkColorless 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Colorless"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   1260
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   816
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Context Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "Variant"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAugments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MaxLevel As Long = 30

Private Type LinkType
    AugmentName As String
    ScaleValue As String
    Augment As Long
    Scale As Long
End Type

Private Type LevelType
    Lines As Long
    Clears As Long
    Clear() As LinkType
    Yellows As Long
    Yellow() As LinkType
    Blues As Long
    Blue() As LinkType
    Reds As Long
    Red() As LinkType
End Type

Private Type ColorType
    Color As Long
    Left As Long
    Width As Long
End Type

Private mtypLevel(1 To MaxLevel) As LevelType
Private mtypColor(4) As ColorType
Private mlngLineColor As Long

Private mlngMarginX As Long
Private mlngRowHeight As Long

Private mlngAugment As Long
Private mlngVariant As Long
Private mlngScale As Long

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    Me.usrAugment.Init aceAny
    LoadData
    SizeControls
    InitGrid
    RefreshColors
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    If Me.picContainer.Visible Then
        If IsOver(Me.picContainer.hwnd, Xpos, Ypos) Then WheelScroll lngValue
    ElseIf Me.usrDetail.Visible Then
        If IsOver(Me.usrDetail.hwnd, Xpos, Ypos) Then Me.usrDetail.Scroll -lngValue
    End If
End Sub

Public Sub RefreshColors()
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        Select Case ctl.Name
            Case "lblML", "lnkColorless", "lnkYellow", "lnkBlue", "lnkRed": ctl.ForeColor = cfg.GetColor(cgeControls, cveText)
        End Select
    Next
    MatchColors
    DrawScaling
    ShowHelpText
End Sub

Public Sub SetAugment(plngAugmentID As Long, plngVariant As Long, plngScale As Long)
    mlngAugment = plngAugmentID
    mlngVariant = plngVariant
    mlngScale = plngScale
    Me.usrAugment.SetSelected mlngAugment, mlngVariant, mlngScale
    ShowDetail
End Sub

Public Function IsMatch(plngAugmentID As Long, plngVariant As Long, plngScale As Long) As Boolean
    If mlngAugment <> plngAugmentID Or mlngVariant <> plngVariant Then Exit Function
    If mlngScale = plngScale Then
        IsMatch = True
    ElseIf mlngAugment <> 0 And mlngVariant <> 0 Then
        If db.Augment(mlngAugment).Scalings = 1 And plngScale = 0 Then IsMatch = True
    End If
End Function


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Back to Scaling": ShowControls True
        Case "Help": ShowHelp "Augments"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    FooterClick pstrCaption
End Sub

Private Sub Form_Click()
    If Me.usrHelp.Tag <> "Scaling" Then ShowControls True
End Sub

Private Sub usrHelp_BackgroundClick()
    If Me.usrHelp.Tag <> "Scaling" Then ShowControls True
End Sub

Private Sub usrDetail_BackgroundClick()
    If Me.usrHelp.Tag <> "Scaling" Then ShowControls True
End Sub

Private Sub chkScaling_Click()
    If UncheckButton(Me.chkScaling, mblnOverride) Then Exit Sub
    ShowControls True
End Sub

Private Sub ShowControls(pblnScaling As Boolean)
    Me.picHeader.Visible = pblnScaling
    Me.picContainer.Visible = pblnScaling
    Me.scrollVertical.Visible = pblnScaling
    Me.usrDetail.Visible = Not pblnScaling
    Me.chkScaling.Visible = False
    If pblnScaling Then
        Me.usrHeader.LeftLinks = vbNullString
        Me.usrAugment.ClearSelection
        Me.Caption = "Augments"
        Me.usrHelp.Tag = "Scaling"
        SetFormIcon Me, "BAGAUGMENT"
    Else
        Me.usrHeader.LeftLinks = "Back to Scaling"
        Me.usrHelp.Tag = "Detail"
    End If
    ShowHelpText
End Sub

Private Sub ShowHelpText()
    With Me.usrHelp
        .Clear
        .MarginX = Me.TextWidth("x")
        .MarginY = Me.TextHeight("Q") \ 2
'        .Height = Me.picContainer.Top + Me.picContainer.Height - .Top
        Select Case .Tag
            Case "Scaling"
                .AddText "Click an augment in the grid on the right to see detail info.", 2
                .AddText "Or choose augments from the list above.", 0
            Case "Detail"
                .AddText "Click anywhere (except the list above) to return to the scaling grid.", 0
        End Select
    End With
End Sub


' ************* DATA *************


Private Sub LoadData()
    Dim enColor As AugmentColorEnum
    Dim lngLevel As Long
    Dim lngCount As Long
    Dim a As Long
    Dim s As Long
    
    For a = 1 To db.Augments
        enColor = db.Augment(a).Color
        For s = 1 To db.Augment(a).Scalings
            lngLevel = db.Augment(a).Scaling(s).ML
            With mtypLevel(lngLevel)
                Select Case enColor
                    Case aceColorless
                        lngCount = AddScale(a, s, .Clear, .Clears)
                    Case aceYellow
                        lngCount = AddScale(a, s, .Yellow, .Yellows)
                    Case aceBlue
                        lngCount = AddScale(a, s, .Blue, .Blues)
                    Case aceRed
                        lngCount = AddScale(a, s, .Red, .Reds)
                    Case aceOrange
                        lngCount = AddScale(a, s, .Red, .Reds)
                        Me.lblHeader(4).Caption = "Red (and Orange)"
                    Case acePurple
                        lngCount = AddScale(a, s, .Blue, .Blues)
                        Me.lblHeader(3).Caption = "Blue (and Purple)"
                    Case aceGreen
                        lngCount = AddScale(a, s, .Yellow, .Yellows)
                        Me.lblHeader(2).Caption = "Yellow (and Green)"
                End Select
                If .Lines < lngCount Then .Lines = lngCount
            End With
        Next
    Next
End Sub

Private Function AddScale(a As Long, s As Long, ptyp() As LinkType, plngCount As Long) As Long
    plngCount = plngCount + 1
    ReDim Preserve ptyp(1 To plngCount)
    With ptyp(plngCount)
        .Augment = a
        .Scale = s
        .AugmentName = db.Augment(a).AugmentName
        .ScaleValue = db.Augment(a).Scaling(s).Value
    End With
    AddScale = plngCount
End Function


' ************* LIST *************


Private Sub usrAugment_AugmentSelected(Augment As Long, Variation As Long, ML As Long)
    mlngAugment = Augment
    mlngVariant = Variation
    mlngScale = ML
    ShowDetail
End Sub


' ************* DETAIL *************


Private Sub ShowDetail()
    Dim strTitle As String
    Dim strText As String
    Dim lngOptions As Long
    Dim strIndent As String
    Dim strID As String
    Dim lngDim As Long
    Dim i As Long
    
    If mlngAugment = 0 Or mlngVariant = 0 Then Exit Sub
    If mlngScale = 0 Then
        If db.Augment(mlngAugment).Scalings = 1 Then mlngScale = 1 Else Exit Sub
    End If
    Me.usrDetail.Clear
    strIndent = "DDO Store:  "
    lngDim = cfg.GetColor(cgeControls, cveTextDim)
    strTitle = AugmentFullName(mlngAugment, mlngVariant, mlngScale)
    Me.Caption = strTitle
    Me.usrDetail.TitleText = strTitle
    With db.Augment(mlngAugment)
        ' Icons
        If Len(.ResourceID) Then
            strID = .ResourceID
        Else
            strID = "AUG" & GetAugmentColorName(.Color, True)
            If .Named Then strID = strID & "NM"
        End If
        Me.usrDetail.SetIcon strID
        SetFormIcon Me, strID
        ' Overview
        If .Named Then
            Me.usrDetail.AddLink .Wiki(mlngVariant), lseURL, MakeWikiItem(.Wiki(mlngVariant)), 0
            Me.usrDetail.AddText " ", 0
        Else
'            strText = Trim(.Scaling(mlngScale).Prefix(mlngVariant) & " " & .Variation(mlngVariant) & " " & .Scaling(mlngScale).Value)
            strText = Trim(.AugmentName & " " & .Scaling(mlngScale).Value)
            Me.usrDetail.AddText strText, 0
        End If
        Me.usrDetail.AddText "(ML" & .Scaling(mlngScale).ML & ")"
        Me.usrDetail.AddText GetAugmentColorName(.Color) & " Augment ", 0
        Me.usrDetail.AddTextFormatted AvailableSlots(.Color), False, , , lngDim
        Me.usrDetail.AddText vbNullString
        ' Description
        If Len(.Descrip) Then ShowNotes .Descrip
        ' Purchase options
        lngOptions = PurchaseOptions()
        Select Case lngOptions
            Case 0: strText = "No purchase options:"
            Case 1: strText = "1 purchase option:"
            Case Else: strText = lngOptions & " purchase options:"
        End Select
        Me.usrDetail.AddText strText
        ' Store
        Me.usrDetail.AddText "DDO Store:", 0
        If .StoreMissing(mlngVariant) Or .Scaling(mlngScale).Store = 0 Then
            Me.usrDetail.AddTextFormatted "Not sold in store", False, , , lngDim, , strIndent
        Else
            Me.usrDetail.AddText .Scaling(mlngScale).Store & " DDO Points", , strIndent
        End If
        ' Astrals
        Me.usrDetail.AddText "Astrals:", 0
        If SoldByVendor(aveCollector) Then
            Me.usrDetail.AddText GetCost(aveCollector), , strIndent
        Else
            Me.usrDetail.AddTextFormatted "Not sold by collectable vendors", False, , , lngDim, , strIndent
        End If
        ' Relics
        Me.usrDetail.AddText "Gianthold:", 0
        If SoldByVendor(aveGianthold) Then
            Me.usrDetail.AddText GetCost(aveGianthold), , strIndent
        Else
            Me.usrDetail.AddTextFormatted "Not available for Gianthold Relics", False, , , lngDim, , strIndent
        End If
        ' Remnants
        Me.usrDetail.AddText "Remnants:", 0
        If .Scaling(mlngScale).Remnants = 0 Then
            Me.usrDetail.AddTextFormatted "Not available for Mysterious Remnants", False, , , lngDim, , strIndent
        Else
            Me.usrDetail.AddText .Scaling(mlngScale).Remnants & " Mysterious Remnants", 0, strIndent
            If .Scaling(mlngScale).RemnantsBonusDays Then Me.usrDetail.AddTextFormatted "(Only during bonus weeks*)", False, , , cfg.GetColor(cgeControls, cveTextDim), 0
            Me.usrDetail.AddText vbNullString
        End If
        ' Tokens
        Me.usrDetail.AddText "Tokens:", 0
        If SoldByVendor(aveLahar) Then
            Me.usrDetail.AddText "20 Tokens of the Twelve", , strIndent
        ElseIf SoldByVendor(aveLahar5) Then
            Me.usrDetail.AddText "5 Greater Tokens of the Twelve", , strIndent
        Else
            Me.usrDetail.AddTextFormatted "Not available for Tokens of the Twelve", False, , , lngDim, , strIndent
        End If
        Me.usrDetail.AddText vbNullString
        ' Champion Hunter Week notice
        If .Scaling(mlngScale).RemnantsBonusDays Then
            Me.usrDetail.AddText "*Exceptional and Insightful stat augments can be bought for Mysterious Remnants only during Champion Hunter bonus weeks, which are pretty rare; maybe once or twice per year.", 2
        End If
        ' Notes
        ShowNotes .Notes
        ' Vendors
        ShowVendors db.Augment(mlngAugment), db.Augment(mlngAugment).Scaling(mlngScale)
        ' Preslotted items
        ShowAugmentItems strTitle
    End With
    ShowControls False
End Sub

Private Function AvailableSlots(penColor As AugmentColorEnum) As String
    Select Case penColor
        Case aceColorless: AvailableSlots = "(goes in any slot)"
        Case aceYellow: AvailableSlots = "(goes in Yellow, Green and Orange slots)"
        Case aceBlue: AvailableSlots = "(goes in Blue, Green and Purple slots)"
        Case aceRed: AvailableSlots = "(goes in Red, Orange and Purple slots)"
        Case aceGreen, acePurple, aceOrange: AvailableSlots = "(goes in " & GetAugmentColorName(penColor) & " slots only)"
    End Select
End Function

Private Function PurchaseOptions() As Long
    Dim lngCount As Long
    
    With db.Augment(mlngAugment)
        If .Scaling(mlngScale).Store > 0 And Not .StoreMissing(mlngVariant) Then lngCount = lngCount + 1
        If SoldByVendor(aveCollector) Then lngCount = lngCount + 1
        If SoldByVendor(aveGianthold) Then lngCount = lngCount + 1
        If .Scaling(mlngScale).Remnants > 0 Then lngCount = lngCount + 1
        If SoldByVendor(aveLahar) Or SoldByVendor(aveLahar5) Then lngCount = lngCount + 1
    End With
    PurchaseOptions = lngCount
End Function

Private Function SoldByVendor(penVendor As AugmentVendorEnum) As Boolean
    Dim i As Long
    
    With db.Augment(mlngAugment).Scaling(mlngScale)
        For i = 1 To .Vendors
            If .Vendor(i) = penVendor Then
                SoldByVendor = True
                Exit For
            End If
        Next
    End With
End Function

Private Function GetCost(penType As AugmentVendorEnum) As String
    Dim blnMatch As Boolean ' Remember not to prematurely exit out of a With...End With block or it will permanantly lock whatever got With'ed
    Dim lngLevel As Long
    Dim i As Long
    
    lngLevel = db.Augment(mlngAugment).Scaling(mlngScale).ML
    For i = 1 To db.AugmentVendors
        With db.AugmentVendor(i)
            If .Style = penType And (.ML = lngLevel Or .AnyLevel = True) Then
                GetCost = .Cost
                blnMatch = True
            End If
        End With
        If blnMatch Then Exit Function
    Next
End Function

Private Sub ShowNotes(pstrRaw As String)
    Dim strLine() As String
    Dim i As Long
    
    If Len(pstrRaw) = 0 Then Exit Sub
    strLine = Split(pstrRaw, vbNewLine)
    For i = 0 To UBound(strLine)
        Me.usrDetail.AddText strLine(i)
    Next
    Me.usrDetail.AddText vbNullString
End Sub

Private Sub ShowVendors(aug As AugmentType, scl As AugmentScaleType)
    Dim blnFound As Boolean
    Dim i As Long
    Dim j As Long
    
    For i = 1 To scl.Vendors
        For j = 1 To db.AugmentVendors
            ShowVendor aug, scl, i, db.AugmentVendor(j), blnFound
        Next
    Next
    If scl.Remnants Then
        If Not blnFound Then VendorHeader
        Me.usrDetail.AddLink "Monster Hunter", lseURL, "https://ddowiki.com/page/Item:Mysterious_Remnant"
        Me.usrDetail.AddTextFormatted "Cost:", True, , , , 0
        Me.usrDetail.AddText scl.Remnants & " Mysterious Remnants"
        Me.usrDetail.AddTextFormatted "Location:", True, , , , 0
        Me.usrDetail.AddText "Hall of Heroes, past bank and auctioneer to far corner"
        Me.usrDetail.AddText vbNullString
    End If
End Sub

Private Sub ShowVendor(aug As AugmentType, scl As AugmentScaleType, i As Long, vnd As AugmentVendorType, pblnFound As Boolean)
    If vnd.Style = aveCollector And aug.Color <> vnd.Color Then Exit Sub
    If vnd.Style <> scl.Vendor(i) Then Exit Sub
    If scl.ML <> vnd.ML And vnd.AnyLevel = False Then Exit Sub
    If Not pblnFound Then
        VendorHeader
        pblnFound = True
    End If
    Me.usrDetail.AddLink vnd.Vendor, lseURL, MakeWiki(vnd.Vendor)
    Me.usrDetail.AddTextFormatted "Cost:", True, , , , 0
    Me.usrDetail.AddText vnd.Cost
    Me.usrDetail.AddTextFormatted "Location:", True, , , , 0
    Me.usrDetail.AddText vnd.Location
    If Len(vnd.Fast) Then
        Me.usrDetail.AddTextFormatted "Fast Access:", True, , , , 0
        Me.usrDetail.AddText vnd.Fast
    End If
    Me.usrDetail.AddText vbNullString
End Sub

Private Sub VendorHeader()
    Me.usrDetail.AddText vbNullString
    Me.usrDetail.AddTextFormatted "Vendors who sell this augment:", False, , , , 2
End Sub

Private Sub ItemHeader()
    Me.usrDetail.AddText vbNullString
    Me.usrDetail.AddTextFormatted "Named Items that come preslotted with this augment:", False, , , , 2
End Sub

Private Sub ShowAugmentItems(pstrAugment As String)
    Dim lngIndex As Long
    Dim strItem As String
    Dim i As Long
    
    lngIndex = SeekAugmentItem(pstrAugment)
    If lngIndex = 0 Then
        If Not db.Augment(mlngAugment).Named Then
            ItemHeader
            Me.usrDetail.AddTextFormatted "None", False, , , cfg.GetColor(cgeWorkspace, cveTextDim)
        End If
        Exit Sub
    End If
    ItemHeader
    With db.AugmentItem(lngIndex)
        For i = 1 To .Items
            If Right$(.Item(i), 1) = "*" Then
                strItem = Left(.Item(i), Len(.Item(i)) - 1)
                Me.usrDetail.AddLink strItem, lseURL, MakeWikiItem(strItem), 0
                Me.usrDetail.AddTextFormatted " (random)", False, , , cfg.GetColor(cgeControls, cveTextDim)
            Else
                Me.usrDetail.AddLink .Item(i), lseURL, MakeWikiItem(.Item(i))
            End If
            
        Next
    End With
End Sub


' ************* GRID *************


Private Sub SizeControls()
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngTop As Long
    Dim lngAvailable As Long
    Dim lngHeight As Long
    Dim lngRows As Long
    
    mlngRowHeight = Me.picScale.TextHeight("Q")
    Me.scrollVertical.Left = Me.ScaleWidth - Me.scrollVertical.Width
    lngLeft = Me.picContainer.Left
    lngWidth = Me.scrollVertical.Left - Me.picContainer.Left
    lngAvailable = Me.usrFooter.Top - (Me.usrHeader.Top + Me.usrHeader.Height)
'    lngRows = (lngAvailable \ mlngRowHeight) - 1
'    lngHeight = lngRows * mlngRowHeight
'    lngTop = Me.usrHeader.Top + Me.usrHeader.Height + (lngAvailable - lngHeight) \ 2
    lngRows = (lngAvailable \ mlngRowHeight) ' - 1
    lngHeight = lngRows * mlngRowHeight
    lngTop = Me.usrFooter.Top - lngHeight
    Me.picHeader.Move lngLeft, lngTop, lngWidth, mlngRowHeight
    Me.picContainer.Move lngLeft, lngTop + mlngRowHeight, lngWidth, (lngRows - 1) * mlngRowHeight
    Me.scrollVertical.Top = Me.picContainer.Top
    Me.scrollVertical.Height = Me.picContainer.Height
End Sub

Private Sub InitGrid()
    Dim lngLeft As Long
    Dim lngRows As Long
    Dim lngWidth As Long
    Dim i As Long
    
    For i = 1 To MaxLevel
        If mtypLevel(i).Lines Then lngRows = lngRows + mtypLevel(i).Lines + 2
    Next
    Me.picScale.Move 0, 0, Me.picContainer.ScaleWidth, mlngRowHeight * lngRows
    With Me.scrollVertical
        .Min = 0
        .Value = 0
        .LargeChange = Me.picContainer.Height \ mlngRowHeight
        .Max = lngRows - .LargeChange
    End With
    Me.picScale.Width = Me.picContainer.ScaleWidth
    With Me.picScale
        mlngMarginX = .TextWidth(" ")
        mtypColor(0).Width = .TextWidth("ML999") + mlngMarginX * 2
        lngWidth = (.ScaleWidth - mtypColor(0).Width) \ 4
    End With
    lngLeft = mtypColor(0).Width
    For i = 1 To 4
        mtypColor(i).Left = lngLeft
        mtypColor(i).Width = lngWidth
        lngLeft = lngLeft + lngWidth
    Next
    For i = 0 To 4
        Me.lblHeader(i).Left = mtypColor(i).Left + mlngMarginX
    Next
    LoadLabels
End Sub

Private Sub LoadLabels()
    Dim lngIndex As Long
    Dim lngWidth As Long
    Dim lngColor As Long
    Dim i As Long
    Dim j As Long
    
    ' ML
    For i = Me.lblML.UBound To 1 Step -1
        Unload Me.lblML(i)
    Next
    lngIndex = 0
    For i = 1 To MaxLevel
        If mtypLevel(i).Lines Then
            lngIndex = lngIndex + 1
            Load Me.lblML(lngIndex)
            With Me.lblML(lngIndex)
                .Caption = "ML" & i
                .Tag = i
                .Visible = False
            End With
        End If
    Next
    ' Colorless
    lngWidth = mtypColor(1).Width
    For i = Me.lnkColorless.UBound To 1 Step -1
        Unload Me.lnkColorless(i)
    Next
    lngIndex = 0
    For i = 1 To MaxLevel
        For j = 1 To mtypLevel(i).Clears
            lngIndex = lngIndex + 1
            Load Me.lnkColorless(lngIndex)
            With Me.lnkColorless(lngIndex)
                .Caption = CreateCaption(mtypLevel(i).Clear(j), lngWidth)
                .Tag = mtypLevel(i).Clear(j).Augment & "|" & mtypLevel(i).Clear(j).Scale
                .Visible = False
            End With
        Next
    Next
    ' Yellow
    lngWidth = mtypColor(1).Width
    For i = Me.lnkYellow.UBound To 1 Step -1
        Unload Me.lnkYellow(i)
    Next
    lngIndex = 0
    For i = 1 To MaxLevel
        For j = 1 To mtypLevel(i).Yellows
            lngIndex = lngIndex + 1
            Load Me.lnkYellow(lngIndex)
            With Me.lnkYellow(lngIndex)
                .Caption = CreateCaption(mtypLevel(i).Yellow(j), lngWidth)
                .Tag = mtypLevel(i).Yellow(j).Augment & "|" & mtypLevel(i).Yellow(j).Scale
                .Visible = False
            End With
        Next
    Next
    ' Blue
    lngWidth = mtypColor(1).Width
    For i = Me.lnkBlue.UBound To 1 Step -1
        Unload Me.lnkBlue(i)
    Next
    lngIndex = 0
    For i = 1 To MaxLevel
        For j = 1 To mtypLevel(i).Blues
            lngIndex = lngIndex + 1
            Load Me.lnkBlue(lngIndex)
            With Me.lnkBlue(lngIndex)
                .Caption = CreateCaption(mtypLevel(i).Blue(j), lngWidth)
                .Tag = mtypLevel(i).Blue(j).Augment & "|" & mtypLevel(i).Blue(j).Scale
                .Visible = False
            End With
        Next
    Next
    ' Red
    lngWidth = mtypColor(1).Width
    For i = Me.lnkRed.UBound To 1 Step -1
        Unload Me.lnkRed(i)
    Next
    lngIndex = 0
    For i = 1 To MaxLevel
        For j = 1 To mtypLevel(i).Reds
            lngIndex = lngIndex + 1
            Load Me.lnkRed(lngIndex)
            With Me.lnkRed(lngIndex)
                .Caption = CreateCaption(mtypLevel(i).Red(j), lngWidth)
                .Tag = mtypLevel(i).Red(j).Augment & "|" & mtypLevel(i).Red(j).Scale
                .Visible = False
            End With
        Next
    Next
End Sub

Private Function CreateCaption(ptypLink As LinkType, plngWidth As Long) As String
    Dim lngAugmentWidth As Long
    Dim lngTotalWidth As Long
    Dim lngValueWidth As Long
    Dim strText As String
    
    lngTotalWidth = plngWidth - mlngMarginX * 2
    If Len(ptypLink.ScaleValue) Then lngValueWidth = Me.picScale.TextWidth(" " & ptypLink.ScaleValue)
    lngAugmentWidth = lngTotalWidth - lngValueWidth
    strText = ptypLink.AugmentName
    With Me.picScale
        If .TextWidth(strText) > lngAugmentWidth Then
            Do While .TextWidth(strText & "...") > lngAugmentWidth
                strText = Left$(strText, Len(strText) - 1)
            Loop
            strText = Trim$(strText) & "..."
        End If
    End With
    If Len(ptypLink.ScaleValue) Then strText = strText & " " & ptypLink.ScaleValue
    CreateCaption = strText
End Function

Private Sub MatchColors()
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim lngMid As Long
    Dim lngHigh As Long
    Dim lngMed As Long
    Dim lngLow As Long
    Dim i As Long
    
    ' RGB
    xp.ColorToRGB cfg.GetColor(cgeControls, cveBackground), lngRed, lngGreen, lngBlue
    lngMid = Int((lngRed + lngGreen + lngBlue) / 3 + 0.5)
    lngMid = ((lngMid + 2) \ 5) * 5
    Select Case lngMid
        Case Is < 60: lngMid = 70
        Case Is < 85: lngMid = 85
        Case Is > 230: lngMid = 230
    End Select
    lngHigh = lngMid + 25
    lngMed = lngMid
    lngLow = lngMid - 25
    mtypColor(0).Color = cfg.GetColor(cgeControls, cveBackground)
    mtypColor(1).Color = RGB(lngMed, lngMed, lngMed)
    mtypColor(2).Color = RGB(lngHigh, lngHigh, lngLow)
    mtypColor(3).Color = RGB(lngLow, lngLow, lngHigh)
    mtypColor(4).Color = RGB(lngHigh, lngLow, lngLow)
    mlngLineColor = cfg.GetColor(cgeWorkspace, cveBackground) ' RGB(lngLow, lngLow, lngLow)
End Sub

Private Sub DrawScaling()
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngLevel As Long
    Dim lngRow As Long
    Dim lngTop As Long
    Dim lngIndex(4) As Long
    Dim i As Long
    
    lngWidth = Me.picScale.ScaleWidth
    lngHeight = Me.picScale.ScaleHeight
    For i = 0 To 4
        With mtypColor(i)
            Me.picScale.Line (.Left, 0)-(.Left + .Width, lngHeight), .Color, BF
        End With
    Next
    lngTop = 0
    For lngLevel = 1 To MaxLevel
        With mtypLevel(lngLevel)
            If .Lines Then
                Me.picScale.Line (0, lngTop)-(lngWidth, lngTop), mlngLineColor
                lngTop = lngTop + mlngRowHeight
                lngIndex(0) = lngIndex(0) + 1
                With Me.lblML(lngIndex(0))
                    .Move mtypColor(0).Left + mlngMarginX, lngTop
                    .Visible = True
                End With
                For lngRow = 1 To .Lines
                    If lngRow <= .Clears Then
                        lngIndex(1) = lngIndex(1) + 1
                        With Me.lnkColorless(lngIndex(1))
                            .Move mtypColor(1).Left + mlngMarginX, lngTop
                            .Visible = True
                        End With
                    End If
                    If lngRow <= .Yellows Then
                        lngIndex(2) = lngIndex(2) + 1
                        With Me.lnkYellow(lngIndex(2))
                            .Move mtypColor(2).Left + mlngMarginX, lngTop
                            .Visible = True
                        End With
                    End If
                    If lngRow <= .Blues Then
                        lngIndex(3) = lngIndex(3) + 1
                        With Me.lnkBlue(lngIndex(3))
                            .Move mtypColor(3).Left + mlngMarginX, lngTop
                            .Visible = True
                        End With
                    End If
                    If lngRow <= .Reds Then
                        lngIndex(4) = lngIndex(4) + 1
                        With Me.lnkRed(lngIndex(4))
                            .Move mtypColor(4).Left + mlngMarginX, lngTop
                            .Visible = True
                        End With
                    End If
                    lngTop = lngTop + mlngRowHeight
                Next
                lngTop = lngTop + mlngRowHeight
            End If
        End With
    Next
End Sub


' ************* SCROLLBAR *************


Private Sub scrollVertical_GotFocus()
    Me.picScale.SetFocus
End Sub

Private Sub scrollVertical_Change()
    Scroll
End Sub

Private Sub scrollVertical_Scroll()
    Scroll
End Sub

Private Sub Scroll()
    Me.picScale.Top = 0 - Me.scrollVertical.Value * mlngRowHeight
End Sub

Private Sub WheelScroll(plngIncrement As Long)
    Dim lngValue As Long
    
    With Me.scrollVertical
        lngValue = .Value - plngIncrement
        If lngValue < 0 Then
            lngValue = 0
        ElseIf lngValue > .Max Then
            lngValue = .Max
        End If
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub


' ************* GRID LINKS *************


Private Sub lnkColorless_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkColorless_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickLink Me.lnkColorless(Index).Tag
End Sub

Private Sub lnkColorless_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkBlue_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkBlue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickLink Me.lnkBlue(Index).Tag
End Sub

Private Sub lnkBlue_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkRed_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkRed_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickLink Me.lnkRed(Index).Tag
End Sub

Private Sub lnkRed_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkYellow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkYellow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClickLink Me.lnkYellow(Index).Tag
End Sub

Private Sub lnkYellow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub ClickLink(pstrTag As String)
    Dim strSplit() As String
    Dim i As Long
    
    strSplit = Split(pstrTag, "|")
    mlngAugment = Val(strSplit(0))
    mlngScale = Val(strSplit(1))
    If db.Augment(mlngAugment).Variations = 1 Then
        mlngVariant = 1
        Me.usrAugment.SetSelected mlngAugment, mlngVariant, mlngScale
        ShowDetail
        Exit Sub
    End If
    With db.Augment(mlngAugment)
        For i = 1 To .Variations + 2
            If i > Me.mnuContext.UBound Then Load Me.mnuContext(i)
            Select Case i
                Case .Variations + 1
                    Me.mnuContext(i).Caption = "-"
                    Me.mnuContext(i).Tag = 0
                Case .Variations + 2
                    Me.mnuContext(i).Caption = "Never Mind"
                    Me.mnuContext(i).Tag = 0
                Case Else
                    Me.mnuContext(i).Caption = AugmentFullName(mlngAugment, i, mlngScale)
                    Me.mnuContext(i).Tag = i
            End Select
            Me.mnuContext(i).Visible = True
        Next
        Me.mnuContext(0).Visible = False
        For i = Me.mnuContext.UBound To .Variations + 3 Step -1
            Unload Me.mnuContext(i)
        Next
    End With
    PopupMenu Me.mnuMain(0)
End Sub

Private Sub mnuContext_Click(Index As Integer)
    mlngVariant = Val(Me.mnuContext(Index).Tag)
    If mlngVariant = 0 Then
        mlngAugment = 0
        mlngScale = 0
    Else
        Me.usrAugment.SetSelected mlngAugment, mlngVariant, mlngScale
        ShowDetail
    End If
End Sub
