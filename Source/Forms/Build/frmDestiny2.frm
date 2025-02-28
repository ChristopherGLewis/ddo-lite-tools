VERSION 5.00
Begin VB.Form frmDestiny2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Destinies"
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
   Icon            =   "frmDestiny2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12225
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6828
      Index           =   1
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   12015
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   12012
      Begin VB.ComboBox cboTree 
         Height          =   312
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   3492
      End
      Begin VB.ListBox lstAbility 
         Appearance      =   0  'Flat
         Height          =   3012
         IntegralHeight  =   0   'False
         ItemData        =   "frmDestiny2.frx":000C
         Left            =   5520
         List            =   "frmDestiny2.frx":000E
         TabIndex        =   16
         Top             =   1080
         Width           =   3492
      End
      Begin VB.ListBox lstSub 
         Appearance      =   0  'Flat
         Height          =   3012
         IntegralHeight  =   0   'False
         ItemData        =   "frmDestiny2.frx":0010
         Left            =   9240
         List            =   "frmDestiny2.frx":0012
         TabIndex        =   18
         Top             =   1080
         Width           =   2652
      End
      Begin CharacterBuilderLite.userList usrList 
         Height          =   6492
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   11456
      End
      Begin CharacterBuilderLite.userDetails usrDetails 
         Height          =   2040
         Left            =   5520
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4452
         Width           =   6372
         _ExtentX        =   7011
         _ExtentY        =   3598
      End
      Begin CharacterBuilderLite.userCheckBox usrchkShowAll 
         Height          =   252
         Left            =   7200
         TabIndex        =   15
         Top             =   816
         Width           =   1812
         _ExtentX        =   3201
         _ExtentY        =   450
         Value           =   0   'False
         Caption         =   "Show All"
         CheckPosition   =   1
      End
      Begin VB.Label lblTier5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tier 5"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   9240
         TabIndex        =   13
         Top             =   406
         Visible         =   0   'False
         Width           =   504
      End
      Begin VB.Label lblTier5Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tier 5 Tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   9240
         TabIndex        =   12
         Top             =   132
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cost: 2 AP per rank"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   8160
         TabIndex        =   23
         Top             =   6540
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblRanks 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Ranks: 3"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7200
         TabIndex        =   22
         Top             =   6540
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblProg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "12 AP spent in tree"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   10200
         TabIndex        =   24
         Top             =   6540
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label lblTree 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   5520
         TabIndex        =   10
         Top             =   132
         Width           =   396
      End
      Begin VB.Label lblSpent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Spent in Tree: 24 AP"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4320
         TabIndex        =   9
         Top             =   6540
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0 / 80 AP"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   480
         TabIndex        =   8
         Top             =   6540
         Width           =   852
      End
      Begin VB.Label lblAbilities 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Abilities"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   5520
         TabIndex        =   14
         Top             =   840
         Width           =   648
      End
      Begin VB.Label lblSelectors 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Selectors"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   9240
         TabIndex        =   17
         Top             =   840
         Width           =   828
      End
      Begin VB.Label lblDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Details"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   5520
         TabIndex        =   19
         Top             =   4212
         Width           =   6372
      End
   End
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   25
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
      LeftLinks       =   "Destiny Trees;Destinies"
      RightLinks      =   "Help"
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6492
      Index           =   0
      Left            =   600
      ScaleHeight     =   6495
      ScaleWidth      =   12015
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   12012
      Begin VB.Frame fraDestinyAP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2595
         Left            =   5800
         TabIndex        =   26
         Top             =   3780
         Width           =   5000
         Begin CharacterBuilderLite.userSpinner usrspnPDP 
            Height          =   300
            Left            =   540
            TabIndex        =   27
            Top             =   360
            Width           =   975
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
         Begin VB.Label lblPDPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- You get 36 FP's for unlocking all Destiny Trees"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   360
            TabIndex        =   33
            Top             =   1920
            Width           =   4470
         End
         Begin VB.Label lblPDPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Fate/Destiny Tomes can add more points"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   3
            Left            =   360
            TabIndex        =   32
            Top             =   1620
            Width           =   4005
         End
         Begin VB.Label lblPDPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Fate Point are earned 1 per 3 EPL "
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   2
            Left            =   360
            TabIndex        =   31
            Top             =   1320
            Width           =   3405
         End
         Begin VB.Label lblPermDestinyPoints 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Perm Destiny Points"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   29
            Top             =   0
            Width           =   1890
         End
         Begin VB.Shape shpPermDestinyPoints 
            Height          =   2235
            Left            =   0
            Top             =   240
            Width           =   4995
         End
         Begin VB.Label lblPDPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Perm Destiny Points can be spent at level 20"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   30
            Top             =   720
            Width           =   4365
         End
         Begin VB.Label lblPDPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- PDPs are earned at 1 per 3 Fate Points"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   28
            Top             =   1020
            Width           =   3840
         End
      End
      Begin VB.Frame fraTreeSelection 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3492
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   10092
         Begin CharacterBuilderLite.userList usrTree 
            Height          =   3012
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   5412
            _ExtentX        =   9551
            _ExtentY        =   5318
         End
         Begin VB.ListBox lstTree 
            Appearance      =   0  'Flat
            Height          =   2652
            IntegralHeight  =   0   'False
            ItemData        =   "frmDestiny2.frx":0014
            Left            =   5880
            List            =   "frmDestiny2.frx":0036
            TabIndex        =   5
            Top             =   360
            Width           =   3972
         End
         Begin VB.Label lblSpentAll 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0 / 80 AP"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   3360
            TabIndex        =   3
            Top             =   3120
            Width           =   4000
         End
         Begin VB.Label lblTree 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tree"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   5880
            TabIndex        =   4
            Top             =   120
            Width           =   2352
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Trees"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuTrees 
         Caption         =   "Delete Tree"
         Index           =   0
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "Delete All Trees"
         Index           =   1
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "Reset Tree"
         Index           =   3
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "Reset All Trees"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Destinies"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuDestinies 
         Caption         =   "Clear this ability"
         Index           =   0
      End
      Begin VB.Menu mnuDestinies 
         Caption         =   "Reset tree"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Guide"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuGuide 
         Caption         =   "Move Row(s)"
         Index           =   0
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Delete Row(s)"
         Index           =   2
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Delete to End"
         Index           =   3
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Delete All"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmDestiny2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Enum MouseShiftEnum
    mseNormal
    mseCtrl
    mseShift
End Enum

Private Type ColumnType
    Header As String
    Left As Long
    Width As Long
    Right As Long
    Align As AlignmentConstants
End Type


Private Col(1 To 6) As ColumnType
Private mlngRow As Long
Private mlngCol As Long
Private mlngLastRowSelected As Long
Private mblnMoveRows As Boolean
Private mlngHeight As Long
Private mlngOffsetX As Long
Private mlngOffsetY As Long
Private mlngInterval As Long
Private mlngDirection As Long

' Destiny
Private mlngDestiny As Long   'This represents the current selected destiny
Private mlngBuildDestiny As Long  'This represents the current destiny in the build
Private mlngMaxTier As Long

' Drag & Drop
Private mblnMouse As Boolean
Private mlngSourceIndex As Long
Private menDragState As DragEnum
Private mblnDragComplete As Boolean
Private msngDownX As Single
Private msngDownY As Single
' General
Private mlngTab As Long  'current tab of focus
Private mblnNoFocus As Boolean
Private mblnOverride As Boolean

Const MAX_DESTINIES As Long = 3


' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.Configure Me
    mlngTab = 0
    LoadData
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Activate()
    ActivateForm oeDestiny2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If mlngTab = 2 Then ActiveCell -1, -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    UnloadForm Me, mblnOverride
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case mlngTab
        Case 0  'DestinyTrees
            If IsOver(Me.usrspnPDP.hwnd, Xpos, Ypos) Then Me.usrspnPDP.WheelScroll lngValue
        Case 1 'Destiny
            Select Case True
                Case IsOver(Me.usrList.hwnd, Xpos, Ypos): Me.usrList.Scroll lngValue
                Case IsOver(Me.usrDetails.hwnd, Xpos, Ypos): Me.usrDetails.Scroll lngValue
            End Select
    End Select
End Sub

Public Sub Cascade()
    Dim i As Long
    
    NoFocus True
    xp.LockWindow Me.hwnd
    LoadData
    Select Case mlngTab
        Case 1 ' Destinies
            If mlngDestiny <> 0 Then
                mlngBuildDestiny = 0
                For i = 1 To build.Destinies
                    If build.Destiny(i).TreeName = db.Destiny(mlngDestiny).TreeName Then
                        mlngBuildDestiny = i
                        Exit For
                    End If
                Next
            End If
    End Select
    ShowTab
    xp.UnlockWindow
    NoFocus False
End Sub

Private Sub NoFocus(pblnNoFocus As Boolean)
    mblnNoFocus = pblnNoFocus
    Me.usrTree.NoFocus = pblnNoFocus
    Me.usrList.NoFocus = pblnNoFocus
    Me.usrDetails.NoFocus = pblnNoFocus
End Sub


Public Property Get CurrentTab() As Variant
    CurrentTab = mlngTab
End Property


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Destiny Trees": ChangeTab 0
        Case "Destinies": ChangeTab 1
        Case "Export": ExportGuide "csv"
        Case "Help": ShowHelp "Destinies"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "< Enhancements": If Not OpenForm("frmEnhancements") Then Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub

Private Sub ChangeTab(plngTab As Long)
    If mlngTab = plngTab Then Exit Sub
    If cboTree.ListCount = 0 Then Exit Sub  'Exit if none are selected
    mlngTab = plngTab
    ShowTab
    SaveBackup
    GenerateOutput oeDestiny
End Sub

Private Sub ShowTab()
    Dim i As Long
    
    xp.LockWindow Me.hwnd
    Me.usrHeader.RightLinks = "Help"
    Select Case mlngTab
        Case 0 ' DTrees
            ShowDestinies False, False
            ShowAvailableDestinies
        Case 1 ' Destinies
            ShowDestinyTier5
            Me.usrList.GotoTop
            If mlngBuildDestiny = 0 Then
                Me.cboTree.ListIndex = 0
                TreeClick
            Else
                ShowAbilities
                ShowAvailable False
            End If
    End Select
    For i = 0 To 1
        Me.picTab(i).Visible = (mlngTab = i)
    Next
    xp.UnlockWindow
End Sub


' ************* INITIALIZE *************


Private Sub LoadData()
    Dim lngTabs() As Long
    Dim lngMax As Long
    Dim i As Long
    
    mlngSourceIndex = 0
    ' Navigation
    Me.usrFooter.LeftLinks = "< Enhancements"
    Me.usrFooter.RightLinks = vbNullString
    ' Destinies
    With Me.usrList
        .DefineDimensions 1, 3, 2
        .DefineColumn 1, vbCenter, "Tier", "Tier"
        .DefineColumn 2, vbCenter, "Ability"
        .DefineColumn 3, vbCenter, "AP", " 20"
        .Refresh
    End With
    Me.usrDetails.Clear
    PopulateCombo
    mblnOverride = True
    ' Destiny Trees
    With Me.usrTree
        .DefineDimensions 1, 4, 2
        'TODO set to longest Dest in consts
        .DefineColumn 1, vbLeftJustify, "Destiny", MAX_DESTINY_NAME
        .DefineColumn 2, vbCenter, "Tree"
        .DefineColumn 3, vbCenter, "Tier", "Tier"
        .DefineColumn 4, vbCenter, "AP", "AP"
        .Refresh
    End With
    ReDim lngTabs(0)  'Tabstops for lstTree
    lngTabs(0) = 86
    lngTabs(0) = 97
    ListboxTabStops Me.lstTree, lngTabs
    ShowDestinies False, False
    mblnOverride = True
    ShowAvailableDestinies
    
    ' Perm Destiny Points
    Me.usrspnPDP.Max = tomes.PermDestinyPointMax
    Me.usrspnPDP.Min = tomes.PermDestinyPointMin
    If build.PermDestinyPoints < tomes.PermDestinyPointMin Then
        build.PermDestinyPoints = tomes.PermDestinyPointMin
    End If
    
    Me.usrspnPDP.Value = build.PermDestinyPoints
    
    Me.lblTier5Label.Visible = (build.MaxLevels > 29)
    Me.lblTier5.Visible = (build.MaxLevels > 29)
    If build.MaxLevels < 29 And Len(build.DestinyTier5) Then
        ShaveTree build.DestinyTier5   'TODO
        build.DestinyTier5 = vbNullString
    End If
    
    ' Leveling Guide
    'Me.cboGuideAbility.ListIndex = 0
    'mblnOverride = False
End Sub

Private Sub PopulateCombo()
    Dim strTree As String  'TODO
    Dim i As Long
    
    mblnOverride = True
    strTree = Me.cboTree.Text
    ComboClear Me.cboTree
    'Load combo with build destinies
    For i = 1 To build.Destinies
        ComboAddItem Me.cboTree, build.Destiny(i).TreeName, i
    Next
    mblnOverride = False
    If Len(strTree) Then ComboSetText Me.cboTree, strTree
    If Me.cboTree.ListIndex = -1 Then
        If Len(build.DestinyTier5) Then
            ComboSetText Me.cboTree, build.DestinyTier5
        End If
        TreeClick
    End If
End Sub


' ************* TREES *************
Private Sub usrspnPDP_Change()
    Dim lngPDP  As Long
    
    'Form based override - used to prevent recalcing???
    If mblnOverride Then
        Exit Sub
    End If
    build.PermDestinyPoints = Me.usrspnPDP.Value
    If build.PermDestinyPoints > tomes.PermDestinyPointMax Then
        lngPDP = tomes.PermDestinyPointMax
    End If
    ShowSpentAll Me.lblSpentAll
    SetDirty
End Sub


Private Sub usrTree_SlotClick(Index As Integer, Button As Integer)
    Dim strTreeName As String
    Dim lngTree As Long
    Dim i As Long
    
    Me.lstTree.ListIndex = -1
    Select Case Button
        Case vbLeftButton
            If Me.usrTree.Selected = Index Then Me.usrTree.Selected = 0 Else Me.usrTree.Selected = Index
        Case vbRightButton
            Me.usrTree.Selected = Index
            Me.usrTree.Active = Index
            lngTree = SeekTree(build.Destiny(Index).TreeName, peDestiny)
            Me.mnuTrees(0).Caption = "Delete " & db.Destiny(lngTree).Abbreviation
            Me.mnuTrees(0).Enabled = (build.Destiny(Index).TreeType <> tseRace)
            Me.mnuTrees(3).Caption = "Reset " & db.Destiny(lngTree).Abbreviation
            PopupMenu Me.mnuMain(0)
    End Select
End Sub

Private Sub mnuTrees_Click(Index As Integer)
    Dim lngBuildDestiny As Long
    Dim i As Long
    
    Select Case Index
        Case 0 ' Delete tree
            lngBuildDestiny = Me.usrTree.Selected
            If lngBuildDestiny = 0 Then Exit Sub
            For i = lngBuildDestiny To build.Destinies - 1
                build.Destiny(i) = build.Destiny(i + 1)
            Next
            build.Destinies = build.Destinies - 1
            ReDim Preserve build.Destiny(1 To build.Destinies)  'Reset.  Destiny is 1 based
            PopulateCombo
        Case 1 ' Delete all trees
            If Not Ask("Delete all trees?") Then Exit Sub
            ReDim build.Destiny(0)
            build.Destinies = 0
            PopulateCombo
        Case 3 ' Reset tree
            lngBuildDestiny = Me.usrTree.Selected
            If lngBuildDestiny = 0 Then Exit Sub
            With build.Destiny(lngBuildDestiny)
                Erase .Ability
                .Abilities = 0
            End With
        Case 4 ' Reset all trees
            If Not Ask("Reset all trees?") Then Exit Sub
            For lngBuildDestiny = 1 To build.Destinies
                With build.Destiny(lngBuildDestiny)
                    Erase .Ability
                    .Abilities = 0
                End With
            Next
    End Select
    ShowDestinies False, False
    ShowAvailableDestinies
    SetDirty
End Sub

Private Sub usrTree_SlotDblClick(Index As Integer)
    Dim strTreeName As String
    Dim i As Long
    
    Me.lstTree.ListIndex = -1
    With build
        If .DestinyTier5 = .Destiny(Index).TreeName Then
            .DestinyTier5 = vbNullString
        End If
        For i = Index To .Destinies - 1
            .Destiny(i) = .Destiny(i + 1)
        Next
        .Destinies = .Destinies - 1
        If .Destinies = 0 Then Erase .Destiny Else ReDim Preserve .Destiny(1 To .Destinies)
    End With
    ShowDestinies False, False
    ShowAvailableDestinies
    PopulateCombo
    SetDirty
End Sub

Private Sub ShowDestinies(pblnDrop As Boolean, pblnAdd As Boolean)
    Dim lngDestiny As Long
    Dim lngMax As Long
    Dim lngRaceTree As Long
    Dim enDropState As DropStateEnum
    Dim i As Long
    
    mblnOverride = True
    If pblnAdd And build.Destinies < MAX_DESTINIES Then
        lngMax = build.Destinies + 1
    Else
        lngMax = build.Destinies
    End If
    Me.usrTree.Rows = lngMax
    For i = 1 To lngMax
        If i <= build.Destinies Then
            With build.Destiny(i)
                Me.usrTree.SetSlot i, .TreeName
                lngDestiny = SeekTree(.TreeName, peDestiny)
                Me.usrTree.SetItemData i, lngDestiny
                Me.usrTree.SetText i, 1, GetMaxTier(i)
                Me.usrTree.SetText i, 4, QuickSpentInDestiny(i)
            End With
        End If
        If pblnDrop And i <> lngRaceTree Then
            enDropState = dsCanDrop
        Else
            enDropState = dsDefault
        End If
        Me.usrTree.SetDropState i, enDropState
    Next
    ShowSpentAll Me.lblSpentAll
    mblnOverride = False
End Sub

'This shows the spent line on the Destiny tab
'TODO move to a bas file so it can be shared btwn Enh and main???
Private Sub ShowSpentAll(plbl As Label)
    Dim lngSpentBase As Long  'Base spent in tree
    Dim lngSpentPDPBonus As Long ' Additional PermDestPoints spent
    Dim lngMaxBase As Long
    Dim lngMaxPDPBonus As Long
    Dim strDisplay As String
    
    'Display should be Spent/Max AP.  Long form is Spent+pdp/Max+pdp AP

    'retrieve each of the spent/maxes from the build tree object
    'This should be a func that returns a class object
    GetDestinyPointsSpentAndMax lngSpentBase, lngSpentPDPBonus, lngMaxBase, lngMaxPDPBonus
    
    'Display
    strDisplay = strDisplay & "Spent: " & lngSpentBase & " +" & lngSpentPDPBonus & "pdp / Max: "
    strDisplay = strDisplay & lngMaxBase & " +" & lngMaxPDPBonus & "pdp AP"
    
    plbl.Caption = strDisplay
    
    'Flag this as an error if spent is > max
    If lngSpentBase > lngMaxBase Then
        plbl.ForeColor = cfg.GetColor(cgeWorkspace, cveTextError)
    Else
        plbl.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    End If
    plbl.Visible = True
End Sub

Private Function PointsSpent(plngBuildDestiny As Long) As Long
    Dim lngDestiny As Long
    Dim lngTotal As Long
    Dim lngSpent() As Long
    
    lngDestiny = SeekTree(build.Destiny(plngBuildDestiny).TreeName, peDestiny)
    GetSpentInTree db.Destiny(lngDestiny), build.Destiny(plngBuildDestiny), lngSpent, lngTotal
    PointsSpent = lngTotal
End Function

Private Sub ShowAvailableDestinies()
    Dim lngDestiny As Long
    Dim i As Long
    
    ListboxClear Me.lstTree
    ' Add destiny trees
    For i = 1 To db.Destinies
        lngDestiny = SeekTree(db.Destiny(i).TreeName, peDestiny)
        If lngDestiny Then AddDestiny lngDestiny, db.Destiny(i).TreeName
    Next
    PopulateCombo
End Sub

Private Sub AddDestiny(plngDestiny As Long, pstrSource As String)
    Dim i As Long
    
    ' Already selected?
    For i = 1 To build.Destinies
        If build.Destiny(i).TreeName = db.Destiny(plngDestiny).TreeName Then Exit Sub
    Next
    ' Already added to available list?
    For i = 0 To Me.lstTree.ListCount - 1
        If Me.lstTree.ItemData(i) = plngDestiny Then
            Exit Sub
        End If
    Next
    ' Locked out?
    If Len(db.Destiny(plngDestiny).Lockout) Then
        For i = 1 To build.Destinies
            If build.Destiny(i).TreeName = db.Destiny(plngDestiny).Lockout Then Exit Sub
        Next
    End If
    ' Add this tree
    ListboxAddItem Me.lstTree, db.Destiny(plngDestiny).TreeName, plngDestiny
End Sub

Private Sub lstTree_Click()
    If mblnMouse Then mblnMouse = False
End Sub

Private Sub lstTree_DblClick()
    If Me.lstTree.ListIndex = -1 Or build.Destinies >= MAX_DESTINIES Then Exit Sub
    AddBuildDestiny build.Destinies + 1
    ShowDestinies False, False
    ShowAvailableDestinies
    PopulateCombo
    Me.usrTree.Active = build.Destinies
    SetDirty
End Sub

Private Sub lstTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.usrTree.Selected = 0
    mblnMouse = True
    menDragState = dragMouseDown
End Sub

Private Sub lstTree_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            ShowDestinies True, True
            Me.lstTree.OLEDropMode = vbOLEDropManual
            Me.lstTree.OLEDrag
        End If
    End If
End Sub

Private Sub lstTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub lstTree_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "Add"
End Sub

Private Sub lstTree_OLECompleteDrag(Effect As Long)
    ShowDestinies False, False
    Me.lstTree.OLEDropMode = vbOLEDropNone
End Sub

Private Sub usrTree_OLEDragDrop(Index As Integer, Data As DataObject)
    Dim strData As String
    
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strData = Data.GetData(vbCFText)
    If strData = "Add" Then
        AddBuildDestiny Index
    Else
        SwapBuildDestinies Index, strData
    End If
    ShowDestinies False, False
    ShowAvailableDestinies
    PopulateCombo
    SetDirty
End Sub

Private Sub usrTree_OLECompleteDrag(Index As Integer, Effect As Long)
    ShowDestinies False, False
End Sub

Private Sub usrTree_RequestDrag(Index As Integer, Allow As Boolean)
    Dim enDropState As DropStateEnum
    Dim i As Long
    
    For i = 1 To build.Destinies
        If i = Index Then enDropState = dsDefault Else enDropState = dsCanDrop
        Me.usrTree.SetDropState i, enDropState
    Next
    Allow = True
End Sub

''Add Clicked destiny to Build
Private Sub AddBuildDestiny(ByVal plngIndex As Long)
    Dim strText() As String
    Dim lngDestiny As Long
    
    If Me.lstTree.ListIndex = -1 Then Exit Sub
    strText = Split(Me.lstTree.Text, vbTab)
    With build
        If plngIndex > .Destinies Then
            .Destinies = plngIndex
            ReDim Preserve .Destiny(1 To plngIndex)
        End If
        With .Destiny(plngIndex)
            .TreeName = strText(0)
            .TreeType = tseDestiny
            .Abilities = 0
            Erase .Ability
        End With
    End With
End Sub

Private Sub SwapBuildDestinies(ByVal plngTree1 As Long, ByVal plngTree2 As Long)
    Dim typSwap As BuildTreeType
    
    typSwap = build.Destiny(plngTree1)
    build.Destiny(plngTree1) = build.Destiny(plngTree2)
    build.Destiny(plngTree2) = typSwap
End Sub


' ************* DISPLAY *************


Private Sub ShowAbilities()
    Dim strCaption As String
    Dim lngCost As Long
    Dim lngRanks As Long
    Dim lngMaxRanks As Long
    Dim lngTotal As Long
    Dim lngSpent() As Long
    Dim blnBlank As Boolean
    Dim lngForceVisible As Long
    Dim i As Long

    If mlngBuildDestiny = 0 Or mlngBuildDestiny > build.Destinies Then
        blnBlank = True
    ElseIf build.Destiny(mlngBuildDestiny).Abilities = 0 Then
        blnBlank = True
    End If
    If blnBlank Then
        Me.usrList.Rows = 1
        SetSlot 1, vbNullString, vbNullString, vbNullString, 0, 0
        Me.usrList.SetError 1, False
        Me.usrList.SetDropState 1, dsDefault
    Else
        Me.usrList.Rows = build.Destiny(mlngBuildDestiny).Abilities
        For i = 1 To build.Destiny(mlngBuildDestiny).Abilities
            If build.Destiny(mlngBuildDestiny).Ability(i).Ability = 0 Then
                Me.usrList.ForceVisible i
                Exit For
            End If
        Next
    End If
    GetSpentInTree db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny), lngSpent, lngTotal
    For i = 1 To build.Destiny(mlngBuildDestiny).Abilities
        If build.Destiny(mlngBuildDestiny).Ability(i).Ability = 0 Then
            SetSlot i, vbNullString, vbNullString, vbNullString, 0, 0
        Else
            GetSlotInfo db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny).Ability(i), strCaption, lngCost, lngRanks, lngMaxRanks
            SetSlot i, build.Destiny(mlngBuildDestiny).Ability(i).Tier, lngCost, strCaption, lngRanks, lngMaxRanks
            Me.usrList.SetError i, CheckErrors(db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny).Ability(i), lngSpent)
        End If
    Next
    With Me.lblSpent
        .Caption = "Spent in Tree: " & lngTotal
        .Visible = True
    End With
    ShowSpentAll Me.lblTotal
End Sub

' Light up valid drop locations during drag operations
Private Sub ShowDropSlots()
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim enDropState As DropStateEnum
    Dim lngSpent() As Long
    Dim typCheck As BuildAbilityType
    Dim i As Long

    If mlngBuildDestiny = 0 Then
        GetUserChoices lngTier, lngAbility, lngSelector
        If lngTier = 0 And lngAbility = 1 Then
            Me.usrList.SetDropState 1, dsCanDrop
        Else
            Me.usrList.SetDropState 1, dsDefault
        End If
    Else
        GetSpentInTree db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny), lngSpent, 0
        With build.Destiny(mlngBuildDestiny)
            For i = 1 To .Abilities
                If .Ability(i).Ability <> 0 Then
                    enDropState = dsDefault
                Else
                    GetUserChoices lngTier, lngAbility, lngSelector
                    With typCheck
                        .Tier = lngTier
                        .Ability = lngAbility
                        .Selector = lngSelector
                        .Rank = 1
                    End With
                    If CheckErrors(db.Destiny(mlngDestiny), typCheck, lngSpent) Then
                        enDropState = dsCanDropError
                    Else
                        enDropState = dsCanDrop
                    End If
                End If
                Me.usrList.SetDropState i, enDropState
                Me.usrList.ForceActive i
            Next
        End With
    End If
End Sub

' Returns TRUE if errors found
Private Function CheckErrors(ptypTree As TreeType, ptypAbility As BuildAbilityType, plngSpent() As Long) As Boolean
    CheckErrors = CheckDestinyAbilityErrors(ptypTree, build.Destiny(mlngBuildDestiny), ptypAbility, plngSpent)
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


Private Function GetUserChoices(plngTier As Long, plngAbility As Long, plngSelector As Long) As Boolean
    Dim lngItemData As Long

    plngTier = 0
    plngAbility = 0
    If plngSelector <> -1 Then plngSelector = 0
    If Me.lstAbility.ListIndex = -1 Then Exit Function
    lngItemData = Me.lstAbility.ItemData(Me.lstAbility.ListIndex)
    SplitAbilityID lngItemData, plngTier, plngAbility, 0
    If plngSelector <> -1 And db.Destiny(mlngDestiny).Tier(plngTier).Ability(plngAbility).SelectorStyle <> sseNone Then
        If Me.lstSub.ListIndex = -1 Then Exit Function
        plngSelector = Me.lstSub.ItemData(Me.lstSub.ListIndex)
    End If
    GetUserChoices = True
End Function

Private Sub StartDrag()
    If mlngBuildDestiny <> 0 Then
        If Not AddAbility(True) Then Exit Sub
    End If
    ShowDropSlots
    If Me.lstSub.ListIndex = -1 Then ListboxDrag Me.lstAbility Else ListboxDrag Me.lstSub
End Sub

Private Sub ListboxDrag(plst As ListBox)
    plst.OLEDropMode = vbOLEDropManual
    plst.OLEDrag
End Sub

Private Function AddAbility(Optional pblnBlank As Boolean = False) As Boolean
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long
    Dim lngInsert As Long
    Dim typBlank As BuildAbilityType
    Dim i As Long

    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Function
    lngInsert = GetInsertionPoint(build.Destiny(mlngBuildDestiny), lngTier, lngAbility)
    If lngAbility Then lngRanks = db.Destiny(mlngDestiny).Tier(lngTier).Ability(lngAbility).Ranks
    With build.Destiny(mlngBuildDestiny)
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
    If Not pblnBlank And lngTier = 5 Then SetDestinyTier5 build.Destiny(mlngBuildDestiny).TreeName
    ShowAbilities
    Me.usrList.Selected = 0
    Me.usrList.Active = lngInsert
    Me.usrList.ForceVisible lngInsert
    AddAbility = True
End Function

Private Sub DropAbility(Index As Integer)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long

    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Sub
    lngRanks = db.Destiny(mlngDestiny).Tier(lngTier).Ability(lngAbility).Ranks
    If mlngBuildDestiny = 0 Then
        With build
            .Destinies = .Destinies + 1
            ReDim Preserve .Destiny(.Destinies)
            mlngBuildDestiny = .Destinies
            With .Destiny(.Destinies)
                .Abilities = 1
                ReDim .Ability(1 To 1)
            End With
        End With
    End If
    With build.Destiny(mlngBuildDestiny).Ability(Index)
        .Tier = lngTier
        .Ability = lngAbility
        .Selector = lngSelector
        .Rank = lngRanks
    End With
End Sub


' ************* TIER 5 *************


Private Sub SetDestinyTier5(pstrDestinyTier5 As String)
    If Len(build.DestinyTier5) = 0 Then
        build.DestinyTier5 = pstrDestinyTier5
    ElseIf build.DestinyTier5 <> pstrDestinyTier5 Then
        ShaveTree build.DestinyTier5
        build.DestinyTier5 = pstrDestinyTier5
    End If
    ShowDestinyTier5
End Sub

Private Sub ShowDestinyTier5()
    Dim blnVisible As Boolean
    
    blnVisible = (Len(build.DestinyTier5) <> 0)
    Me.lblTier5Label.Visible = blnVisible
    Me.lblTier5.Caption = build.DestinyTier5
    Me.lblTier5.Visible = blnVisible
End Sub

Private Function ConfirmDestinyTier5Change() As Boolean
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    ConfirmDestinyTier5Change = True
    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Function
    If lngTier = 5 And Len(build.DestinyTier5) <> 0 And build.DestinyTier5 <> build.Destiny(mlngBuildDestiny).TreeName Then
        If Ask("Make " & build.Destiny(mlngBuildDestiny).TreeName & " your Tier 5 tree?" & vbNewLine & vbNewLine & "(This will clear Tier 5 destinies from the " & build.DestinyTier5 & " tree.)") Then
            SetDestinyTier5 build.Destiny(mlngBuildDestiny).TreeName
        Else
            ConfirmDestinyTier5Change = False
        End If
    End If
End Function


' ************* SLOTS *************

' usrList is the list of selected abilities for this destiny
Private Sub usrList_SlotClick(Index As Integer, Button As Integer)
    Dim typBlank As TwistType
    
    If mlngBuildDestiny = 0 Then Exit Sub
    With Me.lstSub   'Sub selectors
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstAbility  'abilities
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    If Not mblnNoFocus Then Me.usrList.SetFocus
    If Me.usrList.Rows = 0 Then
        Exit Sub
    ElseIf Len(Me.usrList.GetCaption(1)) = 0 Then
        Exit Sub
    ElseIf Me.usrList.Selected = Index Then
        NoSelection
        If Button = vbRightButton Then ClearSlot Index
    Else
        Select Case Button
            Case vbLeftButton
                Me.usrList.Selected = Index
                With build.Destiny(mlngBuildDestiny).Ability(Index)
                    'TODO do we need separate ShowDetails/ShowBuildDetails?
                    ShowDetails .Tier, .Ability, .Selector, Index
                End With
            Case vbRightButton
                If build.Destiny(mlngBuildDestiny).Abilities = 0 Then Exit Sub
                Me.usrList.Selected = Index
                Me.usrList.Active = Index
                With build.Destiny(mlngBuildDestiny).Ability(Index)
                    ShowDetails .Tier, .Ability, .Selector, Index
                End With
                PopupMenu Me.mnuMain(1)
        End Select
    End If
End Sub

Private Sub mnuDestinies_Click(Index As Integer)
    Dim intSlot As Integer
    
    Select Case Index
        Case 0
            intSlot = Me.usrList.Selected
            ClearSlot intSlot
        Case 1
            If Not Ask("Reset " & Me.cboTree.Text & "?") Then Exit Sub
            build.Destiny(mlngBuildDestiny).Abilities = 0
            Erase build.Destiny(mlngBuildDestiny).Ability
            If build.DestinyTier5 = build.Destiny(mlngBuildDestiny).TreeName Then
                SetDestinyTier5 vbNullString
            End If
            ShowAbilities
            ShowAvailable False
            NoSelection
            SetDirty
    End Select
End Sub

Private Sub usrList_SlotDblClick(Index As Integer)
    ClearSlot Index
End Sub

Private Sub ClearSlot(Index As Integer)
    If mlngBuildDestiny = 0 Then Exit Sub
    If build.Destiny(mlngBuildDestiny).Abilities >= Index Then
        build.Destiny(mlngBuildDestiny).Ability(Index).Ability = 0
        RemoveBlanks build.Destiny(mlngBuildDestiny)
        With build.Destiny(mlngBuildDestiny)
            If .Abilities = 0 Then
                SetDestinyTier5 vbNullString
            ElseIf .Ability(.Abilities).Tier < 5 And build.DestinyTier5 = .TreeName Then
                SetDestinyTier5 vbNullString
            End If
        End With
        ShowAbilities
        ShowAvailable True
        NoSelection
        SetDirty
    End If
End Sub

Private Sub usrList_RequestDrag(Index As Integer, Allow As Boolean)
    mlngSourceIndex = Index
    Allow = True
    Me.lstAbility.OLEDropMode = vbOLEDropManual
End Sub

Private Sub usrList_OLEDragDrop(Index As Integer, Data As DataObject)
    If ConfirmDestinyTier5Change() Then
        mblnDragComplete = True
        DropAbility Index
        Me.usrList.SetDropState Index, dsDefault
        ShowAbilities
        Me.usrList.Selected = 0
        Me.usrList.Active = Index
        If Not mblnNoFocus Then Me.usrList.SetFocus
        ShowAvailable True
        SetDirty
    End If
End Sub

Private Sub usrList_RankChange(Index As Integer, Ranks As Long)
    build.Destiny(mlngBuildDestiny).Ability(Index).Rank = Ranks
    ShowAbilities
    ShowAvailable True
    With build.Destiny(mlngBuildDestiny).Ability(Index)
        ShowDetails .Tier, .Ability, .Selector, Index
    End With
    SetDirty
End Sub


' ************* ABILITIES *************


Private Sub lstAbility_Click()
    If mblnMouse Then mblnMouse = False Else ListAbilityClick
End Sub

Private Sub lstAbility_DblClick()
    If Me.lstAbility.ListIndex = -1 Or Me.lstSub.ListCount > 0 Then Exit Sub
    If Not ConfirmDestinyTier5Change() Then Exit Sub
    If AddAbility() Then
        ShowAvailable True
        If Not mblnNoFocus Then Me.usrList.SetFocus
        SetDirty
    End If
End Sub

Private Sub lstAbility_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTier As Long
    Dim lngAbility As Long

    If Button <> vbLeftButton Or Not GetUserChoices(lngTier, lngAbility, -1) Then Exit Sub
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
    Dim lngTier As Long
    Dim lngAbility As Long
    
    ListboxClear Me.lstSub
    If Me.lstAbility.ListIndex = -1 Then Exit Function
    GetUserChoices lngTier, lngAbility, -1
    'TODO do we need separate ShowDetails/ShowBuildDetails?
    ShowDetails lngTier, lngAbility, 0, 0   'Uses the mlngDestiny as tree ID
    If db.Destiny(mlngDestiny).Tier(lngTier).Ability(lngAbility).SelectorStyle <> sseNone Then
        ShowSelectors lngTier, lngAbility
    Else
        ListAbilityClick = True
    End If
End Function

Private Sub lstAbility_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List" ' Me.lstAbility.Text
End Sub

Private Sub lstAbility_OLECompleteDrag(Effect As Long)
    If Not mblnDragComplete Then
        RemoveBlanks build.Destiny(mlngBuildDestiny)
        ShowAbilities
        ShowDropSlots
        Me.usrList.Active = 0
    End If
    mblnDragComplete = False
End Sub

Private Sub lstAbility_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFText) Then
        If Data.GetData(vbCFText) = "List" Then Exit Sub
    End If
    If mlngSourceIndex Then
        ClearSlot CInt(mlngSourceIndex)
        ShowAvailable True
        SetDirty
    End If
End Sub


' ************* SELECTORS *************


Private Sub lstSub_Click()
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    If mblnMouse Then
        mblnMouse = False
    Else
        GetUserChoices lngTier, lngAbility, lngSelector
        ShowDetails lngTier, lngAbility, lngSelector, 0
    End If
End Sub

Private Sub lstSub_DblClick()
    If Me.lstSub.ListIndex = -1 Then Exit Sub
    If Not ConfirmDestinyTier5Change() Then Exit Sub
    If AddAbility() Then
        ShowAvailable True
        If Not mblnNoFocus Then Me.usrList.SetFocus
        SetDirty
    End If
End Sub

Private Sub lstSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long

    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Sub
    Me.usrList.Selected = 0
    ShowDetails lngTier, lngAbility, lngSelector, 0
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
    Data.SetData "List" ' Me.lstSub.Text
End Sub

Private Sub lstSub_OLECompleteDrag(Effect As Long)
    If Not mblnDragComplete Then
        RemoveBlanks build.Destiny(mlngBuildDestiny)
        ShowDropSlots
        Me.usrList.Active = 0
        ShowAbilities
    End If
    mblnDragComplete = False
End Sub


' ************* FILTERS *************


Private Sub usrchkShowAll_UserChange()
    ShowAvailable False
End Sub

Private Sub cboTree_Click()
    If mblnOverride Then Exit Sub
    TreeClick
    SaveBackup
End Sub

Private Sub TreeClick()
    NoSelection
    If Me.cboTree.ListIndex = -1 Then Exit Sub
    mlngDestiny = SeekTree(Me.cboTree.Text, peDestiny)
    mlngBuildDestiny = FindDestinyTree(Me.cboTree.Text)
    mlngMaxTier = GetMaxTier(mlngBuildDestiny)
    Me.usrList.GotoTop
    ShowAbilities
    ShowAvailable False
End Sub

Private Function GetMaxTier(plngBuildTree As Long) As Long
    Dim lngTier As Long
    GetMaxTier = 5
    
    'TODO should be based off of max level + which tier
    'GetMaxTier = CapDestinyTreeTier(build.Destiny(plngBuildTree).TreeName, build.MaxLevels)
End Function

Private Function CapDestinyTreeTier(pstrTreeName As String, ByVal plngLevel As Long) As Long
    Dim lngCap As Long
    ' NEED TO DEAL WITH LEGENDARY HERE
    'Level based
    If plngLevel <= 20 And plngLevel < 22 Then
        lngCap = 2
    ElseIf plngLevel >= 23 And plngLevel < 26 Then
        lngCap = 3
    ElseIf plngLevel >= 26 And plngLevel < 30 Then
        lngCap = 4
    Else
        lngCap = 5
    End If
    
    CapDestinyTreeTier = lngCap
End Function

Private Sub ShaveTree(pstrTreeName As String)
    Dim lngDestinyTree As Long
    Dim lngLast As Long
    Dim i As Long
    
    lngDestinyTree = FindDestinyTree(pstrTreeName)
    If lngDestinyTree = 0 Then Exit Sub
    With build.Destiny(lngDestinyTree)
        For lngLast = .Abilities To 1 Step -1
            If .Ability(lngLast).Tier < 5 Then Exit For
        Next
        If lngLast = 0 Then
            .Abilities = 0
            Erase .Ability
        ElseIf .Abilities <> lngLast Then
            .Abilities = lngLast
            ReDim Preserve .Ability(1 To .Abilities)
        End If
    End With
End Sub

Private Sub ShowAvailable(pblnPreserveTopIndex As Boolean)
    Dim lngTopIndex As Long
    Dim lngSpent() As Long
    Dim typCheck As BuildAbilityType
    Dim lngTier As Long
    Dim i As Long
    
    If pblnPreserveTopIndex Then lngTopIndex = Me.lstAbility.TopIndex
    ListboxClear Me.lstSub
    ListboxClear Me.lstAbility
    If mlngDestiny = 0 Then Exit Sub  'Exit if no current destiny
    'Find out how much has been spent in this destiny - lngSpent is any array based off tier# -
    'contains the max spent
    GetSpentInTree db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny), lngSpent, 0
    For lngTier = 0 To mlngMaxTier
        typCheck.Tier = lngTier
        With db.Destiny(mlngDestiny).Tier(lngTier)
            For i = 1 To .Abilities
                Do  'Ugly loop until ExitDo - I believe this only passes through once no matter what
                    If AbilityTaken(build.Destiny(mlngBuildDestiny), lngTier, i) Then Exit Do
                    typCheck.Ability = i
                    If Not Me.usrchkShowAll.Value Then
                        'See if we need to show this ability
                        If CheckErrors(db.Destiny(mlngDestiny), typCheck, lngSpent) Then Exit Do
                    End If
                    ListboxAddItem Me.lstAbility, lngTier & ": " & .Ability(i).Abbreviation, CreateAbilityID(lngTier, i)
                Loop Until True
            Next
        End With
    Next
    If pblnPreserveTopIndex Then
        With Me.lstAbility
            If lngTopIndex > .ListCount - 1 Then lngTopIndex = .ListCount - 1
            If lngTopIndex <> -1 Then .TopIndex = lngTopIndex
        End With
    End If
End Sub

Private Sub ShowSelectors(plngTier As Long, plngAbility As Long)
    Dim blnSelector() As Boolean
    Dim i As Long

    ListboxClear Me.lstSub
    With db.Destiny(mlngDestiny).Tier(plngTier).Ability(plngAbility)
        GetSelectors db.Destiny(mlngDestiny), plngTier, plngAbility, blnSelector
        For i = 1 To .Selectors
            If blnSelector(i) Then ListboxAddItem Me.lstSub, .Selector(i).SelectorName, i
        Next
    End With
End Sub


' ************* DETAILS *************

'Shows the details of the Destiny/Tier/Ability
'Uses the mlngDestiny as tree ID
Private Sub ShowDetails(ByVal plngTier As Long, ByVal plngAbility As Long, ByVal plngSelector As Long, ByVal plngIndex As Long)
    Dim lngCost As Long
    Dim enReq As ReqGroupEnum
    Dim lngLevels As Long
    Dim lngClassLevels As Long
    Dim lngBuildLevels As Long
    Dim lngSpent() As Long
    Dim lngTotal As Long
    Dim lngProg As Long
    Dim i As Long
    
    ClearDetails False
    With db.Destiny(mlngDestiny).Tier(plngTier).Ability(plngAbility)
        ' Caption
        Me.lblDetails.Caption = "Tier " & plngTier & ": " & .AbilityName
        ' Description
        If Len(.Descrip) Then
            Me.usrDetails.AddDescrip .Descrip, MakeWiki(db.Destiny(mlngDestiny).Wiki) & TierLink(plngTier)
        End If
        'Check for Selector Desc
        If plngSelector > 0 Then
            If Len(.Selector(plngSelector).Descrip) Then
                Me.usrDetails.AddDescrip .Selector(plngSelector).SelectorName, ""
                'See what wiki to use
                If Len(.Selector(plngSelector).Wiki) Then
                    Me.usrDetails.AddDescrip .Selector(plngSelector).Descrip, MakeWiki(.Selector(plngSelector).Wiki) & TierLink(plngTier)
                Else
                    Me.usrDetails.AddDescrip .Selector(plngSelector).Descrip, MakeWiki(db.Destiny(mlngDestiny).Wiki) & TierLink(plngTier)
                End If
            End If
        End If
        
        ' Reqs
        For enReq = rgeAll To rgeNone
            If plngSelector = 0 Then
                ShowDetailsReqs Me.usrDetails, mlngDestiny, .Req(enReq), enReq, 0
            Else
                ShowDetailsReqs Me.usrDetails, mlngDestiny, .Selector(plngSelector).Req(enReq), enReq, 0
            End If
        Next
        ' Rank reqs
        If plngSelector = 0 Then
            ShowRankReqs Me.usrDetails, mlngDestiny, .RankReqs, .Rank
        Else
            ShowRankReqs Me.usrDetails, mlngDestiny, .Selector(plngSelector).RankReqs, .Selector(plngSelector).Rank
        End If
        ' Levels
        'TODO fix for destiny
        GetLevelReqs db.Destiny(mlngDestiny).TreeType, plngTier, plngAbility, lngLevels, lngClassLevels
        'lngBuildLevels = GetBuildLevelReq(lngLevels, lngClassLevels, GetClass(mlngTree))
        'If lngBuildLevels > 1 Or lngBuildLevels = -1 Then
        '    Me.usrDetails.AddText "Level requirements:"
        '    If lngClassLevels Then Me.usrDetails.AddText " - " & lngClassLevels & " class levels"
        '    If lngLevels Then Me.usrDetails.AddText " - " & lngLevels & " character levels"
        '    If lngBuildLevels = -1 Then
        '        Me.usrDetails.AddText " - Unattainable for this build"
        '    ElseIf lngBuildLevels > 1 Then
        '        Me.usrDetails.AddText " - " & lngBuildLevels & " build levels"
        '    End If
        'End If
        
        ' Error?
        If plngIndex <> 0 Then
            gstrError = vbNullString
            GetSpentInTree db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny), lngSpent, lngTotal
            If CheckErrors(db.Destiny(mlngDestiny), build.Destiny(mlngBuildDestiny).Ability(plngIndex), lngSpent) Then
                Me.usrDetails.AddErrorText "Error: " & gstrError
            End If
        End If
        Me.usrDetails.Refresh
        ' Ranks
        Me.lblRanks.Caption = "Ranks: " & .Ranks
        Me.lblRanks.Visible = True
        ' Cost
        Me.lblCost.Caption = CostDescrip(db.Destiny(mlngDestiny).Tier(plngTier).Ability(plngAbility), plngSelector)
        Me.lblCost.Visible = True
        ' Spent in tree
        lngProg = GetSpentReq(db.Destiny(mlngDestiny).TreeType, plngTier, plngAbility)
        If lngProg Then
            Me.lblProg.Caption = lngProg & " AP spent in tree"
            Me.lblProg.Visible = True
        Else
            Me.lblProg.Visible = False
        End If
    End With
End Sub

Private Function GetClass(plngTree As Long) As ClassEnum
    Dim typClassSplit() As ClassSplitType
    Dim lngClass As Long
    Dim lngLevels As Long
    Dim enClass As ClassEnum
    Dim i As Long
    
    For lngClass = 0 To GetClassSplit(typClassSplit) - 1
        enClass = typClassSplit(lngClass).ClassID
        With db.Class(enClass)
            For i = 1 To .Trees
                If .Tree(i) = db.Tree(plngTree).TreeName Then
                    If lngLevels < typClassSplit(lngClass).Levels Then
                        lngLevels = typClassSplit(lngClass).Levels
                        GetClass = enClass
                    End If
                    Exit For
                End If
            Next
        End With
    Next
End Function

Private Sub ShowDetailsReqs(pusrdet As userDetails, plngTree As Long, ptypReqList As ReqListType, penGroup As ReqGroupEnum, plngRank As Long)
    Dim strText As String
    Dim i As Long
    
    If ptypReqList.Reqs = 0 Then Exit Sub
    If plngRank < 2 Then strText = "Requires " Else strText = "Rank " & plngRank & " requires "
    strText = strText & LCase$(GetReqGroupName(penGroup)) & " of:"
    If plngRank = 0 Then
        pusrdet.AddText "Requires " & LCase$(GetReqGroupName(penGroup)) & " of:"
    Else
        pusrdet.AddText "Rank " & plngRank & " requires " & LCase$(GetReqGroupName(penGroup)) & " of:"
    End If
    For i = 1 To ptypReqList.Reqs
        pusrdet.AddText " - " & PointerDisplay(ptypReqList.Req(i), True, plngTree)
    Next
End Sub

Private Sub ShowRankReqs(pusrdet As userDetails, plngTree As Long, pblnRankReqs As Boolean, ptypRank() As RankType)
    Dim enReq As ReqGroupEnum
    Dim lngRank As Long
    Dim i As Long
    
    If Not pblnRankReqs Then Exit Sub
    For lngRank = 2 To 3
        With ptypRank(lngRank)
            ' Class
            If .Class(0) Then
                pusrdet.AddText "Rank " & lngRank & " requires Class:"
                For i = 1 To ceClasses - 1
                    If .Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
                Next
            End If
            ' Reqs
            For enReq = rgeAll To rgeNone
                ShowDetailsReqs pusrdet, plngTree, .Req(enReq), enReq, lngRank
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
    If mlngBuildDestiny = 0 Then Exit Sub
    With Me.lstSub
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstAbility
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.Selected = 0
    ClearDetails True
    On Error Resume Next
    If Not mblnNoFocus Then Me.usrList.SetFocus
End Sub

Private Sub Form_Click()
'    If mlngTab = 1 Then NoSelection
End Sub

Private Sub picTab_Click(Index As Integer)
    If Index = 1 Then NoSelection
End Sub

Private Sub picTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Index = 2 Then ActiveCell -1, -1
End Sub


Private Function ReadMouseShift(Shift As Integer) As MouseShiftEnum
    If (Shift And vbCtrlMask) > 0 Then
        ReadMouseShift = mseCtrl
    ElseIf (Shift And vbShiftMask) > 0 Then
        ReadMouseShift = mseShift
    Else
        ReadMouseShift = mseNormal
    End If
End Function

Private Sub GetCoords(plngRow As Long, plngCol As Long, plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long)
    plngLeft = Col(plngCol).Left
    plngTop = (plngRow - 1) * mlngHeight
    plngRight = Col(plngCol).Right
    plngBottom = plngTop + mlngHeight
End Sub





