VERSION 5.00
Begin VB.Form frmCompendium 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Compendium"
   ClientHeight    =   6096
   ClientLeft      =   132
   ClientTop       =   180
   ClientWidth     =   12312
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompendium.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6096
   ScaleWidth      =   12312
   Begin Compendium.userToolbar usrToolbar 
      Height          =   372
      Left            =   3540
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   8772
      _ExtentX        =   15473
      _ExtentY        =   656
      Captions        =   "Home,XP,Wilderness,Notes,Links"
   End
   Begin Compendium.userLinkLists usrLinks 
      Height          =   2352
      Left            =   7860
      TabIndex        =   6
      Top             =   3180
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   4149
   End
   Begin Compendium.userTextbox usrtxtNotes 
      Height          =   2352
      Left            =   7260
      TabIndex        =   1
      Top             =   2580
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   4149
   End
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   300
      Top             =   120
   End
   Begin Compendium.userAreas usrWilderness 
      Height          =   2352
      Left            =   6660
      TabIndex        =   3
      Top             =   1980
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   4149
   End
   Begin Compendium.userTables usrXP 
      Height          =   2352
      Left            =   6060
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   4149
   End
   Begin Compendium.userOverview usrOverview 
      Height          =   2352
      Left            =   5460
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   4149
   End
   Begin VB.Timer tmrAutoSave 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   720
      Top             =   120
   End
   Begin Compendium.userQuests usrQuests 
      Height          =   3732
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   4332
      _ExtentX        =   7641
      _ExtentY        =   6583
   End
   Begin VB.Menu mnuMain 
      Caption         =   "LinkContext"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuLinkContext 
         Caption         =   "New List"
         Index           =   0
      End
      Begin VB.Menu mnuLinkContext 
         Caption         =   "Undelete"
         Index           =   1
         Visible         =   0   'False
         Begin VB.Menu mnuUndelete 
            Caption         =   "Item 1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuLinkContext 
         Caption         =   "-"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCompendium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RectangleType
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type WindowType
    Left As RectangleType
    Right As RectangleType
    Unified As RectangleType
    Separate As Boolean
End Type

Private win As WindowType

Private mblnOverride As Boolean
Private mblnLoaded As Boolean
Private mblnInitialized As Boolean

Private mlngClientTop As Long


' ************* FORM *************


' Empty form is immediately displayed; controls loaded by timer
Private Sub Form_Load()
    mblnInitialized = False
    SetCaption
    Me.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    cfg.SizeWindow
    If Not xp.DebugMode Then Call WheelHook(Me.Hwnd)
    Me.tmrLoad.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.RunBefore = True
    If Not xp.DebugMode Then Call WheelUnHook(Me.Hwnd)
    If mblnLoaded Then cfg.SaveWindowSize
    CloseApp
End Sub

Private Sub Form_Resize()
    If Not mblnInitialized Then Exit Sub
    CalculateWindows
    MoveControls
    If mblnLoaded Then DirtyFlag dfeSettings
End Sub

Public Sub SetCaption()
    Dim strCaption As String
    
    strCaption = cfg.DataFile
    If strCaption = "Main" Then strCaption = "Compendium"
    If gblnDirtyFlag(dfeAny) Then strCaption = strCaption & "*"
    If Me.Caption <> strCaption Then Me.Caption = strCaption
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    Dim lngStep As Long
    
    lngStep = cfg.WheelStep
    If lngStep = 0 Then lngStep = Me.usrQuests.PageSize
    If lngStep < 1 Then lngStep = 3
    If Rotation < 0 Then lngValue = -lngStep Else lngValue = lngStep
    ' Check only active tabs
    Select Case True
        Case PaneActive(peQuests) And IsOver(Me.usrQuests.Hwnd, Xpos, Ypos): Me.usrQuests.WheelScroll lngValue
        Case PaneActive(peXP) And IsOver(Me.usrXP.Hwnd, Xpos, Ypos): Me.usrXP.WheelScroll lngValue
        Case PaneActive(peWilderness) And IsOver(Me.usrWilderness.Hwnd, Xpos, Ypos): Me.usrWilderness.WheelScroll lngValue
    End Select
End Sub

Private Function PaneActive(penPane As PaneEnum) As Boolean
    Dim enLeft As PaneEnum
    
    If win.Separate Then enLeft = peQuests Else enLeft = cfg.LeftPane
    If penPane = enLeft Or penPane = cfg.RightPane Then PaneActive = True
End Function


' ************* LOAD *************


Private Sub tmrLoad_Timer()
    Me.tmrLoad.Enabled = False
    If Not xp.DebugMode Then xp.LockWindow Me.Hwnd
    ' Colors
    Me.usrToolbar.RefreshColors
    cfg.ApplyColors Me.usrtxtNotes, cgeControls
    cfg.ApplyColors Me.usrXP, cgeWorkspace
    cfg.ApplyColors Me.usrWilderness, cgeWorkspace
    Me.usrLinks.RefreshColors
    ' Initialize data
    Me.usrToolbar.Init
    Me.usrQuests.Init
    Me.usrWilderness.Init
    Me.usrXP.Init
    Me.usrLinks.Init
    mblnInitialized = True
    ' Size and move controls
    CalculateWindows
    MoveControls
    ' Load data
'~    Me.usrOverview.Init
    Me.usrtxtNotes.Text = LoadNotes()
    Me.usrLinks.ShowLinkLists
    Me.usrQuests.Show
    DefineTabs
    ShowTab
    Me.usrToolbar.Visible = True
    mblnLoaded = True
    If Not xp.DebugMode Then xp.UnlockWindow
    If Not cfg.RunBefore Then ShowHelp "Getting_Started"
    cfg.MessageShow
End Sub

Private Sub tmrAutoSave_Timer()
    Me.tmrAutoSave.Enabled = False
    AutoSave
End Sub


' ************* PUBLIC *************


Public Sub RefreshColors()
    If Not xp.DebugMode Then xp.LockWindow Me.Hwnd
    cfg.CompendiumScroll = Me.usrQuests.Scroll
    If cfg.CompendiumBackColor <> cfg.GetColor(cgeControls, cveBackground) Then MatchColors True
    Me.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    Me.usrToolbar.RefreshColors
    Me.usrQuests.Redraw
    Me.usrQuests.Scroll = cfg.CompendiumScroll
'~    Me.usrOverview.RefreshColors
    Me.usrXP.Redraw
    Me.usrWilderness.Redraw
    cfg.ApplyColors Me.usrtxtNotes, cgeControls
    Me.usrLinks.RefreshColors
'~    Me.usrOverview.RefreshColors
    If Not xp.DebugMode Then xp.UnlockWindow
End Sub

Public Sub RedrawQuests()
    xp.Mouse = msAppWait
    If Not xp.DebugMode Then xp.LockWindow Me.Hwnd
    cfg.CompendiumScroll = Me.usrQuests.Scroll
    Me.usrQuests.ReQuery
    CalculateWindows
    MoveControls
    Me.usrQuests.Scroll = cfg.CompendiumScroll
    If Not xp.DebugMode Then xp.UnlockWindow
    xp.Mouse = msNormal
End Sub

Public Sub CharacterChanged()
'~    Me.usrOverview.CharacterChanged
End Sub

Public Sub FavorChange(plngCharacter As Long)
    Me.usrQuests.FavorChange plngCharacter
End Sub

Public Sub ReloadLinkLists()
    LoadLinkLists
    Me.usrLinks.ShowLinkLists
End Sub

Public Sub Notes()
    Me.usrtxtNotes.Text = LoadNotes()
End Sub

Public Sub ShowPlayButton()
    Me.usrToolbar.ShowPlayButton
End Sub

Public Sub GetQuestsFont(pstrName As String, pdblSize As Double)
    Me.usrQuests.GetFont pstrName, pdblSize
End Sub

Public Function SetQuestsFont(pstrName As String, Optional pdblSize As Double) As Double
    SetQuestsFont = Me.usrQuests.SetFont(pstrName, pdblSize)
    RedrawQuests
End Function

Public Sub SaveLinkList(plngIndex As Long)
    Me.usrLinks.SaveData plngIndex
End Sub

Public Sub GetMenuCoords(plngIndex As Long, plngLeft As Long, plngTop As Long)
    Me.usrLinks.GetMenuCoords plngIndex, plngLeft, plngTop
End Sub


' ************* WINDOWS *************


Private Sub CalculateWindows()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngSides As Long
    Dim lngTotalHeight As Long
    Dim lngToolbarHeight As Long
    Dim blnOld As Boolean
    
    lngSides = Me.ScaleX(cfg.Sides, vbPixels, vbTwips)
    lngTotalHeight = Me.ScaleHeight - Me.ScaleY(cfg.Bottom, vbPixels, vbTwips) - lngTop
    lngToolbarHeight = Me.usrToolbar.FitHeight
    ' Dimensions for unified window
    lngLeft = lngSides
    lngWidth = Me.ScaleWidth - (lngSides * 2)
    lngTop = lngToolbarHeight - PixelY
    lngHeight = lngTotalHeight - lngTop
    SetRectangle win.Unified, lngLeft, lngTop, lngWidth, lngHeight
    ' Left window
    lngWidth = Me.usrQuests.FitWidth
    lngTop = 0
    lngHeight = lngTotalHeight
    If lngWidth > win.Unified.Width Then lngWidth = win.Unified.Width
    SetRectangle win.Left, lngLeft, lngTop, lngWidth, lngHeight
    ' Right window
    lngLeft = win.Left.Left + win.Left.Width
    lngWidth = win.Unified.Width - win.Left.Width
    lngTop = lngToolbarHeight - PixelY
    lngHeight = lngTotalHeight - lngTop
    SetRectangle win.Right, lngLeft, lngTop, lngWidth, lngHeight
    ' Unified?
    blnOld = win.Separate
    win.Separate = (win.Right.Width >= Me.usrToolbar.FitWidth)
    If win.Separate <> blnOld Then UnifiedChange
End Sub

Private Sub SetRectangle(ptypRectangle As RectangleType, plngLeft As Long, plngTop As Long, plngWidth As Long, plngHeight As Long)
    With ptypRectangle
        .Left = plngLeft
        .Top = plngTop
        .Width = plngWidth
        .Height = plngHeight
    End With
End Sub

Private Sub UnifiedChange()
    DefineTabs
    If win.Separate Then
        If cfg.LeftPane <> peQuests Then cfg.RightPane = cfg.LeftPane
        Me.usrQuests.Visible = True
    Else
        cfg.LeftPane = peQuests
    End If
    DefineTabs
    ShowTab
End Sub

Private Sub MoveControls()
    If Me.WindowState = vbMinimized Then Exit Sub
    If win.Separate Then
        With win.Left
            Me.usrQuests.Move .Left, .Top, .Width, .Height
        End With
        With win.Right
            Me.usrToolbar.Move .Left, 0, .Width, Me.usrToolbar.FitHeight
            Me.usrOverview.Move .Left, .Top, .Width, .Height
            Me.usrXP.Move .Left, .Top, .Width, .Height
            Me.usrWilderness.Move .Left, .Top, .Width, .Height
            Me.usrtxtNotes.Move .Left, .Top, .Width, .Height
            Me.usrLinks.Move .Left, .Top, .Width, .Height
        End With
    Else
        With win.Unified
            Me.usrToolbar.Move .Left, 0, .Width, Me.usrToolbar.FitHeight
            Me.usrQuests.Move .Left, .Top, .Width, .Height
            Me.usrOverview.Move .Left, .Top, .Width, .Height
            Me.usrXP.Move .Left, .Top, .Width, .Height
            Me.usrWilderness.Move .Left, .Top, .Width, .Height
            Me.usrtxtNotes.Move .Left, .Top, .Width, .Height
            Me.usrLinks.Move .Left, .Top, .Width, .Height
        End With
    End If
End Sub


' ************* TOOLBAR *************


Private Sub DefineTabs()
    Const DefaultCaptions As String = "Notes,Links,XP,Wilderness" ' "Home,XP,Wilderness,Notes,Links"
    Dim strCaptions As String
    Dim strActive As String
    
    If win.Separate Then
        strActive = GetPaneName(cfg.RightPane)
        strCaptions = DefaultCaptions
    Else
        strActive = GetPaneName(cfg.LeftPane)
        strCaptions = "Quests," & DefaultCaptions
    End If
    Me.usrToolbar.BulkChange strCaptions, strActive, TabColor(strActive)
End Sub

Private Function TabColor(pstrCaption As String) As Long
    Select Case pstrCaption
        Case "Quests", "Wilderness": TabColor = cfg.GetColor(cgeDropSlots, cveBackground)
        Case "XP": TabColor = cfg.GetColor(cgeWorkspace, cveBackground)
        Case Else: TabColor = cfg.GetColor(cgeControls, cveBackground)
    End Select
End Function

Private Sub usrToolbar_TabClick(Caption As String)
    Dim lngColor As Long
    Dim lngTop As Long
    Dim enPane As PaneEnum

    If mblnOverride Then Exit Sub
    Me.usrToolbar.TabActiveColor = TabColor(Caption)
    enPane = GetPaneID(Caption)
    If win.Separate Then
        cfg.LeftPane = peQuests
        cfg.RightPane = enPane
    Else
        cfg.LeftPane = enPane
        If enPane <> peQuests Then cfg.RightPane = enPane
    End If
    ShowTab
    DirtyFlag dfeSettings
End Sub

Private Sub ShowTab()
    Dim enPane As PaneEnum
    Dim i As Long
    
    If win.Separate Then enPane = cfg.RightPane Else enPane = cfg.LeftPane
    For i = 0 To pePaneCount - 1
        If i <> enPane Then TabVisible i, False
    Next
    TabVisible enPane, True
    If enPane = peNotes And Me.usrtxtNotes.Visible = True And mblnLoaded Then Me.usrtxtNotes.SetFocus
End Sub

Private Sub TabVisible(penPane As PaneEnum, ByVal pblnVisible As Boolean)
    Select Case penPane
        Case peQuests
            If win.Separate Then pblnVisible = True
            If Me.usrQuests.Visible <> pblnVisible Then Me.usrQuests.Visible = pblnVisible
        Case peHome
'~            If Me.usrOverview.Visible <> pblnVisible Then Me.usrOverview.Visible = pblnVisible
        Case peNotes
            If Me.usrtxtNotes.Visible <> pblnVisible Then Me.usrtxtNotes.Visible = pblnVisible
        Case peLinks
            If Me.usrLinks.Visible <> pblnVisible Then Me.usrLinks.Visible = pblnVisible
        Case peWilderness
            If Me.usrWilderness.Visible <> pblnVisible Then Me.usrWilderness.Visible = pblnVisible
        Case peXP
            If Me.usrXP.Visible <> pblnVisible Then Me.usrXP.Visible = pblnVisible
    End Select
End Sub

Private Sub usrToolbar_ButtonClick(Caption As String)
    Select Case Caption
        Case "Play": PlayGame
        Case "Tools": OpenForm "frmTools"
        Case "Help": ShowHelp "Table_of_Contents"
    End Select
End Sub

Private Sub PlayGame()
    If Len(cfg.PlayEXE) = 0 Then
        If MsgBox("No file to run has been specified. Choose a file now?", vbYesNo + vbQuestion, "Notice") = vbYes Then
            If ChooseEXE() Then PlayGame
        End If
    ElseIf Not xp.File.Exists(cfg.PlayEXE) Then
        MsgBox "File not found: " & vbNewLine & vbNewLine & cfg.PlayEXE, vbInformation, "Notice"
    Else
        xp.File.Run cfg.PlayEXE
    End If
End Sub


' ************* DEBUG TOOLS *************


Public Function GetTextWidth(pstrText As String) As Long
    GetTextWidth = Me.usrQuests.GetTextWidth(pstrText)
End Function
