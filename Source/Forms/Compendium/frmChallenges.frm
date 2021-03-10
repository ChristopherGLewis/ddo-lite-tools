VERSION 5.00
Begin VB.Form frmChallenges 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Challenges"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChallenges.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   10155
   Begin Compendium.userCheckBox usrchkOrder 
      Height          =   312
      Left            =   3060
      TabIndex        =   13
      Tag             =   "nav"
      Top             =   30
      Width           =   2412
      _ExtentX        =   4260
      _ExtentY        =   556
      Caption         =   "Match In-Game Order"
   End
   Begin VB.ComboBox cboCharacter 
      Height          =   312
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "nav"
      Top             =   30
      Width           =   1752
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   5280
      Picture         =   "frmChallenges.frx":08CA
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   5280
      Picture         =   "frmChallenges.frx":1002
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   10
      Top             =   1824
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   5280
      Picture         =   "frmChallenges.frx":17F3
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   9
      Top             =   2088
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   5280
      Picture         =   "frmChallenges.frx":2027
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   8
      Top             =   2352
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   4
      Left            =   5280
      Picture         =   "frmChallenges.frx":285E
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   7
      Top             =   2616
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   5
      Left            =   5280
      Picture         =   "frmChallenges.frx":3064
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picStars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   6
      Left            =   5280
      Picture         =   "frmChallenges.frx":37A8
      ScaleHeight     =   330
      ScaleWidth      =   2310
      TabIndex        =   5
      Top             =   3144
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4332
      Left            =   0
      ScaleHeight     =   4335
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   720
      Width           =   5052
      Begin VB.VScrollBar scrollVertical 
         Height          =   1932
         Left            =   4620
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3792
         Left            =   0
         ScaleHeight     =   3795
         ScaleWidth      =   4575
         TabIndex        =   1
         Tag             =   "wrk"
         Top             =   0
         Width           =   4572
         Begin VB.PictureBox picProgress 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   264
            Index           =   0
            Left            =   1920
            ScaleHeight     =   270
            ScaleWidth      =   1845
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   1848
         End
         Begin VB.Label lnkChallenge 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Challenge Name"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   540
            TabIndex        =   2
            Top             =   420
            Visible         =   0   'False
            Width           =   1440
         End
      End
   End
   Begin Compendium.userHeader usrhdrHeader 
      Height          =   372
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10152
      _ExtentX        =   17912
      _ExtentY        =   661
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
End
Attribute VB_Name = "frmChallenges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMarginX As Long
Private mlngMarginY As Long

Private Enum ColumnEnum
    ceChallenge = 1
    ceGroup
    ceLevel
    ceStars
    cePatron
    ceFavor
    ceColumnCount
End Enum

Private Type ColumType
    Header As String
    Left As Long
    Right As Long
    Width As Long
    Link As Boolean
End Type

Private Type RowType
    Challenge As String
    ChallengeLeft As Long
    ChallengeRight As Long
    Wiki As String
    Group As String
    Patron As String
    LevelLow As Long
    LevelHigh As Long
    Stars As Long
    Index As Long
End Type

Private mtypRow() As RowType
Private mtypCol() As ColumType

Private mlngRow As Long
Private mlngRowHeight As Long
Private mlngTextOffset As Long

Private mlngCharacter As Long

Private mlngStars As Long
Private mlngIndex As Long

Private mlngTextHeight As Long
Private mlngActiveRow As Long

Private mblnOverride As Boolean
Private mblnDirty As Boolean


' ************* FORM *************


Private Sub Form_Load()
    SizeClient
    LoadData
    cfg.Configure Me
    If Not xp.DebugMode Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.Hwnd)
    cfg.SavePosition Me, True
End Sub

Private Sub Form_Resize()
    Dim lngMax As Long
    
    If Me.WindowState = vbMinimized Or mlngRowHeight = 0 Then Exit Sub
    If Me.ScaleHeight < Me.usrhdrHeader.Height + (mlngRowHeight * 3) Then Exit Sub
    Me.usrhdrHeader.Width = Me.ScaleWidth - PixelX
    Me.cboCharacter.Left = mlngMarginX
    Me.usrchkOrder.Width = Me.usrchkOrder.FitWidth
    Me.usrchkOrder.Left = mlngMarginX + Me.cboCharacter.Width + (Me.ScaleWidth - Me.usrhdrHeader.Margin - Me.TextWidth("Help") - (mlngMarginX + Me.cboCharacter.Width) - Me.usrchkOrder.FitWidth) \ 2
    With Me.picContainer
        .Move 0, Me.usrhdrHeader.Height + PixelY * 2, Me.ScaleWidth, Me.ScaleHeight - (Me.usrhdrHeader.Height + PixelY * 2)
    End With
    lngMax = (db.Challenges + 1) - (Me.picContainer.Height \ mlngRowHeight)
    If lngMax > 0 And Me.picClient.Height > Me.picContainer.Height Then
        With Me.scrollVertical
            .Move Me.picContainer.ScaleWidth - .Width, 0, .Width, Me.picContainer.ScaleHeight
            .Min = 0
            .Max = lngMax
            .Value = 0
            .LargeChange = Me.picContainer.Height \ mlngRowHeight
            .Visible = True
        End With
    Else
        Me.picClient.Top = 0
        Me.scrollVertical.Visible = False
    End If
End Sub

Public Property Get Character() As Long
    Character = mlngCharacter
End Property

Public Property Let Character(ByVal plngCharacter As Long)
    mlngCharacter = plngCharacter
    ComboSetValue Me.cboCharacter, mlngCharacter
End Property

Public Sub CharacterListChanged()
    InitCombo True
End Sub

Public Sub DataFileChanged()
    InitCombo False
End Sub

Public Sub ReDrawForm()
    LoadData
    DrawGrid
End Sub

Public Sub ReQueryData(plngCharacter As Long)
    If plngCharacter = 0 Or plngCharacter <> mlngCharacter Then Exit Sub
    SortTable
    DrawGrid
End Sub

Private Sub picStars_Click(Index As Integer)
     MsgBox ("here")
End Sub

Private Sub usrhdrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Challenges"
    End Select
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.picContainer.Hwnd, Xpos, Ypos): WheelScroll lngValue
    End Select
End Sub


' ************* DATA *************


Private Sub LoadData()
    If db.Challenges = 0 Then Exit Sub
    Me.usrchkOrder.Value = (cfg.ChallengeOrder = coeGame)
    SortTable
    InitCombo False
End Sub

Private Sub InitCombo(pblnFindCharacter As Boolean)
    Dim strCharacter As String
    Dim lngIndex As Long
    Dim i As Long
    
    strCharacter = ComboGetText(Me.cboCharacter)
    mblnOverride = True
    ComboClear Me.cboCharacter
    ComboAddItem Me.cboCharacter, vbNullString, 0
    For i = 1 To db.Characters
        ComboAddItem Me.cboCharacter, db.Character(i).Character, i
    Next
    mblnOverride = False
    If pblnFindCharacter Then lngIndex = ComboFindText(Me.cboCharacter, strCharacter) Else lngIndex = 0
    If lngIndex = -1 Then lngIndex = 0
    Me.cboCharacter.ListIndex = lngIndex
End Sub

Private Sub SortTable()
    Dim lngIndex As Long
    Dim i As Long
    
    ReDim mtypRow(1 To db.Challenges)
    For i = 1 To db.Challenges
        Select Case cfg.ChallengeOrder
            Case coeChallenge: lngIndex = i
            Case coeGame: lngIndex = db.Challenge(i).GameOrder
            Case coeGroup: lngIndex = db.Challenge(i).GroupOrder
        End Select
        With mtypRow(lngIndex)
            .Index = i
            .Challenge = db.Challenge(i).Challenge
            .Wiki = db.Challenge(i).Wiki
            .Group = db.Challenge(i).Group
            .Patron = db.Challenge(i).Patron
            .LevelLow = db.Challenge(i).LevelLow
            .LevelHigh = db.Challenge(i).LevelHigh
            If mlngCharacter = 0 Then .Stars = 0 Else .Stars = db.Challenge(i).Stars(mlngCharacter)
        End With
    Next
End Sub


' ************* SIZING *************


Private Sub SizeClient()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngClientHeight As Long
    Dim lngFormHeight As Long
    Dim lngMaxHeight As Long
    Dim i As Long
    
    mlngTextHeight = Me.picClient.TextHeight("Q")
    mlngMarginX = Me.picClient.TextWidth("  ")
    mlngMarginY = PixelY * 4
    ReDim mtypCol(1 To 6)
    With Me.picClient
        For i = 1 To db.Challenges
            CheckWidth ceChallenge, .TextWidth(db.Challenge(i).Challenge)
            CheckWidth ceGroup, .TextWidth(db.Challenge(i).Group)
        Next
        mtypCol(ceLevel).Width = .TextWidth("15-20")
        mtypCol(ceStars).Width = Me.picStars(0).Width
        mtypCol(cePatron).Width = .TextWidth("Purple Dragon Knights")
        mtypCol(ceFavor).Width = .TextWidth("Favor")
    End With
    For i = 1 To ceColumnCount - 1
        With mtypCol(i)
            .Left = lngLeft
            If i <> ceStars And i <> cePatron Then .Width = .Width + mlngMarginX * 2
            .Right = lngLeft + .Width
            lngLeft = .Right
        End With
    Next
    mlngRowHeight = Me.picStars(0).Height + mlngMarginY * 2
    mlngTextOffset = (mlngRowHeight - mlngTextHeight) \ 2
    ' Client
    lngWidth = mtypCol(6).Right + mlngMarginX
    lngClientHeight = Me.usrhdrHeader.Height + mlngRowHeight * db.Challenges
    Me.picClient.Move 0, 0, lngWidth, lngClientHeight
    ' Container
    lngTop = Me.usrhdrHeader.Height + PixelY * 2
    lngFormHeight = lngTop + lngClientHeight + Me.Height - Me.ScaleHeight
    xp.GetDesktop 0, 0, 0, lngMaxHeight
    If lngFormHeight > lngMaxHeight Then lngFormHeight = lngMaxHeight
    lngHeight = lngFormHeight - (Me.Height - Me.ScaleHeight) - lngTop
    Me.picContainer.Move 0, lngTop, Me.picClient.Width + Me.scrollVertical.Width + PixelX, lngHeight
    ' Form
    If Me.picContainer.ScaleHeight < Me.picClient.Height Then
        lngWidth = Me.picContainer.Width + Me.Width - Me.ScaleWidth
    Else
        lngWidth = Me.picClient.Width + Me.Width - Me.ScaleWidth
    End If
    Me.Move Me.Left, Me.Top, lngWidth, lngFormHeight
End Sub

Private Sub CheckWidth(penColumn As ColumnEnum, plngCheck As Long)
    If mtypCol(penColumn).Width < plngCheck Then mtypCol(penColumn).Width = plngCheck
End Sub


' ************* DRAWING *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim strDisplay As String
    Dim blnLine As Boolean
    
    Me.picClient.Cls
    For lngRow = 1 To db.Challenges
        blnLine = False
        Select Case cfg.ChallengeOrder
            Case coeGroup
                If lngRow < db.Challenges Then
                    If mtypRow(lngRow).Group <> mtypRow(lngRow + 1).Group And mtypRow(lngRow).Patron = "House Cannith" Then blnLine = True
                End If
            Case coeGame
                If lngRow Mod 11 = 0 Then blnLine = True
        End Select
        If blnLine Then
            lngTop = lngRow * mlngRowHeight
            Me.picClient.Line (mtypCol(1).Left, lngTop)-(mtypCol(6).Right, lngTop), cfg.GetColor(cgeWorkspace, cveTextDim)
        End If
        For lngCol = 1 To ceColumnCount - 1
            DrawCell lngRow, lngCol, False, False
        Next
    Next
    ShowTotal
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, pblnActive As Boolean, pblnClear As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim strDisplay As String
    
    lngLeft = mtypCol(plngCol).Left
    lngTop = (plngRow - 1) * mlngRowHeight
    With mtypRow(plngRow)
        Select Case plngCol
            Case ceChallenge: strDisplay = .Challenge
            Case ceGroup: strDisplay = .Group
            Case ceLevel: strDisplay = .LevelLow & "-" & .LevelHigh
            Case ceStars: strDisplay = vbNullString
            Case cePatron: strDisplay = .Patron
            Case ceFavor: strDisplay = .Stars
        End Select
    End With
    Select Case plngCol
        Case ceStars
            If plngRow > Me.picProgress.UBound Then Load Me.picProgress(plngRow)
            Me.picProgress(plngRow).Move lngLeft, lngTop + mlngMarginY
            ShowStars plngRow
            Me.picProgress(plngRow).Visible = True
            Exit Sub
        Case ceLevel
            lngLeft = mtypCol(plngCol).Right - Me.picClient.TextWidth(strDisplay)
        Case ceFavor
            lngLeft = lngLeft + (mtypCol(plngCol).Width - Me.picClient.TextWidth(strDisplay)) \ 2
        Case cePatron
        Case Else
            lngLeft = lngLeft + mlngMarginX
    End Select
    If pblnClear Then Me.picClient.Line (lngLeft, lngTop + PixelY)-(mtypCol(plngCol).Right, lngTop + mlngRowHeight - PixelY), cfg.GetColor(cgeWorkspace, cveBackground), BF
    If pblnActive Then Me.picClient.ForeColor = cfg.GetColor(cgeWorkspace, cveTextLink)
    Me.picClient.CurrentX = lngLeft
    Me.picClient.CurrentY = lngTop + mlngTextOffset
    Me.picClient.Print strDisplay;
    If plngCol = 1 Then
        mtypRow(plngRow).ChallengeLeft = mtypCol(plngCol).Left
        mtypRow(plngRow).ChallengeRight = Me.picClient.CurrentX
    End If
    If pblnActive Then Me.picClient.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
End Sub

Private Sub ShowTotal()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim strDisplay As String
    
    lngLeft = mtypCol(6).Left
    lngTop = db.Challenges * mlngRowHeight
    Me.picClient.Line (lngLeft, lngTop)-(mtypCol(6).Right, lngTop + mlngRowHeight), cfg.GetColor(cgeWorkspace, cveBackground), BF
    If mlngCharacter = 0 Then Exit Sub
    Me.picClient.Line (lngLeft, lngTop)-(mtypCol(6).Right, lngTop), cfg.GetColor(cgeWorkspace, cveTextDim)
    strDisplay = db.Character(mlngCharacter).ChallengeFavor
    lngLeft = lngLeft + (mtypCol(6).Width - Me.picClient.TextWidth(strDisplay)) \ 2
    Me.picClient.CurrentX = lngLeft
    Me.picClient.CurrentY = lngTop + PixelY * 2
    Me.picClient.Print strDisplay
End Sub


' ************* MOUSE *************


Private Sub picClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveRow X, Y
End Sub

Private Sub picClient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveRow X, Y
End Sub

Private Sub picClient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveRow X, Y
    If mlngActiveRow Then xp.OpenURL MakeWiki(mtypRow(mlngActiveRow).Wiki)
End Sub

Private Sub ActiveRow(X As Single, Y As Single)
    Dim lngRow As Long
    
    lngRow = Y \ mlngRowHeight + 1
    If lngRow < 1 Or lngRow > db.Challenges Then
        lngRow = 0
    ElseIf X < mtypRow(lngRow).ChallengeLeft Or X > mtypRow(lngRow).ChallengeRight Then
        lngRow = 0
    End If
    If lngRow <> 0 Then xp.SetMouseCursor mcHand
    If lngRow <> mlngActiveRow Then
        If mlngActiveRow <> 0 Then DrawCell mlngActiveRow, 1, False, True
        mlngActiveRow = lngRow
        If mlngActiveRow <> 0 Then DrawCell mlngActiveRow, 1, True, True
    End If
End Sub

Private Sub lnkChallenge_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkChallenge_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkChallenge_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    xp.OpenURL MakeWiki(mtypRow(Index).Wiki)
End Sub

Private Sub picProgress_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngCharacter = 0 Then Exit Sub
    xp.SetMouseCursor mcHand
    If Button = vbLeftButton Then SetStars Index, X
End Sub

Private Sub picProgress_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngCharacter = 0 Then Exit Sub
    xp.SetMouseCursor mcHand
    Select Case Button
        Case vbLeftButton: SetStars Index, X
        Case vbMiddleButton: SetStars Index, -2
        Case vbRightButton: SetStars Index, -1
    End Select
End Sub

Private Sub picProgress_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngCharacter = 0 Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub SetStars(Index As Integer, ByVal X As Long)
    Dim lngStars As Long
    Dim strFile As String
    Dim lngDifference As Long
    
    If mlngCharacter = 0 Then Exit Sub
    Select Case X
        Case -1: lngStars = 0  'Clicking left of all stars clears
        Case -2: lngStars = 6  'Clicking right of all stars sets to 6
        Case Else: lngStars = (X \ (Me.picStars(0).Width \ 7))
    End Select
    If lngStars < 0 Then lngStars = 0
    With db.Challenge(mtypRow(Index).Index)
        If lngStars > .MaxStars Then lngStars = .MaxStars
    End With
    'mlngStars is character current, lngStars is desired
    'We allow clicking 5 twice to toggle 6
    'mlngStars vs mtypRow(Index).Stars?
    If mtypRow(Index).Stars = 5 And lngStars = 5 Then
        lngStars = 6
    End If
    If mtypRow(Index).Stars = 1 And lngStars = 1 Then
        lngStars = 0
    End If
    If mlngStars <> lngStars Or mlngIndex <> Index Then
        mlngStars = lngStars
        mtypRow(Index).Stars = mlngStars
        With db.Challenge(mtypRow(Index).Index)
            lngDifference = mlngStars - .Stars(mlngCharacter)
            .Stars(mlngCharacter) = mlngStars
        End With
        With db.Character(mlngCharacter)
            .ChallengeFavor = .ChallengeFavor + lngDifference
            .TotalFavor = .TotalFavor + lngDifference
        End With
        ShowStars Index
        ShowTotal
        frmCompendium.FavorChange mlngCharacter
        UpdateFavor
        DirtyFlag dfeData
        mlngIndex = Index
    End If
End Sub

Private Sub ShowStars(ByVal plngRow As Long)
    Me.picProgress(plngRow).Picture = Me.picStars(mtypRow(plngRow).Stars).Picture
    DrawCell plngRow, ceFavor, False, True
End Sub


' ************* TOP MENU *************


Private Sub cboCharacter_Click()
    If mblnOverride Then Exit Sub
    mlngCharacter = ComboGetValue(Me.cboCharacter)
    SortTable
    DrawGrid
End Sub

Private Sub usrchkOrder_UserChange()
    If Me.usrchkOrder.Value Then cfg.ChallengeOrder = coeGame Else cfg.ChallengeOrder = coeGroup
    DirtyFlag dfeSettings
    SortTable
    xp.LockWindow Me.Hwnd
    DrawGrid
    xp.UnlockWindow
End Sub


' ************* SCROLLBAR *************


Private Sub WheelScroll(plngIncrement As Long)
    Dim lngValue As Long
    
    If Not Me.scrollVertical.Visible Then Exit Sub
    lngValue = Me.scrollVertical.Value - plngIncrement
    If lngValue < Me.scrollVertical.Min Then lngValue = Me.scrollVertical.Min
    If lngValue > Me.scrollVertical.Max Then lngValue = Me.scrollVertical.Max
    If Me.scrollVertical.Value <> lngValue Then Me.scrollVertical.Value = lngValue
End Sub

Private Sub scrollVertical_GotFocus()
    Me.picClient.SetFocus
End Sub

Private Sub scrollVertical_Change()
    Scroll
End Sub

Private Sub scrollVertical_Scroll()
    Scroll
End Sub

Private Sub Scroll()
    Me.picClient.Top = 0 - Me.scrollVertical.Value * mlngRowHeight
End Sub
