VERSION 5.00
Begin VB.Form frmSagas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sagas"
   ClientHeight    =   6945
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSagas.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9345
   Begin VB.CheckBox chkFontSize 
      Caption         =   "Small Font Size"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   0
      Width           =   1815
   End
   Begin VB.ComboBox cboCharacter 
      Height          =   330
      Left            =   2820
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "nav"
      Top             =   0
      Width           =   1752
   End
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   2472
   End
   Begin Compendium.userTab usrTab 
      Height          =   372
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1632
      _ExtentX        =   2884
      _ExtentY        =   661
      Captions        =   "Epic,Heroic"
   End
   Begin VB.CheckBox chkRedo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Redo"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin VB.CheckBox chkHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   972
   End
   Begin VB.CheckBox chkUndo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Undo"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   972
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   4755
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "drp"
      Top             =   600
      Width           =   4752
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   3252
      Left            =   7620
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   780
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5652
      Left            =   60
      ScaleHeight     =   5655
      ScaleWidth      =   7215
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   7212
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1632
         Left            =   540
         ScaleHeight     =   1635
         ScaleWidth      =   3915
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "ctl"
         Top             =   840
         Width           =   3912
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Quest"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuQuest 
         Caption         =   "Dynamic"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Pack"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuPack 
         Caption         =   "Dynamic"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Progress"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuProgress 
         Caption         =   "Clear"
         Index           =   0
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Casual"
         Index           =   2
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Normal"
         Index           =   3
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Hard"
         Index           =   4
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Elite"
         Index           =   5
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "VIP Skip"
         Index           =   7
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Astral Shard Skip"
         Index           =   8
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Never Mind"
         Index           =   10
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "UserSort"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuUser 
         Caption         =   "Bring to Top"
         Index           =   0
      End
      Begin VB.Menu mnuUser 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuUser 
         Caption         =   "Restore Default Order"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmSagas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ColumnType
    Header As String
    Left As Long
    Width As Long
    Right As Long
    Saga As Long ' Index in db.Saga() or 0 if not a saga column
    Align As AlignmentConstants
End Type

Private Type RowSagaType
    Valid As Boolean
End Type

Private Type RowType
    Quest As Long
    Saga() As RowSagaType
    Common As ProgressEnum
    Group As Long
    Order As Long
    Tier As SagaTierEnum
    Level(1) As Long
    QuestRight As Long
    PackRight As Long
End Type

Private Type ChangeType
    Row As Long
    Column As Long
    Saga As Long
    SagaQuest As Long
    OldValue As ProgressEnum
    NewValue As ProgressEnum
End Type

Private Type ActionType
    Changes As Long
    Change() As ChangeType
End Type

Private Type QueueType
    NewAction As Boolean
    Current As Long
    Actions As Long
    Action() As ActionType
End Type

Private queue As QueueType

Private mtypRow() As RowType
Private mlngRows As Long
Private mlngMaxRows As Long

Private mtypCol() As ColumnType
Private mlngColumns As Long
Private mlngMaxCols As Long

Private mlngCellWidth As Long

Private mlngMarginX As Long
Private mlngMarginY As Long
Private mlngRowHeight As Long

Private mlngCharacter As Long
Private mlngColorValid As Long
Private mlngColorInvalid As Long

Private mlngRow As Long
Private mlngCol As Long

Private mlngSelectRow As Long
Private mlngSelectCol As Long
Private menSelectDifficulty As ProgressEnum

Private mblnOverride As Boolean
Private mblnFormMoved As Boolean

Private menSort As SagaSortEnum
Private mlngSagaMenu As Long


Private Sub chkFontSize_Click()
    If Me.chkFontSize.Value = vbChecked Then
        Me.picClient.FontSize = 6
        Me.picContainer.FontSize = 6
        Me.picHeader.FontSize = 6
    Else
        Me.picClient.FontSize = 10
        Me.picContainer.FontSize = 10
        Me.picHeader.FontSize = 10
    End If
    Redraw
End Sub

' ************* FORM *************


Private Sub Form_Load()
    If cfg.SagaTier <> steEpic Then
        mblnOverride = True
        Me.usrTab.ActiveTab = "Heroic"
        mblnOverride = False
    End If
    menSort = ssePack
    If mlngCharacter = 0 Then Redraw
    If Not xp.DebugMode Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.Hwnd)
    cfg.SavePosition Me
    mlngCharacter = 0
    mblnFormMoved = False
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    Me.chkHelp.Move Me.ScaleWidth - Me.chkHelp.Width, 0
    Me.chkRedo.Move Me.chkHelp.Left - Me.chkRedo.Width, 0
    Me.chkUndo.Move Me.chkRedo.Left - Me.chkUndo.Width, 0
    lngWidth = Me.chkUndo.Left - (Me.usrTab.Left + Me.usrTab.Width)
    lngLeft = Me.usrTab.Left + Me.usrTab.Width + (lngWidth - Me.cboCharacter.Width) \ 2
    lngTop = (Me.chkHelp.Height - Me.cboCharacter.Height) \ 2
    Me.cboCharacter.Move lngLeft, lngTop
    Me.chkFontSize.Move (Me.cboCharacter.Left + Me.cboCharacter.Width), lngTop
    
    With Me.picHeader
        lngLeft = .Left
        lngTop = .Top + .Height - PixelY
        lngWidth = Me.picContainer.Width
    End With
    lngHeight = Me.ScaleHeight - Me.picContainer.Top
    If lngHeight > Me.picClient.Height Then lngHeight = Me.picClient.Height
    Me.picContainer.Move lngLeft, lngTop, lngWidth, lngHeight
    ShowScrollbar
End Sub

Private Sub ShowScrollbar()
    Dim lngMax As Long
    
    With Me.picContainer
        Me.scrollVertical.Move .Left + .Width, .Top, Me.scrollVertical.Width, .Height
    End With
    lngMax = (Me.picClient.Height - Me.picContainer.Height) \ mlngRowHeight
    If lngMax < 0 Then lngMax = 0 Else lngMax = lngMax + 1
    With Me.scrollVertical
        .Value = 0
        .Max = lngMax
        If lngMax Then
            .LargeChange = Me.picContainer.Height \ mlngRowHeight
            .Visible = True
        Else
            .Visible = True
        End If
    End With
    
    With Me.scrollHorizontal
    .Visible = False
    End With
End Sub

Public Sub Redraw()
    ClearQueue
    cfg.RefreshColors Me
    Me.usrTab.TabActiveColor = cfg.GetColor(cgeDropSlots, cveBackground)
    LoadData
    SizeClient
    SortRows menSort
    ShowScrollbar
    InitCombo True
    If Not mblnFormMoved Then
        cfg.MoveForm Me
        mblnFormMoved = True
    End If
End Sub

Public Property Let Character(plngCharacter As Long)
    ComboSetValue Me.cboCharacter, plngCharacter
End Property

Public Sub CharacterListChanged()
    InitCombo True
End Sub

Public Sub DataFileChanged()
    InitCombo False
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.picContainer.Hwnd, Xpos, Ypos): WheelScroll lngValue
    End Select
End Sub

'Public Sub UpdateCharacters()
'    Dim i As Long
'
'    mblnOverride = True
'    ComboClear Me.cboCharacter
'    For i = 1 To db.Characters
'        ComboAddItem Me.cboCharacter, db.Character(i).Character, i
'    Next
'    ComboSetValue Me.cboCharacter, mlngCharacter
'    mlngCharacter = ComboGetValue(Me.cboCharacter)
'    If mlngCharacter = -1 Then mlngCharacter = 0
'    mblnOverride = False
'End Sub


' ************* DATA *************


Private Sub LoadData()
    GatherColumns
    GatherRows
    CommonProgress
End Sub

Private Sub InitCombo(pblnFindCharacter As Boolean)
    Dim strCharacter As String
    Dim lngIndex As Long
    Dim i As Long
    
    strCharacter = ComboGetText(Me.cboCharacter)
    If Len(strCharacter) = 0 And mlngCharacter > 0 And mlngCharacter <= db.Characters Then strCharacter = db.Character(mlngCharacter).Character
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

Private Sub GatherColumns()
    Dim lngIndex() As Long
    Dim strHeader As String
    Dim i As Long
    
    ReDim lngIndex(1 To db.Sagas)
    For i = 1 To db.Sagas
        lngIndex(db.Saga(i).Order) = i
    Next
    ReDim mtypCol(1 To 4 + db.Sagas)
    InitColumn 1, "Lvl", vbRightJustify
    InitColumn 2, "Quest", vbLeftJustify
    InitColumn 3, "Pack", vbLeftJustify
    If mlngCharacter Then strHeader = db.Character(mlngCharacter).Character
    InitColumn 4, strHeader, vbCenter
    mlngColumns = 4
    mlngMaxCols = 0
    For i = 1 To db.Sagas
        With db.Saga(i)
            If .Tier = steEpic Then mlngMaxCols = mlngMaxCols + 1
            If .Tier = cfg.SagaTier Then
                mlngColumns = mlngColumns + 1
                InitColumn 4 + .Order, .Abbreviation, vbCenter
                mtypCol(4 + .Order).Saga = i
            End If
        End With
    Next
    ReDim Preserve mtypCol(1 To mlngColumns)
End Sub

Private Sub InitColumn(plngCol As Long, pstrHeader As String, penAlign As AlignmentConstants)
    With mtypCol(plngCol)
        .Header = pstrHeader
        .Align = penAlign
    End With
End Sub

Private Sub GatherRows()
    Dim lngIndex() As Long
    Dim lngQuest As Long
    Dim s As Long
    Dim q As Long
    
    ReDim lngIndex(1 To db.Quests)
    ReDim mtypRow(1 To db.Quests)
    mlngRows = 0
    mlngMaxRows = 0
    For s = 1 To db.Sagas
        For q = 1 To db.Saga(s).Quests
            If db.Saga(s).Tier = steEpic Then mlngMaxRows = mlngMaxRows + 1
            If db.Saga(s).Tier = cfg.SagaTier Then
                lngQuest = db.Saga(s).Quest(q)
                If lngIndex(lngQuest) = 0 Then
                    mlngRows = mlngRows + 1
                    lngIndex(lngQuest) = mlngRows
                End If
                With mtypRow(lngIndex(lngQuest))
                    ReDim Preserve .Saga(1 To db.Sagas)
                    .Saga(s).Valid = True
                    .Quest = lngQuest
                    .Tier = db.Saga(s).Tier
                    .Group = db.Quest(lngQuest).SagaGroup(.Tier)
                    .Order = db.Quest(.Quest).SagaOrder(.Tier)
                    .Level(steHeroic) = db.Quest(lngQuest).BaseLevel
                    .Level(steEpic) = db.Quest(lngQuest).EpicLevel
                End With
            End If
        Next
    Next
    ReDim Preserve mtypRow(1 To mlngRows)
    CommonProgress
End Sub

Private Sub SortRows(penSort As SagaSortEnum, Optional pblnClearUserSort As Boolean = False)
    Dim enTier As SagaTierEnum
    Dim strBringToTop() As String
    Dim lngBringToTop As Long
    Dim i As Long
    
    If pblnClearUserSort Then cfg.ClearSagaBringToTop
    menSort = penSort
    enTier = cfg.SagaTier
    ' Natural order of quests stays the same within groups
    ' For the most part, sorting is handled at a group level
    For i = 1 To mlngRows
        With mtypRow(i)
            Select Case menSort
                Case ssePack: .Group = db.Quest(.Quest).SagaGroup(enTier)
                Case sseLevel: .Group = .Level(enTier)
                Case sseQuest: .Group = Asc(UCase$(db.Quest(.Quest).Quest))
                Case sseUser: .Group = db.Quest(.Quest).SagaGroup(enTier)
            End Select
        End With
    Next
    ' Apply any user sorts now
    lngBringToTop = cfg.GetSagaBringToTop(strBringToTop)
    For i = 1 To lngBringToTop
        BringToTop strBringToTop(i)
    Next
    ' Sort array and display result
    InsertionSort
    DrawGrid
End Sub

Private Sub BringToTop(pstrSaga As String)
    Dim lngSaga As Long
    Dim i As Long
    
    lngSaga = SeekSaga(pstrSaga)
    If lngSaga = 0 Then Exit Sub
    For i = 1 To mlngRows
        With mtypRow(i)
            If .Saga(lngSaga).Valid Then .Group = 1 Else .Group = .Group + 1
        End With
    Next
End Sub

Private Sub InsertionSort()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As RowType
    
    iMin = 2
    iMax = mlngRows
    For i = iMin To iMax
        typSwap = mtypRow(i)
        For j = i To iMin Step -1
            If CompareRows(typSwap, mtypRow(j - 1)) = -1 Then mtypRow(j) = mtypRow(j - 1) Else Exit For
        Next j
        mtypRow(j) = typSwap
    Next i
End Sub

Private Function CompareRows(ptypLeft As RowType, ptypRight As RowType) As Long
    Select Case menSort
        Case ssePack, sseLevel
            If ptypLeft.Group < ptypRight.Group Then
                CompareRows = -1
            ElseIf ptypLeft.Group > ptypRight.Group Then
                CompareRows = 1
            ElseIf ptypLeft.Order < ptypRight.Order Then
                CompareRows = -1
            ElseIf ptypLeft.Order > ptypRight.Order Then
                CompareRows = 1
            End If
        Case sseQuest
            If db.Quest(ptypLeft.Quest).Quest < db.Quest(ptypRight.Quest).Quest Then
                CompareRows = -1
            ElseIf db.Quest(ptypLeft.Quest).Quest > db.Quest(ptypRight.Quest).Quest Then
                CompareRows = 1
            End If
        Case sseUser
    End Select
End Function

Private Sub CommonProgress()
    Dim i As Long
    
    If mlngCharacter = 0 Then Exit Sub
    For i = 1 To mlngRows
        mtypRow(i).Common = GetCommonProgress(i)
    Next
End Sub

Private Function GetCommonProgress(plngRow As Long) As ProgressEnum
    Dim lngQuest As Long
    Dim lngSaga As Long
    Dim lngSagaQuest As Long
    Dim enProgress As ProgressEnum
    Dim enLow As ProgressEnum
    Dim lngCol As Long
    
    lngQuest = mtypRow(plngRow).Quest
    If mlngCharacter = 0 Or lngQuest = 0 Then Exit Function
    enLow = peElite
    For lngCol = 5 To mlngColumns
        If GetSagaProgress(plngRow, lngCol, enProgress) Then
            Select Case enProgress
                Case peAstrals, peVIP
                Case Else: If enLow > enProgress Then enLow = enProgress
            End Select
        End If
    Next
    GetCommonProgress = enLow
End Function

Private Function GetSagaProgress(plngRow As Long, plngCol As Long, penProgress As ProgressEnum) As Boolean
    Dim lngQuest As Long
    Dim lngSaga As Long
    Dim lngSagaQuest As Long
    
    lngQuest = mtypRow(plngRow).Quest
    If mlngCharacter = 0 Or lngQuest = 0 Then Exit Function
    lngSaga = mtypCol(plngCol).Saga
    If mtypRow(plngRow).Saga(lngSaga).Valid Then
        lngSagaQuest = db.Quest(lngQuest).Saga(lngSaga)
        If lngSagaQuest Then
            penProgress = db.Character(mlngCharacter).Saga(lngSaga).Progress(lngSagaQuest)
            GetSagaProgress = True
        End If
    End If
End Function


' ************* SIZING *************

Private Sub SizeClient()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngMaxHeight As Long
    Dim strPack As String
    Dim i As Long
    
    With Me.picClient
        mlngMarginX = .ScaleX(cfg.MarginX, vbPixels, vbTwips)
        mlngMarginY = .ScaleY(cfg.MarginY, vbPixels, vbTwips)
        mlngRowHeight = .TextHeight("Q") + mlngMarginY * 2
    End With
    GrowColumn 1, "Lvl"
    GrowColumn 1, "32"
    For i = 1 To mlngRows
        With db.Quest(mtypRow(i).Quest)
            mtypRow(i).QuestRight = GrowColumn(2, .Quest)
            If Len(.Pack) Then strPack = .Pack Else strPack = "Free to Play"
            mtypRow(i).PackRight = GrowColumn(3, strPack)
        End With
    Next
    GrowColumn 2, "Garl's Tomb: Troglodyte's Get"
    For i = 1 To db.Characters
        GrowColumn 4, db.Character(i).Character
    Next
    For i = peCasual To peVIP
        GrowColumn 4, GetProgressName(i)
    Next
    For i = 1 To db.Sagas
        GrowColumn 4, db.Saga(i).Abbreviation
    Next
    For i = 5 To mlngColumns
        mtypCol(i).Width = mtypCol(4).Width
    Next
    lngLeft = 0
    For i = 1 To mlngColumns
        With mtypCol(i)
            .Left = lngLeft
            .Width = .Width + mlngMarginX * 2
            .Right = .Left + .Width
            lngLeft = .Right
        End With
    Next
    For i = 1 To mlngRows
        With mtypRow(i)
            .QuestRight = mtypCol(2).Left + .QuestRight
            .PackRight = mtypCol(3).Left + .PackRight
        End With
    Next
    mlngCellWidth = mtypCol(4).Width
    lngLeft = Me.scrollVertical.Width
    Me.usrTab.Move lngLeft, PixelY, Me.usrTab.TabsWidth
    lngTop = Me.usrTab.Top + Me.usrTab.Height - PixelY
    lngWidth = mtypCol(mlngColumns).Right + PixelX
    Me.picHeader.Move lngLeft, lngTop, lngWidth, mlngRowHeight + PixelY
    lngHeight = mlngRowHeight * mlngRows + PixelY
    Me.picClient.Move 0, 0, lngWidth, lngHeight
    With Me.picHeader
        Me.picContainer.Move .Left, .Top + .Height - PixelY, mtypCol(mlngColumns).Right + PixelX
    End With
    ' Form
    If Me.WindowState <> 2 Then  ' don't move if maximized
        Me.Width = Me.Width - Me.ScaleWidth + Me.picHeader.Left + mtypCol(mlngColumns).Right + PixelY + Me.scrollVertical.Width
        lngHeight = Me.Height - Me.ScaleHeight + Me.picHeader.Top + mlngRowHeight * mlngMaxRows
        xp.GetDesktop 0, 0, 0, lngMaxHeight
        If lngHeight > lngMaxHeight Then lngHeight = lngMaxHeight
        Me.Height = lngHeight
    End If
End Sub

Private Function GrowColumn(plngCol As Long, pstrText As String) As Long
    Dim lngWidth As Long
    
    lngWidth = Me.picClient.TextWidth(pstrText)
    If mtypCol(plngCol).Width < lngWidth Then mtypCol(plngCol).Width = lngWidth
    GrowColumn = lngWidth
End Function


' ************* DRAWING *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    mlngColorInvalid = cfg.GetColor(cgeControls, cveBackground)
    If mlngCharacter = 0 Then
        If cfg.SagaTier = steEpic Then mlngColorValid = GetColorValue(gcePink) Else mlngColorValid = GetColorValue(gceGreen)
    Else
        mlngColorValid = db.Character(mlngCharacter).BackColor
    End If
    DrawHeaders
    For lngRow = 1 To mlngRows
        For lngCol = 1 To mlngColumns
            DrawCell lngRow, lngCol, False
        Next
    Next
End Sub

Private Sub DrawHeaders()
    Dim lngActiveColumn As Long
    Dim enColor As ColorValueEnum
    Dim i As Long
    
    Me.picHeader.Cls
    Me.picHeader.ForeColor = cfg.GetColor(cgeDropSlots, cveText)
    Select Case menSort
        Case sseLevel: lngActiveColumn = 1
        Case sseQuest: lngActiveColumn = 2
        Case Else: lngActiveColumn = 3
    End Select
    For i = 1 To mlngColumns
        If i = lngActiveColumn Then enColor = cveBackRelated Else enColor = cveBackground
        Me.picHeader.FillColor = cfg.GetColor(cgeDropSlots, enColor)
        DrawHeader i
    Next
End Sub

Private Sub DrawHeader(plngCol As Long)
    Dim lngLeft As Long
    
    With mtypCol(plngCol)
        Me.picHeader.Line (.Left, 0)-(.Right, mlngRowHeight), cfg.GetColor(cgeControls, cveBorderExterior), B
        Select Case .Align
            Case vbLeftJustify: lngLeft = .Left + mlngMarginX
            Case vbCenter: lngLeft = .Left + (.Width - Me.picHeader.TextWidth(.Header)) \ 2
            Case vbRightJustify: lngLeft = .Right - Me.picHeader.TextWidth(.Header) - mlngMarginX
        End Select
        Me.picHeader.CurrentX = lngLeft
        Me.picHeader.CurrentY = mlngMarginY
        Me.picHeader.Print .Header
    End With
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    Dim strDisplay As String
    Dim lngQuest As Long
    Dim lngSaga As Long
    Dim enProgress As ProgressEnum
    Dim blnBold As Boolean
    
    lngLeft = mtypCol(plngCol).Left
    lngTop = (plngRow - 1) * mlngRowHeight
    lngRight = mtypCol(plngCol).Right
    lngBottom = lngTop + mlngRowHeight
    ' Background
    Select Case plngCol
        Case 1 To 3: lngColor = cfg.GetColor(cgeControls, cveBackground)
        Case 4: If mlngCharacter Then lngColor = mlngColorValid Else lngColor = mlngColorInvalid
        Case Else: If mtypRow(plngRow).Saga(mtypCol(plngCol).Saga).Valid Then lngColor = mlngColorValid Else lngColor = mlngColorInvalid
    End Select
    Me.picClient.FillColor = lngColor
    If pblnActive And plngCol > 3 Then lngColor = cfg.GetColor(cgeControls, cveBorderHighlight) Else lngColor = cfg.GetColor(cgeControls, cveBorderInterior)
    Me.picClient.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngColor, B
    ' Group Borders
    If Not (pblnActive And plngCol > 3) Then
        lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
        ' Top
        If plngRow = 1 Then blnBold = True Else blnBold = (mtypRow(plngRow).Group <> mtypRow(plngRow - 1).Group)
        If blnBold Then Me.picClient.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngColor
        ' Left
        If plngCol = 1 Then Me.picClient.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), lngColor
        ' Right
        If plngCol = mlngColumns Then Me.picClient.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), lngColor
        ' Bottom
        If plngRow = mlngRows Then blnBold = True Else blnBold = (mtypRow(plngRow).Group <> mtypRow(plngRow + 1).Group)
        If blnBold Then Me.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngColor
    End If
    ' Text
    lngQuest = mtypRow(plngRow).Quest
    lngColor = cfg.GetColor(cgeControls, cveText)
    Select Case plngCol
        Case 1
            If cfg.SagaTier = steEpic Then strDisplay = db.Quest(lngQuest).EpicLevel Else strDisplay = db.Quest(lngQuest).BaseLevel
        Case 2
            strDisplay = db.Quest(lngQuest).Quest
            If pblnActive Then lngColor = cfg.GetColor(cgeControls, cveTextLink) Else lngColor = cfg.GetColor(cgeControls, cveText)
        Case 3
            If Len(db.Quest(lngQuest).Pack) = 0 Then strDisplay = "Free to Play" Else strDisplay = db.Quest(lngQuest).Pack
            If pblnActive Then
                lngColor = cfg.GetColor(cgeControls, cveTextLink)
            ElseIf Len(db.Quest(lngQuest).Pack) = 0 Then
                lngColor = cfg.GetColor(cgeControls, cveTextDim)
            End If
        Case 4
            strDisplay = GetProgressName(mtypRow(plngRow).Common)
            Select Case mtypRow(plngRow).Common
                Case peCasual, peNormal, peHard: lngColor = cfg.GetColor(cgeControls, cveTextDim)
            End Select
        Case Else
            If mlngCharacter <> 0 And lngQuest <> 0 Then
                lngSaga = mtypCol(plngCol).Saga
                lngQuest = db.Quest(lngQuest).Saga(lngSaga)
                If lngQuest Then
                    enProgress = db.Character(mlngCharacter).Saga(lngSaga).Progress(lngQuest)
                    strDisplay = GetProgressName(enProgress)
                    Select Case enProgress
                        Case peCasual, peNormal, peHard: lngColor = cfg.GetColor(cgeControls, cveTextDim)
                    End Select
                End If
            End If
    End Select
    Select Case mtypCol(plngCol).Align
        Case vbLeftJustify: lngLeft = lngLeft + mlngMarginX
        Case vbCenter: lngLeft = lngLeft + (mtypCol(plngCol).Width - Me.picClient.TextWidth(strDisplay)) \ 2
        Case vbRightJustify: lngLeft = mtypCol(plngCol).Right - Me.picClient.TextWidth(strDisplay) - mlngMarginX
    End Select
    If Len(strDisplay) Then
        Me.picClient.ForeColor = lngColor
        PrintText strDisplay, lngLeft, lngTop + mlngMarginY
    End If
End Sub

Private Sub PrintText(pstrText As String, plngLeft As Long, plngTop As Long)
    Me.picClient.CurrentX = plngLeft
    Me.picClient.CurrentY = plngTop
    Me.picClient.Print pstrText
End Sub


' ************* TOP MENUS *************


Private Sub usrTab_Click(pstrCaption As String)
    Dim enTier As SagaTierEnum
    
    If mblnOverride Then Exit Sub
    If pstrCaption = "Epic" Then enTier = steEpic Else enTier = steHeroic
    If cfg.SagaTier <> enTier Then
        cfg.SagaTier = enTier
        DirtyFlag dfeSettings
        Redraw
    End If
End Sub

Private Sub cboCharacter_Click()
    Dim lngCharacter As Long
    
    If mblnOverride Then Exit Sub
    With Me.cboCharacter
        If .ListIndex = -1 Then lngCharacter = 0 Else lngCharacter = .ItemData(.ListIndex)
    End With
    If mlngCharacter <> lngCharacter Then
        mlngCharacter = lngCharacter
        Redraw
    End If
End Sub

Private Sub chkHelp_Click()
    If UncheckButton(Me.chkHelp, mblnOverride) Then Exit Sub
    ShowHelp "Sagas"
End Sub


' ************* MOUSE *************


Private Sub picClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngStep As Long
    Dim lngSaga As Long
    Dim i As Long
    
    If ActiveCell(X, Y) Then Exit Sub
    If Button <> vbLeftButton And Button <> vbMiddleButton Then Exit Sub
    If mlngSelectRow <> 0 And mlngSelectCol <> 0 And mlngSelectRow <> mlngRow And mlngCol = mlngSelectCol Then
        lngSaga = mtypCol(mlngSelectCol).Saga
        If mlngSelectRow < mlngRow Then lngStep = 1 Else lngStep = -1
        For i = mlngSelectRow To mlngRow Step lngStep
            If mlngSelectCol = 4 Then
                SetDifficulty i, mlngSelectCol, menSelectDifficulty
            Else
                If mtypRow(i).Saga(lngSaga).Valid Then SetDifficulty i, mlngSelectCol, menSelectDifficulty
            End If
        Next
    End If
End Sub

Private Sub picClient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngSelectRow = 0
    mlngSelectCol = 0
    queue.NewAction = True
    If ActiveCell(X, Y) Then Exit Sub
    Select Case mlngCol
        Case 2: QuestClick Button
        Case 3: PackClick Button
        Case 4 To mlngColumns
            If Button = vbRightButton Then
                ProgressMenu
            Else
                mlngSelectRow = mlngRow
                mlngSelectCol = mlngCol
                If Button = vbMiddleButton Then
                    menSelectDifficulty = peNone
                Else
                    If mlngCol = 4 Then menSelectDifficulty = mtypRow(mlngRow).Common Else GetSagaProgress mlngRow, mlngCol, menSelectDifficulty
                End If
                If menSelectDifficulty = peAstrals Or menSelectDifficulty = peVIP Then
                    mlngSelectRow = 0
                    mlngSelectCol = 0
                    Exit Sub
                End If
                If menSelectDifficulty <> peNone Then cfg.Difficulty = menSelectDifficulty
                SetDifficulty mlngRow, mlngCol, menSelectDifficulty
            End If
    End Select
End Sub

Private Sub picClient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim enDifficulty As ProgressEnum
    
    queue.NewAction = True
    If Not ActiveCell(X, Y) Then
        If mlngCharacter <> 0 Then
            If Button = vbLeftButton Then
                If mlngRow Then
                    If mlngSelectRow = mlngRow Then
                        If mlngSelectCol = mlngCol Then
                            If mlngCol > 3 Then
                                If mlngCol = 4 Then enDifficulty = mtypRow(mlngRow).Common Else GetSagaProgress mlngRow, mlngCol, enDifficulty
                                If enDifficulty = cfg.Difficulty Then enDifficulty = peNone Else enDifficulty = cfg.Difficulty
                                SetDifficulty mlngRow, mlngCol, enDifficulty
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    mlngSelectRow = 0
    mlngSelectCol = 0
End Sub

Private Sub picHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetColumn X
End Sub

Private Sub picHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    Dim lngScroll As Long
    
    lngCol = GetColumn(X)
    Select Case lngCol
        Case 1
            SortRows sseLevel, True
        Case 2
            SortRows sseQuest, True
        Case 3
            SortRows ssePack, True
        Case 4
            frmCharacter.Character = mlngCharacter
            OpenForm "frmCharacter"
        Case 5 To mlngColumns
            Select Case Button
                Case vbLeftButton
                    lngScroll = Me.scrollVertical.Value
                    frmSagaDetail.SetDetail mtypCol(lngCol).Saga, mlngCharacter
                    frmSagaDetail.Show vbModal, Me
                    Redraw
                    If Me.scrollVertical.Visible Then Me.scrollVertical.Value = lngScroll
                Case vbRightButton
                    mlngSagaMenu = lngCol
                    PopupMenu Me.mnuMain(3)
            End Select
    End Select
End Sub

Private Sub mnuUser_Click(Index As Integer)
    Select Case Me.mnuUser(Index).Caption
        Case "Bring to Top"
            cfg.AddSagaBringToTop db.Saga(mtypCol(mlngSagaMenu).Saga).SagaName
            SortRows menSort
        Case "Restore Default Order"
            SortRows ssePack, True
    End Select
End Sub

Private Sub picHeader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetColumn X
End Sub

Private Function GetColumn(X As Single) As Long
    Dim lngCol As Long
    
    ActiveCell 0, 0
    Do
        For lngCol = 1 To mlngColumns
            If X >= mtypCol(lngCol).Left And X <= mtypCol(lngCol).Right Then Exit Do
        Next
        lngCol = 0
    Loop Until True
    If lngCol Then xp.SetMouseCursor mcHand
    GetColumn = lngCol
End Function

' Return TRUE if no active cell
Private Function ActiveCell(X As Single, Y As Single) As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    lngRow = (Y \ mlngRowHeight) + 1
    If lngRow > 0 And lngRow <= mlngRows Then
        If X > mtypCol(2).Left And X <= mtypRow(lngRow).QuestRight Then
            lngCol = 2
        ElseIf X > mtypCol(3).Left And X <= mtypRow(lngRow).PackRight Then
            lngCol = 3
        ElseIf X >= mtypCol(4).Left And X <= mtypCol(4).Right Then
            lngCol = 4
        Else
            Do
                For lngCol = 5 To mlngColumns
                    If X >= mtypCol(lngCol).Left And X <= mtypCol(lngCol).Right Then Exit Do
                Next
                lngCol = 0
            Loop Until True
            If lngCol Then
                If Not mtypRow(lngRow).Saga(mtypCol(lngCol).Saga).Valid Then lngCol = 0
            End If
        End If
    End If
    If lngRow = 0 Or lngCol = 0 Or (mlngCharacter = 0 And lngCol > 3) Then
        If mlngRow <> 0 And mlngCol <> 0 Then DrawCell mlngRow, mlngCol, False
        mlngRow = 0
        mlngCol = 0
        ActiveCell = True
    Else
        xp.SetMouseCursor mcHand
        If lngRow <> mlngRow Or lngCol <> mlngCol Then
            If mlngRow <> 0 And mlngCol <> 0 Then DrawCell mlngRow, mlngCol, False
            mlngRow = lngRow
            mlngCol = lngCol
            DrawCell mlngRow, mlngCol, True
        End If
    End If
End Function

Private Sub QuestClick(ByVal penButton As MouseButtonConstants)
    Dim lngQuest As Long
    
    lngQuest = mtypRow(mlngRow).Quest
    Select Case penButton
        Case vbLeftButton: QuestWiki lngQuest
        Case vbRightButton: QuestMenu lngQuest
    End Select
End Sub

Private Function QuestMenu(plngQuest As Long)
    Dim i As Long
    
    With db.Quest(plngQuest)
        ' Unload
        Me.mnuQuest(0).Caption = db.Quest(plngQuest).Quest
        For i = Me.mnuQuest.UBound To 1 Step -1
            Unload Me.mnuQuest(i)
        Next
        ' Add dynamic items
        For i = 1 To .Links
            AddQuestListItem .Link(i).FullName
        Next
        AddQuestListItem "-"
        ' Add static items
        AddQuestListItem "Never Mind"
        PopupMenu Me.mnuMain(0)
    End With
End Function

Private Sub AddQuestListItem(pstrCaption As String)
    Dim lngIndex As Long
    
    With Me
        lngIndex = .mnuQuest.UBound + 1
        Load .mnuQuest(lngIndex)
        With .mnuQuest(lngIndex)
            .Caption = pstrCaption
            .Visible = True
        End With
    End With
End Sub

Private Sub mnuQuest_Click(Index As Integer)
    Dim lngQuest As Long
    
    lngQuest = mtypRow(mlngRow).Quest
    Select Case Me.mnuQuest(Index).Caption
        Case "Never Mind": Exit Sub
        Case db.Quest(lngQuest).Quest: QuestWiki lngQuest
        Case Else: QuestLink lngQuest, Me.mnuQuest(Index).Caption
    End Select
End Sub

Private Sub PackClick(ByVal penButton As MouseButtonConstants)
    Dim lngQuest As Long
    
    lngQuest = mtypRow(mlngRow).Quest
    Select Case penButton
        Case vbLeftButton: PackWiki db.Quest(lngQuest).Pack
        Case vbRightButton: PackMenu db.Quest(lngQuest).Pack
    End Select
End Sub

Private Sub PackMenu(pstrPack As String)
    Dim lngPack As Long
    Dim i As Long
    
    If Len(pstrPack) = 0 Then Exit Sub
    lngPack = SeekPack(pstrPack)
    If lngPack = 0 Then Exit Sub
    ' Unload
    Me.mnuPack(0).Caption = pstrPack
    For i = Me.mnuPack.UBound To 1 Step -1
        Unload Me.mnuPack(i)
    Next
    ' Add dynamic items
    For i = 1 To db.Pack(lngPack).Links
        AddPackItem db.Pack(lngPack).Link(i).FullName
    Next
    ' Add static items
    AddPackItem "-"
    AddPackItem "Never Mind"
    PopupMenu Me.mnuMain(1)
End Sub

Private Sub AddPackItem(pstrCaption As String)
    Dim lngIndex As Long
    
    With Me
        lngIndex = .mnuPack.UBound + 1
        Load .mnuPack(lngIndex)
        With .mnuPack(lngIndex)
            .Caption = pstrCaption
            .Visible = True
        End With
    End With
End Sub

Private Sub mnuPack_Click(Index As Integer)
    Dim lngQuest As Long
    Dim strPack As String
    Dim lngPack As Long
    
    If mlngRow = 0 Then Exit Sub
    lngQuest = mtypRow(mlngRow).Quest
    If lngQuest Then strPack = db.Quest(lngQuest).Pack
    If Len(strPack) Then lngPack = SeekPack(strPack)
    If lngPack = 0 Then Exit Sub
    Select Case Me.mnuPack(Index).Caption
        Case "Never Mind"
        Case strPack: PackWiki strPack
        Case Else: PackLink lngPack, Me.mnuPack(Index).Caption
    End Select
End Sub

Private Sub ProgressMenu()
    Dim enProgress As ProgressEnum
    Dim blnAlreadySkipped As Boolean
    Dim lngSaga As Long
    Dim lngAstrals As Long
    Dim lngVIP As Long
    Dim i As Long
    
    Select Case mlngCol
        Case 4
            enProgress = mtypRow(mlngRow).Common
            For i = 0 To 1
                Me.mnuProgress(i).Visible = False
            Next
            For i = 2 To 5
                Me.mnuProgress(i).Enabled = (enProgress < i)
            Next
            For i = 6 To 8
                Me.mnuProgress(i).Visible = False
            Next
        Case 5 To mlngColumns
            For i = 0 To 1
                Me.mnuProgress(i).Visible = True
            Next
            GetSagaProgress mlngRow, mlngCol, enProgress
            blnAlreadySkipped = (enProgress = peAstrals Or enProgress = peVIP)
            For i = 2 To 5
                Me.mnuProgress(i).Enabled = Not blnAlreadySkipped
            Next
            lngSaga = mtypCol(mlngCol).Saga
            For i = 1 To db.Saga(lngSaga).Quests
                Select Case db.Character(mlngCharacter).Saga(lngSaga).Progress(i)
                    Case peAstrals: lngAstrals = lngAstrals + 1
                    Case peVIP: lngVIP = lngVIP + 1
                End Select
            Next
            For i = 6 To 8
                Me.mnuProgress(i).Visible = True
            Next
            Me.mnuProgress(7).Enabled = (enProgress = peNone And lngVIP = 0)
            Me.mnuProgress(8).Enabled = (enProgress = peNone And lngAstrals < db.Saga(lngSaga).Astrals)
    End Select
    PopupMenu Me.mnuMain(2)
End Sub

Private Sub mnuProgress_Click(Index As Integer)
    Dim enProgress As ProgressEnum
    
    queue.NewAction = True
    If Me.mnuProgress(Index).Caption = "Never Mind" Then Exit Sub
    enProgress = GetProgressID(Me.mnuProgress(Index).Caption)
    SetDifficulty mlngRow, mlngCol, enProgress
End Sub


' ************* QUEUE *************


Private Sub ClearQueue()
    Dim typBlank As QueueType
    
    queue = typBlank
    queue.NewAction = True
    Me.chkUndo.Enabled = False
    Me.chkRedo.Enabled = False
End Sub

Private Sub SetDifficulty(plngRow As Long, plngCol As Long, penDifficulty As ProgressEnum)
    Dim lngCol As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngSaga As Long
    Dim lngQuest As Long
    Dim enOld As ProgressEnum
    Dim enCommon As ProgressEnum
    Dim blnCommon As Boolean
    
    If plngCol = 4 Then
        lngFirst = 5
        lngLast = mlngColumns
    Else
        lngFirst = plngCol
        lngLast = plngCol
    End If
    blnCommon = False
    For lngCol = lngFirst To lngLast
        If mtypRow(plngRow).Saga(mtypCol(lngCol).Saga).Valid Then
            lngSaga = mtypCol(lngCol).Saga
            lngQuest = db.Quest(mtypRow(plngRow).Quest).Saga(lngSaga)
            enOld = db.Character(mlngCharacter).Saga(lngSaga).Progress(lngQuest)
            If (lngFirst = lngLast And enOld <> penDifficulty) Or (lngFirst <> lngLast And enOld < penDifficulty) Then
                If Not (penDifficulty <> peNone And enOld > peElite) Then
                    ChangeDifficulty lngSaga, lngQuest, penDifficulty, plngRow, lngCol
                    DrawCell plngRow, lngCol, (lngCol = plngCol)
                    DirtyFlag dfeData
                    blnCommon = True
                End If
            End If
        End If
    Next
    If blnCommon Then
        enCommon = GetCommonProgress(plngRow)
        If mtypRow(plngRow).Common <> enCommon Then
            mtypRow(plngRow).Common = enCommon
            DrawCell plngRow, 4, (plngCol = 4)
            If plngCol = 5 Then DrawCell plngRow, plngCol, True
        End If
    End If
End Sub

Private Sub ChangeDifficulty(plngSaga As Long, plngSagaQuest As Long, penDifficulty As ProgressEnum, plngRow As Long, plngCol As Long)
    Dim typBlank As ActionType
    
    With queue
        If .NewAction Then
            .Current = .Current + 1
            If .Current <> .Actions Then
                .Actions = .Current
                ReDim Preserve .Action(1 To .Actions)
            End If
            .Action(.Current) = typBlank
        End If
        .NewAction = False
        With .Action(.Current)
            .Changes = .Changes + 1
            ReDim Preserve .Change(1 To .Changes)
            With .Change(.Changes)
                .Saga = plngSaga
                .SagaQuest = plngSagaQuest
                .Row = plngRow
                .Column = plngCol
                .OldValue = db.Character(mlngCharacter).Saga(plngSaga).Progress(plngSagaQuest)
                .NewValue = penDifficulty
            End With
        End With
    End With
    db.Character(mlngCharacter).Saga(plngSaga).Progress(plngSagaQuest) = penDifficulty
    Me.chkUndo.Enabled = True
    Me.chkRedo.Enabled = False
End Sub

Private Sub chkUndo_Click()
    Dim blnChange As Boolean
    Dim i As Long
    
    If UncheckButton(Me.chkUndo, mblnOverride) Then Exit Sub
    With queue
        If .Current Then
            With .Action(.Current)
                For i = 1 To .Changes
                    With .Change(i)
                        db.Character(mlngCharacter).Saga(.Saga).Progress(.SagaQuest) = .OldValue
                        DrawCell .Row, .Column, False
                        mtypRow(.Row).Common = GetCommonProgress(.Row)
                        DrawCell .Row, 4, False
                        blnChange = True
                    End With
                Next
            End With
        End If
        If blnChange Then
            .Current = .Current - 1
            .NewAction = True
            Me.chkUndo.Enabled = (.Current > 0)
            Me.chkRedo.Enabled = True
            DirtyFlag dfeData
        End If
    End With
End Sub

Private Sub chkRedo_Click()
    Dim blnChange As Boolean
    Dim i As Long
    
    If UncheckButton(Me.chkRedo, mblnOverride) Then Exit Sub
    With queue
        .Current = .Current + 1
        If .Current <= .Actions Then
            With .Action(.Current)
                For i = 1 To .Changes
                    With .Change(i)
                        db.Character(mlngCharacter).Saga(.Saga).Progress(.SagaQuest) = .NewValue
                        DrawCell .Row, .Column, False
                        mtypRow(.Row).Common = GetCommonProgress(.Row)
                        DrawCell .Row, 4, False
                        blnChange = True
                    End With
                Next
            End With
        End If
        If blnChange Then
            .NewAction = True
            Me.chkUndo.Enabled = True
            Me.chkRedo.Enabled = (.Current < .Actions)
            DirtyFlag dfeData
        End If
    End With
End Sub


' ************* SCROLLBARS *************


Private Sub WheelScroll(plngIncrement As Long)
    Dim lngValue As Long
    
    If Not Me.scrollVertical.Visible Then Exit Sub
    lngValue = Me.scrollVertical.Value - plngIncrement
    If lngValue < Me.scrollVertical.Min Then lngValue = Me.scrollVertical.Min
    If lngValue > Me.scrollVertical.Max Then lngValue = Me.scrollVertical.Max
    If Me.scrollVertical.Value <> lngValue Then Me.scrollVertical.Value = lngValue
End Sub

Private Sub scrollVertical_GotFocus()
    Me.picHeader.SetFocus
End Sub

Private Sub scrollVertical_Change()
    VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    Me.picClient.Top = 0 - Me.scrollVertical.Value * mlngRowHeight
End Sub

Private Sub scrollHorizontal_GotFocus()
    Me.picHeader.SetFocus
End Sub

Private Sub scrollHorizontal_Change()
    HorizontalScroll
End Sub

Private Sub scrollHorizontal_Scroll()
    HorizontalScroll
End Sub
Private Sub HorizontalScroll()
    '2021.08.27 figure out how much to HScroll
    Me.picClient.Left = 0 - Me.scrollHorizontal.Value
End Sub

