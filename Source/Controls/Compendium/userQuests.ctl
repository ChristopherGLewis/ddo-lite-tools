VERSION 5.00
Begin VB.UserControl userQuests 
   BackColor       =   &H00FFFFC0&
   ClientHeight    =   3756
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5004
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3756
   ScaleWidth      =   5004
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   1080
      ScaleHeight     =   372
      ScaleWidth      =   2592
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   2592
   End
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   840
      TabIndex        =   3
      Top             =   3300
      Visible         =   0   'False
      Width           =   2472
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   2172
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2472
      Left            =   420
      ScaleHeight     =   2472
      ScaleWidth      =   3672
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3672
      Begin VB.PictureBox picQuests 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1752
         Left            =   420
         ScaleHeight     =   1752
         ScaleWidth      =   3012
         TabIndex        =   1
         Top             =   360
         Width           =   3012
         Begin VB.PictureBox picTotals 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   792
            Left            =   120
            ScaleHeight     =   792
            ScaleWidth      =   2832
            TabIndex        =   5
            Top             =   660
            Visible         =   0   'False
            Width           =   2832
            Begin VB.Label lnkTotalFavor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   0
               Left            =   1584
               TabIndex        =   9
               Tag             =   "wrk"
               Top             =   360
               Visible         =   0   'False
               Width           =   972
            End
            Begin VB.Label lnkChallengeFavor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   252
               Index           =   0
               Left            =   1584
               TabIndex        =   8
               Tag             =   "wrk"
               Top             =   60
               Visible         =   0   'False
               Width           =   972
            End
            Begin VB.Label lblLabel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   " Total Favor:"
               ForeColor       =   &H80000008&
               Height          =   264
               Index           =   1
               Left            =   432
               TabIndex        =   7
               Tag             =   "wrk"
               Top             =   360
               Width           =   960
            End
            Begin VB.Label lblLabel 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   " Challenge Favor:"
               ForeColor       =   &H80000008&
               Height          =   264
               Index           =   0
               Left            =   60
               TabIndex        =   6
               Tag             =   "wrk"
               Top             =   60
               Width           =   1356
            End
         End
      End
   End
   Begin VB.Label lblControl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quests"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   624
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Popups"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuCharacter 
         Caption         =   "Command"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Quest"
      Index           =   1
      Begin VB.Menu mnuQuest 
         Caption         =   "Skip this Quest"
         Index           =   0
      End
      Begin VB.Menu mnuQuest 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuQuest 
         Caption         =   "Never Mind"
         Index           =   2
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Packs"
      Index           =   2
      Begin VB.Menu mnuPack 
         Caption         =   "Skip this Pack"
         Index           =   0
      End
      Begin VB.Menu mnuPack 
         Caption         =   "Don't Skip this Pack"
         Index           =   1
      End
      Begin VB.Menu mnuPack 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPack 
         Caption         =   "Never Mind"
         Index           =   3
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Progress"
      Index           =   3
      Begin VB.Menu mnuProgress 
         Caption         =   "Clear"
         Index           =   0
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Solo"
         Index           =   2
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Casual"
         Index           =   3
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Normal"
         Index           =   4
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Hard"
         Index           =   5
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Elite"
         Index           =   6
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "Never Mind"
         Index           =   8
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "LevelSort"
      Index           =   4
      Begin VB.Menu mnuLevel 
         Caption         =   "Sort Quests by Heroic Level"
         Index           =   0
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Sort Quests by Epic Level"
         Index           =   1
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Match In-Game Order"
         Index           =   2
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "QuestDynamic"
      Index           =   5
      Begin VB.Menu mnuDynamic 
         Caption         =   "Dynamic"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "PackDynamic"
      Index           =   6
      Begin VB.Menu mnuPackDynamic 
         Caption         =   "Dynamic"
         Index           =   0
      End
   End
End
Attribute VB_Name = "userQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ColumnStyleEnum
    cseInactive
    cseLink
    cseCharacter
End Enum

Private Type ColumnType
    Header As String
    Align As AlignmentConstants
    TextColor As Long
    TextActive As Long
    BackColor As Long
    BackDim As Long
    Left As Long
    Width As Long
    Style As ColumnStyleEnum
End Type

Private Type RowType
    QuestID As Long
    BaseLevel As Long
    EpicLevel As Long
    SortLevel As Long
    SortName As String
    Favor As Long
    Patron As String
    Style As QuestStyleEnum
    Pack As String
    Group As Long
    Order As Long
    Skipped As Boolean
    Progress() As ProgressEnum
End Type

Private mlngRowHeight As Long
Private mlngMarginX As Long
Private mlngMarginY As Long

Private mblnInitialized As Boolean

Private mtypRow() As RowType
Private mtypCol() As ColumnType
Private mlngRows As Long
Private mlngCols As Long
Private mlngRow As Long
Private mlngCol As Long

Private mlngSelectCol As Long
Private mlngSelectRow As Long
Private menSelectDifficulty As ProgressEnum

Private mblnCharacters As Boolean
Private mstrPack As String
Private mlngCharacter As Long

Private mlngFitWidth As Long

Private mblnOverride As Boolean
Private mblnLoaded As Boolean



' ************* USERCONTROL *************


Private Sub UserControl_Initialize()
    mblnLoaded = False
    mblnInitialized = False
End Sub

Private Sub UserControl_Resize()
    If mblnInitialized Then SizeClient
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell 0, 0
End Sub


' ************* PUBLIC *************


Public Sub Init()
    With UserControl
        .lblControl.Visible = False
        .picHeader.Visible = True
        .picContainer.Visible = True
    End With
    ReQuery False
    CalculateSizes
    mblnLoaded = True
End Sub

Public Sub Show()
    SizeClient
    Redraw
    Me.Scroll = cfg.CompendiumScroll
    mblnInitialized = True
End Sub

' Load data
Public Sub ReQuery(Optional pblnRedraw As Boolean = True)
    If pblnRedraw Then xp.Mouse = msAppWait
    GatherData
    SortRows
    AssignGroups
    If pblnRedraw Then
        Recalculate
        xp.Mouse = msNormal
    End If
End Sub

' Calculate sizes
Public Sub Recalculate(Optional pblnRedraw As Boolean = True)
    CalculateSizes
    SizeClient
    If pblnRedraw Then Redraw
End Sub

' Draw grid
Public Sub Redraw()
    RefreshColors
    DrawGrid
End Sub

Public Property Get Hwnd() As Long
    Hwnd = UserControl.picContainer.Hwnd
End Property

Private Sub RefreshColors()
    Dim enColor As ColorGroupEnum
    Dim i As Long
    
    UserControl.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    UserControl.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    cfg.ApplyColors UserControl.picHeader, cgeDropSlots
    cfg.ApplyColors UserControl.picQuests, cgeControls
    
    enColor = cgeWorkspace
    cfg.ApplyColors UserControl.picContainer, enColor
    cfg.ApplyColors UserControl.picTotals, enColor
    cfg.ApplyColors UserControl.lblLabel(0), enColor
    cfg.ApplyColors UserControl.lblLabel(1), enColor
    For i = 0 To UserControl.lnkChallengeFavor.UBound
        cfg.ApplyColors UserControl.lnkChallengeFavor(i), enColor
        cfg.ApplyColors UserControl.lnkTotalFavor(i), enColor
    Next
    ColumnColors
End Sub

Public Sub ColumnColors()
    Dim i As Long
    
    ' Standard columns
    For i = 1 To mlngCols
        With mtypCol(i)
            .TextColor = cfg.GetColor(cgeControls, cveText)
            If .Style = cseLink Then .TextActive = cfg.GetColor(cgeControls, cveTextLink) Else .TextActive = .TextColor
            .BackColor = cfg.GetColor(cgeControls, cveBackground)
            .BackDim = .BackColor
        End With
    Next
    ' Character columns
    For i = 1 To db.Characters
        With mtypCol(i + 4)
            .BackColor = db.Character(i).BackColor
            .BackDim = db.Character(i).DimColor
        End With
    Next
End Sub

Public Sub GetFont(pstrName As String, pdblSize As Double)
    With UserControl.picQuests
        pstrName = .FontName
        pdblSize = .FontSize
    End With
End Sub

Public Function SetFont(pstrName As String, Optional pdblSize As Double) As Double
On Error Resume Next
    Dim i As Long
    
    With UserControl
        .FontName = pstrName
        If pdblSize > 1 Then .FontSize = pdblSize
        SetFont = .FontSize ' Return the actual font size we landed on
        .picHeader.FontName = pstrName
        .picHeader.FontSize = pdblSize
        .picContainer.FontName = pstrName
        .picContainer.FontSize = pdblSize
        .picQuests.FontName = pstrName
        .picQuests.FontSize = pdblSize
        .picTotals.FontName = pstrName
        .picTotals.FontSize = pdblSize
        For i = 0 To 1
            .lblLabel(i).FontName = pstrName
            .lblLabel(i).FontSize = pdblSize
        Next
        For i = 0 To .lnkChallengeFavor.UBound
            .lnkChallengeFavor(i).FontName = pstrName
            .lnkChallengeFavor(i).FontSize = pdblSize
        Next
        For i = 0 To .lnkTotalFavor.UBound
            .lnkTotalFavor(i).FontName = pstrName
            .lnkTotalFavor(i).FontSize = pdblSize
        Next
    End With
End Function

' Width this control wants to be (parent reads this property and resizes accordingly)
Public Property Get ClientWidth() As Long
    With UserControl
        ClientWidth = .picQuests.Width + .scrollVertical.Width
    End With
End Property

Public Property Get HeaderHeight() As Long
    HeaderHeight = UserControl.picContainer.Top
End Property

Public Property Get PageSize() As Long
    PageSize = UserControl.scrollVertical.LargeChange
End Property


' ************* DATA *************


Private Sub GatherData()
    Dim lngIndex As Long
    Dim i As Long
    
    ReDim mtypRow(1 To db.Quests)
    mlngRows = 0
    For i = 1 To db.Quests
        If IncludeQuest(i) Then
            mlngRows = mlngRows + 1
            With mtypRow(mlngRows)
                .QuestID = i
                .BaseLevel = db.Quest(i).BaseLevel
                .EpicLevel = db.Quest(i).EpicLevel
                If cfg.LevelSort = lseGame Then
                    .SortName = db.Quest(i).CompendiumName
                Else
                    .SortName = db.Quest(i).SortName
                End If
                Select Case cfg.LevelSort
                    Case lseHeroic: .SortLevel = db.Quest(i).GroupLevel
                    Case lseEpic: If .BaseLevel < .EpicLevel And (cfg.CompendiumOrder = coeLevel Or cfg.CompendiumOrder = coeEpic) Then .SortLevel = .EpicLevel Else .SortLevel = db.Quest(i).GroupLevel
                    Case lseGame: .SortLevel = .BaseLevel
                End Select
                .Favor = db.Quest(i).Favor
                .Patron = db.Quest(i).Patron
                If cfg.AbbreviateColumns And cfg.AbbreviatePatrons Then
                    lngIndex = SeekPatron(.Patron)
                    If lngIndex Then .Patron = db.Patron(lngIndex).Abbreviation
                End If
                .Style = db.Quest(i).Style
                .Pack = db.Quest(i).Pack
                If cfg.AbbreviateColumns And cfg.AbbreviatePacks Then
                    lngIndex = SeekPack(.Pack)
                    If lngIndex Then .Pack = db.Pack(lngIndex).Abbreviation
                End If
                .Order = db.Quest(i).Order
                .Progress = db.Quest(i).Progress
                .Skipped = db.Quest(i).Skipped
            End With
        End If
    Next
    If mlngRows = 0 Then
        Erase mtypRow
    ElseIf mlngRows <> db.Quests Then
        ReDim Preserve mtypRow(1 To mlngRows)
    End If
End Sub

Private Function IncludeQuest(plngQuest As Long) As Boolean
'    If db.Quest(plngQuest).Hidden Then Exit Function
    Select Case cfg.CompendiumOrder
        Case coeEpic: IncludeQuest = (db.Quest(plngQuest).EpicLevel >= 20)
        Case coeStyle: IncludeQuest = (db.Quest(plngQuest).Style = qeRaid)
        Case Else: IncludeQuest = True
    End Select
End Function

' Quicksort (omit plngLeft & plngRight; they are used internally during recursion)
Private Sub SortRows(Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngMid As Long
    Dim typMid As RowType
    Dim typSwap As RowType
    
    If plngRight = 0 Then
        plngLeft = 1
        plngRight = mlngRows
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    typMid = mtypRow((plngLeft + plngRight) \ 2)
    Do
        Do While CompareRows(mtypRow(lngFirst), typMid) = -1 And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While CompareRows(typMid, mtypRow(lngLast)) = -1 And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            typSwap = mtypRow(lngFirst)
            mtypRow(lngFirst) = mtypRow(lngLast)
            mtypRow(lngLast) = typSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then SortRows plngLeft, lngLast
    If lngFirst < plngRight Then SortRows lngFirst, plngRight
End Sub

Private Function CompareRows(ptypLeft As RowType, ptypRight As RowType) As Long
    Do
        Select Case cfg.CompendiumOrder
            Case coeLevel
                Select Case cfg.LevelSort
                    Case lseHeroic, lseEpic
                        If ptypLeft.SortLevel < ptypRight.SortLevel Then
                            CompareRows = -1
                        ElseIf ptypLeft.SortLevel > ptypRight.SortLevel Then
                            CompareRows = 1
                        ElseIf ptypLeft.BaseLevel = 20 And ptypLeft.EpicLevel = 0 Then
                            If ptypRight.BaseLevel <> 20 Or ptypRight.EpicLevel <> 0 Then CompareRows = -1 Else Exit Do
                        ElseIf ptypRight.BaseLevel = 20 And ptypRight.EpicLevel = 0 Then
                            If ptypLeft.BaseLevel <> 20 Or ptypLeft.EpicLevel <> 0 Then CompareRows = 1 Else Exit Do
                        ElseIf ptypLeft.EpicLevel <> 0 And ptypRight.EpicLevel <> 0 And ptypLeft.BaseLevel < ptypRight.BaseLevel Then
                            CompareRows = -1
                        ElseIf ptypLeft.EpicLevel <> 0 And ptypRight.EpicLevel <> 0 And ptypLeft.BaseLevel > ptypRight.BaseLevel Then
                            CompareRows = 1
                        Else
                            Exit Do
                        End If
                    Case lseGame
                        If ptypLeft.SortLevel < ptypRight.SortLevel Then
                            CompareRows = -1
                        ElseIf ptypLeft.SortLevel > ptypRight.SortLevel Then
                            CompareRows = 1
                        ElseIf ptypLeft.SortName > ptypRight.SortName Then
                            CompareRows = -1
                        Else
                            CompareRows = 1
                        End If
                End Select
            Case coeEpic
                If ptypLeft.EpicLevel < ptypRight.EpicLevel Then
                    CompareRows = -1
                ElseIf ptypLeft.EpicLevel > ptypRight.EpicLevel Then
                    CompareRows = 1
                ElseIf ptypLeft.BaseLevel < ptypRight.BaseLevel Then
                    CompareRows = -1
                ElseIf ptypLeft.BaseLevel > ptypRight.BaseLevel Then
                    CompareRows = 1
                Else
                    Exit Do
                End If
            Case coeQuest
                If ptypLeft.SortName < ptypRight.SortName Then
                    CompareRows = -1
                ElseIf ptypLeft.SortName > ptypRight.SortName Then
                    CompareRows = 1
                Else
                    Exit Do
                End If
            Case coePack
                Exit Do
            Case coePatron
                If ptypLeft.Patron < ptypRight.Patron Then
                    CompareRows = -1
                ElseIf ptypLeft.Patron > ptypRight.Patron Then
                    CompareRows = 1
                ElseIf ptypLeft.BaseLevel < ptypRight.BaseLevel Then
                    CompareRows = -1
                ElseIf ptypLeft.BaseLevel > ptypRight.BaseLevel Then
                    CompareRows = 1
                Else
                    Exit Do
                End If
            Case coeFavor
                If ptypLeft.Favor > ptypRight.Favor Then
                    CompareRows = -1
                ElseIf ptypLeft.Favor < ptypRight.Favor Then
                    CompareRows = 1
                Else
                    Exit Do
                End If
            Case coeStyle
                If ptypLeft.Style < ptypRight.Style Then
                    CompareRows = -1
                ElseIf ptypLeft.Style > ptypRight.Style Then
                    CompareRows = 1
                ElseIf ptypLeft.SortLevel < ptypRight.SortLevel Then
                    CompareRows = -1
                ElseIf ptypLeft.SortLevel > ptypRight.SortLevel Then
                    CompareRows = 1
                Else
                    Exit Do
                End If
        End Select
        Exit Function
    Loop Until True
    If ptypLeft.Order < ptypRight.Order Then
        CompareRows = -1
    ElseIf ptypLeft.Order > ptypRight.Order Then
        CompareRows = 1
    End If
End Function

Private Sub AssignGroups()
    Dim lngGroup As Long
    Dim strOld As String
    Dim lngLevel As Long
    Dim lngOldLevel As Long
    Dim blnPacks As Boolean
    Dim i As Long
    
    lngGroup = 0
    For i = 1 To mlngRows
        With mtypRow(i)
            Select Case cfg.CompendiumOrder
                Case coeLevel
                    If cfg.LevelSort = lseGame Then
                        .Group = (i - 1) \ 20
                    Else
                        lngGroup = GetQuestLevel(mtypRow(i)) * 10
                        If lngGroup = 200 And mtypRow(i).EpicLevel = 20 Then lngGroup = 201
                        .Group = lngGroup
                    End If
                Case coeEpic: .Group = .EpicLevel
                Case coeQuest: .Group = Asc(Left$(.SortName, 1))
                Case coePack
                    lngLevel = mtypRow(i).SortLevel
                    If lngLevel < lngOldLevel And Not blnPacks Then
                        lngGroup = 100
                        .Group = lngGroup
                        strOld = .Pack
                        blnPacks = True
                    ElseIf blnPacks Then
                        If .Pack <> strOld Then
                            lngGroup = lngGroup + 1
                            strOld = .Pack
                        End If
                        .Group = lngGroup
                    Else
                        .Group = lngLevel
                    End If
                    lngOldLevel = lngLevel
                Case coePatron
                    If .Patron <> strOld Then
                        lngGroup = lngGroup + 1
                        strOld = .Patron
                    End If
                    .Group = lngGroup
                Case coeFavor: .Group = .Favor
                Case coeStyle: .Group = .Style
            End Select
        End With
    Next
End Sub

Private Function GetQuestLevel(ptypRow As RowType) As Long
    Select Case cfg.LevelSort
        Case lseHeroic: GetQuestLevel = ptypRow.SortLevel
        Case lseEpic: If ptypRow.EpicLevel <> 0 Then GetQuestLevel = ptypRow.EpicLevel Else GetQuestLevel = ptypRow.SortLevel
        Case lseGame: GetQuestLevel = ptypRow.BaseLevel
    End Select
End Function


' ************* SIZING *************


Public Property Get FitHeight() As Long
    FitHeight = -1
End Property

Public Property Get FitWidth() As Long
    If mlngFitWidth = 0 Then CalculateSizes
    FitWidth = mlngFitWidth
End Property

Public Sub CalculateSizes()
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngCol As Long
    Dim lngIndex As Long
    Dim strName As String
    Dim i As Long
    
    ' interior cell margins
    With UserControl.picQuests
        mlngMarginX = .ScaleX(cfg.MarginX, vbPixels, vbTwips)
        mlngMarginY = .ScaleY(cfg.MarginY, vbPixels, vbTwips)
        mlngRowHeight = UserControl.picQuests.TextHeight("Q") + mlngMarginY * 2
    End With
    ' Columns
    mlngCols = db.Characters + 7
    ReDim mtypCol(1 To mlngCols)
    With UserControl.picQuests
        SetColumn 1, "Lvl", lngLeft, .TextWidth("Lvl"), vbRightJustify
        SetColumn 2, "Ep", lngLeft, .TextWidth("Lvl"), vbRightJustify
        For i = 1 To db.Quests
            ' Quest
            GrowWidth 3, .TextWidth(db.Quest(i).Quest)
            ' Pack
            strName = db.Quest(i).Pack
            If cfg.AbbreviateColumns And cfg.AbbreviatePacks Then
                lngIndex = SeekPack(strName)
                If lngIndex Then strName = db.Pack(lngIndex).Abbreviation
            End If
            GrowWidth 4, .TextWidth(strName)
            ' Patron
            strName = db.Quest(i).Patron
            If cfg.AbbreviateColumns And cfg.AbbreviatePatrons Then
                lngIndex = SeekPatron(db.Quest(i).Patron)
                If lngIndex Then strName = db.Patron(lngIndex).Abbreviation
            End If
            GrowWidth db.Characters + 5, .TextWidth(strName)
        Next
        SetColumn 3, "Quest", lngLeft, mtypCol(3).Width, vbLeftJustify, cseLink
        SetColumn 4, "Pack", lngLeft, mtypCol(4).Width, vbLeftJustify, cseLink
        lngWidth = .TextWidth("Normal")
        If lngWidth < .TextWidth("Casual") Then lngWidth = .TextWidth("Casual")
        For i = 1 To db.Characters
            If lngWidth < .TextWidth(db.Character(i).Character) Then lngWidth = .TextWidth(db.Character(i).Character)
        Next
        For i = 1 To db.Characters
            SetColumn i + 4, db.Character(i).Character, lngLeft, lngWidth, vbCenter, cseCharacter
        Next
        lngCol = db.Characters + 5
        SetColumn lngCol, "Patron", lngLeft, mtypCol(lngCol).Width, vbCenter, cseLink
        SetColumn lngCol + 1, "Fvr", lngLeft, .TextWidth("Fvr"), vbRightJustify
        SetColumn lngCol + 2, "Party", lngLeft, .TextWidth("Party")
    End With
    mlngFitWidth = mtypCol(mlngCols).Left + mtypCol(mlngCols).Width + PixelX + UserControl.scrollVertical.Width
End Sub

Private Sub SetColumn(plngIndex As Long, pstrHeader As String, plngLeft As Long, plngWidth As Long, Optional penAlign As AlignmentConstants = vbLeftJustify, Optional penStyle As ColumnStyleEnum = cseInactive)
    With mtypCol(plngIndex)
        .Header = pstrHeader
        .Align = penAlign
        .Left = plngLeft
        .Width = plngWidth + mlngMarginX * 2
        plngLeft = .Left + .Width
        .Style = penStyle
        If .Style = cseLink Then .TextActive = cfg.GetColor(cgeControls, cveTextLink)
    End With
End Sub

Private Sub GrowWidth(plngIndex As Long, plngWidth As Long)
    With mtypCol(plngIndex)
        If .Width < plngWidth Then .Width = plngWidth
    End With
End Sub

Private Sub SizeClient()
    Dim lngRows As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngMax As Long
    Dim blnFavorTotals As Boolean
    
    If (UserControl.ScaleHeight \ mlngRowHeight) < 3 Then Exit Sub
    blnFavorTotals = (mlngRows = db.Quests And db.Characters > 0)
    With UserControl
        ' Header
        lngLeft = 0
        lngTop = 0
        lngWidth = mtypCol(mlngCols).Left + mtypCol(mlngCols).Width + PixelX
        lngHeight = frmCompendium.usrToolbar.FitHeight + PixelY
        .picHeader.Move lngLeft, lngTop, lngWidth, lngHeight
        ' Container
        lngLeft = 0
        lngWidth = .ScaleWidth - .scrollVertical.Width
        lngHeight = .ScaleHeight - .picHeader.Height
        lngTop = .picHeader.Height - PixelY
        .picContainer.Move lngLeft, lngTop, lngWidth, lngHeight
        ' Quests
        lngLeft = 0
        lngTop = 0
        lngWidth = mtypCol(mlngCols).Left + mtypCol(mlngCols).Width + PixelX
        lngHeight = mlngRows * mlngRowHeight + Screen.TwipsPerPixelY
        If blnFavorTotals Then lngHeight = lngHeight + .TextHeight("Q") * 2
        .picQuests.Move lngLeft, lngTop, lngWidth, lngHeight
        lngTop = lngHeight - .TextHeight("Q") * 2
        .picTotals.Move lngLeft, lngTop, lngWidth, .TextHeight("Q") * 2
        ShowFavorTotals
        ' Vertical scrollbar
        lngLeft = .ScaleWidth - .scrollVertical.Width
        lngTop = .picContainer.Top
        lngWidth = .scrollVertical.Width
        lngHeight = .picContainer.Height
        mblnLoaded = False
        With .scrollVertical
            .Move lngLeft, lngTop, lngWidth, lngHeight
            .LargeChange = UserControl.picContainer.ScaleHeight \ mlngRowHeight
            lngMax = mlngRows - .LargeChange + 1
            If blnFavorTotals Then
                Do While UserControl.picQuests.Height > (lngMax - 1) * mlngRowHeight + UserControl.picContainer.Height
                    lngMax = lngMax + 1
                Loop
            End If
            If lngMax < 1 Then
                mblnOverride = True
                .Min = 0
                .Value = 0
                .Max = 0
                .Enabled = False
                .Visible = False
                mblnOverride = False
                UserControl.picContainer.Move 0, 0
            Else
                .Min = 1
                .Value = 1
                .Max = lngMax
                .Enabled = True
                .Visible = True
            End If
        End With
        mblnLoaded = True
    End With
End Sub


' ************* DRAWING *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    DrawColumnHeaders
    UserControl.picQuests.Cls
    For lngRow = 1 To mlngRows
        For lngCol = 1 To mlngCols
            DrawCell lngRow, lngCol, False
        Next
    Next
End Sub

Private Sub DrawColumnHeaders()
    Dim lngSortCol As Long
    Dim lngCol As Long
    
    If cfg.CompendiumOrder < 4 Then lngSortCol = cfg.CompendiumOrder Else lngSortCol = cfg.CompendiumOrder + db.Characters
    UserControl.picHeader.Cls
    For lngCol = 1 To mlngCols
        DrawColumnHeader lngCol, (lngCol = lngSortCol + 1)
    Next
End Sub

Private Sub DrawColumnHeader(plngCol As Long, pblnActive As Boolean)
    Dim strCaption As String
    Dim lngColor As Long
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    With mtypCol(plngCol)
        strCaption = .Header
        lngLeft = .Left
        lngWidth = .Width
        lngHeight = UserControl.picHeader.ScaleHeight - PixelY
    End With
    If pblnActive Then lngColor = cfg.GetColor(cgeDropSlots, cveBackRelated) Else lngColor = cfg.GetColor(cgeDropSlots, cveBackground)
    With UserControl.picHeader
        .FillColor = lngColor
        UserControl.picHeader.Line (lngLeft, 0)-(lngLeft + lngWidth, lngHeight - PixelY), cfg.GetColor(cgeDropSlots, cveBorderExterior), B
        .CurrentX = lngLeft + (lngWidth - .TextWidth(strCaption)) \ 2
        .CurrentY = (lngHeight - .TextHeight("Q")) \ 2
    End With
    UserControl.picHeader.Print mtypCol(plngCol).Header
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Dim strDisplay As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    Dim blnDim As Boolean
    
    Select Case plngCol
        Case 1: strDisplay = mtypRow(plngRow).BaseLevel
        Case 2: strDisplay = mtypRow(plngRow).EpicLevel
        Case 3: If cfg.CompendiumOrder = coeQuest Or (cfg.CompendiumOrder = coeLevel And cfg.LevelSort = lseGame) Then strDisplay = mtypRow(plngRow).SortName Else strDisplay = db.Quest(mtypRow(plngRow).QuestID).Quest
        Case 4
            If Len(mtypRow(plngRow).Pack) = 0 Then
                strDisplay = "Free to Play"
                blnDim = True
            Else
                strDisplay = mtypRow(plngRow).Pack
            End If
        Case mlngCols - 2: strDisplay = mtypRow(plngRow).Patron
        Case mlngCols - 1: strDisplay = mtypRow(plngRow).Favor
        Case mlngCols: strDisplay = GetQuestStyleName(mtypRow(plngRow).Style)
        Case Else
            strDisplay = GetProgressName(mtypRow(plngRow).Progress(plngCol - 4))
            Select Case mtypRow(plngRow).Progress(plngCol - 4)
                Case peElite, peNone
                Case Else: blnDim = True
            End Select
    End Select
    If strDisplay = "0" Then strDisplay = vbNullString
    lngLeft = mtypCol(plngCol).Left
    lngTop = (plngRow - 1) * mlngRowHeight
    lngRight = lngLeft + mtypCol(plngCol).Width
    lngBottom = lngTop + mlngRowHeight
    With UserControl
        ' Borders
        If pblnActive And mtypCol(plngCol).Style = cseCharacter Then
            .picQuests.Line (lngLeft, lngTop)-(lngRight, lngBottom), cfg.GetColor(cgeControls, cveBorderHighlight), B
        Else
            ' Top
            If plngRow = 1 Then
                lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
            ElseIf mtypRow(plngRow).Group <> mtypRow(plngRow - 1).Group Then
                lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
            Else
                lngColor = cfg.GetColor(cgeControls, cveBorderInterior)
            End If
            .picQuests.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngColor
            ' Left
            .picQuests.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), cfg.GetColor(cgeControls, cveBorderInterior)
            ' Right
            .picQuests.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), cfg.GetColor(cgeControls, cveBorderInterior)
            ' Bottom
            If plngRow = mlngRows Then
                lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
            ElseIf mtypRow(plngRow).Group <> mtypRow(plngRow + 1).Group Then
                lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
            Else
                lngColor = cfg.GetColor(cgeControls, cveBorderInterior)
            End If
            .picQuests.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngColor
        End If
        ' Background
        If mtypRow(plngRow).Skipped Then lngColor = mtypCol(plngCol).BackDim Else lngColor = mtypCol(plngCol).BackColor
        .picQuests.Line (lngLeft + PixelX, lngTop + PixelY)-(lngRight - PixelX, lngBottom - PixelY), lngColor, BF
        ' Text
        Select Case mtypCol(plngCol).Align
            Case vbLeftJustify: .picQuests.CurrentX = lngLeft + mlngMarginX
            Case vbCenter: .picQuests.CurrentX = lngLeft + (mtypCol(plngCol).Width - .picQuests.TextWidth(strDisplay)) \ 2
            Case vbRightJustify: .picQuests.CurrentX = lngLeft + mtypCol(plngCol).Width - mlngMarginX - .picQuests.TextWidth(strDisplay)
        End Select
        If pblnActive Then
            lngColor = mtypCol(plngCol).TextActive
        ElseIf blnDim Then
            lngColor = cfg.GetColor(cgeControls, cveTextDim)
        Else
            lngColor = mtypCol(plngCol).TextColor
        End If
        .picQuests.ForeColor = lngColor
        .picQuests.CurrentY = lngTop + mlngMarginY
        .picQuests.Print strDisplay
    End With
End Sub

Private Sub ShowFavorTotals()
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngRowTop(1) As Long
    Dim lngRowHeight As Long
    Dim i As Long
    
    If mlngRows <> db.Quests Or db.Characters = 0 Then
        UserControl.picTotals.Visible = False
        Exit Sub
    End If
    With UserControl
        lngRowTop(0) = 0
        lngRowTop(1) = .picQuests.TextHeight("Q")
        For i = 0 To 1
            .lblLabel(i).Move mtypCol(5).Left - .TextWidth(" ") - .lblLabel(i).Width, lngRowTop(i)
            .lblLabel(i).Visible = True
        Next
        lngRowHeight = .lblLabel(0).Height
        For i = 1 To db.Characters
            If i > .lnkChallengeFavor.UBound Then
                Load .lnkChallengeFavor(i)
                Load .lnkTotalFavor(i)
            End If
            With UserControl.lnkChallengeFavor(i)
                .Move mtypCol(i + 4).Left, lngRowTop(0), mtypCol(i + 4).Width, lngRowHeight
                .Caption = db.Character(i).ChallengeFavor & "  "
                .Visible = True
            End With
            With UserControl.lnkTotalFavor(i)
                .Move mtypCol(i + 4).Left, lngRowTop(1), mtypCol(i + 4).Width, lngRowHeight
                .Caption = db.Character(i).TotalFavor & "  "
                .Visible = True
            End With
        Next
        For i = .lnkChallengeFavor.UBound To db.Characters + 1 Step -1
            Unload .lnkChallengeFavor(i)
            Unload .lnkTotalFavor(i)
        Next
    End With
    UserControl.picTotals.Visible = True
End Sub

Public Sub FavorChange(plngCharacter As Long)
    If mlngRows <> db.Quests Or db.Characters = 0 Then Exit Sub
    With UserControl
        .lnkChallengeFavor(plngCharacter).Caption = db.Character(plngCharacter).ChallengeFavor & "  "
        .lnkTotalFavor(plngCharacter).Caption = db.Character(plngCharacter).TotalFavor & "  "
    End With
End Sub


' ************* MOUSE *************


Private Sub picQuests_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngStep As Long
    Dim lngQuest As Long
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    If ActiveCell(X, Y) Then Exit Sub
    If Button <> vbLeftButton And Button <> vbMiddleButton Then Exit Sub
    If mlngSelectRow <> 0 And mlngSelectCol <> 0 And mlngSelectRow <> mlngRow And mlngCol = mlngSelectCol Then
        If mlngSelectRow < mlngRow Then lngStep = 1 Else lngStep = -1
        For i = mlngSelectRow To mlngRow Step lngStep
            lngQuest = mtypRow(i).QuestID
            If lngQuest Then SetDifficulty i, mlngSelectCol, lngQuest, menSelectDifficulty, (mlngRow = i)
        Next
    End If
End Sub

Private Sub picQuests_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnOverride Then Exit Sub
    mlngSelectRow = 0
    mlngSelectCol = 0
    ActiveCell X, Y
    Select Case Button
        Case vbLeftButton, vbMiddleButton
            Select Case mlngCol
                Case 3: QuestWiki mtypRow(mlngRow).QuestID
                Case 4: PackWiki mtypRow(mlngRow).Pack
                Case mlngCols - 2: PatronWiki mtypRow(mlngRow).Patron
                Case 5 To db.Characters + 4
                    If db.Characters <> 0 Then
                        mlngSelectRow = mlngRow
                        mlngSelectCol = mlngCol
                        If Button = vbMiddleButton Then menSelectDifficulty = peNone Else menSelectDifficulty = mtypRow(mlngRow).Progress(mlngCol - 4)
                        If menSelectDifficulty <> peNone And menSelectDifficulty <> peSolo Then cfg.Difficulty = menSelectDifficulty
                    End If
            End Select
        Case vbRightButton
            Select Case mlngCol
                Case 3: QuestMenu
                Case 4: PackMenu
                Case 5 To db.Characters + 4
                    If db.Characters > 0 Then ProgressMenu
            End Select
        End Select
End Sub

Private Sub picQuests_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngQuest As Long
    Dim enDifficulty As ProgressEnum
    
    If mblnOverride Then Exit Sub
    ActiveCell X, Y
    If Button <> vbLeftButton Then Exit Sub
    If db.Characters <> 0 Then
        Select Case mlngCol
            Case 5 To db.Characters + 4
                    If mlngSelectRow = mlngRow And mlngSelectCol = mlngCol Then
                        If mlngRow Then
                            Select Case mtypRow(mlngRow).Progress(mlngCol - 4)
                                Case peSolo, cfg.Difficulty: enDifficulty = peNone
                                Case Else: enDifficulty = cfg.Difficulty
                            End Select
                            lngQuest = mtypRow(mlngRow).QuestID
                            SetDifficulty mlngRow, mlngCol, lngQuest, enDifficulty, True
                        End If
                    End If
        End Select
    End If
    mlngSelectRow = 0
    mlngSelectCol = 0
End Sub

Private Function ActiveCell(X As Single, Y As Single) As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    lngRow = (Y \ mlngRowHeight) + 1
    For lngCol = mlngCols To 1 Step -1
        If X > mtypCol(lngCol).Left Then Exit For
    Next
    If lngCol > mlngCols Or lngRow > mlngRows Then
        lngCol = 0
        lngRow = 0
        ActiveCell = True
    End If
    If lngCol = mlngCol And lngRow = mlngRow Then
        ActiveCell = True
    Else
        If mlngRow > 0 And mlngRow <= mlngRows And mlngCol > 0 And mlngCol <= mlngCols Then
            If mtypCol(mlngCol).Style <> cseInactive Then DrawCell mlngRow, mlngCol, False
        End If
        mlngCol = lngCol
        mlngRow = lngRow
        If mlngRow <> 0 And mlngCol <> 0 Then
            If mtypCol(mlngCol).Style <> cseInactive Then DrawCell mlngRow, mlngCol, True
        End If
    End If
End Function

Private Sub SetDifficulty(plngRow As Long, plngCol As Long, plngQuest As Long, ByVal penDifficulty As ProgressEnum, pblnActive As Boolean)
    Dim lngCharacter As Long
    Dim lngNewFavor As Long
    Dim lngOldFavor As Long
    Dim lngDifference As Long
    
    lngCharacter = plngCol - 4
    If penDifficulty <> peNone Then
        Select Case db.Quest(plngQuest).Style
            Case qeSolo: penDifficulty = peSolo
            Case qeRaid:: If penDifficulty = peSolo Or penDifficulty = peCasual Then penDifficulty = peNormal
            Case qeQuest: If penDifficulty = peSolo Then penDifficulty = peCasual
        End Select
    End If
    If db.Quest(plngQuest).Progress(lngCharacter) <> penDifficulty Then
        lngOldFavor = QuestFavor(plngQuest, db.Quest(plngQuest).Progress(lngCharacter))
        lngNewFavor = QuestFavor(plngQuest, penDifficulty)
        lngDifference = lngNewFavor - lngOldFavor
        db.Quest(plngQuest).Progress(lngCharacter) = penDifficulty
        mtypRow(plngRow).Progress(lngCharacter) = penDifficulty
        DrawCell plngRow, plngCol, pblnActive
        With db.Character(lngCharacter)
            .QuestFavor = .QuestFavor + lngDifference
            .TotalFavor = .TotalFavor + lngDifference
        End With
        FavorChange lngCharacter
        UpdateFavor
        DirtyFlag dfeData
    End If
End Sub

Private Function QuestMenu()
    Dim i As Long
    
    If mlngRow = 0 Then Exit Function
    With db.Quest(mtypRow(mlngRow).QuestID)
        If .Links = 0 Then
            UserControl.mnuQuest(0).Checked = mtypRow(mlngRow).Skipped
            PopupMenu UserControl.mnuMain(1)
        Else
            ' Unload
            UserControl.mnuDynamic(0).Visible = True
            For i = UserControl.mnuDynamic.UBound To 1 Step -1
                Unload UserControl.mnuDynamic(i)
            Next
            ' Add dynamic items
            For i = 1 To .Links
                AddQuestListItem .Link(i).FullName
            Next
            If .Links Then AddQuestListItem "-"
            ' Add static items
            AddQuestListItem "Skip this Quest", mtypRow(mlngRow).Skipped
            AddQuestListItem "-"
            AddQuestListItem "Never Mind"
            UserControl.mnuDynamic(0).Visible = False
            PopupMenu UserControl.mnuMain(5)
        End If
    End With
End Function

Private Sub AddQuestListItem(pstrCaption As String, Optional pblnChecked As Boolean = False)
    Dim lngIndex As Long
    
    With UserControl
        lngIndex = .mnuDynamic.UBound + 1
        Load .mnuDynamic(lngIndex)
        With .mnuDynamic(lngIndex)
            .Caption = pstrCaption
            .Checked = pblnChecked
            .Visible = True
        End With
    End With
End Sub

Private Sub mnuQuest_Click(Index As Integer)
    Select Case UserControl.mnuQuest(Index).Caption
        Case "Never Mind": Exit Sub
        Case "Skip this Quest": SkipQuest
    End Select
End Sub

Private Sub mnuDynamic_Click(Index As Integer)
    Select Case UserControl.mnuDynamic(Index).Caption
        Case "Never Mind": Exit Sub
        Case "Skip this Quest": SkipQuest
        Case Else: QuestLink mtypRow(mlngRow).QuestID, UserControl.mnuDynamic(Index).Caption
    End Select
End Sub

Private Sub SkipQuest()
    Dim lngQuest As Long
    
    If mlngRow = 0 Then Exit Sub
    lngQuest = mtypRow(mlngRow).QuestID
    db.Quest(lngQuest).Skipped = Not db.Quest(lngQuest).Skipped
    mtypRow(mlngRow).Skipped = db.Quest(lngQuest).Skipped
    cfg.CompendiumScroll = UserControl.scrollVertical.Value
    Redraw
    UserControl.scrollVertical.Value = cfg.CompendiumScroll
    DirtyFlag dfeData
End Sub

' Returns -1 for Free to Play
Private Function PackID() As Long
    Dim lngQuest As Long
    Dim lngPack As Long
    
    If mlngRow = 0 Then Exit Function
    lngQuest = mtypRow(mlngRow).QuestID
    lngPack = SeekPack(db.Quest(lngQuest).Pack)
    If lngPack = 0 Then PackID = -1 Else PackID = lngPack
End Function

Private Sub PackMenu()
    Dim lngPack As Long
    Dim i As Long
    
    lngPack = PackID()
    If lngPack < 1 Then Exit Sub
    If db.Pack(lngPack).Links = 0 Then
        PopupMenu UserControl.mnuMain(2)
    Else
        ' Unload
        UserControl.mnuPackDynamic(0).Visible = True
        For i = UserControl.mnuPackDynamic.UBound To 1 Step -1
            Unload UserControl.mnuPackDynamic(i)
        Next
        ' Add dynamic items
        For i = 1 To db.Pack(lngPack).Links
            AddPackItem db.Pack(lngPack).Link(i).FullName
        Next
        If db.Pack(lngPack).Links Then AddPackItem "-"
        ' Add static items
        AddPackItem "Skip this Pack"
        AddPackItem "Don't Skip this Pack"
        AddPackItem "-"
        AddPackItem "Never Mind"
        UserControl.mnuPackDynamic(0).Visible = False
        PopupMenu UserControl.mnuMain(6)
    End If
End Sub

Private Sub AddPackItem(pstrCaption As String)
    Dim lngIndex As Long
    
    With UserControl
        lngIndex = .mnuPackDynamic.UBound + 1
        Load .mnuPackDynamic(lngIndex)
        With .mnuPackDynamic(lngIndex)
            .Caption = pstrCaption
            .Visible = True
        End With
    End With
End Sub

Private Sub mnuPack_Click(Index As Integer)
    Select Case UserControl.mnuPack(Index).Caption
        Case "Skip this Pack": SkipPack mtypRow(mlngRow).Pack, True
        Case "Don't Skip this Pack": SkipPack mtypRow(mlngRow).Pack, False
    End Select
End Sub

Private Sub SkipPack(ByVal pstrPack As String, pblnSkip As Boolean)
    Dim lngQuest As Long
    Dim i As Long
    
    For i = 1 To mlngRows
        If mtypRow(i).Pack = pstrPack Then
            lngQuest = mtypRow(i).QuestID
            If lngQuest <> 0 Then
                db.Quest(lngQuest).Skipped = pblnSkip
                mtypRow(i).Skipped = pblnSkip
            End If
        End If
    Next
    cfg.CompendiumScroll = UserControl.scrollVertical.Value
    Redraw
    UserControl.scrollVertical.Value = cfg.CompendiumScroll
    DirtyFlag dfeData
End Sub

Private Sub mnuPackDynamic_Click(Index As Integer)
    Select Case UserControl.mnuPackDynamic(Index).Caption
        Case "Skip this Pack": SkipPack mtypRow(mlngRow).Pack, True
        Case "Don't Skip this Pack": SkipPack mtypRow(mlngRow).Pack, False
        Case "Never Mind": Exit Sub
        Case Else: PackLink PackID(), UserControl.mnuPackDynamic(Index).Caption
    End Select
End Sub

Private Sub ProgressMenu()
    Dim lngQuest As Long
    Dim blnSolo As Boolean
    Dim blnRaid As Boolean
    Dim i As Long
    
    If mlngRow = 0 Or mlngCol = 0 Then Exit Sub
    lngQuest = mtypRow(mlngRow).QuestID
    Select Case db.Quest(lngQuest).Style
        Case qeSolo: blnSolo = True
        Case qeRaid: blnRaid = True
    End Select
    UserControl.mnuProgress(2).Visible = blnSolo
    UserControl.mnuProgress(3).Visible = Not (blnSolo Or blnRaid)
    For i = 4 To 6
        UserControl.mnuProgress(i).Visible = Not blnSolo
    Next
    PopupMenu UserControl.mnuMain(3)
End Sub

Private Sub mnuProgress_Click(Index As Integer)
    Dim enDifficulty As ProgressEnum
    Dim lngQuest As Long
    
    If UserControl.mnuProgress(Index).Caption = "Never Mind" Then Exit Sub
    enDifficulty = GetDifficultyID(UserControl.mnuProgress(Index).Caption)
    If mlngRow <> 0 Then
        lngQuest = mtypRow(mlngRow).QuestID
        SetDifficulty mlngRow, mlngCol, lngQuest, enDifficulty, True
        If enDifficulty <> peNone And enDifficulty <> peSolo Then cfg.Difficulty = enDifficulty
    End If
End Sub

Private Function GetDifficultyID(pstrDifficulty As String) As ProgressEnum
    Select Case LCase$(pstrDifficulty)
        Case "solo": GetDifficultyID = peSolo
        Case "casual": GetDifficultyID = peCasual
        Case "normal": GetDifficultyID = peNormal
        Case "hard": GetDifficultyID = peHard
        Case "elite": GetDifficultyID = peElite
    End Select
End Function


' ************* FAVOR TOTALS *************


Private Sub lnkChallengeFavor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    ActiveCell 0, 0
End Sub

Private Sub lnkChallengeFavor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkChallengeFavor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    frmChallenges.Character = Index
    OpenForm "frmChallenges"
End Sub

Private Sub lnkTotalFavor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    ActiveCell 0, 0
End Sub

Private Sub lnkTotalFavor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    ActiveCell 0, 0
End Sub

Private Sub lnkTotalFavor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    OpenForm "frmPatrons"
End Sub


' ************* COLUMN HEADER COMMANDS *************


Private Sub picHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GetColumn(X, Y) <> 0 Then xp.SetMouseCursor mcHand
    ActiveCell 0, 0
End Sub

Private Sub picHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    Dim lngCharacter As Long
    
    lngCol = GetColumn(X, Y)
    If lngCol = 0 Then Exit Sub
    xp.SetMouseCursor mcHand
    If Button = vbRightButton Then
        Select Case lngCol
            Case 1
                ShowLevelOrderMenu
                Exit Sub
            Case mlngCols - 2
                OpenForm "frmPatrons"
                Exit Sub
        End Select
    End If
    lngCharacter = lngCol - 4
    If lngCharacter < 1 Or lngCharacter > db.Characters Then Exit Sub
    Select Case Button
        Case vbLeftButton
            If Not RunLeftClickCommand(lngCharacter) Then
                frmCharacter.Character = lngCharacter
                OpenForm "frmCharacter"
            End If
        Case vbRightButton
            ContextMenu lngCharacter
    End Select
End Sub

Private Sub picHeader_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = GetColumn(X, Y)
    If lngCol = 0 Then Exit Sub
    xp.SetMouseCursor mcHand
    If Button = vbLeftButton Then
        Select Case lngCol
            Case 1: cfg.CompendiumOrder = coeLevel
            Case 2: cfg.CompendiumOrder = coeEpic
            Case 3: cfg.CompendiumOrder = coeQuest
            Case 4: cfg.CompendiumOrder = coePack
            Case mlngCols - 2: cfg.CompendiumOrder = coePatron
            Case mlngCols - 1: cfg.CompendiumOrder = coeFavor
            Case mlngCols: cfg.CompendiumOrder = coeStyle
            Case Else: Exit Sub
        End Select
        ReQuery
        DirtyFlag dfeSettings
    End If
End Sub

Private Function GetColumn(X As Single, Y As Single) As Long
    Dim i As Long
    
    For i = 1 To mlngCols
        If X < mtypCol(i).Left + mtypCol(i).Width Then
            GetColumn = i
            Exit For
        End If
    Next
End Function

Private Sub ShowLevelOrderMenu()
    Dim i As Long
    
    For i = 0 To 2
        UserControl.mnuLevel(i).Checked = (cfg.LevelSort = i)
    Next
    PopupMenu UserControl.mnuMain(4)
End Sub

Private Sub mnuLevel_Click(Index As Integer)
    Dim frm As Form
    
    cfg.CompendiumOrder = coeLevel
    cfg.LevelSort = Index
    ReQuery
    DirtyFlag dfeSettings
    If GetForm(frm, "frmTools") Then frm.LevelOrderChange
End Sub

Private Sub ContextMenu(plngCharacter As Long)
    Dim i As Long
    
    mlngCharacter = plngCharacter
    With db.Character(mlngCharacter).ContextMenu
        For i = 0 To .Commands - 1
            AddMenuItem i, .Command(i + 1).Caption
        Next
        If .Commands > 0 Then AddMenuItem i, "-", True
    End With
    AddMenuItem i, "Challenges", True
    AddMenuItem i, "Sagas", True
    AddMenuItem i, "Total Favor", True
    AddMenuItem i, "Reincarnate", True
    AddMenuItem i, "-", True
    AddMenuItem i, "Edit Characters", True
    AddMenuItem i, "Customize", True
    With UserControl
        For i = i To .mnuCharacter.UBound
            .mnuCharacter(i).Visible = False
        Next
    End With
    PopupMenu UserControl.mnuMain(0)
End Sub

Private Sub AddMenuItem(plngIndex As Long, pstrCaption As String, Optional pblnIncrement As Boolean = False)
    With UserControl
        If plngIndex > .mnuCharacter.UBound Then Load .mnuCharacter(plngIndex)
        With .mnuCharacter(plngIndex)
            .Caption = pstrCaption
            .Visible = True
        End With
    End With
    If pblnIncrement Then plngIndex = plngIndex + 1
End Sub

Private Sub mnuCharacter_Click(Index As Integer)
    Dim i As Long
    
    Select Case mnuCharacter(Index).Caption
        Case "Challenges"
            frmChallenges.Character = mlngCharacter
            OpenForm "frmChallenges"
        Case "Sagas"
            frmSagas.Character = mlngCharacter
            OpenForm "frmSagas"
        Case "Total Favor"
            OpenForm "frmPatrons"
        Case "Reincarnate"
            Reincarnate mlngCharacter
        Case "Edit Characters"
            frmCharacter.Character = mlngCharacter
            OpenForm "frmCharacter"
        Case "Customize"
            AutoSave
            gtypMenu = db.Character(mlngCharacter).ContextMenu
            gtypMenu.Title = db.Character(mlngCharacter).Character & " Context Menu"
            frmMenuEditor.Show vbModal, Me
            With db.Character(mlngCharacter)
                If gtypMenu.Accepted Then
                    .ContextMenu = gtypMenu
                    DirtyFlag dfeData
                End If
                .ContextMenu.Accepted = False
            End With
        Case Else
            With db.Character(mlngCharacter).ContextMenu
                For i = 1 To .Commands
                    If .Command(i).Caption = mnuCharacter(Index).Caption Then
                        RunCommand .Command(i)
                        Exit For
                    End If
                Next
            End With
    End Select
End Sub


' ************* SCROLLBARS *************


Public Sub WheelScroll(plngValue As Long)
    Dim lngValue As Long
    
    With UserControl.scrollVertical
        lngValue = .Value - plngValue
        If lngValue < .Min Then lngValue = .Min
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub

Public Property Get Scroll() As Long
    Scroll = UserControl.scrollVertical.Value
End Property

Public Property Let Scroll(ByVal plngScroll As Long)
    mblnLoaded = False
    With UserControl.scrollVertical
        If plngScroll >= .Min And plngScroll <= .Max And .Enabled = True Then .Value = plngScroll
    End With
    mblnLoaded = True
End Property

Private Sub scrollVertical_GotFocus()
    UserControl.picQuests.SetFocus
End Sub

Private Sub scrollVertical_Change()
    VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    If mblnOverride Then Exit Sub
    If mblnLoaded Then
        cfg.CompendiumScroll = UserControl.scrollVertical.Value
        DirtyFlag dfeSettings
    End If
    UserControl.picQuests.Top = 0 - ((UserControl.scrollVertical.Value - 1) * mlngRowHeight)
End Sub


' ************* DEBUG TOOLS *************


Public Function GetTextWidth(pstrText As String) As Long
    GetTextWidth = UserControl.picQuests.TextWidth(pstrText)
End Function

