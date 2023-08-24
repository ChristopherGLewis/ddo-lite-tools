VERSION 5.00
Begin VB.UserControl userRaceCombo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2415
   ScaleWidth      =   5490
End
Attribute VB_Name = "userRaceCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click(Race As Long)

Private Type RowType
    Race As RaceEnum
    Stats(1 To MAX_STATS) As Long
    Class As String
    Caption As String
End Type

Private mlngTextColor As Long
Private mlngBackColor As Long
Private mlngTextColorHighlight As Long
Private mlngBackColorHighlight As Long
Private mlngTextColorHeader As Long
Private mlngBackColorHeader As Long
Private mlngBorderColor As Long

Private mtypRow() As RowType
Private mlngRows As Long

Private mlngLeft As Long
Private mlngMarginX As Long
Private mlngMarginY As Long
Private mlngStatMargin As Long
Private mlngRaceWidth As Long
Private mlngStatWidth As Long
Private mlngCenterPos As Long
Private mlngCenterNeg As Long
Private mlngClassWidth As Long
Private mlngClassLeft As Long

Private mstrStatHeader() As String
Private mlngClassHeaderRow As Long

Private mlngTextHeight As Long
Private mlngPixelX As Long
Private mlngPixelY As Long

Private mlngRace As Long
Private mlngActive As Long

Private mblnInitialized As Boolean


' ************* METHODS AND PROPERTIES *************


Public Sub RefreshColors()
    mlngTextColor = cfg.GetColor(cgeControls, cveText)
    mlngBackColor = cfg.GetColor(cgeControls, cveBackground)
    mlngTextColorHighlight = vbHighlightText
    mlngBackColorHighlight = vbHighlight
    mlngTextColorHeader = cfg.GetColor(cgeDropSlots, cveText)
    mlngBackColorHeader = cfg.GetColor(cgeDropSlots, cveBackground)
    mlngBorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
    DrawList mlngActive
End Sub

Public Sub DropDown(ByVal penRace As RaceEnum)
    If Not mblnInitialized Then
        InitRows
        InitDimensions
        ResizeList
        mblnInitialized = True
    End If
    DrawList penRace
End Sub

Public Sub KeyDown(ByVal KeyCode As Integer)
    Dim lngActive As Long
    Dim lngRace As Long
    
    Select Case KeyCode
        Case vbKeyUp
            For lngActive = mlngActive - 1 To 2 Step -1
                If mtypRow(lngActive).Race <> reAny Then Exit For
            Next
            If lngActive < 2 Then lngActive = 2
        Case vbKeyDown
            For lngActive = mlngActive + 1 To mlngRows
                If mtypRow(lngActive).Race <> reAny Then Exit For
            Next
            If lngActive > mlngRows Then lngActive = mlngRows
        Case vbKeyPageUp, vbKeyHome
            lngActive = 2
        Case vbKeyPageDown, vbKeyEnd
            lngActive = mlngRows
        Case vbKeyReturn
            lngRace = mtypRow(mlngActive).Race
            RaiseEvent Click(lngRace)
            Exit Sub
        Case Else
            Exit Sub
    End Select
    If mlngActive <> lngActive Then
        DrawRow mlngActive, False
        mlngActive = lngActive
        DrawRow mlngActive, True
    End If
End Sub


' ************* DRAWING *************


Private Sub ResizeList()
    UserControl.Width = mlngClassLeft + mlngClassWidth + UserControl.TextWidth(" ")
    UserControl.Height = mlngTextHeight * mlngRows + mlngPixelY
End Sub

Private Sub DrawList(penRace As RaceEnum)
    Dim i As Long
    
    mlngActive = 0
    UserControl.Cls
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - mlngPixelX, UserControl.ScaleHeight - mlngPixelY), mlngBorderColor, B
    For i = 1 To mlngRows
        DrawRow i, (mtypRow(i).Race = penRace)
        If penRace <> reAny And penRace = mtypRow(i).Race Then mlngActive = i
    Next
End Sub

Private Sub DrawRow(plngRow As Long, pblnHighlight As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngForeColor As Long
    Dim lngBackColor As Long
    Dim i As Long
    
    If plngRow = 0 Then Exit Sub
    ' Calculate dimensions of this row
    lngLeft = mlngPixelX
    lngTop = (plngRow - 1) * mlngTextHeight
    lngRight = UserControl.ScaleWidth - mlngPixelX * 2
    lngBottom = lngTop + mlngTextHeight
    ' Identify colors
    If mtypRow(plngRow).Race = reAny Then
        lngForeColor = mlngTextColorHeader
        lngBackColor = mlngBackColorHeader
    ElseIf pblnHighlight Then
        lngForeColor = mlngTextColorHighlight
        lngBackColor = mlngBackColorHighlight
    Else
        lngForeColor = mlngTextColor
        lngBackColor = mlngBackColor
    End If
    ' Clear area, redrawing borders as needed
    UserControl.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngBackColor, BF
    lngRight = lngRight + mlngPixelX
    If mtypRow(plngRow).Race = reAny Then
        UserControl.Line (lngLeft, lngTop)-(lngRight, lngTop), mlngBorderColor
    ElseIf mtypRow(plngRow - 1).Race = reAny Then
        UserControl.Line (lngLeft, lngTop)-(lngRight, lngTop), mlngBorderColor
    End If
    If mtypRow(plngRow).Race = reAny Or plngRow = mlngRows Then
        UserControl.Line (lngLeft, lngBottom)-(lngRight, lngBottom), mlngBorderColor
    ElseIf mtypRow(plngRow + 1).Race = reAny Then
        UserControl.Line (lngLeft, lngBottom)-(lngRight, lngBottom), mlngBorderColor
    End If
    ' Draw text...
    UserControl.CurrentX = mlngLeft
    UserControl.CurrentY = lngTop
    UserControl.ForeColor = lngForeColor
    With mtypRow(plngRow)
        ' Draw Header
        If .Race = reAny Then
            UserControl.Print .Caption;
            lngLeft = mlngRaceWidth + mlngMarginX
            For i = 1 To 6
                UserControl.CurrentX = lngLeft + (mlngStatWidth - UserControl.TextWidth(mstrStatHeader(i))) \ 2
                UserControl.Print mstrStatHeader(i);
                lngLeft = lngLeft + mlngStatWidth + mlngStatMargin
            Next
            If plngRow = mlngClassHeaderRow Then
                UserControl.CurrentX = mlngClassLeft
                UserControl.Print "Class";
            End If
        Else
            ' Draw Race
            UserControl.Print GetRaceName(.Race);
            For i = 1 To 6
                Select Case .Stats(i)
                    Case Is < 0
                        UserControl.CurrentX = mlngRaceWidth + mlngMarginX + (i - 1) * (mlngStatWidth + mlngStatMargin) + mlngCenterNeg
                        UserControl.Print .Stats(i);
                    Case Is > 0
                        UserControl.CurrentX = mlngRaceWidth + mlngMarginX + (i - 1) * (mlngStatWidth + mlngStatMargin) + mlngCenterPos
                        UserControl.Print "+" & .Stats(i);
                End Select
            Next
            If Len(.Class) Then
                UserControl.CurrentX = mlngClassLeft
                UserControl.Print .Class;
            End If
        End If
    End With
End Sub


' ************* MOUSE *************


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    lngRow = ActiveRow(X, Y)
    If mlngActive <> lngRow Then
        DrawRow mlngActive, False
        mlngActive = lngRow
        DrawRow mlngActive, True
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngRace As Long
    
    lngRow = ActiveRow(X, Y)
    If lngRow Then lngRace = mtypRow(lngRow).Race
    RaiseEvent Click(lngRace)
End Sub

Private Function ActiveRow(X As Single, Y As Single) As Long
    Dim lngRow As Long
    
    lngRow = (Y \ mlngTextHeight) + 1
    If lngRow < 1 Or lngRow > mlngRows Then lngRow = 0
    ActiveRow = lngRow
End Function


' ************* INITIALIZE *************


Private Sub UserControl_Initialize()
    InitColors
    mblnInitialized = False
End Sub

Private Sub InitColors()
    mlngTextColor = vbWindowText
    mlngBackColor = vbWindowBackground
    mlngTextColorHighlight = vbHighlightText
    mlngBackColorHighlight = vbHighlight
    mlngTextColorHeader = vbButtonText
    mlngBackColorHeader = vbButtonFace
    mlngBorderColor = vbBlack
End Sub

Private Sub InitRows()
    Dim enRace() As RaceEnum
    Dim enCurrent As RaceTypeEnum
    Dim enPrevious As RaceTypeEnum
    Dim i As Long
    
    Erase mtypRow
    mlngRows = 0
    mlngRaceWidth = 0
    mlngClassWidth = 0
    enRace = GetRaceList()
    For i = 1 To reRaces - 1
        enCurrent = db.Race(enRace(i)).Type
        If enCurrent <> enPrevious Then AddRow reAny, GetRaceTypeName(enCurrent), (enCurrent = rteIconic)
        enPrevious = enCurrent
        AddRow enRace(i)
    Next
End Sub

Private Function GetRaceList() As RaceEnum()
    Dim enRace() As RaceEnum
    Dim enSwap As RaceEnum
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    
    ReDim enRace(1 To reRaces - 1)
    For i = 1 To reRaces - 1
        enRace(i) = i
    Next
    ' Insertion sort
    iMin = 2
    iMax = reRaces - 1
    For i = iMin To iMax
        enSwap = enRace(i)
        For j = i To iMin Step -1
            If CompareRace(enSwap, enRace(j - 1)) = -1 Then enRace(j) = enRace(j - 1) Else Exit For
        Next j
        enRace(j) = enSwap
    Next i
    GetRaceList = enRace
End Function

Private Function CompareRace(penLeft As RaceEnum, penRight As RaceEnum) As Long
    If db.Race(penLeft).Type < db.Race(penRight).Type Then
        CompareRace = -1
    ElseIf db.Race(penLeft).Type > db.Race(penRight).Type Then
        CompareRace = 1
    ElseIf db.Race(penLeft).ListFirst = db.Race(penRight).ListFirst Then
        If db.Race(penLeft).RaceName < db.Race(penRight).RaceName Then CompareRace = -1 Else CompareRace = 1
    ElseIf db.Race(penLeft).ListFirst Then
        CompareRace = -1
    Else
        CompareRace = 1
    End If
End Function

Private Sub AddRow(penRace As RaceEnum, Optional pstrCaption As String, Optional pblnClassHeaderRow As Boolean)
    Dim lngWidth As Long
    Dim enClass As ClassEnum
    Dim i As Long
    
    mlngRows = mlngRows + 1
    ReDim Preserve mtypRow(1 To mlngRows)
    If pblnClassHeaderRow Then mlngClassHeaderRow = mlngRows
    With mtypRow(mlngRows)
        If penRace = reAny Then
            .Caption = pstrCaption
        Else
            .Race = penRace
            .Caption = GetRaceName(penRace)
            lngWidth = UserControl.TextWidth(.Caption)
            If mlngRaceWidth < lngWidth Then mlngRaceWidth = lngWidth
            For i = 1 To 6
                .Stats(i) = db.Race(penRace).Stats(i) - 8
            Next
            enClass = db.Race(penRace).IconicClass
            If enClass <> ceAny Then
                .Class = GetClassName(enClass)
                lngWidth = UserControl.TextWidth(.Class)
                If mlngClassWidth < lngWidth Then mlngClassWidth = lngWidth
            End If
        End If
    End With
End Sub

Private Sub InitDimensions()
    Dim lngWidth As Long
    Dim i As Long
    
    ' General dimensions
    mlngPixelX = Screen.TwipsPerPixelX
    mlngPixelY = Screen.TwipsPerPixelY
    mlngMarginX = UserControl.TextWidth("  ")
    mlngMarginY = mlngPixelY
    mlngTextHeight = UserControl.TextHeight("Q") + mlngMarginY
    mlngLeft = UserControl.TextWidth(" ")
    ' Stat width (all same as widest)
    mlngStatWidth = 0
    mstrStatHeader = Split(".Str.Dex.Con.Int.Wis.Cha", ".")
    For i = 1 To 6
        lngWidth = UserControl.TextWidth(mstrStatHeader(i))
        If mlngStatWidth < lngWidth Then mlngStatWidth = lngWidth
    Next
    mlngStatMargin = UserControl.TextWidth(" ")
    mlngCenterPos = (mlngStatWidth - UserControl.TextWidth("+2")) \ 2
    mlngCenterNeg = (mlngStatWidth - UserControl.TextWidth("-2")) \ 2
    mlngClassLeft = mlngLeft + mlngRaceWidth + (mlngMarginX * 2) + (mlngStatWidth * 6) + (mlngStatMargin * 5)
End Sub
