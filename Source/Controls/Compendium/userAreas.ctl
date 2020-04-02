VERSION 5.00
Begin VB.UserControl userAreas 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   4296
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4296
   ScaleWidth      =   5160
   Begin VB.PictureBox picHeaderContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   552
      Left            =   540
      ScaleHeight     =   552
      ScaleWidth      =   3552
      TabIndex        =   4
      Top             =   300
      Visible         =   0   'False
      Width           =   3552
      Begin VB.PictureBox picHeader 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   240
         ScaleHeight     =   312
         ScaleWidth      =   2832
         TabIndex        =   5
         Top             =   120
         Width           =   2832
      End
   End
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   1140
      TabIndex        =   3
      Top             =   3300
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   2532
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1932
      Left            =   420
      ScaleHeight     =   1932
      ScaleWidth      =   3612
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3612
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1152
         Left            =   420
         ScaleHeight     =   1152
         ScaleWidth      =   2592
         TabIndex        =   1
         Top             =   360
         Width           =   2592
      End
   End
   Begin VB.Label lblControl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wilderness"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   948
   End
End
Attribute VB_Name = "userAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ColumnType
    Header As String
    Left As Long
    Width As Long
    Right As Long
    Align As AlignmentConstants
End Type

Private Type RowType
    AreaID As Long
    Area As String
    Lowest As Long
    Highest As Long
    Explorer As Long
    PackID As Long
    Pack As String
    Other As String
    OtherLink As String
    AreaRight As Long
    PackRight As Long
    OtherRight As Long
    Group As Long
End Type

Private mtypRow() As RowType
Private mtypCol() As ColumnType

Private mlngRow As Long
Private mlngCol As Long

Private mlngAreaLeft As Long
Private mlngMapLeft As Long
Private mlngMapRight As Long
Private mlngOtherLeft As Long
Private mlngPackLeft As Long
Private mlngSpace As Long

Private mlngMarginX As Long
Private mlngMarginY As Long
Private mlngRowHeight As Long
Private mlngScrollX As Long

Private mblnLoaded As Boolean


' ************* GENERAL *************


Private Sub UserControl_Initialize()
    mblnLoaded = False
    mlngScrollX = UserControl.TextWidth("XXX")
End Sub

Public Sub Init()
    mblnLoaded = False
    
    With UserControl
        .lblControl.Visible = False
        .picHeaderContainer.Visible = True
        .picContainer.Visible = True
    End With
    LoadData
    SizeClient
    RefreshColors
    mblnLoaded = True
    DrawGrid
End Sub

Private Sub RefreshColors()
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .picHeader.BackColor = cfg.GetColor(cgeDropSlots, cveBackground)
        .picHeaderContainer.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .picContainer.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .picClient.BackColor = cfg.GetColor(cgeControls, cveBackground)
    End With
End Sub

Private Sub UserControl_Resize()
    Dim lngHeight As Long
    
    If Not mblnLoaded Then Exit Sub
    If mlngRowHeight = 0 Then Exit Sub
    If UserControl.ScaleHeight < mlngRowHeight * 3 Then Exit Sub
    ShowScrollbars
End Sub

Public Sub Redraw()
    If mblnLoaded Then
        RefreshColors
        mblnLoaded = False
        LoadData
        SizeClient
        DrawGrid
        ShowScrollbars
        mblnLoaded = True
    End If
End Sub

Public Property Get Hwnd() As Long
    Hwnd = UserControl.picContainer.Hwnd
End Property


' ************* DATA *************


Private Sub LoadData()
    Dim i As Long
    
    ReDim mtypCol(1 To 6)
    InitCol 1, "Wilderness Area", vbLeftJustify
    InitCol 2, "Map", vbCenter
    InitCol 3, "Levels", vbCenter
    InitCol 4, "Other", vbLeftJustify
    InitCol 5, "Explorer", vbRightJustify
    InitCol 6, "Adventure Pack", vbLeftJustify
    ReDim mtypRow(db.Areas)
    For i = 1 To db.Areas
        CheckCol 1, db.Area(i).Area
        CheckCol 6, db.Area(i).Pack
        With mtypRow(db.Area(i).Order)
            .AreaID = i
            .Area = db.Area(i).Area
            .Lowest = db.Area(i).Lowest
            Select Case .Lowest
                Case Is < 10: .Group = 1
                Case 10 To 19: .Group = 2
                Case Is >= 20: .Group = 3
            End Select
            .Highest = db.Area(i).Highest
            .Explorer = db.Area(i).Explorer
            If db.Area(i).Links Then
                .Other = db.Area(i).Link(1).Abbreviation
                .OtherLink = db.Area(i).Link(1).Target
                CheckCol 4, .Other
            End If
            .Pack = db.Area(i).Pack
            .PackID = SeekPack(.Pack)
        End With
    Next
End Sub

Private Sub InitCol(plngCol As Long, pstrHeader As String, penAlign As AlignmentConstants)
    With mtypCol(plngCol)
        .Header = pstrHeader
        .Width = UserControl.picClient.TextWidth(.Header)
        .Align = penAlign
    End With
End Sub

Private Sub CheckCol(plngCol As Long, pstrText As String)
    Dim lngWidth As Long
    
    With mtypCol(plngCol)
        lngWidth = UserControl.picClient.TextWidth(pstrText)
        If .Width < lngWidth Then .Width = lngWidth
    End With
End Sub


' ************* SIZING *************


Private Sub SizeClient()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngMaxHeight As Long
    Dim i As Long
    
    mlngMarginX = cfg.MarginX * PixelX
    mlngMarginY = cfg.MarginY * PixelY
    mlngRowHeight = UserControl.picClient.TextHeight("Q") + mlngMarginY * 2
    mlngSpace = UserControl.picClient.TextWidth(" ")
    For i = 1 To 6
        With mtypCol(i)
            .Left = lngLeft
            .Width = .Width + (mlngMarginX + mlngSpace) * 2
            .Right = .Left + .Width
            lngLeft = .Right
        End With
    Next
    mlngAreaLeft = mtypCol(1).Left + mlngMarginX + mlngSpace
    mlngMapLeft = mtypCol(2).Left + mlngMarginX + mlngSpace
    mlngMapRight = mlngMapLeft + UserControl.picClient.TextWidth("Map")
    mlngOtherLeft = mtypCol(4).Left + mlngMarginX + mlngSpace
    mlngPackLeft = mtypCol(6).Left + mlngMarginX + mlngSpace
    For i = 1 To db.Areas
        With mtypRow(i)
            .AreaRight = mlngAreaLeft + UserControl.picClient.TextWidth(.Area)
            .PackRight = mlngPackLeft + UserControl.picClient.TextWidth(.Pack)
            If Len(.Other) Then .OtherRight = mlngOtherLeft + UserControl.picClient.TextWidth(.Other)
        End With
    Next
    lngWidth = mtypCol(6).Right + PixelX
    With UserControl.picClient
        lngHeight = mlngRowHeight * db.Areas + PixelY
        .Move 0, 0, lngWidth, lngHeight
        lngLeft = 0
        lngTop = 0
    End With
    UserControl.picHeaderContainer.Move lngLeft, lngTop, lngWidth, mlngRowHeight + PixelY
    UserControl.picHeader.Move 0, 0, lngWidth, mlngRowHeight + PixelY
    lngTop = lngTop + mlngRowHeight
    UserControl.picContainer.Move lngLeft, lngTop, lngWidth + UserControl.scrollVertical.Width, lngHeight
    With UserControl.scrollVertical
        .Move UserControl.picContainer.ScaleWidth - .Width, 0, .Width, UserControl.picContainer.ScaleHeight
        .Visible = False
    End With
End Sub

Private Sub ShowScrollbars()
    Dim blnVertical As Boolean
    Dim blnHorizontal As Boolean
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngValue As Long
    Dim lngHeaderWidth As Long
    
    With UserControl
        lngLeft = 0
        lngTop = .picHeaderContainer.Height - PixelY
        lngWidth = .picClient.Width
        lngHeaderWidth = .picClient.Width
        lngHeight = .picClient.Height
        blnHorizontal = (.ScaleWidth < lngWidth)
        blnVertical = (.ScaleHeight < lngTop + lngHeight)
        If blnHorizontal And Not blnVertical Then
            blnVertical = (.ScaleHeight < lngTop + lngHeight + .scrollHorizontal.Height)
        ElseIf blnVertical And Not blnHorizontal Then
            blnHorizontal = (.ScaleWidth < lngWidth + .scrollVertical.Width)
        End If
        If blnHorizontal Then lngWidth = .ScaleWidth
        If blnVertical Then
            If lngWidth > .ScaleWidth - .scrollVertical.Width Then
                lngWidth = .ScaleWidth - .scrollVertical.Width
                If .ScaleWidth > .picClient.Width Then lngHeaderWidth = .picClient.Width Else lngHeaderWidth = .ScaleWidth
            End If
            lngHeight = .ScaleHeight - lngTop
            lngHeight = ((lngHeight \ mlngRowHeight) * mlngRowHeight) + PixelY
            If blnHorizontal Then If blnVertical Then lngHeight = lngHeight - .scrollHorizontal.Height
        End If
        .picHeaderContainer.Width = lngHeaderWidth
        .picContainer.Move lngLeft, lngTop, lngWidth, lngHeight
        If blnHorizontal Then
            lngValue = .picClient.Width
            With .scrollHorizontal
                .Move lngLeft, lngTop + lngHeight, lngWidth
                .Max = ((lngValue - lngWidth) \ mlngScrollX) + 1
                .Value = 0
                .LargeChange = lngWidth \ mlngScrollX
            End With
        End If
        .scrollHorizontal.Visible = blnHorizontal
        If blnVertical Then
            lngValue = .picClient.Height
            With .scrollVertical
                .Move lngLeft + lngWidth, lngTop, .Width, lngHeight
                .Max = (lngValue - lngHeight) \ mlngRowHeight
                If blnHorizontal Then .Max = .Max + 1
                .Value = 0
                .LargeChange = lngHeight \ mlngRowHeight
            End With
        End If
        .scrollVertical.Visible = blnVertical
    End With
End Sub


' ************* DRAWING *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    UserControl.picClient.Cls
    DrawHeaders
    For lngRow = 1 To db.Areas
        For lngCol = 1 To 6
            DrawCell lngRow, lngCol, False
        Next
    Next
End Sub

Private Sub DrawHeaders()
    Dim lngLeft As Long
    Dim lngBorderColor As Long
    Dim i As Long
    
    UserControl.picHeader.Cls
    UserControl.picHeader.ForeColor = cfg.GetColor(cgeDropSlots, cveText)
    UserControl.picHeader.FillColor = cfg.GetColor(cgeDropSlots, cveBackground)
    lngBorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
    For i = 1 To 6
        With mtypCol(i)
            Select Case .Align
                Case vbLeftJustify: lngLeft = .Left + mlngSpace + mlngMarginX
                Case vbCenter: lngLeft = .Left + (.Width - UserControl.picClient.TextWidth(.Header)) \ 2
                Case vbRightJustify: lngLeft = .Right - mlngSpace - mlngMarginX - UserControl.picClient.TextWidth(.Header)
            End Select
            UserControl.picHeader.Line (.Left, 0)-(.Right, mlngRowHeight), lngBorderColor, B
        End With
        UserControl.picHeader.CurrentX = lngLeft
        UserControl.picHeader.CurrentY = mlngMarginY
        UserControl.picHeader.Print mtypCol(i).Header
    Next
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, Optional pblnActive As Boolean)
    Dim lngForeColor As Long
    Dim lngBackColor As Long
    Dim lngBorderColor As Long
    Dim strDisplay As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    
    If plngRow = 0 Then
        If plngCol = 0 Then Exit Sub
        lngForeColor = cfg.GetColor(cgeDropSlots, cveText)
        lngBackColor = cfg.GetColor(cgeDropSlots, cveBackground)
        lngBorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
    Else
        If (plngCol < 3 Or plngCol = 4 Or plngCol = 6) And pblnActive Then lngForeColor = cfg.GetColor(cgeControls, cveTextLink) Else lngForeColor = cfg.GetColor(cgeControls, cveText)
        lngBackColor = cfg.GetColor(cgeControls, cveBackground)
        lngBorderColor = cfg.GetColor(cgeControls, cveBorderInterior)
    End If
    UserControl.picClient.ForeColor = lngForeColor
    UserControl.picClient.FillColor = lngBackColor
    lngLeft = mtypCol(plngCol).Left
    lngTop = (plngRow - 1) * mlngRowHeight
    lngRight = mtypCol(plngCol).Right
    lngBottom = lngTop + mlngRowHeight
    UserControl.picClient.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngBorderColor, B
    lngBorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
    lngRight = lngRight
    If plngCol = 1 Then UserControl.picClient.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), lngBorderColor
    If plngCol = 6 Then UserControl.picClient.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), lngBorderColor
    If plngRow = 1 Then
        UserControl.picClient.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngBorderColor
    ElseIf plngRow = db.Areas Then
        UserControl.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngBorderColor
    Else
        If mtypRow(plngRow).Group <> mtypRow(plngRow - 1).Group Then UserControl.picClient.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngBorderColor
        If mtypRow(plngRow).Group <> mtypRow(plngRow + 1).Group Then UserControl.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngBorderColor
    End If
    With mtypRow(plngRow)
        Select Case plngCol
            Case 1: strDisplay = .Area
            Case 2: strDisplay = "Map"
            Case 3: PrintLevelRange .Lowest, .Highest, lngLeft + mlngMarginX, lngTop + mlngMarginY, mtypCol(3).Right - mlngMarginX
            Case 4: strDisplay = .Other
            Case 5: strDisplay = Format(.Explorer, "#,##0")
            Case 6: strDisplay = .Pack
        End Select
        If Len(strDisplay) = 0 Then Exit Sub
        Select Case mtypCol(plngCol).Align
            Case vbLeftJustify: lngLeft = lngLeft + mlngSpace + mlngMarginX
            Case vbCenter: lngLeft = lngLeft + (mtypCol(plngCol).Width - UserControl.picClient.TextWidth(strDisplay)) \ 2
            Case vbRightJustify: lngLeft = mtypCol(plngCol).Right - mlngSpace - mlngMarginX - UserControl.picClient.TextWidth(strDisplay)
        End Select
        PrintText strDisplay, lngLeft, lngTop + mlngMarginY
    End With
End Sub

Private Sub PrintText(pstrText As String, plngX As Long, plngY As Long)
    UserControl.picClient.CurrentX = plngX
    UserControl.picClient.CurrentY = plngY
    UserControl.picClient.Print pstrText
End Sub

Private Sub PrintLevelRange(plngLow As Long, plngHigh As Long, plngLeft As Long, plngTop As Long, plngRight As Long)
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    lngWidth = UserControl.TextWidth("-")
    lngLeft = plngLeft + (plngRight - plngLeft - lngWidth) \ 2
    UserControl.picClient.ForeColor = cfg.GetColor(cgeControls, cveTextDim)
    PrintText "-", lngLeft, plngTop
    PrintText CStr(plngLow), lngLeft - UserControl.picClient.TextWidth(CStr(plngLow)), plngTop
    UserControl.picClient.ForeColor = cfg.GetColor(cgeControls, cveText)
    PrintText CStr(plngHigh), lngLeft + lngWidth, plngTop
End Sub


' ************* MOUSE *************


Private Sub picHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell 0, 0
End Sub

Private Sub picClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
End Sub

Private Sub picClient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngArea As Long
    Dim strLink As String
    Dim lngPack As Long
    
    ActiveCell X, Y
    If mlngRow = 0 Then Exit Sub
    If Button <> vbLeftButton Then Exit Sub
    lngArea = mtypRow(mlngRow).AreaID
    With db.Area(lngArea)
        Select Case mlngCol
            Case 1: xp.OpenURL MakeWiki(.Wiki)
            Case 2: xp.OpenURL WikiImage(.Map)
            Case 4: xp.OpenURL .Link(1).Target
            Case 6
                lngPack = SeekPack(.Pack)
                If lngPack Then xp.OpenURL MakeWiki(db.Pack(lngPack).Wiki)
        End Select
    End With
End Sub

Private Sub picClient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
End Sub

Private Sub ActiveCell(X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    lngRow = (Y \ mlngRowHeight) + 1
    If lngRow < 1 Or lngRow > db.Areas Then
        lngRow = 0
        lngCol = 0
    Else
        If X >= mlngAreaLeft And X <= mtypRow(lngRow).AreaRight Then
            lngCol = 1
        ElseIf X >= mlngMapLeft And X <= mlngMapRight Then
            lngCol = 2
        ElseIf X >= mlngPackLeft And X <= mtypRow(lngRow).PackRight Then
            lngCol = 6
        ElseIf Len(mtypRow(lngRow).Other) Then
            If X >= mlngOtherLeft And X <= mtypRow(lngRow).OtherRight Then lngCol = 4
        End If
        If lngCol = 0 Then lngRow = 0
    End If
    If lngRow <> 0 And lngCol <> 0 Then xp.SetMouseCursor mcHand
    If lngCol <> mlngCol Or lngRow <> mlngRow Then
        DrawCell mlngRow, mlngCol, False
        mlngRow = lngRow
        mlngCol = lngCol
        DrawCell mlngRow, mlngCol, True
    End If
End Sub


' ************* SCROLLBARS *************


Private Sub scrollHorizontal_GotFocus()
    UserControl.picHeaderContainer.SetFocus
End Sub

Private Sub scrollHorizontal_Change()
    HorizontalScroll
End Sub

Private Sub scrollHorizontal_Scroll()
    HorizontalScroll
End Sub

Private Sub HorizontalScroll()
    Dim lngLeft As Long
    
    lngLeft = 0 - UserControl.scrollHorizontal.Value * mlngScrollX
    UserControl.picHeader.Left = lngLeft
    UserControl.picClient.Left = lngLeft
End Sub

Public Sub WheelScroll(plngIncrement As Long)
    Dim lngValue As Long
    
    With UserControl.scrollVertical
        If .Visible Then
            lngValue = .Value - plngIncrement
            If lngValue < .Min Then lngValue = .Min
            If lngValue > .Max Then lngValue = .Max
            If .Value <> lngValue Then .Value = lngValue
        End If
    End With
End Sub

Private Sub scrollVertical_GotFocus()
    UserControl.picHeader.SetFocus
End Sub

Private Sub scrollVertical_Change()
    VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    UserControl.picClient.Top = 0 - UserControl.scrollVertical.Value * mlngRowHeight
End Sub

