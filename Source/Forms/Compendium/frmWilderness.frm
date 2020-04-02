VERSION 5.00
Begin VB.Form frmWilderness 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Wilderness Areas"
   ClientHeight    =   3996
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6744
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWilderness.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3996
   ScaleWidth      =   6744
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   540
      ScaleHeight     =   312
      ScaleWidth      =   1752
      TabIndex        =   0
      Top             =   540
      Width           =   1752
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1812
      Left            =   300
      ScaleHeight     =   1812
      ScaleWidth      =   2892
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   2892
      Begin VB.VScrollBar scrollVertical 
         Height          =   1512
         Left            =   2400
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   252
      End
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   540
         ScaleHeight     =   732
         ScaleWidth      =   1692
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   780
         Width           =   1692
      End
   End
End
Attribute VB_Name = "frmWilderness"
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

Private mblnLoaded As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnLoaded = False
    LoadData
    SizeClient
    cfg.Configure Me
    mblnLoaded = True
    DrawGrid
    If Not XP.DebugMode Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not XP.DebugMode Then Call WheelUnHook(Me.Hwnd)
    cfg.SavePosition Me, True
End Sub

Private Sub Form_Resize()
    If Not mblnLoaded Then Exit Sub
    If Me.WindowState = vbMinimized Or mlngRowHeight = 0 Then Exit Sub
    If Me.ScaleHeight < mlngRowHeight * 3 Then Exit Sub
    With Me.picHeader
        Me.picContainer.Move .Left, .Top + .Height, .Width + Me.scrollVertical.Width, Me.ScaleHeight - (.Top + .Height)
    End With
    ShowScrollbar
End Sub

Private Sub ShowScrollbar()
    Dim lngMax As Long

    lngMax = db.Areas - (Me.picContainer.Height \ mlngRowHeight)
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

Public Sub ReDrawForm()
    If mblnLoaded Then
        mblnLoaded = False
        LoadData
        SizeClient
        DrawGrid
        ShowScrollbar
        mblnLoaded = True
    End If
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
    Dim i As Long
    
    ReDim mtypCol(1 To 6)
    InitCol 1, "Area", vbLeftJustify
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
        .Width = Me.picClient.TextWidth(.Header)
        .Align = penAlign
    End With
End Sub

Private Sub CheckCol(plngCol As Long, pstrText As String)
    Dim lngWidth As Long
    
    With mtypCol(plngCol)
        lngWidth = Me.picClient.TextWidth(pstrText)
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
    mlngRowHeight = Me.picClient.TextHeight("Q") + mlngMarginY * 2
    mlngSpace = Me.picClient.TextWidth(" ")
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
    mlngMapRight = mlngMapLeft + Me.picClient.TextWidth("Map")
    mlngOtherLeft = mtypCol(4).Left + mlngMarginX + mlngSpace
    mlngPackLeft = mtypCol(6).Left + mlngMarginX + mlngSpace
    For i = 1 To db.Areas
        With mtypRow(i)
            .AreaRight = mlngAreaLeft + Me.picClient.TextWidth(.Area)
            .PackRight = mlngPackLeft + Me.picClient.TextWidth(.Pack)
            If Len(.Other) Then .OtherRight = mlngOtherLeft + Me.picClient.TextWidth(.Other)
        End With
    Next
    lngWidth = mtypCol(6).Right + PixelX
    With Me.picHeader
        .Move 0, 0, lngWidth, mlngRowHeight
    End With
    With Me.picClient
        lngHeight = mlngRowHeight * db.Areas + PixelY
        .Move 0, 0, lngWidth, lngHeight
        lngLeft = Me.scrollVertical.Width
        lngTop = mlngRowHeight \ 2
    End With
    Me.picHeader.Move lngLeft, lngTop, lngWidth, mlngRowHeight
    lngTop = lngTop + mlngRowHeight
    Me.picContainer.Move lngLeft, lngTop, lngWidth + Me.scrollVertical.Width, lngHeight
    With Me.scrollVertical
        .Move Me.picContainer.ScaleWidth - .Width, 0, .Width, Me.picContainer.ScaleHeight
        .Visible = False
    End With
    Me.Width = Me.Width - Me.ScaleWidth + Me.picContainer.Left + Me.picContainer.Width
    lngHeight = (Me.Height - Me.ScaleHeight) + Me.picContainer.Top + Me.picContainer.ScaleHeight + mlngRowHeight
    XP.GetDesktop 0, 0, 0, lngMaxHeight
    If lngHeight > lngMaxHeight Then
        Me.picContainer.Height = Me.picContainer.Height - (lngHeight - lngMaxHeight)
        lngHeight = lngMaxHeight
    End If
    Me.Height = lngHeight
End Sub


' ************* DRAWING *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    Me.picClient.Cls
    DrawHeaders
    For lngRow = 1 To db.Areas
        For lngCol = 1 To 6
            DrawCell lngRow, lngCol, False
        Next
    Next
End Sub

Private Sub DrawHeaders()
    Dim lngLeft As Long
    Dim i As Long
    
    Me.picHeader.Cls
    Me.picHeader.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    For i = 1 To 6
        With mtypCol(i)
            Select Case .Align
                Case vbLeftJustify: lngLeft = .Left + mlngSpace + mlngMarginX
                Case vbCenter: lngLeft = .Left + (.Width - Me.picClient.TextWidth(.Header)) \ 2
                Case vbRightJustify: lngLeft = .Right - mlngSpace - mlngMarginX - Me.picClient.TextWidth(.Header)
            End Select
        End With
        Me.picHeader.CurrentX = lngLeft
        Me.picHeader.CurrentY = mlngMarginY
        Me.picHeader.Print mtypCol(i).Header
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
        lngForeColor = cfg.GetColor(cgeWorkspace, cveText)
        lngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        lngBorderColor = cfg.GetColor(cgeWorkspace, cveBackground)
    Else
        If (plngCol < 3 Or plngCol = 4 Or plngCol = 6) And pblnActive Then lngForeColor = cfg.GetColor(cgeControls, cveTextLink) Else lngForeColor = cfg.GetColor(cgeControls, cveText)
        lngBackColor = cfg.GetColor(cgeControls, cveBackground)
        lngBorderColor = cfg.GetColor(cgeControls, cveBorderInterior)
    End If
    Me.picClient.ForeColor = lngForeColor
    Me.picClient.FillColor = lngBackColor
    lngLeft = mtypCol(plngCol).Left
    lngTop = (plngRow - 1) * mlngRowHeight
    lngRight = mtypCol(plngCol).Right
    lngBottom = lngTop + mlngRowHeight
    Me.picClient.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngBorderColor, B
    lngBorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
    lngRight = lngRight
    If plngCol = 1 Then Me.picClient.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), lngBorderColor
    If plngCol = 6 Then Me.picClient.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), lngBorderColor
    If plngRow = 1 Then
        Me.picClient.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngBorderColor
    ElseIf plngRow = db.Areas Then
        Me.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngBorderColor
    Else
        If mtypRow(plngRow).Group <> mtypRow(plngRow - 1).Group Then Me.picClient.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngBorderColor
        If mtypRow(plngRow).Group <> mtypRow(plngRow + 1).Group Then Me.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngBorderColor
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
            Case vbCenter: lngLeft = lngLeft + (mtypCol(plngCol).Width - Me.picClient.TextWidth(strDisplay)) \ 2
            Case vbRightJustify: lngLeft = mtypCol(plngCol).Right - mlngSpace - mlngMarginX - Me.picClient.TextWidth(strDisplay)
        End Select
        PrintText strDisplay, lngLeft, lngTop + mlngMarginY
    End With
End Sub

Private Sub PrintText(pstrText As String, plngX As Long, plngY As Long)
    Me.picClient.CurrentX = plngX
    Me.picClient.CurrentY = plngY
    Me.picClient.Print pstrText
End Sub

Private Sub PrintLevelRange(plngLow As Long, plngHigh As Long, plngLeft As Long, plngTop As Long, plngRight As Long)
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    lngWidth = Me.TextWidth("-")
    lngLeft = plngLeft + (plngRight - plngLeft - lngWidth) \ 2
    Me.picClient.ForeColor = cfg.GetColor(cgeControls, cveTextDim)
    PrintText "-", lngLeft, plngTop
    PrintText CStr(plngLow), lngLeft - Me.picClient.TextWidth(CStr(plngLow)), plngTop
    Me.picClient.ForeColor = cfg.GetColor(cgeControls, cveText)
    PrintText CStr(plngHigh), lngLeft + lngWidth, plngTop
End Sub


' ************* MOUSE *************


Private Sub picClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
End Sub

Private Sub picClient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngArea As Long
    Dim strLink As String
    Dim lngPack As Long
    
    ActiveCell X, Y
    If mlngRow = 0 Then Exit Sub
    lngArea = mtypRow(mlngRow).AreaID
    With db.Area(lngArea)
        Select Case mlngCol
            Case 1: XP.OpenURL MakeWiki(.Wiki)
            Case 2: XP.OpenURL WikiImage(.Map)
            Case 4: XP.OpenURL .Link(1).Target
            Case 6
                lngPack = SeekPack(.Pack)
                If lngPack Then XP.OpenURL MakeWiki(db.Pack(lngPack).Wiki)
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
    If lngRow <> 0 And lngCol <> 0 Then XP.SetMouseCursor mcHand
    If lngCol <> mlngCol Or lngRow <> mlngRow Then
        DrawCell mlngRow, mlngCol, False
        mlngRow = lngRow
        mlngCol = lngCol
        DrawCell mlngRow, mlngCol, True
    End If
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
    Me.picHeader.SetFocus
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

