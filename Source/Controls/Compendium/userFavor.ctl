VERSION 5.00
Begin VB.UserControl userFavor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2808
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2808
   ScaleWidth      =   5520
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   540
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   3612
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   1752
      Left            =   4800
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1932
      Left            =   0
      ScaleHeight     =   1932
      ScaleWidth      =   4512
      TabIndex        =   0
      Top             =   60
      Width           =   4512
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1272
         Left            =   360
         ScaleHeight     =   1272
         ScaleWidth      =   3912
         TabIndex        =   1
         Top             =   240
         Width           =   3912
      End
   End
End
Attribute VB_Name = "userFavor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Type RowType
    Patron As String
    Favor() As Long
    Total As Long
    PatronLeft As Long
    PatronRight As Long
End Type

Private mtypRow() As RowType ' row 0 is totals row
Private mlngRowHeight As Long

Private mlngPatronWidth As Long
Private mlngWidth As Long

Private mlngMarginX As Long
Private mlngMarginY As Long
Private mlngSides As Long

Private mlngActiveRow As Long

Private mblnOverride As Boolean
Private mblnInit As Boolean
Private mlngCharacter As Long

Private mlngFitWidth As Long
Private mlngFitHeight As Long

Private mlngGridText As Long
Private mlngGridBack As Long
Private mlngBackText As Long
Private mlngBackground As Long


' ************* USERCONTROL *************


Private Sub UserControl_Initialize()
    mblnInit = False
    mlngCharacter = 0
End Sub

Public Sub Init(plngCharacter As Long)
    mlngCharacter = plngCharacter
    mblnInit = True
    ReDrawControl
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .picContainer.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
    ShowScrollbars
End Sub

Public Property Get Character() As Long
    Character = mlngCharacter
End Property

Public Property Let Character(plngCharacter As Long)
    mlngCharacter = plngCharacter
End Property

Public Property Get FitWidth() As Long
    FitWidth = mlngFitWidth
End Property

Public Property Get FitHeight() As Long
    FitHeight = mlngFitHeight
End Property

Public Sub ReDrawControl()
    mlngGridText = cfg.GetColor(cgeControls, cveText)
    mlngGridBack = cfg.GetColor(cgeControls, cveBackground)
    mlngBackText = cfg.GetColor(cgeWorkspace, cveText)
    mlngBackground = cfg.GetColor(cgeWorkspace, cveBackground)
    With UserControl
        .ForeColor = mlngBackText
        .BackColor = mlngBackground
        With .picContainer
            .ForeColor = mlngBackText
            .BackColor = mlngBackground
        End With
        With .picClient
            .ForeColor = mlngBackText
            .BackColor = mlngBackground
        End With
    End With
    LoadData
    SizeControl
    DrawControl
End Sub

Public Sub Recalculate()
    LoadData
    DrawControl
End Sub


' ************* DATA *************


Private Sub LoadData()
    Dim lngPatron As Long
    Dim lngRow As Long
    Dim lngFavor As Long
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim i As Long
    Dim c As Long
    
    If mlngCharacter = 0 Then
        lngFirst = 1
        lngLast = db.Characters
    Else
        lngFirst = mlngCharacter
        lngLast = mlngCharacter
    End If
    ReDim mtypRow(db.Patrons)
    ReDim mtypRow(0).Favor(db.Characters)
    If db.Patrons = 0 Then Exit Sub
    For i = 1 To db.Patrons
        With mtypRow(db.Patron(i).Order)
            .Patron = db.Patron(i).Patron
            ReDim .Favor(db.Characters)
        End With
    Next
    For i = 1 To db.Quests
        lngPatron = SeekPatron(db.Quest(i).Patron)
        lngRow = db.Patron(lngPatron).Order
        lngFavor = QuestFavor(i, peElite)
        mtypRow(lngRow).Total = mtypRow(lngRow).Total + lngFavor
        mtypRow(0).Total = mtypRow(0).Total + lngFavor
        For c = lngFirst To lngLast
            lngFavor = QuestFavor(i, db.Quest(i).Progress(c))
            mtypRow(lngRow).Favor(c) = mtypRow(lngRow).Favor(c) + lngFavor
            mtypRow(0).Favor(c) = mtypRow(0).Favor(c) + lngFavor
        Next
    Next
    For i = 1 To db.Challenges
        lngPatron = SeekPatron(db.Challenge(i).Patron)
        lngRow = db.Patron(lngPatron).Order
        mtypRow(lngRow).Total = mtypRow(lngRow).Total + db.Challenge(i).MaxStars
        mtypRow(0).Total = mtypRow(0).Total + db.Challenge(i).MaxStars
        For c = lngFirst To lngLast
            lngFavor = db.Challenge(i).Stars(c)
            mtypRow(lngRow).Favor(c) = mtypRow(lngRow).Favor(c) + lngFavor
            mtypRow(0).Favor(c) = mtypRow(0).Favor(c) + lngFavor
        Next
    Next
End Sub


' ************* SIZING *************


Private Sub SizeControl()
    Dim lngCharacters As Long
    Dim lngWidth As Long
    Dim i As Long
    
    If mlngCharacter = 0 Then lngCharacters = db.Characters Else lngCharacters = 1
    mlngMarginX = UserControl.ScaleX(cfg.MarginX, vbPixels, vbTwips)
    mlngMarginY = UserControl.ScaleY(cfg.MarginY, vbPixels, vbTwips)
    mlngPatronWidth = 0
    For i = 1 To db.Patrons
        lngWidth = UserControl.TextWidth(db.Patron(i).Patron)
        If mlngPatronWidth < lngWidth Then mlngPatronWidth = lngWidth
    Next
    mlngWidth = UserControl.TextWidth("Normal")
    For i = 1 To db.Characters
        lngWidth = UserControl.TextWidth(db.Character(i).Character)
        If mlngWidth < lngWidth Then mlngWidth = lngWidth
    Next
    mlngPatronWidth = mlngPatronWidth + mlngMarginX * 4
    mlngWidth = mlngWidth + mlngMarginX * 2
    mlngRowHeight = UserControl.TextHeight("Q") + mlngMarginY * 1.5
    mlngSides = UserControl.scrollVertical.Width ' UserControl.ScaleX(UserControl.ScaleY(UserControl.TextHeight("Q"), vbTwips, vbPixels), vbPixels, vbTwips) * 2
    mlngFitWidth = mlngPatronWidth + (lngCharacters + 1) * mlngWidth + mlngSides * 2
    mlngFitHeight = mlngRowHeight * (db.Patrons + 2)
    UserControl.picClient.Move 0, 0, mlngFitWidth, mlngFitHeight
    ShowScrollbars
    UserControl.Refresh
End Sub

Private Sub ShowScrollbars()
    Dim blnHorizontal As Boolean
    Dim blnVertical As Boolean
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    mblnOverride = True
    With UserControl
        .scrollHorizontal.Visible = False
        .scrollVertical.Visible = False
        If .picClient.Width > .picContainer.Width Then blnHorizontal = True
        If .picClient.Height > .picContainer.Height Then blnVertical = True
        If blnHorizontal And Not blnVertical Then
            If .picClient.Height + .scrollHorizontal.Height > .picContainer.Height Then blnVertical = True
        ElseIf blnVertical And Not blnHorizontal Then
            If .picClient.Width + .scrollVertical.Width > .picContainer.Height Then blnVertical = True
        End If
        lngWidth = .ScaleWidth
        If blnVertical Then lngWidth = lngWidth - .scrollVertical.Width
        lngHeight = .ScaleHeight
        If blnHorizontal Then lngHeight = lngHeight - .scrollHorizontal.Height
        .picContainer.Move 0, 0, lngWidth, lngHeight
        .picClient.Move 0, 0
        If blnVertical Then
            .scrollVertical.Move .ScaleWidth - .scrollVertical.Width, 0, .scrollVertical.Width, lngHeight
            .scrollVertical.Value = 0
            .scrollVertical.Max = .picClient.Height - .picContainer.Height
            .scrollVertical.SmallChange = Screen.TwipsPerPixelY
            .scrollVertical.LargeChange = .picContainer.Height
            .scrollVertical.Visible = True
        End If
        If blnHorizontal Then
            .scrollHorizontal.Move 0, .ScaleHeight - .scrollHorizontal.Height, lngWidth, .scrollHorizontal.Height
            .scrollHorizontal.Value = 0
            .scrollHorizontal.Max = .picClient.Width - .picContainer.Width
            .scrollHorizontal.SmallChange = Screen.TwipsPerPixelX
            .scrollHorizontal.LargeChange = .picContainer.Width
            .scrollHorizontal.Visible = True
        End If
    End With
    mblnOverride = False
End Sub


' ************* DRAWING *************


Private Sub DrawControl()
    Dim lngRow As Long
    
    UserControl.Cls
    DrawHeaders
    For lngRow = 1 To db.Patrons
        DrawRow lngRow
    Next
    DrawFooters
End Sub

Private Sub DrawHeaders()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngColor As Long
    Dim i As Long
    
    With UserControl.picClient
        .ForeColor = mlngBackText
        If mlngCharacter = 0 Then
            For i = 1 To db.Characters
                .CurrentX = mlngSides + mlngPatronWidth + (mlngWidth * (i - 1)) + (mlngWidth - .TextWidth(db.Character(i).Character)) \ 2
                .CurrentY = mlngMarginY
                UserControl.picClient.Print db.Character(i).Character
            Next
            lngWidth = mlngWidth * db.Characters
        Else
            i = mlngCharacter
            .CurrentX = mlngSides + mlngPatronWidth + (mlngWidth - .TextWidth(db.Character(i).Character)) \ 2
            .CurrentY = mlngMarginY
            UserControl.picClient.Print db.Character(i).Character
            lngWidth = mlngWidth
        End If
        .CurrentX = mlngSides + mlngPatronWidth + lngWidth + (mlngWidth - .TextWidth("Max")) \ 2
        .CurrentY = mlngMarginY
        UserControl.picClient.Print "Max"
    End With
End Sub

Private Sub DrawRow(plngRow As Long)
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim i As Long
    
    DrawPatron plngRow, False
    UserControl.picClient.ForeColor = mlngGridText
    lngTop = plngRow * mlngRowHeight
    If mlngCharacter = 0 Then
        For i = 1 To db.Characters
            DrawCell mlngSides + mlngPatronWidth + mlngWidth * (i - 1), lngTop, mtypRow(plngRow).Favor(i), db.Character(i).BackColor
        Next
        lngWidth = mlngWidth * db.Characters
    Else
        i = mlngCharacter
        DrawCell mlngSides + mlngPatronWidth, lngTop, mtypRow(plngRow).Favor(i), db.Character(i).BackColor
        lngWidth = mlngWidth
    End If
    DrawCell mlngSides + mlngPatronWidth + lngWidth, lngTop, mtypRow(plngRow).Total, mlngGridBack
End Sub

Private Sub DrawPatron(plngRow As Long, pblnActive As Boolean)
    Dim lngTop As Long
    
    lngTop = plngRow * mlngRowHeight
    UserControl.picClient.Line (mlngMarginX, lngTop)-(mtypRow(plngRow).PatronRight, lngTop + mlngRowHeight), mlngBackground, BF
    If pblnActive Then UserControl.picClient.ForeColor = cfg.GetColor(cgeControls, cveTextLink) Else UserControl.picClient.ForeColor = mlngBackText
    mtypRow(plngRow).PatronLeft = mlngSides
    UserControl.picClient.CurrentX = mtypRow(plngRow).PatronLeft
    UserControl.picClient.CurrentY = lngTop + mlngMarginY
    UserControl.picClient.Print mtypRow(plngRow).Patron;
    mtypRow(plngRow).PatronRight = UserControl.picClient.CurrentX
End Sub

Private Sub DrawCell(plngLeft As Long, plngTop As Long, ByVal pstrDisplay As String, plngColor As Long, Optional pblnBorder As Boolean = True)
    If pblnBorder Then
        UserControl.picClient.FillColor = plngColor
        UserControl.picClient.Line (plngLeft, plngTop)-(plngLeft + mlngWidth, plngTop + mlngRowHeight), cfg.GetColor(cgeControls, cveBorderInterior), B
    Else
        UserControl.picClient.Line (plngLeft, plngTop + PixelY)-(plngLeft + mlngWidth, plngTop + mlngRowHeight), plngColor, BF
    End If
    UserControl.picClient.CurrentX = plngLeft + (mlngWidth - UserControl.picClient.TextWidth(pstrDisplay)) \ 2
    UserControl.picClient.CurrentY = plngTop + mlngMarginY
    UserControl.picClient.Print pstrDisplay
End Sub

Private Sub DrawFooters()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim i As Long
    
    UserControl.picClient.ForeColor = mlngBackText
    lngTop = (db.Patrons + 1) * mlngRowHeight
    If mlngCharacter = 0 Then
        For i = 1 To db.Characters
            DrawCell mlngSides + mlngPatronWidth + mlngWidth * (i - 1), lngTop, mtypRow(0).Favor(i), mlngBackground, False
        Next
        lngWidth = mlngWidth * db.Characters
    Else
        i = mlngCharacter
        DrawCell mlngSides + mlngPatronWidth, lngTop, mtypRow(0).Favor(i), mlngBackground, False
        lngWidth = mlngWidth
    End If
    DrawCell mlngSides + mlngPatronWidth + lngWidth, lngTop, mtypRow(0).Total, mlngBackground, False
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
    If mlngActiveRow Then PatronWiki mtypRow(mlngActiveRow).Patron
End Sub

Private Sub ActiveRow(X As Single, Y As Single)
    Dim lngRow As Long
    
    lngRow = Y \ mlngRowHeight
    If lngRow < 1 Or lngRow > db.Patrons Then
        lngRow = 0
    ElseIf X < mtypRow(lngRow).PatronLeft Or X > mtypRow(lngRow).PatronRight Then
        lngRow = 0
    End If
    If lngRow <> 0 Then xp.SetMouseCursor mcHand
    If lngRow <> mlngActiveRow Then
        If mlngActiveRow <> 0 Then DrawPatron mlngActiveRow, False
        mlngActiveRow = lngRow
        If mlngActiveRow <> 0 Then DrawPatron mlngActiveRow, True
    End If
End Sub


' ************* SCROLLBARS *************



Public Sub WheelScroll(plngValue As Long)
    ' Hmmm
End Sub

Private Sub scrollHorizontal_Change()
    HorizontalScroll
End Sub

Private Sub scrollHorizontal_Scroll()
    HorizontalScroll
End Sub

Private Sub HorizontalScroll()
    With UserControl
        .picClient.Left = 0 - .scrollHorizontal.Value
    End With
End Sub

Private Sub scrollVertical_Change()
    VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    With UserControl
        .picClient.Top = 0 - .scrollVertical.Value
    End With
End Sub

Private Sub scrollHorizontal_GotFocus()
    NoFocus
End Sub

Private Sub scrollVertical_GotFocus()
    NoFocus
End Sub

Private Sub NoFocus()
    On Error Resume Next
    UserControl.picContainer.SetFocus
End Sub

