VERSION 5.00
Begin VB.Form frmScaling 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Scaling"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12165
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
   ForeColor       =   &H00404040&
   Icon            =   "frmScaling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   12165
   Begin CannithCrafting.userCheckBox usrchkDarker 
      Height          =   252
      Left            =   1620
      TabIndex        =   7
      Top             =   0
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   450
      Value           =   0   'False
      Caption         =   "Gamma"
   End
   Begin VB.PictureBox picGroup 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1092
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   1260
      Width           =   1032
   End
   Begin VB.PictureBox picML 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   2040
      ScaleHeight     =   435
      ScaleWidth      =   1695
      TabIndex        =   2
      Top             =   240
      Width           =   1692
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2352
      Left            =   1740
      ScaleHeight     =   2355
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   1020
      Width           =   2892
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1152
         Left            =   0
         ScaleHeight     =   1155
         ScaleWidth      =   1695
         TabIndex        =   5
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   1260
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   1632
      Left            =   5820
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   264
      TabIndex        =   6
      Top             =   36
      Width           =   432
   End
   Begin VB.Menu mnuContext 
      Caption         =   "Context"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuShard 
         Caption         =   "ShardName"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmScaling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SimulateLargeDisplay As Boolean = False
Private mlngRows As Long ' Debugging tools to simulate when the screen
Private mlngCols As Long ' is large enough that we don't need scrollbars

Private Enum ScaleColorEnum
    reYellow
    reRed
    rePurple
    reBlue
    reTeal
    reGreen
    reWhite
    reGray
End Enum

Private mlngColor(1, 7) As Long

Private Type ScaleType
    GroupName As String
    ScaleName As String
    Color As ScaleColorEnum
    Table() As String
    Selected As Boolean
End Type

Private mtypScale() As ScaleType
Private mblnColSelected() As Boolean

Private mlngTextHeight As Long
Private mlngRowHeight As Long
Private mlngTextOffset As Long
Private mlngCellWidth As Long
Private mlngSpace As Long
Private mlngMinGroupWidth As Long
Private mlngGroupWidth As Long
Private mlngMinScaleWidth As Long
Private mlngScaleWidth As Long

Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
Private mlngRight As Long
Private mlngBottom As Long

Private mlngShowRows As Long
Private mlngShowCols As Long
' Active selections
Private mlngRow As Long
Private mlngCol As Long
Private mlngScale As Long
Private mlngML As Long


Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    InitColors
    InitScales
    InitTextSize
    InitFormSize
    DrawGrid
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    If Rotation < 0 Then
        KeyScroll Me.scrollVertical, 3
    Else
        KeyScroll Me.scrollVertical, -3
    End If
End Sub


' ************* INITIALIZE *************


Private Sub InitColors()
    Dim ColorOn As Long
    Dim ColorOff As Long
    Dim SelectOn As Long
    Dim SelectOff As Long
    Dim ColorGray As Long
    
    GridColors ColorOn, ColorOff, SelectOn, SelectOff, ColorGray
    ' Standard
    mlngColor(0, reYellow) = RGB(ColorOn, ColorOn, ColorOff)
    mlngColor(0, reRed) = RGB(ColorOn, ColorOff, ColorOff)
    mlngColor(0, rePurple) = RGB(ColorOn, ColorOff, ColorOn)
    mlngColor(0, reBlue) = RGB(ColorOff, ColorOff, ColorOn)
    mlngColor(0, reTeal) = RGB(ColorOff, ColorOn, ColorOn)
    mlngColor(0, reGreen) = RGB(ColorOff, ColorOn, ColorOff)
    mlngColor(0, reWhite) = RGB(ColorOn, ColorOn, ColorOn)
    mlngColor(0, reGray) = RGB(ColorGray, ColorGray, ColorGray)
    ' Selected
    mlngColor(1, reYellow) = RGB(SelectOn, SelectOn, SelectOff)
    mlngColor(1, reRed) = RGB(SelectOn, SelectOff, SelectOff)
    mlngColor(1, rePurple) = RGB(SelectOn, SelectOff, SelectOn)
    mlngColor(1, reBlue) = RGB(SelectOff, SelectOff, SelectOn)
    mlngColor(1, reTeal) = RGB(SelectOff, SelectOn, SelectOn)
    mlngColor(1, reGreen) = RGB(SelectOff, SelectOn, SelectOff)
    mlngColor(1, reWhite) = RGB(SelectOn, SelectOn, SelectOn)
    mlngColor(1, reGray) = RGB(ColorGray, ColorGray, ColorGray)
End Sub

' Transpose scale order back to natural order of data file
Private Sub InitScales()
    Dim i As Long
    
    ReDim mtypScale(1 To db.Scales)
    For i = 1 To db.Scales
        With mtypScale(db.Scaling(i).Order)
            .GroupName = db.Scaling(i).Group
            .ScaleName = db.Scaling(i).ScaleName
            .Table = db.Scaling(i).Table
            .Color = GetFillColor(.GroupName)
        End With
    Next
End Sub

Private Function GetFillColor(pstrGroup As String) As ScaleColorEnum
    Dim enColor As ScaleColorEnum
    
    Select Case pstrGroup
        Case "General": enColor = reYellow
        Case "Offense": enColor = reRed
        Case "Weapon": enColor = rePurple
        Case "Tactics": enColor = reBlue
        Case "Tanking": enColor = reTeal
        Case "Stealth": enColor = reGreen
        Case "Defense": enColor = reYellow
        Case "Saves": enColor = reRed
        Case "Spellcasting": enColor = reBlue
        Case "Non-scaling": enColor = reWhite
    End Select
    GetFillColor = enColor
End Function

Private Sub InitTextSize()
    Dim i As Long
    
    mlngTextHeight = Me.TextHeight("Q")
    mlngRowHeight = ((mlngTextHeight * 1.25) \ PixelY) * PixelY
    mlngTextOffset = (mlngRowHeight - mlngTextHeight) \ 2
    mlngCellWidth = Me.TextWidth(" 1.5[W] ")
    mlngSpace = Me.TextWidth(" ")
    ' Scale size (bolded when active)
    Me.FontBold = True
    For i = 1 To db.Scales
        With mtypScale(i)
            If mlngMinScaleWidth < Me.TextWidth(.ScaleName) Then mlngMinScaleWidth = Me.TextWidth(.ScaleName)
        End With
    Next
    mlngMinScaleWidth = mlngMinScaleWidth + mlngSpace * 2
    ' Group size (never bolded)
    Me.FontBold = False
    For i = 1 To db.Scales
        With mtypScale(i)
            If mlngMinGroupWidth < Me.TextWidth(.GroupName) Then mlngMinGroupWidth = Me.TextWidth(.GroupName)
        End With
    Next
    mlngMinGroupWidth = mlngMinGroupWidth + mlngSpace * 2
End Sub

Private Sub InitFormSize()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngMaxWidth As Long
    Dim lngMaxHeight As Long
    Dim blnHorizontal As Boolean
    Dim blnVertical As Boolean
    
    If SimulateLargeDisplay Then
        mlngRows = 20
        mlngCols = 10
    Else
        mlngRows = db.Scales
        mlngCols = MAX_CRAFTING_LEVEL
    End If
    ReDim mblnColSelected(1 To mlngCols)
    xp.GetDesktop lngLeft, lngTop, lngWidth, lngHeight
    lngMaxWidth = mlngMinGroupWidth + mlngMinScaleWidth + (mlngCols * mlngCellWidth) + Me.Width - Me.ScaleWidth
    lngMaxHeight = (mlngRows + 1) * mlngRowHeight + Me.Height - Me.ScaleHeight
    If lngMaxWidth > lngWidth Then
        blnHorizontal = True
        If lngMaxHeight + Me.scrollHorizontal.Height > lngHeight Then blnVertical = True
    ElseIf lngMaxHeight > lngHeight Then
        blnVertical = True
        If lngMaxWidth + Me.scrollVertical.Width > lngWidth Then blnHorizontal = True
    End If
    If blnHorizontal And blnVertical Then
        Me.WindowState = vbMaximized
    ElseIf blnHorizontal Then
        Me.Move lngLeft, lngTop + (lngHeight - lngMaxHeight) \ 2, lngWidth, lngMaxHeight + Me.scrollHorizontal.Height
    ElseIf blnVertical Then
        Me.Move lngLeft + (lngWidth - lngMaxWidth) \ 2, lngTop, lngMaxWidth + Me.scrollVertical.Width, lngHeight
    Else
        Me.Move lngLeft + (lngWidth - lngMaxWidth) \ 2, lngTop + (lngHeight - lngMaxHeight) \ 2, lngMaxWidth, lngMaxHeight
    End If
End Sub


' ************* RESIZE *************


Private Sub Form_Resize()
    Dim blnHorizontal As Boolean
    Dim blnVertical As Boolean
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleWidth < mlngMinGroupWidth + mlngMinScaleWidth + mlngCellWidth + Me.scrollVertical.Width Then Exit Sub
    If Me.ScaleHeight < mlngRowHeight * 2 + Me.scrollHorizontal.Height Then Exit Sub
    IdentifyScrollbars blnHorizontal, blnVertical
    GetDimensions blnHorizontal, blnVertical
    Me.picML.Move mlngLeft, 0, mlngWidth, mlngTop + PixelY
    Me.picGroup.Move 0, mlngTop, mlngLeft + PixelX, mlngHeight '+ PixelY
    Me.picContainer.Move mlngLeft, mlngTop, mlngWidth, mlngHeight
    SetScrollbars blnHorizontal, blnVertical
    DrawColumnHeaders
    DrawRowHeaders
End Sub

Private Sub IdentifyScrollbars(pblnHorizontal As Boolean, pblnVertical As Boolean)
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    lngWidth = mlngMinGroupWidth + mlngMinScaleWidth + (mlngCols * mlngCellWidth)
    lngHeight = (mlngRows + 1) * mlngRowHeight
    If lngWidth > Me.ScaleWidth Then
        pblnHorizontal = True
        If lngHeight + Me.scrollHorizontal.Height > Me.ScaleHeight Then pblnVertical = True
    ElseIf lngHeight > Me.ScaleHeight Then
        pblnVertical = True
        If lngWidth + Me.scrollVertical.Width > Me.ScaleWidth Then pblnHorizontal = True
    End If
End Sub

Private Sub GetDimensions(pblnHorizontal As Boolean, pblnVertical As Boolean)
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    lngHeight = Me.ScaleHeight
    If pblnHorizontal Then lngHeight = lngHeight - Me.scrollHorizontal.Height
    mlngShowRows = (lngHeight - mlngRowHeight) \ mlngRowHeight
    If mlngShowRows >= mlngRows Then
        mlngShowRows = mlngRows
        mlngTop = mlngRowHeight
    Else
        mlngTop = lngHeight - (mlngShowRows * mlngRowHeight)
    End If
    lngWidth = Me.ScaleWidth
    If pblnVertical Then lngWidth = lngWidth - Me.scrollVertical.Width
    mlngShowCols = (lngWidth - mlngMinGroupWidth - mlngMinScaleWidth) \ mlngCellWidth
    If mlngShowCols >= mlngCols Then
        mlngShowCols = mlngCols
        mlngGroupWidth = mlngMinGroupWidth
        mlngScaleWidth = mlngMinScaleWidth
        mlngLeft = mlngGroupWidth + mlngScaleWidth
    Else
        mlngLeft = lngWidth - (mlngShowCols * mlngCellWidth)
        mlngGroupWidth = mlngMinGroupWidth + (mlngLeft - mlngMinGroupWidth - mlngMinScaleWidth) \ 2
        mlngScaleWidth = mlngLeft - mlngGroupWidth
    End If
    mlngWidth = mlngShowCols * mlngCellWidth + PixelX
    mlngHeight = mlngShowRows * mlngRowHeight + PixelY
    mlngRight = mlngLeft + mlngWidth
    mlngBottom = mlngTop + mlngHeight
End Sub

Private Sub SetScrollbars(pblnHorizontal As Boolean, pblnVertical As Boolean)
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngMax As Long
    
    mblnOverride = True
    With Me.scrollHorizontal
        lngWidth = Me.ScaleWidth
        If pblnVertical Then lngWidth = lngWidth - Me.scrollVertical.Width
        .Move 0, Me.ScaleHeight - .Height, lngWidth
        lngMax = mlngCols - mlngShowCols
        If .Value > lngMax Then
            .Value = lngMax
            Me.picClient.Left = Me.scrollHorizontal.Value * mlngCellWidth * -1
        End If
        .Max = lngMax
        .LargeChange = mlngShowCols
        .Visible = pblnHorizontal
    End With
    With Me.scrollVertical
        lngHeight = Me.ScaleHeight
        If pblnHorizontal Then lngHeight = lngHeight - Me.scrollHorizontal.Height
        .Move Me.ScaleWidth - .Width, 0, .Width, lngHeight
        lngMax = mlngRows - mlngShowRows
        If .Value > lngMax Then
            .Value = lngMax
            Me.picClient.Top = Me.scrollVertical.Value * mlngRowHeight * -1
        End If
        .Max = lngMax
        .LargeChange = mlngShowRows
        .Visible = pblnVertical
    End With
    mblnOverride = False
End Sub


' ************* GRID *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    Me.picClient.Cls
    Me.picClient.FontBold = False
    Me.picClient.Move 0, 0, mlngCellWidth * mlngCols + PixelX, mlngRowHeight * mlngRows + PixelY
    For lngRow = 1 To mlngRows
        For lngCol = 1 To mlngCols
            DrawCell lngRow, lngCol, False
        Next
    Next
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Dim strText As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngBorderColor As Long
    Dim blnSelected As Boolean
    
    If plngRow = 0 Or plngRow > mlngRows Or plngCol = 0 Or plngCol > mlngCols Then Exit Sub
    lngLeft = (plngCol - 1) * mlngCellWidth
    lngTop = (plngRow - 1) * mlngRowHeight
    lngRight = lngLeft + mlngCellWidth
    lngBottom = lngTop + mlngRowHeight
    If pblnActive Then lngBorderColor = vbBlack Else lngBorderColor = GetColor(reGray)
    With mtypScale(plngRow)
        strText = .Table(plngCol)
        blnSelected = .Selected Or mblnColSelected(plngCol)
        Me.picClient.FillColor = GetColor(.Color, blnSelected)
        Me.picClient.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngBorderColor, B
    End With
    If Not pblnActive Then
        If plngRow > 1 Then
            If mtypScale(plngRow).GroupName <> mtypScale(plngRow - 1).GroupName Then Me.picClient.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), vbBlack
        End If
        If plngRow < mlngRows Then
            If mtypScale(plngRow).GroupName <> mtypScale(plngRow + 1).GroupName Then Me.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), vbBlack
        End If
        If plngRow = mlngRows Then Me.picClient.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), vbBlack
        If plngCol Mod 5 = 1 Then
            Me.picClient.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), vbBlack
        ElseIf plngCol = mlngCols Or plngCol Mod 5 = 0 Then
            Me.picClient.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), vbBlack
        End If
    End If
    Me.picClient.CurrentX = lngLeft + (mlngCellWidth - Me.picClient.TextWidth(strText)) \ 2
    Me.picClient.CurrentY = lngTop + mlngTextOffset
    Me.picClient.Print strText
End Sub

Private Sub picClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    ClearActiveML
    ClearActiveScale
    lngRow = (Y \ mlngRowHeight) + 1
    lngCol = (X \ mlngCellWidth) + 1
    If lngRow <> mlngRow Or lngCol <> mlngCol Then
        ActiveCell False
        mlngRow = lngRow
        mlngCol = lngCol
        ActiveCell True
    End If
End Sub

Private Sub picClient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearSelections
End Sub

Private Sub ActiveCell(pblnActive As Boolean)
    If mlngRow = 0 Or mlngCol = 0 Then Exit Sub
    DrawCell mlngRow, mlngCol, pblnActive
    DrawColumnHeader mlngCol - Me.scrollHorizontal.Value, pblnActive, False
    DrawScaleHeader mlngRow - Me.scrollVertical.Value, pblnActive, False
End Sub

Private Sub ClearActiveCell()
    ActiveCell False
    mlngRow = 0
    mlngCol = 0
End Sub

Private Sub ClearSelections()
    Dim i As Long
    
    For i = 1 To mlngRows
        If mtypScale(i).Selected Then SelectScale i, False
    Next
    For i = 1 To mlngCols
        If mblnColSelected(i) Then SelectML i, False
    Next
End Sub


' ************* COLUMN HEADERS (ML) *************


Private Sub DrawColumnHeaders()
    Dim i As Long
    
    For i = 1 To mlngShowCols
        DrawColumnHeader i, False, False
    Next
End Sub

Private Sub DrawColumnHeader(plngCol As Long, pblnActive As Boolean, pblnDarkBorder As Boolean)
    Dim lngCol As Long
    Dim strText As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngColor As Long
    
    lngCol = plngCol + Me.scrollHorizontal.Value
    If lngCol = 0 Or lngCol > mlngCols Then Exit Sub
    If pblnActive Then Me.picML.FontBold = True
    If lngCol > 30 Then strText = "PL" & lngCol Else strText = "ML" & lngCol
    lngLeft = (plngCol - 1) * mlngCellWidth
    lngRight = lngLeft + mlngCellWidth
    If pblnDarkBorder Then lngColor = vbBlack Else lngColor = GetColor(reGray)
    If mblnColSelected(lngCol) Then Me.picML.FillColor = GetColor(reWhite, True)
    Me.picML.Line (lngLeft, 0)-(lngLeft + mlngCellWidth, mlngTop), lngColor, B
    If Not pblnDarkBorder Then
        If plngCol = 1 Or lngCol Mod 5 = 1 Then Me.picML.Line (lngLeft, 0)-(lngLeft, mlngTop), vbBlack
        If lngCol = mlngCols Or lngCol Mod 5 = 0 Then Me.picML.Line (lngRight, 0)-(lngRight, mlngTop), vbBlack
        Me.picML.Line (lngLeft, mlngTop)-(lngRight + PixelX, mlngTop), vbBlack
    End If
    Me.picML.CurrentX = lngLeft + (mlngCellWidth - Me.picML.TextWidth(strText)) \ 2
    Me.picML.CurrentY = (mlngTop - mlngTextHeight) \ 2
    Me.picML.Print strText
    If pblnActive Then Me.picML.FontBold = False
    If mblnColSelected(lngCol) Then Me.picML.FillColor = GetColor(reWhite, False)
End Sub

Private Sub picML_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long

    ClearActiveCell
    ClearActiveScale
    lngCol = (X \ mlngCellWidth) + 1 + Me.scrollHorizontal.Value
    If lngCol = mlngML Then Exit Sub
    DrawColumnHeader mlngML, False, False
    mlngML = lngCol
    DrawColumnHeader mlngML, False, True
End Sub

Private Sub picML_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    Dim blnExclusive As Boolean
    Dim i As Long

    lngCol = (X \ mlngCellWidth) + 1 + Me.scrollHorizontal.Value
    If lngCol < 1 Or lngCol > mlngCols Then Exit Sub
    blnExclusive = (Button = vbLeftButton And Shift = 0)
    If blnExclusive Then
        For i = 1 To mlngCols
            SelectML i, (i = lngCol)
        Next
    Else
        SelectML lngCol, True
    End If
End Sub

Private Sub SelectML(plngCol As Long, pblnSelected As Boolean)
    Dim lngCol As Long
    Dim i As Long

    If mblnColSelected(plngCol) = pblnSelected Then Exit Sub
    mblnColSelected(plngCol) = pblnSelected
    lngCol = plngCol - Me.scrollHorizontal.Value
    DrawColumnHeader plngCol - Me.scrollHorizontal.Value, False, False
    For i = 1 To mlngRows
        DrawCell i, plngCol, False
    Next
End Sub

Private Sub ClearActiveML()
    DrawColumnHeader mlngML, False, False
    mlngML = 0
End Sub


' ************* ROW HEADERS (SCALE) *************


Private Sub DrawRowHeaders()
    Dim strGroup As String
    Dim lngRow As Long
    Dim lngGroupTop As Long
    Dim lngFirst As Long
    Dim lngTop As Long
    Dim lngBottom As Long
    
    Me.picGroup.FontBold = False
    lngFirst = Me.scrollVertical.Value
    With mtypScale(lngFirst + 1)
        strGroup = .GroupName
        Me.picGroup.FillColor = GetColor(.Color)
    End With
    For lngRow = 1 To mlngShowRows
        With mtypScale(lngFirst + lngRow)
            If .GroupName <> strGroup Then
                DrawGroupHeader strGroup, lngGroupTop, lngTop
                strGroup = .GroupName
                lngGroupTop = lngTop
                Me.picGroup.FillColor = GetColor(.Color)
            End If
            If .Selected Then Me.picGroup.FillColor = GetColor(.Color, True)
            Me.picGroup.Line (mlngGroupWidth, lngTop)-(mlngLeft, lngTop + mlngRowHeight), GetColor(reGray), B
            Me.picGroup.CurrentX = mlngGroupWidth + mlngSpace
            Me.picGroup.CurrentY = lngTop + mlngTextOffset
            Me.picGroup.Print .ScaleName
            If .Selected Then Me.picGroup.FillColor = GetColor(.Color, False)
            lngTop = lngTop + mlngRowHeight
        End With
    Next
    DrawGroupHeader strGroup, lngGroupTop, lngTop
    lngBottom = mlngBottom - mlngTop - PixelY
    Me.picGroup.Line (mlngLeft, 0)-(mlngLeft, lngBottom), vbBlack
    If lngFirst + mlngShowRows = mlngRows Then
        Me.picGroup.Line (0, lngBottom)-(mlngLeft, lngBottom), vbBlack
    ElseIf lngFirst + lngRow < mlngRows Then
        If mtypScale(lngFirst + lngRow).GroupName <> strGroup Then Me.picGroup.Line (0, lngBottom)-(mlngLeft, lngBottom), vbBlack
    End If
End Sub

Private Sub DrawGroupHeader(pstrGroup As String, plngTop As Long, plngBottom As Long)
    Me.picGroup.Line (0, plngTop)-(mlngGroupWidth, plngBottom), GetColor(reGray), B
    Me.picGroup.Line (0, plngTop)-(mlngRight, plngTop), vbBlack
    Me.picGroup.CurrentX = (mlngGroupWidth - Me.TextWidth(pstrGroup)) \ 2
    Me.picGroup.CurrentY = plngTop + (plngBottom - plngTop - mlngTextHeight) \ 2
    Me.picGroup.Print pstrGroup
End Sub

Private Sub picGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GroupMouseMove X, Y
End Sub

Private Sub GroupMouseMove(X As Single, Y As Single)
    Dim lngRow As Long
    
    ClearActiveCell
    ClearActiveML
    If X < mlngGroupWidth Then
        ClearActiveScale
    Else
        lngRow = (Y \ mlngRowHeight) + 1
        If lngRow = mlngScale Then Exit Sub
        DrawScaleHeader mlngScale, False, False
        mlngScale = lngRow
        DrawScaleHeader mlngScale, False, True
    End If
End Sub

Private Sub picGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngScale As Long
    Dim blnExclusive As Boolean
    Dim i As Long
    
    GroupMouseMove X, Y
    If X < mlngGroupWidth Then Exit Sub
    lngScale = (Y \ mlngRowHeight) + Me.scrollVertical.Value + 1
    If lngScale < 1 Or lngScale > mlngRows Then Exit Sub
    If Button = vbRightButton Then
        PopupScales lngScale
    Else
        blnExclusive = (Button = vbLeftButton And Shift = 0)
        If blnExclusive Then
            For i = 1 To mlngRows
                SelectScale i, (i = lngScale)
            Next
        Else
            SelectScale lngScale, True
        End If
    End If
End Sub

Private Sub PopupScales(plngScale As Long)
    Dim strShard() As String
    Dim lngShard As Long
    Dim i As Long
    
    ReDim strShard(db.Shards)
    For i = 1 To db.Shards
        If db.Shard(i).ScaleName = mtypScale(plngScale).ScaleName Then
            strShard(lngShard) = db.Shard(i).ShardName
            lngShard = lngShard + 1
        End If
    Next
    If lngShard = 0 Then Exit Sub
    For i = 0 To lngShard - 1
        If i > Me.mnuShard.UBound Then Load Me.mnuShard(i)
        Me.mnuShard(i).Caption = strShard(i)
        Me.mnuShard(i).Visible = True
    Next
    For i = Me.mnuShard.UBound To lngShard Step -1
        Unload Me.mnuShard(i)
    Next
    PopupMenu Me.mnuContext(0)
End Sub

Private Sub mnuShard_Click(Index As Integer)
    OpenShard Me.mnuShard(Index).Caption
End Sub

Private Sub SelectScale(plngScale As Long, pblnSelected As Boolean)
    Dim lngRow As Long
    Dim i As Long
    
    If mtypScale(plngScale).Selected = pblnSelected Then Exit Sub
    mtypScale(plngScale).Selected = pblnSelected
    lngRow = plngScale - Me.scrollVertical.Value
    DrawScaleHeader lngRow, False, False
    For i = 1 To mlngCols
        DrawCell plngScale, i, False
    Next
End Sub

Private Sub DrawScaleHeader(plngRow As Long, pblnBold As Boolean, pblnDarkBorder As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngScale As Long
    Dim lngColor As Long
    
    If plngRow = 0 Or plngRow > mlngRows Then Exit Sub
    lngScale = plngRow + Me.scrollVertical.Value
    If lngScale < 1 Or lngScale > mlngRows Then Exit Sub
    lngLeft = mlngGroupWidth
    lngTop = (plngRow - 1) * mlngRowHeight
    lngRight = lngLeft + mlngScaleWidth
    lngBottom = lngTop + mlngRowHeight
    With mtypScale(lngScale)
        If pblnBold Then Me.picGroup.FontBold = True
        Me.picGroup.FillColor = GetColor(.Color, .Selected)
        If pblnDarkBorder Then lngColor = vbBlack Else lngColor = GetColor(reGray)
        Me.picGroup.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngColor, B
        Me.picGroup.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), vbBlack
        If plngRow = 1 Then Me.picGroup.Line (lngLeft, lngTop)-(lngRight, lngTop), vbBlack
        If lngScale > 1 Then
            If .GroupName <> mtypScale(lngScale - 1).GroupName Then Me.picGroup.Line (lngLeft, lngTop)-(lngRight, lngTop), vbBlack
        End If
        If lngScale < mlngRows Then
            If .GroupName <> mtypScale(lngScale + 1).GroupName Then Me.picGroup.Line (lngLeft, lngBottom)-(lngRight, lngBottom), vbBlack
        End If
        If plngRow = mlngShowRows Then Me.picGroup.Line (lngLeft, lngBottom)-(lngRight, lngBottom), vbBlack
        Me.picGroup.CurrentX = lngLeft + mlngSpace
        Me.picGroup.CurrentY = lngTop + mlngTextOffset
        Me.picGroup.Print .ScaleName
        If pblnBold Then Me.picGroup.FontBold = False
    End With
End Sub

Private Sub ClearActiveScale()
    DrawScaleHeader mlngScale, False, False
    mlngScale = 0
End Sub


' ************* SCROLLING *************


Private Sub scrollHorizontal_GotFocus()
    Me.picContainer.SetFocus
    ClearActive
End Sub

Private Sub scrollHorizontal_Scroll()
    If Not mblnOverride Then HorizontalScroll
End Sub

Private Sub scrollHorizontal_Change()
    If Not mblnOverride Then HorizontalScroll
End Sub

Private Sub scrollVertical_GotFocus()
    Me.picContainer.SetFocus
    ClearActive
End Sub

Private Sub scrollVertical_Scroll()
    If Not mblnOverride Then VerticalScroll
End Sub

Private Sub scrollVertical_Change()
    If Not mblnOverride Then VerticalScroll
End Sub

Private Sub VerticalScroll()
    DrawRowHeaders
    Me.picClient.Top = Me.scrollVertical.Value * mlngRowHeight * -1
End Sub

Private Sub HorizontalScroll()
    DrawColumnHeaders
    Me.picClient.Left = Me.scrollHorizontal.Value * mlngCellWidth * -1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown: KeyScroll Me.scrollVertical, 1
        Case vbKeyUp: KeyScroll Me.scrollVertical, -1
        Case vbKeyLeft: KeyScroll Me.scrollHorizontal, -1
        Case vbKeyRight: KeyScroll Me.scrollHorizontal, 1
        Case vbKeyPageUp: KeyScroll Me.scrollVertical, -2
        Case vbKeyPageDown: KeyScroll Me.scrollVertical, 2
        Case vbKeyHome: KeyScroll Me.scrollHorizontal, 0
        Case vbKeyEnd: KeyScroll Me.scrollHorizontal, 99
        Case vbKeyEscape: Unload Me
        Case vbKeyReturn: ClearActive
    End Select
End Sub

Private Sub KeyScroll(pctl As Control, plngIncrement As Long)
    Dim lngValue As Long
    
    If Not pctl.Visible Then Exit Sub
    Select Case plngIncrement
        Case -3, -1, 1, 3: lngValue = pctl.Value + plngIncrement
        Case -2, 2: lngValue = pctl.Value + (plngIncrement \ 2) * pctl.LargeChange
        Case 0: lngValue = 0
        Case 99: lngValue = pctl.Max
    End Select
    If lngValue < 0 Then lngValue = 0
    If lngValue > pctl.Max Then lngValue = pctl.Max
    If pctl.Value <> lngValue Then pctl.Value = lngValue
End Sub


' ************* GENERAL *************


Private Function GetColor(penScaleColor As ScaleColorEnum, Optional pblnSelected As Boolean = False) As Long
    Dim lngIndex As Long
    
    If pblnSelected Then lngIndex = 1
    GetColor = mlngColor(lngIndex, penScaleColor)
End Function

Private Sub ClearActive()
    ClearActiveCell
    ClearActiveScale
    ClearActiveML
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearActive
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lblHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    ShowHelp "Scaling"
End Sub

Private Sub usrchkDarker_UserChange()
    cfg.DarkColors = Me.usrchkDarker.Value
    InitColors
    DrawRowHeaders
    DrawColumnHeaders
    DrawGrid
End Sub

