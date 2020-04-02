VERSION 5.00
Begin VB.UserControl userList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3792
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4032
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3792
   ScaleWidth      =   4032
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3252
      Left            =   0
      ScaleHeight     =   3252
      ScaleWidth      =   3732
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3732
      Begin VB.VScrollBar scrollVertical 
         Height          =   3252
         Left            =   3480
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   2412
         Left            =   0
         ScaleHeight     =   2412
         ScaleWidth      =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   3480
         Begin CharacterBuilderLite.userSlot usrSlot 
            Height          =   288
            Index           =   0
            Left            =   720
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2172
            _ExtentX        =   3831
            _ExtentY        =   508
         End
      End
   End
   Begin VB.Shape shpBorder 
      Height          =   252
      Left            =   0
      Top             =   0
      Width           =   2772
   End
End
Attribute VB_Name = "userList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Public Event RankChange(Index As Integer, Ranks As Long)
Public Event OLECompleteDrag(Index As Integer, Effect As Long)
Public Event OLEDragDrop(Index As Integer, Data As DataObject)
Public Event OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
Public Event SlotClick(Index As Integer, Button As Integer)
Public Event SlotDblClick(Index As Integer)

Public Event RequestDrag(Index As Integer, Allow As Boolean)

Private Type ColumnType
    Left As Long
    Width As Long
    Right As Long
    Align As AlignmentConstants
    Header As String
    SizeText As String
    Slot As Boolean
End Type

Private mtypColumn() As ColumnType
Private mlngColumns As Long
Private mlngSlotColumn As Long
Private mlngRows As Long

Private mlngTop As Long ' Top of picContainer
Private mlngHeight As Long ' Height of slot control
Private mlngMarginY As Long ' Space between slots
Private mlngMarginX As Long ' Space between columns
Private mlngOffsetY As Long ' Where to print text in relation to slot so captions line up

Private mlngTextColor As Long
Private mlngBackColor As Long
Private mlngBorderColor As Long
Private mlngClientBack As Long

Private grid() As String
Private mlngActive As Long

Private mlngPixelX As Long
Private mlngPixelY As Long

Private mblnNoFocus As Boolean


' ************* INITIALIZE *************


Private Sub UserControl_Initialize()
    Dim lngTextHeight As Long
    
    mlngPixelX = Screen.TwipsPerPixelX
    mlngPixelY = Screen.TwipsPerPixelY
    lngTextHeight = UserControl.TextHeight("X")
    mlngHeight = lngTextHeight * 4 / 3
    mlngOffsetY = (mlngHeight - lngTextHeight) \ 2
    mlngMarginY = mlngHeight / 4
    mlngMarginX = UserControl.TextWidth("  ")
    mlngTop = 30 * mlngPixelY
    mlngTextColor = vbWindowText
    mlngBackColor = vbWindowBackground
    mlngClientBack = vbWindowBackground
    mlngBorderColor = vbWindowText
End Sub

Private Sub UserControl_Resize()
    Dim lngWidth As Long
    
    With UserControl
        .shpBorder.Move 0, mlngTop, .ScaleWidth, .ScaleHeight - mlngTop
        .picContainer.Move mlngPixelX, mlngTop + mlngPixelY, .ScaleWidth - mlngPixelX * 2, .ScaleHeight - mlngTop - mlngPixelY * 2
        .scrollVertical.Move .picContainer.ScaleWidth - .scrollVertical.Width, 0, .scrollVertical.Width, .picContainer.ScaleHeight
        ShowScrollbar
    End With
End Sub

Public Sub RefreshColors()
    Dim lngRow As Long
    Dim lngCol As Long
    
    mlngTextColor = cfg.GetColor(cgeWorkspace, cveText)
    mlngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    mlngClientBack = cfg.GetColor(cgeControls, cveBackground)
    mlngBorderColor = cfg.GetColor(cgeWorkspace, cveBorderExterior)
    ' Get colors
    With UserControl
        .ForeColor = mlngTextColor
        .BackColor = mlngBackColor
        .picContainer.BackColor = mlngClientBack
        .picClient.ForeColor = cfg.GetColor(cgeControls, cveText)
        .picClient.BackColor = mlngClientBack
        .shpBorder.BorderColor = mlngBorderColor
    End With
    ' Redraw everything
    DrawColumnHeaders
    For lngRow = 1 To mlngRows
        UserControl.usrSlot(lngRow).RefreshColors
        For lngCol = 1 To mlngColumns
            DrawText lngRow, lngCol, False
        Next
    Next
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.picContainer.hwnd
End Property


' ************* USER INIT *************


Public Sub DefineDimensions(ByVal plngRows As Long, plngColumns As Long, plngSlotColumn As Long)
    mlngRows = plngRows
    mlngColumns = plngColumns
    mlngSlotColumn = plngSlotColumn
    ReDim mtypColumn(1 To mlngColumns)
    mtypColumn(plngSlotColumn).Slot = True
    ReDim grid(mlngColumns, mlngRows)
    UserControl.picClient.Height = (mlngHeight + mlngMarginY) * mlngRows + mlngMarginY
    ShowScrollbar
End Sub

Public Sub DefineColumn(plngColumn As Long, penAlignment As AlignmentConstants, pstrHeader As String, Optional pstrSizeText As String)
    With mtypColumn(plngColumn)
        .Align = penAlignment
        .Header = pstrHeader
        .SizeText = pstrSizeText
    End With
End Sub

Public Sub Refresh()
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long
    
    With UserControl
        .ForeColor = mlngTextColor
        .BackColor = mlngBackColor
        .picContainer.BackColor = mlngClientBack
        .picClient.BackColor = mlngClientBack
        .shpBorder.BorderColor = mlngBorderColor
        .FontBold = True
    End With
    lngLeft = mlngMarginX
    ' Start from left
    For i = 1 To mlngSlotColumn
        With mtypColumn(i)
            .Left = lngLeft
            If i < mlngSlotColumn Then
                .Width = UserControl.TextWidth(.SizeText)
                .Right = .Left + .Width
                lngLeft = .Right + mlngMarginX
            End If
        End With
    Next
    ' Finish from right
    lngRight = UserControl.picClient.ScaleWidth - mlngMarginX
    If mlngColumns > 0 Then
        For i = mlngColumns To mlngSlotColumn Step -1
            With mtypColumn(i)
                .Right = lngRight
                If i > mlngSlotColumn Then
                    .Width = UserControl.TextWidth(.SizeText)
                    .Left = .Right - .Width
                    lngRight = .Left - mlngMarginX
                Else
                    .Width = .Right - .Left
                End If
            End With
        Next
    End If
    UserControl.FontBold = False
    DrawColumnHeaders
    LoadSlots
    For lngRow = 1 To mlngRows
        For lngCol = 1 To mlngColumns
            DrawText lngRow, lngCol, False
        Next
    Next
End Sub

Public Sub Clear()
    Dim i As Long
    
    For i = 1 To mlngRows
        UserControl.usrSlot(i).Clear
    Next
    mlngRows = 0
    mlngColumns = 0
    mlngSlotColumn = 0
    UserControl.scrollVertical.Value = 0
    UserControl.scrollVertical.Visible = False
    UserControl.picClient.Cls
'    UserControl.picClient.Top = 0
    LoadSlots
End Sub

Public Sub GotoTop()
    UserControl.picClient.Top = 0
    UserControl.scrollVertical.Value = 0
End Sub

Public Sub SwapSlots(ByVal plngFrom As Long, ByVal plngTo As Long)
    Dim strCaption As String
    Dim lngRanks As Long
    Dim lngMaxRanks As Long
    Dim lngItemData As Long
    
    With UserControl.usrSlot(plngTo)
        ' To to Variables
        strCaption = .Caption
        lngRanks = .Ranks
        lngMaxRanks = .MaxRanks
        lngItemData = .ItemData
        ' From to To
        .Caption = UserControl.usrSlot(plngFrom).Caption
        .Ranks = UserControl.usrSlot(plngFrom).Ranks
        .MaxRanks = UserControl.usrSlot(plngFrom).MaxRanks
        .ItemData = UserControl.usrSlot(plngFrom).ItemData
        .Refresh
    End With
    With UserControl.usrSlot(plngFrom)
        ' Variables to From
        .Caption = strCaption
        .Ranks = lngRanks
        .MaxRanks = lngMaxRanks
        .ItemData = lngItemData
        .Refresh
    End With
End Sub


' ************* ROWS *************


Public Sub ForceVisible(plngRow As Long)
    With UserControl
        If .picClient.Top + .usrSlot(plngRow).Top < 0 Then
            If plngRow = 1 Then .scrollVertical.Value = 0 Else .scrollVertical.Value = .usrSlot(plngRow).Top
        ElseIf .picClient.Top + .usrSlot(plngRow).Top + mlngHeight > .picContainer.ScaleHeight Then
            If plngRow = mlngRows Then .scrollVertical.Value = .scrollVertical.Max Else .scrollVertical.Value = .usrSlot(plngRow).Top + mlngHeight - .picContainer.ScaleHeight
        End If
    End With
End Sub

Public Property Get Selected() As Long
    Dim enDropState As DropStateEnum
    Dim lngSelected As Long
    Dim i As Long
    
    With UserControl
        For i = 1 To mlngRows
            Select Case .usrSlot(i).DropState
                Case dsDefault
                Case dsCanDrop
                    If lngSelected = 0 Then
                        lngSelected = i
                    Else
                        lngSelected = 0
                        Exit For
                    End If
                Case Else
                    lngSelected = 0
                    Exit For
            End Select
        Next
    End With
    Selected = lngSelected
End Property

Public Property Let Selected(ByVal plngRow As Long)
    Dim enDropState As DropStateEnum
    Dim i As Long
    
    With UserControl
        For i = 1 To mlngRows
            If i = plngRow Then
                enDropState = dsCanDrop
            Else
                enDropState = dsDefault
            End If
            .usrSlot(i).DropState = enDropState
        Next
    End With
End Property

Public Sub ClearSlot(ByVal plngRow As Long)
    UserControl.usrSlot(plngRow).Clear
End Sub

Public Sub SetSlot(ByVal plngRow As Long, pstrText As String, Optional plngRanks As Long, Optional plngMaxRanks As Long)
    With UserControl.usrSlot(plngRow)
        .Caption = pstrText
        .MaxRanks = plngMaxRanks
        .Ranks = plngRanks
        .Refresh
    End With
End Sub

Public Function GetCaption(plngRow As Long) As String
    GetCaption = UserControl.usrSlot(plngRow).Caption
End Function

Public Sub SetText(plngRow As Long, plngColumn As Long, ByVal pstrValue As String)
    grid(plngColumn, plngRow) = pstrValue
    DrawText plngRow, plngColumn, False
End Sub

Public Function GetText(plngRow As Long, plngColumn As Long) As String
    GetText = grid(plngColumn, plngRow)
End Function

Public Sub SetItemData(plngRow As Long, plngItemData As Long)
    UserControl.usrSlot(plngRow).ItemData = plngItemData
End Sub

Public Function GetItemData(plngRow As Long) As Long
    If plngRow <= mlngRows Then GetItemData = UserControl.usrSlot(plngRow).ItemData
End Function

Public Sub SetError(plngRow As Long, pblnError As Boolean)
    UserControl.usrSlot(plngRow).Error = pblnError
End Sub

Public Function GetError(plngRow As Long) As Boolean
    GetError = UserControl.usrSlot(plngRow).Error
End Function

Public Property Get NoFocus() As Boolean
    NoFocus = mblnNoFocus
End Property

Public Property Let NoFocus(ByVal pblnNoFocus As Boolean)
    mblnNoFocus = pblnNoFocus
End Property

Public Property Get Rows() As Long
    Rows = mlngRows
End Property

Public Property Let Rows(ByVal plngRows As Long)
    mlngRows = plngRows
    ReDim Preserve grid(mlngColumns, mlngRows)
    UserControl.picClient.Height = mlngMarginY * 2 + mlngRows * (mlngHeight + mlngMarginY)
    ShowScrollbar
    LoadSlots
End Property

Public Sub SetMandatory(plngRow As Long, pblnMandatory As Boolean)
    UserControl.usrSlot(plngRow).Mandatory = pblnMandatory
End Sub

Public Function GetMandatory(ByVal plngRow As Long) As Boolean
    GetMandatory = UserControl.usrSlot(plngRow).Mandatory
End Function

Public Sub SetFree(plngRow As Long, pblnFree As Boolean)
    UserControl.usrSlot(plngRow).Free = pblnFree
End Sub

Public Function GetFree(ByVal plngRow As Long) As Boolean
    GetFree = UserControl.usrSlot(plngRow).Free
End Function


' ************* DRAWING *************


Private Function LoadSlots()
    Dim lngTop As Long
    Dim i As Long
    
    With UserControl
        For i = .usrSlot.UBound To mlngRows + 1 Step -1
            Unload .usrSlot(i)
        Next
        For i = .usrSlot.UBound + 1 To mlngRows
            Load .usrSlot(i)
            .usrSlot(i).RefreshColors
        Next
        For i = 1 To mlngRows
            lngTop = GetTop(i)
            With mtypColumn(mlngSlotColumn)
                UserControl.usrSlot(i).Move .Left, lngTop, .Width, mlngHeight
                UserControl.usrSlot(i).SetValue vbNullString, 0, 0
                UserControl.usrSlot(i).Visible = True
            End With
        Next
    End With
End Function

Private Sub DrawColumnHeaders()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim i As Long
    
    With UserControl
        .Cls
        lngTop = .picContainer.Top - UserControl.TextHeight("X") - mlngPixelX * 2
    End With
    For i = 1 To mlngColumns
        With mtypColumn(i)
            Select Case .Align
                Case vbLeftJustify: lngLeft = .Left
                Case vbCenter: lngLeft = .Left + (.Width - UserControl.TextWidth(.Header)) \ 2
                Case vbRightJustify: lngLeft = .Right - UserControl.TextWidth(.Header)
            End Select
            UserControl.CurrentX = lngLeft
            UserControl.CurrentY = lngTop
            UserControl.Print .Header
        End With
    Next
End Sub

Private Sub DrawText(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Dim strText As String
    Dim lngTextWidth As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngBottom As Long
    
    If mlngRows = 0 Or mlngColumns = 0 Then Exit Sub
    lngTop = GetTop(plngRow)
    lngBottom = lngTop + mlngHeight
    With mtypColumn(plngCol)
        If .Slot Then
            UserControl.usrSlot(plngRow).Active = pblnActive
        Else
            UserControl.picClient.Line (.Left, lngTop)-(.Right, lngBottom), mlngClientBack, BF
            strText = grid(plngCol, plngRow)
            If Len(strText) Then
                UserControl.picClient.FontBold = pblnActive
                lngTextWidth = UserControl.picClient.TextWidth(strText)
                Select Case mtypColumn(plngCol).Align
                    Case vbLeftJustify: lngLeft = .Left
                    Case vbCenter: lngLeft = .Left + (.Width - lngTextWidth) \ 2
                    Case vbRightJustify: lngLeft = .Right - lngTextWidth
                End Select
                UserControl.picClient.CurrentX = lngLeft
                UserControl.picClient.CurrentY = lngTop + mlngOffsetY
                UserControl.picClient.Print strText
            End If
        End If
    End With
End Sub

Private Function GetTop(plngRow As Long) As Long
    GetTop = (mlngHeight + mlngMarginY) * (plngRow - 1) + mlngMarginY
End Function


' ************* SCROLLBAR *************


Private Sub ShowScrollbar()
    With UserControl
        If .picClient.Height - mlngMarginY > .picContainer.ScaleHeight Then
            .picClient.Width = .picContainer.ScaleWidth - .scrollVertical.Width
            .scrollVertical.Max = .picClient.Height - .picContainer.ScaleHeight
            .scrollVertical.SmallChange = mlngHeight + mlngMarginY
            .scrollVertical.LargeChange = .picContainer.ScaleHeight
            .scrollVertical.Visible = True
        Else
            .picClient.Width = .picContainer.ScaleWidth
            .scrollVertical.Visible = False
        End If
    End With
    Me.Refresh
End Sub

Public Property Get ScrollPosition() As Long
    With UserControl.scrollVertical
        If .Visible Then ScrollPosition = .Value Else ScrollPosition = -1
    End With
End Property

Public Property Let ScrollPosition(ByVal plngScrollPosition As Long)
    Dim lngValue As Long
    
    With UserControl.scrollVertical
        If .Visible And plngScrollPosition <> -1 Then
            If plngScrollPosition > .Max Then lngValue = .Max Else lngValue = plngScrollPosition
            .Value = lngValue
        End If
    End With
End Property

Private Sub scrollVertical_GotFocus()
    If Not mblnNoFocus Then UserControl.picClient.SetFocus
End Sub

Private Sub scrollVertical_Change()
    ScrollClient
End Sub

Private Sub scrollVertical_Scroll()
    ScrollClient
End Sub

Private Sub ScrollClient()
    With UserControl
        .picClient.Top = 0 - .scrollVertical.Value
    End With
End Sub

Public Sub Scroll(plngValue As Long)
    Dim lngIncrement As Long
    Dim lngValue As Long
    
    If Not UserControl.scrollVertical.Visible Then Exit Sub
    lngIncrement = plngValue * (mlngHeight + mlngMarginY)
    With UserControl.scrollVertical
        lngValue = .Value - lngIncrement
        If lngValue < .Min Then lngValue = .Min
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub


' ************* DRAGGING *************


Public Sub SetDropState(ByVal plngRow As Long, penDropState As DropStateEnum)
    UserControl.usrSlot(plngRow).DropState = penDropState
End Sub

Public Function GetDropState(ByVal plngRow As Long) As DropStateEnum
    GetDropState = UserControl.usrSlot(plngRow).DropState
End Function

Private Sub usrSlot_MouseClick(Index As Integer, Button As Integer)
    RaiseEvent SlotClick(Index, Button)
End Sub

Private Sub usrSlot_DblClick(Index As Integer)
    RaiseEvent SlotDblClick(Index)
End Sub

Private Sub usrSlot_OLEStartDrag(Index As Integer, Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Index, Data, AllowedEffects)
End Sub

Private Sub usrSlot_OLEDragDrop(Index As Integer, Data As DataObject)
    RaiseEvent OLEDragDrop(Index, Data)
End Sub

Private Sub usrSlot_OLECompleteDrag(Index As Integer, Effect As Long)
    RaiseEvent OLECompleteDrag(Index, Effect)
End Sub


' ************* ACTIVE *************


Public Property Get Active() As Long
    Active = mlngActive
End Property

Public Property Let Active(ByVal plngSlot As Long)
    If mlngActive = plngSlot Then Exit Property
    If mlngActive And (mlngActive <= mlngRows) Then HighlightRow mlngActive, False
    If plngSlot = 0 Or plngSlot > mlngRows Then
        mlngActive = 0
    Else
        mlngActive = plngSlot
        HighlightRow mlngActive, True
    End If
End Property

Public Sub ForceActive(ByVal plngSlot As Long)
    HighlightRow plngSlot, True
End Sub

Private Sub HighlightRow(plngRow As Long, pblnActive As Boolean)
    Dim i As Long
    
    For i = 1 To mlngColumns
        If i = mlngSlotColumn Then
            If mlngRows > 0 And mlngActive <= mlngRows Then UserControl.usrSlot(mlngActive).Active = pblnActive
        Else
            DrawText plngRow, i, pblnActive
        End If
    Next
End Sub

Private Sub usrSlot_MouseMove(Index As Integer)
    Active = Index
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Active = 0
End Sub

Private Sub picClient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Active = 0
End Sub

Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Active = 0
End Sub

Private Sub usrSlot_RankChange(Index As Integer, Ranks As Long)
    RaiseEvent RankChange(Index, Ranks)
End Sub

Private Sub usrSlot_RequestDrag(Index As Integer, Allow As Boolean, DragIndex As Long)
    DragIndex = Index
    RaiseEvent RequestDrag(Index, Allow)
End Sub

