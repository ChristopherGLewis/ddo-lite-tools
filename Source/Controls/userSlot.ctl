VERSION 5.00
Begin VB.UserControl userSlot 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   288
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   300.522
   ScaleMode       =   0  'User
   ScaleWidth      =   3000
   Begin VB.PictureBox picArrows 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2520
      ScaleHeight     =   252
      ScaleWidth      =   372
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   372
   End
End
Attribute VB_Name = "userSlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Public Enum DropStateEnum
    dsDefault
    dsCanDrop
    dsCanDropError
    dsSubordinate
End Enum

Private Enum ArrowStyle
    asFatTriangle
    asSkinnyTriangle
    asArrow
End Enum

Private Enum MouseArrowEnum
    maeNone
    maeUp
    maeDown
    maeRemember
End Enum

Public Event MouseClick(Button As Integer)
Public Event DblClick()
Public Event MouseMove()
Public Event OLEDragDrop(Data As DataObject)
Public Event OLECompleteDrag(Effect As Long)
Public Event RankChange(Ranks As Long)
Public Event RequestDrag(Allow As Boolean, DragIndex As Long)

Private mlngForeColorDefault As Long
Private mlngForeColorError As Long

Private mlngBackColorDefault As Long
Private mlngBackColorError As Long
Private mlngBackColorDrop As Long
Private mlngBackColorSub As Long

Private mlngBorderColorDefault As Long
Private mlngBorderColorActive As Long

Private mlngMarginLeft As Long
Private mlngMarginTop As Long
Private mlngRight As Long
Private mlngBottom As Long

Private mlngArrowWidth As Long
Private mlngArrowUp As Long
Private mlngArrowDown As Long
Private menMouseArrow As MouseArrowEnum

Private menDropState As DropStateEnum
Private mblnError As Boolean
Private mblnActive As Boolean
Private mstrCaption As String

Private mlngRanks As Long
Private mlngMaxRanks As Long
Private mlngItemData As Long

Private mblnMandatory As Boolean
Private mblnFree As Boolean

Private mlngDragIndex As Long
Private menDragState As DragEnum
Private msngDownX As Single
Private msngDownY As Single

Private mlngPixelX As Long
Private mlngPixelY As Long


' ************* INITIALIZE *************


Private Sub UserControl_Initialize()
    mlngForeColorDefault = vbBlack
    mlngForeColorError = vbRed
    mlngBackColorDefault = RGB(245, 245, 245)
    mlngBackColorError = RGB(255, 245, 245)
    mlngBackColorDrop = RGB(255, 255, 245)
    mlngBackColorSub = RGB(245, 245, 255)
    mlngBorderColorDefault = RGB(175, 175, 175)
    menDropState = dsDefault
    mlngBorderColorActive = vbBlack
    mlngRanks = 0
    mlngMaxRanks = 0
    With UserControl
        mlngMarginLeft = .TextWidth(" ")
        mlngMarginTop = (.Height - .TextHeight("X")) \ 2
        mlngRight = .ScaleWidth - Screen.TwipsPerPixelX
        mlngBottom = .ScaleHeight - Screen.TwipsPerPixelY
    End With
    Refresh
End Sub

Private Sub UserControl_Resize()
    With UserControl
        mlngMarginTop = (.Height - .TextHeight("X")) \ 2
        mlngRight = .ScaleWidth - Screen.TwipsPerPixelX
        mlngBottom = .ScaleHeight - Screen.TwipsPerPixelY
    End With
    Refresh
End Sub


' ************* DRAWING *************


Public Sub RefreshColors()
    mlngForeColorDefault = cfg.GetColor(cgeDropSlots, cveText)
    mlngForeColorError = cfg.GetColor(cgeDropSlots, cveTextError)
    mlngBackColorDefault = cfg.GetColor(cgeDropSlots, cveBackground)
    mlngBackColorError = cfg.GetColor(cgeDropSlots, cveBackError)
    mlngBackColorDrop = cfg.GetColor(cgeDropSlots, cveBackHighlight)
    mlngBackColorSub = cfg.GetColor(cgeDropSlots, cveBackRelated)
    mlngBorderColorDefault = cfg.GetColor(cgeDropSlots, cveBorderExterior)
    mlngBorderColorActive = cfg.GetColor(cgeDropSlots, cveBorderHighlight)
    UserControl.picArrows.BackColor = mlngBackColorDefault
    Refresh
End Sub

Public Sub Refresh()
    Dim strText As String
    Dim lngColor As Long
    
    With UserControl
        ' Background
        Select Case menDropState
            Case dsDefault: lngColor = mlngBackColorDefault
            Case dsCanDrop: lngColor = mlngBackColorDrop
            Case dsCanDropError: lngColor = mlngBackColorError
            Case dsSubordinate: lngColor = mlngBackColorSub
        End Select
        .BackColor = lngColor
        .picArrows.BackColor = lngColor
        .Cls
        ' Up/Down arrows
        DrawArrows
        ' Text
        If Len(mstrCaption) Then
            If mblnError Then lngColor = mlngForeColorError Else lngColor = mlngForeColorDefault
            .ForeColor = lngColor
            .CurrentX = mlngMarginLeft
            .CurrentY = mlngMarginTop
            UserControl.Print DisplayText()
        End If
        ' Border
        If mblnActive Then lngColor = mlngBorderColorActive Else lngColor = mlngBorderColorDefault
        UserControl.Line (0, 0)-(mlngRight, mlngBottom), lngColor, B
    End With
End Sub

Private Function DisplayText() As String
    Const Dots As String = "..."
    Dim strBase As String
    Dim strRank As String
    Dim lngWidth As Long
    Dim strText As String
    Dim strWhitespace As String
    
    With UserControl
        strBase = mstrCaption
        lngWidth = .ScaleWidth - mlngMarginLeft * 2
        If mlngMaxRanks > 1 Then
            lngWidth = lngWidth - mlngArrowWidth
           strRank = Left$("III", mlngRanks)
            strWhitespace = " "
        End If
         Do
            strText = strBase & strWhitespace & strRank
            If .TextWidth(strText) <= lngWidth Then Exit Do
            strWhitespace = Dots
            strBase = Trim$(Left$(strBase, Len(strBase) - 1))
        Loop
    End With
    DisplayText = strText
End Function

Private Sub DrawArrows()
    Dim lngCenter As Long
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim i As Long
    
    If mlngMaxRanks < 2 Then
        UserControl.picArrows.Visible = False
        mlngArrowWidth = 0
        Exit Sub
    End If
    With Screen
        mlngPixelX = .TwipsPerPixelX
        mlngPixelY = .TwipsPerPixelY
    End With
    With UserControl
        lngHeight = ((.ScaleHeight - (mlngPixelY * 5)) \ 2) \ mlngPixelY
        ' Convert from Y to X
        mlngArrowWidth = lngHeight * mlngPixelX
        mlngArrowWidth = (mlngArrowWidth * 2)
        .picArrows.Move .ScaleWidth - mlngArrowWidth - mlngPixelX, mlngPixelY, mlngArrowWidth, .ScaleHeight - mlngPixelY * 2
        lngCenter = mlngArrowWidth \ 2
        DrawUpDown lngHeight, lngCenter, asSkinnyTriangle
        .picArrows.Visible = True
        mlngArrowDown = .picArrows.ScaleHeight \ 2
        mlngArrowUp = mlngArrowDown - mlngPixelY
    End With
End Sub

Private Sub DrawUpDown(plngHeight As Long, plngCenter As Long, penStyle As ArrowStyle)
    Select Case penStyle
        Case asFatTriangle: DrawFatTriangle plngHeight, plngCenter
        Case asSkinnyTriangle: DrawSkinnyTriangle plngHeight, plngCenter
        Case asArrow: DrawArrow plngHeight, plngCenter
    End Select
End Sub

Private Sub DrawFatTriangle(plngHeight As Long, plngCenter As Long)
    Dim lngWidth As Long
    Dim i As Long
    
    For i = 1 To plngHeight
        lngWidth = i * 2 - 1
        DrawLines i, plngCenter, lngWidth
    Next
End Sub

Private Sub DrawSkinnyTriangle(plngHeight As Long, plngCenter As Long)
    Dim i As Long
    
    For i = 1 To plngHeight
        DrawLines i, plngCenter, i
    Next
End Sub

Private Sub DrawArrow(plngHeight As Long, plngCenter As Long)
    Dim lngWidth As Long
    Dim lngStem As Long
    Dim lngStemWidth As Long
    Dim i As Long
    
    lngStem = plngHeight * 2 \ 3
    lngStemWidth = plngHeight \ 2
    For i = 1 To plngHeight
        If i > lngStem Then lngWidth = lngStemWidth Else lngWidth = i * 2 - 1
        DrawLines i, plngCenter, lngWidth
    Next
End Sub

Private Sub DrawLines(plngY As Long, plngCenter As Long, plngWidth As Long)
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngY As Long
    Dim lngWidth As Long
    Dim lngUpper As Long
    Dim lngLower As Long
    
    lngWidth = plngWidth * mlngPixelX
    lngY = plngY * mlngPixelY
    lngLeft = plngCenter - lngWidth \ 2
    lngRight = lngLeft + lngWidth
    lngUpper = lngY - mlngPixelY
    lngLower = UserControl.ScaleHeight - lngY - mlngPixelY * 3
    UserControl.picArrows.Line (lngLeft, lngUpper)-(lngRight, lngUpper), mlngForeColorDefault
    UserControl.picArrows.Line (lngLeft, lngLower)-(lngRight, lngLower), mlngForeColorDefault
End Sub


' ************* PROPERTIES *************


Public Sub Clear()
    mlngRanks = 0
    mlngMaxRanks = 0
    mlngItemData = 0
    mblnError = False
    mblnFree = False
    mblnMandatory = False
    menDropState = dsDefault
    mstrCaption = vbNullString
    Me.Refresh
End Sub

Public Sub SetValue(pstrCaption As String, Optional plngRanks As Long, Optional plngMaxRanks As Long)
    mstrCaption = pstrCaption
    mlngRanks = plngRanks
    mlngMaxRanks = plngMaxRanks
    Refresh
End Sub

Public Property Get Active() As Boolean
    Active = mblnActive
End Property

Public Property Let Active(ByVal pblnActive As Boolean)
    mblnActive = pblnActive
    Refresh
End Property

Public Property Get Caption() As String
    Caption = mstrCaption
End Property

Public Property Let Caption(ByVal pstrCaption As String)
    mstrCaption = pstrCaption
    If Len(mstrCaption) = 0 Then
        mlngRanks = 0
        mlngMaxRanks = 0
    End If
End Property

Public Property Get DropState() As DropStateEnum
    DropState = menDropState
End Property

Public Property Let DropState(penDropState As DropStateEnum)
    menDropState = penDropState
    With UserControl
        If menDropState = dsDefault Then .OLEDropMode = vbOLEDropNone Else .OLEDropMode = vbOLEDropManual
    End With
    Refresh
End Property

Public Property Get Error() As Boolean
    Error = mblnError
End Property

Public Property Let Error(ByVal pblnError As Boolean)
    mblnError = pblnError
    Refresh
End Property

Public Property Get Free() As Boolean
    Free = mblnFree
End Property

Public Property Let Free(ByVal pblnFree As Boolean)
    mblnFree = pblnFree
    Refresh
End Property

Public Property Get ItemData() As Long
    ItemData = mlngItemData
End Property

Public Property Let ItemData(ByVal plngItemData As Long)
    mlngItemData = plngItemData
End Property

Public Property Get Mandatory() As Boolean
    Mandatory = mblnMandatory
End Property

Public Property Let Mandatory(ByVal pblnMandatory As Boolean)
    mblnMandatory = pblnMandatory
    Refresh
End Property

Public Property Get MarginLeft() As Long
    MarginLeft = mlngMarginLeft
End Property

Public Property Let MarginLeft(ByVal plngMarginLeft As Long)
    mlngMarginLeft = plngMarginLeft
    If Len(mstrCaption) Then Refresh
End Property

Public Property Get MarginTop() As Long
    MarginTop = mlngMarginTop
End Property

Public Property Get MaxRanks() As Long
    MaxRanks = mlngMaxRanks
End Property

Public Property Let MaxRanks(ByVal plngMaxRanks As Long)
    If mlngMaxRanks <> plngMaxRanks Then
        mlngMaxRanks = plngMaxRanks
        If mlngRanks = 0 Or mlngRanks > mlngMaxRanks Then mlngRanks = mlngMaxRanks
    End If
End Property

Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal penDropMode As OLEDropConstants)
    UserControl.OLEDropMode = OLEDropMode
End Property

Public Property Get Ranks() As Long
    Ranks = mlngRanks
End Property

Public Property Let Ranks(ByVal plngRanks As Long)
    If mlngRanks <> plngRanks Then
        mlngRanks = plngRanks
        If mlngRanks > mlngMaxRanks Then mlngRanks = mlngMaxRanks
    End If
End Property


' ************* RANKS *************


Private Sub picArrows_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove
    If Y < mlngArrowUp Or Y > mlngArrowDown Then xp.SetMouseCursor mcHand Else menMouseArrow = maeNone
End Sub

Private Sub picArrows_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Select Case Y
            Case Is < mlngArrowUp: ClickArrow maeUp
            Case Is > mlngArrowDown: ClickArrow maeDown
            Case Else: menMouseArrow = maeNone
        End Select
    End If
End Sub

Private Sub picArrows_DblClick()
    If menMouseArrow <> maeNone Then ClickArrow maeRemember
End Sub

Private Sub ClickArrow(penMouseArrow As MouseArrowEnum)
    If penMouseArrow <> maeRemember Then menMouseArrow = penMouseArrow
    If menMouseArrow <> maeNone Then xp.SetMouseCursor mcHand
    Select Case menMouseArrow
        Case maeUp
            If mlngRanks < mlngMaxRanks Then
                mlngRanks = mlngRanks + 1
                Refresh
                RaiseEvent RankChange(mlngRanks)
            End If
        Case maeDown
            If mlngRanks > 1 And mlngMaxRanks <> 0 Then
                mlngRanks = mlngRanks - 1
                Refresh
                RaiseEvent RankChange(mlngRanks)
            End If
    End Select
End Sub


' ************* MOUSE *************


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnAllow As Boolean
    
    menMouseArrow = maeNone
    RaiseEvent MouseMove
    If Button = vbLeftButton Then
        If menDragState = dragMouseDown Then
            menDragState = dragMouseMove
            msngDownX = X
            msngDownY = Y
        ElseIf menDragState = dragMouseMove Then
            ' Only start dragging if mouse actually moved
            If X <> msngDownX Or Y <> msngDownY Then
                menDragState = dragNormal
                RaiseEvent RequestDrag(blnAllow, mlngDragIndex)
                If blnAllow Then UserControl.OLEDrag
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        menDragState = dragMouseDown
    End If
    RaiseEvent MouseClick(Button)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData mlngDragIndex
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub
