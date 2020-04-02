VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatusbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' // Written by Ellis Dee of vbforums.com //
Option Explicit

Private Const Margin = 7
Private Const Spacer = "o"

Public Event SliderChange(Value As Long)
Public Event PanelClick(Panel As Integer)
Public Event PanelDoubleClick(Panel As Integer)

Public Enum PanelSizeEnum
    pseFixed = 0
    pseAutoFit
    pseSpring
End Enum

Private Type PanelType
    Left As Long
    Width As Long
    Right As Long
    PublicWidth As Long
    Text As String
    Align As AlignmentConstants
    Size As PanelSizeEnum
End Type

Private Type ProgressType
    Panel As Long
    Value As Long
    Max As Long
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    Side As Long
    Blocks As Long
    NextCutoff As Long
    NextBlock As Long
    LastWidth As Long
End Type

Private Type SliderType
    Hot As Boolean
    Panel As Long
    Value As Long
    Min As Long
    Max As Long
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    Width As Long
End Type

' General
Private mpic As PictureBox
Private mtypProgress As ProgressType
Private mtypSlider As SliderType
Private mtypPanel() As PanelType
Private mlngPanels As Long
Private mstrSimpleText As String
Private mblnMaximized As Boolean
Private mblnLoaded As Boolean
' Coordinates / dimensions
Private mlngX As Long
Private mlngY As Long
Private mlngOffsetX As Long
Private mlngOffsetY As Long
Private mlngLineBottom As Long

' API for showing hand cursor when over slider
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long



' METHODS


Public Sub Init(ppic As PictureBox)
    Set mpic = ppic
    With mpic
        If .Parent.ScaleMode = vbPixels Then
            mlngOffsetX = 1
            mlngOffsetY = 1
        Else
            mlngOffsetX = Screen.TwipsPerPixelX
            mlngOffsetY = Screen.TwipsPerPixelY
        End If
        .Appearance = 0
        .BorderStyle = vbBSNone
        .AutoRedraw = True
        Set .Font = .Parent.Font
        .ScaleMode = vbPixels
        .BorderStyle = vbBSNone
        .BackColor = vbButtonFace
        .Cls
        .Visible = True
    End With
    mtypProgress.Panel = 0
    mtypSlider.Panel = 0
    mlngPanels = 0
    mblnLoaded = True
End Sub

Public Sub AddPanel(pstrCaption As String, penAlign As AlignmentConstants, penSize As PanelSizeEnum, plngWidthInPixels As Long, Optional pblnRedraw As Boolean = True)
    mlngPanels = mlngPanels + 1
    ReDim Preserve mtypPanel(1 To mlngPanels)
    With mtypPanel(mlngPanels)
        .Text = pstrCaption
        .Align = penAlign
        .Size = penSize
        .PublicWidth = plngWidthInPixels
    End With
    If mblnLoaded And pblnRedraw Then CalculateDimensions
End Sub

Public Sub Click()
    RaiseEvent PanelClick(CurrentPanel())
End Sub

Public Sub DoubleClick()
    RaiseEvent PanelDoubleClick(CurrentPanel())
End Sub

Public Function GetPanelAlign(plngPanel As Long) As AlignmentConstants
    Select Case plngPanel
        Case 1 To mlngPanels: GetPanelAlign = mtypPanel(plngPanel).Align
    End Select
End Function

Public Function GetPanelCaption(plngPanel As Long) As String
    Select Case plngPanel
        Case 1 To mlngPanels: GetPanelCaption = mtypPanel(plngPanel).Text
    End Select
End Function

Public Function GetPanelSize(plngPanel As Long) As PanelSizeEnum
    Select Case plngPanel
        Case 1 To mlngPanels: GetPanelSize = mtypPanel(plngPanel).Size
    End Select
End Function

Public Function GetPanelWidth(plngPanel As Long) As Long
    Select Case plngPanel
        Case 1 To mlngPanels: GetPanelWidth = mtypPanel(plngPanel).PublicWidth
    End Select
End Function

Public Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngX = X
    mlngY = Y
    If mtypSlider.Panel = 0 Then Exit Sub
    If HandCursor(X, Y) Then
        If Button = vbLeftButton Then
            mtypSlider.Hot = True
            HotSlider X
        End If
    End If
End Sub

Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngX = X
    mlngY = Y
    If mtypSlider.Panel = 0 Then Exit Sub
    If mtypSlider.Hot Then
        ' Slider has already been engaged, so use input even if mouse left area
        HotSlider X
    Else
        ' Slider hasn't been engaged; show hand cursor if mouse over slider
        HandCursor X, Y
    End If
End Sub

Public Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngX = X
    mlngY = Y
    If mtypSlider.Panel = 0 Then Exit Sub
    mtypSlider.Hot = False
End Sub

Public Sub ProgressbarIncrement()
    With mtypProgress
        If .Value < .Max Then
            .Value = .Value + 1
            DrawProgress
        End If
    End With
End Sub

Public Sub ProgressbarInit(plngPanel As Long, plngMax As Long)
    Select Case plngPanel
        Case 1 To mlngPanels
            With mtypProgress
                .Panel = plngPanel
                If plngMax < 1 Then .Max = 1 Else .Max = plngMax
                .Value = 0
            End With
            PanelWidths
    End Select
End Sub

Public Sub ProgressbarRemove()
    Dim lngPanel As Long
    
    If mtypProgress.Panel <> 0 Then
        lngPanel = mtypProgress.Panel
        mtypProgress.Panel = 0
        Redraw lngPanel
    End If
End Sub

Public Sub Refresh()
    If mblnLoaded Then mpic.Refresh
End Sub

Public Sub Resize()
    If mblnLoaded Then CalculateDimensions
End Sub

Public Sub SetPanel(plngPanel As Long, pstrCaption As String, penAlign As AlignmentConstants, penSize As PanelSizeEnum, plngWidthInPixels As Long, Optional pblnRedraw As Boolean = True)
    Select Case plngPanel
        Case 1 To mlngPanels
            With mtypPanel(plngPanel)
                .Text = pstrCaption
                .Align = penAlign
                .Size = penSize
                .PublicWidth = plngWidthInPixels
            End With
            If mblnLoaded And pblnRedraw Then CalculateDimensions
    End Select
End Sub

Public Sub SetPanelAlign(plngPanel As Long, penAlign As AlignmentConstants, Optional pblnRedraw As Boolean = True)
    Select Case plngPanel
        Case 1 To mlngPanels
            mtypPanel(plngPanel).Align = penAlign
            If mblnLoaded And pblnRedraw Then Redraw plngPanel
    End Select
End Sub

Public Sub SetPanelCaption(plngPanel As Long, pstrCaption As String, Optional pblnRedraw As Boolean = True)
    Select Case plngPanel
        Case 1 To mlngPanels
            With mtypPanel(plngPanel)
                If .Text <> pstrCaption Then
                    .Text = pstrCaption
                    If mblnLoaded And pblnRedraw Then
                        If .Size = pseAutoFit Then PanelWidths Else Redraw plngPanel
                    End If
                End If
            End With
    End Select
End Sub

Public Sub SetPanelSize(plngPanel As Long, penSize As PanelSizeEnum, Optional pblnRedraw As Boolean = True)
    Select Case plngPanel
        Case 1 To mlngPanels
            mtypPanel(plngPanel).Size = penSize
            If mblnLoaded And pblnRedraw Then PanelWidths
    End Select
End Sub

Public Sub SetPanelWidth(plngPanel As Long, ByVal plngPixels As Long, Optional pblnRedraw As Boolean = True)
    If plngPixels < 0 Then plngPixels = 0
    Select Case plngPanel
        Case 1 To mlngPanels
            mtypPanel(plngPanel).PublicWidth = plngPixels
            If mblnLoaded And pblnRedraw Then PanelWidths
    End Select
End Sub

Public Sub SliderInit(plngPanel As Long, plngMin As Long, plngMax As Long, plngValue As Long)
    With mtypSlider
        .Panel = plngPanel
        .Value = plngValue
        .Min = plngMin
        .Max = plngMax
        If .Value < .Min Then .Value = .Min
        If .Value > .Max Then .Value = .Max
        .Hot = False
    End With
    CalculateDimensions
End Sub

Public Sub SliderRemove()
    mtypSlider.Panel = 0
    CalculateDimensions
End Sub



' PROPERTIES


' Panels
Public Property Let Panels(plngPanels As Long)
    If plngPanels > 0 Then
        mlngPanels = plngPanels
        ReDim Preserve mtypPanel(1 To mlngPanels)
        If mtypProgress.Panel > mlngPanels Then mtypProgress.Panel = mlngPanels
        If mtypSlider.Panel > mlngPanels Then mtypSlider.Panel = mlngPanels
        If mblnLoaded Then CalculateDimensions
    End If
End Property

Public Property Get Panels() As Long
    Panels = mlngPanels
End Property


' Progress Max
Public Property Let ProgressMax(ByVal plngMax As Long)
    With mtypProgress
        If .Panel <> 0 And plngMax > 0 Then
            .Max = plngMax
            .Value = 0
            Redraw .Panel
            ProgressDimensions
        End If
    End With
End Property

Public Property Get ProgressMax() As Long
    If mtypProgress.Panel <> 0 Then ProgressMax = mtypProgress.Max
End Property


' Progressbar value
Public Property Let ProgressbarValue(ByVal plngValue As Long)
    With mtypProgress
        If .Panel <> 0 Then
            Select Case plngValue
                Case Is <= 0: plngValue = 0
                Case Is >= .Max: plngValue = .Max
            End Select
            If .Value <> plngValue Then
                .Value = plngValue
                Redraw .Panel
                ProgressDimensions
            End If
        End If
    End With
End Property

Public Property Get ProgressbarValue() As Long
    ProgressbarValue = mtypProgress.Value
End Property


' Simple Text
Public Property Let SimpleText(pstrSimpleText As String)
    mstrSimpleText = pstrSimpleText
    Redraw
End Property

Public Property Get SimpleText() As String
    SimpleText = mstrSimpleText
End Property


' Slider value
Public Property Let SliderValue(ByVal plngValue As Long)
    With mtypSlider
        If .Panel <> 0 Then
            Select Case plngValue
                Case Is <= .Min: plngValue = .Min
                Case Is >= .Max: plngValue = .Max
            End Select
            If .Value <> plngValue Then
                .Value = plngValue
                Redraw
            End If
        End If
    End With
End Property

Public Property Get SliderValue() As Long
    SliderValue = mtypSlider.Value
End Property



' INTERNAL FUNCTIONS


Private Sub Class_Terminate()
    Erase mtypPanel
    mpic.Cls
    mpic.Visible = False
    Set mpic = Nothing
End Sub

Private Sub CalculateDimensions()
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngOffset As Long
    
    With mpic
        mblnMaximized = (.Parent.WindowState = vbMaximized)
        If Not mblnMaximized Then lngOffset = 1
        lngHeight = (.TextHeight(Spacer) + 4) * mlngOffsetY
        With .Parent
            lngTop = .ScaleHeight - lngHeight
            lngWidth = .ScaleWidth + 2 * mlngOffsetX
        End With
        .Move -mlngOffsetX, lngTop, lngWidth, lngHeight
        mlngLineBottom = .ScaleHeight - lngOffset
    End With
    PanelWidths
End Sub

Private Function CurrentPanel() As Integer
    Dim i As Integer
    
    CurrentPanel = -1
    For i = 1 To mlngPanels
        If mlngX >= mtypPanel(i).Left And mlngX <= mtypPanel(i).Right Then CurrentPanel = i
    Next
End Function

Private Sub DrawProgress()
    Dim lngLeft As Long
    
    With mtypProgress
        If .Blocks > 0 Then
            Do While .Value >= .NextCutoff And .NextBlock <= .Blocks
                lngLeft = .Left + 3 + .NextBlock * (.Side + 2)
                If .NextBlock = .Blocks Then
                    mpic.Line (lngLeft, .Top)-(lngLeft + .LastWidth - 1, .Bottom), vbHighlight, BF
                Else
                    mpic.Line (lngLeft, .Top)-(lngLeft + .Side - 1, .Bottom), vbHighlight, BF
                End If
                mpic.Refresh
                .NextBlock = .NextBlock + 1
                .NextCutoff = (.Max * .NextBlock) / .Blocks
            Loop
        End If
    End With
End Sub

Private Sub DrawSlider()
    Dim lngX As Long
    Dim lngY As Long
    Dim lngValue As Long
    Dim lngMax As Long
    
    With mtypSlider
        Select Case .Value
            Case Is <= .Min: lngX = .Left
            Case Is >= .Max: lngX = .Right
            Case Else: lngX = Round(((.Value - .Min) / (.Max - .Min)) * .Width) + .Left
        End Select
        lngY = .Top + (.Bottom - .Top) \ 2 - 2
    End With
    Select Case mpic.TextHeight(Spacer)
        Case Is > 28
            lngX = lngX - 5
            lngY = lngY - 8
            mpic.Line (lngX, lngY)-(lngX + 11, lngY), vbWhite
            mpic.Line (lngX, lngY)-(lngX, lngY + 15), vbWhite
            mpic.Line (lngX, lngY + 15)-(lngX + 5, lngY + 20), vbWhite
            mpic.Line (lngX + 1, lngY + 1)-(lngX + 8, lngY + 14), vbButtonFace, BF
            lngX = lngX + 9
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 16), vbButtonShadow
            mpic.Line (lngX - 1, lngY + 16)-(lngX - 5, lngY + 20), vbButtonShadow
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 16), vbBlack
            mpic.Line (lngX - 1, lngY + 16)-(lngX - 6, lngY + 21), vbBlack
        Case 25 To 28
            lngX = lngX - 4
            lngY = lngY - 6
            mpic.Line (lngX, lngY)-(lngX + 8, lngY), vbWhite
            mpic.Line (lngX, lngY)-(lngX, lngY + 13), vbWhite
            mpic.Line (lngX, lngY + 12)-(lngX + 4, lngY + 16), vbWhite
            mpic.Line (lngX + 1, lngY + 1)-(lngX + 6, lngY + 12), vbButtonFace, BF
            lngX = lngX + 7
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 13), vbButtonShadow
            mpic.Line (lngX - 1, lngY + 13)-(lngX - 4, lngY + 16), vbButtonShadow
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 13), vbBlack
            mpic.Line (lngX - 1, lngY + 13)-(lngX - 5, lngY + 17), vbBlack
        Case 20 To 24
            lngX = lngX - 3
            lngY = lngY - 5
            mpic.Line (lngX, lngY)-(lngX + 6, lngY), vbWhite
            mpic.Line (lngX, lngY)-(lngX, lngY + 11), vbWhite
            mpic.Line (lngX, lngY + 10)-(lngX + 3, lngY + 13), vbWhite
            mpic.Line (lngX + 1, lngY + 1)-(lngX + 4, lngY + 10), vbButtonFace, BF
            lngX = lngX + 5
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 11), vbButtonShadow
            mpic.Line (lngX - 1, lngY + 11)-(lngX - 3, lngY + 13), vbButtonShadow
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 11), vbBlack
            mpic.Line (lngX - 1, lngY + 11)-(lngX - 4, lngY + 14), vbBlack
        Case 15 To 19
            lngX = lngX - 3
            lngY = lngY - 3
            mpic.Line (lngX, lngY)-(lngX + 6, lngY), vbWhite
            mpic.Line (lngX, lngY)-(lngX, lngY + 9), vbWhite
            mpic.Line (lngX, lngY + 8)-(lngX + 3, lngY + 11), vbWhite
            mpic.Line (lngX + 1, lngY + 1)-(lngX + 4, lngY + 8), vbButtonFace, BF
            lngX = lngX + 5
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 9), vbButtonShadow
            mpic.Line (lngX - 1, lngY + 9)-(lngX - 3, lngY + 11), vbButtonShadow
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 9), vbBlack
            mpic.Line (lngX - 1, lngY + 9)-(lngX - 4, lngY + 12), vbBlack
        Case 11 To 14
            lngX = lngX - 2
            lngY = lngY - 1
            mpic.Line (lngX, lngY)-(lngX + 5, lngY), vbWhite
            mpic.Line (lngX, lngY)-(lngX, lngY + 6), vbWhite
            lngX = lngX + 1
            mpic.Line (lngX, lngY + 1)-(lngX + 1, lngY + 4), vbButtonFace, BF
            mpic.PSet (lngX, lngY + 6), vbWhite
            lngX = lngX + 1
            mpic.PSet (lngX, lngY + 6), vbButtonShadow
            mpic.PSet (lngX, lngY + 7), vbBlack
            lngX = lngX + 1
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 6), vbButtonShadow
            mpic.PSet (lngX, lngY + 6), vbBlack
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 6), vbBlack
        Case 7 To 10
            lngX = lngX - 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 5), vbWhite
            lngX = lngX + 1
            mpic.PSet (lngX, lngY), vbWhite
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 4), vbButtonFace
            mpic.PSet (lngX, lngY + 4), vbButtonShadow
            mpic.PSet (lngX, lngY + 5), vbBlack
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 5), vbBlack
        Case 1 To 6
            lngX = lngX - 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 4), vbWhite
            lngX = lngX + 1
            mpic.PSet (lngX, lngY), vbWhite
            mpic.Line (lngX, lngY + 1)-(lngX, lngY + 3), vbButtonFace
            mpic.PSet (lngX, lngY + 3), vbButtonShadow
            If mpic.TextHeight(Spacer) > 3 Then mpic.PSet (lngX, lngY + 4), vbBlack
            lngX = lngX + 1
            mpic.Line (lngX, lngY)-(lngX, lngY + 4), vbBlack
    End Select
End Sub

Private Sub EraseSlider(pblnClearArea As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    
    With mtypSlider
        lngLeft = .Left
        lngTop = .Top + (.Bottom - .Top) \ 2 - 2
        lngRight = .Right
        If pblnClearArea Then mpic.Line (.Left - Margin + 2, .Top + 1)-(.Right + Margin - 2, .Bottom - 2), vbButtonFace, BF
        mpic.Line (lngLeft, lngTop + 2)-(lngLeft, lngTop), vbButtonShadow
        mpic.Line (lngLeft, lngTop)-(lngRight, lngTop), vbButtonShadow
        lngLeft = lngLeft + 1
        lngTop = lngTop + 1
        mpic.Line (lngLeft, lngTop)-(lngRight - 1, lngTop), vbBlack
        lngTop = lngTop + 2
        lngLeft = lngLeft - 1
        mpic.Line (lngLeft, lngTop)-(lngRight, lngTop), vbWhite
        mpic.Line (lngRight, lngTop - 3)-(lngRight, lngTop + 1), vbWhite
    End With
End Sub

Private Function HandCursor(ByVal plngX As Long, ByVal plngY As Long) As Boolean
    Const IDC_HAND = 32649&
    
    ' Short-circuit comparison logic for efficiency
    If mtypSlider.Panel = 0 Then Exit Function
    If plngX < mtypSlider.Left Then Exit Function
    If plngX > mtypSlider.Right Then Exit Function
    If plngY < mtypSlider.Top Then Exit Function
    If plngY > mtypSlider.Bottom Then Exit Function
    ' There is a slider and the mouse is currently over it
    SetCursor LoadCursor(0, IDC_HAND)
    HandCursor = True
End Function

Private Sub HotSlider(ByVal X As Long)
    Dim lngValue As Long
    
    With mtypSlider
        lngValue = Round(((X - .Left) / .Width) * (.Max - .Min)) + .Min
        Select Case lngValue
            Case Is < .Min: lngValue = .Min
            Case Is > .Max: lngValue = .Max
        End Select
        If .Value <> lngValue Then
            EraseSlider True
            .Value = lngValue
            DrawSlider
            RaiseEvent SliderChange(.Value)
        End If
    End With
End Sub

Private Sub PanelWidths()
    Dim lngTotalWidth As Long
    Dim lngScaleWidth As Long
    Dim lngWidth As Long
    Dim lngExtra As Long
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim i As Long
    Dim lngSpring As Long
    
    For i = 1 To mlngPanels
        With mtypPanel(i)
            Select Case .Size
                Case pseSpring
                    lngSpring = lngSpring + 1
                Case pseFixed
                    .Width = .PublicWidth
                    lngTotalWidth = lngTotalWidth + .Width
                Case pseAutoFit
                    If Len(.Text) = 0 Then .Width = .PublicWidth Else .Width = mpic.TextWidth(.Text) + 10
                    lngTotalWidth = lngTotalWidth + .Width
            End Select
        End With
    Next
    If lngSpring <> 0 Then
        lngScaleWidth = mpic.ScaleWidth
        lngWidth = (lngScaleWidth - lngTotalWidth) \ lngSpring
        If lngWidth < 0 Then lngWidth = 0 Else lngExtra = (lngScaleWidth - lngTotalWidth - lngWidth * lngSpring)
        For i = mlngPanels To 1 Step -1
            With mtypPanel(i)
                If .Size = pseSpring Then
                    .Width = lngWidth + lngExtra
                    lngExtra = 0
                End If
            End With
        Next
    End If
    For i = 1 To mlngPanels
        With mtypPanel(i)
            .Left = lngLeft
            lngLeft = lngLeft + .Width
            .Right = lngLeft - 1
        End With
    Next
    With mtypSlider
        If .Panel <> 0 Then
            .Left = mtypPanel(.Panel).Left + Margin
            .Top = 2
            .Right = mtypPanel(.Panel).Right - Margin
            .Bottom = mpic.ScaleHeight
            .Width = .Right - .Left + 1
        End If
    End With
    Redraw
End Sub

Private Sub ProgressDimensions()
    Dim lngRemaining As Long
    
    With mtypProgress
        .Side = mpic.TextWidth(Spacer)
        If (mlngLineBottom - 2 - .Side) Mod 2 <> 0 Then .Side = .Side - 1
        If .Side < 0 Then .Side = 0
        .Left = mtypPanel(.Panel).Left
        .Top = 2 + (mlngLineBottom - 2 - .Side) \ 2
        .Right = mtypPanel(.Panel).Right
        .Bottom = .Top + .Side
        .Blocks = (.Right - .Left - 2) \ (.Side + 2)
        lngRemaining = .Right - .Left - 2 - (.Side + 2) * .Blocks
        Select Case lngRemaining
            Case 0, 1
                .LastWidth = .Side - (2 - lngRemaining)
            Case 2 To 4
                .LastWidth = .Side
            Case Else
                .Blocks = .Blocks + 1
                .LastWidth = lngRemaining - 4
        End Select
        ' Convert to zero-based
        .Blocks = .Blocks - 1
        .NextBlock = 0
        .NextCutoff = 1
    End With
    DrawProgress
End Sub

Private Sub Redraw(Optional plngPanel As Long)
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngWidth As Long
    Dim strText As String
    Dim lngLen As Long
    Dim lngTextWidth As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    
    ' Simple text
    If Len(mstrSimpleText) <> 0 Then
        If plngPanel = 0 Then
            lngLeft = 1
            lngRight = mpic.ScaleWidth - 2
            mpic.Line (lngLeft, 0)-(lngRight, mlngLineBottom), vbButtonFace, BF
            mpic.Line (lngLeft, 2)-(lngRight, 2), vbButtonShadow
            mpic.Line (lngLeft, 2)-(lngLeft, mlngLineBottom), vbButtonShadow
            mpic.Line (lngRight, 2)-(lngRight, mlngLineBottom), vbWhite
            If Not mblnMaximized Then mpic.Line (lngLeft, mlngLineBottom)-(lngRight + 1, mlngLineBottom), vbWhite
            ' Draw text
            lngWidth = lngRight - lngLeft - 7
            strText = mstrSimpleText
            lngLen = Len(strText)
            lngTextWidth = mpic.TextWidth(strText)
            If lngLen <> 0 And lngTextWidth > lngWidth Then
                lngLen = lngLen + 2
                strText = Space$(lngLen)
                Mid$(strText, 1, lngLen - 2) = mstrSimpleText
                Mid$(strText, lngLen - 2, 3) = "..."
                lngTextWidth = mpic.TextWidth(strText)
            End If
            Do While lngLen <> 0 And lngTextWidth > lngWidth
                lngLen = lngLen - 1
                strText = Left$(strText, lngLen)
                If lngLen > 2 Then Mid$(strText, lngLen - 2) = "."
                lngTextWidth = mpic.TextWidth(strText)
            Loop
            lngLeft = lngLeft + 4
            mpic.CurrentX = lngLeft
            mpic.CurrentY = 3
            mpic.Print Left$(strText, lngLen)
        End If
    ' Panels
    Else
        If plngPanel = 0 Then
            iMin = 1
            iMax = mlngPanels
        Else
            iMin = plngPanel
            iMax = plngPanel
        End If
        For i = iMin To iMax
            With mtypPanel(i)
                ' Clear panel area
                mpic.Line (.Left, 0)-(.Right, mlngLineBottom), vbButtonFace, BF
                ' Draw borders
                lngLeft = .Left + 1
                lngRight = .Right - 1
                mpic.Line (lngLeft, 2)-(lngRight, 2), vbButtonShadow
                mpic.Line (lngLeft, 2)-(lngLeft, mlngLineBottom), vbButtonShadow
                mpic.Line (lngRight, 2)-(lngRight, mlngLineBottom), vbWhite
                If Not mblnMaximized Then mpic.Line (lngLeft, mlngLineBottom)-(lngRight + 1, mlngLineBottom), vbWhite
                If mtypProgress.Panel = i Then
                    ProgressDimensions
                ElseIf mtypSlider.Panel = i Then
                    EraseSlider False
                    DrawSlider
                Else
                    ' Draw text
                    lngWidth = lngRight - lngLeft - 7
                    strText = .Text
                    lngLen = Len(strText)
                    lngTextWidth = mpic.TextWidth(strText)
                    If lngLen <> 0 And lngTextWidth > lngWidth Then
                        lngLen = lngLen + 2
                        strText = Space$(lngLen)
                        Mid$(strText, 1, lngLen - 2) = .Text
                        Mid$(strText, lngLen - 2, 3) = "..."
                        lngTextWidth = mpic.TextWidth(strText)
                    End If
                    Do While lngLen <> 0 And lngTextWidth > lngWidth
                        lngLen = lngLen - 1
                        strText = Left$(strText, lngLen)
                        If lngLen > 2 Then Mid$(strText, lngLen - 2) = "."
                        lngTextWidth = mpic.TextWidth(strText)
                    Loop
                    Select Case .Align
                        Case vbLeftJustify: lngLeft = .Left + 5
                        Case vbCenter: lngLeft = .Left + (.Right - .Left - lngTextWidth) \ 2
                        Case vbRightJustify: lngLeft = .Right - 5 - lngTextWidth
                    End Select
                    mpic.CurrentX = lngLeft
                    mpic.CurrentY = 3
                    mpic.Print Left$(strText, lngLen)
                End If
            End With
        Next
    End If
End Sub
