VERSION 5.00
Begin VB.UserControl userDropdown 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   312
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1416
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   312
   ScaleWidth      =   1416
   Begin VB.Menu mnuMain 
      Caption         =   "LeftClick"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuDropdown 
         Caption         =   "Dropdown"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "RightClick"
      Index           =   1
      Begin VB.Menu mnuContext 
         Caption         =   "Context"
         Index           =   0
      End
   End
End
Attribute VB_Name = "userDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SpinWidth As Long = 21 ' Pixels

Public Event ListChange(Index As Long, Caption As String, ItemData As Long)
Public Event ContextMenuShow(CommaList As String, Show As Boolean)
Public Event ContextMenuClick(Caption As String)

Private Type DropdownListType
    Text As String
    ItemData As Long
    BackColor As Long
End Type

Private Type UserDropdownType
    Caption As String
    Align As AlignmentConstants
    List() As DropdownListType
    ListIndex As Long
    ListMax As Long
    Enabled As Boolean
    ForeColor As Long
    DimColor As Long
    BackColor As Long
    BorderColor As Long
    Spinner As Boolean
    SpinLeft As Long ' X coordinate
    SpinRight As Long ' X coordinate
    X As Long ' Mouse coordinate
End Type

Private drp As UserDropdownType


' ************* PUBLIC *************


Public Sub RefreshColors()
    drp.ForeColor = cfg.GetColor(cgeControls, cveText)
    drp.DimColor = cfg.GetColor(cgeControls, cveTextDim)
    drp.BackColor = cfg.GetColor(cgeControls, cveBackground)
    drp.BorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
    DrawDropdown
End Sub

Public Sub Redraw()
    DrawDropdown
End Sub

Public Property Get SpinLeft() As Long
    SpinLeft = drp.SpinLeft
End Property

Public Property Get SpinRight() As Long
    SpinRight = drp.SpinRight
End Property


' ************* DROPDOWN *************


Public Sub ListClear(Optional pblnRedraw As Boolean = True)
    drp.ListIndex = 0
    Erase drp.List
    drp.ListMax = 0
    If pblnRedraw Then Redraw
End Sub

Public Sub AddItem(pstrText As String, plngItemData As Long, Optional plngBackColor As Long = -1)
    drp.ListMax = drp.ListMax + 1
    ReDim Preserve drp.List(1 To drp.ListMax)
    With drp.List(drp.ListMax)
        .Text = pstrText
        .ItemData = plngItemData
        .BackColor = plngBackColor
    End With
End Sub

Public Property Get ListIndex() As Long
    ListIndex = drp.ListIndex
End Property

Public Property Let ListIndex(Index As Long)
    If Index > drp.ListMax Then Exit Property
    drp.ListIndex = Index
    Redraw
End Property

Public Property Get ListMax() As Long
    ListMax = drp.ListMax
End Property

' Set dropdown value based on Text
Public Sub SetText(pstrText As String)
    For drp.ListIndex = 1 To drp.ListMax
        If drp.List(drp.ListIndex).Text = pstrText Then Exit For
    Next
    If drp.ListIndex > drp.ListMax Then drp.ListIndex = 0
    Redraw
End Sub

' Set dropdown value based on ItemData
Public Sub SetData(plngItemData As Long)
    For drp.ListIndex = 1 To drp.ListMax
        If drp.List(drp.ListIndex).ItemData = plngItemData Then Exit For
    Next
    If drp.ListIndex > drp.ListMax Then drp.ListIndex = 0
    Redraw
End Sub


' ************* PROPERTIES *************


Public Property Get Caption() As String
    Caption = drp.Caption
End Property

Public Property Let Caption(ByVal pstrCaption As String)
    drp.Caption = pstrCaption
    PropertyChanged "Caption"
    Redraw
End Property

Public Property Get Spinner() As Boolean
    Spinner = drp.Spinner
End Property

Public Property Let Spinner(ByVal pblnSpinner As Boolean)
    drp.Spinner = pblnSpinner
    PropertyChanged "Spinner"
    Redraw
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = drp.Align
End Property

Public Property Let Alignment(ByVal penAlign As AlignmentConstants)
    drp.Align = penAlign
    PropertyChanged "Alignment"
    Redraw
End Property

Public Property Get Enabled() As Boolean
    Enabled = drp.Enabled
End Property

Public Property Let Enabled(ByVal pblnEnabled As Boolean)
    drp.Enabled = pblnEnabled
    PropertyChanged "Enabled"
    Redraw
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = drp.ForeColor
End Property

Public Property Let ForeColor(ByVal poleColor As OLE_COLOR)
    drp.ForeColor = poleColor
    PropertyChanged "ForeColor"
    Redraw
End Property

Public Property Get DimColor() As OLE_COLOR
    DimColor = drp.DimColor
End Property

Public Property Let DimColor(ByVal poleColor As OLE_COLOR)
    drp.DimColor = poleColor
    PropertyChanged "DimColor"
    Redraw
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = drp.BackColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    drp.BackColor = poleColor
    PropertyChanged "BackColor"
    Redraw
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = drp.BorderColor
End Property

Public Property Let BorderColor(ByVal poleColor As OLE_COLOR)
    drp.BorderColor = poleColor
    PropertyChanged "BorderColor"
    Redraw
End Property


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    drp.Caption = "Dropdown"
    drp.Spinner = True
    drp.Align = vbCenter
    drp.Enabled = True
    drp.ForeColor = vbButtonText
    drp.DimColor = vbGrayText
    drp.BackColor = vbButtonFace
    drp.BorderColor = vbBlack
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", drp.Caption, "Dropdown"
    PropBag.WriteProperty "Spinner", drp.Spinner, True
    PropBag.WriteProperty "Alignment", drp.Align, vbCenter
    PropBag.WriteProperty "Enabled", drp.Enabled, True
    PropBag.WriteProperty "ForeColor", drp.ForeColor, vbButtonText
    PropBag.WriteProperty "DimColor", drp.DimColor, vbGrayText
    PropBag.WriteProperty "BackColor", drp.BackColor, vbButtonFace
    PropBag.WriteProperty "BorderColor", drp.BorderColor, vbBlack
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    drp.Caption = PropBag.ReadProperty("Caption", "Dropdown")
    drp.Spinner = PropBag.ReadProperty("Spinner", True)
    drp.Align = PropBag.ReadProperty("Alignment", vbCenter)
    drp.Enabled = PropBag.ReadProperty("Enabled", True)
    drp.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    drp.DimColor = PropBag.ReadProperty("DimColor", vbGrayText)
    drp.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    drp.BorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
    Redraw
End Sub


' ************* DRAWING *************


Private Sub UserControl_Resize()
    DrawDropdown
End Sub

Private Sub DrawDropdown(Optional pblnRefresh As Boolean = False)
    Dim PixelX As Long
    Dim PixelY As Long
    Dim strText As String
    Dim lngColor As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngRight As Long
    Dim lngMiddle As Long
    Dim i As Long
    
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    With UserControl
        ' Data
        If drp.ListIndex > 0 And drp.ListIndex <= drp.ListMax Then
            strText = drp.List(drp.ListIndex).Text
            lngColor = drp.List(drp.ListIndex).BackColor
            If lngColor = -1 Then lngColor = drp.BackColor
        Else
            strText = vbNullString
            lngColor = drp.BackColor
        End If
        ' Clear
        If drp.Enabled Then .ForeColor = drp.ForeColor Else .ForeColor = drp.DimColor
        .BackColor = lngColor
        .Cls
        UserControl.Line (0, 0)-(.ScaleWidth - PixelX, .ScaleHeight - PixelY), drp.BorderColor, B
        ' Coordinates
        lngTop = (.ScaleHeight - .TextHeight(strText)) \ 2 - PixelY
        drp.SpinLeft = .ScaleX(SpinWidth - 1, vbPixels, vbTwips)
        drp.SpinRight = .ScaleWidth - drp.SpinLeft - PixelX
        If drp.Spinner Then
            lngLeft = drp.SpinLeft
            lngRight = drp.SpinRight
        Else
            lngLeft = 0
            lngRight = .ScaleWidth
        End If
        lngWidth = lngRight - lngLeft
        ' Text
        Select Case drp.Align
            Case vbCenter: lngLeft = lngLeft + (lngWidth - .TextWidth(strText)) \ 2
            Case vbLeftJustify: lngLeft = lngLeft + .TextWidth(" ")
            Case vbRightJustify: lngLeft = lngRight - .TextWidth(strText & " ")
        End Select
        .CurrentX = lngLeft
        .CurrentY = lngTop
        UserControl.Print strText
        ' Spinner
        If drp.Spinner Then
            lngHeight = .ScaleY(.ScaleHeight, vbTwips, vbPixels)
            lngMiddle = lngHeight \ 2
            UserControl.Line (drp.SpinLeft, PixelX)-(drp.SpinLeft, .ScaleHeight - PixelY), drp.BorderColor
            UserControl.Line (drp.SpinRight, PixelX)-(drp.SpinRight, .ScaleHeight - PixelY), drp.BorderColor
            lngTop = .ScaleHeight \ 2
            ' Draw arrows
            If lngHeight Mod 2 = 1 Then ' Odd height
                DrawLine 0, 5, lngMiddle, 9, drp.ForeColor
                DrawLine drp.SpinRight, 6, lngMiddle, 9, drp.ForeColor
                For i = 1 To 4
                    ' Left
                    DrawLine 0, 5 + i * 2, lngMiddle - i, 9 - i * 2, drp.ForeColor
                    DrawLine 0, 5 + i * 2, lngMiddle + i, 9 - i * 2, drp.ForeColor
                    ' Right
                    DrawLine drp.SpinRight, 6, lngMiddle - i, 9 - i * 2, drp.ForeColor
                    DrawLine drp.SpinRight, 6, lngMiddle + i, 9 - i * 2, drp.ForeColor
                Next
            Else ' Even height
                For i = 0 To 3
                    ' Left
                    DrawLine 0, 6 + i * 2, lngMiddle - i - 1, 8 - (i * 2), drp.ForeColor
                    DrawLine 0, 6 + i * 2, lngMiddle + i, 8 - (i * 2), drp.ForeColor
                    ' Right
                    DrawLine drp.SpinRight, 6, lngMiddle - i - 1, 8 - i * 2, drp.ForeColor
                    DrawLine drp.SpinRight, 6, lngMiddle + i, 8 - i * 2, drp.ForeColor
                Next
            End If
        End If
    End With
    If pblnRefresh Then UserControl.Refresh
End Sub

Private Sub DrawLine(plngOffsetX As Long, plngLeft As Long, plngTop As Long, plngWidth As Long, plngColor As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    
    With UserControl
        lngLeft = .ScaleX(plngLeft, vbPixels, vbTwips) + plngOffsetX
        lngTop = .ScaleY(plngTop, vbPixels, vbTwips)
        lngRight = .ScaleX(plngLeft + plngWidth + 1, vbPixels, vbTwips) + plngOffsetX
    End With
    UserControl.Line (lngLeft, lngTop)-(lngRight, lngTop), plngColor
End Sub


' ************* MOUSE *************


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not drp.Enabled Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub UserControl_DblClick()
    If Not drp.Enabled Then Exit Sub
    xp.SetMouseCursor mcHand
    If drp.Spinner Then SpinnerClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not drp.Enabled Then Exit Sub
    xp.SetMouseCursor mcHand
    drp.X = X
    If Button = vbLeftButton Then
        If drp.Spinner Then SpinnerClick Else ShowDropdown
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Long
    Dim strCaption As String
    Dim lngItemData As Long
    
    If Not drp.Enabled Then Exit Sub
    xp.SetMouseCursor mcHand
    drp.X = X
    Select Case Button
        Case vbLeftButton
            If drp.ListIndex < 1 Or drp.ListIndex > drp.ListMax Then
                strCaption = drp.Caption
            Else
                lngIndex = drp.ListIndex
                With drp.List(drp.ListIndex)
                    strCaption = .Text
                    lngItemData = .ItemData
                End With
            End If
            RaiseEvent ListChange(lngIndex, strCaption, lngItemData)
        Case vbRightButton
            ShowContextMenu
    End Select
End Sub

Private Sub SpinnerClick()
    Dim lngIncrement As Long
    Dim strList() As String
    Dim lngNext As Long
    Dim i As Long
    
    Select Case drp.X
        Case Is <= drp.SpinLeft: lngIncrement = -1
        Case Is >= drp.SpinRight: lngIncrement = 1
        Case Else
            ShowDropdown
            Exit Sub
    End Select
    lngNext = drp.ListIndex + lngIncrement
    If lngNext < 1 Then
        lngNext = drp.ListMax
    ElseIf lngNext > drp.ListMax Then
        lngNext = 1
    End If
    SelectListIndex lngNext
End Sub

Private Sub SelectListIndex(plngIndex As Long)
    Dim strText As String
    Dim lngItemData As Long
    
    If plngIndex > 0 And plngIndex <= drp.ListMax Then
        drp.ListIndex = plngIndex
        With drp.List(plngIndex)
            strText = .Text
            lngItemData = .ItemData
        End With
    End If
    Redraw
    RaiseEvent ListChange(plngIndex, strText, lngItemData)
End Sub

Private Sub ShowDropdown()
    Dim i As Long
    
    If drp.ListMax = 0 Then Exit Sub
    With UserControl
        For i = 1 To drp.ListMax
            If i > .mnuDropdown.ubound Then Load .mnuDropdown(i)
            With .mnuDropdown(i)
                .Caption = drp.List(i).Text
                .Tag = drp.List(i).ItemData
                .Enabled = True
                .Checked = (i = drp.ListIndex)
                .Visible = True
            End With
        Next
        .mnuDropdown(0).Visible = False
        For i = .mnuDropdown.ubound To i + 1 Step -1
            Unload .mnuDropdown(i)
        Next
        PopupMenu .mnuMain(0), , 0, .ScaleHeight
    End With
End Sub

Private Sub mnuDropdown_Click(Index As Integer)
    SelectListIndex CLng(Index)
End Sub


' ************* CONTEXT MENU *************


Private Sub ShowContextMenu()
    Dim blnShow As Boolean
    Dim strRaw As String
    Dim strList() As String
    Dim i As Long
    
    RaiseEvent ContextMenuShow(strRaw, blnShow)
    If Len(strRaw) = 0 Or blnShow = False Then Exit Sub
    strList = Split(strRaw, ",")
    With UserControl
        For i = 0 To UBound(strList)
            If i > .mnuContext.ubound Then Load .mnuContext(i)
            With .mnuContext(i)
                .Checked = False
                .Caption = Trim$(strList(i))
                If Left$(.Caption, 1) = "+" Then
                    .Caption = Mid$(.Caption, 2)
                    If .Caption <> "-" Then .Checked = True
                End If
                .Enabled = True
                .Visible = True
            End With
        Next
        For i = .mnuContext.ubound To i + 1 Step -1
            Unload .mnuContext(i)
        Next
        PopupMenu .mnuMain(1)
    End With
End Sub

Private Sub mnuContext_Click(Index As Integer)
    Dim strCaption As String
    
    strCaption = UserControl.mnuContext(Index).Caption
    RaiseEvent ContextMenuClick(strCaption)
End Sub

