VERSION 5.00
Begin VB.UserControl userButton 
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
      Begin VB.Menu mnuDropdown 
         Caption         =   "Dropdown"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "RightClick"
      Index           =   1
      Begin VB.Menu mnuContext 
         Caption         =   "ContextMenu"
         Index           =   0
      End
   End
End
Attribute VB_Name = "userButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click(Caption As String)
Public Event DropdownShow(CommaList As String, Show As Boolean)
Public Event DropdownClick(Caption As String)
Public Event ContextMenuShow(CommaList As String, Show As Boolean)
Public Event ContextMenuClick(Caption As String)

Public Enum ButtonAppearanceEnum
    bseRaised
    bseSunken
    bseSolid
    baeCaption
End Enum

Public Enum ButtonStyleEnum
    bseButton
    bseCaption
    bseMenu
End Enum

Private Type UserButtonType
    Style As ButtonStyleEnum
    Appearance As ButtonAppearanceEnum
    Caption As String
    Align As AlignmentConstants
    Enabled As Boolean
    ForeColor As Long
    DimColor As Long
    BackColor As Long
    BorderLight As Long
    BorderDark As Long
    BorderSolid As Long
End Type

Private btn As UserButtonType


' ************* PUBLIC *************


'? Unique to DDO Lite tools
Public Sub RefreshColors()
    Dim enGroup As ColorGroupEnum
    
    If btn.Style = bseCaption Then enGroup = cgeWorkspace Else enGroup = cgeControls
    btn.ForeColor = cfg.GetColor(enGroup, cveText)
    btn.DimColor = cfg.GetColor(enGroup, cveTextDim)
    btn.BackColor = cfg.GetColor(enGroup, cveBackground)
    btn.BorderLight = cfg.GetColor(enGroup, cveBorderExterior)
    btn.BorderDark = cfg.GetColor(enGroup, cveBorderInterior)
    btn.BorderSolid = btn.BorderLight
    Redraw
End Sub

Public Sub Redraw()
    DrawButton btn.Appearance
End Sub

Public Property Get TextLeft()
    Dim strText As String
    
    strText = btn.Caption
    With UserControl
        Select Case btn.Align
            Case vbCenter: TextLeft = (.ScaleWidth - .TextWidth(strText)) \ 2
            Case vbLeftJustify: TextLeft = .TextWidth(" ")
            Case vbRightJustify: TextLeft = .ScaleWidth - .TextWidth(strText & " ")
        End Select
    End With
End Property

Public Property Get TextTop() As Long
    TextTop = (UserControl.ScaleHeight - UserControl.TextHeight(btn.Caption)) \ 2 - PixelY
End Property


' ************* PROPERTIES *************


Public Property Get Caption() As String
    Caption = btn.Caption
End Property

Public Property Let Caption(ByVal pstrCaption As String)
    btn.Caption = pstrCaption
    PropertyChanged "Caption"
    Redraw
End Property

Public Property Get Style() As ButtonStyleEnum
    Style = btn.Style
End Property

Public Property Let Style(ByVal penStyle As ButtonStyleEnum)
    Dim enAppearance As ButtonAppearanceEnum
    
    btn.Style = penStyle
    Select Case btn.Style
        Case bseButton: enAppearance = bseRaised
        Case bseCaption: enAppearance = baeCaption
    End Select
    PropertyChanged "Style"
    If btn.Appearance <> enAppearance Then
        btn.Appearance = enAppearance
        PropertyChanged "Appearance"
    End If
    Redraw
End Property

Public Property Get Appearance() As ButtonAppearanceEnum
    Appearance = btn.Appearance
End Property

Public Property Let Appearance(ByVal penAppearance As ButtonAppearanceEnum)
    btn.Appearance = penAppearance
    PropertyChanged "Appearance"
    Redraw
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = btn.Align
End Property

Public Property Let Alignment(ByVal penAlign As AlignmentConstants)
    btn.Align = penAlign
    PropertyChanged "Alignment"
    Redraw
End Property

Public Property Get Enabled() As Boolean
    Enabled = btn.Enabled
End Property

Public Property Let Enabled(ByVal pblnEnabled As Boolean)
    btn.Enabled = pblnEnabled
    PropertyChanged "Enabled"
    Redraw
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = btn.ForeColor
End Property

Public Property Let ForeColor(ByVal poleColor As OLE_COLOR)
    btn.ForeColor = poleColor
    PropertyChanged "ForeColor"
    Redraw
End Property

Public Property Get DimColor() As OLE_COLOR
    DimColor = btn.DimColor
End Property

Public Property Let DimColor(ByVal poleColor As OLE_COLOR)
    btn.DimColor = poleColor
    PropertyChanged "DimColor"
    Redraw
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = btn.BackColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    btn.BackColor = poleColor
    PropertyChanged "BackColor"
    Redraw
End Property

Public Property Get BorderLight() As OLE_COLOR
    BorderLight = btn.BorderLight
End Property

Public Property Let BorderLight(ByVal poleColor As OLE_COLOR)
    btn.BorderLight = poleColor
    PropertyChanged "BorderLight"
    Redraw
End Property

Public Property Get BorderDark() As OLE_COLOR
    BorderDark = btn.BorderDark
End Property

Public Property Let BorderDark(ByVal poleColor As OLE_COLOR)
    btn.BorderDark = poleColor
    PropertyChanged "BorderDark"
    Redraw
End Property

Public Property Get BorderSolid() As OLE_COLOR
    BorderSolid = btn.BorderSolid
End Property

Public Property Let BorderSolid(ByVal poleColor As OLE_COLOR)
    btn.BorderSolid = poleColor
    PropertyChanged "BorderSolid"
    Redraw
End Property


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    btn.Caption = "Button"
    btn.Style = bseButton
    btn.Appearance = bseRaised
    btn.Align = vbCenter
    btn.Enabled = True
    btn.ForeColor = vbButtonText
    btn.DimColor = vbGrayText
    btn.BackColor = vbButtonFace
    btn.BorderLight = vb3DHighlight
    btn.BorderDark = vb3DDKShadow
    btn.BorderSolid = vb3DLight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", btn.Caption, "Button"
    PropBag.WriteProperty "Style", btn.Style, bseButton
    PropBag.WriteProperty "Appearance", btn.Appearance, bseRaised
    PropBag.WriteProperty "Alignment", btn.Align, vbCenter
    PropBag.WriteProperty "Enabled", btn.Enabled, True
    PropBag.WriteProperty "ForeColor", btn.ForeColor, vbButtonText
    PropBag.WriteProperty "DimColor", btn.DimColor, vbGrayText
    PropBag.WriteProperty "BackColor", btn.BackColor, vbButtonFace
    PropBag.WriteProperty "BorderLight", btn.BorderLight, vb3DHighlight
    PropBag.WriteProperty "BorderDark", btn.BorderDark, vb3DDKShadow
    PropBag.WriteProperty "BorderSolid", btn.BorderSolid, vb3DLight
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    btn.Caption = PropBag.ReadProperty("Caption", "Button")
    btn.Style = PropBag.ReadProperty("Style", bseButton)
    btn.Appearance = PropBag.ReadProperty("Appearance", bseRaised)
    btn.Align = PropBag.ReadProperty("Alignment", vbCenter)
    btn.Enabled = PropBag.ReadProperty("Enabled", True)
    btn.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    btn.DimColor = PropBag.ReadProperty("DimColor", vbGrayText)
    btn.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    btn.BorderLight = PropBag.ReadProperty("BorderLight", vb3DHighlight)
    btn.BorderDark = PropBag.ReadProperty("BorderDark", vb3DDKShadow)
    btn.BorderSolid = PropBag.ReadProperty("BorderSolid", vb3DLight)
    Redraw
End Sub


' ************* DRAWING *************


Private Sub UserControl_Resize()
    DrawButton btn.Appearance
End Sub

Private Sub DrawButton(penAppearance As ButtonAppearanceEnum, Optional pblnRefresh As Boolean = False)
    Dim PixelX As Long
    Dim PixelY As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim strText As String
    
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    With UserControl
        ' Background
        If btn.Enabled Then .ForeColor = btn.ForeColor Else .ForeColor = btn.DimColor
        .BackColor = btn.BackColor
        .Cls
        ' Borders
        Select Case penAppearance
            Case bseRaised
                UserControl.Line (.ScaleWidth - PixelX, 0)-(.ScaleWidth - PixelX, .ScaleHeight + PixelY), btn.BorderDark
                UserControl.Line (0, .ScaleHeight - PixelY)-(.ScaleWidth + PixelX, .ScaleHeight - PixelY), btn.BorderDark
                UserControl.Line (0, 0)-(0, .ScaleHeight + PixelY), btn.BorderLight
                UserControl.Line (0, 0)-(.ScaleWidth + PixelX, 0), btn.BorderLight
            Case bseSunken
                UserControl.Line (0, 0)-(0, .ScaleHeight + PixelY), btn.BorderDark
                UserControl.Line (0, 0)-(.ScaleWidth + PixelX, 0), btn.BorderDark
                UserControl.Line (.ScaleWidth - PixelX, 0)-(.ScaleWidth - PixelX, .ScaleHeight + PixelY), btn.BorderLight
                UserControl.Line (0, .ScaleHeight - PixelY)-(.ScaleWidth + PixelX, .ScaleHeight - PixelY), btn.BorderLight
            Case bseSolid
                UserControl.Line (0, 0)-(.ScaleWidth - PixelX, .ScaleHeight - PixelY), btn.BorderLight, B
        End Select
        lngLeft = TextLeft
        lngTop = TextTop
        If penAppearance = bseSunken Then
            lngLeft = lngLeft + PixelX
            lngTop = lngTop + PixelY
        End If
        .CurrentX = lngLeft
        .CurrentY = lngTop
    End With
    UserControl.Print btn.Caption
    If pblnRefresh Then UserControl.Refresh
End Sub


' ************* MOUSE *************


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If btn.Style <> bseCaption And btn.Enabled = True Then xp.SetMouseCursor mcHand
End Sub

Private Sub UserControl_DblClick()
    If btn.Style <> bseCaption And btn.Enabled = True Then xp.SetMouseCursor mcHand
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If btn.Style = bseCaption Or btn.Enabled = False Then Exit Sub
    xp.SetMouseCursor mcHand
    If Button = vbLeftButton Then
        If btn.Appearance = bseRaised Then DrawButton bseSunken
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strCaption As String
    
    If btn.Style = bseCaption Or btn.Enabled = False Then Exit Sub
    xp.SetMouseCursor mcHand
    Select Case Button
        ' Left button
        Case vbLeftButton
            If btn.Appearance = bseRaised Then DrawButton bseRaised, True
            strCaption = btn.Caption
            If btn.Style = bseMenu Then ShowDropdown Else RaiseEvent Click(strCaption)
        ' Right button
        Case vbRightButton
            ShowContextMenu
    End Select
End Sub


' ************* DROPDOWN MENU *************


Private Sub ShowDropdown()
    Dim blnShow As Boolean
    Dim strRaw As String
    Dim strList() As String
    Dim i As Long
    
    RaiseEvent DropdownShow(strRaw, blnShow)
    If Len(strRaw) = 0 Or blnShow = False Then Exit Sub
    strList = Split(strRaw, ",")
    With UserControl
        For i = 0 To UBound(strList)
            If i > .mnuDropdown.ubound Then Load .mnuDropdown(i)
            With .mnuDropdown(i)
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
        For i = .mnuDropdown.ubound To i + 1 Step -1
            Unload .mnuDropdown(i)
        Next
        PopupMenu .mnuMain(0), , 0, .ScaleHeight
    End With
End Sub

Private Sub mnuDropdown_Click(Index As Integer)
    Dim strCaption As String
    
    strCaption = UserControl.mnuDropdown(Index).Caption
    RaiseEvent DropdownClick(strCaption)
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


