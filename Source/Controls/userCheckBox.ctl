VERSION 5.00
Begin VB.UserControl userCheckBox 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   372
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4332
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
   ScaleHeight     =   372
   ScaleWidth      =   4332
End
Attribute VB_Name = "userCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Const CheckNormal As String = "X"
Private Const CheckImplicit As String = "x"

Public Enum CheckCharEnum
    chkCheck
    chkX
End Enum

Public Enum CheckPositionEnum
    chkLeft
    chkRight
End Enum

Public Enum CheckStyleEnum
    cseStandard
    cseRed
    cseOrange
    csePurple
    cseBlue
    cseGreen
    cseYellow
    cseColorless
End Enum

Public Event UserChange()
Public Event CodeChange()
Public Event MouseOver()

Private mlngBoxWidth As Long
Private mlngHeight As Long
Private mlngMargin As Long
Private mlngPixelX As Long
Private mlngPixelY As Long

Private menCheckChar As CheckCharEnum
Private menCheckPosition As CheckPositionEnum
Private mstrCaption As String
Private mblnValue As Boolean
Private mblnEnabled As Boolean
Private mblnBold As Boolean
Private mblnImplicit As Boolean
Private mlngFitWidth As Long
Private menStyle As CheckStyleEnum
Private menGroup As ColorGroupEnum

Private mlngBackColor As Long
Private mlngForeColor As Long
Private mlngDimColor As Long
Private mlngFillColor As Long
Private mlngCheckColor As Long
Private mlngCheckDimColor As Long
Private mlngBorderColor As Long
Private mlngBorderInterior As Long

Private mblnHasFocus As Boolean


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    menCheckChar = chkCheck
    mblnValue = True
    mblnImplicit = False
    mstrCaption = "Checkbox"
    menCheckPosition = chkLeft
    mblnEnabled = True
    mblnBold = False
    mlngBorderColor = vbWindowText
    mlngBorderInterior = vbGrayText
    menStyle = cseStandard
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "CheckCharacter", menCheckChar, chkCheck
    PropBag.WriteProperty "Value", mblnValue, True
    PropBag.WriteProperty "Implicit", mblnImplicit, False
    PropBag.WriteProperty "Caption", mstrCaption, "Checkbox"
    PropBag.WriteProperty "CheckPosition", menCheckPosition, chkLeft
    PropBag.WriteProperty "Enabled", mblnEnabled, True
    PropBag.WriteProperty "Bold", mblnBold, False
    PropBag.WriteProperty "Style", menStyle, cseStandard
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    menCheckChar = PropBag.ReadProperty("CheckCharacter", chkCheck)
    mblnValue = PropBag.ReadProperty("Value", True)
    mblnImplicit = PropBag.ReadProperty("Implicit", False)
    mstrCaption = PropBag.ReadProperty("Caption", "Checkbox")
    menCheckPosition = PropBag.ReadProperty("CheckPosition", chkLeft)
    mblnEnabled = PropBag.ReadProperty("Enabled", True)
    mblnBold = PropBag.ReadProperty("Bold", False)
    menStyle = PropBag.ReadProperty("Style", cseStandard)
    Refresh
End Sub


' ************* PROPERTIES *************


Public Property Get CheckChar() As CheckCharEnum
    CheckChar = menCheckChar
End Property

Public Property Let CheckChar(ByVal penCheckChar As CheckCharEnum)
    menCheckChar = penCheckChar
    PropertyChanged "CheckChar"
    ShowValue
End Property


Public Property Get Value() As Boolean
    Value = mblnValue
End Property

Public Property Let Value(ByVal pblnValue As Boolean)
    If mblnValue = pblnValue Then Exit Property
    mblnValue = pblnValue
    PropertyChanged "Value"
    ShowValue
    RaiseEvent CodeChange
End Property


Public Property Get Implicit() As Boolean
    Implicit = mblnImplicit
End Property

Public Property Let Implicit(ByVal pblnImplicit As Boolean)
    If mblnImplicit = pblnImplicit Then Exit Property
    mblnImplicit = pblnImplicit
    PropertyChanged "Implicit"
    ShowValue
    RaiseEvent CodeChange
End Property


Public Property Get Caption() As String
    Caption = mstrCaption
End Property

Public Property Let Caption(ByVal pstrCaption As String)
    mstrCaption = pstrCaption
    PropertyChanged "Caption"
    Refresh
End Property


Public Property Get CheckPosition() As CheckPositionEnum
    CheckPosition = menCheckPosition
End Property

Public Property Let CheckPosition(ByVal penCheckPosition As CheckPositionEnum)
    menCheckPosition = penCheckPosition
    PropertyChanged "Check Position"
    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property

Public Property Let Enabled(ByVal pblnEnabled As Boolean)
    If mblnEnabled = pblnEnabled Then Exit Property
    mblnEnabled = pblnEnabled
    PropertyChanged "Enabled"
    Refresh
End Property

Public Property Get Bold() As Boolean
    Bold = mblnBold
End Property

Public Property Let Bold(ByVal pblnBold As Boolean)
    If mblnBold = pblnBold Then Exit Property
    mblnBold = pblnBold
    PropertyChanged "Bold"
    Refresh
End Property

Public Property Get Style() As CheckStyleEnum
    Style = menStyle
End Property

Public Property Let Style(ByVal CheckStyle As CheckStyleEnum)
    menStyle = CheckStyle
    PropertyChanged "Style"
    RefreshColors menGroup
End Property

Public Property Get FitWidth() As Long
    FitWidth = mlngFitWidth
End Property

Public Property Get BoxWidth() As Long
    BoxWidth = mlngBoxWidth
End Property


' ************* DRAWING *************


Private Sub UserControl_Initialize()
    mlngPixelX = Screen.TwipsPerPixelX
    mlngPixelY = Screen.TwipsPerPixelY
    With UserControl
        .FontBold = True
        mlngBoxWidth = .TextWidth(CheckNormal) + UserControl.ScaleX(7, vbPixels, vbTwips)
        mlngHeight = .TextHeight(CheckNormal)
        mlngMargin = .TextWidth("  ")
        .FontBold = False
        mlngBackColor = vbWindowBackground
        mlngForeColor = vbWindowText
        mlngDimColor = vbGrayText
        mlngFillColor = vbWindowBackground
        mlngCheckColor = vbWindowText
        mlngCheckDimColor = vbGrayText
        mlngBorderColor = vbWindowText
        mlngBorderInterior = vbGrayText
    End With
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Public Sub RefreshColors(Optional penGroup As ColorGroupEnum = cgeWorkspace)
    Dim enTextColor As ColorValueEnum
    
    If penGroup = cgeNavigation Then enTextColor = cveTextLink Else enTextColor = cveText
    mlngBackColor = cfg.GetColor(penGroup, cveBackground)
    mlngForeColor = cfg.GetColor(penGroup, enTextColor)
    mlngDimColor = cfg.GetColor(penGroup, cveTextDim)
    mlngFillColor = cfg.GetColor(cgeControls, cveBackground)
    mlngCheckColor = cfg.GetColor(cgeControls, cveText)
    mlngCheckDimColor = cfg.GetColor(cgeControls, cveTextDim)
    mlngBorderColor = cfg.GetColor(cgeWorkspace, cveBorderExterior)
    mlngBorderInterior = cfg.GetColor(cgeWorkspace, cveBorderInterior)
    menGroup = penGroup
    ApplyStyleColors
    Refresh
'    mlngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
'    mlngForeColor = cfg.GetColor(cgeWorkspace, cveText)
'    mlngDimColor = cfg.GetColor(cgeWorkspace, cveTextDim)
'    mlngFillColor = cfg.GetColor(cgeControls, cveBackground)
'    mlngCheckColor = cfg.GetColor(cgeControls, cveText)
'    mlngCheckDimColor = cfg.GetColor(cgeControls, cveTextDim)
'    mlngBorderColor = cfg.GetColor(cgeWorkspace, cveBorderExterior)
'    mlngBorderInterior = cfg.GetColor(cgeWorkspace, cveBorderInterior)
'    Refresh
End Sub

Private Sub ApplyStyleColors()
    Dim enStyleColor As ColorValueEnum
    
    Select Case menStyle
        Case cseStandard: Exit Sub
        Case cseRed: enStyleColor = cveRed
        Case cseOrange: enStyleColor = cveOrange
        Case csePurple: enStyleColor = cvePurple
        Case cseBlue: enStyleColor = cveBlue
        Case cseGreen: enStyleColor = cveGreen
        Case cseYellow: enStyleColor = cveYellow
        Case cseColorless: enStyleColor = cveLightGray
    End Select
    mlngCheckColor = cfg.GetColor(menGroup, enStyleColor)
    mlngBorderColor = mlngCheckColor
End Sub

Public Sub CustomColors(penGroup As ColorGroupEnum, plngIndex As Long)
    Dim enTextColor As ColorValueEnum
    
    If penGroup = cgeNavigation Then enTextColor = cveTextLink Else enTextColor = cveText
    mlngBackColor = cfg.GetColor(penGroup, cveBackground)
    mlngForeColor = cfg.GetColor(penGroup, enTextColor)
    mlngDimColor = cfg.GetColor(penGroup, cveTextDim)
    mlngFillColor = cfg.GetColor(cgeControls, cveBackground)
    mlngCheckColor = cfg.GetColor(cgeControls, cveText)
    mlngCheckDimColor = cfg.GetColor(cgeControls, cveTextDim)
    mlngBorderColor = cfg.GetColor(cgeWorkspace, cveBorderExterior)
    mlngBorderInterior = cfg.GetColor(cgeWorkspace, cveBorderInterior)
    menGroup = penGroup
    ApplyStyleColors
    Refresh
End Sub

Public Sub Refresh()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngCaptionWidth As Long
    Dim lngColor As Long
    
    With UserControl
        .BackColor = mlngBackColor
        .Cls
        .FontBold = mblnBold
        lngColor = mlngBackColor
        If mblnEnabled Then
            If mblnHasFocus Then
                .ForeColor = vbHighlightText
                lngColor = vbHighlight
            Else
                .ForeColor = mlngForeColor
            End If
        Else
            .ForeColor = mlngDimColor
        End If
        If Len(mstrCaption) Then
            lngCaptionWidth = .TextWidth(mstrCaption)
            If menCheckPosition = chkLeft Then lngLeft = mlngBoxWidth + mlngMargin Else lngLeft = .ScaleWidth - lngCaptionWidth - mlngBoxWidth - mlngMargin
            lngRight = lngLeft + lngCaptionWidth
            lngTop = (.ScaleHeight - .TextHeight(mstrCaption)) \ 2
            lngBottom = lngTop + .TextHeight(mstrCaption)
            UserControl.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngColor, BF
            .CurrentX = lngLeft
            .CurrentY = lngTop
            UserControl.Print mstrCaption
            mlngFitWidth = mlngBoxWidth + mlngMargin + lngCaptionWidth + .TextWidth(" ")
        Else
            mlngFitWidth = mlngBoxWidth + mlngMargin
        End If
    End With
    ShowValue
End Sub

Private Sub ShowValue()
    Dim strCheckChar As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    Dim strCheck As String
    
    ' Draw box
    With UserControl
        If menCheckPosition = chkLeft Then lngLeft = 0 Else lngLeft = .ScaleWidth - mlngBoxWidth - mlngPixelX
        lngTop = (.ScaleHeight - mlngHeight) \ 2
        lngRight = lngLeft + mlngBoxWidth
        lngBottom = lngTop + mlngHeight
        .FillColor = mlngFillColor
        UserControl.Line (lngLeft, lngTop)-(lngRight, lngBottom), mlngBorderColor, B
        UserControl.Line (lngLeft + mlngPixelX, lngTop + mlngPixelY)-(lngRight - mlngPixelX, lngBottom - mlngPixelY), mlngBorderInterior, B
        If mblnValue Then
            If mblnEnabled And Not mblnImplicit Then .ForeColor = mlngCheckColor Else .ForeColor = mlngCheckDimColor
            If mblnImplicit Then strCheck = CheckImplicit Else strCheck = CheckNormal
            .FontBold = True
            .CurrentX = lngLeft + (mlngBoxWidth - .TextWidth(strCheck)) \ 2
            .CurrentY = lngTop
            UserControl.Print strCheck
            .FontBold = False
        End If
    End With
End Sub


' ************* USER INTERACTION *************


Public Sub Click()
    If Not mblnEnabled Then Exit Sub
    If mblnValue And mblnImplicit Then
        mblnImplicit = False
    Else
        mblnValue = Not mblnValue
    End If
    ShowValue
    RaiseEvent UserChange
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Click
End Sub

Private Sub UserControl_DblClick()
    Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseOver
End Sub


' ************* KEYBOARD *************


Private Sub UserControl_GotFocus()
    mblnHasFocus = mblnEnabled
    Refresh
End Sub

Private Sub UserControl_LostFocus()
    mblnHasFocus = False
    Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyReturn
            mblnValue = Not mblnValue
            ShowValue
            RaiseEvent UserChange
    End Select
End Sub


