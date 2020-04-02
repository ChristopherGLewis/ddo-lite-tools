VERSION 5.00
Begin VB.UserControl userSpinner 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   636
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2208
   FillColor       =   &H80000005&
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
   ScaleHeight     =   636
   ScaleWidth      =   2208
   Begin VB.Timer tmrTyping 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   780
      Top             =   240
   End
   Begin VB.Timer tmrTooltip 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   60
      Top             =   240
   End
   Begin VB.Timer tmrRepeat 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1200
      Top             =   0
   End
   Begin VB.Label lblTooltip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Large "
      ForeColor       =   &H80000017&
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   636
   End
End
Attribute VB_Name = "userSpinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Public Event Change()
Public Event RequestChange(OldValue As Long, NewValue As Long, Cancel As Boolean)

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long

Public Enum PositionEnum
    peStandAlone
    peTop
    peMiddle
    peBottom
End Enum

Private mlngLeft As Long
Private mlngRight As Long
Private mlngOffsetX As Long
Private mlngOffsetY As Long
Private mlngPixelX As Long
Private mlngPixelY As Long

Private mlngMin As Long
Private mlngMax As Long
Private mlngValue As Long
Private mlngStepSmall As Long
Private mlngStepLarge As Long
Private mblnAllowJump As Boolean
Private mlngBorderColor As Long
Private mlngBorderInterior As Long
Private menPosition As PositionEnum
Private mblnEnabled As Boolean
Private mlngDisabledColor As Long
Private mlngForeColor As Long

Private mblnAppearance3D As Boolean
Private mblnShowZero As Boolean

Private mlngIncrement As Long
Private mlngWheelIncrement As Long

Private mblnAllowTyping As Boolean
Private mblnHasFocus As Boolean
Private mblnTyping As Boolean


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mblnAppearance3D = False
    mlngMin = 1
    mlngMax = 10
    mlngValue = 10
    mlngStepSmall = 1
    mlngStepLarge = 5
    mblnAllowJump = True
    mblnAllowTyping = True
    mlngForeColor = vbWindowText
    UserControl.BackColor = vbWindowBackground
    mlngBorderColor = vbBlack
    mlngBorderInterior = vbGrayText
    menPosition = peStandAlone
    mblnEnabled = True
    mlngDisabledColor = vbGrayText
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Appearance3D", mblnAppearance3D, False
    PropBag.WriteProperty "Min", mlngMin, 1
    PropBag.WriteProperty "Max", mlngMax, 10
    PropBag.WriteProperty "Value", mlngValue, 10
    PropBag.WriteProperty "StepSmall", mlngStepSmall, 1
    PropBag.WriteProperty "StepLarge", mlngStepLarge, 5
    PropBag.WriteProperty "AllowJump", mblnAllowJump, True
    PropBag.WriteProperty "AllowTyping", mblnAllowTyping, True
    PropBag.WriteProperty "ShowZero", mblnShowZero, True
    PropBag.WriteProperty "ForeColor", mlngForeColor
    PropBag.WriteProperty "BackColor", UserControl.FillColor
    PropBag.WriteProperty "BorderColor", mlngBorderColor
    PropBag.WriteProperty "BorderInterior", mlngBorderInterior
    PropBag.WriteProperty "Position", menPosition
    PropBag.WriteProperty "Enabled", mblnEnabled
    PropBag.WriteProperty "DisabledColor", mlngDisabledColor
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mblnAppearance3D = PropBag.ReadProperty("Appearance3D", False)
    mlngMin = PropBag.ReadProperty("Min", 1)
    mlngMax = PropBag.ReadProperty("Max", 10)
    mlngValue = PropBag.ReadProperty("Value", 10)
    mlngStepSmall = PropBag.ReadProperty("StepSmall", 1)
    mlngStepLarge = PropBag.ReadProperty("StepLarge", 5)
    mblnAllowJump = PropBag.ReadProperty("AllowJump", True)
    mblnAllowTyping = PropBag.ReadProperty("AllowTyping", True)
    mblnShowZero = PropBag.ReadProperty("ShowZero", True)
    mlngForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    UserControl.FillColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    mlngBorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
    mlngBorderInterior = PropBag.ReadProperty("BorderInterior", vbGrayText)
    menPosition = PropBag.ReadProperty("Position", peStandAlone)
    mblnEnabled = PropBag.ReadProperty("Enabled", True)
    mlngDisabledColor = PropBag.ReadProperty("DisabledColor", vbGrayText)
    mlngWheelIncrement = mlngStepSmall
    DrawBorders
    ShowValue
End Sub


' ************* PROPERTIES *************


Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property

Public Property Let Enabled(ByVal pblnEnabled As Boolean)
    mblnEnabled = pblnEnabled
    PropertyChanged "Enabled"
    UserControl.Enabled = mblnEnabled
    DrawBorders
    ShowValue
End Property

' Check out the fancy property type, complete with dropdowns in the Properties window of Design view
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mlngForeColor
End Property

Public Property Let ForeColor(ByVal poleColor As OLE_COLOR)
    mlngForeColor = poleColor
    PropertyChanged "ForeColor"
    DrawBorders
    ShowValue
End Property

Public Property Get DisabledColor() As OLE_COLOR
    DisabledColor = mlngDisabledColor
End Property

Public Property Let DisabledColor(ByVal poleColor As OLE_COLOR)
    mlngDisabledColor = poleColor
    PropertyChanged "DisabledColor"
    DrawBorders
    ShowValue
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.FillColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    UserControl.BackColor = poleColor
    UserControl.FillColor = poleColor
    PropertyChanged "BackColor"
    DrawBorders
    ShowValue
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mlngBorderColor
End Property

Public Property Let BorderColor(ByVal poleColor As OLE_COLOR)
    mlngBorderColor = poleColor
    PropertyChanged "BorderColor"
    DrawBorders
    ShowValue
End Property

Public Property Get BorderInterior() As OLE_COLOR
    BorderInterior = mlngBorderInterior
End Property

Public Property Let BorderInterior(ByVal poleInterior As OLE_COLOR)
    mlngBorderInterior = poleInterior
    PropertyChanged "BorderInterior"
    DrawBorders
    ShowValue
End Property

Public Property Get Position() As PositionEnum
    Position = menPosition
End Property

Public Property Let Position(ByVal penPosition As PositionEnum)
    menPosition = penPosition
    PropertyChanged "Position"
    DrawBorders
    ShowValue
End Property

Public Property Get Appearance3D() As Boolean
    Appearance3D = mblnAppearance3D
End Property

Public Property Let Appearance3D(ByVal pblnAppearance3D As Boolean)
    mblnAppearance3D = pblnAppearance3D
    PropertyChanged "Appearance3D"
    DrawBorders
    ShowValue
End Property

Public Property Get AllowJump() As Boolean
    AllowJump = mblnAllowJump
End Property

Public Property Let AllowJump(ByVal pblnAllowJump As Boolean)
    mblnAllowJump = pblnAllowJump
    PropertyChanged "AllowJump"
End Property

Public Property Get AllowTyping() As Boolean
    AllowTyping = mblnAllowTyping
End Property

Public Property Let AllowTyping(ByVal pblnAllowTyping As Boolean)
    mblnAllowTyping = pblnAllowTyping
    PropertyChanged "AllowTyping"
End Property

Public Property Get ShowZero() As Boolean
    ShowZero = mblnShowZero
End Property

Public Property Let ShowZero(ByVal pblnShowZero As Boolean)
    mblnShowZero = pblnShowZero
    PropertyChanged "ShowZero"
    If mlngValue = 0 Then ShowValue
End Property

Public Property Get Min() As Long
    Min = mlngMin
End Property

Public Property Let Min(ByVal plngMin As Long)
    mlngMin = plngMin
    PropertyChanged "Min"
End Property

Public Property Get Max() As Long
    Max = mlngMax
End Property

Public Property Let Max(ByVal plngMax As Long)
    mlngMax = plngMax
    PropertyChanged "Max"
End Property

Public Property Get Value() As Long
    Value = mlngValue
End Property

Public Property Let Value(ByVal plngValue As Long)
    Dim lngValue As Long
    
    If plngValue < mlngMin Then
        lngValue = mlngMin
    ElseIf plngValue > mlngMax Then
        lngValue = mlngMax
    Else
        lngValue = plngValue
    End If
    If mlngValue <> lngValue Then
        If mblnAllowJump Then
            mlngValue = lngValue
            ShowValue
            RaiseEvent Change
            PropertyChanged "Value"
        Else
            JumpInSteps lngValue
        End If
    End If
End Property

Public Property Get Override() As Long
    Override = mlngValue
End Property

Public Property Let Override(ByVal plngValue As Long)
    Dim lngValue As Long
    
    If plngValue < mlngMin Then
        lngValue = mlngMin
    ElseIf plngValue > mlngMax Then
        lngValue = mlngMax
    Else
        lngValue = plngValue
    End If
    If mlngValue <> lngValue Then
        mlngValue = lngValue
        ShowValue
        RaiseEvent Change
        PropertyChanged "Value"
    End If
End Property

Private Sub JumpInSteps(plngValue As Long)
    Dim blnCancel As Boolean
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngStep As Long
    Dim lngNew As Long
    Dim i As Long
    
    If plngValue > mlngValue Then lngStep = 1 Else lngStep = -1
    lngFirst = mlngValue + lngStep
    lngLast = plngValue
    For lngNew = lngFirst To lngLast Step lngStep
        RaiseEvent RequestChange(mlngValue, lngNew, blnCancel)
        If blnCancel Then Exit For
        mlngValue = lngNew
    Next
    ShowValue
    RaiseEvent Change
End Sub

Public Sub WheelScroll(ByVal plngValue As Long)
    If Not mblnEnabled Then Exit Sub
    If plngValue < 0 Then Me.Value = Me.Value - mlngWheelIncrement Else Me.Value = Me.Value + mlngWheelIncrement
End Sub

Public Property Get StepSmall() As Long
    StepSmall = mlngStepSmall
End Property

Public Property Let StepSmall(ByVal plngStepSmall As Long)
    mlngStepSmall = plngStepSmall
    PropertyChanged "StepSmall"
End Property

Public Property Get StepLarge() As Long
    StepLarge = mlngStepLarge
End Property

Public Property Let StepLarge(ByVal plngStepLarge As Long)
    mlngStepLarge = plngStepLarge
    PropertyChanged "StepLarge"
End Property


' ************* DRAWING *************


Private Sub UserControl_Resize()
    With UserControl
        .Cls
        mlngPixelX = Screen.TwipsPerPixelX
        mlngPixelY = Screen.TwipsPerPixelY
        mlngLeft = .TextWidth("<") * 1.75
        mlngRight = UserControl.ScaleWidth - mlngLeft
        mlngOffsetX = (mlngLeft - .TextWidth("<")) \ 2
        mlngOffsetY = (.ScaleHeight - .TextHeight("<")) \ 2
        mlngOffsetY = mlngOffsetY - mlngPixelY
    End With
    DrawBorders
    ShowValue
End Sub

Private Sub DrawBorders()
    Dim lngBorderColor As Long
    Dim lngColor As Long
    
    If mblnEnabled Then
        lngBorderColor = mlngBorderColor
        UserControl.ForeColor = mlngForeColor
    Else
        lngBorderColor = mlngBorderInterior
        UserControl.ForeColor = mlngDisabledColor
    End If
    If mblnAppearance3D Then
        UserControl.Cls
        ' Top
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), vb3DShadow
        UserControl.Line (mlngPixelX, mlngPixelY)-(UserControl.ScaleWidth - mlngPixelX, mlngPixelY), vb3DDKShadow
        ' Left
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), vb3DShadow
        UserControl.Line (mlngPixelX, mlngPixelY)-(mlngPixelX, UserControl.ScaleHeight - mlngPixelY), vb3DDKShadow
        ' Right
        UserControl.Line (UserControl.ScaleWidth - mlngPixelX, 0)-(UserControl.ScaleWidth - mlngPixelX, UserControl.ScaleHeight + mlngPixelY), vb3DHighlight
        UserControl.Line (UserControl.ScaleWidth - mlngPixelX * 2, mlngPixelY)-(UserControl.ScaleWidth - mlngPixelX * 2, UserControl.ScaleHeight), vb3DLight
        ' Bottom
        UserControl.Line (0, UserControl.ScaleHeight - mlngPixelY)-(UserControl.ScaleWidth, UserControl.ScaleHeight - mlngPixelY), vb3DHighlight
        UserControl.Line (mlngPixelX, UserControl.ScaleHeight - mlngPixelY * 2)-(UserControl.ScaleWidth - mlngPixelX * 2, UserControl.ScaleHeight - mlngPixelY * 2), vb3DLight
        ' Separators
        UserControl.Line (mlngLeft, mlngPixelY * 2)-(mlngLeft, UserControl.ScaleHeight - mlngPixelY * 2), UserControl.ForeColor
        UserControl.Line (mlngRight, mlngPixelY * 2)-(mlngRight, UserControl.ScaleHeight - mlngPixelY * 2), UserControl.ForeColor
    Else
        ' Top
        If menPosition = peStandAlone Or menPosition = peTop Then lngColor = lngBorderColor Else lngColor = mlngBorderInterior
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), lngColor
        ' Left
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), lngBorderColor
        ' Right
        UserControl.Line (UserControl.ScaleWidth - mlngPixelX, 0)-(UserControl.ScaleWidth - mlngPixelX, UserControl.ScaleHeight + mlngPixelY), lngBorderColor
        ' Bottom
        If menPosition = peStandAlone Or menPosition = peBottom Then lngColor = lngBorderColor Else lngColor = mlngBorderInterior
        UserControl.Line (0, UserControl.ScaleHeight - mlngPixelY)-(UserControl.ScaleWidth, UserControl.ScaleHeight - mlngPixelY), lngColor
        ' Separators
        UserControl.Line (mlngLeft, mlngPixelY)-(mlngLeft, UserControl.ScaleHeight - mlngPixelY), mlngBorderInterior
        UserControl.Line (mlngRight, mlngPixelY)-(mlngRight, UserControl.ScaleHeight - mlngPixelY), mlngBorderInterior
    End If
    UserControl.CurrentX = mlngOffsetX
    UserControl.CurrentY = mlngOffsetY
    UserControl.Print "<"
    UserControl.CurrentX = mlngRight + mlngOffsetX
    UserControl.CurrentY = mlngOffsetY
    UserControl.Print ">"
End Sub

Private Sub ShowValue()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    
    lngColor = UserControl.BackColor
    If mblnEnabled Then
        If mblnHasFocus Then
            UserControl.ForeColor = vbHighlightText
            lngColor = vbHighlight
        Else
            UserControl.ForeColor = mlngForeColor
        End If
    Else
        UserControl.ForeColor = mlngDisabledColor
    End If
    lngLeft = mlngLeft + mlngPixelX
    lngRight = mlngRight - mlngPixelX
    If mblnAppearance3D Then
        lngTop = mlngPixelY * 2
        lngBottom = UserControl.ScaleHeight - mlngPixelY * 3
    Else
        lngTop = mlngPixelY
        lngBottom = UserControl.ScaleHeight - mlngPixelY * 2
    End If
    UserControl.Line (lngLeft, lngTop)-(lngRight, lngBottom), UserControl.BackColor, BF
    UserControl.Line (lngLeft + mlngPixelX, lngTop + mlngPixelY)-(lngRight - mlngPixelX, lngBottom - mlngPixelY), lngColor, BF
'    UserControl.Line (mlngLeft + mlngPixelX * 2, mlngPixelY * 2)-(mlngRight - mlngPixelX * 4, UserControl.ScaleHeight - mlngPixelY * 4), lngColor, BF
    If mlngValue <> 0 Or mblnShowZero Then
        UserControl.CurrentX = mlngLeft + (mlngRight - mlngLeft - UserControl.TextWidth(mlngValue)) \ 2
        UserControl.CurrentY = mlngOffsetY
        UserControl.Print CStr(mlngValue)
    End If
End Sub


' ************* USER INTERACTION *************


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X < mlngLeft Or X > mlngRight Then SetCursor LoadCursor(0, 32649&)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngNew As Long
    
    mlngIncrement = 0
    If X < mlngLeft Or X > mlngRight Then
        SetCursor LoadCursor(0, 32649&)
    Else
        If Button <> vbLeftButton Then ToggleStep
        Exit Sub
    End If
    If Button = vbMiddleButton Then
        If X < mlngLeft Then
            mlngIncrement = mlngMin - mlngValue
        ElseIf X > mlngRight Then
            mlngIncrement = mlngMax - mlngValue
        End If
        MouseDown
        Exit Sub
    End If
    If X < mlngLeft Then
        Select Case Button
            Case vbLeftButton: mlngIncrement = -mlngStepSmall
            Case vbRightButton: mlngIncrement = -mlngStepLarge
        End Select
    ElseIf X > mlngRight Then
        Select Case Button
            Case vbLeftButton: mlngIncrement = mlngStepSmall
            Case vbRightButton: mlngIncrement = mlngStepLarge
        End Select
    End If
    MouseDown
    UserControl.tmrStart.Enabled = True
End Sub

Private Sub ToggleStep()
    If mlngWheelIncrement = mlngStepSmall Then WheelStep mlngStepLarge Else WheelStep mlngStepSmall
End Sub

Private Sub WheelStep(plngStep As Long)
    mlngWheelIncrement = plngStep
    With UserControl
        If mlngWheelIncrement = mlngStepLarge Then .lblTooltip.Caption = " Large " Else .lblTooltip.Caption = " Small "
        .lblTooltip.Move (.ScaleWidth - .lblTooltip.Width) \ 2, (.ScaleHeight - .lblTooltip.Height) \ 2
        .lblTooltip.Visible = True
        .tmrTooltip.Enabled = True
    End With
End Sub

Private Sub tmrTooltip_Timer()
    UserControl.tmrTooltip.Enabled = False
    UserControl.lblTooltip.Visible = False
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngIncrement <> 0 Then SetCursor LoadCursor(0, 32649&)
    UserControl.tmrStart.Enabled = False
    UserControl.tmrRepeat.Enabled = False
End Sub

Private Sub tmrStart_Timer()
    UserControl.tmrStart.Enabled = False
    UserControl.tmrRepeat.Enabled = True
End Sub

Private Sub tmrRepeat_Timer()
    MouseDown
End Sub

Public Sub MouseDown()
    Increment mlngIncrement
End Sub

Private Sub Increment(plngIncrement As Long)
    Dim lngNew As Long
    Dim blnCancel As Boolean
    
    lngNew = mlngValue + plngIncrement
    If Abs(plngIncrement) > mlngStepSmall And Not mblnAllowJump Then
        JumpInSteps mlngValue + plngIncrement
    Else
        If lngNew < mlngMin Then lngNew = mlngMin
        If lngNew > mlngMax Then lngNew = mlngMax
        If mlngValue = lngNew Then Exit Sub
        RaiseEvent RequestChange(mlngValue, lngNew, blnCancel)
        If blnCancel Then Exit Sub
        mlngValue = lngNew
        ShowValue
        RaiseEvent Change
    End If
End Sub

Private Sub UserControl_DblClick()
    If mlngIncrement <> 0 Then
        SetCursor LoadCursor(0, 32649&)
        MouseDown
    End If
End Sub


' ************* KEYBOARD *************


Private Sub UserControl_GotFocus()
    mblnHasFocus = mblnEnabled
    ShowValue
End Sub

Private Sub UserControl_LostFocus()
    mblnHasFocus = False
    ShowValue
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngDigit As Long
    Dim lngIncrement As Long
    
    lngDigit = -1
    Select Case KeyCode
        Case vbKey0 To vbKey9: lngDigit = KeyCode - vbKey0
        Case vbKeyNumpad0 To vbKeyNumpad9: lngDigit = KeyCode - vbKeyNumpad0
        Case vbKeyLeft, vbKeySubtract: lngIncrement = -mlngStepSmall
        Case vbKeyRight, vbKeyAdd: lngIncrement = mlngStepSmall
        Case vbKeyUp, vbKeyPageUp: lngIncrement = mlngStepLarge
        Case vbKeyDown, vbKeyPageDown: lngIncrement = -mlngStepLarge
        Case vbKeyDelete, vbKeyEnd: lngIncrement = mlngMin - mlngValue
        Case vbKeyInsert, vbKeyHome: lngIncrement = mlngMax - mlngValue
        Case vbKeySpace: ToggleStep
        Case vbKeyReturn
    End Select
    UserControl.tmrTyping.Enabled = False
    If lngDigit = -1 Then
        mblnTyping = False
        If lngIncrement Then Increment lngIncrement
    ElseIf mblnAllowTyping Then
        If mblnTyping Then
            Me.Value = Val(mlngValue & lngDigit)
        Else
            mblnTyping = (lngDigit <> 0 And mlngMax > 9)
            Me.Value = lngDigit
        End If
        If mblnTyping Then UserControl.tmrTyping.Enabled = True
    End If
End Sub

Private Sub tmrTyping_Timer()
    UserControl.tmrTyping.Enabled = False
    mblnTyping = False
End Sub

