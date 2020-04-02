VERSION 5.00
Begin VB.UserControl userTab 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1932
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5148
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1932
   ScaleWidth      =   5148
End
Attribute VB_Name = "userTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click(pstrCaption As String)

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Type TabType
    Caption As String
    Left As Long
    Width As Long
    Right As Long
End Type

Private mlngTop As Long

Private mtypTab() As TabType
Private mstrCaptions As String
Private mlngActiveTab As Long
Private mlngTextActiveColor As Long
Private mlngTextInactiveColor As Long
Private mlngTabActiveColor As Long
Private mlngTabInactiveColor As Long
Private mlngBackColor As Long
Private mlngBorderColor As Long
Private mlngHighlightColor As Long
Private mlngShadowColor As Long
Private mlngTabs As Long

Public Sub RefreshColors()
    mlngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    mlngTextActiveColor = cfg.GetColor(cgeControls, cveText)
    mlngTabActiveColor = cfg.GetColor(cgeControls, cveBackground)
    mlngTextInactiveColor = cfg.GetColor(cgeWorkspace, cveText)
    mlngTabInactiveColor = cfg.GetColor(cgeWorkspace, cveBackground)
    DrawTabs
End Sub

' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mstrCaptions = "Tab 1,Tab 2,Tab 3"
    DefineTabs
    mlngActiveTab = 1
    mlngTextActiveColor = vbWindowText
    mlngTextInactiveColor = vbWindowText
    mlngTabActiveColor = vbWindowBackground
    mlngTabInactiveColor = vbButtonFace
    mlngBackColor = vbButtonFace
'    mlngBorderColor = vbBlack
'    mlngHighlightColor = vbWhite
'    mlngShadowColor = vbButtonShadow
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Captions", mstrCaptions, "One,Two,Three"
    PropBag.WriteProperty "ActiveTab", mlngActiveTab, 1
    PropBag.WriteProperty "TextActiveColor", mlngTextActiveColor, vbWindowText
    PropBag.WriteProperty "TextInactiveColor", mlngTextInactiveColor, vbWindowText
    PropBag.WriteProperty "TabActiveColor", mlngTabActiveColor, vbWindowBackground
    PropBag.WriteProperty "TabInactiveColor", mlngTabInactiveColor, vbButtonFace
    PropBag.WriteProperty "BackColor", mlngBackColor, vbButtonFace
'    PropBag.WriteProperty "BorderColor", mlngBorderColor, vbBlack
'    PropBag.WriteProperty "HighlightColor", mlngHighlightColor, vbWhite
'    PropBag.WriteProperty "ShadowColor", mlngShadowColor, vbButtonShadow
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mstrCaptions = PropBag.ReadProperty("Captions", "One,Two,Three")
    DefineTabs
    mlngActiveTab = PropBag.ReadProperty("ActiveTab", 1)
    mlngTextActiveColor = PropBag.ReadProperty("TextActiveColor", vbWindowText)
    mlngTextInactiveColor = PropBag.ReadProperty("TextInactiveColor", vbWindowText)
    mlngTabActiveColor = PropBag.ReadProperty("TabActiveColor", vbWindowBackground)
    mlngTabInactiveColor = PropBag.ReadProperty("TabInactiveColor", vbButtonFace)
    mlngBackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
'    mlngBorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
'    mlngHighlightColor = PropBag.ReadProperty("HighlightColor", vbWhite)
'    mlngShadowColor = PropBag.ReadProperty("ShadowColor", vbButtonShadow)
    DrawTabs
End Sub

Private Sub UserControl_Resize()
    DrawTabs
End Sub

' Set all three changes with only a single redraw
Public Sub BulkChange(pstrCaptions As String, pstrActiveTab As String, plngTabActiveColor As Long)
    mstrCaptions = pstrCaptions
    DefineTabs
    SetActiveTab pstrActiveTab
    mlngTabActiveColor = plngTabActiveColor
    DrawTabs
End Sub

Public Property Get Captions() As String
    Captions = mstrCaptions
End Property

Public Property Let Captions(ByVal pstrCaptions As String)
    mstrCaptions = pstrCaptions
    DefineTabs
    DrawTabs
    PropertyChanged "Captions"
End Property

Public Property Get ActiveTab() As String
    ActiveTab = mtypTab(mlngActiveTab).Caption
End Property

Public Property Let ActiveTab(ByVal pstrActiveTab As String)
    SetActiveTab pstrActiveTab
    DrawTabs
    PropertyChanged "ActiveTab"
    ClickEvent
End Property

' The Click() event may need to redimension the mtypTab() array, so
' it can't be called using an element from mtypTab() as a parameter
' because that locks the array.
Private Sub ClickEvent()
    Dim strCaption As String
    
    strCaption = mtypTab(mlngActiveTab).Caption
    RaiseEvent Click(strCaption)
End Sub

Public Property Get ActiveTabIndex() As Long
    ActiveTabIndex = mlngActiveTab - 1
End Property

Private Sub SetActiveTab(pstrActiveTab As String)
    For mlngActiveTab = mlngTabs To 1 Step -1
        If mtypTab(mlngActiveTab).Caption = pstrActiveTab Then Exit For
    Next
End Sub

Public Sub ClickTab(pstrCaption As String)
    Dim i As Long
    
    For i = 1 To mlngTabs
        If mtypTab(i).Caption = pstrCaption Then
            mlngActiveTab = i
            DrawTabs
            ClickEvent
            Exit Sub
        End If
    Next
End Sub

Public Property Get TextActiveColor() As OLE_COLOR
    TextActiveColor = mlngTextActiveColor
End Property

Public Property Let TextActiveColor(ByVal poleColor As OLE_COLOR)
    mlngTextActiveColor = poleColor
    DrawTabs
    PropertyChanged "TextActiveColor"
End Property

Public Property Get TextInactiveColor() As OLE_COLOR
    TextInactiveColor = mlngTextInactiveColor
End Property

Public Property Let TextInactiveColor(ByVal poleColor As OLE_COLOR)
    mlngTextInactiveColor = poleColor
    DrawTabs
    PropertyChanged "TextInactiveColor"
End Property

Public Property Get TabActiveColor() As OLE_COLOR
    TabActiveColor = mlngTabActiveColor
End Property

Public Property Let TabActiveColor(ByVal poleColor As OLE_COLOR)
    mlngTabActiveColor = poleColor
    DrawTabs
    PropertyChanged "TabActiveColor"
End Property

Public Property Get TabInactiveColor() As OLE_COLOR
    TabInactiveColor = mlngTabInactiveColor
End Property

Public Property Let TabInactiveColor(ByVal poleColor As OLE_COLOR)
    mlngTabInactiveColor = poleColor
    DrawTabs
    PropertyChanged "TabInactiveColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mlngBackColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    mlngBackColor = poleColor
    DrawTabs
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As Long
    BorderColor = mlngBorderColor
End Property

'Public Property Get BorderColor() As OLE_COLOR
'    BorderColor = mlngBorderColor
'End Property
'
'Public Property Let BorderColor(ByVal poleColor As OLE_COLOR)
'    mlngBorderColor = poleColor
'    DrawTabs
'    PropertyChanged "BorderColor"
'End Property
'
'Public Property Get HighlightColor() As OLE_COLOR
'    HighlightColor = mlngHighlightColor
'End Property
'
'Public Property Let HighlightColor(ByVal poleColor As OLE_COLOR)
'    mlngHighlightColor = poleColor
'    DrawTabs
'    PropertyChanged "HighlightColor"
'End Property
'
'Public Property Get ShadowColor() As OLE_COLOR
'    ShadowColor = mlngShadowColor
'End Property
'
'Public Property Let ShadowColor(ByVal poleColor As OLE_COLOR)
'    mlngShadowColor = poleColor
'    DrawTabs
'    PropertyChanged "ShadowColor"
'End Property

Public Property Get Tabs() As Long
    Tabs = mlngTabs
End Property

Public Property Get TabsWidth()
    TabsWidth = mtypTab(mlngTabs).Right + Screen.TwipsPerPixelX
End Property

Public Property Get TabHeight()
    TabHeight = UserControl.ScaleHeight - mlngTop
End Property

Private Sub DrawTabs()
    Dim PixelX As Long
    Dim PixelY As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    Dim lngTab As Long
    Dim i As Long
    
    DefineColors
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    With UserControl
        mlngTop = .ScaleHeight - .TextHeight("Q") - .ScaleY(8, vbPixels, vbTwips)
        lngBottom = .ScaleHeight - PixelY
        .BackColor = mlngBackColor
        .Cls
    End With
    For lngTab = 1 To mlngTabs
        With mtypTab(lngTab)
            ' Background
            If lngTab = mlngActiveTab Then lngColor = mlngTabActiveColor Else lngColor = mlngTabInactiveColor
            UserControl.Line (.Left, mlngTop)-(.Right, lngBottom), lngColor, BF
            ' Black border
            UserControl.Line (.Left, mlngTop + 4 * PixelY)-(.Left, lngBottom + PixelY), mlngBorderColor
            UserControl.Line (.Left, mlngTop + 4 * PixelY)-(.Left + 4 * PixelX, mlngTop), mlngBorderColor
            UserControl.Line (.Left + 4 * PixelY, mlngTop)-(.Right - 4 * PixelX, mlngTop), mlngBorderColor
            UserControl.Line (.Right - 4 * PixelX, mlngTop)-(.Right, mlngTop + 4 * PixelY), mlngBorderColor
            UserControl.Line (.Right, mlngTop + 4 * PixelY)-(.Right, lngBottom + PixelY), mlngBorderColor
            ' Highlight
            UserControl.Line (.Left + PixelX, mlngTop + 4 * PixelY)-(.Left + PixelX, lngBottom + PixelY), mlngHighlightColor
            UserControl.Line (.Left + PixelX, mlngTop + 4 * PixelY)-(.Left + 4 * PixelX, mlngTop + PixelY), mlngHighlightColor
            UserControl.Line (.Left + 4 * PixelX, mlngTop + PixelY)-(.Right - 4 * PixelY, mlngTop + PixelY), mlngHighlightColor
            ' Shadow
            UserControl.Line (.Right - 4 * PixelX, mlngTop + PixelY)-(.Right - PixelX, mlngTop + 4 * PixelY), mlngShadowColor
            UserControl.Line (.Right - PixelX, mlngTop + 4 * PixelY)-(.Right - PixelX, lngBottom + PixelY), mlngShadowColor
            ' Bottom border
            If (lngTab <> mlngActiveTab) Then
                UserControl.Line (.Left, lngBottom)-(.Right + PixelX, lngBottom), mlngBorderColor
            End If
            ' Caption
            If lngTab = mlngActiveTab Then lngColor = mlngTextActiveColor Else lngColor = mlngTextInactiveColor
            UserControl.ForeColor = lngColor
            UserControl.FontBold = (lngTab = mlngActiveTab)
            UserControl.CurrentX = .Left + (.Right - .Left - UserControl.TextWidth(.Caption)) \ 2
            UserControl.CurrentY = mlngTop + 4 * PixelY
            UserControl.Print .Caption
            ' Corners
            For i = 0 To 3
                UserControl.Line (.Left, mlngTop + i * PixelY)-(.Left + (4 - i) * PixelX, mlngTop + i * PixelY), mlngBackColor
                UserControl.Line (.Right, mlngTop + i * PixelY)-(.Right - (4 - i) * PixelX, mlngTop + i * PixelY), mlngBackColor
            Next
        End With
    Next
End Sub

Private Sub DefineTabs()
    Dim strCaption() As String
    Dim lngLeft As Long
    Dim i As Long
    
    strCaption = Split("," & mstrCaptions, ",")
    mlngTabs = UBound(strCaption)
    ReDim mtypTab(mlngTabs)
    UserControl.FontBold = True
'    lngLeft = Screen.TwipsPerPixelX
    For i = 1 To mlngTabs
        With mtypTab(i)
            .Caption = strCaption(i)
            If Len(.Caption) = 0 Then .Caption = "Tab " & i
            .Left = lngLeft
            .Width = UserControl.TextWidth(.Caption & "    ")
            .Right = .Left + .Width
            lngLeft = .Right
        End With
    Next
    UserControl.FontBold = False
End Sub

Private Sub DefineColors()
    If GetBrightest() < 129 Then
        mlngBorderColor = RGB(128, 128, 128)
        mlngHighlightColor = RGB(160, 160, 160)
'        mlngShadowColor = RGB(96, 96, 96)
        mlngShadowColor = RGB(112, 112, 112)
'        mlngTextInactiveColor = vbWhite
    Else
        mlngBorderColor = vbBlack
        mlngHighlightColor = vbWhite
        mlngShadowColor = RGB(160, 160, 160)
'        mlngTextInactiveColor = vbBlack
    End If
'    mlngTabActiveColor = mlngBackColor
'    mlngTabInactiveColor = mlngBackColor
End Sub

Private Function GetBrightest() As Long
    Dim lngColor As Long
    Dim lngHighest As Long
    Dim lngCheck As Long
    
    If mlngBackColor < 0 Then OleTranslateColor mlngBackColor, 0, lngColor Else lngColor = mlngBackColor
    lngHighest = &HFF& And lngColor
    lngCheck = (&HFF00& And lngColor) \ 256
    If lngHighest < lngCheck Then lngHighest = lngCheck
    lngCheck = (&HFF0000 And lngColor) \ 65536
    If lngHighest < lngCheck Then lngHighest = lngCheck
    GetBrightest = lngHighest
End Function
'
'Private Sub ColorToRGB(ByVal plngColor As Long, plngRed As Long, plngGreen As Long, plngBlue As Long)
'    If plngColor < 0 Then plngColor = SystemColorRGB(plngColor)
'    plngRed = &HFF& And plngColor
'    plngGreen = (&HFF00& And plngColor) \ 256
'    plngBlue = (&HFF0000 And plngColor) \ 65536
'End Sub
'
'' Translate system color into its resulting color value
'Private Function SystemColorRGB(ByVal plngSystemColor As Long) As Long
'    Const S_OK = &H0
'    Dim lngReturn As Long
'
'    If OleTranslateColor(plngSystemColor, 0, lngReturn) = S_OK Then SystemColorRGB = lngReturn Else SystemColorRGB = plngSystemColor
'End Function


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Y < mlngTop Then Exit Sub
    For i = 1 To mlngTabs
        If X < mtypTab(i).Right Then Exit For
    Next
    If i = 0 Or i > mlngTabs Then Exit Sub
    If mlngActiveTab <> i Then
        mlngActiveTab = i
        DrawTabs
        ClickEvent
    End If
End Sub

