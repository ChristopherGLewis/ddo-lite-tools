VERSION 5.00
Begin VB.UserControl userToolbar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   408
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9072
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   408
   ScaleWidth      =   9072
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   0
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   972
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tools"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   1
      Left            =   7032
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   972
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Play"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   2
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
   Begin Compendium.userTab usrTab 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4812
      _ExtentX        =   8488
      _ExtentY        =   656
      Captions        =   "Home,XP,Wilderness,Notes,Links"
      TabActiveColor  =   -2147483633
   End
   Begin VB.Line lin 
      X1              =   5040
      X2              =   5940
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "userToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ButtonClick(Caption As String)
Public Event TabClick(Caption As String)

Private mblnInit As Boolean
Private mblnOverride As Boolean

Private mstrCaptions As String
Private mlngFitWidth As Long


' ************* PUBLIC *************


Public Property Get Captions() As String
    Captions = mstrCaptions
End Property

Public Property Let Captions(ByVal pstrCaptions As String)
    mstrCaptions = pstrCaptions
    UserControl.usrTab.Captions = mstrCaptions
End Property

Public Sub BulkChange(pstrCaptions As String, pstrActiveTab As String, plngTabActiveColor As Long)
    mstrCaptions = pstrCaptions
    UserControl.usrTab.BulkChange pstrCaptions, pstrActiveTab, plngTabActiveColor
End Sub

Public Sub Init()
    With UserControl
        mlngFitWidth = .usrTab.TabsWidth + .chkButton(0).Width * 3
    End With
    mblnInit = True
End Sub

Public Property Get FitWidth() As Long
    FitWidth = mlngFitWidth
End Property

Public Property Get FitHeight() As Long
    FitHeight = UserControl.chkButton(0).Height
End Property

Public Sub RefreshColors()
    Dim i As Long
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        With .usrTab
            .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
            .TextActiveColor = cfg.GetColor(cgeControls, cveText)
            .TabActiveColor = cfg.GetColor(cgeControls, cveBackground)
            .TextInactiveColor = cfg.GetColor(cgeWorkspace, cveText)
            .TabInactiveColor = cfg.GetColor(cgeWorkspace, cveBackground)
        End With
        For i = 0 To .chkButton.UBound
            cfg.ApplyColors .chkButton(i), cgeControls
        Next
        .lin.BorderColor = .usrTab.BorderColor
    End With
End Sub

Public Sub ShowPlayButton()
    UserControl.chkButton(2).Visible = cfg.PlayButton
End Sub

Public Property Get TabActiveColor() As Long
    TabActiveColor = UserControl.usrTab.TabActiveColor
End Property

Public Property Let TabActiveColor(plngColor As Long)
    With UserControl.usrTab
        If .TabActiveColor <> plngColor Then .TabActiveColor = plngColor
    End With
End Property


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mstrCaptions = "Tab 1,Tab 2,Tab 3"
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Captions", mstrCaptions, "Tab 1,Tab 2,Tab 3"
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mstrCaptions = PropBag.ReadProperty("Captions", "Tab 1,Tab 2,Tab 3")
    UserControl.usrTab.Captions = mstrCaptions
End Sub


' ************* PRIVATE *************


Private Sub UserControl_Resize()
    Redraw
End Sub

Private Sub Redraw()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim i As Long
    
    With UserControl
        ' Buttons
        lngLeft = .ScaleWidth
        lngTop = 0
        lngWidth = .chkButton(0).Width
        lngHeight = .chkButton(0).Height
        For i = 0 To 2
            lngLeft = lngLeft - lngWidth
            .chkButton(i).Move lngLeft, lngTop, lngWidth, lngHeight
        Next
        ' Tabstrip
        lngWidth = .usrTab.TabsWidth
        lngHeight = .usrTab.TabHeight
        lngLeft = 0
        lngTop = .chkButton(0).Height - .usrTab.TabHeight
        .usrTab.Move lngLeft, lngTop, lngWidth, lngHeight
        ' Line
        .lin.X1 = lngWidth
        .lin.X2 = .ScaleWidth
        .lin.Y1 = .ScaleHeight - Screen.TwipsPerPixelY
        .lin.Y2 = .ScaleHeight - Screen.TwipsPerPixelY
    End With
End Sub

Private Sub chkButton_Click(Index As Integer)
    Dim strCaption As String
    
    If UncheckButton(UserControl.chkButton(Index), mblnOverride) Then Exit Sub
    strCaption = UserControl.chkButton(Index).Caption
    RaiseEvent ButtonClick(strCaption)
End Sub

Private Sub usrTab_Click(pstrCaption As String)
    RaiseEvent TabClick(pstrCaption)
End Sub

