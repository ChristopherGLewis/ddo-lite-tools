VERSION 5.00
Begin VB.UserControl userHeader 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   384
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   384
   ScaleWidth      =   12216
   Begin VB.Label lblCenter 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Left            =   6300
      TabIndex        =   19
      Top             =   84
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   9
      Left            =   5040
      TabIndex        =   18
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   8
      Left            =   4500
      TabIndex        =   17
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   7
      Left            =   3972
      TabIndex        =   16
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   6
      Left            =   3444
      TabIndex        =   15
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   5
      Left            =   2916
      TabIndex        =   14
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblSlash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " / "
      Height          =   216
      Left            =   8100
      TabIndex        =   13
      Top             =   84
      Visible         =   0   'False
      Width           =   204
   End
   Begin VB.Label lblColon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Colon: "
      Height          =   216
      Left            =   7440
      TabIndex        =   12
      Top             =   84
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblParenClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ") "
      Height          =   216
      Left            =   7200
      TabIndex        =   11
      Top             =   84
      Visible         =   0   'False
      Width           =   144
   End
   Begin VB.Label lblParenOpen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "("
      Height          =   216
      Left            =   7020
      TabIndex        =   10
      Top             =   84
      Visible         =   0   'False
      Width           =   84
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   4
      Left            =   8436
      TabIndex        =   9
      Top             =   84
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   1
      Left            =   10704
      TabIndex        =   8
      Top             =   84
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   2
      Left            =   9948
      TabIndex        =   7
      Top             =   84
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   3
      Left            =   9192
      TabIndex        =   6
      Top             =   84
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   4
      Left            =   2388
      TabIndex        =   5
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   3
      Left            =   1860
      TabIndex        =   4
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   1
      Left            =   792
      TabIndex        =   2
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   0
      Left            =   11460
      TabIndex        =   1
      Top             =   84
      Width           =   516
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   0
      Left            =   264
      TabIndex        =   0
      Top             =   84
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   216
      Left            =   5520
      TabIndex        =   20
      Top             =   84
      Visible         =   0   'False
      Width           =   684
   End
   Begin VB.Shape shpBackground 
      BackColor       =   &H8000000F&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   384
      Left            =   0
      Top             =   0
      Width           =   12216
   End
End
Attribute VB_Name = "userHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Advanced styles:
'
' Use a pipe (|) character to separate the first part in parentheses. For example:
' "New|Gearset" will appear as "(New) Gearset" and will include two clickable links
' The actual click will send the caption of the individual link ("New" or "Gearset")
'
' Use a colon (":") to denote non-clickable link to start with, and a "/" to split the
' clickable links into two. For example:
' "Gearset: New / Load" will have two clickable links: "New" and "Load"
Option Explicit

' I don't know offhand how to expose user-defined properties that are lists to design view
' So I'm just taking the poor man's route and using delimited strings

Private mstrLeftLinks As String
Private mstrRightLinks As String
Private mstrCenterLink As String
Private mlngMargin As Long
Private mlngSpacing As Long
Private mlngCaptionTop As Long
Private mstrCaption As String
Private mlngTextColor As Long
Private mlngLinkColor As Long
Private mlngErrorColor As Long
Private mblnUseTabs As Boolean
Private mlngCurrent As Long
Private mblnEnabled As Boolean

Public Event Click(pstrCaption As String)
Public Event MouseOver()

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long

Private Const CURSOR_HAND As Long = 32649&


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mlngMargin = 264
    mlngSpacing = 564
    mlngCaptionTop = 84
    mstrCaption = vbNullString
    mblnUseTabs = True
    mlngTextColor = vbBlack
    mlngLinkColor = vbBlue
    mlngErrorColor = vbRed
    mlngCurrent = 0
    mblnEnabled = True
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Margin", mlngMargin, 264
    PropBag.WriteProperty "Spacing", mlngMargin, 564
    PropBag.WriteProperty "CaptionTop", mlngCaptionTop, 84
    PropBag.WriteProperty "Caption", mstrCaption, vbNullString
    PropBag.WriteProperty "UseTabs", mblnUseTabs, True
    PropBag.WriteProperty "TextColor", mlngTextColor, vbBlack
    PropBag.WriteProperty "LinkColor", mlngLinkColor, vbBlue
    PropBag.WriteProperty "ErrorColor", mlngErrorColor, vbRed
    PropBag.WriteProperty "BackColor", UserControl.shpBackground.FillColor, vbButtonFace
    PropBag.WriteProperty "BorderColor", UserControl.shpBackground.BorderColor, vbBlack
    PropBag.WriteProperty "LeftLinks", mstrLeftLinks, vbNullString
    PropBag.WriteProperty "CenterLink", mstrCenterLink, vbNullString
    PropBag.WriteProperty "RightLinks", mstrRightLinks, vbNullString
    PropBag.WriteProperty "Enabled", mblnEnabled, True
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mlngMargin = PropBag.ReadProperty("Margin", 264)
    mlngSpacing = PropBag.ReadProperty("Margin", 564)
    mlngCaptionTop = PropBag.ReadProperty("CaptionTop", 84)
    mstrCaption = PropBag.ReadProperty("Caption", vbNullString)
    mblnUseTabs = PropBag.ReadProperty("UseTabs", True)
    mlngTextColor = PropBag.ReadProperty("TextColor", vbBlack)
    mlngLinkColor = PropBag.ReadProperty("LinkColor", vbBlue)
    mlngErrorColor = PropBag.ReadProperty("ErrorColor", vbRed)
    UserControl.shpBackground.FillColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    UserControl.shpBackground.BorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
    mstrLeftLinks = PropBag.ReadProperty("LeftLinks", vbNullString)
    mstrCenterLink = PropBag.ReadProperty("CenterLink", vbNullString)
    mstrRightLinks = PropBag.ReadProperty("RightLinks", vbNullString)
    mblnEnabled = PropBag.ReadProperty("Enabled", True)
    ShowLinks
End Sub


' ************* METHODS *************


Public Sub SetTab(plngTab As Long)
    UserControl.lblLeft(mlngCurrent).FontUnderline = False
    mlngCurrent = plngTab
    UserControl.lblLeft(mlngCurrent).FontUnderline = True
    RaiseEvent Click(UserControl.lblLeft(mlngCurrent).Caption)
End Sub

' When dynamically recreating tabs, use this to force the underline without throwing a change event
Public Sub SyncTab(pstrCaption As String)
    Dim i As Long
    
    If Not mblnUseTabs Then Exit Sub
    For i = 0 To UserControl.lblLeft.UBound
        If UserControl.lblLeft(i).Caption = pstrCaption Then
            If mlngCurrent <> i Then UserControl.lblLeft(mlngCurrent).FontUnderline = False
            mlngCurrent = i
            Exit For
        End If
    Next
    If mlngCurrent > UserControl.lblLeft.UBound Then mlngCurrent = 0
    UserControl.lblLeft(mlngCurrent).FontUnderline = True
End Sub

Public Sub SetError(pstrTabCaption As String, pblnError As Boolean)
    Dim lngColor As Long
    Dim ctl As Control
    
    For Each ctl In UserControl.Controls
        Select Case ctl.Name
            Case "lblLeft", "lblCenter", "lblRight"
                If ctl.Caption = pstrTabCaption Then
                    If pblnError Then
                        ctl.Tag = "Error"
                        lngColor = mlngErrorColor
                    Else
                        ctl.Tag = vbNullString
                        lngColor = mlngLinkColor
                    End If
                    If ctl.ForeColor <> lngColor Then ctl.ForeColor = lngColor
                    Exit For
                End If
        End Select
    Next
End Sub


' ************* PROPERTIES *************


Public Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property

Public Property Let Enabled(ByVal pblnEnabled As Boolean)
    mblnEnabled = pblnEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get Margin() As Long
    Margin = mlngMargin
End Property

Public Property Let Margin(ByVal plngMargin As Long)
    mlngMargin = plngMargin
    PropertyChanged "Margin"
    ShowLinks
End Property

Public Property Get Spacing() As Long
    Spacing = mlngSpacing
End Property

Public Property Let Spacing(ByVal plngSpacing As Long)
    mlngSpacing = plngSpacing
    PropertyChanged "Spacing"
    ShowLinks
End Property

Public Property Get CaptionTop() As Long
    CaptionTop = mlngCaptionTop
End Property

Public Property Let CaptionTop(plngCaptionTop As Long)
    Dim ctl As Control
    
    mlngCaptionTop = plngCaptionTop
    PropertyChanged "CaptionTop"
    ShowLinks
End Property

Public Property Get Caption() As String
    Caption = mstrCaption
End Property

Public Property Let Caption(pstrCaption As String)
    If mstrCaption = pstrCaption Then Exit Property
    mstrCaption = pstrCaption
    UserControl.lblCaption.Caption = mstrCaption
    PropertyChanged "Caption"
    ShowLinks
End Property

Public Property Get UseTabs() As Boolean
    UseTabs = mblnUseTabs
End Property

Public Property Let UseTabs(ByVal pblnUseTabs As Boolean)
    mblnUseTabs = pblnUseTabs
    ShowLinks
    PropertyChanged "UseTabs"
End Property

' Check out the fancy property type, complete with dropdowns in the Properties window of Design view
Public Property Get TextColor() As OLE_COLOR
    TextColor = mlngTextColor
End Property

Public Property Let TextColor(ByVal poleColor As OLE_COLOR)
    Dim ctl As Control
    
    mlngTextColor = poleColor
    For Each ctl In UserControl.Controls
        If TypeOf ctl Is Label Then
            Select Case ctl.Name
                Case "lblLeft", "lblRight", "lblCenter"
                Case Else: ctl.ForeColor = mlngTextColor
            End Select
        End If
    Next
    PropertyChanged "TextColor"
End Property

Public Property Get LinkColor() As OLE_COLOR
    LinkColor = mlngLinkColor
End Property

Public Property Let LinkColor(ByVal poleColor As OLE_COLOR)
    Dim lngColor As Long
    Dim ctl As Control
    
    mlngLinkColor = poleColor
    For Each ctl In UserControl.Controls
        Select Case ctl.Name
            Case "lblLeft", "lblRight", "lblCenter"
                If ctl.Tag = "Error" Then lngColor = mlngErrorColor Else lngColor = mlngLinkColor
                If ctl.ForeColor <> lngColor Then ctl.ForeColor = lngColor
        End Select
    Next
    PropertyChanged "LinkColor"
End Property

Public Property Get ErrorColor() As OLE_COLOR
    ErrorColor = mlngErrorColor
End Property

Public Property Let ErrorColor(ByVal poleColor As OLE_COLOR)
    Dim lngColor As Long
    Dim i As Long
    
    mlngErrorColor = poleColor
    For i = 0 To UserControl.lblLeft.UBound
        With UserControl.lblLeft(i)
            If .Tag = "Error" Then lngColor = mlngErrorColor Else lngColor = mlngLinkColor
            If .ForeColor <> lngColor Then .ForeColor = lngColor
        End With
    Next
    PropertyChanged "ErrorColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.shpBackground.FillColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    UserControl.shpBackground.FillColor = poleColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = UserControl.shpBackground.BorderColor
End Property

Public Property Let BorderColor(ByVal poleColor As OLE_COLOR)
    UserControl.shpBackground.BorderColor = poleColor
    PropertyChanged "BorderColor"
End Property

Public Property Get LeftLinks() As String
    LeftLinks = mstrLeftLinks
End Property

Public Property Let LeftLinks(ByVal pstrLeftLinks As String)
    Dim i As Long
    
    For i = 0 To UserControl.lblLeft.UBound
        With UserControl.lblLeft(i)
            If .Tag = "Error" Then
                .Tag = vbNullString
                .ForeColor = mlngLinkColor
            End If
        End With
    Next
    mstrLeftLinks = pstrLeftLinks
    PropertyChanged "LeftLinks"
    ShowLinks
End Property

Public Property Get CenterLink() As String
    CenterLink = mstrCenterLink
End Property

Public Property Let CenterLink(ByVal pstrCenterLink As String)
    mstrCenterLink = pstrCenterLink
    PropertyChanged "CenterLink"
    ShowLinks
End Property

Public Property Get RightLinks() As String
    RightLinks = mstrRightLinks
End Property

Public Property Let RightLinks(ByVal pstrRightLinks As String)
    mstrRightLinks = pstrRightLinks
    PropertyChanged "RightLinks"
    ShowLinks
End Property


' ************* DRAWING *************


Private Sub UserControl_Resize()
    With UserControl
        .shpBackground.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
    ShowLinks
End Sub

Private Sub ShowLinks()
    Dim strLink() As String
    Dim lngLinks As Long
    Dim lngLeft As Long
    Dim lngIndex As Long
    Dim lngLabel As Long
    Dim lngLink As Long
    Dim strOld As String
    Dim ctl As Control
    Dim i As Long
    
    With UserControl
        For Each ctl In UserControl.Controls
            If TypeOf ctl Is Label Then ctl.Top = mlngCaptionTop
        Next
        .lblParenOpen.Visible = False
        .lblParenClose.Visible = False
        .lblColon.Visible = False
        .lblSlash.Visible = False
        .lblCaption.Left = mlngMargin
        .lblCaption.Visible = (Len(mstrCaption) > 0)
        ' Left links
        If Len(mstrLeftLinks) Then
            lngLeft = mlngMargin
            strLink = Split(mstrLeftLinks, ";")
            If mblnUseTabs Then
                strOld = .lblLeft(mlngCurrent).Caption
                For mlngCurrent = 0 To UBound(strLink)
                    If strLink(mlngCurrent) = strOld Then Exit For
                Next
                If mlngCurrent > UBound(strLink) Then mlngCurrent = 0
                If strOld = "Selected" And strLink(0) = "General" Then mlngCurrent = 0
            End If
            For lngLink = 0 To UBound(strLink)
                If InStr(strLink(lngLink), "|") Then
                    ShowPipeLink strLink(lngLink), lngLeft, lngLabel
                ElseIf InStr(strLink(lngLink), ": ") Then
                    ShowColonLink strLink(lngLink), lngLeft, lngLabel
                Else
                    SetLabel .lblLeft(lngLabel), strLink(lngLink), lngLeft, True, (mblnUseTabs And mlngCurrent = lngLink)
                    lngLabel = lngLabel + 1
                End If
            Next
        End If
        For i = lngLink To .lblLeft.UBound
            .lblLeft(i).Visible = False
        Next
        ' Center link
        If Len(mstrCenterLink) Then
            .lblCenter.Caption = mstrCenterLink
            .lblCenter.Left = (.ScaleWidth - .lblCenter.Width) \ 2
        End If
        .lblCenter.Visible = (Len(mstrCenterLink) > 0)
        ' Right links
        strLink = Split(mstrRightLinks, ";")
        lngLinks = UBound(strLink) + 1
        lngLeft = .shpBackground.Width - mlngMargin
        For i = lngLinks - 1 To 0 Step -1
            lngIndex = lngLinks - 1 - i
            With .lblRight(lngIndex)
                .Caption = strLink(i)
                .Left = lngLeft - .Width
                lngLeft = .Left - mlngSpacing
                .Visible = True
            End With
        Next
        For i = lngLinks To .lblRight.UBound
            .lblRight(i).Visible = False
        Next
    End With
End Sub

Private Sub SetLabel(plbl As Label, pstrCaption As String, plngLeft As Long, pblnSpacer As Boolean, Optional pblnUnderline As Boolean = False)
    plbl.Caption = pstrCaption
    plbl.Left = plngLeft
    plbl.FontUnderline = pblnUnderline
    plbl.Visible = True
    plngLeft = plngLeft + plbl.Width
    If pblnSpacer Then plngLeft = plngLeft + mlngSpacing
End Sub

Private Sub ShowPipeLink(pstrLink As String, plngLeft As Long, plngLabel As Long)
    Dim strToken() As String
    
    strToken = Split(pstrLink, "|")
    With UserControl
        SetLabel .lblParenOpen, "(", plngLeft, False
        SetLabel .lblLeft(plngLabel), strToken(0), plngLeft, False
        SetLabel .lblParenClose, ") ", plngLeft, False
        SetLabel .lblLeft(plngLabel + 1), strToken(1), plngLeft, True
    End With
    plngLabel = plngLabel + 2
End Sub

Private Sub ShowColonLink(pstrLink As String, plngLeft As Long, plngLabel As Long)
    Dim strColon As String
    Dim strToken() As String
    Dim lngPos As Long
    
    lngPos = InStr(pstrLink, ": ")
    If Len(lngPos) = 0 Then Exit Sub
    strColon = Left$(pstrLink, lngPos + 1)
    strToken = Split(Mid$(pstrLink, lngPos + 2), " / ")
    With UserControl
        SetLabel .lblColon, strColon, plngLeft, False
        SetLabel .lblLeft(plngLabel), strToken(0), plngLeft, False
        SetLabel .lblSlash, " / ", plngLeft, False
        SetLabel .lblLeft(plngLabel + 1), strToken(1), plngLeft, True
    End With
    plngLabel = plngLabel + 2
End Sub


' ************* MOUSE *************


Private Sub lblLeft_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnEnabled Then SetCursor LoadCursor(0, CURSOR_HAND)
    RaiseEvent MouseOver
End Sub

Private Sub lblLeft_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnEnabled Then Exit Sub
    SetCursor LoadCursor(0, CURSOR_HAND)
    If Button = vbLeftButton Then
        RaiseEvent Click(UserControl.lblLeft(Index).Caption)
        If mblnUseTabs Then
            UserControl.lblLeft(mlngCurrent).FontUnderline = False
            mlngCurrent = Index
            UserControl.lblLeft(mlngCurrent).FontUnderline = True
        End If
    End If
End Sub

Private Sub lblLeft_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnEnabled Then SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lblRight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnEnabled Then SetCursor LoadCursor(0, CURSOR_HAND)
    RaiseEvent MouseOver
End Sub

Private Sub lblCenter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnEnabled Then SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lblCenter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnEnabled Then Exit Sub
    SetCursor LoadCursor(0, CURSOR_HAND)
    If Button = vbLeftButton Then RaiseEvent Click(UserControl.lblCenter.Caption)
End Sub

Private Sub lblCenter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnEnabled Then SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lblRight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnEnabled Then Exit Sub
    SetCursor LoadCursor(0, CURSOR_HAND)
    If Button = vbLeftButton Then RaiseEvent Click(UserControl.lblRight(Index).Caption)
End Sub

Private Sub lblRight_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnEnabled Then SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseOver
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseOver
End Sub

