VERSION 5.00
Begin VB.UserControl userInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3072
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4656
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3072
   ScaleWidth      =   4656
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2652
      Left            =   0
      ScaleHeight     =   2628
      ScaleWidth      =   4608
      TabIndex        =   1
      Top             =   420
      Width           =   4632
      Begin VB.VScrollBar scrollVertical 
         Height          =   2652
         Left            =   4380
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1272
         Left            =   360
         ScaleHeight     =   1272
         ScaleWidth      =   3912
         TabIndex        =   2
         Top             =   600
         Width           =   3912
         Begin VB.Timer tmrTooltip 
            Interval        =   1000
            Left            =   1560
            Top             =   420
         End
         Begin VB.Label lnkAugment 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Augment"
            ForeColor       =   &H00FF0000&
            Height          =   216
            Index           =   0
            Left            =   2700
            TabIndex        =   8
            Tag             =   "ctl"
            Top             =   0
            Visible         =   0   'False
            Width           =   804
         End
         Begin VB.Label lnkCopy 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Copy to Clipboard"
            ForeColor       =   &H00FF0000&
            Height          =   216
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Tag             =   "ctl"
            Top             =   0
            Visible         =   0   'False
            Width           =   1608
         End
         Begin VB.Label lblTooltip 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   " Copied! "
            ForeColor       =   &H80000017&
            Height          =   216
            Left            =   360
            TabIndex        =   6
            Tag             =   "tip"
            Top             =   240
            Visible         =   0   'False
            Width           =   804
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Link"
            ForeColor       =   &H00FF0000&
            Height          =   216
            Index           =   0
            Left            =   1980
            TabIndex        =   5
            Tag             =   "ctl"
            Top             =   0
            Visible         =   0   'False
            Width           =   348
         End
      End
   End
   Begin VB.Label lblTitleLarge 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Large Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   540
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1224
   End
   Begin VB.Label lblTitleSmall 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Small Title"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgIcon 
      Height          =   384
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "userInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Click(strLink As String)
Public Event BackgroundClick()

Public Enum LinkStyle
    lseHelp
    lseURL
    lseShard
    lseMaterial
    lseItem
    lseSchool
    lseCommand
    lseForm
End Enum

Public Enum InfoTitleSizeEnum
    iteLarge
    iteSmall
    iteNone
End Enum

Private Type WordType
    Text As String
    X As Long
    Y As Long
    ErrorColor As Boolean
    Bold As Boolean
    Italics As Boolean
    Underline As Boolean
    FontName As String
    FontSize As Long
End Type

Private Type AugmentLinkType
    ColorValue As ColorValueEnum
    Augment As Long
    Variation As Long
    Scaling As Long
End Type

Private mtypWord() As WordType
Private mlngWords As Long
Private mlngBuffer As Long

Private mtypAugment() As AugmentLinkType
Private mlngAugments As Long

Private menTitleSize As InfoTitleSizeEnum
Private mblnTitleIcon As Boolean
Private mstrTitleText As String
Private mlngTitleForeColor As Long
Private mlngTitleBackColor As Long

Private mlngBackColor As Long
Private mlngTextColor As Long
Private mlngErrorColor As Long
Private mlngLinkColor As Long

Private mblnCanScroll As Boolean

Private mlngLinkIndex As Long
Private mlngCopyIndex As Long

Private mlngIndent As Long
Private mlngIndentWrap As Long
Private mlngLastLinkLeft As Long ' Left coordinate of the last link added

Private mlngMarginX As Long
Private mlngMarginY As Long

Private mlngTextHeight As Long
Private mblnOverride As Boolean

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long

Private Const CURSOR_HAND As Long = 32649&


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    menTitleSize = iteLarge
    mblnTitleIcon = True
    mstrTitleText = "Title"
    mlngTitleForeColor = vbBlack
    mlngTitleBackColor = vbWhite
    
    mlngBackColor = vbWhite
    mlngTextColor = vbBlack
    mlngErrorColor = vbRed
    mlngLinkColor = vbBlue
    
    mlngMarginX = 0
    mlngMarginY = 0
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TitleSize", menTitleSize, iteLarge
    PropBag.WriteProperty "TitleIcon", mblnTitleIcon, True
    PropBag.WriteProperty "TitleText", mstrTitleText, "Title"
    PropBag.WriteProperty "TitleForeColor", mlngTitleForeColor, vbBlack
    PropBag.WriteProperty "TitleBackColor", mlngTitleBackColor, vbWhite
    PropBag.WriteProperty "BackColor", mlngBackColor, vbWhite
    PropBag.WriteProperty "TextColor", mlngTextColor, vbBlack
    PropBag.WriteProperty "ErrorColor", mlngErrorColor, vbRed
    PropBag.WriteProperty "LinkColor", mlngLinkColor, vbBlue
    PropBag.WriteProperty "CanScroll", mblnCanScroll, True
    PropBag.WriteProperty "MarginX", mlngMarginX, 0
    PropBag.WriteProperty "MarginY", mlngMarginY, 0
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    menTitleSize = PropBag.ReadProperty("TitleSize", iteLarge)
    mblnTitleIcon = PropBag.ReadProperty("TitleIcon", True)
    mstrTitleText = PropBag.ReadProperty("TitleText", "Title")
    mlngTitleForeColor = PropBag.ReadProperty("TitleForeColor", vbBlack)
    mlngTitleBackColor = PropBag.ReadProperty("TitleBackColor", vbWhite)
    mlngBackColor = PropBag.ReadProperty("BackColor", vbWhite)
    mlngTextColor = PropBag.ReadProperty("TextColor", vbBlack)
    mlngErrorColor = PropBag.ReadProperty("ErrorColor", vbRed)
    mlngLinkColor = PropBag.ReadProperty("LinkColor", vbBlue)
    mblnCanScroll = PropBag.ReadProperty("CanScroll", True)
    mlngMarginX = PropBag.ReadProperty("MarginX", 0)
    mlngMarginY = PropBag.ReadProperty("MarginY", 0)
    ClearContents
    DrawControl
End Sub

Private Sub UserControl_Click()
    RaiseEvent BackgroundClick
End Sub

Private Sub picClient_Click()
    RaiseEvent BackgroundClick
End Sub


' ************* METHODS *************


Public Sub Clear()
    With UserControl
        .imgIcon.Visible = False
        .imgIcon.Picture = LoadPicture()
        .lblTitleLarge.Visible = False
        .lblTitleSmall.Visible = False
    End With
    ClearContents
End Sub

Public Sub ClearContents()
    Dim i As Long
    
    mlngWords = 0
    mlngBuffer = 127
    ReDim mtypWord(mlngBuffer)
    mlngAugments = 0
    Erase mtypAugment
    mlngLinkIndex = 0
    mlngCopyIndex = 0
    mlngIndent = 0
    mlngIndentWrap = 0
    With UserControl
        mlngTextHeight = .picClient.TextHeight("Q")
        .scrollVertical.Value = 0
        .scrollVertical.Visible = False
        .picClient.Move 0, 0, .picClient.Width, .picContainer.ScaleHeight
        .picClient.Cls
        For i = 0 To .lnkLink.UBound
            .lnkLink(i).Visible = False
        Next
        For i = 0 To .lnkCopy.UBound
            .lnkCopy(i).Visible = False
        Next
        For i = 0 To .lnkAugment.UBound
            .lnkAugment(i).Visible = False
        Next
        .lblTooltip.Visible = False
        .tmrTooltip.Enabled = False
    End With
End Sub

Public Sub Scroll(plngValue As Long)
    Dim lngIncrement As Long
    Dim lngValue As Long
    
    If Not UserControl.scrollVertical.Visible Then Exit Sub
    lngIncrement = plngValue * (UserControl.picClient.TextHeight("Q") \ Screen.TwipsPerPixelY)
    With UserControl.scrollVertical
        lngValue = .Value + lngIncrement
        If lngValue < .Min Then lngValue = .Min
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub

Public Sub AddText(pstrText As String, Optional plngNewLines As Long = 1, Optional pstrIndent As String, Optional pstrIndentAfterWrap As String)
    Dim strLine() As String
    Dim i As Long
    
    If Len(pstrIndent) Then mlngIndent = UserControl.picClient.TextWidth(pstrIndent)
    If Len(pstrIndentAfterWrap) Then mlngIndentWrap = UserControl.picClient.TextWidth(pstrIndentAfterWrap)
    If InStr(pstrText, vbNewLine) Then
        strLine = Split(pstrText, vbNewLine)
        For i = 0 To UBound(strLine)
            DrawText UserControl.picClient, strLine(i), False, 1
        Next
'        NewLine 1
    Else
        DrawText UserControl.picClient, pstrText, False, plngNewLines
    End If
    FinishDraw plngNewLines
End Sub

Public Sub AddTextFormatted(pstrText As String, pblnBold As Boolean, Optional pblnItalics As Boolean, Optional pblnUnderline As Boolean, Optional plngColor As Long = -1, Optional plngNewLines As Long = 1, Optional pstrIndent As String, Optional pstrIndentAfterWrap As String, Optional pblnFixed As Boolean)
    Dim strLine() As String
    Dim i As Long
    
    With UserControl.picClient
        If pblnBold Then .FontBold = True
        If pblnItalics Then .FontItalic = True
        If pblnUnderline Then .FontUnderline = True
        If pblnFixed Then
            .FontName = "Courier"
            .FontSize = 10
        End If
        If plngColor <> -1 Then .ForeColor = plngColor
    End With
    If Len(pstrIndent) Then mlngIndent = UserControl.picClient.TextWidth(pstrIndent)
    If Len(pstrIndentAfterWrap) Then mlngIndentWrap = UserControl.picClient.TextWidth(pstrIndentAfterWrap)
    If InStr(pstrText, vbNewLine) Then
        strLine = Split(pstrText, vbNewLine)
        For i = 0 To UBound(strLine)
            DrawText UserControl.picClient, strLine(i), False, 1
        Next
        NewLine plngNewLines
    Else
        DrawText UserControl.picClient, pstrText, False, plngNewLines
    End If
    FinishDraw plngNewLines
    With UserControl.picClient
        If pblnBold Then .FontBold = False
        If pblnItalics Then .FontItalic = False
        If pblnUnderline Then .FontUnderline = False
        If pblnFixed Then
            .FontName = "Verdana"
            .FontSize = 9
        End If
        If plngColor <> -1 Then .ForeColor = mlngTextColor
    End With
End Sub

Private Sub FinishDraw(plngNewLines As Long)
    If (mlngIndent <> 0 Or mlngIndentWrap <> 0) And plngNewLines <> 0 Then UserControl.picClient.CurrentX = 0
    mlngIndent = 0
    mlngIndentWrap = 0
End Sub

Public Sub AddError(pstrText As String, Optional plngNewLines As Long = 1)
    DrawText UserControl.picClient, pstrText, True, plngNewLines
End Sub

Public Sub AddLink(pstrCaption As String, penStyle As LinkStyle, Optional pstrLink As String, Optional plngNewLines As Long = 1, Optional pblnCanWrap As Boolean = True)
    Dim strLink As String
    
    If Len(pstrLink) Then strLink = pstrLink Else strLink = pstrCaption
    With UserControl
        If mlngLinkIndex > .lnkLink.UBound Then Load .lnkLink(mlngLinkIndex)
        With .lnkLink(mlngLinkIndex)
            .Caption = pstrCaption
            .Tag = penStyle & strLink
            If penStyle = lseURL Then .TooltipText = pstrLink Else .TooltipText = vbNullString
            If pblnCanWrap Then
                If UserControl.picClient.CurrentX + .Width > UserControl.picClient.ScaleWidth Then NewLine 1
            End If
            .Move UserControl.picClient.CurrentX, UserControl.picClient.CurrentY
            mlngLastLinkLeft = .Left
            .Visible = True
            UserControl.picClient.CurrentX = .Left + .Width
        End With
    End With
    mlngLinkIndex = mlngLinkIndex + 1
    NewLine plngNewLines
End Sub

Public Sub AddLinkParentheses(pstrCaption As String, penStyle As LinkStyle, Optional pstrLink As String, Optional plngNewLines As Long = 1, Optional pblnCanWrap As Boolean = True)
    Dim strLink As String
    
    If Len(pstrLink) Then strLink = pstrLink Else strLink = pstrCaption
    With UserControl
        If mlngLinkIndex > .lnkLink.UBound Then Load .lnkLink(mlngLinkIndex)
        With .lnkLink(mlngLinkIndex)
            .Caption = pstrCaption
            .Tag = penStyle & strLink
            If penStyle = lseURL Then .TooltipText = pstrLink Else .TooltipText = vbNullString
            If pblnCanWrap Then
                If UserControl.picClient.CurrentX + .Width + UserControl.picClient.TextWidth("()") > UserControl.picClient.ScaleWidth Then NewLine 1
            End If
            PrintWord "("
            .Move UserControl.picClient.CurrentX, UserControl.picClient.CurrentY
            .Visible = True
            UserControl.picClient.CurrentX = .Left + .Width
            PrintWord ")"
        End With
    End With
    mlngLinkIndex = mlngLinkIndex + 1
    NewLine plngNewLines
End Sub

Public Sub AddClipboard(pstrText As String, Optional plngNewLines As Long = 1)
    PrintWord "("
    With UserControl
        If mlngCopyIndex > .lnkCopy.UBound Then Load .lnkCopy(mlngCopyIndex)
        With .lnkCopy(mlngCopyIndex)
            .Tag = pstrText
            .Move UserControl.picClient.CurrentX, UserControl.picClient.CurrentY
            .Visible = True
            UserControl.picClient.CurrentX = .Left + .Width
        End With
    End With
    PrintWord ")"
    mlngCopyIndex = mlngCopyIndex + 1
    NewLine plngNewLines
End Sub

Public Sub AddNumber(plngNumber As Long, plngPadding As Long, Optional pstrIndent As String)
    Dim strText As String
    Dim lngLeft As Long
    Dim lngPadding As Long
    Dim strLine() As String
    Dim i As Long
    
    strText = plngNumber
    lngPadding = plngPadding - Len(strText)
    If lngPadding < 0 Then lngPadding = 0
    UserControl.picClient.CurrentX = UserControl.picClient.TextWidth(pstrIndent & String(lngPadding, 57))
    PrintWord strText & " "
End Sub

Public Sub AddAugment(plngAugment As Long, plngVariant As Long, plngScale As Long, Optional plngNewLines As Long = 1, Optional pstrIndent As String, Optional plngLeft As Long = -1)
    Dim enColorValue As ColorValueEnum
    Dim lngLeft As Long
    
    If plngAugment = 0 Or plngVariant = 0 Or plngScale = 0 Then Exit Sub
    enColorValue = GetAugmentColorValue(db.Augment(plngAugment).Color)
    mlngAugments = mlngAugments + 1
    ReDim Preserve mtypAugment(1 To mlngAugments)
    With mtypAugment(mlngAugments)
        .Augment = plngAugment
        .Variation = plngVariant
        .Scaling = plngScale
        .ColorValue = enColorValue
    End With
    With UserControl
        If plngLeft <> -1 Then
            .picClient.CurrentX = plngLeft
        ElseIf Len(pstrIndent) > 0 And .picClient.CurrentX = 0 Then
            .picClient.CurrentX = UserControl.picClient.TextWidth(pstrIndent)
        End If
        If mlngAugments > .lnkAugment.UBound Then Load .lnkAugment(mlngAugments)
        With .lnkAugment(mlngAugments)
            .Caption = AugmentFullName(plngAugment, plngVariant, plngScale)
            .ForeColor = cfg.GetColor(cgeControls, enColorValue)
            .Move UserControl.picClient.CurrentX, UserControl.picClient.CurrentY
            .Visible = True
            UserControl.picClient.CurrentX = .Left + .Width
        End With
    End With
    NewLine plngNewLines
End Sub

Public Sub Indent(plngWidth As Long)
    mlngIndent = plngWidth
End Sub

Public Sub BackupOneSpace()
    Dim lngIndent As Long
    Dim lngSpace As Long
    
    With UserControl.picClient
        lngIndent = mlngIndent + mlngIndentWrap
        lngSpace = .TextWidth(" ")
        If .CurrentX >= lngIndent + lngSpace Then .CurrentX = .CurrentX - lngSpace
    End With
End Sub

Private Sub DrawText(pic As PictureBox, pstrText As String, pblnError As Boolean, plngNewLines As Long)
    Dim strWord() As String
    Dim strText As String
    Dim strLink As String
    Dim i As Long
    
    If pblnError Then pic.ForeColor = mlngErrorColor
    If mlngIndent Then pic.CurrentX = mlngIndent
    strWord = Split(pstrText, " ")
    For i = 0 To UBound(strWord)
        If Left$(strWord(i), 1) = "{" Then
            ParseText strWord(i), strText, strLink
            AddLink strText, lseURL, strLink, 0
            pic.Print " ";
        Else
            If pic.CurrentX + pic.TextWidth(strWord(i)) > pic.ScaleWidth Then NewLine 1
            If pic.FontUnderline Then
                PrintWord strWord(i)
            Else
                PrintWord strWord(i) & " "
            End If
        End If
    Next
    If pblnError Then pic.ForeColor = mlngTextColor
    NewLine plngNewLines
End Sub

Private Sub ParseText(pstrRaw As String, pstrText As String, pstrLink As String)
    Dim lngPos As Long
    
    pstrText = vbNullString
    pstrLink = vbNullString
    lngPos = InStr(pstrRaw, "}")
    If lngPos = 0 Then Exit Sub
    pstrLink = Left$(pstrRaw, lngPos - 1)
    pstrText = Mid$(pstrRaw, lngPos + 1)
    If InStr(pstrText, "_") Then pstrText = Replace(pstrText, "_", " ")
    If Left$(pstrLink, 5) = "{url=" Then pstrLink = Mid$(pstrLink, 6)
End Sub

Private Sub NewLine(plngNewLines As Long)
    Dim lngIndent As Long
    Dim i As Long
    
    If plngNewLines = 0 Then Exit Sub
    For i = 1 To plngNewLines
        UserControl.picClient.Print vbNullString
    Next
    lngIndent = mlngIndent + mlngIndentWrap
    If lngIndent Then UserControl.picClient.CurrentX = lngIndent
    SetScrollbar
End Sub

Public Sub SetIcon(pstrResourceID As String, Optional penStyle As LoadResConstants = vbResIcon)
    If mblnTitleIcon And menTitleSize = iteLarge Then
        UserControl.imgIcon.Picture = LoadResPicture(pstrResourceID, penStyle)
        UserControl.imgIcon.Visible = True
    End If
End Sub

Private Sub PrintWord(pstrWord As String)
    Dim typNew As WordType
    
    With UserControl.picClient
        typNew.Text = pstrWord
        typNew.X = .CurrentX
        typNew.Y = .CurrentY
        typNew.Bold = .FontBold
        typNew.Italics = .FontItalic
        typNew.Underline = .FontUnderline
        typNew.FontName = .FontName
        typNew.FontSize = .FontSize
        If .ForeColor = mlngErrorColor Then typNew.ErrorColor = True
    End With
    mlngWords = mlngWords + 1
    If mlngWords > mlngBuffer Then
        mlngBuffer = (mlngBuffer * 3) \ 2
        ReDim Preserve mtypWord(mlngBuffer)
    End If
    mtypWord(mlngWords) = typNew
    UserControl.picClient.Print pstrWord;
End Sub

Public Sub Redraw()
    Dim i As Long
    
    With UserControl
        With .picClient
            .Cls
            For i = 1 To mlngWords
                .FontName = mtypWord(i).FontName
                .FontSize = mtypWord(i).FontSize
                .FontBold = mtypWord(i).Bold
                .FontItalic = mtypWord(i).Italics
                .FontUnderline = mtypWord(i).Underline
                .CurrentX = mtypWord(i).X
                .CurrentY = mtypWord(i).Y
                If mtypWord(i).ErrorColor Then .ForeColor = mlngErrorColor
                UserControl.picClient.Print mtypWord(i).Text;
                If mtypWord(i).ErrorColor Then .ForeColor = mlngTextColor
            Next
        End With
        For i = 1 To mlngAugments
            .lnkAugment(i).ForeColor = cfg.GetColor(cgeControls, mtypAugment(i).ColorValue)
        Next
    End With
End Sub


' ************* PROPERTIES *************


Public Property Get hwnd() As Long
    hwnd = UserControl.picContainer.hwnd
End Property


Public Property Get CanScroll() As Boolean
    CanScroll = mblnCanScroll
End Property

Public Property Let CanScroll(ByVal pblnCanScroll As Boolean)
    mblnCanScroll = pblnCanScroll
    PropertyChanged "CanScroll"
    DrawControl
End Property


Public Property Get TitleForeColor() As OLE_COLOR
    TitleForeColor = mlngTitleForeColor
End Property

Public Property Let TitleForeColor(ByVal poleColor As OLE_COLOR)
    mlngTitleForeColor = poleColor
    PropertyChanged "TitleForeColor"
    UserControl.lblTitleLarge.ForeColor = mlngTitleForeColor
    UserControl.lblTitleSmall.ForeColor = mlngTitleForeColor
End Property


Public Property Get TitleBackColor() As OLE_COLOR
    TitleBackColor = mlngTitleBackColor
End Property

Public Property Let TitleBackColor(ByVal poleColor As OLE_COLOR)
    mlngTitleBackColor = poleColor
    PropertyChanged "TitleBackColor"
    UserControl.BackColor = mlngTitleBackColor
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = mlngBackColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    mlngBackColor = poleColor
    PropertyChanged "BackColor"
    UserControl.picContainer.BackColor = mlngBackColor
    UserControl.picClient.BackColor = mlngBackColor
End Property


Public Property Get TextColor() As OLE_COLOR
    TextColor = mlngTextColor
End Property

Public Property Let TextColor(ByVal poleColor As OLE_COLOR)
    mlngTextColor = poleColor
    PropertyChanged "TextColor"
    UserControl.picClient.ForeColor = mlngTextColor
End Property


Public Property Get ErrorColor() As OLE_COLOR
    ErrorColor = mlngErrorColor
End Property

Public Property Let ErrorColor(ByVal poleColor As OLE_COLOR)
    mlngErrorColor = poleColor
    PropertyChanged "ErrorColor"
End Property


Public Property Get LinkColor() As OLE_COLOR
    LinkColor = mlngLinkColor
End Property

Public Property Let LinkColor(ByVal poleColor As OLE_COLOR)
    Dim i As Long
    
    mlngLinkColor = poleColor
    PropertyChanged "LinkColor"
    With UserControl
        For i = 0 To .lnkLink.UBound
            .lnkLink(i).ForeColor = mlngLinkColor
        Next
        For i = 0 To .lnkCopy.UBound
            .lnkCopy(i).ForeColor = mlngLinkColor
        Next
    End With
End Property


Public Property Get TitleText() As String
    TitleText = mstrTitleText
End Property

Public Property Let TitleText(ByVal pstrTitleText As String)
    mstrTitleText = pstrTitleText
    PropertyChanged "TitleText"
    Select Case menTitleSize
        Case iteLarge
            UserControl.lblTitleLarge.Caption = mstrTitleText
            UserControl.lblTitleLarge.Visible = True
        Case iteSmall
            UserControl.lblTitleSmall.Caption = mstrTitleText
            UserControl.lblTitleSmall.Visible = True
    End Select
End Property


Public Property Get TitleSize() As InfoTitleSizeEnum
    TitleSize = menTitleSize
End Property

Public Property Let TitleSize(ByVal penSize As InfoTitleSizeEnum)
    menTitleSize = penSize
    PropertyChanged "TitleSize"
    DrawControl
End Property


Public Property Get TitleIcon() As Boolean
    TitleIcon = mblnTitleIcon
End Property

Public Property Let TitleIcon(ByVal pblnTitleIcon As Boolean)
    mblnTitleIcon = pblnTitleIcon
    PropertyChanged "TitleIcon"
    DrawControl
End Property


Public Property Get MarginX() As Long
    MarginX = mlngMarginX
End Property

Public Property Let MarginX(ByVal plngMarginX As Long)
    mlngMarginX = plngMarginX
    DrawControl
End Property


Public Property Get MarginY() As Long
    MarginY = mlngMarginY
End Property

Public Property Let MarginY(ByVal plngMarginY As Long)
    mlngMarginY = plngMarginY
    DrawControl
End Property


Public Property Get LastLinkLeft() As Long
    LastLinkLeft = mlngLastLinkLeft
End Property


' ************* RESIZE *************


Private Sub UserControl_Resize()
    DrawControl
End Sub

Private Sub DrawControl()
    Dim lngPixelX As Long
    Dim lngPixelY As Long
    Dim lngTop As Long
    Dim i As Long
    
    lngPixelX = Screen.TwipsPerPixelX
    lngPixelY = Screen.TwipsPerPixelY
    With UserControl
        .BackColor = mlngTitleBackColor
        .picContainer.BackColor = mlngBackColor
        .picClient.BackColor = mlngBackColor
        .picClient.ForeColor = mlngTextColor
        .imgIcon.Visible = False
        With .lblTitleSmall
            .Visible = False
            .Caption = mstrTitleText
            .ForeColor = mlngTitleForeColor
        End With
        With .lblTitleLarge
            .Visible = False
            .Caption = mstrTitleText
            .ForeColor = mlngTitleForeColor
        End With
        .scrollVertical.Visible = False
        Select Case menTitleSize
            Case iteLarge
                .lblTitleLarge.Caption = mstrTitleText
                If mblnTitleIcon Then
                    .imgIcon.Move 0, 0, 32 * lngPixelX, 32 * lngPixelY
                    .lblTitleLarge.Move .imgIcon.Width + 8 * lngPixelX, (.imgIcon.Height - .lblTitleLarge.Height) \ 2
                    lngTop = 35 * lngPixelY
                    .imgIcon.Visible = True
                Else
                    .lblTitleLarge.Move 0, 0
                    lngTop = .lblTitleLarge.Height + 5 * lngPixelY
                End If
                .lblTitleLarge.Visible = True
            Case iteSmall
                lngTop = .lblTitleSmall.Height + 2 * lngPixelY
                .lblTitleSmall.Visible = True
            Case iteNone
                lngTop = 0
        End Select
        .picContainer.Move 0, lngTop, .ScaleWidth, .ScaleHeight - lngTop
        If mblnCanScroll Then
            .picClient.Move 0, 0, .picContainer.ScaleWidth - .scrollVertical.Width, .picContainer.ScaleHeight
            .scrollVertical.Move .picContainer.ScaleWidth - .scrollVertical.Width + Screen.TwipsPerPixelX, 0, .scrollVertical.Width, .picContainer.ScaleHeight
            .scrollVertical.SmallChange = .picClient.TextHeight("Q") \ Screen.TwipsPerPixelY
            .scrollVertical.LargeChange = .picContainer.ScaleHeight \ Screen.TwipsPerPixelY
            .scrollVertical.Visible = True
        Else
            .picClient.Move mlngMarginX, mlngMarginY, .picContainer.ScaleWidth - (mlngMarginX * 2), .picContainer.ScaleHeight - mlngMarginY
        End If
    End With
End Sub

Private Sub SetScrollbar()
    Dim lngHeight As Long
    Dim lngMax As Long
    
    With UserControl
        lngHeight = .picClient.CurrentY + mlngTextHeight
        If .picClient.Height < lngHeight Then
            .picClient.Height = lngHeight
            mblnOverride = True
            .scrollVertical.Value = 0
            lngMax = (.picClient.Height - .picContainer.Height) \ Screen.TwipsPerPixelY
            If lngMax < 0 Then lngMax = 0
            .scrollVertical.Max = lngMax
            .scrollVertical.Visible = lngMax
            mblnOverride = False
        End If
    End With
End Sub


' ************* CLIPBOARD *************


Private Sub lnkCopy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkCopy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkCopy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    SetCursor LoadCursor(0, CURSOR_HAND)
    With UserControl.lnkCopy(Index)
        Clipboard.Clear
        Clipboard.SetText .Tag
        lngLeft = .Left + (.Width - UserControl.lblTooltip.Width) \ 2
        lngTop = .Top + Int(.Height * 1.25)
    End With
    With UserControl.lblTooltip
        .Move lngLeft, lngTop
        .Visible = True
    End With
    UserControl.tmrTooltip.Enabled = True
End Sub

Private Sub tmrTooltip_Timer()
    UserControl.tmrTooltip.Enabled = False
    UserControl.lblTooltip.Visible = False
End Sub


' ************* LINKS *************


Private Sub lnkLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim enStyle As LinkStyle
    Dim strLink As String
    
    SetCursor LoadCursor(0, CURSOR_HAND)
    With UserControl.lnkLink(Index)
        enStyle = Val(Left(.Tag, 1))
        strLink = Mid(.Tag, 2)
    End With
    Select Case enStyle
        Case lseHelp: ShowHelp strLink
        Case lseURL: xp.OpenURL strLink
        Case lseShard: OpenShard strLink
        Case lseMaterial: OpenMaterial strLink
        Case lseItem: OpenForm "frmItem"
        Case lseSchool: OpenSchool GetSchoolID(strLink)
        Case lseForm: OpenForm strLink
        Case lseCommand: RaiseEvent Click(strLink)
    End Select
End Sub

Private Sub lnkAugment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkAugment_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
    If Index < 1 Or Index > mlngAugments Then Exit Sub
    With mtypAugment(Index)
        OpenAugment .Augment, .Variation, .Scaling
    End With
End Sub

Private Sub lnkAugment_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub


' ************* SCROLLING *************


Private Sub scrollVertical_GotFocus()
    UserControl.picClient.SetFocus
End Sub

Private Sub scrollVertical_Change()
    If Not mblnOverride Then VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    If Not mblnOverride Then VerticalScroll
End Sub

Private Sub VerticalScroll()
    With UserControl
        .picClient.Top = 0 - (.scrollVertical.Value * Screen.TwipsPerPixelY)
    End With
End Sub
