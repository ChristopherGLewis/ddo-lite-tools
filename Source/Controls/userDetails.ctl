VERSION 5.00
Begin VB.UserControl userDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3180
   ScaleWidth      =   3240
   Begin VB.VScrollBar scrollVertical 
      Height          =   2052
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   252
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2772
      Left            =   0
      ScaleHeight     =   2775
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   0
      Width           =   2292
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Wiki"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC4C4&
         Height          =   216
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   360
      End
   End
End
Attribute VB_Name = "userDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Type DetailType
    Text As String
    Link As String
    HasError As Boolean
End Type

Private Type DetailLineType
    Text As String
    HasError As Boolean
End Type

Private mtypLine() As DetailLineType
Private mlngLines As Long
Private mlngWikiLine As Long
Private mlngWidth As Long
Private mlngSpace As Long

Private mtypDetail() As DetailType
Private mlngDetails As Long
Private mlngLinkIndex As Long
Private mlngMarginX As Long
Private mlngMarginY As Long

Private mblnNoFocus As Boolean


' ************* INITIALIZE *************


Private Sub UserControl_Resize()
    UserControl.scrollVertical.Value = 0
    DrawText
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lblLink_Click()
    If mlngLinkIndex = 0 Then Exit Sub
    If Len(mtypDetail(mlngLinkIndex).Link) = 0 Then Exit Sub
    xp.OpenURL mtypDetail(mlngLinkIndex).Link
End Sub


' ************* METHODS *************


Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub AddText(pstrText As String)
    mlngDetails = mlngDetails + 1
    ReDim Preserve mtypDetail(1 To mlngDetails)
    mtypDetail(mlngDetails).Text = pstrText
End Sub

Public Sub AddErrorText(pstrText As String)
    mlngDetails = mlngDetails + 2
    ReDim Preserve mtypDetail(1 To mlngDetails)
    mtypDetail(mlngDetails).Text = pstrText
    mtypDetail(mlngDetails).HasError = True
End Sub

Public Sub AddDescrip(ByVal pstrDescrip As String, pstrLink As String)
    Dim strLine() As String
    Dim lngLines As Long
    Dim lngPos As Long
    Dim i As Long
    
    strLine = Split(pstrDescrip, "}")
    lngLines = UBound(strLine)
    ReDim Preserve mtypDetail(1 To mlngDetails + lngLines + 2)
    For i = 0 To lngLines
        mlngDetails = mlngDetails + 1
        strLine(i) = Trim$(strLine(i))
        lngPos = InStr(strLine(i), "{")
        If lngPos Then strLine(i) = " " & Mid$(strLine(i), lngPos + 1, 1) & " " & Trim$(Mid$(strLine(i), lngPos + 2))
        If Len(pstrLink) And i = lngLines Then
            'Add wiki link
            strLine(i) = strLine(i) & " (Wiki)"
            mlngLinkIndex = mlngDetails
        End If
        mtypDetail(mlngDetails).Text = strLine(i)
        mtypDetail(mlngDetails).Link = pstrLink
    Next
    mlngDetails = mlngDetails + 1
End Sub

Public Sub Clear()
    mlngDetails = 0
    mlngLinkIndex = 0
    Erase mtypDetail
    UserControl.scrollVertical.Value = 0
    DrawText
 End Sub

Public Sub Refresh()
    DrawText
End Sub

Public Sub RefreshColors()
    With UserControl
        .BackColor = cfg.GetColor(cgeControls, cveBackground)
        With .picClient
            .ForeColor = cfg.GetColor(cgeControls, cveText)
            .BackColor = cfg.GetColor(cgeControls, cveBackground)
        End With
        .lblLink.ForeColor = cfg.GetColor(cgeControls, cveTextLink)
        .lblLink.BackColor = cfg.GetColor(cgeControls, cveBackground)
    End With
    DrawText
End Sub

Public Property Get NoFocus() As Boolean
    NoFocus = mblnNoFocus
End Property

Public Property Let NoFocus(ByVal pblnNoFocus As Boolean)
    mblnNoFocus = pblnNoFocus
End Property


' ************* DRAW TEXT *************


Private Sub DrawText()
    Dim lngHeight As Long
    Dim lngFirst As Long
    Dim i As Long
    
    ' Remove starting blank lines
    If mlngDetails > 0 Then
        For lngFirst = 1 To mlngDetails
            If Len(mtypDetail(lngFirst).Text) > 0 Then Exit For
        Next
        If lngFirst > 1 Then
            For i = lngFirst To mlngDetails
                mtypDetail(i - lngFirst + 1) = mtypDetail(i)
            Next
            mlngDetails = mlngDetails - lngFirst + 1
            If mlngDetails < 1 Then
                mlngDetails = 0
                Erase mtypDetail
            Else
                ReDim Preserve mtypDetail(1 To mlngDetails)
            End If
        End If
    End If
    ' Remove trailing blank lines
    If mlngDetails > 0 Then
        Do While Len(mtypDetail(mlngDetails).Text) = 0
            mlngDetails = mlngDetails - 1
            If mlngDetails = 0 Then
                Erase mtypDetail
                Exit Do
            Else
                ReDim Preserve mtypDetail(1 To mlngDetails)
            End If
        Loop
    End If
    With UserControl
        .picClient.Cls
        .lblLink.Visible = False
        .scrollVertical.Visible = False
        ' Size for no scrollbar
        .picClient.Move 0, 0, .ScaleWidth, .ScaleHeight
        mlngMarginX = Screen.TwipsPerPixelX * 5
        mlngMarginY = Screen.TwipsPerPixelY * 2
        mlngWidth = .picClient.ScaleWidth - mlngMarginX * 2
        mlngSpace = .picClient.TextWidth(" ")
        lngHeight = .picClient.TextHeight("X")
        ProcessLines
        If (mlngLines * .TextHeight("X")) + (mlngMarginY * 2) > .ScaleHeight Then
            ' Doesn't fit, add scrollbar
            .picClient.Width = .ScaleWidth - .scrollVertical.Width
            mlngWidth = mlngWidth - .scrollVertical.Width
            ProcessLines
        End If
        .picClient.Height = lngHeight * mlngLines + mlngMarginY * 2
        ShowScrollbar
        ' Draw text
        For i = 1 To mlngLines
            .picClient.CurrentX = mlngMarginX
            .picClient.CurrentY = mlngMarginY + lngHeight * (i - 1)
            If mtypLine(i).HasError Then .picClient.ForeColor = cfg.GetColor(cgeControls, cveTextError)
            UserControl.picClient.Print mtypLine(i).Text;
            If mtypLine(i).HasError Then .picClient.ForeColor = cfg.GetColor(cgeControls, cveText)
            'Move the WIKI link lable ON TOP OF the '(Wiki)' text in the desc.
            'This means that we can have only one link...
            If i = mlngWikiLine Then
                .lblLink.Move .picClient.CurrentX - .picClient.TextWidth("Wiki)"), .picClient.CurrentY
                .lblLink.Visible = True
            End If
        Next
    End With
End Sub

Private Sub ProcessLines()
    Dim i As Long
    
    Erase mtypLine
    mlngLines = 0
    mlngWikiLine = 0
    For i = 1 To mlngDetails
        With mtypDetail(i)
            ProcessLine .Text, .HasError
            If Len(.Link) Then mlngWikiLine = mlngLines
        End With
    Next
End Sub

Private Sub ProcessLine(ByVal pstrText As String, pblnError As Boolean)
    Dim strWord() As String
    Dim lngWords As Long
    Dim lngWord As Long
    Dim lngWordWidth As Long
    Dim lngWidth As Long
    Dim lngStart As Long
    Dim lngLength As Long
    Dim blnIndent As Boolean
    
    If Left$(pstrText, 1) = " " Then blnIndent = True
    strWord = Split(pstrText, " ")
    lngWords = UBound(strWord)
    lngStart = 1
    Do Until lngWord > lngWords
        lngWordWidth = UserControl.picClient.TextWidth(strWord(lngWord))
        Do While lngWordWidth > mlngWidth
            strWord(lngWord) = Left$(strWord(lngWord), Len(strWord(lngWord)) - 1)
            lngWordWidth = UserControl.picClient.TextWidth(strWord(lngWord))
        Loop
        If lngWidth + mlngSpace + lngWordWidth > mlngWidth Then
            NextLine pstrText, lngStart, lngLength, blnIndent, pblnError
            lngStart = lngStart + lngLength
            If blnIndent Then lngWidth = mlngSpace * 2 Else lngWidth = 0
            lngLength = 0
        Else
            lngLength = lngLength + Len(strWord(lngWord)) + 1
            lngWidth = lngWidth + lngWordWidth + mlngSpace
            lngWord = lngWord + 1
        End If
    Loop
    NextLine pstrText, lngStart, lngLength, blnIndent, pblnError
End Sub

Private Sub NextLine(pstrText As String, plngStart As Long, plngLength As Long, pblnIndent As Boolean, pblnError As Boolean)
    mlngLines = mlngLines + 1
    ReDim Preserve mtypLine(1 To mlngLines)
    If pblnIndent And plngStart > 1 Then
        mtypLine(mlngLines).Text = "   " & Mid$(pstrText, plngStart, plngLength)
    Else
        mtypLine(mlngLines).Text = Mid$(pstrText, plngStart, plngLength)
    End If
    mtypLine(mlngLines).HasError = pblnError
End Sub


' ************* SCROLLBAR *************


Private Sub ShowScrollbar()
    With UserControl
        If .picClient.Height > .ScaleHeight Then
            .scrollVertical.Move .ScaleWidth - .scrollVertical.Width, 0, .scrollVertical.Width, .ScaleHeight
            .scrollVertical.Max = (.picClient.Height - .ScaleHeight) \ Screen.TwipsPerPixelY
            .scrollVertical.SmallChange = .TextHeight("X") \ Screen.TwipsPerPixelY
            .scrollVertical.LargeChange = .ScaleHeight \ Screen.TwipsPerPixelY
            .scrollVertical.Visible = True
        Else
            .scrollVertical.Visible = False
        End If
    End With
End Sub

Private Sub scrollVertical_GotFocus()
    If Not mblnNoFocus Then UserControl.picClient.SetFocus
End Sub

Private Sub scrollVertical_Change()
    VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    With UserControl
        .picClient.Top = 0 - .scrollVertical.Value * Screen.TwipsPerPixelY
    End With
End Sub

Public Sub Scroll(plngValue As Long)
    Dim lngIncrement As Long
    Dim lngValue As Long
    
    If Not UserControl.scrollVertical.Visible Then Exit Sub
    lngIncrement = plngValue * (UserControl.TextHeight("Q") \ Screen.TwipsPerPixelY)
    With UserControl.scrollVertical
        lngValue = .Value - lngIncrement
        If lngValue < .Min Then lngValue = .Min
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub

