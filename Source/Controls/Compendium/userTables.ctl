VERSION 5.00
Begin VB.UserControl userTables 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   5604
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7716
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5604
   ScaleWidth      =   7716
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   1500
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   3852
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   3312
      Left            =   7080
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4512
      Left            =   300
      ScaleHeight     =   4512
      ScaleWidth      =   6612
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   6612
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3852
         Left            =   240
         ScaleHeight     =   3852
         ScaleWidth      =   5952
         TabIndex        =   1
         Top             =   240
         Width           =   5952
         Begin Compendium.userTable usrtblFred 
            Height          =   1272
            Left            =   3960
            TabIndex        =   5
            Top             =   540
            Width           =   1452
            _ExtentX        =   2561
            _ExtentY        =   2244
         End
         Begin Compendium.userTable usrtblDestiny 
            Height          =   1152
            Left            =   2220
            TabIndex        =   4
            Top             =   480
            Width           =   1452
            _ExtentX        =   2561
            _ExtentY        =   2032
         End
         Begin Compendium.userTable usrtblHeroic 
            Height          =   1272
            Left            =   480
            TabIndex        =   2
            Top             =   420
            Width           =   1392
            _ExtentX        =   2455
            _ExtentY        =   2244
         End
         Begin Compendium.userTable usrtblEpic 
            Height          =   1272
            Left            =   540
            TabIndex        =   3
            Top             =   2100
            Width           =   1392
            _ExtentX        =   2455
            _ExtentY        =   2244
         End
      End
   End
   Begin VB.Label lblControl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "XP"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   228
   End
End
Attribute VB_Name = "userTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mblnLoaded As Boolean
Private mblnOverride As Boolean

Private mlngScrollX As Long
Private mlngScrollY As Long

Private mlngMarginX As Long

Public Property Get Hwnd() As Long
    Hwnd = UserControl.picContainer.Hwnd
End Property

Public Sub Init()
    With UserControl
        .lblControl.Visible = False
        mlngScrollX = .TextWidth("XXX")
        mlngScrollY = .TextHeight("Q") * 2
        mlngMarginX = .TextWidth("   ")
    End With
    Redraw
End Sub

Public Sub Redraw()
    Dim lngLeft As Long
    Dim lngTop As Long
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .picClient.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .picContainer.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        lngTop = UserControl.TextHeight("Q") \ 2
        With .usrtblHeroic
            .TableID = "Heroic"
            .Move lngLeft, lngTop
            lngTop = .Top + .Height
        End With
        With .usrtblEpic
            .TableID = "Epic"
            .Move lngLeft, lngTop - PixelY
        End With
        With .usrtblHeroic
            lngLeft = .Left + .Width + mlngMarginX
            lngTop = .Top
        End With
        With .usrtblDestiny
            .TableID = "Destiny"
            .Move lngLeft, lngTop
            lngLeft = .Left + .Width + mlngMarginX
        End With
        With .usrtblFred
            .TableID = "Fred"
            .Move lngLeft, lngTop
        End With
        .picClient.Move 0, 0, .usrtblFred.Left + .usrtblFred.Width + PixelX, .usrtblEpic.Top + .usrtblEpic.Height + PixelX
        .picContainer.Visible = True
    End With
    mblnLoaded = True
    MoveContainer
    ShowScrollbars
End Sub

Private Sub UserControl_Resize()
    If Not mblnLoaded Then Exit Sub
    MoveContainer
    ShowScrollbars
End Sub

Private Sub MoveContainer()
    Dim lngLeft As Long
    Dim lngWidth As Long
    
    With UserControl
        If .ScaleWidth <= .picClient.Width Then
            lngLeft = 0
        Else
            lngLeft = (.ScaleWidth - .picClient.Width) \ 2
            If lngLeft > mlngMarginX Then lngLeft = mlngMarginX
        End If
        .picContainer.Move lngLeft, 0
    End With
End Sub

Private Sub ShowScrollbars()
    Dim blnVertical As Boolean
    Dim blnHorizontal As Boolean
    Dim lngValue As Long
    Dim lngClient As Long
    
    With UserControl
        If .picContainer.Width < .picClient.Width Then
        End If
        lngClient = .picClient.Width + .picContainer.Left
        blnHorizontal = (.ScaleWidth < lngClient)
        blnVertical = (.ScaleHeight < .picClient.Height)
        If blnHorizontal And Not blnVertical Then
            blnVertical = (.ScaleHeight - .scrollHorizontal.Height < .picClient.Height)
        ElseIf blnVertical And Not blnHorizontal Then
            blnHorizontal = (.ScaleWidth - .scrollVertical.Width < lngClient)
        End If
        If blnHorizontal Then .picContainer.Height = .ScaleHeight - .scrollHorizontal.Height Else .picContainer.Height = .ScaleHeight
        If blnVertical Then .picContainer.Width = .ScaleWidth - .scrollVertical.Width Else .picContainer.Width = .ScaleWidth
        .scrollHorizontal.Move 0, .ScaleHeight - .scrollHorizontal.Height, .picContainer.Width
        .scrollHorizontal.Visible = blnHorizontal
        If blnHorizontal Then
            .scrollHorizontal.Value = 0
            lngValue = lngClient - .picContainer.Width
            .scrollHorizontal.Max = lngValue \ mlngScrollX + 1
            If .picContainer.Width > mlngScrollX Then .scrollHorizontal.LargeChange = .picContainer.Width \ mlngScrollX
        End If
        .scrollVertical.Move .ScaleWidth - .scrollVertical.Width, 0, .scrollVertical.Width, .picContainer.Height
        .scrollVertical.Visible = blnVertical
        If blnVertical Then
            .scrollVertical.Value = 0
            lngValue = .picClient.Height - .picContainer.Height
            .scrollVertical.Max = lngValue \ mlngScrollY + 1
            If .picContainer.Height > mlngScrollY Then .scrollVertical.LargeChange = .picContainer.Height \ mlngScrollY
        End If
    End With
End Sub

Private Sub scrollHorizontal_GotFocus()
    UserControl.picClient.SetFocus
End Sub

Private Sub scrollHorizontal_Change()
    HorizontalScroll
End Sub

Private Sub scrollHorizontal_Scroll()
    HorizontalScroll
End Sub

Private Sub HorizontalScroll()
    With UserControl
        .picClient.Left = 0 - (.scrollHorizontal.Value * mlngScrollX)
    End With
End Sub

Public Sub WheelScroll(ByVal plngValue As Long)
    Dim lngValue As Long
    
    With UserControl.scrollVertical
        lngValue = .Value - plngValue
        If lngValue < 0 Then lngValue = 0
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub

Private Sub scrollVertical_GotFocus()
    UserControl.picClient.SetFocus
End Sub

Private Sub scrollVertical_Change()
    VerticalScroll
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    With UserControl
        .picClient.Top = 0 - (.scrollVertical.Value * mlngScrollY)
    End With
End Sub
