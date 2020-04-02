VERSION 5.00
Begin VB.UserControl userStats 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1752
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   1752
   Begin VB.PictureBox picSpent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   720
      ScaleHeight     =   228
      ScaleWidth      =   468
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1920
      Width           =   492
   End
   Begin VB.Timer tmrRepeat 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1260
      Top             =   1800
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   445
      Min             =   6
      Max             =   20
      Value           =   8
      AllowJump       =   0   'False
      AllowTyping     =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   1
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   1800
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   252
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   445
      Min             =   6
      Max             =   20
      Value           =   8
      AllowJump       =   0   'False
      AllowTyping     =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   2
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   252
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   445
      Min             =   6
      Max             =   20
      Value           =   8
      AllowJump       =   0   'False
      AllowTyping     =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   2
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   252
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   1080
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   445
      Min             =   6
      Max             =   20
      Value           =   8
      AllowJump       =   0   'False
      AllowTyping     =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   2
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   252
      Index           =   5
      Left            =   480
      TabIndex        =   14
      Top             =   1320
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   445
      Min             =   6
      Max             =   20
      Value           =   8
      AllowJump       =   0   'False
      AllowTyping     =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   2
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   252
      Index           =   6
      Left            =   480
      TabIndex        =   17
      Top             =   1560
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   445
      Min             =   6
      Max             =   20
      Value           =   8
      AllowJump       =   0   'False
      AllowTyping     =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   3
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin CharacterBuilderLite.userCheckBox usrchkInclude 
      Height          =   252
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   445
      Caption         =   "Adventurer"
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   6
      Left            =   1500
      TabIndex        =   18
      Top             =   1596
      Width           =   204
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   5
      Left            =   1500
      TabIndex        =   15
      Top             =   1356
      Width           =   204
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   4
      Left            =   1500
      TabIndex        =   12
      Top             =   1116
      Width           =   204
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   3
      Left            =   1500
      TabIndex        =   9
      Top             =   876
      Width           =   204
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   2
      Left            =   1500
      TabIndex        =   6
      Top             =   636
      Width           =   204
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "+3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   1
      Left            =   1500
      TabIndex        =   3
      Top             =   396
      Width           =   204
   End
   Begin VB.Label lblAbility 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Str"
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   492
   End
   Begin VB.Label lblAbility 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dex"
      Height          =   252
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   492
   End
   Begin VB.Label lblAbility 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Con"
      Height          =   252
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   492
   End
   Begin VB.Label lblAbility 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Int"
      Height          =   252
      Index           =   4
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label lblAbility 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wis"
      Height          =   252
      Index           =   5
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   492
   End
   Begin VB.Label lblAbility 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cha"
      Height          =   252
      Index           =   6
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   492
   End
End
Attribute VB_Name = "userStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Public Event StatChange(Stat As StatEnum, Increase As Boolean)
Public Event Include(Include As Boolean)

Private menBuildPoints As BuildPointsEnum
Private mlngBase(6) As Long
Private mlngMax As Long

Private mlngCurrentIndex As Long
Private mblnCurrentDirection As Boolean

Private mlngBackColor As Long
Private mlngBorderColor As Long
Private mlngTextColor As Long
Private mlngLeft As Long

Private mblnOverride As Boolean

Private Sub UserControl_Initialize()
    mlngTextColor = vbWindowText
    mlngBackColor = vbWindowBackground
End Sub


' ************* GENERAL *************


Public Sub RefreshColors()
    Dim i As Long
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .usrchkInclude.RefreshColors
        For i = 1 To 6
            cfg.ApplyColors .lblAbility(i), cgeWorkspace
            cfg.ApplyColors .usrSpinner(i), cgeControls
            .lblCost(i).ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
            .lblCost(i).BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        Next
        mlngTextColor = cfg.GetColor(cgeControls, cveText)
        mlngBackColor = cfg.GetColor(cgeControls, cveBackground)
        mlngBorderColor = cfg.GetColor(cgeControls, cveBorderExterior)
        ShowSpent
    End With
End Sub

Public Sub Refresh()
    Dim i As Long
    
    build.StatPoints(menBuildPoints, 0) = 0
    For i = 1 To 6
        mlngBase(i) = db.Race(build.Race).Stats(i)
        build.StatPoints(menBuildPoints, 0) = build.StatPoints(menBuildPoints, 0) + build.StatPoints(menBuildPoints, i)
    Next
    ShowPoints
End Sub

Private Sub ShowPoints()
    Dim blnVisible As Boolean
    Dim i As Long
    
    mblnOverride = True
    blnVisible = (build.IncludePoints(menBuildPoints) = 1)
    With UserControl
        .usrchkInclude.Value = blnVisible
        For i = 1 To 6
            .usrSpinner(i).Override = CalculateBaseStat(mlngBase(i), build.StatPoints(menBuildPoints, i))
            ShowCost i
            .lblAbility(i).Visible = blnVisible
            .usrSpinner(i).Visible = blnVisible
            .lblCost(i).Visible = blnVisible
        Next
        .picSpent.Visible = blnVisible
        ShowSpent
        If .usrchkInclude.Enabled And build.Race = reDrow And menBuildPoints = beChampion Then
            .usrchkInclude.Enabled = Not (build.Race = reDrow)
            RaiseEvent Include(False)
        End If
    End With
    mblnOverride = False
End Sub

Private Sub ShowCost(penStat As StatEnum)
    Dim lngCost As Long
    
    With UserControl.lblCost(penStat)
        lngCost = NextChange(penStat)
        If lngCost = 0 Then .Caption = "-" Else .Caption = "+" & lngCost
    End With
End Sub

Private Sub ShowSpent()
    Dim strSpent As String
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngPixelX As Long
    Dim lngPixelY As Long
    
    lngPixelX = Screen.TwipsPerPixelX
    lngPixelY = Screen.TwipsPerPixelY
    strSpent = build.StatPoints(menBuildPoints, 0)
    With UserControl.picSpent
        lngWidth = .ScaleWidth
        lngHeight = .ScaleHeight
        .ForeColor = mlngTextColor
        .BackColor = mlngBackColor
        .FillColor = mlngBackColor
        .BorderStyle = vbBSNone
        UserControl.picSpent.Line (0, 0)-(lngWidth - lngPixelX, lngHeight - lngPixelY), mlngBorderColor, B
        .CurrentX = (lngWidth - .TextWidth(strSpent)) \ 2
        .CurrentY = (lngHeight - .TextHeight(strSpent)) \ 2 - lngPixelY
    End With
    UserControl.picSpent.Print strSpent
End Sub

Public Property Get BuildPoints() As BuildPointsEnum
    BuildPoints = menBuildPoints
End Property

Public Property Let BuildPoints(ByVal penBuildPoints As BuildPointsEnum)
    menBuildPoints = penBuildPoints
    mlngMax = GetBuildPoints(menBuildPoints)
    Select Case menBuildPoints
        Case beAdventurer: SetIncludeCaption "Adventurer"
        Case beChampion: SetIncludeCaption "Champion"
        Case beHero: SetIncludeCaption "Hero"
        Case beLegend: SetIncludeCaption "Legend"
    End Select
End Property

Private Sub SetIncludeCaption(pstrCaption As String)
    Dim lngLeft As Long
    
    With UserControl
        .usrchkInclude.Caption = pstrCaption
        If mlngLeft = 0 Then mlngLeft = .usrchkInclude.Left
        .usrchkInclude.Left = mlngLeft + (.TextWidth("Adventurer") - .TextWidth(pstrCaption)) \ 2
    End With
End Sub

Private Sub usrchkInclude_UserChange()
    If UserControl.usrchkInclude.Value Then build.IncludePoints(menBuildPoints) = 1 Else build.IncludePoints(menBuildPoints) = 0
    RaiseEvent Include(UserControl.usrchkInclude.Value)
    CascadeChanges cceStats
    SetDirty
    ShowPoints
End Sub


' ************* SPINNERS *************


Private Sub usrSpinner_RequestChange(Index As Integer, OldValue As Long, NewValue As Long, Cancel As Boolean)
    If NewValue > OldValue Then Cancel = Not Increase(Index) Else Cancel = Not Decrease(Index)
End Sub

Private Function Increase(ByVal penStat As StatEnum) As Boolean
    Dim lngChange As Long
    
    lngChange = NextChange(penStat)
    If lngChange = 0 Or build.StatPoints(menBuildPoints, 0) + lngChange > mlngMax Then Exit Function
    Increase = True
    ChangePoints penStat, lngChange
End Function

Private Function NextChange(ByVal penStat As StatEnum) As Long
    Select Case build.StatPoints(menBuildPoints, penStat)
        Case 0 To 5: NextChange = 1
        Case 6, 8: NextChange = 2
        Case 10, 13: NextChange = 3
        Case Else: NextChange = 0
    End Select
End Function

Private Function Decrease(ByVal penStat As StatEnum) As Boolean
    Dim lngChange As Long
    
    Select Case build.StatPoints(menBuildPoints, penStat)
        Case 1 To 6: lngChange = -1
        Case 8, 10: lngChange = -2
        Case 13, 16: lngChange = -3
        Case Else: Exit Function
    End Select
    Decrease = True
    ChangePoints penStat, lngChange
End Function

Private Function ChangePoints(penStat As StatEnum, plngChange As Long) As Boolean
    ' Individual stat
    build.StatPoints(menBuildPoints, penStat) = build.StatPoints(menBuildPoints, penStat) + plngChange
    UserControl.usrSpinner(penStat).Override = CalculateBaseStat(mlngBase(penStat), build.StatPoints(menBuildPoints, penStat))
    ShowCost penStat
    ' Total spent
    build.StatPoints(menBuildPoints, 0) = build.StatPoints(menBuildPoints, 0) + plngChange
    ShowSpent
    ' Propagate
    If build.BuildPoints = menBuildPoints Then RaiseEvent StatChange(penStat, (plngChange > 0))
    If Not mblnOverride Then
        CascadeChanges cceStats
        SetDirty
    End If
End Function

' Call that handles propagation
Public Sub IncrementStat(penStat As StatEnum, pblnIncrement As Boolean)
    mblnOverride = True
    If pblnIncrement Then Increase penStat Else Decrease penStat
    mblnOverride = False
End Sub
