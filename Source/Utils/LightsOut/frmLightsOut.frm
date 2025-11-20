VERSION 5.00
Begin VB.Form frmLightsOut 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solver"
   ClientHeight    =   4164
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3264
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLightsOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4164
   ScaleWidth      =   3264
   Tag             =   "4020"
   Begin LightsOut.userCheckBox usrchkAlwaysOnTop 
      Height          =   252
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1932
      _ExtentX        =   3408
      _ExtentY        =   445
      Caption         =   "Always On Top"
   End
   Begin LightsOut.userCheckBox usrchkAction 
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   445
      Caption         =   "Play"
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   2172
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   960
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   960
   End
   Begin VB.Timer tmrAlwaysOnTop 
      Interval        =   1000
      Left            =   2220
      Top             =   3780
   End
   Begin LightsOut.userCheckBox usrchkAction 
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   445
      Value           =   0   'False
      Caption         =   "Edit"
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2040
      ScaleHeight     =   252
      ScaleWidth      =   132
      TabIndex        =   0
      Top             =   60
      Width           =   132
   End
   Begin VB.Label lnkLink 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   2748
      TabIndex        =   10
      Top             =   3864
      Width           =   384
   End
   Begin VB.Label lnkLink 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Circle"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   4
      Left            =   2652
      TabIndex        =   8
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lnkLink 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "5x5"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   3
      Left            =   2108
      TabIndex        =   7
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lnkLink 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "4x4"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   2
      Left            =   1564
      TabIndex        =   6
      Top             =   480
      Width           =   360
   End
   Begin VB.Label lnkLink 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "3x3"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   1020
      TabIndex        =   5
      Top             =   480
      Width           =   360
   End
   Begin VB.Image imgStepOn 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   4116
      Picture         =   "frmLightsOut.frx":1CFA
      Top             =   132
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgStepOff 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   4116
      Picture         =   "frmLightsOut.frx":383C
      Top             =   732
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgOn 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   3516
      Picture         =   "frmLightsOut.frx":537E
      Top             =   132
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgOff 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   3516
      Picture         =   "frmLightsOut.frx":6EC0
      Top             =   732
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   9
      Left            =   2532
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   8
      Left            =   1932
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   7
      Left            =   1332
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   6
      Left            =   732
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   5
      Left            =   132
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   24
      Left            =   2532
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   23
      Left            =   1932
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   22
      Left            =   1332
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   21
      Left            =   732
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   20
      Left            =   132
      Stretch         =   -1  'True
      Top             =   3180
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   19
      Left            =   2532
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   18
      Left            =   1932
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   17
      Left            =   1332
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   16
      Left            =   732
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   15
      Left            =   132
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   14
      Left            =   2532
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   13
      Left            =   1932
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   12
      Left            =   1332
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   11
      Left            =   732
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   10
      Left            =   132
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   4
      Left            =   2532
      Stretch         =   -1  'True
      Top             =   780
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   3
      Left            =   1932
      Stretch         =   -1  'True
      Top             =   780
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   2
      Left            =   1332
      Stretch         =   -1  'True
      Top             =   780
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   1
      Left            =   732
      Stretch         =   -1  'True
      Top             =   780
      Width           =   600
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   0
      Left            =   132
      Stretch         =   -1  'True
      Top             =   780
      Width           =   600
   End
End
Attribute VB_Name = "frmLightsOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Enum ActiveEnum
    aeActive
    aeInactive
    aeDefault
End Enum

Private Enum LitEnum
    leLit
    leUnlit
    leDefault
End Enum

Private Enum StepEnum
    seStep
    seNoStep
    seDefault
End Enum

Private Enum ActionEnum
    aePlay
    aeLights
    aeLayout
End Enum

Private Type TileType
    Active As Boolean
    Lit As Boolean
    Solve As Boolean
End Type

Private Type CircleType
    X As Long
    Y As Long
End Type

Private mtypGrid(4, 4) As TileType

Private mstrStyle As String
Private menAction As ActionEnum

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    Dim i As Long
    
    cfg.RefreshColors Me
    PositionForm Me, "Solver"
    mstrStyle = "5x5"
    menAction = aePlay
    For i = 0 To 24
        SetTile i, aeActive, leLit, seNoStep
    Next
    SetAlwaysOnTop Me.hWnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteCoords "Solver", Me.Left, Me.Top
    CloseApp
End Sub

Private Sub usrchkAlwaysOnTop_UserChange()
    Me.tmrAlwaysOnTop.Enabled = Me.usrchkAlwaysOnTop.Value
    SetAlwaysOnTop Me.hWnd, Me.usrchkAlwaysOnTop.Value
End Sub

Private Sub tmrAlwaysOnTop_Timer()
    Me.tmrAlwaysOnTop.Enabled = False
    SetAlwaysOnTop Me.hWnd, True
    Me.tmrAlwaysOnTop.Enabled = True
End Sub


' ************* CONTROLS *************


Private Sub usrchkAction_UserChange(Index As Integer)
    Dim i As Long
    
    Me.usrchkAction(Index).Value = True
    Me.usrchkAction(1 - Index).Value = False
    If Index = 0 Then
        menAction = aePlay
    Else
        menAction = aeLights
        For i = 0 To 24
            SetTile i, aeDefault, leDefault, seNoStep
        Next
    End If
End Sub

Private Sub lnkLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub lnkLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub lnkLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
    Select Case Me.lnkLink(Index).Caption
        Case "Help": ShowHelp "Lights_Out"
        Case Else: SetStyle Me.lnkLink(Index).Caption
    End Select
End Sub

Private Sub SetStyle(pstrStyle As String)
    Dim i As Long
    Dim lngX As Long
    Dim lngY As Long
    
    mstrStyle = pstrStyle
    Select Case mstrStyle
        Case "3x3"
            For i = 0 To 24
                IndexToCoords i, lngX, lngY
                With mtypGrid(lngX, lngY)
                    .Active = (lngX > 0 And lngY > 0 And lngX < 4 And lngY < 4)
                    .Lit = (lngX > 0 And lngY > 0 And lngX < 4 And lngY < 4)
                End With
                SetTile i, aeDefault, leDefault, seNoStep
            Next
        Case "4x4"
            For i = 0 To 24
                IndexToCoords i, lngX, lngY
                With mtypGrid(lngX, lngY)
                    .Active = (lngX < 4 And lngY < 4)
                    .Lit = (lngX < 4 And lngY < 4)
                End With
                SetTile i, aeDefault, leDefault, seNoStep
            Next
        Case "5x5"
            For i = 0 To 24
                IndexToCoords i, lngX, lngY
                With mtypGrid(lngX, lngY)
                    .Active = True
                    .Lit = True
                End With
                SetTile i, aeDefault, leDefault, seNoStep
            Next
        Case "Circle"
            For i = 0 To 24
                SetTile i, aeInactive, leUnlit, seNoStep
            Next
            SetTile 1, aeActive, leLit, seNoStep
            SetTile 2, aeActive, leLit, seNoStep
            SetTile 5, aeActive, leLit, seNoStep
            SetTile 8, aeActive, leLit, seNoStep
            SetTile 10, aeActive, leLit, seNoStep
            SetTile 13, aeActive, leLit, seNoStep
            SetTile 16, aeActive, leLit, seNoStep
            SetTile 17, aeActive, leLit, seNoStep
    End Select
End Sub

Private Sub chkButton_Click(Index As Integer)
    If UncheckButton(Me.chkButton(Index), mblnOverride) Then Exit Sub
    Select Case Me.chkButton(Index).Caption
        Case "Shuffle": Shuffle
        Case "Solve": SolvePuzzle
    End Select
    Me.picFocus.SetFocus
End Sub

Private Sub Shuffle()
    Dim lngX As Long
    Dim lngY As Long
    Dim i As Long
    
    For i = 0 To 24
        IndexToCoords i, lngX, lngY
        mtypGrid(lngX, lngY).Lit = mtypGrid(lngX, lngY).Active
        mtypGrid(lngX, lngY).Solve = False
    Next
    For i = 1 To Int(91 * Rnd + 10)
        IndexToCoords Int(25 * Rnd), lngX, lngY
        If mtypGrid(lngX, lngY).Active Then ToggleLight lngX, lngY, True, False
    Next
    For i = 0 To 24
        SetTile i, aeDefault, leDefault, seDefault
    Next
End Sub

Private Sub SolvePuzzle()
    Dim typTemp(4, 4) As TileType
    Dim blnSolved As Boolean
    Dim lngX As Long
    Dim lngY As Long
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
    For i = 0 To 24
        IndexToCoords i, lngX, lngY
        mtypGrid(lngX, lngY).Solve = False
        typTemp(lngX, lngY) = mtypGrid(lngX, lngY)
    Next
    Select Case mstrStyle
        Case "3x3": blnSolved = Solve3x3()
        Case "4x4": blnSolved = Solve4x4()
        Case "5x5": blnSolved = Solve5x5()
        Case "Circle": blnSolved = SolveCircle()
    End Select
    For i = 0 To 24
        IndexToCoords i, lngX, lngY
        If blnSolved Then typTemp(lngX, lngY).Solve = mtypGrid(lngX, lngY).Solve
        mtypGrid(lngX, lngY) = typTemp(lngX, lngY)
        SetTile i, aeDefault, leDefault, seDefault
    Next
    Screen.MousePointer = vbDefault
    If Not blnSolved Then MsgBox "No solution.", vbInformation, "Notice"
End Sub


' ************* PLAY *************


Private Sub img_Click(Index As Integer)
    Dim lngX As Long
    Dim lngY As Long
    
    IndexToCoords Index, lngX, lngY
    Select Case menAction
        Case aePlay: ToggleLight lngX, lngY, True
        Case aeLights: ToggleLight lngX, lngY, False
    End Select
End Sub

Private Function IndexToCoords(ByVal plngIndex As Long, plngX As Long, plngY As Long)
    plngX = plngIndex Mod 5
    plngY = plngIndex \ 5
End Function

Private Function CoordsToIndex(plngX As Long, plngY As Long) As Long
    CoordsToIndex = plngY * 5 + plngX
End Function

Private Sub ToggleLight(plngX As Long, plngY As Long, pblnPlay As Boolean, Optional pblnDraw As Boolean = True)
    Dim lngIndex As Long
    
    If plngX < 0 Or plngX > 4 Or plngY < 0 Or plngY > 4 Then Exit Sub
    lngIndex = CoordsToIndex(plngX, plngY)
    With mtypGrid(plngX, plngY)
        If .Active Then
            .Lit = Not .Lit
            If pblnDraw Then SetTile lngIndex, aeDefault, leDefault, seDefault
        End If
    End With
    If pblnPlay Then
        ToggleLight plngX - 1, plngY, False
        ToggleLight plngX + 1, plngY, False
        ToggleLight plngX, plngY - 1, False
        ToggleLight plngX, plngY + 1, False
        If mstrStyle = "Circle" Then
            ToggleLight plngX - 1, plngY - 1, False
            ToggleLight plngX + 1, plngY - 1, False
            ToggleLight plngX - 1, plngY + 1, False
            ToggleLight plngX + 1, plngY + 1, False
        End If
    End If
End Sub

Private Sub SetTile(plngIndex As Long, penActive As ActiveEnum, penLit As LitEnum, penStep As StepEnum)
    Dim lngX As Long
    Dim lngY As Long
    Dim strImage As String
    
    IndexToCoords plngIndex, lngX, lngY
    With mtypGrid(lngX, lngY)
        ' Image
        If penLit <> leDefault Then .Lit = (penLit = leLit)
        If penStep <> seDefault Then .Solve = (penStep = seStep)
        strImage = "img"
        If .Solve Then strImage = strImage & "Step"
        If .Lit Then strImage = strImage & "On" Else strImage = strImage & "Off"
        If penActive <> aeDefault Then .Active = (penActive = aeActive)
        If Not .Active Then strImage = ""
        If Me.img(plngIndex).Tag <> strImage Then
            Me.img(plngIndex).Tag = strImage
            If .Active Then
                Set Me.img(plngIndex).Picture = Me(strImage).Picture
            Else
                Set Me.img(plngIndex).Picture = LoadPicture
            End If
        End If
    End With
End Sub


' ************* SOLVE *************


Private Function Solve3x3() As Boolean
    Dim lngX As Long
    Dim i As Long
    
    SolveDown 1, 3, 1, 3
    For lngX = 1 To 3
        If Not mtypGrid(lngX, 3).Lit Then
            For i = lngX - 1 To lngX + 1
                If mtypGrid(i, 1).Active Then ToggleSolve i, 1
            Next
        End If
    Next
    SolveDown 1, 3, 1, 3
    Solve3x3 = CheckSolution()
End Function

Private Function Solve4x4() As Boolean
    SolveDown 0, 3, 0, 3
    Solve4x4 = CheckSolution()
End Function

Private Function Solve5x5() As Boolean
    Dim lngX As Long
    Dim i As Long
    
    SolveDown 0, 4, 0, 4
    For lngX = 0 To 2
        If Not mtypGrid(lngX, 4).Lit Then
            For i = lngX - 1 To lngX + 1
                If i >= 0 And i <= 2 Then
                    If mtypGrid(i, 0).Active Then ToggleSolve i, 0
                End If
            Next
        End If
    Next
    SolveDown 0, 4, 0, 4
    Solve5x5 = CheckSolution()
End Function

Private Function SolveCircle() As Boolean
    Dim lngMap(7) As CircleType
    Dim i As Long
    
    ' All possibilities solvable
    SolveCircle = True
    ' Remap circle tiles into a linear array
    SetMap lngMap(0), 1, 0
    SetMap lngMap(1), 2, 0
    SetMap lngMap(2), 3, 1
    SetMap lngMap(3), 3, 2
    SetMap lngMap(4), 2, 3
    SetMap lngMap(5), 1, 3
    SetMap lngMap(6), 0, 2
    SetMap lngMap(7), 0, 1
    ' Chase down lights to 1 unlit or 2 adjacent unlit
    For i = 0 To 5
        If Not mtypGrid(lngMap(i).X, lngMap(i).Y).Lit Then ToggleSolve lngMap(i + 1).X, lngMap(i + 1).Y
    Next
    ' Check if already solved
    If mtypGrid(0, 2).Lit And mtypGrid(0, 1).Lit Then Exit Function
    ' If last 2 tiles unlit, hit one to leave only one unlit
    If Not (mtypGrid(0, 2).Lit Or mtypGrid(0, 1).Lit) Then ToggleSolve 0, 1
    ' Now only 1 light unlit; identify which one
    Select Case False
        Case mtypGrid(0, 2).Lit: i = 6
        Case mtypGrid(0, 1).Lit: i = 7
        Case mtypGrid(1, 0).Lit: i = 0
    End Select
    ' To solve with 1 unlit, toggle that one, skip one, toggle 2, skip 1, toggle 2
    ToggleSolve lngMap(i).X, lngMap(i).Y
    i = i + 2: If i > 7 Then i = i - 8
    ToggleSolve lngMap(i).X, lngMap(i).Y
    i = i + 1: ToggleSolve lngMap(i).X, lngMap(i).Y
    i = i + 2: ToggleSolve lngMap(i).X, lngMap(i).Y
    i = i + 1: ToggleSolve lngMap(i).X, lngMap(i).Y
    ' Puzzle is now solved
End Function

Private Sub SolveDown(plngX1 As Long, plngX2 As Long, plngY1 As Long, plngY2 As Long)
    Dim lngX As Long
    Dim lngY As Long
    
    For lngY = plngY1 To plngY2 - 1
        For lngX = plngX1 To plngX2
            If Not mtypGrid(lngX, lngY).Lit Then ToggleSolve lngX, lngY + 1
        Next
    Next
End Sub

Private Sub SetMap(plngMap As CircleType, plngX As Long, plngY As Long)
    plngMap.X = plngX
    plngMap.Y = plngY
End Sub

Private Sub ToggleSolve(plngX As Long, plngY As Long)
    mtypGrid(plngX, plngY).Solve = Not mtypGrid(plngX, plngY).Solve
    ToggleLight plngX, plngY, True, False
End Sub

Private Function CheckSolution() As Boolean
    Dim lngX As Long
    Dim lngY As Long
    
    For lngX = 0 To 4
        For lngY = 0 To 4
            With mtypGrid(lngX, lngY)
                If .Active And Not .Lit Then Exit Function
            End With
        Next
    Next
    CheckSolution = True
End Function
