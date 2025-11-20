VERSION 5.00
Begin VB.Form frmADQ 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Against the Demon Queen"
   ClientHeight    =   2004
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmADQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   Begin VB.Timer tmrAlwaysOnTop 
      Interval        =   1000
      Left            =   5580
      Top             =   1200
   End
   Begin VB.Image imgHelp 
      Height          =   216
      Left            =   4620
      Picture         =   "frmADQ.frx":57E2
      Top             =   1572
      Width           =   420
   End
   Begin VB.Image imgAccept 
      Height          =   216
      Left            =   3408
      Picture         =   "frmADQ.frx":5FBC
      Top             =   1572
      Width           =   600
   End
   Begin VB.Image imgReset 
      Height          =   228
      Left            =   2340
      Picture         =   "frmADQ.frx":6AAE
      Top             =   1572
      Width           =   456
   End
   Begin VB.Image imgBlank 
      Height          =   228
      Index           =   6
      Left            =   4260
      Picture         =   "frmADQ.frx":738C
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgBlank 
      Height          =   228
      Index           =   5
      Left            =   1800
      Picture         =   "frmADQ.frx":812A
      Top             =   1008
      Width           =   720
   End
   Begin VB.Image imgBlank 
      Height          =   228
      Index           =   4
      Left            =   1800
      Picture         =   "frmADQ.frx":8EC8
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgBlank 
      Height          =   228
      Index           =   3
      Left            =   3156
      Picture         =   "frmADQ.frx":9C66
      Top             =   432
      Width           =   720
   End
   Begin VB.Image imgBlank 
      Height          =   228
      Index           =   2
      Left            =   1800
      Picture         =   "frmADQ.frx":AA04
      Top             =   432
      Width           =   720
   End
   Begin VB.Image imgLine4 
      Height          =   228
      Left            =   2700
      Picture         =   "frmADQ.frx":B7A2
      Top             =   1008
      Width           =   1800
   End
   Begin VB.Image imgLine3b 
      Height          =   228
      Left            =   5160
      Picture         =   "frmADQ.frx":D970
      Top             =   720
      Width           =   72
   End
   Begin VB.Image imgLine3a 
      Height          =   228
      Left            =   2700
      Picture         =   "frmADQ.frx":DB2E
      Top             =   720
      Width           =   1248
   End
   Begin VB.Image imgLine2b 
      Height          =   228
      Left            =   4056
      Picture         =   "frmADQ.frx":F298
      Top             =   432
      Width           =   1008
   End
   Begin VB.Image imgLine2a 
      Height          =   228
      Left            =   2700
      Picture         =   "frmADQ.frx":1058E
      Top             =   432
      Width           =   360
   End
   Begin VB.Image imgLine1 
      Height          =   228
      Left            =   2700
      Picture         =   "frmADQ.frx":10CA4
      Top             =   156
      Width           =   1872
   End
   Begin VB.Image imgBlank 
      Height          =   228
      Index           =   1
      Left            =   1800
      Picture         =   "frmADQ.frx":12FA2
      Top             =   156
      Width           =   720
   End
   Begin VB.Image imgKeyword 
      Height          =   228
      Index           =   6
      Left            =   156
      Picture         =   "frmADQ.frx":13D40
      Top             =   1572
      Width           =   972
   End
   Begin VB.Image imgKeyword 
      Height          =   228
      Index           =   5
      Left            =   156
      Picture         =   "frmADQ.frx":14F9E
      Top             =   1296
      Width           =   648
   End
   Begin VB.Image imgKeyword 
      Height          =   228
      Index           =   4
      Left            =   156
      Picture         =   "frmADQ.frx":15C0C
      Top             =   1008
      Width           =   948
   End
   Begin VB.Image imgKeyword 
      Height          =   228
      Index           =   3
      Left            =   156
      Picture         =   "frmADQ.frx":16E1E
      Top             =   720
      Width           =   828
   End
   Begin VB.Image imgKeyword 
      Height          =   228
      Index           =   2
      Left            =   156
      Picture         =   "frmADQ.frx":17DD0
      Top             =   432
      Width           =   972
   End
   Begin VB.Image imgKeyword 
      Height          =   228
      Index           =   1
      Left            =   156
      Picture         =   "frmADQ.frx":1902E
      Top             =   156
      Width           =   828
   End
End
Attribute VB_Name = "frmADQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Const ADQ_Solutions As Long = 23

Private Enum KeywordEnum
    keDevious = 1
    keGrasping
    keHungry
    keMockery
    keNight
    kePoisoner
End Enum

Private Enum CrestEnum
    ceBat = 1
    ceMonkey
    ceOctopus
    ceScorpion
    ceSnake
    ceWolf
End Enum

Private Type ADQType
    Keyword(1 To 6) As KeywordEnum
    Crest(1 To 6) As CrestEnum
    Valid As Boolean ' Calculated on the fly
End Type

Private mtypADQ(ADQ_Solutions) As ADQType
Private mlngArray(1 To 6) As Long ' Holds the current state of the puzzle
Private mblnValid(1 To 6) As Boolean
Private mlngCurrent As Long ' Dual-use
Private mblnAccept As Boolean
Private mlngLeft As Long
Private mlngTop As Long
Private mblnComplete As Boolean

' ************* FORM *************


Private Sub Form_Load()
    mblnAccept = False
    Unload frmADQ2
    LoadData
    ResetChoices
    DrawScreen
    PositionForm Me, "ADQ1"
    SetAlwaysOnTop Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteCoords "ADQ1", Me.Left, Me.Top
    ResetChoices
    If Not mblnAccept Then CloseApp
End Sub

Private Sub tmrAlwaysOnTop_Timer()
    SetAlwaysOnTop Me.hWnd, True
End Sub


' ************* DATA *************


Private Sub LoadData()
    mlngCurrent = 0
    AddSolution "Devious Mockery Grasping Poisoner Hungry Night", "Snake Octopus Wolf Scorpion Monkey Bat"
    AddSolution "Devious Mockery Poisoner Grasping Hungry Night", "Snake Monkey Octopus Wolf Scorpion Bat"
    AddSolution "Grasping Mockery Night Devious Poisoner Hungry", "Octopus Bat Scorpion Snake Monkey Wolf"
    AddSolution "Grasping Night Mockery Devious Hungry Poisoner", "Octopus Monkey Snake Wolf Bat Scorpion"
    AddSolution "Grasping Poisoner Night Hungry Mockery Devious", "Octopus Scorpion Wolf Monkey Bat Snake"
    AddSolution "Grasping Devious Hungry Night Poisoner Mockery", "Octopus Snake Bat Scorpion Wolf Monkey"
    AddSolution "Hungry Grasping Devious Mockery Night Poisoner", "Wolf Octopus Monkey Bat Snake Scorpion"
    AddSolution "Hungry Devious Night Mockery Grasping Poisoner", "Wolf Snake Monkey Octopus Bat Scorpion"
    AddSolution "Hungry Night Grasping Poisoner Devious Mockery", "Wolf Bat Snake Scorpion Octopus Monkey"
    AddSolution "Hungry Poisoner Mockery Grasping Night Devious", "Wolf Scorpion Octopus Bat Monkey Snake"
    AddSolution "Hungry Poisoner Night Mockery Devious Grasping", "Wolf Bat Monkey Snake Scorpion Octopus"
    AddSolution "Mockery Grasping Poisoner Hungry Devious Night", "Monkey Scorpion Wolf Snake Octopus Bat"
    AddSolution "Mockery Hungry Devious Night Grasping Poisoner", "Monkey Snake Octopus Bat Wolf Scorpion"
    AddSolution "Mockery Night Devious Poisoner Hungry Grasping", "Monkey Snake Scorpion Wolf Bat Octopus"
    AddSolution "Mockery Poisoner Grasping Devious Night Hungry", "Monkey Octopus Bat Snake Scorpion Wolf"
    AddSolution "Night Grasping Hungry Poisoner Devious Mockery", "Bat Octopus Scorpion Snake Wolf Monkey"
    AddSolution "Night Mockery Poisoner Devious Grasping Hungry", "Bat Scorpion Snake Octopus Monkey Wolf"
    AddSolution "Poisoner Grasping Devious Night Mockery Hungry", "Scorpion Snake Bat Monkey Octopus Wolf"
    AddSolution "Poisoner Hungry Devious Mockery Night Grasping", "Scorpion Wolf Bat Monkey Snake Octopus"
    AddSolution "Poisoner Hungry Mockery Grasping Night Devious", "Scorpion Monkey Bat Octopus Wolf Snake"
    AddSolution "Poisoner Hungry Night Mockery Grasping Devious", "Scorpion Bat Octopus Monkey Wolf Snake"
    AddSolution "Poisoner Mockery Grasping Night Hungry Devious", "Scorpion Monkey Wolf Bat Octopus Snake"
    AddSolution "Poisoner Mockery Hungry Devious Night Grasping", "Scorpion Monkey Bat Snake Wolf Octopus"
    AddSolution "Poisoner Night Devious Grasping Hungry Mockery", "Scorpion Bat Wolf Octopus Snake Monkey"
End Sub

Private Sub AddSolution(pstrKeywords As String, pstrCrests As String)
    Dim strKeys() As String
    Dim strCrests() As String
    Dim i As Long
    
    strKeys = Split(pstrKeywords, " ")
    strCrests = Split(pstrCrests, " ")
    With mtypADQ(mlngCurrent)
        For i = 0 To 5
            .Keyword(i + 1) = GetKeyCode(strKeys(i))
            .Crest(i + 1) = GetCrestCode(strCrests(i))
        Next
    End With
    mlngCurrent = mlngCurrent + 1
End Sub

Private Function GetKeyCode(pstrKey As String) As KeywordEnum
    Select Case pstrKey
        Case "Devious": GetKeyCode = keDevious
        Case "Grasping": GetKeyCode = keGrasping
        Case "Hungry": GetKeyCode = keHungry
        Case "Mockery": GetKeyCode = keMockery
        Case "Night": GetKeyCode = keNight
        Case "Poisoner": GetKeyCode = kePoisoner
    End Select
End Function

Private Function GetKeyName(penKey As KeywordEnum) As String
    Select Case penKey
        Case keDevious: GetKeyName = "Devious"
        Case keGrasping: GetKeyName = "Grasping"
        Case keHungry: GetKeyName = "Hungry"
        Case keMockery: GetKeyName = "Mockery"
        Case keNight: GetKeyName = "Night"
        Case kePoisoner: GetKeyName = "Poisoner"
    End Select
End Function

Private Function GetCrestCode(pstrCrest As String) As CrestEnum
    Select Case pstrCrest
        Case "Bat": GetCrestCode = ceBat
        Case "Monkey": GetCrestCode = ceMonkey
        Case "Octopus": GetCrestCode = ceOctopus
        Case "Scorpion": GetCrestCode = ceScorpion
        Case "Snake": GetCrestCode = ceSnake
        Case "Wolf": GetCrestCode = ceWolf
    End Select
End Function

Private Function GetCrestName(penCrest As CrestEnum) As String
    Select Case penCrest
        Case ceBat: GetCrestName = "Bat"
        Case ceMonkey: GetCrestName = "Monkey"
        Case ceOctopus: GetCrestName = "Octopus"
        Case ceScorpion: GetCrestName = "Scorpion"
        Case ceSnake: GetCrestName = "Snake"
        Case ceWolf: GetCrestName = "Wolf"
    End Select
End Function

Private Function GetCrestDisplay(penCrest As CrestEnum) As String
    Select Case penCrest
        Case ceBat: GetCrestDisplay = "Bat-SE"
        Case ceMonkey: GetCrestDisplay = "Monkey-E"
        Case ceOctopus: GetCrestDisplay = "Octopus-NE"
        Case ceScorpion: GetCrestDisplay = "Scorpion-NW"
        Case ceSnake: GetCrestDisplay = "Snake-W"
        Case ceWolf: GetCrestDisplay = "Wolf-SW"
    End Select
End Function


' ************* KEYWORDS *************


Private Sub imgKeyword_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Don't show as clickable unless this keyword is a valid choice
    If mblnValid(Index) Then SetMouseCursor mcHand
End Sub

Private Sub imgKeyword_Click(Index As Integer)
    Dim i As Long
    
    ' Ignore if this keyword has already been chosen
    For i = 1 To 6
        If mlngArray(i) = Index Then Exit Sub
    Next
    ChooseKey Index
End Sub

Private Sub ChooseKey(ByVal plngIndex As Long)
    mlngArray(mlngCurrent) = plngIndex
    mlngCurrent = mlngCurrent + 1
    AutoComplete
End Sub

Private Sub AutoComplete()
    Dim lngSolution As Long
    Dim lngCount As Long
    Dim lngIndex As Long
    Dim i As Long
    
    ' Find all valid solutions
    For lngSolution = 0 To ADQ_Solutions
        With mtypADQ(lngSolution)
            .Valid = False
            For i = 1 To mlngCurrent
                If .Keyword(i) <> mlngArray(i) Then Exit For
            Next
            If i = mlngCurrent Then
                .Valid = True
                lngCount = lngCount + 1
                lngIndex = lngSolution
            End If
        End With
    Next
    ' If only one solution, solve and exit
    If lngCount = 1 Then
        Me.imgAccept.Tag = ""
        With mtypADQ(lngIndex)
            For i = 1 To 6
                mblnValid(i) = False
                mlngArray(i) = .Keyword(i)
                ' Create solution text
                Me.imgAccept.Tag = Me.imgAccept.Tag & i & ". " & GetCrestDisplay(.Crest(i)) & "  "
                If i = 3 Then Me.imgAccept.Tag = Trim$(Me.imgAccept.Tag) & vbNewLine
            Next
            Me.imgAccept.Tag = Trim$(Me.imgAccept.Tag)
        End With
        DrawScreen
        Exit Sub
    End If
    ' More than one solution; check how many choices for next keyword
    lngCount = 0
    Erase mblnValid
    For lngSolution = 0 To ADQ_Solutions
        With mtypADQ(lngSolution)
            If .Valid Then
                If Not mblnValid(.Keyword(mlngCurrent)) Then lngCount = lngCount + 1
                mblnValid(.Keyword(mlngCurrent)) = True
                lngIndex = .Keyword(mlngCurrent)
            End If
        End With
    Next
    ' If only one choice for next keyword, choose it and recurse
    If lngCount = 1 Then
        ChooseKey lngIndex
        AutoComplete
        Exit Sub
    End If
    ' Multiple choices are available; show results
    DrawScreen
End Sub


' ************* DRAWING *************


Private Sub DrawScreen()
    Const RowHeight As Long = 19
    Dim blnChosen(6) As Boolean
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngOffsetX As Long
    Dim lngOffsetY As Long
    Dim lngBlank As Long
    Dim i As Long
    
    lngOffsetX = 10
    lngOffsetY = 10
    ' Identify the keywords that have been chosen so far
    For i = 1 To 6
        blnChosen(mlngArray(i)) = True
        Me.imgKeyword(i).Visible = False
        Me.imgBlank(i).Visible = False
    Next
    ' Show the unselected keywords along left side
    lngLeft = lngOffsetX
    lngTop = lngOffsetY
    For i = 1 To 6
        If mblnValid(i) And Not blnChosen(i) Then
            Me.imgKeyword(i).Move lngLeft, lngTop
            lngTop = lngTop + RowHeight
            Me.imgKeyword(i).Visible = True
        End If
    Next
    ' Move to right colum
    lngOffsetX = 120
    lngBlank = 1
    ' Line 1
    lngLeft = lngOffsetX
    lngTop = lngOffsetY
    ShowRiddle 1, lngBlank, Me.imgLine1, lngLeft, lngTop
    ' Line 2
    lngLeft = lngOffsetX
    lngTop = lngTop + RowHeight
    ShowRiddle 2, lngBlank, Me.imgLine2a, lngLeft, lngTop
    ShowRiddle 3, lngBlank, Me.imgLine2b, lngLeft, lngTop
    ' Line 3
    lngLeft = lngOffsetX
    lngTop = lngTop + RowHeight
    ShowRiddle 4, lngBlank, Me.imgLine3a, lngLeft, lngTop
    ShowRiddle 5, lngBlank, Me.imgLine3b, lngLeft, lngTop
    ' Line 4
    lngLeft = lngOffsetX
    lngTop = lngTop + RowHeight
    ShowRiddle 6, lngBlank, Me.imgLine4, lngLeft, lngTop
End Sub

Private Sub ShowRiddle(plngIndex As Long, plngBlank As Long, pimgLine As Image, plngLeft As Long, plngTop As Long)
    Dim lngKeyword As Long
    
    lngKeyword = mlngArray(plngIndex)
    If lngKeyword <> 0 Then
        With Me.imgKeyword(lngKeyword)
            .Move plngLeft, plngTop
            plngLeft = plngLeft + .Width
            .Visible = True
        End With
    Else
        With Me.imgBlank(plngBlank)
            .Move plngLeft, plngTop
            plngLeft = plngLeft + .Width
            .Visible = True
            plngBlank = plngBlank + 1
        End With
    End If
    pimgLine.Move plngLeft, plngTop
    plngLeft = plngLeft + pimgLine.Width
End Sub


' ************* BUTTONS *************


Private Sub imgReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub imgReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub imgReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
    ResetChoices
    DrawScreen
End Sub

Private Sub ResetChoices()
    Dim i As Long
    
    mlngCurrent = 1
    Erase mlngArray
    For i = 1 To 6
        mblnValid(i) = True
    Next
    Me.imgAccept.Tag = vbNullString
End Sub

Private Sub imgAccept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub imgAccept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub imgAccept_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Len(Me.imgAccept.Tag) = 0 Then
        MsgBox "Complete the riddle first", vbInformation, "Notice"
        Exit Sub
    End If
    SetMouseCursor mcHand
    Load frmADQ2
    frmADQ2.Solution = Me.imgAccept.Tag
    frmADQ2.Show
    mblnAccept = True
    Unload Me
End Sub

Private Sub imgHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
    ShowHelp "ADQ"
End Sub
