VERSION 5.00
Begin VB.Form frmSagaDetail 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saga Detail"
   ClientHeight    =   5556
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7164
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSagaDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5556
   ScaleWidth      =   7164
   ShowInTaskbar   =   0   'False
   Begin Compendium.userHeader usrHeader 
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7152
      _ExtentX        =   12615
      _ExtentY        =   656
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin VB.CheckBox chkClaim 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Claim Reward"
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4860
      Visible         =   0   'False
      Width           =   1872
   End
   Begin VB.ComboBox cboSaga 
      Height          =   312
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   732
      Width           =   1632
   End
   Begin VB.Label lblSkip 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "You can skip 1 quest with Astral Shards (Plus 1 with VIP)"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   1440
      TabIndex        =   10
      Top             =   1860
      Width           =   6432
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Skip:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   1860
      Width           =   1152
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Chalice has completed 10 of 12 quests and earned 30 points"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   60
      TabIndex        =   32
      Top             =   4380
      Visible         =   0   'False
      Width           =   7032
   End
   Begin VB.Label lblRenown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "15,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   5580
      TabIndex        =   31
      Top             =   3840
      Width           =   1152
   End
   Begin VB.Label lblRenown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "10,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   4200
      TabIndex        =   30
      Top             =   3840
      Width           =   1152
   End
   Begin VB.Label lblRenown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "7,500"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   2820
      TabIndex        =   29
      Top             =   3840
      Width           =   1152
   End
   Begin VB.Label lblRenown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1440
      TabIndex        =   28
      Top             =   3840
      Width           =   1152
   End
   Begin VB.Label lblXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "26,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   5580
      TabIndex        =   26
      Top             =   3480
      Width           =   1152
   End
   Begin VB.Label lblXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "20,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   4200
      TabIndex        =   25
      Top             =   3480
      Width           =   1152
   End
   Begin VB.Label lblXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "13,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   2820
      TabIndex        =   24
      Top             =   3480
      Width           =   1152
   End
   Begin VB.Label lblXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "8,000"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1440
      TabIndex        =   23
      Top             =   3480
      Width           =   1152
   End
   Begin VB.Label lblRenown 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Renown:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   60
      TabIndex        =   27
      Top             =   3840
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "True Elite"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   5580
      TabIndex        =   16
      Top             =   2760
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Elite"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   4200
      TabIndex        =   15
      Top             =   2760
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   2820
      TabIndex        =   14
      Top             =   2760
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   1440
      TabIndex        =   13
      Top             =   2760
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tomes:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   2220
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NPC:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1500
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tier:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1140
      Width           =   1152
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Saga:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   1152
   End
   Begin VB.Label lblSkillTome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "+1 Skill Tomes"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1380
      TabIndex        =   12
      Top             =   2220
      Width           =   1992
   End
   Begin VB.Label lnkNPC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Gizla"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   1
      Left            =   2136
      TabIndex        =   8
      Top             =   1500
      Width           =   432
   End
   Begin VB.Label lblComma 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   ", "
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   2016
      TabIndex        =   34
      Top             =   1500
      Width           =   120
   End
   Begin VB.Label lnkNPC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Grazla"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   1500
      Width           =   576
   End
   Begin VB.Label lblTier 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Epic"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1440
      TabIndex        =   5
      Top             =   1140
      Width           =   1692
   End
   Begin VB.Label lblSagaName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "The Pirates of the Thunder Sea (heroic)"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3240
      TabIndex        =   3
      Top             =   780
      Width           =   5652
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "30"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   5580
      TabIndex        =   21
      Top             =   3120
      Width           =   1152
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "25"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   4200
      TabIndex        =   20
      Top             =   3120
      Width           =   1152
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "17"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   2820
      TabIndex        =   19
      Top             =   3120
      Width           =   1152
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "10"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1440
      TabIndex        =   18
      Top             =   3120
      Width           =   1152
   End
   Begin VB.Label lblXP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "XP:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   60
      TabIndex        =   22
      Top             =   3480
      Width           =   1152
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Points:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   3120
      Width           =   1152
   End
End
Attribute VB_Name = "frmSagaDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSaga As Long
Private mlngCharacter As Long

Private mblnOverride As Boolean

Private Sub Form_Load()
    cfg.Configure Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
End Sub

Private Sub usrHeader_Click(pstrCaption As String)
    If pstrCaption = "Help" Then ShowHelp "Saga_Detail"
End Sub

Public Sub SetDetail(plngSaga As Long, plngCharacter As Long)
    mlngSaga = plngSaga
    mlngCharacter = plngCharacter
    InitCombos
    ShowDetail
End Sub

Private Sub InitCombos()
    Dim i As Long
    
    mblnOverride = True
    If mlngSaga < 1 Then mlngSaga = 1
    ComboClear Me.cboSaga
    For i = 1 To db.Sagas
        If db.Saga(i).Tier = db.Saga(mlngSaga).Tier Then ComboAddItem Me.cboSaga, db.Saga(i).Abbreviation, i
    Next
    ComboSetValue Me.cboSaga, mlngSaga
    mblnOverride = False
End Sub

Private Sub cboSaga_Click()
    If mblnOverride Then Exit Sub
    mlngSaga = ComboGetValue(Me.cboSaga)
    ShowDetail
End Sub

Private Sub ShowDetail()
    Dim i As Long
    
    If mlngSaga < 1 Then mlngSaga = 1
    With db.Saga(mlngSaga)
        Me.lblSagaName.Caption = .SagaName
        If .Tier = steEpic Then Me.lblTier.Caption = "Epic" Else Me.lblTier.Caption = "Heroic"
        Me.lnkNPC(0).Caption = .NPC(1)
        Me.lblComma.Left = Me.lnkNPC(0).Left + Me.lnkNPC(0).Width
        Me.lnkNPC(1).Left = Me.lblComma.Left + Me.lblComma.Width
        If .NPCs = 2 Then
            Me.lnkNPC(1).Caption = .NPC(2)
            Me.lblComma.Visible = True
            Me.lnkNPC(1).Visible = True
        Else
            Me.lblComma.Visible = False
            Me.lnkNPC(1).Visible = False
        End If
        If .Astrals = 1 Then
            Me.lblSkip.Caption = "You can skip 1 quest with Astral Shards (Plus 1 with VIP)"
        Else
            Me.lblSkip.Caption = "You can skip " & .Astrals & " quests with Astral Shards (Plus 1 with VIP)"
        End If
        Me.lblSkillTome.Caption = "+" & .Tome & " Skill Tomes"
        For i = 1 To 4
            With .Reward(i - 1)
                Me.lblXP(i).Caption = Format(.xp, "#,##0")
                Me.lblRenown(i).Caption = Format(.Renown, "#,##0")
                Me.lblPoints(i).Caption = .Points
            End With
        Next
    End With
    ShowCharacterDetail
End Sub

Private Sub ShowCharacterDetail()
    Dim lngDone As Long
    Dim lngTotal As Long
    Dim lngPoints As Long
    Dim i As Long
    
    If mlngCharacter = 0 Then
        Me.lblStatus.Visible = False
        Me.chkClaim.Visible = False
        Exit Sub
    End If
    lngTotal = db.Saga(mlngSaga).Quests
    For i = 1 To lngTotal
        lngDone = lngDone + 1
        Select Case db.Character(mlngCharacter).Saga(mlngSaga).Progress(i)
            Case peNone: lngDone = lngDone - 1
            Case peSolo, peCasual, peNormal: lngPoints = lngPoints + 1
            Case peHard: lngPoints = lngPoints + 2
            Case peElite, peAstrals, peVIP: lngPoints = lngPoints + 3
        End Select
    Next
    Me.lblStatus.Caption = db.Character(mlngCharacter).Character & " has completed " & lngDone & " of " & lngTotal & " quests and earned " & lngPoints & " points"
    Me.lblStatus.Visible = True
    Me.chkClaim.Enabled = (lngDone = lngTotal)
    Me.chkClaim.Visible = True
End Sub

Private Sub lnkNPC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNPC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    xp.OpenURL MakeWiki(Me.lnkNPC(Index).Caption)
End Sub

Private Sub lnkNPC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub chkClaim_Click()
    Dim i As Long
    
    If UncheckButton(Me.chkClaim, mblnOverride) Then Exit Sub
    If MsgBox("Clear all progress for saga " & db.Saga(mlngSaga).Abbreviation & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    With db.Character(mlngCharacter).Saga(mlngSaga)
        For i = 1 To db.Saga(mlngSaga).Quests
            .Progress(i) = peNone
        Next
    End With
    ShowDetail
    DirtyFlag dfeData
    Unload Me
End Sub

