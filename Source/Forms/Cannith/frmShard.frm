VERSION 5.00
Begin VB.Form frmShard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shards"
   ClientHeight    =   9015
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   12390
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShard.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userInfo usrInfo 
      Height          =   7812
      Left            =   4140
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   6612
      _ExtentX        =   11668
      _ExtentY        =   13785
   End
   Begin VB.PictureBox picML 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7392
      Left            =   10920
      ScaleHeight     =   7395
      ScaleWidth      =   1335
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1332
   End
   Begin VB.ComboBox cboScaling 
      Height          =   312
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   8100
      Width           =   2712
   End
   Begin VB.ComboBox cboGroup 
      Height          =   312
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   7680
      Width           =   2712
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   1260
      TabIndex        =   2
      Top             =   600
      Width           =   2412
   End
   Begin VB.ListBox lstShard 
      Appearance      =   0  'Flat
      Height          =   6330
      Left            =   300
      TabIndex        =   3
      Top             =   1020
      Width           =   3672
   End
   Begin CannithCrafting.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8640
      Width           =   12384
      _ExtentX        =   21855
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "New Item;New Gearset;Load Gearset"
      RightLinks      =   "Materials;Augments;Scaling"
   End
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12384
      _ExtentX        =   21855
      _ExtentY        =   688
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin VB.Label lblScaling 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Scaling"
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
      Height          =   216
      Left            =   11040
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Scaling:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   8136
      Width           =   852
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Group:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   7716
      Width           =   852
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Search:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   624
      Width           =   852
   End
End
Attribute VB_Name = "frmShard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrShard() As String
Private mstrShardName As String

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    InitControls
    ShowMatches
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseApp
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Shards"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    FooterClick pstrCaption
End Sub


' ************* INITIALIZE *************


Private Sub InitControls()
    Dim i As Long
    
    mblnOverride = True
    ReDim mstrShard(1 To db.Shards)
    For i = 1 To db.Shards
        mstrShard(i) = LCase$(db.Shard(i).ShardName)
    Next
    With Me.cboGroup
        .AddItem "Any"
        For i = 1 To db.Groups
            .AddItem db.Group(i)
            .ItemData(.NewIndex) = i
        Next
        .ListIndex = 0
    End With
    With Me.cboScaling
        .AddItem "Any"
        For i = 1 To db.Scales
            .AddItem db.Scaling(i).ScaleName
            .ItemData(.NewIndex) = i
        Next
        .ListIndex = 0
    End With
    mblnOverride = False
End Sub


' ************* OPEN TO SHARD *************


Public Property Get ShardName() As String
    ShardName = mstrShardName
End Property

Public Property Let ShardName(ByVal pstrShardName As String)
    Dim lngTopIndex As Long
    Dim blnFound As Boolean
    Dim i As Long
    
    With Me.lstShard
        For i = 0 To .ListCount
            If .List(i) = pstrShardName Then
                blnFound = True
                Exit For
            End If
        Next
        If blnFound Then
            .ListIndex = i
            lngTopIndex = i - (.Height \ Me.TextHeight("Q")) \ 2
            If lngTopIndex < 0 Then lngTopIndex = 0
            .TopIndex = lngTopIndex
            mstrShardName = pstrShardName
        Else
            mstrShardName = vbNullString
        End If
    End With
End Property


' ************* SEARCH *************


Private Sub txtSearch_GotFocus()
    With Me.txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyReturn
            If Me.lstShard.ListCount > 0 Then
                Me.lstShard.ListIndex = 0
                Me.lstShard.SetFocus
            End If
    End Select
End Sub

Private Sub txtSearch_Change()
    ShowMatches
End Sub

Private Sub cboGroup_Click()
    ShowMatches
End Sub

Private Sub cboScaling_Click()
    ShowMatches
End Sub

Private Sub ShowMatches()
    Dim strSearch As String
    Dim strGroup As String
    Dim strScaling As String
    Dim blnMatch As Boolean
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    Me.lstShard.ListIndex = -1
    Me.lstShard.Clear
    strSearch = LCase$(Me.txtSearch.Text)
    If Me.cboGroup.ListIndex > 0 Then strGroup = Me.cboGroup.Text
    If Me.cboScaling.ListIndex > 0 Then strScaling = Me.cboScaling.Text
    For i = 1 To db.Shards
        blnMatch = True
        If InStr(mstrShard(i), strSearch) = 0 Then blnMatch = False
        If Len(strGroup) Then
            If db.Shard(i).Group <> strGroup Then blnMatch = False
        End If
        If Len(strScaling) Then
            If db.Shard(i).ScaleName <> strScaling Then blnMatch = False
        End If
        If blnMatch Then
            Me.lstShard.AddItem db.Shard(i).ShardName
            Me.lstShard.ItemData(Me.lstShard.NewIndex) = i
        End If
    Next
End Sub


' ************* DETAILS *************


Private Sub lstShard_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Me.txtSearch.SetFocus
        Case vbKeyUp
            If Me.lstShard.ListIndex = 0 Then Me.txtSearch.SetFocus
    End Select
End Sub

Private Sub lstShard_Click()
    ShardDetails
End Sub

Private Sub ShardDetails()
    Dim strText As String
    Dim strPower As String
    Dim strOldPower As String
    Dim lngScale As Long
    Dim lngPos As Long
    Dim i As Long
    
    Me.usrInfo.Clear
    Me.lblScaling.Visible = False
    Me.picML.Visible = False
    If Me.lstShard.ListIndex = -1 Then
        Me.Caption = "Shards"
        SetFormIcon Me, GetGeneralResource(gieML)
        Exit Sub
    End If
    With db.Shard(Me.lstShard.ItemData(Me.lstShard.ListIndex))
        ' Form
        mstrShardName = .ShardName
        Me.Caption = .ShardName
        SetFormIcon Me, GetGeneralResource(gieShard, .Bound.Level)
        ' Header
        Me.usrInfo.TitleText = .ShardName
        Me.usrInfo.SetIcon GetGeneralResource(gieShard, .Bound.Level)
        If Len(.Descrip) Then Me.usrInfo.AddText .Descrip, 2
        If Len(.Notes) Then Me.usrInfo.AddText .Notes, 2
        Me.usrInfo.AddText "Group: " & .Group
        Me.usrInfo.AddText "Scaling: " & .ScaleName, 2
'        If .ML <> 0 Then
'            Me.usrInfo.AddText "Minimum Level: " & .ML
'            If .ML < 0 Then Me.usrInfo.AddText "(Supposed to be ML" & Abs(.ML) & " but can currently be applied to any ML item.)"
'            Me.usrInfo.AddText vbNullString
'        End If
        ' Slots
        Me.usrInfo.AddText "Prefix: " & Slots(.Prefix), 1, vbNullString, "Prefix: "
        Me.usrInfo.AddText "Suffix: " & Slots(.Suffix), 1, vbNullString, "Suffix: "
        Me.usrInfo.AddText "Extra: " & Slots(.Extra), 2, vbNullString, "Extra: "
        ' Bound
        Me.usrInfo.AddText "Bound Crafting:", 0
        Me.usrInfo.AddClipboard .ShardName & " (Bound)" & ClipboardText(.Bound), 1
        Me.usrInfo.AddText "Level " & .Bound.Level, 1, "   "
        AddRecipeToInfo .Bound, Me.usrInfo, "   "
        Me.usrInfo.AddText vbNullString, 2
        ' Unbound
        Me.usrInfo.AddText "Unbound Crafting:", 0
        Me.usrInfo.AddClipboard .ShardName & " (Unbound)" & ClipboardText(.Unbound), 1
        Me.usrInfo.AddText "Level " & .Unbound.Level, 1, "   "
        AddRecipeToInfo .Unbound, Me.usrInfo, "   "
        ' Scaling
        Me.picML.Cls
        If .ScaleName <> "None" Then lngScale = SeekScaling(.ScaleName)
        If lngScale Then
            lngPos = Me.picML.TextWidth("ML30:  ")
            For i = 1 To MAX_CRAFTING_LEVEL
                If i < .ML Then
                    strPower = vbNullString
                Else
                    strPower = db.Scaling(lngScale).Table(i)
                End If
                If strPower <> strOldPower Then
                    Me.picML.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
                Else
                    Me.picML.ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
                End If
                If i > 30 Then
                    strText = "PL" & i & ":  "
                Else
                    strText = "ML" & i & ":  "
                End If
                Me.picML.CurrentX = lngPos - Me.picML.TextWidth(strText)
                Me.picML.Print strText & strPower
                strOldPower = strPower
            Next
            Me.lblScaling.Visible = True
            Me.picML.Visible = True
        End If
    End With
    Me.usrInfo.Visible = True
End Sub

Private Function Slots(pblnSlot() As Boolean) As String
    Dim strText As String
    Dim i As Long
    
    For i = 0 To geGearCount - 1
        If pblnSlot(i) Then
            If Len(strText) Then strText = strText & ", "
            strText = strText & GetGearName(i)
        End If
    Next
    Slots = strText
End Function

Private Function ClipboardText(ptypRecipe As RecipeType) As String
    Dim strText As String
    Dim i As Long
    
    strText = vbNewLine
    With ptypRecipe
        strText = strText & "Level " & .Level & vbNewLine
        strText = strText & .Essences & " Essences" & vbNewLine
        For i = 1 To .Ingredients
            strText = strText & .Ingredient(i).Count & " " & Pluralized(.Ingredient(i)) & vbNewLine
        Next
    End With
    ClipboardText = strText
End Function
