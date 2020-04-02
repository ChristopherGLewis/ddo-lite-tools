VERSION 5.00
Begin VB.UserControl userLinkLists 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0FF&
   ClientHeight    =   5148
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5928
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   5148
   ScaleWidth      =   5928
   Begin VB.PictureBox picLinks 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2652
      Left            =   720
      ScaleHeight     =   2652
      ScaleWidth      =   2832
      TabIndex        =   0
      Tag             =   "ctl"
      Top             =   900
      Visible         =   0   'False
      Width           =   2832
      Begin Compendium.userMenu usrmnu 
         Height          =   1632
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   1872
         _ExtentX        =   2815
         _ExtentY        =   1207
      End
   End
   Begin VB.Label lblControl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Links"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   444
   End
   Begin VB.Shape shpLinks 
      Height          =   3072
      Left            =   540
      Top             =   720
      Visible         =   0   'False
      Width           =   3192
   End
   Begin VB.Menu mnuMain 
      Caption         =   "LinkContext"
      Index           =   0
      Begin VB.Menu mnuLinkContext 
         Caption         =   "New List"
         Index           =   0
      End
      Begin VB.Menu mnuLinkContext 
         Caption         =   "Undelete"
         Index           =   1
         Begin VB.Menu mnuUndelete 
            Caption         =   "Item 1"
            Index           =   0
         End
      End
      Begin VB.Menu mnuLinkContext 
         Caption         =   "-"
         Index           =   2
      End
   End
End
Attribute VB_Name = "userLinkLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngX As Long
Private mlngY As Long

Private Sub UserControl_Resize()
    Dim lngX As Long
    Dim lngY As Long
    
    lngX = Screen.TwipsPerPixelX
    lngY = Screen.TwipsPerPixelY
    With UserControl
        .shpLinks.Move 0, 0, .ScaleWidth, .ScaleHeight
        .picLinks.Move lngX, lngY, .ScaleWidth - lngX * 2, .ScaleHeight - lngY * 2
    End With
End Sub


' ************* LINKS *************


Public Sub Init()
    With UserControl
        .lblControl.Visible = False
        .shpLinks.Visible = True
        .picLinks.Visible = True
    End With
End Sub

Public Sub RefreshColors()
    Dim enGroup As ColorGroupEnum
    Dim i As Long
    
    enGroup = cgeControls
    With UserControl
        .BackColor = cfg.GetColor(enGroup, cveBackground)
        .picLinks.BackColor = cfg.GetColor(enGroup, cveBackground)
        .shpLinks.BorderColor = cfg.GetColor(enGroup, cveBorderExterior)
        For i = 0 To .usrmnu.UBound
            cfg.ApplyColors .usrmnu(i), enGroup
        Next
    End With
End Sub

Public Sub ShowLinkLists()
    Dim i As Long
    
    For i = 1 To db.LinkLists
        If i > UserControl.usrmnu.UBound Then Load UserControl.usrmnu(i)
        gtypMenu = db.LinkList(i)
        With UserControl.usrmnu(i)
            .LoadData
            .Move db.LinkList(i).Left, db.LinkList(i).Top
            .Visible = Not db.LinkList(i).Deleted
        End With
    Next
    For i = UserControl.usrmnu.UBound To db.LinkLists + 1 Step -1
        Unload UserControl.usrmnu(i)
    Next
End Sub

Public Sub SaveData(plngIndex As Long)
    UserControl.usrmnu(plngIndex).SaveData
End Sub

Public Sub GetMenuCoords(plngIndex As Long, plngLeft As Long, plngTop As Long)
    With UserControl.usrmnu(plngIndex)
        plngLeft = .Left
        plngTop = .Top
    End With
End Sub

Private Sub usrmnu_Changed(Index As Integer)
    db.LinkList(Index) = gtypMenu
    DirtyFlag dfeLinks
End Sub

Private Sub usrmnu_Deleted(Index As Integer)
    db.LinkList(Index) = gtypMenu
    DirtyFlag dfeLinks
End Sub

Private Sub usrmnu_Copy(Index As Integer)
    AddList 0, 0
End Sub

Private Sub AddList(plngX As Long, plngY As Long)
    db.LinkLists = db.LinkLists + 1
    ReDim Preserve db.LinkList(1 To db.LinkLists)
    gtypMenu.Left = plngX
    gtypMenu.Top = plngY
    db.LinkList(db.LinkLists) = gtypMenu
    Load UserControl.usrmnu(db.LinkLists)
    With UserControl.usrmnu(db.LinkLists)
        .Left = plngX
        .Top = plngY
        .Init
        .ZOrder vbBringToFront
        .Visible = True
    End With
    DirtyFlag dfeLinks
End Sub

Private Sub picLinks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngIndex As Long
    Dim i As Long
    
    If Button <> vbRightButton Then Exit Sub
    mlngX = X
    mlngY = Y
    lngIndex = -1
    For i = 1 To db.LinkLists
        If db.LinkList(i).Deleted Then
            lngIndex = lngIndex + 1
            If lngIndex > UserControl.mnuUndelete.UBound Then Load UserControl.mnuUndelete(lngIndex)
            With UserControl.mnuUndelete(lngIndex)
                .Caption = db.LinkList(i).Title
                .Tag = i
                .Visible = True
            End With
        End If
    Next
    For i = UserControl.mnuUndelete.UBound To lngIndex + 1 Step -1
        If i > 0 Then Unload UserControl.mnuUndelete(i)
    Next
    UserControl.mnuLinkContext(1).Visible = (lngIndex <> -1)
    UserControl.mnuLinkContext(2).Visible = (db.Templates > 0)
    For i = 1 To db.Templates
        lngIndex = i + 2
        If lngIndex > UserControl.mnuLinkContext.UBound Then Load UserControl.mnuLinkContext(lngIndex)
        With UserControl.mnuLinkContext(lngIndex)
            .Caption = db.Template(i).Title
            .Tag = i
            .Visible = True
        End With
    Next
    For i = UserControl.mnuLinkContext.UBound To lngIndex + 1 Step -1
        If i > 3 Then Unload UserControl.mnuUndelete(i)
    Next
    PopupMenu UserControl.mnuMain(0)
End Sub

Private Sub mnuLinkContext_Click(Index As Integer)
    Dim typBlank As MenuType
    Dim lngIndex As Long
    
    Select Case UserControl.mnuLinkContext(Index).Caption
        Case "New List"
            AutoSave
            gtypMenu = typBlank
            gtypMenu.Title = "Untitled"
            gtypMenu.LinkList = True
            gtypMenu.Selected = -1
            frmMenuEditor.Show vbModal, UserControl
            If gtypMenu.Accepted Then AddList mlngX, mlngY
        Case Else
            lngIndex = Val(UserControl.mnuLinkContext(Index).Tag)
            If lngIndex < 1 Or lngIndex > db.Templates Then Exit Sub
            gtypMenu = db.Template(lngIndex)
            AddList mlngX, mlngY
    End Select
End Sub

Private Sub mnuUndelete_Click(Index As Integer)
    Dim lngIndex As Long
    
    lngIndex = UserControl.mnuUndelete(Index).Tag
    gtypMenu = db.LinkList(lngIndex)
    gtypMenu.Deleted = False
    UserControl.usrmnu(lngIndex).LoadData
    UserControl.usrmnu(lngIndex).Visible = True
    DirtyFlag dfeLinks
End Sub
