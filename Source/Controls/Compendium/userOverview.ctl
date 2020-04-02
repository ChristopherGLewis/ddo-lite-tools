VERSION 5.00
Begin VB.UserControl userOverview 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   11424
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11676
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   11424
   ScaleWidth      =   11676
   Begin Compendium.userCheckBox usrCheck 
      Height          =   312
      Left            =   3720
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   1212
      _ExtentX        =   2773
      _ExtentY        =   550
   End
   Begin Compendium.userButton usrButton 
      Height          =   372
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1152
      _ExtentX        =   2561
      _ExtentY        =   762
      Caption         =   "ADQ"
   End
   Begin Compendium.userNotes usrNotes 
      Height          =   3312
      Left            =   4620
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   5772
      _ExtentX        =   10181
      _ExtentY        =   4784
   End
   Begin Compendium.userCharacter usrCharacter 
      Height          =   2412
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   4392
      _ExtentX        =   7747
      _ExtentY        =   4255
   End
   Begin Compendium.userFavor usrFavor 
      Height          =   2532
      Left            =   1440
      TabIndex        =   0
      Top             =   4020
      Width           =   5952
      _ExtentX        =   10499
      _ExtentY        =   4466
   End
   Begin Compendium.userButton usrButton 
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1152
      _ExtentX        =   2561
      _ExtentY        =   762
      Caption         =   "Shroud"
   End
   Begin Compendium.userButton usrButton 
      Height          =   432
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3420
      Visible         =   0   'False
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   762
      Caption         =   "Shroud"
   End
   Begin Compendium.userButton usrButton 
      Height          =   372
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1152
      _ExtentX        =   2561
      _ExtentY        =   762
      Caption         =   "Timer"
   End
   Begin VB.Label lblControl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Overview"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   840
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuCharacter 
         Caption         =   "Add New"
         Index           =   0
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Delete"
         Index           =   1
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Rename"
         Index           =   2
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Import"
         Index           =   4
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Export"
         Index           =   5
         Begin VB.Menu mnuExport 
            Caption         =   "Character Builder Lite"
            Index           =   0
         End
         Begin VB.Menu mnuExport 
            Caption         =   "Ron's Character Planner"
            Index           =   1
         End
         Begin VB.Menu mnuExport 
            Caption         =   "DDO Builder"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "userOverview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mblnInit As Boolean
Private mblnOverride As Boolean


' ************* PUBLIC *************


Public Sub Init()
    With UserControl
        .lblControl.Visible = False
        .usrCharacter.Init
        .usrNotes.Init
        .usrFavor.Init cfg.Character
        .usrFavor.Visible = True
    End With
    RefreshColors
    SizeControls
    mblnInit = True
    SizeFavor
End Sub

Public Sub RefreshColors()
    Dim ctl As Control
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        For Each ctl In .Controls
            Select Case TypeName(ctl)
                Case "userButton"
                    ctl.RefreshColors
                    ctl.Redraw
                Case "userNotes"
                    ctl.RefreshColors
                Case Else
                    Select Case ctl.Tag
                        Case "ctl": cfg.ApplyColors ctl, cgeControls
                        Case Else: cfg.ApplyColors ctl, cgeWorkspace
                    End Select
            End Select
        Next
        RedrawFavor
        If mblnInit Then .usrCharacter.RefreshColors
    End With
End Sub


' ************* CHARACTERS *************


Public Sub CharacterChanged()
    UserControl.usrCharacter.CharacterChanged
End Sub

Private Sub usrCharacter_CharacterChanged()
    Dim frm As Form

    DirtyFlag dfeSettings
    With UserControl.usrFavor
        .Character = cfg.Character
        .ReDrawControl
        .Move .Left, .Top, .FitWidth, .FitHeight
    End With
    If GetForm(frm, "frmPatrons") Then frm.ReDrawForm
    If GetForm(frm, "frmChallenges") Then frm.CharacterListChanged
    If GetForm(frm, "frmSagas") Then frm.CharacterListChanged
End Sub

Private Sub usrCharacter_CharacterListChanged()
    Dim frm As Form

    frmCompendium.RedrawQuests
    DirtyFlag dfeData
    If GetForm(frm, "frmPatrons") Then frm.ReDrawForm
    If GetForm(frm, "frmChallenges") Then frm.CharacterListChanged
    If GetForm(frm, "frmSagas") Then frm.CharacterListChanged
    UserControl.usrFavor.ReDrawControl
End Sub


' ************* SIZING *************


Private Sub UserControl_Resize()
    If mblnInit Then SizeControls
End Sub

Private Sub SizeControls()
    Dim lngWidth As Long
    
    With UserControl
        lngWidth = .ScaleWidth - .usrNotes.Left
        If lngWidth >= .usrCharacter.Width Then .usrNotes.Width = lngWidth
    End With
    SizeFavor
End Sub

Private Sub SizeFavor()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    With UserControl
        With .usrFavor
            lngLeft = 0
            lngTop = .Top
            lngWidth = .FitWidth
            lngHeight = .FitHeight
        End With
        If lngLeft + lngWidth > .ScaleWidth Then lngWidth = .ScaleWidth - lngLeft
        If lngWidth < 120 Then lngWidth = 120
        If lngTop + lngHeight > .ScaleHeight Then lngHeight = .ScaleHeight - lngTop
        If lngHeight < 120 Then lngHeight = 120
        .usrFavor.Move lngLeft, lngTop, lngWidth, lngHeight
    End With
End Sub


' ************* NOTES *************



' ************* TOMES *************


Private Sub RedrawFavor()
    With UserControl.usrFavor
        .ReDrawControl
        .Move .Left, .Top, .FitWidth, .FitHeight
    End With
End Sub
