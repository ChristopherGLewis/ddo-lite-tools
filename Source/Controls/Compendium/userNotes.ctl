VERSION 5.00
Begin VB.UserControl userNotes 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0FF&
   ClientHeight    =   3792
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3792
   ScaleWidth      =   4620
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   3540
      ScaleHeight     =   240
      ScaleWidth      =   252
      TabIndex        =   4
      ToolTipText     =   "Properties"
      Top             =   60
      Width           =   252
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3840
      ScaleHeight     =   240
      ScaleWidth      =   252
      TabIndex        =   3
      ToolTipText     =   "Maximize"
      Top             =   60
      Width           =   252
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4140
      ScaleHeight     =   240
      ScaleWidth      =   252
      TabIndex        =   2
      ToolTipText     =   "Help"
      Top             =   60
      Width           =   252
   End
   Begin Compendium.userTab usrtab 
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3372
      _ExtentX        =   5948
      _ExtentY        =   656
      Captions        =   "Magi,Shared,Public,...,+"
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      Height          =   3132
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "ctl"
      Top             =   360
      Width           =   4392
   End
End
Attribute VB_Name = "userNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Maximize()

Private mblnInit As Boolean

Public Sub Init()
    SizeControls
    mblnInit = True
    RefreshColors
End Sub

Public Sub RefreshColors()
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .usrtab.RefreshColors
        cfg.ApplyColors .txtNotes, cgeControls
    End With
End Sub

Private Sub SizeControls()
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngStep As Long
    Dim i As Long
    
    With UserControl
        With .usrtab
            .Move 0, 0, .TabsWidth, .TabHeight
            .ZOrder vbBringToFront
        End With
        lngTop = .usrtab.Height - Screen.TwipsPerPixelY
        .txtNotes.Move 0, lngTop, .ScaleWidth, .ScaleHeight - lngTop
        .txtNotes.ZOrder vbSendToBack
        lngLeft = .ScaleWidth - .pic(0).Width - Screen.TwipsPerPixelX
        lngTop = (.txtNotes.Top - .pic(0).Height) \ 2
        lngStep = .pic(0).Left - .pic(1).Left
        If lngStep < .pic(0).Width Then lngStep = .pic(0).Width
        For i = 0 To 2
            .pic(i).Move lngLeft, lngTop
            lngLeft = lngLeft - lngStep
        Next
    End With
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub pic_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub pic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    Select Case Index
        Case 0
        Case 1: RaiseEvent Maximize
        Case 2
    End Select
End Sub

Private Sub UserControl_Resize()
    SizeControls
End Sub
