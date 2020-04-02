VERSION 5.00
Begin VB.Form frmAugSlots 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Augment Slots"
   ClientHeight    =   2292
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   2928
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2292
   ScaleWidth      =   2928
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSnap 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   360
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   7
      Left            =   2520
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "Colorless"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   6
      Left            =   2100
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Yellow"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   5
      Left            =   1680
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "Green"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   4
      Left            =   1260
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "Blue"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   3
      Left            =   840
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "Purple"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   2
      Left            =   420
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "Orange"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   1
      Left            =   0
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "Red"
      Top             =   1620
      Width           =   408
   End
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   3
      Left            =   0
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   408
   End
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   2
      Left            =   0
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   420
      Width           =   408
   End
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   1
      Left            =   0
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   408
   End
   Begin VB.Label lnkHelp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   2460
      TabIndex        =   15
      Top             =   2040
      Width           =   384
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   60
      TabIndex        =   14
      Top             =   2040
      Width           =   60
   End
   Begin VB.Label lblAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   3
      Left            =   540
      TabIndex        =   13
      Top             =   960
      Width           =   60
   End
   Begin VB.Label lblAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   2
      Left            =   540
      TabIndex        =   12
      Top             =   540
      Width           =   60
   End
   Begin VB.Label lblAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   540
      TabIndex        =   11
      Top             =   120
      Width           =   60
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Choose up to three slots:"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   60
      TabIndex        =   1
      Top             =   1320
      Width           =   2304
   End
End
Attribute VB_Name = "frmAugSlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Form

Private mlngParentLeft As Long
Private mlngParentTop As Long
Private mlngLeft As Long
Private mlngTop As Long
Private mlngRight As Long
Private mlngBottom As Long

Private Sub Form_Load()
    cfg.RefreshColors Me
    InitIcons
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblColor.Caption = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseApp
End Sub

Public Sub Position(plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long)
    Me.tmrSnap.Enabled = True
    mlngLeft = plngLeft
    mlngTop = plngTop
    mlngRight = plngRight
    mlngBottom = plngBottom
    FollowParent
End Sub

Private Sub FollowParent()
    Dim lngLeft As Long
    Dim lngTop As Long
    
    mlngParentLeft = mfrmParent.Left
    mlngParentTop = mfrmParent.Top
    Do
        If FitsInScreen(mlngLeft, mlngBottom, lngLeft, lngTop) Then Exit Do
        If FitsInScreen(mlngRight, mlngTop, lngLeft, lngTop) Then Exit Do
        If FitsInScreen(mlngLeft, mlngTop - Me.Height, lngLeft, lngTop) Then Exit Do
        If FitsInScreen(mlngRight - Me.Width, mlngTop - Me.Height, lngLeft, lngTop) Then Exit Do
        If FitsInScreen(mlngLeft - Me.Width, mlngTop - Me.Height, lngLeft, lngTop) Then Exit Do
    Loop Until True
    If mlngParentLeft + lngLeft < 0 Then
        lngLeft = -mlngParentLeft
    ElseIf mlngParentLeft + lngLeft + Me.Width > Screen.Width Then
        lngLeft = Screen.Width - Me.Width - mlngParentLeft
    End If
    If mlngParentTop + lngTop < 0 Then
        lngTop = -mlngParentTop
    ElseIf mlngParentTop + lngTop + Me.Height > Screen.Height Then
        lngTop = Screen.Height - Me.Height - mlngParentTop
    End If
    Me.Move mlngParentLeft + lngLeft, mlngParentTop + lngTop
End Sub

Private Sub tmrSnap_Timer()
    If mlngParentLeft <> mfrmParent.Left Or mlngParentTop <> mfrmParent.Top Then FollowParent
End Sub

Private Function FitsInScreen(plngLeft As Long, plngTop As Long, plngSetLeft As Long, plngSetTop As Long) As Boolean
    Dim lngLeft As Long
    Dim lngTop As Long
    
    lngLeft = mlngParentLeft + plngLeft
    lngTop = mlngParentTop + plngTop
    If lngLeft + Me.Width <= Screen.Width Then
        If lngTop + Me.Height <= Screen.Height Then
           If lngLeft >= 0 Then
                If lngTop >= 0 Then
                    FitsInScreen = True
                End If
            End If
        End If
    End If
    plngSetLeft = plngLeft
    plngSetTop = plngTop
End Function

Public Sub SetParent(pfrm As Form)
    Set mfrmParent = pfrm
End Sub

Private Sub InitIcons()
    Dim enColor As AugmentColorEnum
    Dim i As Long
    
    For i = 1 To 3
        Me.lblAugment(i).ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        Me.lblAugment(i).Caption = "Unused slot"
        DrawPicture Me.picActive(i), "MSCBARTER"
    Next
    For enColor = 1 To 7
        DrawPicture Me.picColor(enColor), "AUG" & GetAugmentColorName(enColor, True)
    Next
End Sub

Private Sub DrawPicture(ppic As PictureBox, pstrResourceID As String, Optional pblnActive As Boolean = False)
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    With ppic
        lngWidth = .ScaleWidth
        lngHeight = .ScaleHeight
        .Cls
        On Error Resume Next
        .PaintPicture LoadResPicture(pstrResourceID, vbResIcon), PixelX, PixelY, lngWidth - (PixelX * 2), lngHeight - (PixelY * 2)
        If Err.Number Then .PaintPicture LoadResPicture("UNKNOWN", vbResIcon), PixelX, PixelY, lngWidth - (PixelX * 2), lngHeight - (PixelY * 2)
        On Error GoTo 0
    End With
    If pblnActive Then DrawBorder ppic
End Sub

Private Sub DrawBorder(ppic As PictureBox)
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngColor As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    
    lngWidth = ppic.ScaleWidth
    lngHeight = ppic.ScaleHeight
    lngColor = RGB(60, 160, 60)
    ' Top and Bottom
    lngLeft = PixelX * 2
    lngRight = lngWidth - PixelX * 2
    lngTop = 0
    lngBottom = lngHeight - PixelY
    ppic.Line (lngLeft, lngTop)-(lngRight, lngTop), lngColor
    ppic.Line (lngLeft, lngBottom)-(lngRight, lngBottom), lngColor
    ' Left and Right
    lngLeft = 0
    lngRight = lngWidth - PixelX
    lngTop = PixelY * 2
    lngBottom = lngHeight - PixelY * 2
    ppic.Line (lngLeft, lngTop)-(lngLeft, lngBottom), lngColor
    ppic.Line (lngRight, lngTop)-(lngRight, lngBottom), lngColor
    ' Draw boxes
    lngLeft = PixelX
    lngRight = lngWidth - PixelX * 2
    lngTop = PixelX
    lngBottom = lngHeight - PixelY * 2
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(165, 240, 165), PixelX, PixelY
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(250, 255, 250), PixelX, PixelY
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(160, 225, 160), PixelX, PixelY
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(70, 185, 90), PixelX, PixelY
End Sub

Private Sub DrawBox(ppic As PictureBox, plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long, plngColor As Long, plngPixelX As Long, plngPixelY As Long)
    ppic.Line (plngLeft, plngTop)-(plngRight, plngBottom), plngColor, B
    ' Shrink next box one pixel on all sides
    plngLeft = plngLeft + plngPixelX
    plngTop = plngTop + plngPixelY
    plngRight = plngRight - plngPixelX
    plngBottom = plngBottom - plngPixelY
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblColor.Caption = vbNullString
End Sub

Private Sub picColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    If Me.lblColor.Caption <> Me.picColor(Index).Tag Then Me.lblColor.Caption = Me.picColor(Index).Tag
End Sub

Private Sub picColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub picColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

