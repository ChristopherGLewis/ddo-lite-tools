VERSION 5.00
Begin VB.Form frmStopwatch 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stopwatch"
   ClientHeight    =   720
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   2592
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStopwatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   2592
   Begin VB.Timer tmrAlwaysOnTop 
      Interval        =   1000
      Left            =   1500
      Top             =   0
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1140
      Top             =   0
   End
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   780
      Top             =   0
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   720
      ScaleHeight     =   216
      ScaleWidth      =   1212
      TabIndex        =   0
      Top             =   60
      Width           =   1212
   End
   Begin Stopwatch.userSpinner usrspnCountdown 
      Height          =   312
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   912
      _ExtentX        =   1609
      _ExtentY        =   550
      Min             =   0
      Max             =   9
      Value           =   3
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Label lnkMore 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "More"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   2040
      TabIndex        =   4
      Top             =   60
      Width           =   456
   End
   Begin VB.Label lblIntro 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Countdown"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   120
      TabIndex        =   3
      Top             =   396
      Visible         =   0   'False
      Width           =   1032
   End
   Begin VB.Label lnkStart 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Start"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   444
   End
End
Attribute VB_Name = "frmStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngIntro As Long

Private mlngTop As Long
Private mlngHeight As Long

Private mlngLeft(4) As Long
Private mlngWidth(4) As Long

Private mblnOverride As Boolean

Private Sub Form_Load()
    StopwatchInit
    cfg.RefreshColors Me
    Me.picTime.ForeColor = vbWhite
    PositionForm Me, "Stopwatch"
    mlngLeft(0) = Me.picTime.Left
    mlngTop = Me.picTime.Top
    mlngWidth(0) = Me.picTime.TextWidth(" Start in 9 ")
    mlngHeight = Me.picTime.TextHeight("Q")
    CalculateDimensions 1, " 0.000 "
    CalculateDimensions 2, " 30.0 "
    CalculateDimensions 3, " 1:00 "
    CalculateDimensions 4, " 10:00 "
    Me.picTime.Move mlngLeft(0), mlngTop, mlngWidth(0), mlngHeight
    Me.lnkMore.Left = Me.picTime.Left + Me.picTime.Width + (Me.picTime.Left - (Me.lnkStart.Left + Me.lnkStart.Width))
    Me.Width = Me.lnkStart.Left + Me.lnkMore.Left + Me.lnkMore.Width + Me.Width - Me.ScaleWidth
    mlngIntro = Me.usrspnCountdown.Value
    SmallForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteCoords "Stopwatch", Me.Left, Me.Top
    CloseApp
End Sub

Private Sub CalculateDimensions(plngIndex As Long, pstrText As String)
    mlngWidth(plngIndex) = Me.picTime.TextWidth(pstrText)
    mlngLeft(plngIndex) = mlngLeft(0) + (mlngWidth(0) - mlngWidth(plngIndex)) \ 2
End Sub

Private Sub SmallForm()
    Me.lblIntro.Visible = False
    Me.usrspnCountdown.Visible = False
    Me.Height = Me.picTime.Top * 2 + Me.picTime.Height + Me.Height - Me.ScaleHeight
End Sub

Private Sub LargeForm()
    Me.Height = Me.picTime.Top + Me.usrspnCountdown.Top + Me.usrspnCountdown.Height + Me.Height - Me.ScaleHeight
    Me.lblIntro.Visible = True
    Me.usrspnCountdown.Visible = True
End Sub

Private Sub lnkMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub lnkMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
    Select Case Me.lnkMore.Caption
        Case "More"
            Me.lnkMore.Caption = "Less"
            LargeForm
        Case "Less"
            Me.lnkMore.Caption = "More"
            SmallForm
    End Select
End Sub

Private Sub lnkMore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub lnkStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub lnkStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
    StartClick
End Sub

Private Sub lnkStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetMouseCursor mcHand
End Sub

Private Sub lnkStart_DblClick()
    SetMouseCursor mcHand
    StartClick
End Sub

Private Sub StartClick()
    If Me.tmrCountdown.Enabled Then Me.picTime.Cls
    Me.tmrCountdown.Enabled = False
    Me.tmr.Enabled = False
    If Len(Me.lnkStart.Caption) = 5 Then ' Start
        Me.lnkStart.Caption = "Stop"
        mlngIntro = Me.usrspnCountdown.Value
        If mlngIntro Then StartIntro Else StartTimer
    Else ' Stop
        Me.tmr.Enabled = False
        Me.lnkStart.Caption = "Start"
    End If
End Sub

Private Sub StartTimer()
    Me.picTime.Move mlngLeft(1), mlngTop, mlngWidth(1)
    Me.tmr.Interval = 10
    StopwatchStart
    Me.tmr.Enabled = True
End Sub

Private Sub tmr_Timer()
    Dim dblTime As Double
    Dim lngMinutes As Long
    
    dblTime = StopwatchStop()
    Me.picTime.Cls
    Select Case dblTime
        Case Is < 10
            Me.picTime.Print Format(dblTime, " 0.000")
        Case Is < 30
            If Me.tmr.Interval = 10 Then Me.tmr.Interval = 50
            Me.picTime.Print Format(dblTime, " 0.00")
        Case Is < 60
            If Me.tmr.Interval = 50 Then
                Me.tmr.Interval = 100
                Me.picTime.Move mlngLeft(2), mlngTop, mlngWidth(2)
            End If
            Me.picTime.Print Format(dblTime, " 0.0")
        Case Is < 600
            If Me.tmr.Interval = 100 Then
                Me.tmr.Interval = 200
                Me.picTime.Move mlngLeft(3), mlngTop, mlngWidth(3)
            End If
            lngMinutes = dblTime \ 60
            Me.picTime.Print " " & lngMinutes & ":" & Format(dblTime Mod 60, "00")
            Me.picTime.Print Format(dblTime, "0.0")
        Case Else
            If Me.tmr.Interval = 200 Then
                Me.tmr.Interval = 1000
                Me.picTime.Move mlngLeft(4), mlngTop, mlngWidth(4)
            End If
            lngMinutes = dblTime \ 60
            Me.picTime.Print " " & lngMinutes & ":" & Format(dblTime Mod 60, "00")
    End Select
End Sub

Private Sub tmrAlwaysOnTop_Timer()
    SetAlwaysOnTop Me.hWnd
End Sub

Private Sub tmrCountdown_Timer()
    mlngIntro = mlngIntro - 1
    ShowIntro
End Sub

Private Sub StartIntro()
    ShowIntro
    Me.tmrCountdown.Enabled = True
End Sub

Private Sub ShowIntro()
    If mlngIntro > 0 Then
        Me.picTime.Move mlngLeft(0), mlngTop, mlngWidth(0)
        Me.picTime.Cls
        Me.picTime.Print " Start in " & mlngIntro
    Else
        Me.tmrCountdown.Enabled = False
        StartTimer
    End If
End Sub
