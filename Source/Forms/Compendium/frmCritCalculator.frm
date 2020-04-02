VERSION 5.00
Begin VB.Form frmCritCalculator 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crit Calculator"
   ClientHeight    =   3216
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5028
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCritCalculator.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3216
   ScaleWidth      =   5028
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   780
      Width           =   1092
   End
   Begin VB.Timer tmrCalculate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3060
      Top             =   1020
   End
   Begin VB.TextBox txtExploitRange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   3300
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   612
   End
   Begin VB.TextBox txtCritRange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   3300
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   612
   End
   Begin VB.TextBox txtExploitPercent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   2160
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1032
   End
   Begin VB.TextBox txtCritPercent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   2160
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1032
   End
   Begin VB.TextBox txtRolls 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   2160
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1032
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Roll"
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   300
      Width           =   1092
   End
   Begin Compendium.userSpinner usrspnRange 
      Height          =   312
      Left            =   1740
      TabIndex        =   3
      Top             =   300
      Width           =   1032
      _ExtentX        =   1820
      _ExtentY        =   550
      Min             =   2
      Max             =   20
      Value           =   15
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin Compendium.userSpinner usrspnFort 
      Height          =   312
      Left            =   1740
      TabIndex        =   6
      Top             =   660
      Width           =   1032
      _ExtentX        =   1820
      _ExtentY        =   550
      Min             =   0
      Max             =   100
      Value           =   0
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin Compendium.userSpinner usrspnBypass 
      Height          =   312
      Left            =   1740
      TabIndex        =   8
      Top             =   1020
      Width           =   1032
      _ExtentX        =   1820
      _ExtentY        =   550
      Min             =   0
      Max             =   100
      Value           =   0
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Line lin 
      X1              =   -180
      X2              =   5040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "to 20"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   8
      Left            =   4020
      TabIndex        =   18
      Top             =   2544
      Width           =   492
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "to 20"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   7
      Left            =   4020
      TabIndex        =   14
      Top             =   2184
      Width           =   492
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Exploit Weakness"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   6
      Left            =   444
      TabIndex        =   15
      Top             =   2544
      Width           =   1560
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Regular Crits"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   5
      Left            =   876
      TabIndex        =   11
      Top             =   2184
      Width           =   1128
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Rolls"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   4
      Left            =   1596
      TabIndex        =   9
      Top             =   1824
      Width           =   408
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Crit Range"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   336
      Width           =   1224
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fortification"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   696
      Width           =   1224
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fort Bypass"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   1056
      Width           =   1224
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "to 20"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   492
   End
End
Attribute VB_Name = "frmCritCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngRolls As Long
Private mlngCrit As Long
Private mlngControl As Long
Private mlngTest As Long
Private mlngMiss As Long
Private mlngFort As Long

Private mblnOverride As Boolean

Private Sub Form_Load()
    cfg.Configure Me
    If Not XP.DebugMode Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not XP.DebugMode Then Call WheelUnHook(Me.Hwnd)
    cfg.SavePosition Me
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.usrspnRange.Hwnd, Xpos, Ypos): Me.usrspnRange.WheelScroll lngValue
        Case IsOver(Me.usrspnFort.Hwnd, Xpos, Ypos): Me.usrspnFort.WheelScroll lngValue
        Case IsOver(Me.usrspnBypass.Hwnd, Xpos, Ypos): Me.usrspnBypass.WheelScroll lngValue
    End Select
End Sub

Private Sub chkHelp_Click()
    If UncheckButton(Me.chkHelp, mblnOverride) Then Exit Sub
    ShowHelp "Crit_Calculator"
End Sub

Private Sub chkButton_Click()
    If UncheckButton(Me.chkButton, mblnOverride) Then Exit Sub
    If Me.chkButton.Caption = "Roll" Then
        EnableControls False
        mlngRolls = 0
        mlngControl = 0
        mlngTest = 0
        mlngMiss = 0
        mlngCrit = Me.usrspnRange.Value
        mlngFort = Me.usrspnFort.Value - Me.usrspnBypass.Value
        If mlngFort < 0 Then mlngFort = 0
        If mlngFort > 100 Then mlngFort = 100
        Me.chkButton.Caption = "Stop"
        Me.tmrCalculate.Enabled = True
    Else
        Me.tmrCalculate.Enabled = False
        Me.chkButton.Caption = "Roll"
        EnableControls True
    End If
End Sub

Private Sub EnableControls(pblnEnabled As Boolean)
    Me.usrspnRange.Enabled = pblnEnabled
    Me.usrspnFort.Enabled = pblnEnabled
    Me.usrspnBypass.Enabled = pblnEnabled
End Sub

Private Sub tmrCalculate_Timer()
    Dim lngRoll As Long
    Dim lngFortRoll As Long
    Dim dblPercent As Double
    Dim lngRange As Long
    Dim i As Long
    
    Me.tmrCalculate.Enabled = False
    For i = 1 To 100
        mlngRolls = mlngRolls + 1
        lngRoll = Int(20 * Rnd + 1)
        If lngRoll = 1 Then
            mlngMiss = mlngMiss + 1
        Else
            If lngRoll >= mlngCrit - mlngMiss Then
                mlngMiss = 0
                lngFortRoll = Int(100 * Rnd + 1)
                If lngFortRoll >= mlngFort Then mlngTest = mlngTest + 1
            Else
                mlngMiss = mlngMiss + 1
            End If
            If lngRoll >= mlngCrit And lngFortRoll > mlngFort Then mlngControl = mlngControl + 1
        End If
    Next
    Me.txtRolls.Text = Format(mlngRolls, "0,000")
    ' Control
    dblPercent = mlngControl / mlngRolls
    Me.txtCritPercent.Text = Format(dblPercent, "0.00%")
    lngRange = 21 - (dblPercent * 20)
    Me.txtCritRange.Text = lngRange
    ' Test
    dblPercent = mlngTest / mlngRolls
    Me.txtExploitPercent.Text = Format(dblPercent, "0.00%")
    lngRange = 21 - (dblPercent * 20)
    Me.txtExploitRange.Text = lngRange
    Me.tmrCalculate.Enabled = True
End Sub

