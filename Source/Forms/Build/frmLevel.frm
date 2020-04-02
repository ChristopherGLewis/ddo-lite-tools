VERSION 5.00
Begin VB.Form frmLevel 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Exchange Feat"
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3552
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
   ScaleHeight     =   1260
   ScaleWidth      =   3552
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   660
      Width           =   1032
   End
   Begin VB.CheckBox chkOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   660
      Width           =   1032
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   300
      Left            =   2340
      TabIndex        =   1
      Top             =   180
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   529
      Appearance3D    =   -1  'True
      Max             =   30
      Value           =   30
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   -2147483631
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Shape shpBorder 
      Height          =   192
      Left            =   0
      Top             =   60
      Width           =   192
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Exchange feat at level:"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   180
      TabIndex        =   0
      Top             =   216
      Width           =   2040
   End
End
Attribute VB_Name = "frmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOverride As Boolean

Private Sub Form_Load()
    cfg.RefreshColors Me
    With Me.usrSpinner
        .Min = glngLevel
        .Max = build.MaxLevels
        .Value = glngLevel
    End With
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    If IsOver(Me.usrSpinner.hwnd, Xpos, Ypos) Then Me.usrSpinner.WheelScroll lngValue
End Sub

Private Sub chkOK_Click()
    If UncheckButton(Me.chkOK, mblnOverride) Then Exit Sub
    glngLevel = Me.usrSpinner.Value
    Unload Me
End Sub

Private Sub chkCancel_Click()
    If UncheckButton(Me.chkCancel, mblnOverride) Then Exit Sub
    glngLevel = 0
    Unload Me
End Sub

Private Sub Form_Resize()
    Me.shpBorder.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
