VERSION 5.00
Begin VB.Form frmColorPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Named Colors"
   ClientHeight    =   6300
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   12048
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   12048
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBackground 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Background"
      ForeColor       =   &H80000008&
      Height          =   1752
      Left            =   6960
      TabIndex        =   18
      Tag             =   "ctl"
      Top             =   240
      Width           =   2412
      Begin VB.PictureBox picDisabled 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1092
         Left            =   1260
         ScaleHeight     =   1092
         ScaleWidth      =   852
         TabIndex        =   28
         Tag             =   "ctl"
         Top             =   420
         Width           =   852
         Begin VB.TextBox txtBackground 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   324
            Index           =   0
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   0
            Width           =   792
         End
         Begin VB.TextBox txtBackground 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   324
            Index           =   1
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   792
         End
         Begin VB.TextBox txtBackground 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   324
            Index           =   2
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   792
         End
      End
      Begin VB.Label lnkMatch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Match"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   1752
         TabIndex        =   19
         Tag             =   "ctl"
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Background"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Tag             =   "ctl"
         Top             =   0
         Width           =   1056
      End
      Begin VB.Shape shpBackground 
         Height          =   1632
         Left            =   0
         Top             =   120
         Width           =   2412
      End
      Begin VB.Label lblRGB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Blue"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label lblRGB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Green"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lblRGB 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Red"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   336
      End
   End
   Begin VB.Frame fraRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "RGB Values"
      ForeColor       =   &H80000008&
      Height          =   1752
      Left            =   3900
      TabIndex        =   10
      Tag             =   "ctl"
      Top             =   240
      Width           =   2652
      Begin Compendium.userSpinner usrspnValue 
         Height          =   312
         Index           =   0
         Left            =   1260
         TabIndex        =   13
         Top             =   420
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         Min             =   0
         Max             =   255
         Value           =   105
         StepSmall       =   5
         StepLarge       =   50
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   0
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin Compendium.userSpinner usrspnValue 
         Height          =   312
         Index           =   1
         Left            =   1260
         TabIndex        =   15
         Top             =   780
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         Min             =   0
         Max             =   255
         Value           =   80
         StepSmall       =   5
         StepLarge       =   50
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   0
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin Compendium.userSpinner usrspnValue 
         Height          =   312
         Index           =   2
         Left            =   1260
         TabIndex        =   17
         Top             =   1140
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         Min             =   0
         Max             =   255
         Value           =   55
         StepSmall       =   5
         StepLarge       =   50
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   0
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "RGB Values"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Tag             =   "ctl"
         Top             =   0
         Width           =   1020
      End
      Begin VB.Shape shpRGB 
         Height          =   1632
         Left            =   0
         Top             =   120
         Width           =   2652
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Medium"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   5
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   696
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   4
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   384
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   1752
      Left            =   540
      TabIndex        =   4
      Tag             =   "ctl"
      Top             =   240
      Width           =   3012
      Begin Compendium.userSpinner usrspnValue 
         Height          =   312
         Index           =   3
         Left            =   1620
         TabIndex        =   9
         Top             =   1140
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         Min             =   0
         Max             =   100
         Value           =   90
         StepSmall       =   5
         StepLarge       =   25
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   0
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin Compendium.userSpinner usrspnBrightness 
         Height          =   312
         Left            =   1620
         TabIndex        =   7
         Top             =   420
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         Min             =   0
         Max             =   255
         Value           =   105
         StepSmall       =   5
         StepLarge       =   50
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   0
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "General"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Tag             =   "ctl"
         Top             =   0
         Width           =   684
      End
      Begin VB.Shape shpGeneral 
         Height          =   1632
         Left            =   0
         Top             =   120
         Width           =   3012
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Skipped %"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Width           =   948
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Brightness"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   924
      End
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   0
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1212
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   1
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   660
      Width           =   1212
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Apply"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   2
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1140
      Width           =   1212
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   3
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1740
      Width           =   1212
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3192
      Left            =   300
      ScaleHeight     =   3192
      ScaleWidth      =   11112
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "ctl"
      Top             =   2400
      Width           =   11112
   End
End
Attribute VB_Name = "frmColorPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOriginal(3) As Long

Private mlngTop As Long
Private mlngColWidth As Long
Private mlngRowHeight As Long

Private mblnOverride As Boolean
Private mblnOKCancel As Boolean
Private mblnDirty As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    If Not XP.DebugMode Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
    If mblnDirty And Not mblnOKCancel Then
        Select Case MsgBox("Apply changes?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Notice")
            Case vbYes: ApplyColors
            Case vbNo: RestoreColors
            Case vbCancel: Cancel = 1
        End Select
    End If
    If Not XP.DebugMode Then Call WheelUnHook(Me.Hwnd)
End Sub

Public Sub ReDrawForm()
    Me.BackColor = cfg.GetColor(cgeControls, cveBackground)
    mblnDirty = False
    mblnOKCancel = False
    SizeForm
    CenterFrames
    LoadData
    DrawGrid
End Sub

Private Sub LoadData()
    Dim lngRGB(2) As Long
    Dim i As Long
    
    SaveOriginal
    mblnOverride = True
    XP.ColorToRGB cfg.GetColor(cgeControls, cveBackground), lngRGB(0), lngRGB(1), lngRGB(2)
    Me.usrspnValue(0).Value = cfg.NamedHigh
    Me.usrspnValue(1).Value = cfg.NamedMed
    Me.usrspnValue(2).Value = cfg.NamedLow
    Me.usrspnValue(3).Value = cfg.NamedDim
    For i = 0 To 2
        Me.lblRGB(i).ForeColor = cfg.GetColor(cgeControls, cveTextDim)
        Me.txtBackground(i).ForeColor = cfg.GetColor(cgeControls, cveTextDim)
        Me.txtBackground(i).Text = lngRGB(i)
    Next
    SetBrightness
    mblnOverride = False
End Sub

Private Sub SaveOriginal()
    mlngOriginal(0) = cfg.NamedHigh
    mlngOriginal(1) = cfg.NamedMed
    mlngOriginal(2) = cfg.NamedLow
    mlngOriginal(3) = cfg.NamedDim
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.usrspnBrightness.Hwnd, Xpos, Ypos): Me.usrspnBrightness.WheelScroll lngValue
        Case IsOver(Me.usrspnValue(0).Hwnd, Xpos, Ypos): Me.usrspnValue(0).WheelScroll lngValue
        Case IsOver(Me.usrspnValue(1).Hwnd, Xpos, Ypos): Me.usrspnValue(1).WheelScroll lngValue
        Case IsOver(Me.usrspnValue(2).Hwnd, Xpos, Ypos): Me.usrspnValue(2).WheelScroll lngValue
        Case IsOver(Me.usrspnValue(3).Hwnd, Xpos, Ypos): Me.usrspnValue(3).WheelScroll lngValue
    End Select
End Sub


' ************* SIZING *************


Private Sub SizeForm()
    Dim lngWidth As Long
    Dim i As Long
    
    mlngColWidth = 0
    For i = 1 To gceColors - 2
        lngWidth = Me.picPreview.TextWidth(GetColorName(i))
        If mlngColWidth < lngWidth Then mlngColWidth = lngWidth
    Next
    mlngColWidth = mlngColWidth + PixelX * 4
    mlngRowHeight = Me.picPreview.TextHeight("Q") + PixelY * 2
    mlngTop = 52 * PixelY
    Me.picPreview.Width = mlngColWidth * (gceColors - 2) + PixelX
    Me.picPreview.Height = mlngTop + mlngRowHeight * 10 + PixelY
    For i = 0 To 3
        Me.chkButton(i).Left = Me.picPreview.Left + Me.picPreview.Width - Me.chkButton(i).Width
    Next
    Me.Width = Me.picPreview.Left * 2 + Me.picPreview.Width + Me.Width - Me.ScaleWidth
    Me.Height = Me.picPreview.Top + Me.picPreview.Height + mlngRowHeight + Me.Height - Me.ScaleHeight
End Sub

Private Sub CenterFrames()
    Dim lngMargin As Long
    Dim lngLeft As Long
    
    lngMargin = (Me.chkButton(0).Left - Me.fraGeneral.Width - Me.fraRGB.Width - Me.fraBackground.Width) \ 4
    lngLeft = lngMargin
    Me.fraGeneral.Left = lngLeft
    lngLeft = lngLeft + Me.fraGeneral.Width + lngMargin
    Me.fraRGB.Left = lngLeft
    lngLeft = lngLeft + Me.fraRGB.Width + lngMargin
    Me.fraBackground.Left = lngLeft
End Sub


' ************* DRAWING *************


Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    
    Me.picPreview.Cls
    For lngCol = 1 To gceColors - 2
        Me.picPreview.CurrentX = (lngCol - 1) * mlngColWidth + (mlngColWidth - Me.picPreview.TextWidth(GetColorName(lngCol))) \ 2
        Me.picPreview.CurrentY = mlngTop + PixelY
        Me.picPreview.Print GetColorName(lngCol)
        ColorBars lngCol
        For lngRow = 1 To 10
            lngLeft = (lngCol - 1) * mlngColWidth
            lngTop = lngRow * mlngRowHeight + mlngTop
            lngRight = lngLeft + mlngColWidth
            lngBottom = lngTop + mlngRowHeight
            lngColor = GetColorValue(lngCol)
            If lngRow > 3 And lngRow < 7 Then lngColor = GetColorDim(lngColor)
            Me.picPreview.FillColor = lngColor
            Me.picPreview.Line (lngLeft, lngTop)-(lngRight, lngBottom), cfg.GetColor(cgeControls, cveBorderInterior), B
        Next
    Next
End Sub

Private Sub ColorBars(plngCol As Long)
    Dim lngColor(1 To 3) As Long
    Dim lngBar(1 To 3) As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim i As Long
    
    lngWidth = mlngColWidth \ 7
    lngBottom = mlngTop
    XP.ColorToRGB GetColorValue(plngCol), lngColor(1), lngColor(2), lngColor(3)
    lngBar(1) = cfg.GetColor(cgeControls, cveRed)
    lngBar(2) = cfg.GetColor(cgeControls, cveGreen)
    lngBar(3) = cfg.GetColor(cgeControls, cveBlue)
    For i = 1 To 3
        lngHeight = Me.picPreview.ScaleY(lngColor(i) \ 5, vbPixels, vbTwips)
        lngTop = lngBottom - lngHeight
        lngLeft = (plngCol - 1) * mlngColWidth + (i * 2 - 1) * lngWidth
        lngRight = lngLeft + lngWidth
        Me.picPreview.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngBar(i), BF
    Next
End Sub


' ************* EDITING *************


Private Property Let Dirty(pblnDirty As Boolean)
    mblnDirty = pblnDirty
    Me.chkButton(2).Enabled = mblnDirty
End Property

Private Sub usrspnBrightness_Change()
    Dim lngChange As Long
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    Dirty = True
    mblnOverride = True
    lngChange = Me.usrspnBrightness.Value - Me.usrspnValue(0).Value
    For i = 0 To 2
        Me.usrspnValue(i).Value = Me.usrspnValue(i).Value + lngChange
    Next
    cfg.NamedHigh = Me.usrspnValue(0).Value
    cfg.NamedMed = Me.usrspnValue(1).Value
    cfg.NamedLow = Me.usrspnValue(2).Value
    mblnOverride = False
    DrawGrid
End Sub

Private Sub usrspnValue_Change(Index As Integer)
    Dim lngValue As Long
    Dim lngMin As Long
    Dim lngMax As Long
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    If Index = 3 Then
        cfg.NamedDim = Me.usrspnValue(3).Value
        Dirty = True
        DrawGrid
        Exit Sub
    End If
    Dirty = True
    mblnOverride = True
    For i = Index + 1 To 2
        If Me.usrspnValue(i).Value > Me.usrspnValue(Index).Value Then Me.usrspnValue(i).Value = Me.usrspnValue(Index).Value
    Next
    For i = Index - 1 To 0 Step -1
        If Me.usrspnValue(i).Value < Me.usrspnValue(Index).Value Then Me.usrspnValue(i).Value = Me.usrspnValue(Index).Value
    Next
    cfg.NamedHigh = Me.usrspnValue(0).Value
    cfg.NamedMed = Me.usrspnValue(1).Value
    cfg.NamedLow = Me.usrspnValue(2).Value
    SetBrightness
    mblnOverride = False
    DrawGrid
End Sub

Private Sub SetBrightness()
    Me.usrspnBrightness.Min = Me.usrspnValue(0).Value - Me.usrspnValue(2).Value
    Me.usrspnBrightness.Value = Me.usrspnValue(0).Value
End Sub

Private Sub lnkMatch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XP.SetMouseCursor mcHand
End Sub

Private Sub lnkMatch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XP.SetMouseCursor mcHand
End Sub

Private Sub lnkMatch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XP.SetMouseCursor mcHand
    If mblnOverride Then Exit Sub
    MatchColors False
    mblnOverride = True
    Me.usrspnBrightness.Value = cfg.NamedHigh
    Me.usrspnValue(0).Value = cfg.NamedHigh
    Me.usrspnValue(1).Value = cfg.NamedMed
    Me.usrspnValue(2).Value = cfg.NamedLow
    Me.usrspnValue(3).Value = cfg.NamedDim
    mblnOverride = False
    DrawGrid
    Dirty = True
End Sub


' ************* BUTTONS *************


Private Sub chkButton_Click(Index As Integer)
    Dim i As Long
    
    If UncheckButton(Me.chkButton(Index), mblnOverride) Then Exit Sub
    Me.chkButton(Index).Refresh
    Select Case Me.chkButton(Index).Caption
        Case "OK"
            If mblnDirty Then ApplyColors
            Unload Me
        Case "Cancel"
            RestoreColors
            Unload Me
        Case "Apply"
            ApplyColors
            Me.chkButton(0).SetFocus
        Case "Help"
            ShowHelp "Named Colors"
    End Select
End Sub

Private Sub ApplyColors()
    Dim i As Long
    
    For i = 1 To db.Characters
        With db.Character(i)
            If Not .CustomColor Then
                .BackColor = GetColorValue(.GeneratedColor)
                .DimColor = GetColorDim(.BackColor)
            End If
        End With
    Next
    frmCompendium.RedrawQuests
    Dirty = False
    SaveOriginal
    DirtyFlag dfeSettings
End Sub

Private Sub RestoreColors()
    cfg.NamedHigh = mlngOriginal(0)
    cfg.NamedMed = mlngOriginal(1)
    cfg.NamedLow = mlngOriginal(2)
    cfg.NamedDim = mlngOriginal(3)
    Dirty = False
End Sub
