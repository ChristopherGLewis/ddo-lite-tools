VERSION 5.00
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6420
   ClientLeft      =   30
   ClientTop       =   405
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7110
   Begin CharacterBuilderLite.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7032
      _ExtentX        =   12409
      _ExtentY        =   688
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin VB.Frame fraAppearance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2592
      Left            =   300
      TabIndex        =   10
      Top             =   3480
      Width           =   3672
      Begin CharacterBuilderLite.userCheckBox usrchkIconSkills 
         Height          =   252
         Left            =   600
         TabIndex        =   14
         Top             =   1200
         Width           =   2292
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "Skills Screen"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkIconOverview 
         Height          =   252
         Left            =   600
         TabIndex        =   13
         Top             =   840
         Width           =   2292
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "Overview Screen"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkUseIcons 
         Height          =   252
         Left            =   300
         TabIndex        =   12
         Top             =   480
         Width           =   2592
         _ExtentX        =   4577
         _ExtentY        =   450
         Caption         =   "Use Class Icons"
      End
      Begin VB.CheckBox chkColors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Colors"
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2100
         Width           =   1392
      End
      Begin CharacterBuilderLite.userSpinner usrspnOutputMargin 
         Height          =   300
         Left            =   1632
         TabIndex        =   16
         Top             =   1620
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   529
         Appearance3D    =   -1  'True
         Min             =   0
         Value           =   0
         StepLarge       =   3
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Spaces"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   8
         Left            =   2700
         TabIndex        =   17
         Top             =   1656
         Width           =   636
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Indent Output"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   7
         Left            =   300
         TabIndex        =   15
         Top             =   1656
         Width           =   1236
      End
      Begin VB.Label lblAppearance 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Appearance"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   1044
      End
      Begin VB.Shape shpAppearance 
         Height          =   2472
         Left            =   60
         Top             =   120
         Width           =   3612
      End
   End
   Begin VB.Frame fraGeneral 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "General"
      ForeColor       =   &H80000008&
      Height          =   2712
      Left            =   300
      TabIndex        =   1
      Top             =   600
      Width           =   3732
      Begin VB.ComboBox cboWindow 
         Height          =   312
         ItemData        =   "frmOptions.frx":000C
         Left            =   1620
         List            =   "frmOptions.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1692
      End
      Begin CharacterBuilderLite.userSpinner usrspnMRU 
         Height          =   300
         Left            =   2100
         TabIndex        =   6
         Top             =   840
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   529
         Appearance3D    =   -1  'True
         Max             =   9
         Value           =   9
         StepLarge       =   3
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userCheckBox usrchkErrorLog 
         Height          =   252
         Left            =   300
         TabIndex        =   7
         Top             =   1320
         Width           =   3192
         _ExtentX        =   5636
         _ExtentY        =   450
         Caption         =   "Show Error Log on startup"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkChildWindows 
         Height          =   252
         Left            =   300
         TabIndex        =   8
         Top             =   1740
         Width           =   3192
         _ExtentX        =   5636
         _ExtentY        =   450
         Caption         =   "Child Windows"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkConfirm 
         Height          =   252
         Left            =   300
         TabIndex        =   9
         Top             =   2160
         Width           =   3192
         _ExtentX        =   5636
         _ExtentY        =   450
         Caption         =   "Confirmation Prompts"
      End
      Begin VB.Label lblGeneral 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "General"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   684
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Main Window"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   300
         TabIndex        =   3
         Top             =   408
         Width           =   1176
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Recent Files Menu"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   876
         Width           =   1596
      End
      Begin VB.Shape shpGeneral 
         Height          =   2592
         Left            =   0
         Top             =   120
         Width           =   3732
      End
   End
   Begin VB.Frame fraOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Output"
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   4260
      TabIndex        =   27
      Top             =   3480
      Width           =   2532
      Begin VB.ComboBox cboFeatOrderOutput 
         Height          =   312
         ItemData        =   "frmOptions.frx":0045
         Left            =   300
         List            =   "frmOptions.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1380
         Width           =   1932
      End
      Begin VB.CheckBox chkFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "BBCodes"
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2280
         Width           =   1392
      End
      Begin VB.ComboBox cboSkillOrderOutput 
         Height          =   312
         ItemData        =   "frmOptions.frx":0072
         Left            =   300
         List            =   "frmOptions.frx":007C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   660
         Width           =   1932
      End
      Begin CharacterBuilderLite.userCheckBox usrchkFeatsShowAll 
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "Show All Feats"
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Feat Order"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   6
         Left            =   300
         TabIndex        =   31
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Skill Order"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   5
         Left            =   300
         TabIndex        =   29
         Top             =   420
         Width           =   900
      End
      Begin VB.Label lblOutput 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Output"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   120
         TabIndex        =   28
         Top             =   0
         Width           =   612
      End
      Begin VB.Shape shpOutput 
         Height          =   2715
         Left            =   0
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Frame fraBuild 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Build"
      ForeColor       =   &H80000008&
      Height          =   2712
      Left            =   4260
      TabIndex        =   19
      Top             =   600
      Width           =   2532
      Begin VB.ComboBox cboSkillOrderScreen 
         Height          =   312
         ItemData        =   "frmOptions.frx":009B
         Left            =   300
         List            =   "frmOptions.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1380
         Width           =   1932
      End
      Begin VB.ComboBox cboBuildPoints 
         Height          =   312
         ItemData        =   "frmOptions.frx":00C5
         Left            =   300
         List            =   "frmOptions.frx":00D5
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   660
         Width           =   1932
      End
      Begin VB.ComboBox cboFeatOrder 
         Height          =   312
         ItemData        =   "frmOptions.frx":00FD
         Left            =   300
         List            =   "frmOptions.frx":0107
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2100
         Width           =   1932
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Skill Order"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   3
         Left            =   300
         TabIndex        =   23
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Default Build Points"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   300
         TabIndex        =   21
         Top             =   420
         Width           =   1692
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Feat Order"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   4
         Left            =   300
         TabIndex        =   25
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label lblBuild 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Build"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   408
      End
      Begin VB.Shape shpBuild 
         Height          =   2592
         Left            =   0
         Top             =   120
         Width           =   2532
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private mblnOverride As Boolean
Private mlngIncrement As Long


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    LoadData
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Activate()
    ActivateForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
    cfg.SaveSettings
    frmMain.UpdateWindowMenu
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.usrspnMRU.hwnd, Xpos, Ypos): Me.usrspnMRU.WheelScroll lngValue
        Case IsOver(Me.usrspnOutputMargin.hwnd, Xpos, Ypos): Me.usrspnOutputMargin.WheelScroll lngValue
    End Select
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Options"
    End Select
End Sub


' ************* INITIALIZE *************


Private Sub LoadData()
    Dim i As Long
    
    mblnOverride = True
    ' General
    ComboSetValue Me.cboWindow, cfg.WindowSize
    Me.usrchkErrorLog.Value = cfg.ShowErrors
    Me.usrchkChildWindows.Value = cfg.ChildWindows
    Me.usrchkConfirm.Value = cfg.Confirm
    Me.usrspnMRU.Value = cfg.MRUCount
    ' Build
    ComboSetValue Me.cboBuildPoints, cfg.BuildPoints
    ComboSetValue Me.cboSkillOrderScreen, cfg.SkillOrderScreen
    ComboSetValue Me.cboFeatOrder, cfg.FeatOrder
    ComboSetValue Me.cboFeatOrderOutput, cfg.FeatOrderOutput
    ' Output
    ComboSetValue Me.cboSkillOrderOutput, cfg.SkillOrderOutput
    ' Appearance
    Me.usrchkUseIcons.Value = cfg.UseIcons
    Me.usrchkIconOverview.Value = cfg.IconOverview
    Me.usrchkIconOverview.Enabled = Me.usrchkUseIcons.Value
    Me.usrchkIconSkills.Value = cfg.IconSkills
    Me.usrchkIconSkills.Enabled = Me.usrchkUseIcons.Value
    Me.usrspnOutputMargin.Value = cfg.OutputMargin
    
    Me.usrchkFeatsShowAll.Value = cfg.FeatShowAll
    
    mblnOverride = False
End Sub


' ************* GENERAL *************


Private Sub cboWindow_Click()
    If mblnOverride Or Me.cboWindow.ListIndex = -1 Then Exit Sub
    cfg.WindowSize = ComboGetValue(Me.cboWindow)
End Sub

Private Sub usrspnMRU_Change()
    If mblnOverride Then Exit Sub
    cfg.MRUCount = Me.usrspnMRU.Value
    EnableMenus
End Sub

Private Sub usrspnOutputMargin_Change()
    If mblnOverride Then Exit Sub
    cfg.OutputMargin = Me.usrspnOutputMargin.Value
    ResizeOutput
End Sub

Private Sub cboBuildPoints_Click()
    If mblnOverride Or Me.cboBuildPoints.ListIndex = -1 Then Exit Sub
    cfg.BuildPoints = ComboGetValue(Me.cboBuildPoints)
End Sub

Private Sub cboSkillOrderScreen_Click()
    Dim frm As Form
    
    If mblnOverride Or Me.cboSkillOrderScreen.ListIndex = -1 Then Exit Sub
    cfg.SkillOrderScreen = ComboGetValue(Me.cboSkillOrderScreen)
    If GetForm(frm, "frmSkills") Then frm.Cascade
End Sub

Private Sub cboSkillOrderOutput_Click()
    If mblnOverride Or Me.cboSkillOrderOutput.ListIndex = -1 Then Exit Sub
    cfg.SkillOrderOutput = ComboGetValue(Me.cboSkillOrderOutput)
    GenerateOutput oeAll
End Sub

Private Sub cboFeatOrder_Click()
    Dim frm As Form
    
    If mblnOverride Or Me.cboFeatOrder.ListIndex = -1 Then Exit Sub
    cfg.FeatOrder = ComboGetValue(Me.cboFeatOrder)
    IndexFeatDisplay
    If GetForm(frm, "frmFeats") Then frm.OrderChanged
End Sub

Private Sub cboFeatOrderOutput_Click()
    If mblnOverride Or Me.cboFeatOrderOutput.ListIndex = -1 Then Exit Sub
    cfg.FeatOrderOutput = ComboGetValue(Me.cboFeatOrderOutput)
    GenerateOutput oeAll
End Sub

Private Sub usrchkErrorLog_UserChange()
    If mblnOverride Then Exit Sub
    cfg.ShowErrors = Me.usrchkErrorLog.Value
End Sub

Private Sub usrchkChildWindows_UserChange()
    If mblnOverride Then Exit Sub
    cfg.ChildWindows = Me.usrchkChildWindows.Value
End Sub

Private Sub usrchkConfirm_UserChange()
    If mblnOverride Then Exit Sub
    cfg.Confirm = Me.usrchkConfirm.Value
End Sub

Private Sub chkFormat_Click()
    If UncheckButton(Me.chkFormat, mblnOverride) Then Exit Sub
    OpenForm "frmFormat"
End Sub

Private Sub chkColors_Click()
    If UncheckButton(Me.chkColors, mblnOverride) Then Exit Sub
    cfg.RunUtil ueColors
End Sub

Private Sub usrchkUseIcons_UserChange()
    If mblnOverride Then Exit Sub
    cfg.UseIcons = Me.usrchkUseIcons.Value
    Me.usrchkIconOverview.Enabled = Me.usrchkUseIcons.Value
    Me.usrchkIconSkills.Enabled = Me.usrchkUseIcons.Value
    UpdateIcons
End Sub

Private Sub usrchkIconOverview_UserChange()
    If mblnOverride Then Exit Sub
    cfg.IconOverview = Me.usrchkIconOverview.Value
    UpdateIcons
End Sub

Private Sub usrchkIconSkills_UserChange()
    If mblnOverride Then Exit Sub
    cfg.IconSkills = Me.usrchkIconSkills.Value
    UpdateIcons
End Sub

Private Sub UpdateIcons()
    Dim frm As Form
    
    If GetForm(frm, "frmOverview") Then frm.RefreshIcons
    If GetForm(frm, "frmSkills") Then frm.RefreshIcons
End Sub

Private Sub usrchkFeatsShowAll_UserChange()
    cfg.FeatShowAll = Me.usrchkFeatsShowAll.Value
End Sub
