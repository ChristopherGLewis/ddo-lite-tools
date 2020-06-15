VERSION 5.00
Begin VB.Form frmStats 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stats"
   ClientHeight    =   7770
   ClientLeft      =   30
   ClientTop       =   405
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12225
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   7380
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "< Overview"
      CenterLink      =   "Clear Stats"
      RightLinks      =   "Skills >"
   End
   Begin CharacterBuilderLite.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   688
      Spacing         =   264
      BorderColor     =   -2147483640
      LeftLinks       =   "Stats;Schedule"
      RightLinks      =   "Help"
   End
   Begin VB.Frame fraGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2772
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   4440
      Width           =   3492
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   0
         Left            =   1920
         TabIndex        =   31
         Top             =   480
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   0
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   1
         Left            =   1920
         TabIndex        =   33
         Top             =   960
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   1
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   35
         Top             =   1200
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   2
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   37
         Top             =   1440
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   2
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   4
         Left            =   1920
         TabIndex        =   39
         Top             =   1680
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   2
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   5
         Left            =   1920
         TabIndex        =   41
         Top             =   1920
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   2
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin CharacterBuilderLite.userSpinner usrSpinner 
         Height          =   252
         Index           =   6
         Left            =   1920
         TabIndex        =   43
         Top             =   2160
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         Min             =   0
         Max             =   6
         Value           =   0
         StepLarge       =   1
         ShowZero        =   0   'False
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderColor     =   -2147483631
         BorderInterior  =   -2147483631
         Position        =   3
         Enabled         =   -1  'True
         DisabledColor   =   -2147483631
      End
      Begin VB.Label lblLimitTomes 
         BackColor       =   &H80000005&
         Caption         =   "Tomes locked to max level"
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   2500
         Width           =   2655
      End
      Begin VB.Label lblGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tomes"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   0
         Width           =   612
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Strength"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   1332
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Dexterity"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   34
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Constitution"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   36
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Intelligence"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   360
         TabIndex        =   38
         Top             =   1680
         Width           =   1332
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wisdom"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   5
         Left            =   360
         TabIndex        =   40
         Top             =   1920
         Width           =   1332
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Charisma"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   6
         Left            =   360
         TabIndex        =   42
         Top             =   2160
         Width           =   1332
      End
      Begin VB.Label lblAbility 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Supreme"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   1332
      End
      Begin VB.Shape shpGroup 
         Height          =   2652
         Index           =   2
         Left            =   0
         Top             =   120
         Width           =   3492
      End
   End
   Begin VB.Frame fraGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3732
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7932
      Begin VB.ComboBox cboBuildPoints 
         Height          =   312
         ItemData        =   "frmStats.frx":000C
         Left            =   2520
         List            =   "frmStats.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1572
      End
      Begin CharacterBuilderLite.userStats usrStats 
         Height          =   2292
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1752
         _ExtentX        =   3096
         _ExtentY        =   4048
      End
      Begin CharacterBuilderLite.userStats usrStats 
         Height          =   2292
         Index           =   1
         Left            =   2280
         TabIndex        =   7
         Top             =   1080
         Width           =   1752
         _ExtentX        =   3096
         _ExtentY        =   4048
      End
      Begin CharacterBuilderLite.userStats usrStats 
         Height          =   2292
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   1080
         Width           =   1752
         _ExtentX        =   3096
         _ExtentY        =   4048
      End
      Begin CharacterBuilderLite.userStats usrStats 
         Height          =   2292
         Index           =   3
         Left            =   6120
         TabIndex        =   9
         Top             =   1080
         Width           =   1752
         _ExtentX        =   3096
         _ExtentY        =   4048
      End
      Begin VB.Label lblGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Base Stats"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   960
      End
      Begin VB.Label lblBuildPoints 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Preferred Build Points:"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   360
         TabIndex        =   3
         Top             =   516
         Width           =   2412
      End
      Begin VB.Label lblPoints 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4200
         TabIndex        =   5
         Top             =   516
         Width           =   1572
      End
      Begin VB.Shape shpGroup 
         Height          =   3612
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   7932
      End
   End
   Begin VB.Frame fraGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2772
      Index           =   3
      Left            =   3960
      TabIndex        =   44
      Top             =   4440
      Width           =   7932
      Begin VB.CheckBox chkTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Template 4"
         ForeColor       =   &H80000008&
         Height          =   432
         Index           =   3
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2040
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.CheckBox chkTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Template 2"
         ForeColor       =   &H80000008&
         Height          =   432
         Index           =   1
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.CheckBox chkTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Template 3"
         ForeColor       =   &H80000008&
         Height          =   432
         Index           =   2
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1500
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.CheckBox chkTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Template 1"
         ForeColor       =   &H80000008&
         Height          =   432
         Index           =   0
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   420
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Timer tmrFocus 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4800
         Top             =   1800
      End
      Begin VB.Timer tmrWarning 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   4740
         Top             =   780
      End
      Begin VB.Label lnkCustomize 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Customize"
         ForeColor       =   &H00FF0000&
         Height          =   216
         Left            =   6804
         TabIndex        =   57
         Top             =   0
         Width           =   960
      End
      Begin VB.Label lblWarning 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "tmrWarning: Show warning if template doesn't meet all reqs"
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   5220
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   2352
      End
      Begin VB.Label lblTimer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "tmrFocus: To force user controls to lose focus to MsgBox"
         ForeColor       =   &H80000008&
         Height          =   732
         Left            =   5280
         TabIndex        =   55
         Top             =   1680
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label lblTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description of Template 4"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   1920
         TabIndex        =   52
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   5892
      End
      Begin VB.Label lblTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description of Template 3"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   1920
         TabIndex        =   50
         Top             =   1620
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   5892
      End
      Begin VB.Label lblTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description of Template 2"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   1920
         TabIndex        =   48
         Top             =   1080
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   5892
      End
      Begin VB.Label lblTemplate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Description of Template 1"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   1920
         TabIndex        =   46
         Top             =   540
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   5892
      End
      Begin VB.Label lnkTemplate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Templates"
         ForeColor       =   &H00FF0000&
         Height          =   216
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   924
      End
      Begin VB.Shape shpGroup 
         Height          =   2652
         Index           =   3
         Left            =   0
         Top             =   120
         Width           =   7932
      End
   End
   Begin VB.Frame fraGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3732
      Index           =   1
      Left            =   8400
      TabIndex        =   10
      Top             =   600
      Width           =   3492
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   7
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3120
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   6
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2760
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   5
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2400
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   4
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2040
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   3
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1320
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   1932
      End
      Begin VB.ComboBox cboLevelup 
         Height          =   312
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   480
         Width           =   1932
      End
      Begin VB.Label lblGroup 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Levelups"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   768
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 28:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   3156
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 24:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   2796
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 20:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   22
         Top             =   2436
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 16:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   2076
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 12:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1716
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 8:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1356
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level 4:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   996
         Width           =   1092
      End
      Begin VB.Label lblLevelup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "All Levels:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   516
         Width           =   1092
      End
      Begin VB.Shape shpGroup 
         Height          =   3612
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   3492
      End
   End
   Begin VB.PictureBox picSchedule 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6792
      Left            =   120
      ScaleHeight     =   6795
      ScaleWidth      =   11955
      TabIndex        =   58
      Top             =   480
      Visible         =   0   'False
      Width           =   11952
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuTemplate 
         Caption         =   "Class Name"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Customize"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuCustomize 
         Caption         =   "Edit Custom Templates"
         Index           =   0
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "Reload Templates"
         Index           =   1
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "Delete Custom Templates"
         Index           =   3
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuCustomize 
         Caption         =   "Edit System Templates"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private menTemplateClass As ClassEnum

Private Type GridDimensionsType
    CellLeft(30) As Long
    CellWidth As Long
    HeaderWidth As Long
    RowHeight As Long
    OffsetX1 As Long
    OffsetX2 As Long
    OffsetY As Long
    Top As Long
End Type

Private grid As GridDimensionsType

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.Configure Me
    InitTemplatesMenu
    Cascade
End Sub

Private Sub Form_Activate()
    ActivateForm oeStats
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me, mblnOverride
End Sub

Public Sub Cascade()
    LoadData
    EnforceLevelMax
    DefaultTemplates
    If Me.picSchedule.Visible Then ShowSchedule
End Sub

Private Sub StatsChanged()
    CascadeChanges cceStats
    SetDirty
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Stats": HideSchedule
        Case "Schedule": ShowSchedule
        Case "Help": ShowHelp "Stats"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "Clear Stats"
            ResetStats
            Exit Sub
        Case "< Overview"
            OpenForm "frmOverview"
        Case "Skills >"
            If Not OpenForm("frmSkills") Then Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub

Private Sub ResetStats()
On Error Resume Next
    Me.usrFooter.SetFocus
    Me.tmrFocus.Enabled = True
End Sub

Private Sub tmrFocus_Timer()
    Dim blnTomes As Boolean
    Dim i As Long
    
    Me.tmrFocus.Enabled = False
    For i = 1 To 6
        If build.Tome(i) > 0 Then
            blnTomes = True
            Exit For
        End If
    Next
    If blnTomes Then
        Select Case MsgBox("Keep tomes?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Clear Stats")
            Case vbYes: blnTomes = False
            Case vbNo: blnTomes = True
            Case vbCancel: Exit Sub
        End Select
    Else
        If Not Ask("Clear stats and levelups?") Then Exit Sub
    End If
    If blnTomes Then Erase build.Tome
    Erase build.StatPoints
    Erase build.Levelups
    For i = 0 To 3
        build.IncludePoints(i) = 1
    Next
    build.BuildPoints = cfg.BuildPoints
    If build.Race = reDrow Then build.IncludePoints(1) = 0
    ShowTemplates
    LoadData
    EnforceLevelMax
    StatsChanged
    On Error Resume Next
    Me.cboBuildPoints.SetFocus
End Sub


' ************* INITIALIZE *************


Private Sub InitTemplatesMenu()
    Dim i As Long
    
    For i = 1 To ceClasses - 1
        Load Me.mnuTemplate(i)
        Me.mnuTemplate(i).Caption = db.Class(i).ClassName
        Me.mnuTemplate(i).Visible = True
    Next
    If i > 1 Then Me.mnuTemplate(0).Visible = False
End Sub

Private Sub LoadData()
On Error GoTo LoadDataErr
    Dim lngIndex As Long
    Dim lngMax As Long
    Dim i As Long
    
    If Me.Visible Then xp.LockWindow Me.hwnd
    mblnOverride = True
    ' Build Points
    ComboClear Me.cboBuildPoints
    ComboAddItem Me.cboBuildPoints, "Adventurer", beAdventurer
    If build.Race <> reDrow Then ComboAddItem Me.cboBuildPoints, "Champion", beChampion
    ComboAddItem Me.cboBuildPoints, "Hero", beHero
    ComboAddItem Me.cboBuildPoints, "Legend", beLegend
    ComboSetValue Me.cboBuildPoints, build.BuildPoints
    If build.Race = reDrow Then build.IncludePoints(beChampion) = 0
    ' Stats
    For lngIndex = 0 To 3
        With Me.usrStats(lngIndex)
            .BuildPoints = lngIndex
            .Refresh
        End With
    Next
    ' Levelups
    For lngIndex = 0 To 7
        ComboClear Me.cboLevelup(lngIndex)
        For i = 0 To 6
            ComboAddItem Me.cboLevelup(lngIndex), GetStatName(i), i
        Next
        ComboSetValue Me.cboLevelup(lngIndex), build.Levelups(lngIndex)
    Next
    mblnOverride = False
    
LoadDataExit:
    If Me.Visible Then xp.UnlockWindow
    Exit Sub
    
LoadDataErr:
    Resume LoadDataExit
End Sub

Private Sub EnforceLevelMax()
    Dim blnVisible As Boolean
    Dim lngMax As Long
    Dim i As Long
    
    mblnOverride = True
    ' Tomes
    lngMax = TomeLevel(build.MaxLevels)
    If lngMax >= tomes.Stat.Max Then
        lngMax = tomes.Stat.Max
        lblLimitTomes.Visible = False
    Else
        lblLimitTomes.Visible = True
    End If
    
    blnVisible = (lngMax > 0)
    For i = 0 To 6
        Me.usrSpinner(i).Max = lngMax
        If build.Tome(i) > lngMax Then build.Tome(i) = lngMax
        Me.usrSpinner(i).Value = build.Tome(i)
        Me.lblAbility(i).Visible = blnVisible
        Me.usrSpinner(i).Visible = blnVisible
    Next
    ' Levelups
    For i = 1 To 7
        blnVisible = (build.MaxLevels >= i * 4)
        Me.lblLevelup(i).Visible = blnVisible
        Me.cboLevelup(i).Visible = blnVisible
    Next
    Me.lblLevelup(0).Visible = Me.lblLevelup(1).Visible
    Me.cboLevelup(0).Visible = Me.lblLevelup(0).Visible
    mblnOverride = False
End Sub


' ************* BUILD POINTS *************


Private Sub cboBuildPoints_Click()
    If mblnOverride Then Exit Sub
    build.BuildPoints = ComboGetValue(Me.cboBuildPoints)
    build.IncludePoints(build.BuildPoints) = 1
    Me.usrStats(build.BuildPoints).Refresh
    ShowBuildPoints
    EnforceLevelMax
    StatsChanged
End Sub

Private Sub ShowBuildPoints()
    With Me.lblPoints
        Select Case build.BuildPoints
            Case beAdventurer: .Caption = "28pt"
            Case beChampion: If build.Race = reDrow Then .Caption = "28pt" Else .Caption = "32pt"
            Case beHero: If build.Race = reDrow Then .Caption = "30pt" Else .Caption = "34pt"
            Case beLegend: If build.Race = reDrow Then .Caption = "32pt" Else .Caption = "36pt"
        End Select
    End With
End Sub

Private Sub usrStats_StatChange(Index As Integer, Stat As StatEnum, Increase As Boolean)
    Dim i As Long
    
    For i = 0 To 3
        If i <> Index Then Me.usrStats(i).IncrementStat Stat, Increase
    Next
End Sub

Private Sub usrStats_Include(Index As Integer, Include As Boolean)
    Dim lngBuildPoints As Long
    Dim i As Long
    
    If Include Or build.BuildPoints <> Index Then Exit Sub
    For i = 1 To 4
        lngBuildPoints = Index + i
        If lngBuildPoints > 3 Then lngBuildPoints = lngBuildPoints - 4
        If build.IncludePoints(lngBuildPoints) = 1 Then
            ComboSetValue Me.cboBuildPoints, lngBuildPoints
            Exit Sub
        End If
    Next
End Sub


' ************* LEVELUPS *************


Private Sub cboLevelup_Click(Index As Integer)
    Dim enStat As StatEnum
    Dim blnUniform As Boolean
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    mblnOverride = True
    build.Levelups(Index) = ComboGetValue(Me.cboLevelup(Index))
    ' Propagate "All Levels" to all the levels
    If Index = 0 Then
        For i = 1 To 7
            build.Levelups(i) = build.Levelups(0)
            ComboSetValue Me.cboLevelup(i), build.Levelups(i)
        Next
    Else ' Set "All Levels" if all levels are the same, or unset if not
        enStat = build.Levelups(1)
        blnUniform = True
        For i = 2 To 7
            If build.Levelups(i) <> enStat Then blnUniform = False
        Next
        If Not blnUniform Then enStat = aeAny
        If build.Levelups(0) <> enStat Then
            build.Levelups(0) = enStat
            ComboSetValue Me.cboLevelup(0), enStat
        End If
    End If
    mblnOverride = False
    StatsChanged
End Sub


' ************* TOMES *************


' This used to be a much larger section of code before I wrapped it into the spinner user control
Private Sub usrSpinner_Change(Index As Integer)
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    build.Tome(Index) = Me.usrSpinner(Index).Value
    For i = 1 To 6
        If build.Tome(i) < build.Tome(aeAny) Then
            build.Tome(i) = build.Tome(aeAny)
            Me.usrSpinner(i).Value = build.Tome(aeAny)
        End If
    Next
    StatsChanged
End Sub


' ************* TEMPLATES *************


Private Sub DefaultTemplates()
    Dim typClassSplit() As ClassSplitType
    
    GetClassSplit typClassSplit
    menTemplateClass = typClassSplit(0).ClassID
    ShowTemplates
End Sub

Private Sub ShowTemplates()
    Dim i As Long
    
    AddTemplates menTemplateClass
    ShowWarning -1
    For i = 1 To ceClasses - 1
        Me.mnuTemplate(i).Checked = (i = menTemplateClass)
    Next
End Sub

Private Sub AddTemplates(penClass As ClassEnum)
    Dim strTemplate() As String
    Dim blnTrapping As Boolean
    Dim strName As String
    Dim lngTemplate As Long
    Dim lngIndex As Long
    Dim i As Long
    
    ' Hide
    For i = 3 To 0 Step -1
        Me.chkTemplate(i).Visible = False
        Me.lblTemplate(i).Visible = False
    Next
    blnTrapping = IsTrapper()
    For i = 1 To db.Templates
        If db.Template(i).Class = penClass And (db.Template(i).Always = True Or db.Template(i).Trapping = blnTrapping) Then
            With Me.chkTemplate(lngIndex)
                .Tag = i
                .Caption = db.Template(i).Caption
                .Visible = True
            End With
            With Me.lblTemplate(lngIndex)
                .Caption = db.Template(i).Descrip
                .Visible = True
            End With
            lngIndex = lngIndex + 1
            If lngIndex > 3 Then Exit For
        End If
    Next
    If lngIndex <> 0 Then Me.lnkTemplate.Caption = GetClassName(penClass) & " Templates"
    Me.lnkTemplate.Visible = (lngIndex <> 0)
End Sub

Private Function IsTrapper() As Boolean
    Dim i As Long
    
    For i = 1 To 20
        If build.Class(i) = ceRogue Or build.Class(i) = ceArtificer Then
            IsTrapper = True
            Exit Function
        End If
    Next
End Function

Private Sub chkTemplate_Click(Index As Integer)
    Dim lngTemplate As Long
    
    If UncheckButton(Me.chkTemplate(Index), mblnOverride) Then Exit Sub
    lngTemplate = Me.chkTemplate(Index).Tag
    ApplyTemplate lngTemplate
    ShowWarning Index
    Me.tmrWarning.Enabled = True
End Sub

Private Sub ShowWarning(Index As Integer)
    Dim lngTemplate As Long
    Dim lngColor As Long
    Dim i As Long
    
    Me.tmrWarning.Enabled = False
    lngColor = cfg.GetColor(cgeWorkspace, cveText)
    For i = 0 To 3
        lngTemplate = Val(Me.chkTemplate(i).Tag)
        If lngTemplate Then
            With Me.lblTemplate(i)
                If .Caption <> db.Template(lngTemplate).Descrip Then .Caption = db.Template(lngTemplate).Descrip
                If .ForeColor <> lngColor Then .ForeColor = lngColor
            End With
        End If
    Next
    If Index = -1 Then Exit Sub
    lngTemplate = Val(Me.chkTemplate(Index).Tag)
    If lngTemplate = 0 Then Exit Sub
    If Len(db.Template(lngTemplate).Warning) = 0 Then Exit Sub
    With Me.lblTemplate(Index)
        .Caption = db.Template(lngTemplate).Warning
        .ForeColor = cfg.GetColor(cgeWorkspace, cveTextError)
    End With
End Sub

Private Sub tmrWarning_Timer()
    ShowWarning -1
End Sub

Private Sub ApplyTemplate(plngIndex As Long)
    Dim i As Long
    
    If plngIndex < 1 Or plngIndex > db.Templates Then
        Notice "Unknown template"
        Exit Sub
    End If
    Erase build.StatPoints
    If build.Race = reDrow Then
        ApplyStats plngIndex, 0, 0
        ApplyStats plngIndex, 0, 1
        ApplyStats plngIndex, 1, 2
        ApplyStats plngIndex, 2, 3
    Else
        ApplyStats plngIndex, 0, 0
        ApplyStats plngIndex, 2, 1
        ApplyStats plngIndex, 3, 2
        ApplyStats plngIndex, 4, 3
    End If
    For i = 0 To 7
        build.Levelups(i) = db.Template(plngIndex).Levelups
    Next
    LoadData
    EnforceLevelMax
    StatsChanged
End Sub

Private Sub ApplyStats(plngIndex As Long, plngTemp As Long, plngBuild As Long)
    Dim i As Long
    
    For i = 1 To 6
        build.StatPoints(plngBuild, i) = db.Template(plngIndex).StatPoints(plngTemp, i)
        build.StatPoints(plngBuild, 0) = build.StatPoints(plngBuild, 0) + build.StatPoints(plngBuild, i)
    Next
End Sub

Private Sub lnkTemplate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkTemplate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    PopupMenu Me.mnuMain(0)
End Sub

Private Sub lnkTemplate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub mnuTemplate_Click(Index As Integer)
    Dim strClass As String
    Dim enClass As ClassEnum
    
    strClass = Me.mnuTemplate(Index).Caption
    menTemplateClass = GetClassID(strClass)
    If menTemplateClass = ceAny Then
        Notice "Class not recognized: " & strClass
        DefaultTemplates
    Else
        ShowTemplates
    End If
End Sub


' ************* CUSTOM TEMPLATES *************


Private Sub lnkCustomize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkCustomize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    Me.mnuCustomize(3).Enabled = xp.File.Exists(UserTemplates())
    PopupMenu Me.mnuMain(1)
End Sub

Private Sub lnkCustomize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub mnuCustomize_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuCustomize(Index).Caption)
        Case "Edit Custom Templates": EditUserTemplates
        Case "Delete Custom Templates": DeleteCustomTemplates
        Case "Edit System Templates": EditSystemTemplates
        Case "Reload Templates": ReloadTemplates
    End Select
End Sub

Private Sub EditUserTemplates()
    Dim strSystem As String
    Dim strUser As String
    
    strSystem = SystemTemplates()
    strUser = UserTemplates()
    If Not xp.File.Exists(strUser) Then
        If xp.File.Exists(strSystem) Then
            xp.File.Copy strSystem, strUser
        Else
            Notice "System templates not found"
            Exit Sub
        End If
    End If
    If xp.File.Exists(strUser) Then
        xp.File.Run strUser
    Else
        Notice "User templates not found"
    End If
End Sub

Private Sub EditSystemTemplates()
    Dim strSystem As String
    
    strSystem = SystemTemplates()
    If xp.File.Exists(strSystem) Then
        xp.File.Run strSystem
    Else
        Notice "System templates not found"
    End If
End Sub

Private Sub DeleteCustomTemplates()
    Dim strUser As String
    
    strUser = UserTemplates()
    If Not xp.File.Exists(strUser) Then Exit Sub
    If Not AskAlways("Delete your custom templates?") Then Exit Sub
    xp.File.Delete strUser
    ReloadTemplates
End Sub

Private Sub ReloadTemplates()
    LoadTemplates
    ProcessData
    CreateErrorLog True
    ClearLog
    DefaultTemplates
End Sub


' ************* SCHEDULE *************


Private Sub ShowSchedule()
    Dim enStat As StatEnum
    Dim lngLevel As Long
    Dim lngTop As Long
    
    EnableStats False
    InitGrid
    Me.picSchedule.Cls
    DrawHeaderRow
    lngTop = grid.RowHeight * 2
    For enStat = aeStr To aeCha
        DrawStatRow enStat, lngTop
        DrawTomeRow enStat, lngTop
        DrawLevelupRow enStat, lngTop
        lngTop = lngTop + grid.RowHeight
    Next
    Me.picSchedule.ZOrder vbBringToFront
    Me.picSchedule.Visible = True
End Sub

Private Sub InitGrid()
    Dim lngLeft As Long
    Dim i As Long
    
    With Me.picSchedule
        grid.HeaderWidth = .TextWidth("Constitution ")
        grid.CellWidth = (.ScaleWidth - grid.HeaderWidth) \ 31
        grid.CellWidth = .ScaleX(.ScaleX(grid.CellWidth, vbTwips, vbPixels), vbPixels, vbTwips)
        grid.OffsetX1 = (grid.CellWidth - Me.TextWidth("9")) \ 2
        grid.OffsetX2 = (grid.CellWidth - Me.TextWidth("99")) \ 2
        grid.RowHeight = .TextHeight("Q") + .ScaleY(4, vbPixels, vbTwips)
        grid.OffsetY = (grid.RowHeight - .TextHeight("Q")) \ 2
    End With
    With grid
        lngLeft = .HeaderWidth
        For i = 0 To 30
            .CellLeft(i) = lngLeft
            lngLeft = lngLeft + .CellWidth
        Next
    End With
End Sub

Private Sub DrawHeaderRow()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim i As Long
    
    Me.picSchedule.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    With grid
        lngTop = .OffsetY
        For i = 0 To 30
            If i < 10 Then lngLeft = .CellLeft(i) + .OffsetX1 Else lngLeft = .CellLeft(i) + .OffsetX2
            Me.picSchedule.CurrentX = lngLeft
            Me.picSchedule.CurrentY = lngTop
            Me.picSchedule.Print i
        Next
    End With
End Sub

Private Sub DrawStatRow(penStat As StatEnum, plngTop As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngValue As Long
    Dim lngOldValue As Long
    Dim lngColor As Long
    Dim i As Long
    
    lngTop = plngTop + grid.OffsetY
    Me.picSchedule.CurrentX = 0
    Me.picSchedule.CurrentY = lngTop
    Me.picSchedule.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    Me.picSchedule.Print GetStatName(penStat)
    For i = 0 To 30
        lngValue = CalculateStat(penStat, i, , (i <> 0))
        If lngValue <> lngOldValue Then lngColor = cfg.GetColor(cgeWorkspace, cveText) Else lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        If lngValue < 10 Then lngLeft = grid.CellLeft(i) + grid.OffsetX1 Else lngLeft = grid.CellLeft(i) + grid.OffsetX2
        Me.picSchedule.CurrentX = lngLeft
        Me.picSchedule.CurrentY = lngTop
        Me.picSchedule.ForeColor = lngColor
        Me.picSchedule.Print lngValue
        lngOldValue = lngValue
    Next
    plngTop = plngTop + grid.RowHeight
End Sub

Private Sub DrawTomeRow(penStat As StatEnum, plngTop As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngValue As Long
    Dim lngOldValue As Long
    Dim lngColor As Long
    Dim i As Long
    
    If build.Tome(penStat) = 0 Then Exit Sub
    lngTop = plngTop + grid.OffsetY
    Me.picSchedule.CurrentX = 0
    Me.picSchedule.CurrentY = lngTop
    Me.picSchedule.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    Me.picSchedule.Print "Tome"
    For i = 0 To 30
        lngValue = CalculateTome(penStat, i, (i <> 0))
        If lngValue <> lngOldValue Then lngColor = cfg.GetColor(cgeWorkspace, cveText) Else lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        If lngValue < 10 Then lngLeft = grid.CellLeft(i) + grid.OffsetX1 Else lngLeft = grid.CellLeft(i) + grid.OffsetX2
        Me.picSchedule.CurrentX = lngLeft
        Me.picSchedule.CurrentY = lngTop
        Me.picSchedule.ForeColor = lngColor
        Me.picSchedule.Print lngValue
        lngOldValue = lngValue
    Next
    plngTop = plngTop + grid.RowHeight
End Sub

Private Sub DrawLevelupRow(penStat As StatEnum, plngTop As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngValue As Long
    Dim lngOldValue As Long
    Dim lngColor As Long
    Dim i As Long
    
    For i = 1 To 7
        If build.Levelups(i) = penStat Then Exit For
    Next
    If i > 7 Then Exit Sub
    lngTop = plngTop + grid.OffsetY
    Me.picSchedule.CurrentX = 0
    Me.picSchedule.CurrentY = lngTop
    Me.picSchedule.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    Me.picSchedule.Print "Levelups"
    For i = 0 To 30
        lngValue = CalculateLevelup(penStat, i)
        If lngValue <> lngOldValue Then lngColor = cfg.GetColor(cgeWorkspace, cveText) Else lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        If lngValue < 10 Then lngLeft = grid.CellLeft(i) + grid.OffsetX1 Else lngLeft = grid.CellLeft(i) + grid.OffsetX2
        Me.picSchedule.CurrentX = lngLeft
        Me.picSchedule.CurrentY = lngTop
        Me.picSchedule.ForeColor = lngColor
        Me.picSchedule.Print lngValue
        lngOldValue = lngValue
    Next
    plngTop = plngTop + grid.RowHeight
End Sub

Private Sub HideSchedule()
    EnableStats True
    Me.picSchedule.Visible = False
End Sub

Private Sub EnableStats(pblnEnabled As Boolean)
    Dim i As Long
    
    For i = 0 To 3
        Me.fraGroup(i).Enabled = pblnEnabled
    Next
End Sub
