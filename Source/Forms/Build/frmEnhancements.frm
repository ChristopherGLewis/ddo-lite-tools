VERSION 5.00
Begin VB.Form frmEnhancements 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enhancements"
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
   Icon            =   "frmEnhancements.frx":0000
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
      LeftLinks       =   "< Previous"
      RightLinks      =   "Destiny >"
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
      LeftLinks       =   "Trees;Enhancements;Leveling Guide"
      RightLinks      =   "Help"
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6828
      Index           =   2
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   12015
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   12012
      Begin VB.ComboBox cboGuideAbility 
         Height          =   312
         ItemData        =   "frmEnhancements.frx":000C
         Left            =   2340
         List            =   "frmEnhancements.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   180
         Width           =   3432
      End
      Begin CharacterBuilderLite.userDetails usrdetGuide 
         Height          =   2040
         Left            =   240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4452
         Width           =   5532
         _ExtentX        =   9763
         _ExtentY        =   3598
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6132
         Left            =   6000
         ScaleHeight     =   6135
         ScaleWidth      =   5970
         TabIndex        =   50
         TabStop         =   0   'False
         Tag             =   "ctl"
         Top             =   360
         Width           =   5976
         Begin VB.Timer tmrInterval 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   5340
            Top             =   360
         End
         Begin VB.Timer tmrScroll 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   4860
            Top             =   360
         End
         Begin VB.PictureBox picGuide 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2112
            Left            =   300
            ScaleHeight     =   2115
            ScaleWidth      =   3555
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   480
            Width           =   3552
         End
         Begin VB.VScrollBar scrollVertical 
            Enabled         =   0   'False
            Height          =   2568
            LargeChange     =   20
            Left            =   4260
            Max             =   20
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   360
            Value           =   1
            Width           =   252
         End
      End
      Begin VB.ListBox lstGuideAbility 
         Appearance      =   0  'Flat
         Height          =   3552
         IntegralHeight  =   0   'False
         ItemData        =   "frmEnhancements.frx":005B
         Left            =   2340
         List            =   "frmEnhancements.frx":005D
         TabIndex        =   38
         Top             =   540
         Width           =   3432
      End
      Begin VB.PictureBox picGuideTree 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3912
         Left            =   240
         ScaleHeight     =   3915
         ScaleWidth      =   2055
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   2052
         Begin VB.ListBox lstGuideTree 
            Appearance      =   0  'Flat
            Height          =   3552
            IntegralHeight  =   0   'False
            ItemData        =   "frmEnhancements.frx":005F
            Left            =   0
            List            =   "frmEnhancements.frx":0061
            TabIndex        =   36
            Top             =   300
            Width           =   1872
         End
         Begin CharacterBuilderLite.userCheckBox usrchkGuideTree 
            Height          =   252
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   1872
            _ExtentX        =   3307
            _ExtentY        =   450
            Value           =   0   'False
            Caption         =   "Show All"
         End
      End
      Begin VB.PictureBox picSelector 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3912
         Left            =   240
         ScaleHeight     =   3915
         ScaleWidth      =   2055
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   2052
         Begin VB.CheckBox chkBackToTrees 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Back to Trees"
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   3480
            Width           =   1632
         End
         Begin VB.ListBox lstGuideSub 
            Appearance      =   0  'Flat
            Height          =   3072
            IntegralHeight  =   0   'False
            ItemData        =   "frmEnhancements.frx":0063
            Left            =   0
            List            =   "frmEnhancements.frx":006A
            TabIndex        =   58
            Top             =   300
            Width           =   1872
         End
      End
      Begin VB.Label lblGuideAP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Available: 4 AP"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7800
         TabIndex        =   54
         Top             =   6540
         Width           =   1335
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Prog"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   6
         Left            =   11340
         TabIndex        =   49
         Top             =   120
         Width           =   588
      End
      Begin VB.Label lblGuideSpent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Spent: 80 AP"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   10080
         TabIndex        =   55
         Top             =   6540
         Width           =   1185
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tier"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   3
         Left            =   9120
         TabIndex        =   46
         Top             =   120
         Width           =   348
      End
      Begin VB.Label lblGuideDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Details"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   240
         TabIndex        =   39
         Top             =   4212
         Width           =   5532
      End
      Begin VB.Label lblGuideSpentInTree 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "12 AP spent in tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   2184
         TabIndex        =   42
         Top             =   6540
         Visible         =   0   'False
         Width           =   1716
      End
      Begin VB.Label lblGuideCost 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cost: 2 AP"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   300
         TabIndex        =   41
         Top             =   6540
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label lblGuideLevels 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Levels: 20"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4560
         TabIndex        =   43
         Top             =   6540
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AP"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   5
         Left            =   10740
         TabIndex        =   48
         Top             =   120
         Width           =   528
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   4
         Left            =   9540
         TabIndex        =   47
         Top             =   120
         Width           =   1044
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Ability"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   6660
         TabIndex        =   45
         Top             =   120
         Width           =   2316
      End
      Begin VB.Label lblGuide 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Level"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   6000
         TabIndex        =   44
         Top             =   120
         Width           =   588
      End
      Begin VB.Label lblGuideLevel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Current Level: 1"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5760
         TabIndex        =   53
         Top             =   6540
         Width           =   1455
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6492
      Index           =   0
      Left            =   600
      ScaleHeight     =   6495
      ScaleWidth      =   12015
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   12012
      Begin VB.Frame fraUniversalAP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2352
         Left            =   5800
         TabIndex        =   60
         Top             =   3780
         Width           =   5000
         Begin CharacterBuilderLite.userSpinner usrspnUniversalAP 
            Height          =   300
            Left            =   540
            TabIndex        =   61
            Top             =   480
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   529
            Min             =   0
            Value           =   0
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   -2147483631
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblUniversalAP 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Universal AP"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   1170
         End
         Begin VB.Shape shpUniversalAP 
            Height          =   2232
            Left            =   0
            Top             =   120
            Width           =   5000
         End
         Begin VB.Label lblUniversalAPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Universal AP can be spent at level 1"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   64
            Top             =   1020
            Width           =   3525
         End
         Begin VB.Label lblUniversalAPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Universal AP Tomes can add +?"
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   62
            Top             =   1320
            Width           =   3090
         End
      End
      Begin VB.Frame fraRacialAP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2352
         Left            =   600
         TabIndex        =   7
         Top             =   3780
         Width           =   5000
         Begin CharacterBuilderLite.userSpinner usrspnRacialAP 
            Height          =   300
            Left            =   540
            TabIndex        =   9
            Top             =   480
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   529
            Min             =   0
            Value           =   0
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   -2147483631
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblRacialAPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Racial AP Tomes can add another +?"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   360
            TabIndex        =   13
            Tag             =   "- Racial AP Tomes can add up to +"
            Top             =   1620
            Width           =   3420
         End
         Begin VB.Label lblRacialAPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Racial Past Lives offer up to +? Racial AP"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   360
            TabIndex        =   12
            Top             =   1320
            Width           =   3756
         End
         Begin VB.Label lblRacialAP 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Racial Past Life AP"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   1596
         End
         Begin VB.Label lblRacialAPreq 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Requires 30 racial past lives"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   1680
            TabIndex        =   10
            Tag             =   "Requires # racial past lives"
            Top             =   510
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.Label lblRacialAPhelp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "- Racial AP can be spent at level 1"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   360
            TabIndex        =   11
            Top             =   1020
            Width           =   3012
         End
         Begin VB.Shape shpRacialAP 
            Height          =   2232
            Left            =   0
            Top             =   120
            Width           =   5000
         End
      End
      Begin VB.Frame fraTreeSelection 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3492
         Left            =   480
         TabIndex        =   28
         Top             =   240
         Width           =   10092
         Begin CharacterBuilderLite.userList usrTree 
            Height          =   3012
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   5412
            _ExtentX        =   9551
            _ExtentY        =   5318
         End
         Begin VB.ListBox lstTree 
            Appearance      =   0  'Flat
            Height          =   2652
            IntegralHeight  =   0   'False
            ItemData        =   "frmEnhancements.frx":007B
            Left            =   5880
            List            =   "frmEnhancements.frx":009D
            TabIndex        =   6
            Top             =   360
            Width           =   3972
         End
         Begin VB.Label lblSpentAll 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0 / 80 AP"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   3360
            TabIndex        =   3
            Top             =   3120
            Width           =   4000
         End
         Begin VB.Label lblSource 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Class"
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   8520
            TabIndex        =   5
            Top             =   120
            Width           =   1332
         End
         Begin VB.Label lblTree 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tree"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   5880
            TabIndex        =   4
            Top             =   120
            Width           =   2352
         End
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6828
      Index           =   1
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   12015
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   12012
      Begin VB.ComboBox cboTree 
         Height          =   312
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   3492
      End
      Begin VB.ListBox lstAbility 
         Appearance      =   0  'Flat
         Height          =   3012
         IntegralHeight  =   0   'False
         ItemData        =   "frmEnhancements.frx":00EA
         Left            =   5520
         List            =   "frmEnhancements.frx":00EC
         TabIndex        =   24
         Top             =   1080
         Width           =   3492
      End
      Begin VB.ListBox lstSub 
         Appearance      =   0  'Flat
         Height          =   3012
         IntegralHeight  =   0   'False
         ItemData        =   "frmEnhancements.frx":00EE
         Left            =   9240
         List            =   "frmEnhancements.frx":00F0
         TabIndex        =   26
         Top             =   1080
         Width           =   2652
      End
      Begin CharacterBuilderLite.userList usrList 
         Height          =   6492
         Left            =   360
         TabIndex        =   15
         Top             =   0
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   11456
      End
      Begin CharacterBuilderLite.userDetails usrDetails 
         Height          =   2040
         Left            =   5520
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4452
         Width           =   6372
         _ExtentX        =   7011
         _ExtentY        =   3598
      End
      Begin CharacterBuilderLite.userCheckBox usrchkShowAll 
         Height          =   252
         Left            =   7200
         TabIndex        =   23
         Top             =   816
         Width           =   1812
         _ExtentX        =   3201
         _ExtentY        =   450
         Value           =   0   'False
         Caption         =   "Show All"
         CheckPosition   =   1
      End
      Begin VB.Label lblTier5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tier 5"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   9240
         TabIndex        =   21
         Top             =   406
         Visible         =   0   'False
         Width           =   504
      End
      Begin VB.Label lblTier5Label 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tier 5 Tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   9240
         TabIndex        =   20
         Top             =   132
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cost: 2 AP per rank"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   8040
         TabIndex        =   31
         Top             =   6540
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblRanks 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Ranks: 3"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   6960
         TabIndex        =   30
         Top             =   6540
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblProg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "12 AP spent in tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   10080
         TabIndex        =   32
         Top             =   6540
         Visible         =   0   'False
         Width           =   1704
      End
      Begin VB.Label lblTree 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tree"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   5520
         TabIndex        =   18
         Top             =   132
         Width           =   396
      End
      Begin VB.Label lblSpent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Spent in Tree: 24 AP"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4560
         TabIndex        =   17
         Top             =   6540
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0 / 80 AP"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   480
         TabIndex        =   16
         Top             =   6540
         Width           =   852
      End
      Begin VB.Label lblAbilities 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Abilities"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   5520
         TabIndex        =   22
         Top             =   840
         Width           =   648
      End
      Begin VB.Label lblSelectors 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Selectors"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   9240
         TabIndex        =   25
         Top             =   840
         Width           =   828
      End
      Begin VB.Label lblDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Details"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   5520
         TabIndex        =   27
         Top             =   4212
         Width           =   6372
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Trees"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuTrees 
         Caption         =   "Delete Tree"
         Index           =   0
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "Delete All Trees"
         Index           =   1
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "Reset Tree"
         Index           =   3
      End
      Begin VB.Menu mnuTrees 
         Caption         =   "Reset All Trees"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Enhancements"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuEnhancements 
         Caption         =   "Clear this ability"
         Index           =   0
      End
      Begin VB.Menu mnuEnhancements 
         Caption         =   "Reset tree"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Guide"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuGuide 
         Caption         =   "Move Row(s)"
         Index           =   0
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Delete Row(s)"
         Index           =   2
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Delete to End"
         Index           =   3
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Delete All"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmEnhancements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Enum MouseShiftEnum
    mseNormal
    mseCtrl
    mseShift
End Enum

Private Type ColumnType
    Header As String
    Left As Long
    Width As Long
    Right As Long
    Align As AlignmentConstants
End Type

Private Enum GuideFilterEnum
    gfeSelectedAvailable
    gfeAllSelected
    gfeAllAvailable
    gfeAllEnhancements
    gfeUnknown
End Enum

Private Col(1 To 6) As ColumnType
Private mlngGuideTree As Long
Private mlngGuideLevel As Long
Private mlngGuideAP As Long
Private mlngRow As Long
Private mlngCol As Long
Private mlngLastRowSelected As Long
Private mblnMoveRows As Boolean
Private mlngHeight As Long
Private mlngOffsetX As Long
Private mlngOffsetY As Long
Private mlngInterval As Long
Private mlngDirection As Long

' Enhancements
Private mlngTree As Long
Private mlngBuildTree As Long
Private mlngMaxTier As Long
' Drag & Drop
Private mblnMouse As Boolean
Private mlngSourceIndex As Long
Private menDragState As DragEnum
Private mblnDragComplete As Boolean
Private msngDownX As Single
Private msngDownY As Single
' General
Private mlngTab As Long
Private mblnNoFocus As Boolean
Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.Configure Me
    mlngTab = 0
    LoadData
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Activate()
    ActivateForm oeEnhancements
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngTab = 2 Then ActiveCell -1, -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    UnloadForm Me, mblnOverride
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case mlngTab
        Case 0
            If IsOver(Me.usrspnRacialAP.hwnd, Xpos, Ypos) Then Me.usrspnRacialAP.WheelScroll lngValue
            If IsOver(Me.usrspnUniversalAP.hwnd, Xpos, Ypos) Then Me.usrspnUniversalAP.WheelScroll lngValue
        Case 1
            Select Case True
                Case IsOver(Me.usrList.hwnd, Xpos, Ypos): Me.usrList.Scroll lngValue
                Case IsOver(Me.usrDetails.hwnd, Xpos, Ypos): Me.usrDetails.Scroll lngValue
            End Select
        Case 2
            Select Case True
                Case IsOver(Me.usrdetGuide.hwnd, Xpos, Ypos): Me.usrdetGuide.Scroll lngValue
                Case IsOver(Me.picContainer.hwnd, Xpos, Ypos): GuideWheelScroll lngValue
            End Select
    End Select
End Sub

Public Sub Cascade()
    Dim i As Long
    
    NoFocus True
    xp.LockWindow Me.hwnd
    LoadData
    Select Case mlngTab
        Case 1 ' Enhancements
            If mlngTree <> 0 Then
                mlngBuildTree = 0
                For i = 1 To build.Trees
                    If build.Tree(i).TreeName = db.Tree(mlngTree).TreeName Then
                        mlngBuildTree = i
                        Exit For
                    End If
                Next
            End If
    End Select
    ShowTab
    xp.UnlockWindow
    NoFocus False
End Sub

Private Sub NoFocus(pblnNoFocus As Boolean)
    mblnNoFocus = pblnNoFocus
    Me.usrTree.NoFocus = pblnNoFocus
    Me.usrList.NoFocus = pblnNoFocus
    Me.usrDetails.NoFocus = pblnNoFocus
End Sub

Public Sub RefreshColors()
    If mlngTab = 2 Then DrawGrid
End Sub

Public Property Get CurrentTab() As Variant
    CurrentTab = mlngTab
End Property


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Trees": ChangeTab 0
        Case "Enhancements": ChangeTab 1
        Case "Leveling Guide": ChangeTab 2
        Case "Export": ExportGuide "csv"
        Case "Help": If mlngTab = 2 Then ShowHelp "Leveling_Guide" Else ShowHelp "Enhancements"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "< Feats": If Not OpenForm("frmFeats") Then Exit Sub
        Case "< Spells": If Not OpenForm("frmSpells") Then Exit Sub
        Case "Destiny >": If Not OpenForm("frmDestiny") Then Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub

Private Sub ChangeTab(plngTab As Long)
    If mlngTab = plngTab Then Exit Sub
    mlngTab = plngTab
    ShowTab
    SaveBackup
    GenerateOutput oeEnhancements
End Sub

Private Sub ShowTab()
    Dim i As Long
    
    xp.LockWindow Me.hwnd
    If mlngTab = 2 Then Me.usrHeader.RightLinks = "Export;Help" Else Me.usrHeader.RightLinks = "Help"
    Select Case mlngTab
        Case 0 ' Trees
            ShowTrees False, False
            ShowAvailableTrees
        Case 1 ' Enhancements
            ShowTier5
            Me.usrList.GotoTop
            If mlngBuildTree = 0 Then
                Me.cboTree.ListIndex = 0
                TreeClick
            Else
                ShowAbilities
                ShowAvailable False
            End If
        Case 2 ' Leveling Guide
            InitGuide
            ClearSelection
    End Select
    For i = 0 To 2
        Me.picTab(i).Visible = (mlngTab = i)
    Next
    xp.UnlockWindow
End Sub


' ************* INITIALIZE *************


Private Sub LoadData()
    Dim lngTabs() As Long
    Dim lngMax As Long
    Dim i As Long
    
    mlngSourceIndex = 0
    ' Navigation
    If build.CanCastSpell(1) <> 0 Then Me.usrFooter.LeftLinks = "< Spells" Else Me.usrFooter.LeftLinks = "< Feats"
    If build.MaxLevels > 19 Then Me.usrFooter.RightLinks = "Destiny >" Else Me.usrFooter.RightLinks = vbNullString
    ' Enhancements
    With Me.usrList
        .DefineDimensions 1, 3, 2
        .DefineColumn 1, vbCenter, "Tier", "Tier"
        .DefineColumn 2, vbCenter, "Ability"
        .DefineColumn 3, vbCenter, "AP", " 20"
        .Refresh
    End With
    Me.usrDetails.Clear
    PopulateCombo
    mblnOverride = True
    ' Trees
    With Me.usrTree
        .DefineDimensions 1, 4, 2
        .DefineColumn 1, vbLeftJustify, "Class", "Favored Soul"
        .DefineColumn 2, vbCenter, "Tree"
        .DefineColumn 3, vbCenter, "Tier", "Tier"
        .DefineColumn 4, vbCenter, "AP", "AP"
        .Refresh
    End With
    ReDim lngTabs(0)
    lngTabs(0) = 86
    lngTabs(0) = 97
    ListboxTabStops Me.lstTree, lngTabs
    ShowTrees False, False
    mblnOverride = True
    ShowAvailableTrees
    ' Racial AP
    For i = 1 To reRaces - 1
        If db.Race(i).Type <> rteIconic And db.Race(i).SubRace = reAny Then lngMax = lngMax + 1
    Next
    Me.lblRacialAPhelp(1).Caption = "- Racial Past Lives offer up to +" & lngMax & " Racial AP"
    Me.usrspnRacialAP.Max = lngMax + tomes.RacialAPMax
    Me.usrspnRacialAP.Value = build.RacialAP
    Me.lblRacialAPhelp(2).Caption = "- Racial AP Tomes can add another +" & tomes.RacialAPMax & " Racial AP"
    ' Universal AP
    Me.lblUniversalAPhelp(1).Caption = Replace(Me.lblUniversalAPhelp(1).Caption, "?", tomes.UniversalAPMax)
    Me.usrspnUniversalAP.Max = tomes.UniversalAPMax
    Me.usrspnUniversalAP.Value = build.UniversalAP
    
    
    Me.lblTier5Label.Visible = (build.MaxLevels > 11)
    Me.lblTier5.Visible = (build.MaxLevels > 11)
    If build.MaxLevels < 12 And Len(build.Tier5) Then
        ShaveTree build.Tier5
        build.Tier5 = vbNullString
    End If
    ' Leveling Guide
    Me.cboGuideAbility.ListIndex = 0
    mblnOverride = False
End Sub

Private Sub PopulateCombo()
    Dim strTree As String
    Dim i As Long
    
    mblnOverride = True
    strTree = Me.cboTree.Text
    ComboClear Me.cboTree
    For i = 1 To build.Trees
        ComboAddItem Me.cboTree, build.Tree(i).TreeName, i
    Next
    mblnOverride = False
    If Len(strTree) Then ComboSetText Me.cboTree, strTree
    If Me.cboTree.ListIndex = -1 Then
        If Len(build.Tier5) Then
            ComboSetText Me.cboTree, build.Tier5
        Else
            ComboSetText Me.cboTree, db.Race(build.Race).RaceName
        End If
        TreeClick
    End If
End Sub


' ************* TREES *************
Private Sub usrspnRacialAP_Change()
    Dim lngLives As Long
    
    If mblnOverride Then Exit Sub
    build.RacialAP = Me.usrspnRacialAP.Value
    If build.RacialAP > tomes.RacialAPMax Then lngLives = build.RacialAP - tomes.RacialAPMax
    With Me.lblRacialAPreq
        If lngLives = 0 Then
            .Caption = "Can be reached with tomes"
        Else
            .Caption = Replace(.Tag, "#", lngLives * 3)
        End If
        Me.lblRacialAPreq.Visible = (build.RacialAP <> 0)
    End With
    InitGuideEnhancements
    ShowSpentAll Me.lblSpentAll
    SetDirty
End Sub

Private Sub usrspnUniversalAP_Change()
    Dim lngLives As Long
    
    If mblnOverride Then Exit Sub
    build.UniversalAP = Me.usrspnUniversalAP.Value
    If build.UniversalAP > tomes.UniversalAPMax Then lngLives = build.UniversalAP - tomes.UniversalAPMax
    InitGuideEnhancements
    ShowSpentAll Me.lblSpentAll
    SetDirty
End Sub


Private Sub usrTree_SlotClick(Index As Integer, Button As Integer)
    Dim strTreeName As String
    Dim lngTree As Long
    Dim i As Long
    
    Me.lstTree.ListIndex = -1
    Select Case Button
        Case vbLeftButton
            If Me.usrTree.Selected = Index Then Me.usrTree.Selected = 0 Else Me.usrTree.Selected = Index
        Case vbRightButton
            Me.usrTree.Selected = Index
            Me.usrTree.Active = Index
            lngTree = SeekTree(build.Tree(Index).TreeName, peEnhancement)
            Me.mnuTrees(0).Caption = "Delete " & db.Tree(lngTree).Abbreviation
            Me.mnuTrees(0).Enabled = (build.Tree(Index).TreeType <> tseRace)
            Me.mnuTrees(3).Caption = "Reset " & db.Tree(lngTree).Abbreviation
            PopupMenu Me.mnuMain(0)
    End Select
End Sub

Private Sub mnuTrees_Click(Index As Integer)
    Dim lngBuildTree As Long
    Dim i As Long
    
    Select Case Index
        Case 0 ' Delete tree
            lngBuildTree = Me.usrTree.Selected
            If lngBuildTree = 0 Then Exit Sub
            If build.Tree(lngBuildTree).TreeType = tseRace Then Exit Sub
            For i = lngBuildTree To build.Trees - 1
                build.Tree(i) = build.Tree(i + 1)
            Next
            build.Trees = build.Trees - 1
            ReDim Preserve build.Tree(1 To build.Trees)
            PopulateCombo
        Case 1 ' Delete all trees
            If Not Ask("Delete all trees?") Then Exit Sub
            ReDim build.Tree(1 To 1)
            build.Trees = 1
            With build.Tree(1)
                .TreeName = db.Race(build.Race).RaceName
                .TreeType = tseRace
            End With
            PopulateCombo
        Case 3 ' Reset tree
            lngBuildTree = Me.usrTree.Selected
            If lngBuildTree = 0 Then Exit Sub
            With build.Tree(lngBuildTree)
                Erase .Ability
                .Abilities = 0
            End With
        Case 4 ' Reset all trees
            If Not Ask("Reset all trees?") Then Exit Sub
            For lngBuildTree = 1 To build.Trees
                With build.Tree(lngBuildTree)
                    Erase .Ability
                    .Abilities = 0
                End With
            Next
    End Select
    ShowTrees False, False
    ShowAvailableTrees
    SetDirty
    InitLevelingGuide
End Sub

Private Sub usrTree_SlotDblClick(Index As Integer)
    Dim strTreeName As String
    Dim i As Long
    
    Me.lstTree.ListIndex = -1
    If build.Tree(Index).TreeType = tseRace Then Exit Sub
    With build
        If .Tier5 = .Tree(Index).TreeName Then .Tier5 = vbNullString
        For i = Index To .Trees - 1
            .Tree(i) = .Tree(i + 1)
        Next
        .Trees = .Trees - 1
        If .Trees = 0 Then Erase .Tree Else ReDim Preserve .Tree(1 To .Trees)
    End With
    ShowTrees False, False
    ShowAvailableTrees
    PopulateCombo
    SetDirty
    InitLevelingGuide
End Sub

Private Sub ShowTrees(pblnDrop As Boolean, pblnAdd As Boolean)
    Dim lngTree As Long
    Dim lngMax As Long
    Dim lngRaceTree As Long
    Dim enDropState As DropStateEnum
    Dim i As Long
    
    mblnOverride = True
    If pblnAdd And build.Trees < 7 Then lngMax = build.Trees + 1 Else lngMax = build.Trees
    Me.usrTree.Rows = lngMax
    For i = 1 To lngMax
        If i <= build.Trees Then
            With build.Tree(i)
                Me.usrTree.SetSlot i, .TreeName
                lngTree = SeekTree(.TreeName, peEnhancement)
                Me.usrTree.SetItemData i, lngTree
                If .TreeType = tseRace Then
                    lngRaceTree = i
                    Me.usrTree.SetText i, 1, vbNullString
                    Me.usrTree.SetText i, 3, "4"
                Else
                    Me.usrTree.SetText i, 1, GetClassName(.Source)
                    Me.usrTree.SetText i, 3, GetMaxTier(i)
                End If
                Me.usrTree.SetText i, 4, QuickSpentInTree(i)
            End With
        End If
        If pblnDrop And i <> lngRaceTree Then enDropState = dsCanDrop Else enDropState = dsDefault
        Me.usrTree.SetDropState i, enDropState
    Next
    ShowSpentAll Me.lblSpentAll
    mblnOverride = False
End Sub

'This shows the spend line on the Enhancement tab
'TODO move to a bas file so it can be shared btwn Enh and main
Private Sub ShowSpentAll(plbl As Label)
    Dim lngSpentBase As Long
    Dim lngSpentRPLBonus As Long
    Dim lngSpentUniBonus As Long
    Dim lngMaxBase As Long
    Dim lngMaxRPLBonus As Long
    Dim lngMaxUniBonus As Long
    Dim strDisplay As String
    
    'Display should be Spent/Max AP.  Long form is Spent+r+u/Max+r+u AP

    'retrieve each of the spent/maxes from the build tree object
    'This should be a func that returns a class object
    GetPointsSpentAndMax lngSpentBase, lngSpentRPLBonus, lngSpentUniBonus, lngMaxBase, lngMaxRPLBonus, lngMaxUniBonus
    
    'Display
    strDisplay = strDisplay & "Spent: " & lngSpentBase & "+" & lngSpentRPLBonus & "r +" & lngSpentUniBonus & "u / Max: "
    strDisplay = strDisplay & lngMaxBase & "+" & lngMaxRPLBonus & "r +" & lngMaxUniBonus & "u AP"
    
    plbl.Caption = strDisplay
    
    'Flag this as an error if spent is > max
    If lngSpentBase > lngMaxBase Then plbl.ForeColor = cfg.GetColor(cgeWorkspace, cveTextError) Else plbl.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    plbl.Visible = True
End Sub

Private Function PointsSpent(plngBuildTree As Long) As Long
    Dim lngTree As Long
    Dim lngTotal As Long
    Dim lngSpent() As Long
    
    lngTree = SeekTree(build.Tree(plngBuildTree).TreeName, peEnhancement)
    GetSpentInTree db.Tree(lngTree), build.Tree(plngBuildTree), lngSpent, lngTotal
    PointsSpent = lngTotal
End Function

Private Sub ShowAvailableTrees()
    Dim typClassSplit() As ClassSplitType
    Dim lngClass As Long
    Dim enClass As ClassEnum
    Dim lngTree As Long
    Dim i As Long
    
    ListboxClear Me.lstTree
    ' Add class trees
    For lngClass = 0 To GetClassSplit(typClassSplit) - 1
        enClass = typClassSplit(lngClass).ClassID
        With db.Class(enClass)
            For i = 1 To .Trees
                lngTree = SeekTree(.Tree(i), peEnhancement)
                If lngTree Then AddTree lngTree, .ClassName
            Next
        End With
    Next
    ' Add racial class tree(s)
    With db.Race(build.Race)
        For i = 1 To .Trees
            lngTree = SeekTree(.Tree(i), peEnhancement)
            If lngTree Then AddTree lngTree, .RaceName
        Next
    End With
    ' Add global tree(s)
    For i = 1 To db.Trees
        If db.Tree(i).TreeType = tseGlobal Then AddTree i, vbNullString
    Next
'    PopulateCombo
End Sub

Private Sub AddTree(plngTree As Long, pstrSource As String)
    Dim blnFound As Boolean
    Dim lngTree As Long
    Dim i As Long
    
    ' Already selected?
    For i = 1 To build.Trees
        If build.Tree(i).TreeName = db.Tree(plngTree).TreeName Then Exit Sub
    Next
    ' Already added to available list?
    For i = 0 To Me.lstTree.ListCount - 1
        If Me.lstTree.ItemData(i) = plngTree Then
            Exit Sub
        End If
    Next
    ' Locked out?
    If Len(db.Tree(plngTree).Lockout) Then
        For i = 1 To build.Trees
            If build.Tree(i).TreeName = db.Tree(plngTree).Lockout Then Exit Sub
        Next
    End If
    ' Add this tree
    ListboxAddItem Me.lstTree, db.Tree(plngTree).TreeName & vbTab & pstrSource, plngTree
End Sub

Private Sub lstTree_Click()
    If mblnMouse Then mblnMouse = False
End Sub

Private Sub lstTree_DblClick()
    If Me.lstTree.ListIndex = -1 Or build.Trees > 6 Then Exit Sub
    AddBuildTree build.Trees + 1
    ShowTrees False, False
    ShowAvailableTrees
    PopulateCombo
    Me.usrTree.Active = build.Trees
    SetDirty
End Sub

Private Sub lstTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.usrTree.Selected = 0
    mblnMouse = True
    menDragState = dragMouseDown
End Sub

Private Sub lstTree_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            ShowTrees True, True
            Me.lstTree.OLEDropMode = vbOLEDropManual
            Me.lstTree.OLEDrag
        End If
    End If
End Sub

Private Sub lstTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub lstTree_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "Add"
End Sub

Private Sub lstTree_OLECompleteDrag(Effect As Long)
    ShowTrees False, False
    Me.lstTree.OLEDropMode = vbOLEDropNone
End Sub

Private Sub usrTree_OLEDragDrop(Index As Integer, Data As DataObject)
    Dim strData As String
    
    If Not Data.GetFormat(vbCFText) Then Exit Sub
    strData = Data.GetData(vbCFText)
    If strData = "Add" Then AddBuildTree Index Else SwapBuildTrees Index, strData
    ShowTrees False, False
    ShowAvailableTrees
    PopulateCombo
    SetDirty
End Sub

Private Sub usrTree_OLECompleteDrag(Index As Integer, Effect As Long)
    ShowTrees False, False
End Sub

Private Sub usrTree_RequestDrag(Index As Integer, Allow As Boolean)
    Dim enDropState As DropStateEnum
    Dim i As Long
    
    For i = 1 To build.Trees
        If i = Index Then enDropState = dsDefault Else enDropState = dsCanDrop
        Me.usrTree.SetDropState i, enDropState
    Next
    Allow = True
End Sub

Private Sub AddBuildTree(ByVal plngIndex As Long)
    Dim strText() As String
    Dim enClass As ClassEnum
    Dim lngTree As Long
    
    If Me.lstTree.ListIndex = -1 Then Exit Sub
    strText = Split(Me.lstTree.Text, vbTab)
    enClass = GetClassID(strText(1))
    With build
        If plngIndex > .Trees Then
            .Trees = plngIndex
            ReDim Preserve .Tree(1 To plngIndex)
        End If
        With .Tree(plngIndex)
            .TreeName = strText(0)
            .Abilities = 0
            Erase .Ability
            .Source = enClass
            If .Source = 0 Then
                .ClassLevels = build.MaxLevels
                lngTree = SeekTree(.TreeName, peEnhancement)
                If lngTree <> 0 Then .TreeType = db.Tree(lngTree).TreeType
            Else
                .ClassLevels = ClassLevels(enClass)
                .TreeType = tseClass
            End If
        End With
    End With
    InitLevelingGuide
End Sub

Private Function ClassLevels(penClass As ClassEnum) As Long
    Dim lngClassLevels As Long
    Dim i As Long
    
    If penClass = ceAny Then
        lngClassLevels = HeroicLevels()
    Else
        For i = 1 To HeroicLevels()
            If build.Class(i) = penClass Then lngClassLevels = lngClassLevels + 1
        Next
    End If
    ClassLevels = lngClassLevels
End Function

Private Sub SwapBuildTrees(ByVal plngTree1 As Long, ByVal plngTree2 As Long)
    Dim typSwap As BuildTreeType
    
    typSwap = build.Tree(plngTree1)
    build.Tree(plngTree1) = build.Tree(plngTree2)
    build.Tree(plngTree2) = typSwap
    InitLevelingGuide
End Sub


' ************* DISPLAY *************


Private Sub ShowAbilities()
    Dim strCaption As String
    Dim lngCost As Long
    Dim lngRanks As Long
    Dim lngMaxRanks As Long
    Dim lngTotal As Long
    Dim lngSpent() As Long
    Dim blnBlank As Boolean
    Dim lngForceVisible As Long
    Dim i As Long

    If mlngBuildTree = 0 Or mlngBuildTree > build.Trees Then
        blnBlank = True
    ElseIf build.Tree(mlngBuildTree).Abilities = 0 Then
        blnBlank = True
    End If
    If blnBlank Then
        Me.usrList.Rows = 1
        SetSlot 1, vbNullString, vbNullString, vbNullString, 0, 0
        Me.usrList.SetError 1, False
        Me.usrList.SetDropState 1, dsDefault
    Else
        Me.usrList.Rows = build.Tree(mlngBuildTree).Abilities
        For i = 1 To build.Tree(mlngBuildTree).Abilities
            If build.Tree(mlngBuildTree).Ability(i).Ability = 0 Then
                Me.usrList.ForceVisible i
                Exit For
            End If
        Next
    End If
    GetSpentInTree db.Tree(mlngTree), build.Tree(mlngBuildTree), lngSpent, lngTotal
    For i = 1 To build.Tree(mlngBuildTree).Abilities
        If build.Tree(mlngBuildTree).Ability(i).Ability = 0 Then
            SetSlot i, vbNullString, vbNullString, vbNullString, 0, 0
        Else
            GetSlotInfo db.Tree(mlngTree), build.Tree(mlngBuildTree).Ability(i), strCaption, lngCost, lngRanks, lngMaxRanks
            SetSlot i, build.Tree(mlngBuildTree).Ability(i).Tier, lngCost, strCaption, lngRanks, lngMaxRanks
            Me.usrList.SetError i, CheckErrors(db.Tree(mlngTree), build.Tree(mlngBuildTree).Ability(i), lngSpent)
        End If
    Next
    With Me.lblSpent
        .Caption = "Spent in Tree: " & lngTotal
        .Visible = True
    End With
    ShowSpentAll Me.lblTotal
End Sub

' Light up valid drop locations during drag operations
Private Sub ShowDropSlots()
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim enDropState As DropStateEnum
    Dim lngSpent() As Long
    Dim typCheck As BuildAbilityType
    Dim i As Long

    If mlngBuildTree = 0 Then
        GetUserChoices lngTier, lngAbility, lngSelector
        If lngTier = 0 And lngAbility = 1 Then
            Me.usrList.SetDropState 1, dsCanDrop
        Else
            Me.usrList.SetDropState 1, dsDefault
        End If
    Else
        GetSpentInTree db.Tree(mlngTree), build.Tree(mlngBuildTree), lngSpent, 0
        With build.Tree(mlngBuildTree)
            For i = 1 To .Abilities
                If .Ability(i).Ability <> 0 Then
                    enDropState = dsDefault
                Else
                    GetUserChoices lngTier, lngAbility, lngSelector
                    With typCheck
                        .Tier = lngTier
                        .Ability = lngAbility
                        .Selector = lngSelector
                        .Rank = 1
                    End With
                    If CheckErrors(db.Tree(mlngTree), typCheck, lngSpent) Then enDropState = dsCanDropError Else enDropState = dsCanDrop
                End If
                Me.usrList.SetDropState i, enDropState
                Me.usrList.ForceActive i
            Next
        End With
    End If
End Sub

' Returns TRUE if errors found
Private Function CheckErrors(ptypTree As TreeType, ptypAbility As BuildAbilityType, plngSpent() As Long) As Boolean
    CheckErrors = CheckAbilityErrors(ptypTree, build.Tree(mlngBuildTree), ptypAbility, plngSpent)
End Function

Private Sub GetSlotInfo(ptypTree As TreeType, ptypAbility As BuildAbilityType, pstrCaption As String, plngCost As Long, plngRanks, plngMaxRanks)
    With ptypAbility
        plngRanks = .Rank
        With ptypTree.Tier(.Tier).Ability(.Ability)
            If ptypAbility.Selector = 0 Then
                pstrCaption = .Abbreviation
                plngCost = .Cost
            Else
                pstrCaption = .Selector(ptypAbility.Selector).SelectorName
                If Not .SelectorOnly Then pstrCaption = .Abbreviation & ": " & pstrCaption
                plngCost = .Selector(ptypAbility.Selector).Cost
            End If
            If plngRanks <> 0 Then plngCost = plngCost * plngRanks
            plngMaxRanks = .Ranks
        End With
    End With
End Sub

Private Sub SetSlot(plngSlot As Long, ByVal pstrTier As String, ByVal pstrCost As String, pstrCaption As String, plngRanks As Long, plngMaxRanks As Long)
    With Me.usrList
        .SetText plngSlot, 1, pstrTier
        .SetText plngSlot, 3, pstrCost
        .SetSlot plngSlot, pstrCaption, plngRanks, plngMaxRanks
    End With
End Sub


' ************* GENERAL *************


Private Function GetUserChoices(plngTier As Long, plngAbility As Long, plngSelector As Long) As Boolean
    Dim lngItemData As Long

    plngTier = 0
    plngAbility = 0
    If plngSelector <> -1 Then plngSelector = 0
    If Me.lstAbility.ListIndex = -1 Then Exit Function
    lngItemData = Me.lstAbility.ItemData(Me.lstAbility.ListIndex)
    SplitAbilityID lngItemData, plngTier, plngAbility, 0
    If plngSelector <> -1 And db.Tree(mlngTree).Tier(plngTier).Ability(plngAbility).SelectorStyle <> sseNone Then
        If Me.lstSub.ListIndex = -1 Then Exit Function
        plngSelector = Me.lstSub.ItemData(Me.lstSub.ListIndex)
    End If
    GetUserChoices = True
End Function

Private Sub StartDrag()
    If mlngBuildTree <> 0 Then
        If Not AddAbility(True) Then Exit Sub
    End If
    ShowDropSlots
    If Me.lstSub.ListIndex = -1 Then ListboxDrag Me.lstAbility Else ListboxDrag Me.lstSub
End Sub

Private Sub ListboxDrag(plst As ListBox)
    plst.OLEDropMode = vbOLEDropManual
    plst.OLEDrag
End Sub

Private Function AddAbility(Optional pblnBlank As Boolean = False) As Boolean
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long
    Dim lngInsert As Long
    Dim typBlank As BuildAbilityType
    Dim i As Long

    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Function
    lngInsert = GetInsertionPoint(build.Tree(mlngBuildTree), lngTier, lngAbility)
    If lngAbility Then lngRanks = db.Tree(mlngTree).Tier(lngTier).Ability(lngAbility).Ranks
    With build.Tree(mlngBuildTree)
        .Abilities = .Abilities + 1
        ReDim Preserve .Ability(1 To .Abilities)
        For i = .Abilities To lngInsert + 1 Step -1
            .Ability(i) = .Ability(i - 1)
        Next
        If pblnBlank Then
            .Ability(lngInsert) = typBlank
            .Ability(lngInsert).Tier = 0
        Else
            With .Ability(lngInsert)
                .Tier = lngTier
                .Ability = lngAbility
                .Selector = lngSelector
                .Rank = lngRanks
            End With
        End If
    End With
    If Not pblnBlank And lngTier = 5 Then SetTier5 build.Tree(mlngBuildTree).TreeName
    ShowAbilities
    Me.usrList.Selected = 0
    Me.usrList.Active = lngInsert
    Me.usrList.ForceVisible lngInsert
    AddAbility = True
End Function

Private Sub DropAbility(Index As Integer)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRanks As Long

    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Sub
    lngRanks = db.Tree(mlngTree).Tier(lngTier).Ability(lngAbility).Ranks
    If mlngBuildTree = 0 Then
        With build
            .Trees = .Trees + 1
            ReDim Preserve .Tree(.Trees)
            mlngBuildTree = .Trees
            With .Tree(.Trees)
                .Abilities = 1
                ReDim .Ability(1 To 1)
            End With
        End With
    End If
    With build.Tree(mlngBuildTree).Ability(Index)
        .Tier = lngTier
        .Ability = lngAbility
        .Selector = lngSelector
        .Rank = lngRanks
    End With
End Sub


' ************* TIER 5 *************


Private Sub SetTier5(pstrTier5 As String)
    If Len(build.Tier5) = 0 Then
        build.Tier5 = pstrTier5
    ElseIf build.Tier5 <> pstrTier5 Then
        ShaveTree build.Tier5
        build.Tier5 = pstrTier5
    End If
    ShowTier5
End Sub

Private Sub ShowTier5()
    Dim blnVisible As Boolean
    
    blnVisible = (Len(build.Tier5) <> 0)
    Me.lblTier5Label.Visible = blnVisible
    Me.lblTier5.Caption = build.Tier5
    Me.lblTier5.Visible = blnVisible
End Sub

Private Function ConfirmTier5Change() As Boolean
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    ConfirmTier5Change = True
    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Function
    If lngTier = 5 And Len(build.Tier5) <> 0 And build.Tier5 <> build.Tree(mlngBuildTree).TreeName Then
        If Ask("Make " & build.Tree(mlngBuildTree).TreeName & " your Tier 5 tree?" & vbNewLine & vbNewLine & "(This will clear Tier 5 enhancements from the " & build.Tier5 & " tree.)") Then
            SetTier5 build.Tree(mlngBuildTree).TreeName
        Else
            ConfirmTier5Change = False
        End If
    End If
End Function


' ************* SLOTS *************


Private Sub usrList_SlotClick(Index As Integer, Button As Integer)
    Dim typBlank As TwistType
    
    If mlngBuildTree = 0 Then Exit Sub
    With Me.lstSub
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstAbility
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    If Not mblnNoFocus Then Me.usrList.SetFocus
    If Me.usrList.Rows = 0 Then
        Exit Sub
    ElseIf Len(Me.usrList.GetCaption(1)) = 0 Then
        Exit Sub
    ElseIf Me.usrList.Selected = Index Then
        NoSelection
        If Button = vbRightButton Then ClearSlot Index
    Else
        Select Case Button
            Case vbLeftButton
                Me.usrList.Selected = Index
                With build.Tree(mlngBuildTree).Ability(Index)
                    ShowDetails .Tier, .Ability, .Selector, Index
                End With
            Case vbRightButton
                If build.Tree(mlngBuildTree).Abilities = 0 Then Exit Sub
                Me.usrList.Selected = Index
                Me.usrList.Active = Index
                With build.Tree(mlngBuildTree).Ability(Index)
                    ShowDetails .Tier, .Ability, .Selector, Index
                End With
                PopupMenu Me.mnuMain(1)
        End Select
    End If
End Sub

Private Sub mnuEnhancements_Click(Index As Integer)
    Dim intSlot As Integer
    
    Select Case Index
        Case 0
            intSlot = Me.usrList.Selected
            ClearSlot intSlot
        Case 1
            If Not Ask("Reset " & Me.cboTree.Text & "?") Then Exit Sub
            build.Tree(mlngBuildTree).Abilities = 0
            Erase build.Tree(mlngBuildTree).Ability
            If build.Tier5 = build.Tree(mlngBuildTree).TreeName Then SetTier5 vbNullString
            ShowAbilities
            ShowAvailable False
            NoSelection
            SetDirty
    End Select
End Sub

Private Sub usrList_SlotDblClick(Index As Integer)
    ClearSlot Index
End Sub

Private Sub ClearSlot(Index As Integer)
    If mlngBuildTree = 0 Then Exit Sub
    If build.Tree(mlngBuildTree).Abilities >= Index Then
        build.Tree(mlngBuildTree).Ability(Index).Ability = 0
        RemoveBlanks build.Tree(mlngBuildTree)
        With build.Tree(mlngBuildTree)
            If .Abilities = 0 Then
                SetTier5 vbNullString
            ElseIf .Ability(.Abilities).Tier < 5 And build.Tier5 = .TreeName Then
                SetTier5 vbNullString
            End If
        End With
        ShowAbilities
        ShowAvailable True
        NoSelection
        SetDirty
    End If
End Sub

Private Sub usrList_RequestDrag(Index As Integer, Allow As Boolean)
    mlngSourceIndex = Index
    Allow = True
    Me.lstAbility.OLEDropMode = vbOLEDropManual
End Sub

Private Sub usrList_OLEDragDrop(Index As Integer, Data As DataObject)
    If ConfirmTier5Change() Then
        mblnDragComplete = True
        DropAbility Index
        Me.usrList.SetDropState Index, dsDefault
        ShowAbilities
        Me.usrList.Selected = 0
        Me.usrList.Active = Index
        If Not mblnNoFocus Then Me.usrList.SetFocus
        ShowAvailable True
        SetDirty
    End If
End Sub

Private Sub usrList_RankChange(Index As Integer, Ranks As Long)
    build.Tree(mlngBuildTree).Ability(Index).Rank = Ranks
    ShowAbilities
    ShowAvailable True
    With build.Tree(mlngBuildTree).Ability(Index)
        ShowDetails .Tier, .Ability, .Selector, Index
    End With
    SetDirty
End Sub


' ************* ABILITIES *************


Private Sub lstAbility_Click()
    If mblnMouse Then mblnMouse = False Else ListAbilityClick
End Sub

Private Sub lstAbility_DblClick()
    If Me.lstAbility.ListIndex = -1 Or Me.lstSub.ListCount > 0 Then Exit Sub
    If Not ConfirmTier5Change() Then Exit Sub
    If AddAbility() Then
        ShowAvailable True
        If Not mblnNoFocus Then Me.usrList.SetFocus
        SetDirty
    End If
End Sub

Private Sub lstAbility_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTier As Long
    Dim lngAbility As Long

    If Button <> vbLeftButton Or Not GetUserChoices(lngTier, lngAbility, -1) Then Exit Sub
    Me.usrList.Selected = 0
    mblnMouse = ListAbilityClick()
    If mblnMouse Then menDragState = dragMouseDown
End Sub

Private Sub lstAbility_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            Me.lstAbility.OLEDropMode = vbOLEDropNone
            StartDrag
        End If
    End If
End Sub

Private Sub lstAbility_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
    ListAbilityClick
End Sub

' Show Details and Selectors, and return TRUE if we can drag this ability (ie: it has no selectors)
Private Function ListAbilityClick() As Boolean
    Dim lngTier As Long
    Dim lngAbility As Long
    
    ListboxClear Me.lstSub
    If Me.lstAbility.ListIndex = -1 Then Exit Function
    GetUserChoices lngTier, lngAbility, -1
    ShowDetails lngTier, lngAbility, 0, 0
    If db.Tree(mlngTree).Tier(lngTier).Ability(lngAbility).SelectorStyle <> sseNone Then ShowSelectors lngTier, lngAbility Else ListAbilityClick = True
End Function

Private Sub lstAbility_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List" ' Me.lstAbility.Text
End Sub

Private Sub lstAbility_OLECompleteDrag(Effect As Long)
    If Not mblnDragComplete Then
        RemoveBlanks build.Tree(mlngBuildTree)
        ShowAbilities
        ShowDropSlots
        Me.usrList.Active = 0
    End If
    mblnDragComplete = False
End Sub

Private Sub lstAbility_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFText) Then
        If Data.GetData(vbCFText) = "List" Then Exit Sub
    End If
    If mlngSourceIndex Then
        ClearSlot CInt(mlngSourceIndex)
        ShowAvailable True
        SetDirty
    End If
End Sub


' ************* SELECTORS *************


Private Sub lstSub_Click()
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    
    If mblnMouse Then
        mblnMouse = False
    Else
        GetUserChoices lngTier, lngAbility, lngSelector
        ShowDetails lngTier, lngAbility, lngSelector, 0
    End If
End Sub

Private Sub lstSub_DblClick()
    If Me.lstSub.ListIndex = -1 Then Exit Sub
    If Not ConfirmTier5Change() Then Exit Sub
    If AddAbility() Then
        ShowAvailable True
        If Not mblnNoFocus Then Me.usrList.SetFocus
        SetDirty
    End If
End Sub

Private Sub lstSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long

    If Not GetUserChoices(lngTier, lngAbility, lngSelector) Then Exit Sub
    Me.usrList.Selected = 0
    ShowDetails lngTier, lngAbility, lngSelector, 0
    mblnMouse = True
    menDragState = dragMouseDown
End Sub

Private Sub lstSub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Or Not mblnMouse Then Exit Sub
    If menDragState = dragMouseDown Then
        menDragState = dragMouseMove
        msngDownX = X
        msngDownY = Y
    ElseIf menDragState = dragMouseMove Then
        ' Only start dragging if mouse actually moved
        If X <> msngDownX Or Y <> msngDownY Then
            Me.lstAbility.OLEDropMode = vbOLEDropNone
            StartDrag
        End If
    End If
End Sub

Private Sub lstSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    menDragState = dragNormal
End Sub

Private Sub lstSub_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List" ' Me.lstSub.Text
End Sub

Private Sub lstSub_OLECompleteDrag(Effect As Long)
    If Not mblnDragComplete Then
        RemoveBlanks build.Tree(mlngBuildTree)
        ShowDropSlots
        Me.usrList.Active = 0
        ShowAbilities
    End If
    mblnDragComplete = False
End Sub


' ************* FILTERS *************


Private Sub usrchkShowAll_UserChange()
    ShowAvailable False
End Sub

Private Sub cboTree_Click()
    If mblnOverride Then Exit Sub
    TreeClick
    SaveBackup
End Sub

Private Sub TreeClick()
    NoSelection
    If Me.cboTree.ListIndex = -1 Then Exit Sub
    mlngTree = SeekTree(Me.cboTree.Text, peEnhancement)
    mlngBuildTree = FindBuildTree(Me.cboTree.Text)
    mlngMaxTier = GetMaxTier(mlngBuildTree)
    Me.usrList.GotoTop
    ShowAbilities
    ShowAvailable False
End Sub

Private Function GetMaxTier(plngBuildTree As Long) As Long
    Dim lngTier As Long
    
    With build.Tree(plngBuildTree)
        Select Case .TreeType
            Case tseClass: GetMaxTier = CapClassTreeTier(.TreeName, .ClassLevels)
            Case tseGlobal, tseRaceClass: GetMaxTier = CapClassTreeTier(.TreeName, build.MaxLevels)
            Case tseRace: GetMaxTier = 4
        End Select
    End With
End Function

Private Function CapClassTreeTier(pstrTreeName As String, ByVal plngTier As Long) As Long
    Dim lngCap As Long
    
    If plngTier > 5 Then lngCap = 5 Else lngCap = plngTier
'    If lngCap = 5 And Len(build.Tier5) <> 0 And pstrTreeName <> build.Tier5 Then lngCap = 4
    If lngCap = 5 And build.MaxLevels < 12 Then lngCap = 4
    CapClassTreeTier = lngCap
End Function

Private Sub ShaveTree(pstrTreeName As String)
    Dim lngTree As Long
    Dim lngLast As Long
    Dim i As Long
    
    lngTree = FindBuildTree(pstrTreeName)
    If lngTree = 0 Then Exit Sub
    With build.Tree(lngTree)
        For lngLast = .Abilities To 1 Step -1
            If .Ability(lngLast).Tier < 5 Then Exit For
        Next
        If lngLast = 0 Then
            .Abilities = 0
            Erase .Ability
        ElseIf .Abilities <> lngLast Then
            .Abilities = lngLast
            ReDim Preserve .Ability(1 To .Abilities)
        End If
    End With
End Sub

Private Sub ShowAvailable(pblnPreserveTopIndex As Boolean)
    Dim lngTopIndex As Long
    Dim lngSpent() As Long
    Dim typCheck As BuildAbilityType
    Dim lngTier As Long
    Dim i As Long
    
    If pblnPreserveTopIndex Then lngTopIndex = Me.lstAbility.TopIndex
    ListboxClear Me.lstSub
    ListboxClear Me.lstAbility
    If mlngTree = 0 Then Exit Sub
    GetSpentInTree db.Tree(mlngTree), build.Tree(mlngBuildTree), lngSpent, 0
    For lngTier = 0 To mlngMaxTier
        typCheck.Tier = lngTier
        With db.Tree(mlngTree).Tier(lngTier)
            For i = 1 To .Abilities
                Do
                    If AbilityTaken(build.Tree(mlngBuildTree), lngTier, i) Then Exit Do
                    typCheck.Ability = i
                    If Not Me.usrchkShowAll.Value Then
                        If CheckErrors(db.Tree(mlngTree), typCheck, lngSpent) Then Exit Do
                    End If
                    ListboxAddItem Me.lstAbility, lngTier & ": " & .Ability(i).Abbreviation, CreateAbilityID(lngTier, i)
                Loop Until True
            Next
        End With
    Next
    If pblnPreserveTopIndex Then
        With Me.lstAbility
            If lngTopIndex > .ListCount - 1 Then lngTopIndex = .ListCount - 1
            If lngTopIndex <> -1 Then .TopIndex = lngTopIndex
        End With
    End If
End Sub

Private Sub ShowSelectors(plngTier As Long, plngAbility As Long)
    Dim blnSelector() As Boolean
    Dim i As Long

    ListboxClear Me.lstSub
    With db.Tree(mlngTree).Tier(plngTier).Ability(plngAbility)
        GetSelectors db.Tree(mlngTree), plngTier, plngAbility, blnSelector
        For i = 1 To .Selectors
            If blnSelector(i) Then ListboxAddItem Me.lstSub, .Selector(i).SelectorName, i
        Next
    End With
End Sub


' ************* DETAILS *************


Private Sub ShowDetails(ByVal plngTier As Long, ByVal plngAbility As Long, ByVal plngSelector As Long, ByVal plngIndex As Long)
    Dim lngCost As Long
    Dim enReq As ReqGroupEnum
    Dim lngLevels As Long
    Dim lngClassLevels As Long
    Dim lngBuildLevels As Long
    Dim lngSpent() As Long
    Dim lngTotal As Long
    Dim lngProg As Long
    Dim i As Long
    
    ClearDetails False
    With db.Tree(mlngTree).Tier(plngTier).Ability(plngAbility)
        ' Caption
        Me.lblDetails.Caption = "Tier " & plngTier & ": " & .AbilityName
        ' Description
        If Len(.Descrip) Then Me.usrDetails.AddDescrip .Descrip, MakeWiki(db.Tree(mlngTree).Wiki) & TierLink(plngTier)
'        ' Class
'        If .Class(0) Then
'            Me.usrDetails.AddText "Requires class:"
'            For i = 1 To ceClasses - 1
'                If .Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
'            Next
'        End If
        ' Reqs
        For enReq = rgeAll To rgeNone
            If plngSelector = 0 Then ShowDetailsReqs Me.usrDetails, mlngTree, .Req(enReq), enReq, 0 Else ShowDetailsReqs Me.usrDetails, mlngTree, .Selector(plngSelector).Req(enReq), enReq, 0
        Next
        ' Rank reqs
        If plngSelector = 0 Then ShowRankReqs Me.usrDetails, mlngTree, .RankReqs, .Rank Else ShowRankReqs Me.usrDetails, mlngTree, .Selector(plngSelector).RankReqs, .Selector(plngSelector).Rank
        ' Levels
        GetLevelReqs db.Tree(mlngTree).TreeType, plngTier, plngAbility, lngLevels, lngClassLevels
        lngBuildLevels = GetBuildLevelReq(lngLevels, lngClassLevels, GetClass(mlngTree))
        If lngBuildLevels > 1 Or lngBuildLevels = -1 Then
            Me.usrDetails.AddText "Level requirements:"
            If lngClassLevels Then Me.usrDetails.AddText " - " & lngClassLevels & " class levels"
            If lngLevels Then Me.usrDetails.AddText " - " & lngLevels & " character levels"
            If lngBuildLevels = -1 Then
                Me.usrDetails.AddText " - Unattainable for this build"
            ElseIf lngBuildLevels > 1 Then
                Me.usrDetails.AddText " - " & lngBuildLevels & " build levels"
            End If
        End If
        ' Error?
        If plngIndex <> 0 Then
            gstrError = vbNullString
            GetSpentInTree db.Tree(mlngTree), build.Tree(mlngBuildTree), lngSpent, lngTotal
            If CheckErrors(db.Tree(mlngTree), build.Tree(mlngBuildTree).Ability(plngIndex), lngSpent) Then Me.usrDetails.AddErrorText "Error: " & gstrError
        End If
        Me.usrDetails.Refresh
        ' Ranks
        Me.lblRanks.Caption = "Ranks: " & .Ranks
        Me.lblRanks.Visible = True
        ' Cost
        Me.lblCost.Caption = CostDescrip(db.Tree(mlngTree).Tier(plngTier).Ability(plngAbility), plngSelector)
        Me.lblCost.Visible = True
        ' Spent in tree
        lngProg = GetSpentReq(db.Tree(mlngTree).TreeType, plngTier, plngAbility)
        If lngProg Then
            Me.lblProg.Caption = lngProg & " AP spent in tree"
            Me.lblProg.Visible = True
        Else
            Me.lblProg.Visible = False
        End If
    End With
End Sub

Private Function GetClass(plngTree As Long) As ClassEnum
    Dim typClassSplit() As ClassSplitType
    Dim lngClass As Long
    Dim lngLevels As Long
    Dim enClass As ClassEnum
    Dim i As Long
    
    For lngClass = 0 To GetClassSplit(typClassSplit) - 1
        enClass = typClassSplit(lngClass).ClassID
        With db.Class(enClass)
            For i = 1 To .Trees
                If .Tree(i) = db.Tree(plngTree).TreeName Then
                    If lngLevels < typClassSplit(lngClass).Levels Then
                        lngLevels = typClassSplit(lngClass).Levels
                        GetClass = enClass
                    End If
                    Exit For
                End If
            Next
        End With
    Next
End Function

Private Sub ShowDetailsReqs(pusrdet As userDetails, plngTree As Long, ptypReqList As ReqListType, penGroup As ReqGroupEnum, plngRank As Long)
    Dim strText As String
    Dim i As Long
    
    If ptypReqList.Reqs = 0 Then Exit Sub
    If plngRank < 2 Then strText = "Requires " Else strText = "Rank " & plngRank & " requires "
    strText = strText & LCase$(GetReqGroupName(penGroup)) & " of:"
    If plngRank = 0 Then
        pusrdet.AddText "Requires " & LCase$(GetReqGroupName(penGroup)) & " of:"
    Else
        pusrdet.AddText "Rank " & plngRank & " requires " & LCase$(GetReqGroupName(penGroup)) & " of:"
    End If
    For i = 1 To ptypReqList.Reqs
        pusrdet.AddText " - " & PointerDisplay(ptypReqList.Req(i), True, plngTree)
    Next
End Sub

Private Sub ShowRankReqs(pusrdet As userDetails, plngTree As Long, pblnRankReqs As Boolean, ptypRank() As RankType)
    Dim enReq As ReqGroupEnum
    Dim lngRank As Long
    Dim i As Long
    
    If Not pblnRankReqs Then Exit Sub
    For lngRank = 2 To 3
        With ptypRank(lngRank)
            ' Class
            If .Class(0) Then
                pusrdet.AddText "Rank " & lngRank & " requires Class:"
                For i = 1 To ceClasses - 1
                    If .Class(i) Then Me.usrDetails.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
                Next
            End If
            ' Reqs
            For enReq = rgeAll To rgeNone
                ShowDetailsReqs pusrdet, plngTree, .Req(enReq), enReq, lngRank
            Next
        End With
    Next
End Sub

Private Sub ClearDetails(pblnClearLabel As Boolean)
    If pblnClearLabel Then Me.lblDetails.Caption = "Details"
    Me.usrDetails.Clear
    Me.lblRanks.Visible = False
    Me.lblCost.Visible = False
    Me.lblProg.Visible = False
End Sub

Private Sub NoSelection()
    If mlngBuildTree = 0 Then Exit Sub
    With Me.lstSub
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    With Me.lstAbility
        If .ListIndex <> -1 Then .Selected(.ListIndex) = False
    End With
    Me.usrList.Selected = 0
    ClearDetails True
    On Error Resume Next
    If Not mblnNoFocus Then Me.usrList.SetFocus
End Sub

Private Sub Form_Click()
'    If mlngTab = 1 Then NoSelection
End Sub

Private Sub picTab_Click(Index As Integer)
    If Index = 1 Then NoSelection
End Sub

Private Sub picTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then ActiveCell -1, -1
End Sub


' ************* GUIDE *************


Private Sub InitGuide()
    GuideTotals
    LoadGuideTrees
    InitGrid
End Sub

Private Sub GuideTotals()
    Dim lngSpent As Long
    Dim lngAP As Long
    Dim lngRacialSpent As Long
    Dim lngRacialAvailable
    Dim lngUniversalSpent As Long  'New as of 20210421
    Dim lngUniversalAvailable
    Dim i As Long
    
    If Guide.Enhancements = 0 Then
        lngRacialAvailable = build.RacialAP
        lngUniversalAvailable = build.UniversalAP
        If db.Race(build.Race).Type = rteIconic Then
            If build.MaxLevels > 14 Then mlngGuideLevel = 14 Else mlngGuideLevel = build.MaxLevels
            mlngGuideAP = mlngGuideLevel * 4
        Else
            mlngGuideLevel = 1
            mlngGuideAP = 4
        End If
        mlngGuideAP = mlngGuideAP
    Else
        ' Calculate how many RacialAP/UniversalAP spent
        For i = 1 To Guide.Enhancements
            With Guide.Enhancement(i)
                If db.Tree(.TreeID).TreeType = tseRace Then lngRacialSpent = lngRacialSpent + .Cost
                If db.Tree(.TreeID).TreeType = tseGlobal Then lngUniversalSpent = lngUniversalSpent + .Cost
            End With
        Next
        If lngRacialSpent > build.RacialAP Then lngRacialSpent = build.RacialAP
        lngRacialAvailable = build.RacialAP - lngRacialSpent
        If lngUniversalSpent > build.UniversalAP Then lngUniversalSpent = build.UniversalAP
        lngUniversalAvailable = build.UniversalAP - lngUniversalSpent
        
        ' Get totals
        With Guide.Enhancement(Guide.Enhancements)
            If .Style = geBankAP Then mlngGuideLevel = .Level + 1 Else mlngGuideLevel = .Level
            If mlngGuideLevel > build.MaxLevels Then mlngGuideLevel = build.MaxLevels
            If mlngGuideLevel > 20 Then lngAP = 80 Else lngAP = mlngGuideLevel * 4
            lngSpent = .Spent - lngRacialSpent - lngUniversalSpent
            mlngGuideAP = lngAP - lngSpent
            If mlngGuideAP = 0 And mlngGuideLevel < 20 Then
                mlngGuideLevel = mlngGuideLevel + 1
                mlngGuideAP = mlngGuideAP + 4
            End If
        End With
    End If
    
    Me.lblGuideLevel.Caption = "Cur. Level: " & mlngGuideLevel
    'If lngRacialAvailable Then
        Me.lblGuideAP.Caption = "Available: " & mlngGuideAP & " + " & lngRacialAvailable & "r + " & lngUniversalAvailable & "u"
    'Else
    '    Me.lblGuideAP.Caption = "Available: " & mlngGuideAP & " AP"
    'End If
    If mlngGuideAP < 0 Then Me.lblGuideAP.ForeColor = cfg.GetColor(cgeWorkspace, cveTextError) Else Me.lblGuideAP.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    'If lngRacialSpent Then
        Me.lblGuideSpent.Caption = "Spent: " & lngSpent & " + " & lngRacialSpent & "r +" & lngUniversalSpent & "u"
    'Else
    '    Me.lblGuideSpent.Caption = "Spent: " & lngSpent & " AP"
    'End If
    If lngSpent > HeroicLevels() * 4 Then Me.lblGuideSpent.ForeColor = cfg.GetColor(cgeWorkspace, cveTextError) Else Me.lblGuideSpent.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
End Sub

Private Sub usrchkGuideTree_UserChange()
    LoadGuideTrees
End Sub

Private Sub LoadGuideTrees()
    Dim enClass As ClassEnum
    Dim i As Long

    ListboxClear Me.lstGuideTree
    For i = 1 To Guide.Trees
        If Me.usrchkGuideTree.Value Or Guide.Tree(i).BuildTreeID <> 0 Then
            ListboxAddItem Me.lstGuideTree, Guide.Tree(i).Display, i
        End If
    Next
    ListboxAddItem Me.lstGuideTree, "(Commands)", 0
    Me.picSelector.Visible = False
    Me.picGuideTree.Visible = True
End Sub

Private Sub lstGuideTree_Click()
    ListboxClear Me.lstGuideAbility
    mlngGuideTree = 0
    If Me.lstGuideTree.ListIndex = -1 Then Exit Sub
    With Me.lstGuideTree
        mlngGuideTree = .ItemData(.ListIndex)
    End With
    ShowGuideAvailable False
    SaveBackup
End Sub


' ************* GUIDE (ABILITY FILTERS) *************


Private Sub cboGuideAbility_Click()
    If mblnOverride Then Exit Sub
    ShowGuideAvailable False
    SaveBackup
End Sub

Private Sub ShowGuideAvailable(pblnPreserveTopIndex As Boolean)
    Dim ActiveTree() As Boolean
    Dim blnResetAll As Boolean
    Dim strDisplay As String
    Dim lngTree As Long
    Dim lngBuildTree As Long
    Dim lngSpent() As Long
    Dim typCheck As BuildAbilityType
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngRank As Long
    Dim lngMaxRanks As Long
    Dim lngTopIndex As Long
    Dim enGuideFilter As GuideFilterEnum
    Dim i As Long

    If pblnPreserveTopIndex Then lngTopIndex = Me.lstGuideAbility.TopIndex
    ListboxClear Me.lstGuideAbility
    If mlngGuideTree = 0 Then
        If Me.lstGuideTree.ListIndex = -1 Then Exit Sub
        ' 0=unknown, 1-99=GuideTreeID, 100=Reset all trees, 101-199=Reset tree (ID-100), 200=Bank remaining AP
        ListboxAddItem Me.lstGuideAbility, "Bank " & mlngGuideAP & " AP", 200
        ReDim ActiveTree(Guide.Trees)
        For i = 1 To Guide.Enhancements
            ActiveTree(Guide.Enhancement(i).GuideTreeID) = True
        Next
        For i = 1 To Guide.Trees
            If ActiveTree(i) Then
                With Guide.Tree(i)
                    If .Duplicate Then strDisplay = .TreeName & " (" & db.Class(.Class).Initial(3) & ")" Else strDisplay = .TreeName
                    ListboxAddItem Me.lstGuideAbility, "Reset " & strDisplay, 100 + Guide.Tree(i).BuildGuideTreeID
                End With
                blnResetAll = True
            End If
        Next
        If blnResetAll Then ListboxAddItem Me.lstGuideAbility, "Reset all trees", 100
        Exit Sub
    End If
    enGuideFilter = GuideFilter()
    With Guide.Tree(mlngGuideTree)
        lngTree = .TreeID
        lngBuildTree = .BuildTreeID
    End With
    GetSpentInTree db.Tree(lngTree), Guide.Tree(mlngGuideTree).BuildTree, lngSpent, 0
    For lngTier = 0 To Guide.Tree(mlngGuideTree).MaxTier
        typCheck.Tier = lngTier
        With db.Tree(lngTree).Tier(lngTier)
            For lngAbility = 1 To .Abilities
                lngMaxRanks = .Ability(lngAbility).Ranks
                If lngMaxRanks = 0 Then lngMaxRanks = 1
                For lngRank = 1 To lngMaxRanks
                    Do
                        typCheck.Tier = lngTier
                        typCheck.Ability = lngAbility
                        typCheck.Selector = 0 ' Filled in later by GuideSelected()
                        typCheck.Rank = lngRank
                        If GuideAbilityTaken(typCheck) Then Exit Do
                        If Not GuideSelected(typCheck) Then Exit Do
                        If Not (enGuideFilter = gfeAllSelected Or enGuideFilter = gfeAllEnhancements) Then
                            If GuideAbilityML(mlngGuideTree, lngTier, lngAbility, typCheck.Selector, lngRank) > mlngGuideLevel Then Exit Do
                            If CheckGuideAvailable(db.Tree(lngTree), Guide.Tree(mlngGuideTree).BuildTree, typCheck, lngSpent) Then Exit Do
                        End If
                        ListboxAddItem Me.lstGuideAbility, GuideBuildAbilityDisplay(lngTree, typCheck), CreateAbilityID(lngTier, lngAbility, lngRank, typCheck.Selector)
                    Loop Until True
                Next
            Next
        End With
    Next
    If pblnPreserveTopIndex Then
        With Me.lstAbility
            If lngTopIndex > .ListCount - 1 Then lngTopIndex = .ListCount - 1
            If lngTopIndex <> -1 Then .TopIndex = lngTopIndex
        End With
    End If
End Sub

Private Function GuideFilter() As GuideFilterEnum
    If Me.cboGuideAbility.ListIndex = -1 Then GuideFilter = gfeUnknown Else GuideFilter = Me.cboGuideAbility.ListIndex
End Function

Private Function GuideBuildAbilityDisplay(plngTree As Long, ptypAbility As BuildAbilityType) As String
    Dim strReturn As String
    Dim typPointer As PointerType
    
    With ptypAbility
        With db.Tree(plngTree).Tier(.Tier).Ability(.Ability)
            If ptypAbility.Selector = 0 Then
                strReturn = .Abbreviation
            ElseIf .SelectorOnly Then
                strReturn = .Selector(ptypAbility.Selector).SelectorName
            Else
                strReturn = .Abbreviation & ": " & .Selector(ptypAbility.Selector).SelectorName
            End If
            strReturn = ptypAbility.Tier & ": " & strReturn
            If .Ranks > 1 Then strReturn = strReturn & Left$(" III", ptypAbility.Rank + 1)
        End With
    End With
    GuideBuildAbilityDisplay = strReturn
End Function

' Returns TRUE if this tier+ability+rank has been taken, but FALSE if it is reset later
Private Function GuideAbilityTaken(ptypAbility As BuildAbilityType) As Boolean
    Dim blnTaken As Boolean
    Dim i As Long
    
    With build.Guide
        For i = 1 To .Enhancements
            With .Enhancement(i)
                Select Case .ID
                    Case 100, 100 + Guide.Tree(mlngGuideTree).BuildGuideTreeID: blnTaken = False
                    Case Guide.Tree(mlngGuideTree).BuildGuideTreeID: If .Tier = ptypAbility.Tier And .Ability = ptypAbility.Ability And .Rank = ptypAbility.Rank Then blnTaken = True
                End Select
            End With
        Next
    End With
    GuideAbilityTaken = blnTaken
End Function

' Returns TRUE if we satisfy the "only choose selected enhancements" filter
Private Function GuideSelected(ptypAbility As BuildAbilityType) As Boolean
    Dim lngBuildTree As Long
    Dim i As Long

    Select Case GuideFilter()
        Case gfeAllAvailable, gfeAllEnhancements
            GuideSelected = True
            Exit Function
    End Select
    If mlngGuideTree = 0 Then Exit Function
    lngBuildTree = Guide.Tree(mlngGuideTree).BuildTreeID
    If lngBuildTree = 0 Then Exit Function
    With build.Tree(lngBuildTree)
        For i = 1 To .Abilities
            If .Ability(i).Tier = ptypAbility.Tier And .Ability(i).Ability = ptypAbility.Ability Then
                ptypAbility.Selector = .Ability(i).Selector
                If .Ability(i).Rank >= ptypAbility.Rank Then GuideSelected = True
                Exit For
            End If
        Next
    End With
End Function

' Returns TRUE if errors found
Private Function CheckGuideAvailable(ptypTree As TreeType, ptypBuildTree As BuildTreeType, ptypAbility As BuildAbilityType, plngSpent() As Long) As Boolean
    Dim blnPassChecks As Boolean
    
    CheckGuideAvailable = True
    ' Levels (eg: 3 class levels for tier 3, 12 build levels for tier 5, 4 character levels for racial core 2, etc...)
    If CheckLevels(ptypBuildTree, ptypAbility.Tier, ptypAbility.Ability) Then Exit Function
    ' Points in tree
    If CheckSpentInTreeAbility(ptypTree, ptypAbility, plngSpent) Then Exit Function
    With ptypAbility
        Do
'            ' Class checks (eg: Half-Orc Tier 4: Power Rage requires 2 barbarian levels for rank 1, 6 for rank 2)
'            With ptypTree.Tier(.Tier).Ability(.Ability)
'                Do
'                    If CheckAbilityClassLevels(.Class, .ClassLevel, 1) Then Exit Do
'                    If .RankReqs And ptypAbility.Rank > 1 Then
'                        If CheckAbilityClassLevels(.Rank(ptypAbility.Rank).Class, .Rank(ptypAbility.Rank).ClassLevel, ptypAbility.Rank) Then Exit Do
'                    End If
'                    blnPassChecks = True
'                Loop Until True
'            End With
'            If Not blnPassChecks Then Exit Do
'            blnPassChecks = False
            ' All/One/None
            If .Selector = 0 Then
                If Not CheckGuideReqs(ptypTree.Tier(.Tier).Ability(.Ability).Req, .Rank, False) Then Exit Do
            Else
                If Not CheckGuideReqs(ptypTree.Tier(.Tier).Ability(.Ability).Selector(.Selector).Req, .Rank, False) Then Exit Do
            End If
            ' Ranks
            If .Rank > 1 And ptypTree.Tier(.Tier).Ability(.Ability).RankReqs Then
                If Not CheckGuideReqs(ptypTree.Tier(.Tier).Ability(.Ability).Rank(.Rank).Req, .Rank, True) Then Exit Do
                If .Selector <> 0 Then
                    If Not CheckGuideReqs(ptypTree.Tier(.Tier).Ability(.Ability).Selector(.Selector).Rank(.Rank).Req, .Rank, True) Then Exit Do
                End If
            End If
            blnPassChecks = True
        Loop Until True
    End With
    CheckGuideAvailable = Not blnPassChecks
End Function

' Returns TRUE if we successfully satisfy all abilityreqs
Private Function CheckGuideReqs(ptypReqList() As ReqListType, ByVal plngRanks As Long, pblnRankReq As Boolean) As Boolean
    Dim enReq As ReqGroupEnum
    
    For enReq = rgeAll To rgeNone
        If CheckGuideReq(ptypReqList(enReq), enReq, plngRanks, pblnRankReq) Then Exit Function
    Next
    CheckGuideReqs = True
End Function

' Returns TRUE if we fail any reqs
Private Function CheckGuideReq(ptypReqList As ReqListType, penReq As ReqGroupEnum, plngRanks As Long, pblnRankReq As Boolean) As Boolean
    Dim lngMatches As Long
    Dim lngGuide As Long
    Dim strTaken As String
    Dim strMissing As String
    Dim i As Long

    If ptypReqList.Reqs = 0 Then Exit Function
    For i = 1 To ptypReqList.Reqs
        With ptypReqList.Req(i)
            Select Case .Style
                Case peFeat
                    If CheckAbilityFeat(.Feat, .Selector) Then
                        If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, 1)
                        lngMatches = lngMatches + 1
                    Else
                        If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, 1)
                    End If
                Case peDestiny
                    If CheckAbility(build.Destiny, ptypReqList.Req(i), penReq, plngRanks) Then
                        If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                        lngMatches = lngMatches + 1
                    Else
                        If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                    End If
                Case peEnhancement
                    lngGuide = FindGuideTree(.Tree)
                    If lngGuide <> 0 Then
                        If CheckAbility(Guide.Tree(lngGuide).BuildTree, ptypReqList.Req(i), penReq, plngRanks) Then
                            If Len(strTaken) = 0 Then strTaken = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                            lngMatches = lngMatches + 1
                        Else
                            If Len(strMissing) = 0 Then strMissing = PointerDisplay(ptypReqList.Req(i), True, .Tree)
                        End If
                    End If
            End Select
        End With
    Next
    Select Case penReq
        Case rgeAll
            If lngMatches < ptypReqList.Reqs Then
                If pblnRankReq Then gstrError = "Rank " & plngRanks & " requires " Else gstrError = "Requires "
                gstrError = gstrError & strMissing
                CheckGuideReq = True
            End If
        Case rgeOne
            If lngMatches < 1 Then
                gstrError = "Nothing taken from the 'One of' list"
                CheckGuideReq = True
            End If
        Case rgeNone
            If lngMatches > 0 Then
                If pblnRankReq Then gstrError = "Rank " & plngRanks & " excludes " Else gstrError = "Antireq for "
                gstrError = gstrError & strMissing
                CheckGuideReq = True
            End If
    End Select
End Function

Private Function FindGuideTree(ByVal plngTree As Long)
    Dim i As Long

    For i = 1 To Guide.Trees
        If Guide.Tree(i).TreeID = plngTree Then
            FindGuideTree = i
            Exit Function
        End If
    Next
End Function


' ************* GUIDE (SELECTORS) *************


Private Sub lstGuideAbility_Click()
    If mblnOverride Then Exit Sub
    If Me.picSelector.Visible Then Me.lstGuideSub.ListIndex = -1
    ClearSelection
    GuideDetails 0
End Sub

Private Sub lstGuideSub_Click()
    If mblnOverride Then Exit Sub
    ClearSelection
    GuideDetails 0
End Sub

Private Sub ClearSelection()
    Dim i As Long
    
    For i = 1 To Guide.Enhancements
        If Guide.Enhancement(i).Selected Then
            Guide.Enhancement(i).Selected = False
            DrawCell i, 2
        End If
    Next
    mlngLastRowSelected = 0
    mblnMoveRows = False
End Sub

Private Sub lstGuideAbility_DblClick()
    GuideChooseAbility
End Sub

Private Sub lstGuideSub_DblClick()
    GuideChooseAbility
End Sub

Private Sub ShowGuideSelectors(plngTree As Long, plngTier As Long, plngAbility As Long)
    Dim blnSelector() As Boolean
    Dim i As Long

    ListboxClear Me.lstGuideSub
    With db.Tree(plngTree).Tier(plngTier).Ability(plngAbility)
        GetGuideSelectors db.Tree(plngTree), plngTier, plngAbility, blnSelector
        For i = 1 To .Selectors
            If blnSelector(i) Then ListboxAddItem Me.lstGuideSub, .Selector(i).SelectorName, i
        Next
    End With
    Me.picGuideTree.Visible = False
    Me.picSelector.Visible = True
End Sub

' The key difference between ability and feat selectors is that any given ability can only be taken once.
' eg: Improved Critical can be taken multiple times, but Tier 1 Elemental Arrows can only ever be taken once.
Private Sub GetGuideSelectors(ptypTree As TreeType, plngTier As Long, plngAbility As Long, pblnSelector() As Boolean)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim typTaken() As FeatTakenType
    Dim lngFeat As Long
    Dim i As Long

    With ptypTree.Tier(plngTier).Ability(plngAbility)
        ReDim pblnSelector(.Selectors)
        Select Case ptypTree.Tier(plngTier).Ability(plngAbility).SelectorStyle
            Case sseRoot
                For i = 1 To .Selectors
                    pblnSelector(i) = True
                Next
            Case sseShared
                If .Parent.Style = peFeat Then
                    ' Call IdentifyTakenFeats() to find all selectors taken
                    lngFeat = .Parent.Feat
                    IdentifyTakenFeats typTaken, build.MaxLevels
                    For i = 1 To .Selectors
                        pblnSelector(i) = typTaken(lngFeat).Selector(i)
                    Next
                Else
                    ' Flag the only selector taken
                    lngSelector = GetGuideSelector(.Parent)
                    pblnSelector(lngSelector) = True
                End If
            Case sseExclusive
                ' Initialize all choices as valid
                For i = 1 To .Selectors
                    pblnSelector(i) = True
                Next
                ' Parent choice is already taken
                lngSelector = GetGuideSelector(.Parent)
                pblnSelector(lngSelector) = False
                ' Siblings are also taken
                For i = 1 To .Siblings
                    lngSelector = GetGuideSelector(.Sibling(i))
                    pblnSelector(lngSelector) = False
                Next
        End Select
    End With
End Sub

' The following three functions are copy & pastes from the generic versions in basBuild
' Had to change this first one to properly point to the guide's build tree
Private Function GetGuideSelector(ptypPointer As PointerType) As Long
    With ptypPointer
        Select Case .Style
            Case peEnhancement
                GetGuideSelector = GetGuideSelectorChosen(Guide.Tree(mlngGuideTree).BuildTree, ptypPointer)
        End Select
    End With
End Function

Private Function GetGuideSelectorChosen(ptypBuildTree As BuildTreeType, ptypPointer As PointerType) As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim i As Long
    
    With ptypPointer
        lngTier = .Tier
        lngAbility = .Ability
    End With
    With ptypBuildTree
        For i = 1 To .Abilities
            If .Ability(i).Tier = lngTier Then
                If .Ability(i).Ability = lngAbility Then
                    GetGuideSelectorChosen = .Ability(i).Selector
                    Exit For
                End If
            End If
        Next
    End With
End Function

Private Sub chkBackToTrees_Click()
    If UncheckButton(Me.chkBackToTrees, mblnOverride) Then Exit Sub
    Me.picSelector.Visible = False
    ListboxClear Me.lstGuideSub
    Me.picGuideTree.Visible = True
End Sub

Private Sub HideGuideSelectors()
    Me.picSelector.Visible = False
    Me.picGuideTree.Visible = True
    Me.lstGuideSub.ListIndex = -1
    Me.lstGuideSub.Clear
End Sub


' ************* GUIDE (CHOOSE) *************


Private Sub GuideChooseAbility()
    Dim lngTree As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRank As Long
    Dim blnHideSelector As Boolean
    Dim lngEnhancements As Long
    Dim i As Long
    
    If Me.lstGuideAbility.ListIndex = -1 Then Exit Sub
    lngEnhancements = build.Guide.Enhancements
    If mlngGuideTree = 0 Then
        With build.Guide
            .Enhancements = .Enhancements + 1
            ReDim Preserve .Enhancement(1 To .Enhancements)
            With .Enhancement(.Enhancements)
                .ID = Me.lstGuideAbility.ItemData(Me.lstGuideAbility.ListIndex)
            End With
        End With
    Else
        lngTree = Guide.Tree(mlngGuideTree).TreeID
        SplitAbilityID Me.lstGuideAbility.ItemData(Me.lstGuideAbility.ListIndex), lngTier, lngAbility, lngRank, lngSelector
        If db.Tree(lngTree).Tier(lngTier).Ability(lngAbility).SelectorStyle <> sseNone And lngSelector = 0 Then
            If Me.lstGuideSub.ListIndex = -1 Then Exit Sub
            lngSelector = Me.lstGuideSub.ItemData(Me.lstGuideSub.ListIndex)
            HideGuideSelectors
            If lngSelector = 0 Then Exit Sub
        End If
        AddGuideAbility lngTree, lngTier, lngAbility, lngSelector, lngRank
    End If
    InitGuideEnhancements
    GuideTotals
    DrawGrid True
    ShowGuideAvailable True
    SetDirty
    If lngEnhancements = 0 Then EnableMenus
End Sub

Private Sub AddGuideAbility(plngTree As Long, plngTier As Long, plngAbility As Long, plngSelector As Long, plngRanks As Long)
    Dim lngHighestRank As Long
    Dim i As Long

    ' Identify how many ranks this choice actually slots (choosing rank 3 first slots ranks 1 and 2 if not already slotted)
    If plngRanks > 1 Then
        For i = 1 To Guide.Enhancements
            With Guide.Enhancement(i)
                Select Case .Style
                    Case geResetAllTrees
                        lngHighestRank = 0
                    Case geResetTree
                        If .GuideTreeID = mlngGuideTree Then lngHighestRank = 0
                    Case geEnhancement
                        If .GuideTreeID = mlngGuideTree And .Tier = plngTier And .Ability = plngAbility Then
                            If lngHighestRank < .Rank Then lngHighestRank = .Rank
                        End If
                End Select
            End With
        Next
    End If
    ' Now add each rank individually
    For i = lngHighestRank + 1 To plngRanks
        AddGuideAbilityRank plngTree, plngTier, plngAbility, plngSelector, i
    Next
End Sub

Private Sub AddGuideAbilityRank(plngTree As Long, plngTier As Long, plngAbility As Long, plngSelector As Long, plngRanks As Long)
    Dim lngID As Long

    ' Add new GuideTree if this is first enhancement taken from the tree
    If Guide.Tree(mlngGuideTree).BuildGuideTreeID = 0 Then
        With build.Guide
            .Trees = .Trees + 1
            ReDim Preserve .Tree(1 To .Trees)
            With .Tree(.Trees)
                .Class = Guide.Tree(mlngGuideTree).Class
                .TreeName = db.Tree(Guide.Tree(mlngGuideTree).TreeID).TreeName
            End With
            ReDim Preserve Guide.TreeLookup(.Trees)
            Guide.TreeLookup(.Trees) = mlngGuideTree
            Guide.Tree(mlngGuideTree).BuildGuideTreeID = .Trees
            lngID = .Trees
        End With
    Else
        lngID = Guide.Tree(mlngGuideTree).BuildGuideTreeID
    End If
    ' Slot this ability to end of guide tree
    With build.Guide
        .Enhancements = .Enhancements + 1
        ReDim Preserve .Enhancement(1 To .Enhancements)
        With .Enhancement(.Enhancements)
            .Tier = plngTier
            .Ability = plngAbility
            .Selector = plngSelector
            .Rank = plngRanks
            .ID = lngID
        End With
    End With
End Sub


' ************* GUIDE (SELECT ROWS) *************


Private Sub picGuide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim enMouse As MouseShiftEnum
    Dim blnActive As Boolean
    
    If ActiveCell(X, Y) Then
        If mlngRow < 1 Or mlngRow > Guide.Enhancements Or mlngCol <> 2 Then Exit Sub
        If Button <> vbLeftButton Then Exit Sub
        Guide.Enhancement(mlngRow).Selected = True
        DrawCell mlngRow, mlngCol
        DrawBorders mlngRow, mlngCol, True
        mlngLastRowSelected = mlngRow
    End If
    If Button = vbLeftButton And Not Me.tmrScroll.Enabled Then
        If Me.picGuide.Top + Y < 0 Then
            mlngDirection = -1
        ElseIf Me.picGuide.Top + Y > Me.picContainer.Height Then
            mlngDirection = 1
        Else
            Exit Sub
        End If
        If (mlngDirection = -1 And Me.picGuide.Top < 0) Or (mlngDirection = 1 And Me.picGuide.Top + Me.picGuide.Height > Me.picContainer.ScaleHeight) Then
            Me.tmrScroll.Interval = mlngInterval
            If mlngInterval > 30 Then mlngInterval = mlngInterval - 5
            Me.tmrScroll.Enabled = True
        End If
    End If
End Sub

Private Sub picGuide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim enMouse As MouseShiftEnum
    Dim blnCtrl As Boolean
    Dim blnShift As Boolean
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngSelections As Long
    Dim i As Long
     
    mlngInterval = 100
    ActiveCell X, Y
    If mlngRow < 1 Or mlngRow > Guide.Enhancements Or mlngCol <> 2 Then Exit Sub
    enMouse = ReadMouseShift(Shift)
    If (Button = vbLeftButton And enMouse = mseNormal) Or (Button = vbRightButton And mblnMoveRows = False And Guide.Enhancement(mlngRow).Selected = False) Then
        mblnMoveRows = False
        If Button = vbLeftButton Or Not Guide.Enhancement(mlngRow).Selected Then
            For i = 1 To Guide.Enhancements
                If i <> mlngRow And Guide.Enhancement(i).Selected Then
                    lngSelections = lngSelections + 1
                    Guide.Enhancement(i).Selected = False
                    DrawCell i, 2
                End If
            Next
        End If
    End If
    Select Case Button
        Case vbLeftButton
            Select Case enMouse
                Case mseNormal
                    Guide.Enhancement(mlngRow).Selected = (lngSelections > 1 Or Guide.Enhancement(mlngRow).Selected = False)
                    DrawCell mlngRow, mlngCol
                Case mseCtrl
                    Guide.Enhancement(mlngRow).Selected = Not Guide.Enhancement(mlngRow).Selected
                    DrawCell mlngRow, mlngCol
                Case mseShift
                    Guide.Enhancement(mlngRow).Selected = True
                    lngStart = mlngRow
                    lngEnd = mlngRow
                    If mlngLastRowSelected <> 0 Then
                        If mlngRow > mlngLastRowSelected Then lngStart = mlngLastRowSelected Else lngEnd = mlngLastRowSelected
                    End If
                    For i = lngStart To lngEnd
                        Guide.Enhancement(i).Selected = True
                        DrawCell i, 2
                    Next
            End Select
            If Guide.Enhancement(mlngRow).Selected Then mlngLastRowSelected = mlngRow Else mlngLastRowSelected = 0
            DrawBorders mlngRow, mlngCol, True
            GuideClicked True
        Case vbRightButton
            If Not mblnMoveRows Then
                Guide.Enhancement(mlngRow).Selected = True
                DrawCell mlngRow, 2
            End If
            DrawBorders mlngRow, mlngCol, True
            GuideClicked True
            PopupGuideMenu
    End Select
End Sub

Private Sub GuideClicked(pblnShowDetails As Boolean)
    mblnOverride = True
    Me.lstGuideAbility.ListIndex = -1
    Me.lstGuideSub.ListIndex = -1
    Me.picSelector.Visible = False
    Me.picGuideTree.Visible = True
    mblnOverride = False
    If pblnShowDetails Then
        With Guide.Enhancement(mlngRow)
            If .Selected And .Style = geEnhancement Then
                ShowGuideDetails .GuideTreeID, .Tier, .Ability, .Selector, .Rank, mlngRow
            Else
                ClearGuideDetails True
            End If
        End With
    End If
End Sub

Private Function ReadMouseShift(Shift As Integer) As MouseShiftEnum
    If (Shift And vbCtrlMask) > 0 Then
        ReadMouseShift = mseCtrl
    ElseIf (Shift And vbShiftMask) > 0 Then
        ReadMouseShift = mseShift
    Else
        ReadMouseShift = mseNormal
    End If
End Function

Private Sub picGuide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
    If mlngRow >= 1 And mlngRow <= Guide.Enhancements And mlngCol = 2 Then
        With Guide.Enhancement(mlngRow)
            If .Selected And .Style = geEnhancement Then ShowGuideDetails .GuideTreeID, .Tier, .Ability, .Selector, .Rank, mlngRow
        End With
    End If
End Sub

Private Function ActiveCell(psngX As Single, psngY As Single) As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngMaxRow As Long
    Dim i As Long
    
    lngMaxRow = build.Guide.Enhancements
    If lngMaxRow < 20 Then lngMaxRow = 20
    lngRow = (psngY \ mlngHeight) + 1
    For i = 1 To 6
        If psngX >= Col(i).Left And psngX <= Col(i).Right Then
            lngCol = i
            Exit For
        End If
    Next
    If lngRow < 1 Or lngRow > lngMaxRow Or lngCol < 1 Or lngCol > 6 Then
        DrawBorders mlngRow, mlngCol, False
        mlngRow = 0
        mlngCol = 0
    ElseIf lngRow <> mlngRow Or lngCol <> mlngCol Then
        DrawBorders mlngRow, mlngCol, False
        mlngRow = lngRow
        mlngCol = lngCol
        If mlngCol = 2 Then
            DrawBorders mlngRow, mlngCol, True
            ActiveCell = True
        End If
    End If
End Function

Private Sub GetCoords(plngRow As Long, plngCol As Long, plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long)
    plngLeft = Col(plngCol).Left
    plngTop = (plngRow - 1) * mlngHeight
    plngRight = Col(plngCol).Right
    plngBottom = plngTop + mlngHeight
End Sub

Private Sub PopupGuideMenu()
    Dim lngSelected As Long
    Dim strPlural As String
    Dim i As Long
    
    For i = 1 To Guide.Enhancements
        If Guide.Enhancement(i).Selected Then lngSelected = lngSelected + 1
    Next
    If lngSelected > 1 Then strPlural = "s"
    If mblnMoveRows Then
        Me.mnuGuide(0).Caption = "Insert Row" & strPlural & " Here"
        Me.mnuGuide(0).Enabled = Not Guide.Enhancement(mlngRow).Selected
        Me.mnuGuide(2).Enabled = Guide.Enhancement(mlngRow).Selected
        Me.mnuGuide(3).Enabled = (Guide.Enhancement(mlngRow).Selected And lngSelected = 1)
    Else
        Me.mnuGuide(0).Caption = "Move Row" & strPlural
        Me.mnuGuide(0).Enabled = True
        Me.mnuGuide(2).Enabled = True
        Me.mnuGuide(3).Enabled = (lngSelected = 1)
    End If
    Me.mnuGuide(2).Caption = "Delete Row" & strPlural
    PopupMenu Me.mnuMain(2)
End Sub

Private Sub mnuGuide_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuGuide(Index).Caption)
        Case "Insert Row Here", "Insert Rows Here": InsertRows
        Case "Move Row", "Move Rows": MoveRows
        Case "Delete Row", "Delete Rows": DeleteRows
        Case "Delete to End": If mlngRow = 1 Then ClearGuide Else SetRows mlngRow - 1
        Case "Delete All": If Ask("Clear Leveling Guide?") Then ClearGuide
    End Select
End Sub

Private Sub ClearGuide()
    With build.Guide
        Erase .Enhancement, .Tree
        .Enhancements = 0
        .Trees = 0
    End With
    With Guide
        Erase .Enhancement, .Tree, .TreeLookup
        .Enhancements = 0
        .Trees = 0
    End With
    RefreshGuide
    EnableMenus
End Sub

Private Sub DeleteRows()
    Dim lngInsert As Long
    Dim i As Long
    
    lngInsert = 1
    For i = 1 To Guide.Enhancements
        If Not Guide.Enhancement(i).Selected Then MoveRow i, lngInsert
    Next
    SetRows lngInsert - 1
End Sub

Private Sub MoveRow(plngFrom As Long, plngTo As Long)
    If plngFrom <> plngTo Then
        build.Guide.Enhancement(plngTo) = build.Guide.Enhancement(plngFrom)
        Guide.Enhancement(plngTo) = Guide.Enhancement(plngFrom)
    End If
    plngTo = plngTo + 1
End Sub

Private Sub MoveRows()
    Dim i As Long
    
    mblnMoveRows = True
    For i = 1 To Guide.Enhancements
        If Guide.Enhancement(i).Selected Then DrawCell i, 2
    Next
End Sub

Private Sub InsertRows()
    Dim typBuild() As BuildGuideEnhancementType
    Dim typGuide() As GuideEnhancementType
    Dim lngInsertionPoint As Long
    Dim lngInsert As Long
    Dim lngSelected As Long
    Dim lngPos As Long
    Dim i As Long
    
    If mlngRow >= 1 And mlngRow <= Guide.Enhancements Then
        ' Identify eventual insertion point
        For i = 1 To mlngRow
            If Not Guide.Enhancement(i).Selected Then lngInsertionPoint = lngInsertionPoint + 1
        Next
        ' Copy all selected rows to temporary holding area and "delete" from arrays
        ReDim typGuide(Guide.Enhancements)
        ReDim typBuild(Guide.Enhancements)
        lngInsert = 1
        For i = 1 To Guide.Enhancements
            If Guide.Enhancement(i).Selected Then
                lngSelected = lngSelected + 1
                typBuild(lngSelected) = build.Guide.Enhancement(i)
                typGuide(lngSelected) = Guide.Enhancement(i)
            Else
                MoveRow i, lngInsert
            End If
        Next
        ' Open hole in array at the insertion point
        For i = Guide.Enhancements To lngInsertionPoint + lngSelected Step -1
            MoveRow i - lngSelected, (i + 0) ' The 2nd parameter is incremented inside MoveRow, so we add 0 to avoid passing by reference
        Next
        ' Copy the moved rows back in from the temporary arrays
        For i = 1 To lngSelected
            lngPos = lngInsertionPoint + i - 1
            build.Guide.Enhancement(lngPos) = typBuild(i)
            Guide.Enhancement(lngPos) = typGuide(i)
        Next
    End If
    RefreshGuide lngInsertionPoint, lngSelected
End Sub

Private Sub SetRows(plngRows As Long)
    With build.Guide
        .Enhancements = plngRows
        If .Enhancements = 0 Then Erase .Enhancement Else ReDim Preserve .Enhancement(1 To .Enhancements)
    End With
    With Guide
        .Enhancements = plngRows
        If .Enhancements = 0 Then Erase .Enhancement Else ReDim Preserve .Enhancement(.Enhancements)
    End With
    RefreshGuide
    If build.Guide.Enhancements = 0 Then EnableMenus
End Sub

Private Sub RefreshGuide(Optional plngSelectFirst As Long, Optional plngSelectCount As Long)
    Dim lngScroll As Long
    Dim i As Long
    
    lngScroll = Me.scrollVertical.Value
    mlngLastRowSelected = 0
    mblnMoveRows = False
    ClearGuideDetails True
    InitLevelingGuide
    If plngSelectCount > 1 Then
        For i = plngSelectFirst To plngSelectFirst + plngSelectCount - 1
            Guide.Enhancement(i).Selected = True
        Next
    End If
    InitGuide
    If plngSelectCount = 1 Then DrawBorders plngSelectFirst, 2, True
    If Me.scrollVertical.Enabled Then
        Select Case lngScroll
            Case Is < Me.scrollVertical.Min: lngScroll = Me.scrollVertical.Min
            Case Is > Me.scrollVertical.Max: lngScroll = Me.scrollVertical.Max
        End Select
        Me.scrollVertical.Value = lngScroll
    End If
    SetDirty
End Sub


' ************* GUIDE (DETAILS) *************


Private Sub GuideDetails(plngIndex As Long)
    Dim lngItemData As Long
    Dim lngTree As Long
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngSelector As Long
    Dim lngRank As Long

    ' Identify ability
    If Me.lstGuideAbility.ListIndex = -1 Then
        ClearGuideDetails True
        HideGuideSelectors
        Exit Sub
    End If
    lngItemData = Me.lstGuideAbility.ItemData(Me.lstGuideAbility.ListIndex)
    If mlngGuideTree = 0 Then
        Me.usrdetGuide.Clear
        Select Case lngItemData
            Case 100
                Me.lblGuideDetails.Caption = "Reset All Trees"
                Me.usrdetGuide.AddText "Reset all your trees and re-spend all available AP starting at the current level."
            Case 101 To 199
                lngTree = lngItemData - 100
                Me.lblGuideDetails.Caption = "Reset " & Guide.Tree(lngTree).Display
                Me.usrdetGuide.AddText "You can re-spend all recovered APs starting at the current level."
            Case 200
                Me.lblGuideDetails.Caption = "Bank " & mlngGuideAP & " AP"
                Me.usrdetGuide.AddText "Save the remainder of your available AP until next level."
        End Select
        Me.usrdetGuide.Refresh
    Else
        SplitAbilityID lngItemData, lngTier, lngAbility, lngRank, lngSelector
        ' Selectors
        lngTree = Guide.Tree(mlngGuideTree).TreeID
        If db.Tree(lngTree).Tier(lngTier).Ability(lngAbility).SelectorStyle <> sseNone And lngSelector = 0 Then
            If Me.lstGuideSub.ListIndex = -1 Then
                ShowGuideSelectors lngTree, lngTier, lngAbility
            Else
                lngSelector = Me.lstGuideSub.ItemData(Me.lstGuideSub.ListIndex)
            End If
        Else
            HideGuideSelectors
        End If
        ShowGuideDetails mlngGuideTree, lngTier, lngAbility, lngSelector, lngRank, plngIndex
    End If
End Sub

Private Sub ShowGuideDetails(plngGuide As Long, ByVal plngTier As Long, ByVal plngAbility As Long, ByVal plngSelector As Long, plngRanks As Long, plngIndex As Long)
    Dim lngTree As Long
    Dim lngCost As Long
    Dim enReq As ReqGroupEnum
    Dim lngLevels As Long
    Dim lngClassLevels As Long
    Dim lngBuildLevels As Long
    Dim lngSpent() As Long
    Dim lngTotal As Long
    Dim lngProg As Long
    Dim lngColor As Long
    Dim i As Long

    ClearGuideDetails False
    lngTree = Guide.Tree(plngGuide).TreeID
    With db.Tree(lngTree).Tier(plngTier).Ability(plngAbility)
        ' Caption
        Me.lblGuideDetails.Caption = "Tier " & plngTier & ": " & .AbilityName
        ' Description
        If Len(.Descrip) Then Me.usrdetGuide.AddDescrip .Descrip, MakeWiki(db.Tree(lngTree).Wiki) & TierLink(plngTier)
        ' Class
'        If .Class(0) Then
'            Me.usrdetGuide.AddText "Requires class:"
'            For i = 1 To ceClasses - 1
'                If .Class(i) Then Me.usrdetGuide.AddText " - " & GetClassName(i) & " " & .ClassLevel(i)
'            Next
'        End If
        ' Reqs
        For enReq = rgeAll To rgeNone
            If plngSelector = 0 Then ShowDetailsReqs Me.usrdetGuide, lngTree, .Req(enReq), enReq, 0 Else ShowDetailsReqs Me.usrdetGuide, lngTree, .Selector(plngSelector).Req(enReq), enReq, 0
        Next
        ' Rank reqs
        If plngSelector = 0 Then ShowRankReqs Me.usrdetGuide, lngTree, .RankReqs, .Rank Else ShowRankReqs Me.usrdetGuide, lngTree, .Selector(plngSelector).RankReqs, .Selector(plngSelector).Rank
        ' Levels
        GetLevelReqs db.Tree(lngTree).TreeType, plngTier, plngAbility, lngLevels, lngClassLevels
        lngBuildLevels = GetBuildLevelReq(lngLevels, lngClassLevels, Guide.Tree(plngGuide).Class)
        If lngBuildLevels > 1 Or lngBuildLevels = -1 Then
            Me.usrDetails.AddText "Level requirements:"
            If lngClassLevels Then Me.usrDetails.AddText " - " & lngClassLevels & " class levels"
            If lngLevels Then Me.usrDetails.AddText " - " & lngLevels & " character levels"
            If lngBuildLevels = -1 Then
                Me.usrDetails.AddText " - Unattainable for this build"
            ElseIf lngBuildLevels > 1 Then
                Me.usrDetails.AddText " - " & lngBuildLevels & " build levels"
            End If
        End If
        ' Error?
        If plngIndex <> 0 Then
            With Guide.Enhancement(plngIndex)
                If .ErrorState Then Me.usrdetGuide.AddErrorText "Error: " & .ErrorText
            End With
        End If
        Me.usrdetGuide.Refresh
        ' Cost
        Me.lblGuideCost.Caption = CostDescrip(db.Tree(lngTree).Tier(plngTier).Ability(plngAbility), plngSelector, False)
        Me.lblGuideCost.Visible = True
        ' Spent in tree
        lngProg = GetSpentReq(db.Tree(lngTree).TreeType, plngTier, plngAbility)
        If lngProg Then
            Me.lblGuideSpentInTree.Caption = "Spent in Tree: " & lngProg & " AP"
            Me.lblGuideSpentInTree.Visible = True
        End If
        ' Build levels
        lngBuildLevels = GuideAbilityML(plngGuide, plngTier, plngAbility, plngSelector, plngRanks)
        With Me.lblGuideLevels
            If lngBuildLevels < 1 Then
                lngColor = cfg.GetColor(cgeWorkspace, cveTextError)
                .Caption = "Untrainable"
            Else
                lngColor = cfg.GetColor(cgeWorkspace, cveText)
                .Caption = "Min Level: " & lngBuildLevels
            End If
            If .ForeColor <> lngColor Then .ForeColor = lngColor
        End With
        Me.lblGuideLevels.Visible = True
    End With
End Sub

Private Sub ClearGuideDetails(pblnClearLabel As Boolean)
    If pblnClearLabel Then Me.lblGuideDetails.Caption = "Details"
    Me.usrdetGuide.Clear
    Me.lblGuideCost.Visible = False
    Me.lblGuideSpentInTree.Visible = False
    Me.lblGuideLevels.Visible = False
End Sub


' ************* GUIDE (DRAW GRID) *************


Private Sub InitGrid()
    Dim strWidestTree As String
    Dim strTree As String
    Dim lngWidth As Long
    Dim lngLeft As Long
    Dim lngBottom As Long
    Dim lngHeight As Long
    Dim lngAdjust As Long
    Dim i As Long

    With Me.picContainer
        mlngOffsetX = .TextWidth(" ")
        mlngHeight = .ScaleY(.ScaleY(Me.picContainer.ScaleHeight, vbTwips, vbPixels) \ 20, vbPixels, vbTwips)
        lngHeight = mlngHeight * 20 + PixelY
        lngBottom = .Top + .Height
        lngAdjust = .Height - lngHeight
        .Move .Left, lngBottom - lngHeight, .Width, lngHeight
    End With
    For i = 1 To 6
        Me.lblGuide(i).Top = Me.lblGuide(i).Top + lngAdjust
    Next
    With Me.scrollVertical
        .Move Me.picContainer.ScaleWidth - .Width, 0, .Width, Me.picContainer.ScaleHeight
        Me.picGuide.Move 0, 0, .Left, .Height
    End With
    For i = 1 To Guide.Trees
'        strTree = db.Tree(Guide.Tree(i).TreeID).Abbreviation
        strTree = Guide.Tree(i).Display
        With Me.picGuide
            If .TextWidth(strTree) > .TextWidth(strWidestTree) Then strWidestTree = strTree
        End With
    Next
    Erase Col
    InitColumn 1, "Level", vbCenter, lngWidth
    InitColumn 2, "Ability", vbLeftJustify, lngWidth, "Skip"
    InitColumn 3, "Tier", vbCenter, lngWidth
    InitColumn 4, "Tree", vbLeftJustify, lngWidth, strWidestTree
    InitColumn 5, "AP", vbCenter, lngWidth, "Tier"
    InitColumn 6, "Prog", vbCenter, lngWidth, "Tier"
    Col(2).Width = Me.picGuide.ScaleWidth - lngWidth - PixelX
    For i = 1 To 6
        With Col(i)
            .Left = lngLeft
            .Right = lngLeft + .Width
            lngLeft = .Right
        End With
    Next
    DrawGrid
End Sub

Private Sub InitColumn(plngCol As Long, pstrHeader As String, penAlign As AlignmentConstants, plngWidth As Long, Optional pstrWidth As String = vbNullString)
    Dim lngSpace As Long
    Dim strCheck As String
    
    If Len(pstrWidth) = 0 Then strCheck = pstrHeader Else strCheck = pstrWidth
    With Col(plngCol)
        .Align = penAlign
        .Header = pstrHeader
        If pstrWidth <> "Skip" Then .Width = Me.picGuide.TextWidth(strCheck) + mlngOffsetX * 2
        plngWidth = plngWidth + .Width
    End With
End Sub

Private Sub DrawGrid(Optional pblnScrollToEnd As Boolean = False)
    Dim lngLines As Long
    Dim lngHeight As Long
    Dim i As Long
    
    mlngLastRowSelected = 0
    mblnMoveRows = False
    With Me.picGuide
        .ForeColor = cfg.GetColor(cgeControls, cveText)
        .BackColor = cfg.GetColor(cgeControls, cveBackground)
        ' Columns
        For i = 1 To 6
            Me.lblGuide(i).Left = Col(i).Left + Me.picContainer.Left
            Me.lblGuide(i).Width = Col(i).Width
        Next
        Me.lblGuideSpent.Left = Me.picContainer.Left + Col(6).Right - Me.lblGuideSpent.Width
        ' Rows
        mlngHeight = .ScaleY(.ScaleY(Me.picContainer.ScaleHeight, vbTwips, vbPixels) \ 20, vbPixels, vbTwips)
        mlngOffsetY = (mlngHeight - .TextHeight("X")) \ 2
        lngLines = build.Guide.Enhancements
        If lngLines < 20 Then lngLines = 20
        .Move .Left, 0, .Width, mlngHeight * lngLines + PixelY
        .Cls
        For mlngRow = 1 To lngLines
            For mlngCol = 1 To 6
                DrawCell mlngRow, mlngCol
            Next
        Next
    End With
    With Me.scrollVertical
        If build.Guide.Enhancements > 20 Then
            .Max = lngLines - 20
            If pblnScrollToEnd Then .Value = .Max Else .Value = 0
            .Enabled = True
        Else
            .Enabled = False
        End If
    End With
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngForeColor As Long
    Dim lngBackColor As Long
    Dim blnHighlight As Boolean
    Dim strDisplay As String
    
    DrawBorders plngRow, plngCol, False
    If plngRow > 0 And plngRow <= Guide.Enhancements Then
        With Guide.Enhancement(plngRow)
            If plngCol = 1 Then
                strDisplay = .Level
            Else
                Select Case .Style
                    Case geBankAP
                        If plngCol = 2 Then strDisplay = "Bank " & .Bank & " AP"
                    Case geResetTree
                        If plngCol = 2 Then strDisplay = "Reset " & Guide.Tree(.GuideTreeID).Display
                    Case geResetAllTrees
                        If plngCol = 2 Then strDisplay = "Reset All Trees"
                    Case geEnhancement
                        Select Case plngCol
                            Case 2: strDisplay = FitDisplay(.Display, .RankText, plngCol)
                            Case 3: strDisplay = .Tier
                            Case 4: strDisplay = Guide.Tree(.GuideTreeID).Display
                            Case 5: strDisplay = .Cost
                            Case 6: strDisplay = .SpentInTree
                        End Select
                End Select
            End If
            If plngCol = 2 And .ErrorState = True Then lngForeColor = cfg.GetColor(cgeControls, cveTextError) Else lngForeColor = cfg.GetColor(cgeControls, cveText)
        End With
        ' Background
        GetCoords plngRow, plngCol, lngLeft, lngTop, lngRight, lngBottom
        If plngCol = 2 And Guide.Enhancement(plngRow).Selected Then
            If mblnMoveRows Then
                lngBackColor = cfg.GetColor(cgeControls, cveBackRelated)
            Else
                lngBackColor = cfg.GetColor(cgeControls, cveBackHighlight)
            End If
        Else
            lngBackColor = cfg.GetColor(cgeControls, cveBackground)
        End If
        Me.picGuide.Line (lngLeft + PixelX, lngTop + PixelY)-(lngRight - PixelX, lngBottom - PixelY), lngBackColor, BF
        With Me.picGuide
            Select Case Col(plngCol).Align
                Case vbLeftJustify: .CurrentX = lngLeft + mlngOffsetX
                Case vbCenter: .CurrentX = lngLeft + (Col(plngCol).Width - .TextWidth(strDisplay)) \ 2
                Case vbRightJustify: .CurrentX = lngRight - mlngOffsetX - .TextWidth(strDisplay)
            End Select
            .CurrentY = lngTop + mlngOffsetY
            .ForeColor = lngForeColor
        End With
        Me.picGuide.Print strDisplay
    End If
End Sub

Private Function FitDisplay(pstrName As String, ByVal pstrRank As String, plngCol As Long) As String
    Dim lngWidth As Long
    Dim lngLen As Long
    Dim strCheck As String
    
    lngWidth = Col(plngCol).Width - mlngOffsetX * 2
    strCheck = pstrName & pstrRank
    If Me.picGuide.TextWidth(strCheck) > lngWidth Then
        lngLen = Len(pstrName) + 1
        pstrRank = Trim$(pstrRank)
        Do
            lngLen = lngLen - 1
            strCheck = Left$(pstrName, lngLen) & "..." & pstrRank
        Loop Until Me.picGuide.TextWidth(strCheck) <= lngWidth
    End If
    FitDisplay = strCheck
End Function

Private Sub DrawBorders(plngRow As Long, plngCol As Long, pblnHighlight As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    Dim blnDraw As Boolean

    If plngRow < 1 Or (plngRow > 20 And plngRow > Guide.Enhancements) Or mlngCol < 1 Or mlngCol > 6 Then Exit Sub
    If pblnHighlight And plngRow > Guide.Enhancements Then Exit Sub
    GetCoords plngRow, plngCol, lngLeft, lngTop, lngRight, lngBottom
    If pblnHighlight Then
        Me.picGuide.Line (lngLeft, lngTop)-(lngRight, lngBottom), cfg.GetColor(cgeControls, cveBorderHighlight), B
    Else
        Me.picGuide.Line (lngLeft, lngTop)-(lngRight, lngBottom), cfg.GetColor(cgeControls, cveBorderInterior), B
        lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
        ' Left
        If mlngCol = 1 Then Me.picGuide.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), lngColor
        ' Top
        If plngRow = 1 Or plngRow = Guide.Enhancements + 1 Then
            blnDraw = True
        ElseIf plngRow > Guide.Enhancements + 1 Then
            blnDraw = False
        ElseIf Guide.Enhancement(plngRow).Level <> Guide.Enhancement(plngRow - 1).Level Then
            blnDraw = True
        Else
            blnDraw = False
        End If
        If blnDraw Then Me.picGuide.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngColor
        ' Right
        If mlngCol = 6 Then Me.picGuide.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), lngColor
        ' Bottom
        If plngRow = Guide.Enhancements Then
            blnDraw = True
        ElseIf plngRow > Guide.Enhancements Then
            blnDraw = False
        ElseIf Guide.Enhancement(plngRow).Level <> Guide.Enhancement(plngRow + 1).Level Then
            blnDraw = True
        Else
            blnDraw = False
        End If
        If blnDraw Then Me.picGuide.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngColor
    End If
End Sub


' ************* GUIDE (SCROLLBAR) *************


Private Sub scrollVertical_GotFocus()
    Me.picGuide.SetFocus
End Sub

Private Sub scrollVertical_Change()
    GuideScroll
End Sub

Private Sub scrollVertical_Scroll()
    GuideScroll
End Sub

Private Sub GuideScroll()
    If mblnOverride Then Exit Sub
    Me.picGuide.Top = 0 - (Me.scrollVertical.Value * mlngHeight)
End Sub

Private Sub GuideWheelScroll(plngValue As Long)
    Dim lngValue As Long
    
    If Not Me.scrollVertical.Visible Then Exit Sub
    With Me.scrollVertical
        lngValue = .Value - plngValue
        If lngValue < .Min Then lngValue = .Min
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub

Private Sub picGuide_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    ' placeholder
End Sub

Private Sub picGuide_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    ' placeholder
End Sub

Private Sub picGuide_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' placeholder
End Sub

Private Sub picGuide_OLECompleteDrag(Effect As Long)
    ' placeholder
End Sub

Private Sub tmrScroll_Timer()
    Dim lngNewValue As Long
    
    Me.tmrInterval.Enabled = False
    Me.tmrScroll.Enabled = False
    With Me.scrollVertical
        lngNewValue = .Value + mlngDirection
        If lngNewValue >= .Min And lngNewValue <= .Max Then .Value = lngNewValue
    End With
    Me.tmrInterval.Enabled = True
End Sub

Private Sub tmrInterval_Timer()
    Me.tmrInterval.Enabled = False
    Me.tmrScroll.Interval = 100
    mlngInterval = 100
End Sub

