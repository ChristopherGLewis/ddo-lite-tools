VERSION 5.00
Begin VB.Form frmSkills 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Skills"
   ClientHeight    =   7755
   ClientLeft      =   30
   ClientTop       =   3930
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
   Icon            =   "frmSkills.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   12225
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7380
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "< Stats"
      CenterLink      =   "Clear Skills"
      RightLinks      =   "Feats >"
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
      LeftLinks       =   "Skill Ranks;Skill Tomes"
      RightLinks      =   "Contrast;Help"
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   9120
      Top             =   7440
   End
   Begin VB.Timer tmrRepeat 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9480
      Top             =   7440
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6804
      Left            =   180
      ScaleHeight     =   6810
      ScaleWidth      =   11895
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   11892
   End
   Begin VB.PictureBox picTomes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6804
      Left            =   180
      ScaleHeight     =   6810
      ScaleWidth      =   11895
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   11892
      Begin VB.Timer tmrFocus 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3000
         Top             =   6240
      End
      Begin VB.Frame fraSkills 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1572
         Index           =   5
         Left            =   2160
         TabIndex        =   21
         Top             =   4440
         Width           =   3372
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   6
            Left            =   2100
            TabIndex        =   28
            Top             =   840
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   11
            Left            =   2100
            TabIndex        =   30
            Top             =   1080
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   14
            Left            =   2100
            TabIndex        =   26
            Top             =   600
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   21
            Left            =   2100
            TabIndex        =   24
            Top             =   360
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Miscellaneous"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   0
            Width           =   1212
         End
         Begin VB.Shape linFrame 
            Height          =   1452
            Index           =   2
            Left            =   0
            Top             =   120
            Width           =   3372
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Haggle"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   6
            Left            =   1332
            TabIndex        =   27
            Top             =   852
            Width           =   600
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Listen"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   11
            Left            =   1416
            TabIndex        =   29
            Top             =   1092
            Width           =   516
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Perform"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   14
            Left            =   1212
            TabIndex        =   25
            Top             =   612
            Width           =   720
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Use Magic Device"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   21
            Left            =   384
            TabIndex        =   23
            Top             =   372
            Width           =   1548
         End
      End
      Begin VB.Frame fraSkills 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1572
         Index           =   4
         Left            =   5820
         TabIndex        =   51
         Top             =   4440
         Width           =   3372
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   12
            Left            =   2100
            TabIndex        =   56
            Top             =   600
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   8
            Left            =   2100
            TabIndex        =   54
            Top             =   360
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Stealth"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   5
            Left            =   120
            TabIndex        =   52
            Top             =   0
            Width           =   624
         End
         Begin VB.Shape linFrame 
            Height          =   1452
            Index           =   5
            Left            =   0
            Top             =   120
            Width           =   3372
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Hide"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   8
            Left            =   1548
            TabIndex        =   53
            Top             =   372
            Width           =   384
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Move Silently"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   12
            Left            =   756
            TabIndex        =   55
            Top             =   612
            Width           =   1176
         End
      End
      Begin VB.Frame fraSkills 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1572
         Index           =   3
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   3372
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   1
            Left            =   2100
            TabIndex        =   6
            Top             =   360
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   10
            Left            =   2100
            TabIndex        =   8
            Top             =   600
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   19
            Left            =   2100
            TabIndex        =   10
            Top             =   840
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   20
            Left            =   2100
            TabIndex        =   12
            Top             =   1080
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Athletics"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   744
         End
         Begin VB.Shape linFrame 
            Height          =   1452
            Index           =   0
            Left            =   0
            Top             =   120
            Width           =   3372
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Balance"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   1248
            TabIndex        =   5
            Top             =   372
            Width           =   684
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Jump"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   10
            Left            =   1452
            TabIndex        =   7
            Top             =   612
            Width           =   480
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Swim"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   19
            Left            =   1440
            TabIndex        =   9
            Top             =   852
            Width           =   492
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Tumble"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   20
            Left            =   1284
            TabIndex        =   11
            Top             =   1092
            Width           =   648
         End
      End
      Begin VB.Frame fraSkills 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1572
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   2460
         Width           =   3372
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   2
            Left            =   2100
            TabIndex        =   16
            Top             =   360
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   4
            Left            =   2100
            TabIndex        =   18
            Top             =   600
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   9
            Left            =   2100
            TabIndex        =   20
            Top             =   840
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Social"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   0
            Width           =   516
         End
         Begin VB.Shape linFrame 
            Height          =   1452
            Index           =   1
            Left            =   0
            Top             =   120
            Width           =   3372
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Bluff"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   1548
            TabIndex        =   15
            Top             =   372
            Width           =   384
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Diplomacy"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   1008
            TabIndex        =   17
            Top             =   612
            Width           =   924
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Intimidate"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   9
            Left            =   1044
            TabIndex        =   19
            Top             =   852
            Width           =   888
         End
      End
      Begin VB.Frame fraSkills 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1572
         Index           =   1
         Left            =   5820
         TabIndex        =   41
         Top             =   2460
         Width           =   3372
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   5
            Left            =   2100
            TabIndex        =   48
            Top             =   840
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   13
            Left            =   2100
            TabIndex        =   50
            Top             =   1080
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   16
            Left            =   2100
            TabIndex        =   44
            Top             =   360
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   18
            Left            =   2100
            TabIndex        =   46
            Top             =   600
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Trapping"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   120
            TabIndex        =   42
            Top             =   0
            Width           =   756
         End
         Begin VB.Shape linFrame 
            Height          =   1452
            Index           =   4
            Left            =   0
            Top             =   120
            Width           =   3372
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Disable Device"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   5
            Left            =   660
            TabIndex        =   47
            Top             =   852
            Width           =   1272
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Open Lock"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   13
            Left            =   984
            TabIndex        =   49
            Top             =   1092
            Width           =   948
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Search"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   16
            Left            =   1320
            TabIndex        =   43
            Top             =   372
            Width           =   612
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Spot"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   18
            Left            =   1512
            TabIndex        =   45
            Top             =   612
            Width           =   420
         End
      End
      Begin VB.Frame fraSkills 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1572
         Index           =   0
         Left            =   5820
         TabIndex        =   31
         Top             =   480
         Width           =   3372
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   3
            Left            =   2100
            TabIndex        =   34
            Top             =   360
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   7
            Left            =   2100
            TabIndex        =   38
            Top             =   840
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   15
            Left            =   2100
            TabIndex        =   40
            Top             =   1080
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin CharacterBuilderLite.userSpinner usrspnSkill 
            Height          =   252
            Index           =   17
            Left            =   2100
            TabIndex        =   36
            Top             =   600
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   450
            Min             =   0
            Max             =   5
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
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Spellcasting"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   1032
         End
         Begin VB.Shape linFrame 
            Height          =   1452
            Index           =   3
            Left            =   0
            Top             =   120
            Width           =   3372
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Concentration"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   672
            TabIndex        =   33
            Top             =   372
            Width           =   1260
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Heal"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   7
            Left            =   1548
            TabIndex        =   37
            Top             =   852
            Width           =   384
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Repair"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   15
            Left            =   1380
            TabIndex        =   39
            Top             =   1092
            Width           =   552
         End
         Begin VB.Label lblSkill 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Spellcraft"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   17
            Left            =   1116
            TabIndex        =   35
            Top             =   612
            Width           =   816
         End
      End
      Begin VB.Label lblTimer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "tmrFocus: To force user controls to lose focus to MsgBox"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3480
         TabIndex        =   57
         Top             =   6300
         Visible         =   0   'False
         Width           =   6312
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuSkill 
         Caption         =   "Max Ranks (Half Ranks)"
         Index           =   0
      End
      Begin VB.Menu mnuSkill 
         Caption         =   "Max Ranks (Full Ranks)"
         Index           =   1
      End
      Begin VB.Menu mnuSkill 
         Caption         =   "Max Ranks (Even Levels)"
         Index           =   2
      End
      Begin VB.Menu mnuSkill 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSkill 
         Caption         =   "Clear Ranks"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Contrast"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuContrast 
         Caption         =   "None"
         Index           =   0
      End
      Begin VB.Menu mnuContrast 
         Caption         =   "Low"
         Index           =   1
      End
      Begin VB.Menu mnuContrast 
         Caption         =   "Medium"
         Index           =   2
      End
      Begin VB.Menu mnuContrast 
         Caption         =   "High"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private mlngMaxRanks As Long
Private mlngMaxRanksRow As Long

Private mlngLeft As Long
Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
Private mlngIconWidth As Long
Private mlngIconHeight As Long
Private mlngIconOffset As Long

Private mlngNativeColor As Long
Private mlngCrossColor As Long

Private mlngRow As Long ' mlngRow is the grid row; data row is mtypRow(mlngRow).Skill
Private mlngCol As Long

Private mblnClassRow As Boolean

Private mintButton As Integer
Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnOverride = False
    cfg.MoveForm Me
    InitBuildSkills
    Cascade
End Sub

Private Sub Form_Activate()
    ActivateForm oeSkills
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me, mblnOverride
End Sub

Public Sub Cascade()
    LoadData
    InitRows
    InitGrid
    cfg.RefreshColors Me
End Sub

Public Sub RefreshIcons()
    InitRows
    InitGrid
End Sub

Private Sub SkillsChanged()
    CascadeChanges cceSkill
    SetDirty
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Skill Ranks"
            Me.picTomes.Visible = False
            Me.usrFooter.CenterLink = "Clear Skills"
            Me.picGrid.Visible = True
        Case "Skill Tomes"
            Me.picGrid.Visible = False
            Me.usrFooter.CenterLink = "Clear Tomes"
            Me.picTomes.Visible = True
        Case "Contrast"
            ContrastMenu
        Case "Help"
            ShowHelp "Skills"
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "< Stats"
            cfg.SavePosition Me
            If Not OpenForm("frmStats") Then Exit Sub
        Case "Feats >"
            cfg.SavePosition Me
            If Not OpenForm("frmFeats") Then Exit Sub
        Case "Clear Skills"
            If Ask("Clear all skills?") Then
                Erase build.Skills
                InitBuildSkills
                InitGrid
                SkillsChanged
            End If
            Exit Sub
        Case "Clear Tomes"
            Me.picTomes.SetFocus
            Me.tmrFocus.Enabled = True
            Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub


Private Sub tmrFocus_Timer()
    Me.tmrFocus.Enabled = False
    If Not Ask("Clear all skill tomes?") Then Exit Sub
    Erase build.SkillTome
    LoadData
    SkillsChanged
End Sub

Private Sub ContrastMenu()
    Dim i As Long
    
    For i = 0 To 3
        Me.mnuContrast(i).Checked = (cfg.Contrast = i)
    Next
    PopupMenu Me.mnuMain(1)
End Sub

Private Sub mnuContrast_Click(Index As Integer)
    cfg.Contrast = Index
    SetFillColors
    DrawGrid
End Sub


' ************* INITIALIZE *************


Private Sub InitRows()
    Dim typClassSplit() As ClassSplitType
    Dim lngClasses As Long
    Dim lngClass As Long
    Dim blnSet(21) As Boolean
    Dim lngRow As Long
    Dim i As Long
    
    ' Set font to bold for sizing purposes
    Me.picGrid.FontBold = True
    ' Class row?
    lngClasses = GetClassSplit(typClassSplit)
    If lngClasses = 0 Then Exit Sub
    mblnClassRow = (lngClasses > 1)
    Select Case cfg.SkillOrderScreen
        Case sosAlphabetical
            For i = 1 To 21
                SetRow lngRow, i
            Next
        Case sosNativeFirst
            ' Primary class native skills first...
            For i = 1 To 21
                If db.Class(typClassSplit(0).ClassID).NativeSkills(i) Then blnSet(i) = SetRow(lngRow, i)
            Next
            ' All splash class native skills second, mixed together
            For i = 1 To 21
                For lngClass = 1 To lngClasses - 1
                    If db.Class(typClassSplit(lngClass).ClassID).NativeSkills(i) And Not blnSet(i) Then blnSet(i) = SetRow(lngRow, i)
                Next
            Next
            ' Cross-class skills at end of list
            For lngClass = 1 To lngClasses
                For i = 1 To 21
                    If Not blnSet(i) Then blnSet(i) = SetRow(lngRow, i)
                Next
            Next
    End Select
End Sub

Private Function SetRow(plngRow As Long, ByVal penSkill As SkillsEnum) As Boolean
    plngRow = plngRow + 1
    With Skill.Map(plngRow)
        .Skill = penSkill
        .SkillName = GetSkillName(penSkill)
    End With
    SetRow = True
End Function

Private Sub InitGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngLines As Long
    Dim lngFudge As Long ' True centering looks uneven because the total ranks colum on the right won't normally have ½ ranks
    
    With Me.picGrid
        .FontBold = True
        mlngLeft = .TextWidth("Use Magic Device  ")
        mlngWidth = .TextWidth("10½")
        .FontBold = False
        lngFudge = .TextWidth("½") \ 4 ' Came up with this number via trial and error
        mlngIconWidth = mlngWidth * 4 / 5
        mlngIconWidth = .ScaleX(.ScaleX(Int(mlngIconWidth), vbTwips, vbPixels), vbPixels, vbTwips)
        mlngIconOffset = (mlngWidth - mlngIconWidth) \ 2
        mlngIconHeight = .ScaleY(.ScaleX(mlngIconWidth, vbTwips, vbPixels), vbPixels, vbTwips)
        lngWidth = mlngLeft + mlngWidth * 21 '+ PixelX
        If mblnClassRow Then
            If cfg.UseIcons And cfg.IconSkills Then
                mlngHeight = (Me.usrFooter.Top - Me.usrHeader.Height - mlngIconHeight) \ 25
            Else
                mlngHeight = (Me.usrFooter.Top - Me.usrHeader.Height) \ 26
            End If
        Else
            mlngHeight = (Me.usrFooter.Top - Me.usrHeader.Height) \ 25
        End If
        mlngHeight = .ScaleY(Int(.ScaleY(mlngHeight, vbTwips, vbPixels)), vbPixels, vbTwips)
        ' Resize and center grid
        If mblnClassRow Then
            If cfg.UseIcons And cfg.IconSkills Then
                mlngTop = mlngIconHeight + mlngHeight
                lngHeight = mlngIconHeight + mlngHeight * 24
            Else
                mlngTop = mlngHeight * 2
                lngHeight = mlngHeight * 25
            End If
        Else
            mlngTop = mlngHeight
            lngHeight = mlngHeight * 24
        End If
        lngLeft = (Me.ScaleWidth - lngWidth) \ 2 + lngFudge
    End With
    lngTop = Me.usrHeader.Height + ((Me.usrFooter.Top - Me.usrHeader.Height) - lngHeight) \ 2
    Me.picGrid.Move lngLeft, lngTop, lngWidth, lngHeight
    ' Abbreviate skill names if they don't fit in the space we have to work with
    Me.picGrid.FontBold = True
    For lngRow = 1 To 21
        With Skill.Map(lngRow)
            If Me.picGrid.TextWidth(.SkillName) > mlngLeft Then .SkillName = GetSkillName(.Skill, True)
        End With
    Next
    Me.picGrid.FontBold = False
    DrawGrid
End Sub


' ************* SKILL TOMES *************


Private Sub LoadData()
    Dim lngValue As Long
    Dim i As Long
    
    mblnOverride = True
    For i = 1 To 21
        Me.usrspnSkill(i).Max = tomes.Skill.Max
        lngValue = build.SkillTome(i)
        If lngValue > tomes.Skill.Max Then lngValue = tomes.Skill.Max
        Me.usrspnSkill(i).Value = lngValue
    Next
    mblnOverride = False
End Sub

Private Sub usrspnSkill_Change(Index As Integer)
    If mblnOverride Then Exit Sub
    build.SkillTome(Index) = Me.usrspnSkill(Index).Value
    SkillsChanged
End Sub


' ************* DRAW GRID *************


Public Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    SetFillColors
    Me.picGrid.Cls
    For lngCol = 1 To HeroicLevels()
        DrawColumnHeader lngCol, False
    Next
    For lngCol = 1 To HeroicLevels()
        DrawColumnFooter lngCol
    Next
    ' Draw skill names
    For lngRow = 1 To 21
        DrawRowHeader lngRow, False
        DrawRowFooter lngRow
    Next
    ' Draw cells
    For lngCol = 1 To HeroicLevels()
        For lngRow = 1 To 21
            DrawCell lngRow, lngCol
        Next
    Next
End Sub

Private Sub SetFillColors()
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim dblContrast As Double
    
    mlngCrossColor = cfg.GetColor(cgeControls, cveBackground)
    mlngNativeColor = mlngCrossColor
    If cfg.Contrast = ceNone Then Exit Sub
    xp.ColorToRGB mlngNativeColor, lngRed, lngGreen, lngBlue
    Select Case 256 - ((lngRed + lngGreen + lngBlue) \ 3)
        Case Is <= 15: dblContrast = 0.975
        Case 16 To 25: dblContrast = 1.03
        Case 26 To 156: dblContrast = 1.05
        Case 157 To 181: dblContrast = 1.1
        Case 182 To 191: dblContrast = 1.2
        Case Is >= 192: dblContrast = 1.25
    End Select
    Select Case cfg.Contrast
        Case ceMedium: dblContrast = 1 + (dblContrast - 1) * 1.5
        Case ceHigh: dblContrast = 1 + (dblContrast - 1) * 2
    End Select
'    Me.Caption = "From " & (lngRed + lngGreen + lngBlue) \ 3
    lngRed = lngRed * dblContrast
    lngGreen = lngGreen * dblContrast
    lngBlue = lngBlue * dblContrast
'    Me.Caption = Me.Caption & " to " & (lngRed + lngGreen + lngBlue) \ 3
    If dblContrast > 1 Then
        mlngNativeColor = RGB(lngRed, lngGreen, lngBlue)
    Else
        mlngNativeColor = mlngCrossColor
        mlngCrossColor = RGB(lngRed, lngGreen, lngBlue)
    End If
End Sub

Private Sub DrawRowHeader(plngRow As Long, pblnBold As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    
    If plngRow < 1 Or plngRow > 21 Then Exit Sub
    lngLeft = 0
    lngTop = mlngTop + mlngHeight * (plngRow - 1)
    lngRight = mlngLeft - PixelX
    lngBottom = lngTop + mlngHeight - PixelY
    With Me.picGrid
        .ForeColor = cfg.GetColor(cgeWorkspace, cveText)
        If .FontBold <> pblnBold Then .FontBold = pblnBold
        Me.picGrid.Line (lngLeft, lngTop)-(lngRight, lngBottom), cfg.GetColor(cgeWorkspace, cveBackground), BF
        Me.picGrid.CurrentX = 0
        Me.picGrid.CurrentY = lngTop
        Me.picGrid.Print Skill.Map(plngRow).SkillName
        If pblnBold Then Me.picGrid.FontBold = False
    End With
End Sub

Private Sub DrawRowFooter(plngRow As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngMin As Long
    Dim lngMax As Long
    Dim lngColor As Long
    Dim lngSkill As Long
    
    lngSkill = Skill.Map(plngRow).Skill
    If Me.picGrid.FontBold Then Me.picGrid.FontBold = False
    ' Get coordinates
    lngLeft = mlngLeft + mlngWidth * 20
    lngTop = mlngTop + mlngHeight * (plngRow - 1)
    ' Clear
    Me.picGrid.Line (lngLeft + PixelX, lngTop)-(lngLeft + mlngWidth, lngTop + mlngHeight), cfg.GetColor(cgeWorkspace, cveBackground), BF
    If Skill.Row(lngSkill).Ranks = 0 Then Exit Sub
    ' Color
    lngMax = Skill.Row(lngSkill).MaxRanks
    If lngMax Mod 2 <> 0 Then lngMin = lngMax - 1 Else lngMin = lngMax
    Select Case Skill.Row(lngSkill).Ranks
        Case lngMin To lngMax: lngColor = cfg.GetColor(cgeWorkspace, cveText)
        Case Is > lngMax: lngColor = cfg.GetColor(cgeWorkspace, cveTextError)
        Case Else: lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
    End Select
    ' Draw text
    PrintText FormatRanks(Skill.Row(lngSkill).Ranks), lngLeft, lngTop, lngColor
End Sub

Private Sub DrawColumnHeader(plngCol As Long, pblnBold As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim strDisplay As String
'    Dim strFile As String
    Dim strResource As String
    
    If plngCol < 1 Or plngCol > HeroicLevels() Then Exit Sub
    ' Clear
    lngLeft = mlngLeft + mlngWidth * (plngCol - 1) + PixelX
    lngRight = lngLeft + mlngWidth - PixelX * 2
    Me.picGrid.Line (lngLeft, lngTop)-(lngRight, mlngTop - PixelY), cfg.GetColor(cgeWorkspace, cveBackground), BF
    ' Draw
    With Me.picGrid
        .ForeColor = cfg.GetColor(cgeWorkspace, cveText)
        If .FontBold <> pblnBold Then .FontBold = pblnBold
        ' Class header
        If mblnClassRow Then
            If cfg.UseIcons And cfg.IconSkills Then
                strResource = GetClassResourceID(build.Class(plngCol))
                .PaintPicture LoadResPicture(strResource, vbResBitmap), lngLeft + mlngIconOffset, 0, mlngIconWidth, mlngIconHeight
                lngTop = mlngIconHeight
            Else
                strDisplay = Skill.Col(plngCol).Initial
                .CurrentX = lngLeft + (mlngWidth - .TextWidth(strDisplay)) \ 2
                .CurrentY = lngTop
                Me.picGrid.Print strDisplay
                lngTop = mlngHeight
            End If
        End If
        ' Level
        strDisplay = CStr(plngCol)
        .CurrentX = lngLeft + (mlngWidth - .TextWidth(strDisplay)) \ 2
        .CurrentY = lngTop
        Me.picGrid.Print strDisplay
        If pblnBold Then Me.picGrid.FontBold = False
    End With
End Sub

Private Sub DrawColumnFooter(plngCol As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngColor As Long
    
    If plngCol < 1 Or plngCol > HeroicLevels() Then Exit Sub
    lngLeft = mlngLeft + mlngWidth * (plngCol - 1)
    lngTop = mlngTop + mlngHeight * 21
    Me.picGrid.Line (lngLeft, lngTop + PixelY)-(lngLeft + mlngWidth, lngTop + mlngHeight * 2), cfg.GetColor(cgeWorkspace, cveBackground), BF
    With Skill.Col(plngCol)
        Select Case .Points
            Case Is < .MaxPoints: lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
            Case Is > .MaxPoints: lngColor = cfg.GetColor(cgeWorkspace, cveTextError)
            Case Else: lngColor = cfg.GetColor(cgeWorkspace, cveText)
        End Select
        PrintText CStr(.Points), lngLeft, lngTop, lngColor
        PrintText CStr(.MaxPoints), lngLeft, lngTop + mlngHeight, cfg.GetColor(cgeWorkspace, cveText)
    End With
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, Optional pblnActive As Boolean = False)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngColor As Long
    Dim lngRanks As Long
    Dim lngMaxRanks As Long
    
    If plngRow < 1 Or plngRow > 21 Then Exit Sub
    If plngCol < 1 Or plngCol > HeroicLevels() Then Exit Sub
    If Me.picGrid.FontBold Then Me.picGrid.FontBold = False
    ' Get coordinates
    lngLeft = mlngLeft + mlngWidth * (plngCol - 1)
    lngTop = mlngTop + mlngHeight * (plngRow - 1)
    ' Background
    If Skill.grid(Skill.Map(plngRow).Skill, plngCol).Native = 2 Then lngColor = mlngNativeColor Else lngColor = mlngCrossColor
    If Me.picGrid.FillColor <> lngColor Then Me.picGrid.FillColor = lngColor
    If pblnActive Then lngColor = cfg.GetColor(cgeControls, cveBorderHighlight) Else lngColor = cfg.GetColor(cgeControls, cveBorderInterior)
    Me.picGrid.Line (lngLeft, lngTop)-(lngLeft + mlngWidth, lngTop + mlngHeight), lngColor, B
    ' Borders
    If Not pblnActive Then
        lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
        If plngRow = 1 Then Me.picGrid.Line (lngLeft, lngTop)-(lngLeft + mlngWidth + PixelX, lngTop), lngColor
        If plngRow = 21 Then Me.picGrid.Line (lngLeft, lngTop + mlngHeight)-(lngLeft + mlngWidth + PixelX, lngTop + mlngHeight), lngColor
        If plngCol = 1 Then Me.picGrid.Line (lngLeft, lngTop)-(lngLeft, lngTop + mlngHeight + PixelY), lngColor
        If plngCol = HeroicLevels() Then Me.picGrid.Line (lngLeft + mlngWidth, lngTop)-(lngLeft + mlngWidth, lngTop + mlngHeight + PixelY), lngColor
    End If
    ' Ranks
    lngRanks = GetRanks(plngRow, plngCol)
    If lngRanks = 0 Then Exit Sub
    ' Color
    lngMaxRanks = Skill.grid(Skill.Map(plngRow).Skill, plngCol).MaxRanks
    Select Case CumulativeRanks(plngRow, plngCol)
        Case Is > lngMaxRanks: lngColor = cfg.GetColor(cgeControls, cveTextError)
        Case Is < lngMaxRanks: lngColor = cfg.GetColor(cgeControls, cveTextDim)
        Case Else: lngColor = cfg.GetColor(cgeControls, cveText)
    End Select
    ' Draw text
    PrintText FormatRanks(lngRanks), lngLeft, lngTop, lngColor
End Sub

Private Function GetRanks(plngRow As Long, plngCol As Long)
    GetRanks = Skill.grid(Skill.Map(plngRow).Skill, plngCol).Ranks
End Function

Private Function FormatRanks(plngRanks As Long) As String
    Select Case plngRanks
        Case Is < 1
        Case 1: FormatRanks = "½"
        Case Else: If plngRanks Mod 2 <> 0 Then FormatRanks = CStr(plngRanks \ 2) & "½" Else FormatRanks = CStr(plngRanks \ 2)
    End Select
End Function

Private Function CumulativeRanks(plngRow As Long, plngCol As Long) As Long
    Dim lngRanks As Long
    Dim lngLast As Long
    Dim i As Long
    
    If plngCol > HeroicLevels() Then lngLast = HeroicLevels() Else lngLast = plngCol
    For i = 1 To lngLast
        lngRanks = lngRanks + GetRanks(plngRow, i)
    Next
    CumulativeRanks = lngRanks
End Function

Private Sub PrintText(pstrText As String, plngLeft As Long, plngTop As Long, plngColor As Long)
    If Len(pstrText) = 0 Then Exit Sub
    With Me.picGrid
        .CurrentX = plngLeft + (mlngWidth - .TextWidth(pstrText)) \ 2
        .CurrentY = plngTop + (mlngHeight - .TextHeight(pstrText)) \ 2
        .ForeColor = plngColor
    End With
    Me.picGrid.Print pstrText
End Sub


' ************* GRID *************


Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then ActiveCell X, Y
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
    If mlngRow < 0 Or mlngRow > 21 Then Exit Sub
    If mlngCol < 0 Or mlngCol > 20 Then Exit Sub
    If mlngRow = 0 Then
        If mlngCol = 0 Then Exit Sub
        If Button = vbRightButton Then ClearColumn
    Else
        mintButton = Button
        Select Case mintButton
            Case vbLeftButton
                If mlngCol = 0 Then
                    If mlngMaxRanksRow = mlngRow Then
                        MaxRanks
                    Else
                        mlngMaxRanksRow = mlngRow
                        MaxRanks 1
                    End If
                Else
                    Increment
                    mlngMaxRanksRow = 0
                End If
            Case vbMiddleButton
                If mlngCol > 0 Then ClearCell
                mlngMaxRanksRow = 0
            Case vbRightButton
                If mlngCol = 0 Then
                    mlngMaxRanksRow = mlngRow
                    PopupMenu Me.mnuMain(0)
                Else
                    Decrement
                End If
                mlngMaxRanksRow = 0
        End Select
        If mlngCol <> 0 Then Me.tmrStart.Enabled = True
    End If
End Sub

Private Sub picGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.tmrStart.Enabled = False
    Me.tmrRepeat.Enabled = False
    If Button = 0 Then ActiveCell X, Y
End Sub

Private Sub mnuSkill_Click(Index As Integer)
    If Index = 4 Then SkillClear Else MaxRanks Index + 1
End Sub

Private Sub picGrid_DblClick()
    MouseDown
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowActiveCell -1, -1
End Sub

Private Sub tmrStart_Timer()
    Me.tmrStart.Enabled = False
    Me.tmrRepeat.Enabled = True
End Sub

Private Sub tmrRepeat_Timer()
    MouseDown
End Sub

Private Sub ActiveCell(ByVal lngX As Long, ByVal lngY As Long)
    Dim lngRow As Long
    Dim lngCol As Long

    ' Identify skill
    If lngY < 0 Then
        lngRow = -1
    ElseIf lngY < mlngTop Then
        lngRow = 0
    Else
        lngY = lngY - mlngTop
        lngRow = lngY \ mlngHeight + 1
        If lngRow > 21 Then lngRow = -1
    End If
    ' Identify level
    If lngX < 0 Then
        lngCol = -1
    ElseIf lngX < mlngLeft Then
        lngCol = 0
    Else
        lngX = lngX - mlngLeft
        lngCol = lngX \ mlngWidth + 1
        If lngCol > 20 Then lngCol = -1
    End If
    ShowActiveCell lngRow, lngCol
End Sub

Private Sub ShowActiveCell(plngRow As Long, plngCol As Long)
'    If mlngRow > 0 And mlngCol > 0 Then
'        If Skill.grid(Skill.Map(mlngRow).Skill, mlngCol).Native = 2 Then xp.SetMouseCursor mcHand
'    End If
    If mlngRow <> plngRow Or mlngCol <> plngCol Then
        HighlightCell mlngRow, mlngCol, False
        HighlightCell plngRow, plngCol, True
    End If
    If mlngRow <> plngRow Then
        DrawRowHeader mlngRow, False
        mlngRow = plngRow
        DrawRowHeader mlngRow, True
    End If
    If mlngCol <> plngCol Then
        DrawColumnHeader mlngCol, False
        mlngCol = plngCol
        DrawColumnHeader mlngCol, True
    End If
End Sub

Private Sub HighlightCell(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Select Case plngRow
        Case 1 To 21
        Case Else: Exit Sub
    End Select
    Select Case plngCol
        Case 1 To 20
        Case Else: Exit Sub
    End Select
    DrawCell plngRow, plngCol, pblnActive
End Sub

Private Sub MouseDown()
    If mlngRow < 1 Or mlngRow > 21 Then Exit Sub
    If mlngCol < 0 Or mlngCol > 20 Then Exit Sub
    If mlngCol = 0 Then
        MaxRanks
    Else
        Select Case mintButton
            Case vbLeftButton: Increment
            Case vbRightButton: Decrement
        End Select
    End If
End Sub

Private Sub Increment()
    Dim lngCumulative As Long
    Dim lngRanks As Long
    Dim i As Long
    Dim anyTrapper As Boolean
    
    anyTrapper = False
    ' Is this a Trapper skill, and have we taken a trapper level?
    If Skill.Map(mlngRow).Skill = seDisableDevice Or Skill.Map(mlngRow).Skill = seOpenLock Then
        For i = 1 To mlngCol
            'Have we taken a trapper level yet
            If Skill.Col(i).Trapper Then
                anyTrapper = True
            End If
        Next
        If Not anyTrapper Then
            Exit Sub
        End If
    End If
    ' Do we have points left to spend this level?
    If Skill.Col(mlngCol).Points >= Skill.Col(mlngCol).MaxPoints Then Exit Sub
    ' Add point if we can fit it into the Level+3 cap
    With Skill.grid(Skill.Map(mlngRow).Skill, mlngCol)
        If CumulativeRanks(mlngRow, mlngCol) + .Native <= .MaxRanks Then
            AddPoint 1
        End If
    End With
End Sub

Private Sub Decrement()
    If build.Skills(Skill.Map(mlngRow).Skill, mlngCol) > 0 Then AddPoint -1
End Sub

Private Sub AddPoint(ByVal plngPoints As Long)
    Dim lngNew As Long
    Dim lngSkill As Long
    Dim i As Long
    
    lngSkill = Skill.Map(mlngRow).Skill
    lngNew = build.Skills(lngSkill, mlngCol) + plngPoints
    If lngNew < 0 Then lngNew = 0
    build.Skills(lngSkill, mlngCol) = lngNew
    CalculateCell mlngRow, mlngCol
    ' Redraw cells from this level through total (may violate MaxRanks in subsequent levels)
    For i = mlngCol + 1 To 20
        DrawCell mlngRow, i
    Next
    ' Draw current cell last because cell borders overlap (current cell will show active border)
    DrawCell mlngRow, mlngCol, True
    ' Refresh totals
    RecalculateRow mlngRow
    RecalculateColumn mlngCol
    SkillsChanged
End Sub

Private Sub CalculateCell(plngRow As Long, plngCol As Long)
    Dim lngSkill As Long
    
    lngSkill = Skill.Map(plngRow).Skill
    With Skill.grid(lngSkill, plngCol)
        .Ranks = build.Skills(lngSkill, plngCol) * .Native
    End With
End Sub

Private Sub ClearCell()
    Dim lngSkill As Long
    Dim i As Long
    
    lngSkill = Skill.Map(mlngRow).Skill
    build.Skills(lngSkill, mlngCol) = 0
    Skill.grid(lngSkill, mlngCol).Ranks = 0
    ' Redraw cells from this level through total (may no longer violate MaxRanks in subsequent levels)
    For i = mlngCol + 1 To 20
        DrawCell mlngRow, i
    Next
    ' Draw current cell last because cell borders overlap (current cell will show active border)
    DrawCell mlngRow, mlngCol, True
    ' Refresh totals
    RecalculateRow mlngRow
    RecalculateColumn mlngCol
    SkillsChanged
End Sub

Private Sub ClearColumn()
    Dim lngSkill As Long
    Dim lngRow As Long
    
    For lngRow = 1 To 21
        lngSkill = Skill.Map(lngRow).Skill
        build.Skills(lngSkill, mlngCol) = 0
        Skill.grid(lngSkill, mlngCol).Ranks = 0
        DrawCell lngRow, mlngCol
        RecalculateRow lngRow
    Next
    RecalculateColumn mlngCol
    SkillsChanged
End Sub

Private Sub MaxRanks(Optional plngMaxRanks As Long = -1)
    If plngMaxRanks = -1 Then
        mlngMaxRanks = mlngMaxRanks + 1
    Else
        mlngMaxRanks = plngMaxRanks
    End If
    If mlngMaxRanks > 4 Then mlngMaxRanks = 1
    Select Case mlngMaxRanks
        Case 1: SkillMaxHalfRanks
        Case 2: SkillMaxFullRanks
        Case 3: SkillMaxEvenLevels
        Case 4: SkillClear
    End Select
End Sub

Private Sub SkillClear()
    Dim lngSkill As Long
    Dim lngCol As Long
   
    lngSkill = Skill.Map(mlngRow).Skill
    For lngCol = 1 To 20
        build.Skills(lngSkill, lngCol) = 0
        Skill.grid(lngSkill, lngCol).Ranks = 0
        DrawCell mlngRow, lngCol
        RecalculateColumn lngCol
    Next
    RecalculateRow mlngRow
    SkillsChanged
End Sub

Private Sub SkillMaxHalfRanks()
    Dim lngCumulative As Long
    Dim lngCol As Long
    Dim lngSkill As Long
    
    lngSkill = Skill.Map(mlngRow).Skill
    For lngCol = FirstLevel() To 20
        With Skill.grid(lngSkill, lngCol)
            .Ranks = .MaxRanks - lngCumulative
            If .Native = 0 Then
                .Ranks = 0
                build.Skills(lngSkill, lngCol) = 0
            Else
                .Ranks = (.Ranks \ .Native) * .Native
                build.Skills(lngSkill, lngCol) = .Ranks \ .Native
            End If
            lngCumulative = lngCumulative + .Ranks
        End With
        DrawCell mlngRow, lngCol
        RecalculateColumn lngCol
    Next
    RecalculateRow mlngRow
    SkillsChanged
End Sub

Private Sub SkillMaxFullRanks()
    Dim lngCumulative As Long
    Dim lngCol As Long
    Dim lngSkill As Long
    
    lngSkill = Skill.Map(mlngRow).Skill
    'Loop through columns from FirstLevel() to 20
    For lngCol = FirstLevel() To 20
        With Skill.grid(lngSkill, lngCol)
            .Ranks = .MaxRanks - lngCumulative
            If .Ranks Mod 2 = 1 Then
                .Ranks = .Ranks - 1
            End If
            If .Native = 0 Then
                build.Skills(lngSkill, lngCol) = 0
            Else
                'max this out to max for cell
                'If (.Ranks \ .Native) > 4 Then
                '    build.Skills(lngSkill, lngCol) = 4 '.Ranks \ .Native
                'Else
                    build.Skills(lngSkill, lngCol) = .Ranks \ .Native
                'End If
            End If
            lngCumulative = lngCumulative + .Ranks
        End With
        DrawCell mlngRow, lngCol
        RecalculateColumn lngCol
    Next
    RecalculateRow mlngRow
    SkillsChanged
End Sub

Private Sub SkillMaxEvenLevels()
    Dim lngCumulative As Long
    Dim lngCol As Long
    Dim lngSkill As Long
    
    lngSkill = Skill.Map(mlngRow).Skill
    For lngCol = FirstLevel() To 20
        With Skill.grid(lngSkill, lngCol)
            If lngCol = 1 Then
                .Ranks = .MaxRanks \ 2
            ElseIf lngCol Mod 2 = 0 Then
                .Ranks = .MaxRanks - lngCumulative
                If .Ranks Mod 2 = 1 Then .Ranks = .Ranks - 1
            Else
                .Ranks = 0
            End If
            If .Native = 0 Then build.Skills(lngSkill, lngCol) = 0 Else build.Skills(lngSkill, lngCol) = .Ranks \ .Native
            lngCumulative = lngCumulative + .Ranks
        End With
        DrawCell mlngRow, lngCol
        RecalculateColumn lngCol
    Next
    RecalculateRow mlngRow
    SkillsChanged
End Sub

''All skills should have a first level of 1 other than DD and OL
Private Function FirstLevel() As Long
    Dim enSkill As Long
    Dim lngFirst As Long
    
    enSkill = Skill.Map(mlngRow).Skill
    If enSkill = seDisableDevice Or enSkill = seOpenLock Then
        For lngFirst = 1 To 20
            If Skill.Col(lngFirst).Trapper Then Exit For
        Next
    Else
        lngFirst = 1
    End If
    FirstLevel = lngFirst
End Function

Private Sub RecalculateColumn(plngCol As Long)
    Dim lngRow As Long
    
    With Skill.Col(plngCol)
        .Points = 0
        For lngRow = 1 To 21
            .Points = .Points + build.Skills(lngRow, plngCol)
        Next
    End With
    DrawColumnFooter plngCol
End Sub

Private Sub RecalculateRow(plngRow As Long)
    Dim lngSkill As Long
    Dim lngCol As Long
    
    lngSkill = Skill.Map(plngRow).Skill
    With Skill.Row(lngSkill)
        .Ranks = 0
        For lngCol = 1 To HeroicLevels()
            .Ranks = .Ranks + Skill.grid(lngSkill, lngCol).Ranks
        Next
    End With
    DrawRowFooter plngRow
End Sub

