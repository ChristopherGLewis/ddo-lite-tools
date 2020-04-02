VERSION 5.00
Begin VB.Form frmTools 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   6876
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9132
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6876
   ScaleWidth      =   9132
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4092
      Index           =   0
      Left            =   0
      ScaleHeight     =   4092
      ScaleMode       =   0  'User
      ScaleWidth      =   9132
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "ctl"
      Top             =   360
      Width           =   9132
      Begin VB.TextBox txtSavePath 
         Height          =   324
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3480
         Width           =   7932
      End
      Begin VB.Frame fraData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   4680
         TabIndex        =   13
         Tag             =   "ctl"
         Top             =   240
         Width           =   4092
         Begin VB.TextBox txtDataFile 
            Appearance      =   0  'Flat
            Height          =   324
            Left            =   2040
            TabIndex        =   18
            Top             =   600
            Width           =   1692
         End
         Begin VB.ListBox lstDataFile 
            Appearance      =   0  'Flat
            Height          =   1752
            ItemData        =   "frmTools.frx":000C
            Left            =   240
            List            =   "frmTools.frx":000E
            TabIndex        =   15
            Tag             =   "ctl"
            Top             =   360
            Width           =   1572
         End
         Begin VB.CheckBox chkDelete 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Delete"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2160
            Width           =   1572
         End
         Begin VB.Label lnkNew 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "New"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   2040
            TabIndex        =   17
            Tag             =   "ctl"
            Top             =   360
            Width           =   396
         End
         Begin VB.Label lnkSave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Save"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   3300
            TabIndex        =   19
            Tag             =   "ctl"
            Top             =   360
            Width           =   444
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Revert Characters"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   10
            Left            =   2100
            TabIndex        =   20
            Tag             =   "ctl"
            Top             =   1260
            Width           =   1620
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Revert Notes"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   11
            Left            =   2100
            TabIndex        =   21
            Tag             =   "ctl"
            Top             =   1560
            Width           =   1176
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Revert Links"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   12
            Left            =   2100
            TabIndex        =   22
            Tag             =   "ctl"
            Top             =   1860
            Width           =   1092
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Data"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   8
            Left            =   120
            TabIndex        =   14
            Tag             =   "ctl"
            Top             =   0
            Width           =   420
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Refresh Backups"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   13
            Left            =   2100
            TabIndex        =   23
            Tag             =   "ctl"
            Top             =   2340
            Width           =   1476
         End
         Begin VB.Shape shpData 
            Height          =   2652
            Left            =   0
            Top             =   120
            Width           =   4092
         End
      End
      Begin VB.Frame fraQuickAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   360
         TabIndex        =   2
         Tag             =   "ctl"
         Top             =   240
         Width           =   4092
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "About"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   14
            Left            =   300
            TabIndex        =   53
            Tag             =   "ctl"
            Top             =   2220
            Width           =   528
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Custom Colors"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   2160
            TabIndex        =   52
            Tag             =   "ctl"
            Top             =   2220
            Width           =   1344
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Named Colors"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   5
            Left            =   2160
            TabIndex        =   12
            Tag             =   "ctl"
            Top             =   1920
            Width           =   1272
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Characters"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   300
            TabIndex        =   4
            Tag             =   "ctl"
            Top             =   420
            Width           =   972
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Challenges"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   300
            TabIndex        =   5
            Tag             =   "ctl"
            Top             =   720
            Width           =   948
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Patrons"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   300
            TabIndex        =   6
            Tag             =   "ctl"
            Top             =   1020
            Width           =   684
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Shroud Puzzles"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   6
            Left            =   2160
            TabIndex        =   8
            Tag             =   "ctl"
            Top             =   420
            Width           =   1368
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "ADQ Riddle"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   7
            Left            =   2160
            TabIndex        =   9
            Tag             =   "ctl"
            Top             =   720
            Width           =   972
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Stopwatch"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   8
            Left            =   2160
            TabIndex        =   10
            Tag             =   "ctl"
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Sagas"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   300
            TabIndex        =   7
            Tag             =   "ctl"
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Crit Calculator"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   9
            Left            =   2160
            TabIndex        =   11
            Tag             =   "ctl"
            Top             =   1320
            Width           =   1260
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Quick Access"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   6
            Left            =   120
            TabIndex        =   3
            Tag             =   "ctl"
            Top             =   0
            Width           =   1164
         End
         Begin VB.Shape shpQuickAccess 
            Height          =   2652
            Left            =   0
            Top             =   120
            Width           =   4092
         End
      End
      Begin VB.Label lnkFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Switch"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   6960
         TabIndex        =   25
         Tag             =   "ctl"
         Top             =   3204
         Width           =   588
      End
      Begin VB.Label lnkMove 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Move"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   7920
         TabIndex        =   26
         Tag             =   "ctl"
         Top             =   3204
         Width           =   492
      End
      Begin VB.Shape shpBorder 
         Height          =   492
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   312
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Save Path"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   11
         Left            =   600
         TabIndex        =   24
         Tag             =   "ctl"
         Top             =   3204
         Width           =   900
      End
   End
   Begin VB.CheckBox chkHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   0
      Width           =   972
   End
   Begin Compendium.userTab usrTab 
      Height          =   372
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4512
      _ExtentX        =   7959
      _ExtentY        =   656
      Captions        =   "Tools,Options"
   End
   Begin VB.ComboBox cboSize 
      Height          =   312
      Left            =   7260
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   1260
      Width           =   912
   End
   Begin VB.ComboBox cboFont 
      Height          =   312
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   1260
      Width           =   1992
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6492
      Index           =   1
      Left            =   0
      ScaleHeight     =   6492
      ScaleMode       =   0  'User
      ScaleWidth      =   9132
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "ctl"
      Top             =   360
      Visible         =   0   'False
      Width           =   9132
      Begin VB.CheckBox chkFonts 
         Caption         =   "Check Fonts"
         Height          =   492
         Index           =   2
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   4740
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.CheckBox chkFonts 
         Caption         =   "Create Patrons.csv"
         Height          =   492
         Index           =   1
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4080
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.CheckBox chkFonts 
         Caption         =   "Create Packs.csv"
         Height          =   492
         Index           =   0
         Left            =   900
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3420
         Visible         =   0   'False
         Width           =   2172
      End
      Begin VB.Frame fraQuests 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5592
         Left            =   4680
         TabIndex        =   38
         Tag             =   "ctl"
         Top             =   240
         Width           =   4092
         Begin VB.PictureBox picFont 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   3540
            ScaleHeight     =   348
            ScaleWidth      =   468
            TabIndex        =   61
            Top             =   840
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.ComboBox cboLevelSort 
            Height          =   312
            ItemData        =   "frmTools.frx":0010
            Left            =   840
            List            =   "frmTools.frx":001D
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   3300
            Width           =   2712
         End
         Begin Compendium.userSpinner usrspnVertical 
            Height          =   312
            Left            =   2580
            TabIndex        =   45
            Top             =   4440
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   550
            Min             =   0
            Value           =   2
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnHorizontal 
            Height          =   312
            Left            =   2580
            TabIndex        =   43
            Top             =   4080
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   550
            Min             =   0
            Value           =   4
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnWheel 
            Height          =   312
            Left            =   2580
            TabIndex        =   47
            Top             =   4980
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   550
            Min             =   0
            Max             =   20
            Value           =   6
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userCheckBox usrchkAbbrev 
            Height          =   252
            Index           =   0
            Left            =   480
            TabIndex        =   54
            Tag             =   "ctl"
            Top             =   1440
            Width           =   2832
            _ExtentX        =   4995
            _ExtentY        =   445
            Caption         =   "Abbreviate Columns"
         End
         Begin Compendium.userCheckBox usrchkAbbrev 
            Height          =   252
            Index           =   1
            Left            =   900
            TabIndex        =   55
            Tag             =   "ctl"
            Top             =   1800
            Width           =   1752
            _ExtentX        =   3090
            _ExtentY        =   445
            Caption         =   "Pack names"
         End
         Begin Compendium.userCheckBox usrchkAbbrev 
            Height          =   252
            Index           =   2
            Left            =   900
            TabIndex        =   56
            Tag             =   "ctl"
            Top             =   2160
            Width           =   2172
            _ExtentX        =   3831
            _ExtentY        =   445
            Caption         =   "Patron names"
         End
         Begin VB.Label lnkDefaultFont 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Default"
            ForeColor       =   &H00FF0000&
            Height          =   216
            Left            =   3300
            TabIndex        =   65
            Tag             =   "ctl"
            Top             =   0
            Width           =   624
         End
         Begin VB.Line linQuests 
            Index           =   2
            X1              =   0
            X2              =   4080
            Y1              =   3840
            Y2              =   3840
         End
         Begin VB.Line linQuests 
            Index           =   1
            X1              =   0
            X2              =   4080
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line linQuests 
            Index           =   0
            X1              =   0
            X2              =   4080
            Y1              =   1260
            Y2              =   1260
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Size"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   13
            Left            =   2580
            TabIndex        =   58
            Tag             =   "ctl"
            Top             =   420
            Width           =   372
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Font"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   12
            Left            =   480
            TabIndex        =   57
            Tag             =   "ctl"
            Top             =   420
            Width           =   408
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Vertical Margins"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   5
            Left            =   1020
            TabIndex        =   44
            Tag             =   "ctl"
            Top             =   4476
            Width           =   1392
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Horizontal Margins"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   756
            TabIndex        =   42
            Tag             =   "ctl"
            Top             =   4116
            Width           =   1656
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "When Sorting Quests by Level"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   600
            TabIndex        =   40
            Tag             =   "ctl"
            Top             =   2940
            Width           =   2676
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Quests"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   9
            Left            =   120
            TabIndex        =   39
            Tag             =   "ctl"
            Top             =   0
            Width           =   624
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Mouse Wheel Step"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   10
            Left            =   756
            TabIndex        =   46
            Tag             =   "ctl"
            Top             =   5016
            Width           =   1656
         End
         Begin VB.Shape shpQuests 
            Height          =   5472
            Left            =   0
            Top             =   120
            Width           =   4092
         End
      End
      Begin VB.TextBox txtPlay 
         Height          =   324
         Left            =   540
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   6000
         Width           =   7632
      End
      Begin VB.CheckBox chkPlay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "..."
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   8172
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5400
         Width           =   312
      End
      Begin VB.Frame fraMainWindow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2772
         Left            =   360
         TabIndex        =   29
         Tag             =   "ctl"
         Top             =   240
         Width           =   4092
         Begin Compendium.userCheckBox usrchkChildWindows 
            Height          =   252
            Left            =   1080
            TabIndex        =   37
            Tag             =   "ctl"
            Top             =   2160
            Width           =   2412
            _ExtentX        =   4255
            _ExtentY        =   445
            Caption         =   "Child Windows"
            CheckPosition   =   1
         End
         Begin VB.ComboBox cboWindow 
            Height          =   312
            ItemData        =   "frmTools.frx":0060
            Left            =   780
            List            =   "frmTools.frx":006D
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   780
            Width           =   2712
         End
         Begin Compendium.userSpinner usrspnBottom 
            Height          =   312
            Left            =   2520
            TabIndex        =   36
            Top             =   1620
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   550
            Min             =   0
            Max             =   20
            Value           =   8
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSides 
            Height          =   312
            Left            =   2520
            TabIndex        =   34
            Top             =   1260
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   550
            Min             =   0
            Max             =   20
            Value           =   8
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Main Window Startup Position"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   540
            TabIndex        =   31
            Tag             =   "ctl"
            Top             =   420
            Width           =   2892
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Bottom Margin"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   1020
            TabIndex        =   35
            Tag             =   "ctl"
            Top             =   1656
            Width           =   1332
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Side Margins"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   1236
            TabIndex        =   33
            Tag             =   "ctl"
            Top             =   1296
            Width           =   1116
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Main Window"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Tag             =   "ctl"
            Top             =   0
            Width           =   1176
         End
         Begin VB.Shape shpMainWindow 
            Height          =   2652
            Left            =   0
            Top             =   120
            Width           =   4092
         End
      End
      Begin Compendium.userCheckBox usrchkPlay 
         Height          =   252
         Left            =   540
         TabIndex        =   48
         Tag             =   "ctl"
         Top             =   5700
         Width           =   1872
         _ExtentX        =   3302
         _ExtentY        =   445
         Caption         =   "Play Button"
      End
      Begin VB.Shape shpBorder 
         Height          =   492
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   312
      End
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnNew As Boolean

Private mblnOverride As Boolean


' ************* TEMP *************


Private Sub chkFonts_Click(Index As Integer)
    If UncheckButton(Me.chkFonts(Index), mblnOverride) Then Exit Sub
    Select Case Index
        Case 0: PackWidths
        Case 1: PatronWidths
        Case 2: TestFonts
    End Select
End Sub

Private Sub TestFonts()
    Dim strFont() As String
    
    strFont = GetFontList(Me.picFont)
End Sub


' ************* FORM *************


Private Sub Form_Load()
    mblnNew = False
    cfg.Configure Me
    SizeBorders
    LoadData
    If Not xp.DebugMode Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.Hwnd)
    cfg.SavePosition Me
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case True
        Case IsOver(Me.usrspnSides.Hwnd, Xpos, Ypos): Me.usrspnSides.WheelScroll lngValue
        Case IsOver(Me.usrspnBottom.Hwnd, Xpos, Ypos): Me.usrspnBottom.WheelScroll lngValue
        Case IsOver(Me.usrspnHorizontal.Hwnd, Xpos, Ypos): Me.usrspnHorizontal.WheelScroll lngValue
        Case IsOver(Me.usrspnVertical.Hwnd, Xpos, Ypos): Me.usrspnVertical.WheelScroll lngValue
        Case IsOver(Me.usrspnWheel.Hwnd, Xpos, Ypos): Me.usrspnWheel.WheelScroll lngValue
    End Select
End Sub

Public Sub LevelOrderChange()
    mblnOverride = True
    Me.cboLevelSort.ListIndex = cfg.LevelSort
    mblnOverride = False
End Sub

Private Sub chkHelp_Click()
    If UncheckButton(Me.chkHelp, mblnOverride) Then Exit Sub
    ShowHelp Me.usrTab.ActiveTab
End Sub

Private Sub SizeBorders()
    Dim i As Long
    
    Me.usrTab.Width = Me.usrTab.TabsWidth
    For i = 0 To 1
        With Me.picTab(i)
            Me.shpBorder(i).Move 0, 0, .ScaleWidth, .ScaleHeight
        End With
    Next
End Sub

Private Sub usrTab_Click(pstrCaption As String)
    Dim i As Long
    
    For i = 0 To 1
        Me.picTab(i).Visible = (Me.usrTab.ActiveTabIndex = i)
    Next
End Sub


' ************* LOAD *************


Private Sub LoadData()
    Dim ctl As Control
    Dim lngMax As Long
    Dim strFontName As String
    Dim dblFontSize As Double
    Dim strFont() As String
    Dim i As Long
    
    mblnOverride = True
    Me.cboLevelSort.ListIndex = cfg.LevelSort
    Me.cboWindow.ListIndex = cfg.WindowSize
    Me.usrspnSides.Value = cfg.Sides
    Me.usrspnBottom.Value = cfg.Bottom
    Me.usrchkChildWindows.Value = cfg.ChildWindows
    Me.usrspnHorizontal.Value = cfg.MarginX
    Me.usrspnVertical.Value = cfg.MarginY
    lngMax = frmCompendium.usrQuests.PageSize
    If lngMax < 10 Then lngMax = 10
    Me.usrspnWheel.Max = lngMax
    Me.usrspnWheel.Value = cfg.WheelStep
    Me.usrchkAbbrev(0).Value = cfg.AbbreviateColumns
    Me.usrchkAbbrev(1).Value = cfg.AbbreviatePacks
    Me.usrchkAbbrev(1).Enabled = cfg.AbbreviateColumns
    Me.usrchkAbbrev(2).Value = cfg.AbbreviatePatrons
    Me.usrchkAbbrev(2).Enabled = cfg.AbbreviateColumns
    PopulateListbox
    Me.txtDataFile.Text = cfg.DataFile
    Me.usrchkPlay.Value = cfg.PlayButton
    Me.usrchkPlay.Width = Me.usrchkPlay.FitWidth
    Me.txtPlay.Text = cfg.PlayEXE
    Me.txtSavePath.Text = cfg.CompendiumPath
    ' Font
    ComboListHeightChild Me.cboFont, 20, Me
    ComboListHeightChild Me.cboSize, 20, Me
    frmCompendium.GetQuestsFont strFontName, dblFontSize
    ComboClear Me.cboFont
    strFont = GetFontList(Me.picFont)
    For i = 1 To UBound(strFont)
        Me.cboFont.AddItem strFont(i)
    Next
    ComboSetText Me.cboFont, strFontName
    PopulateFontSizes dblFontSize
    mblnOverride = False
End Sub

Private Sub PopulateFontSizes(Optional pdblSelected As Double = 0)
    Dim strFontName As String
    Dim strSelected As String
    Dim strSize() As String
    Dim i As Long
    
    strSelected = FontSizeToString(pdblSelected)
    ComboClear Me.cboSize
    strSize = GetFontSizes(Me.cboFont.Text, Me.picFont)
    For i = 1 To UBound(strSize)
        Me.cboSize.AddItem strSize(i)
    Next
    ComboSetText Me.cboSize, strSelected
End Sub

Private Sub lnkDefaultFont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkDefaultFont_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dblSize As Double
    
    xp.SetMouseCursor mcHand
    dblSize = frmCompendium.SetQuestsFont("Arial Narrow", 10.2)
    mblnOverride = True
    ComboSetText Me.cboFont, "Arial Narrow"
    PopulateFontSizes dblSize
    mblnOverride = False
End Sub

Private Sub lnkDefaultFont_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub PopulateListbox()
    Dim strDataFile As String
    Dim strFile As String
    Dim lngCount As Long
    
    ListboxClear Me.lstDataFile
    If xp.File.Exists(DataFileToFileName("Main")) Then
        Me.lstDataFile.AddItem "Main"
        lngCount = 1
    End If
    strFile = Dir(cfg.CompendiumPath & "\*.compendium")
    Do While Len(strFile)
        strDataFile = FileNameToDataFile(strFile)
        If strDataFile <> "Main" Then
            If LCase$(Right$(strDataFile, 7)) <> "_backup" Then
                Me.lstDataFile.AddItem strDataFile
                lngCount = lngCount + 1
            End If
        End If
        strFile = Dir
    Loop
    If lngCount = 0 Then
        Me.lstDataFile.AddItem "Main"
        cfg.DataFile = "Main"
    End If
    ListboxSetText Me.lstDataFile, cfg.DataFile
End Sub

Private Sub EnableControls()
    Dim blnEnabled As Boolean
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        Do
            Select Case TypeName(ctl)
                Case "Shape", "userTab": Exit Do
            End Select
            Select Case ctl.Name
                Case "lnkNew", "fraData", "txtDataFile", "picTab": ctl.Enabled = True
                Case "lnkSave"
                    If Len(Me.txtDataFile.Text) = 0 Then
                        ctl.Enabled = False
                    ElseIf mblnNew Then
                        ctl.Enabled = True
                    Else
                        ctl.Enabled = (Me.txtDataFile.Text <> Me.lstDataFile.Text)
                    End If
                Case Else: ctl.Enabled = Not mblnNew
            End Select
        Loop Until True
    Next
End Sub


' ************* QUICK ACCESS *************


Private Sub lnkLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    Select Case Me.lnkLink(Index).Caption
        Case "Characters": OpenForm "frmCharacter"
        Case "Challenges": OpenForm "frmChallenges"
        Case "Patrons": OpenForm "frmPatrons"
        Case "Sagas": OpenForm "frmSagas"
        Case "Colors", "Custom Colors": cfg.RunUtil ueColors
        Case "Named", "Named Colors": OpenForm "frmColorPreview"
        Case "Shroud Puzzles": cfg.RunUtil ueLightsOut
        Case "ADQ Riddle": cfg.RunUtil ueADQ
        Case "Stopwatch": cfg.RunUtil ueStopwatch
        Case "Crit Calculator": OpenForm "frmCritCalculator"
        Case "Revert Characters": RevertCharacters
        Case "Revert Links": RevertLinks
        Case "Revert Notes": RevertNotes
        Case "Refresh Backups": RefreshBackups
        Case "Wilderness Areas": OpenForm "frmWilderness"
        Case "About": frmAbout.Show vbModal, Me
        Case Else: MsgBox "Feature under construction.", vbInformation, "Sorry..."
    End Select
End Sub


' ************* DATAFILE *************


Private Sub lstDataFile_Click()
    If mblnOverride Then Exit Sub
    If Me.lstDataFile.ListIndex = -1 Then
        If Me.lstDataFile.ListCount > 0 Then Me.lstDataFile.ListIndex = 0
        Exit Sub
    End If
    mblnOverride = True
    Me.txtDataFile.Text = Me.lstDataFile.Text
    mblnOverride = False
    If cfg.DataFile <> Me.lstDataFile.Text Then ChangeCompendiumFile Me.lstDataFile.Text
End Sub

Private Sub chkDelete_Click()
    Dim lngIndex As Long
    Dim strFile As String
    
    If UncheckButton(Me.chkDelete, mblnOverride) Then Exit Sub
    If Me.lstDataFile.ListIndex = -1 Then Exit Sub
    AutoSave
    If MsgBox("Delete " & Me.lstDataFile.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    With Me.lstDataFile
        mblnOverride = True
        DeleteFile CompendiumBackupFile()
        DeleteFile CompendiumFile()
        ClearData
        lngIndex = .ListIndex
        .RemoveItem .ListIndex
        If lngIndex > .ListCount - 1 Then lngIndex = .ListCount - 1
        If lngIndex = -1 Then
            cfg.DataFile = "Main"
            LoadCompendium True
            .AddItem "Main"
            lngIndex = 0
        End If
        mblnOverride = False
        cfg.DataFile = vbNullString
        .ListIndex = lngIndex
    End With
End Sub

Private Sub DeleteFile(pstrFile As String)
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile
End Sub

Private Sub ClearData()
    Dim i As Long
    
    Erase db.Character
    db.Characters = 0
    For i = 1 To db.Quests
        Erase db.Quest(i).Progress
    Next
    For i = 1 To db.Challenges
        Erase db.Challenge(i).Stars
    Next
End Sub

Private Sub lnkNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    If Me.lnkNew.Caption = "Cancel" And Not mblnNew Then
        Me.lnkNew.Caption = "New"
        mblnOverride = True
        Me.txtDataFile.Text = Me.lstDataFile.Text
        mblnOverride = False
        Me.lstDataFile.SetFocus
        Exit Sub
    End If
    mblnNew = Not mblnNew
    EnableControls
    If mblnNew Then
        Me.lnkNew.Caption = "Cancel"
        Me.txtDataFile.Text = vbNullString
        Me.txtDataFile.SetFocus
    Else
        Me.lnkNew.Caption = "New"
        Me.txtDataFile.Text = Me.lstDataFile.Text
        Me.lstDataFile.SetFocus
    End If
End Sub

Private Sub txtDataFile_GotFocus()
    TextboxGotFocus Me.txtDataFile
End Sub

Private Sub txtDataFile_Change()
    If mblnNew Then
        Me.lnkSave.Enabled = (Len(Me.txtDataFile.Text) > 0)
    Else
        Me.lnkSave.Enabled = (Me.txtDataFile.Text <> Me.lstDataFile.Text) And (Len(Me.txtDataFile.Text) > 0)
        If Me.lnkSave.Enabled Then Me.lnkNew.Caption = "Cancel" Else Me.lnkNew.Caption = "New"
    End If
End Sub

Private Sub lnkSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strFile As String
    
    xp.SetMouseCursor mcHand
    mblnOverride = True
    Me.txtDataFile.Text = ProperName(Me.txtDataFile.Text)
    mblnOverride = False
    Me.txtDataFile.Refresh
    If Len(Me.txtDataFile.Text) = 0 Then
        MsgBox "Invalid file name", vbInformation, "Notice"
        Exit Sub
    End If
    If cfg.DataFile = Me.txtDataFile.Text And Not mblnNew Then
        MsgBox "Name not changed.", vbInformation, "Notice"
        Exit Sub
    End If
    If xp.File.Exists(DataFileToFileName(Me.txtDataFile.Text)) Then
        MsgBox "File already exists:" & vbNewLine & vbNewLine & DataFileToFileName(Me.txtDataFile.Text), vbInformation, "Notice"
        Exit Sub
    End If
    If mblnNew Then
        If MsgBox("Create new Compendium file named " & Me.txtDataFile.Text & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
        ChangeCompendiumFile Me.txtDataFile.Text
        mblnOverride = True
        Me.lstDataFile.AddItem Me.txtDataFile.Text
        Me.lstDataFile.ListIndex = Me.lstDataFile.NewIndex
        mblnOverride = False
        Me.txtDataFile.Text = Me.lstDataFile.Text
    Else
        If MsgBox("Rename " & cfg.DataFile & " to " & Me.txtDataFile.Text & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
        RenameCompendiumFile Me.txtDataFile.Text
        With Me.lstDataFile
            If .ListIndex <> -1 Then .List(.ListIndex) = Me.txtDataFile.Text
        End With
    End If
    Me.lnkNew.Caption = "New"
    mblnNew = False
    EnableControls
End Sub


' ************* SAVE PATH *************


Private Sub lnkFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strFile As String
    
    xp.SetMouseCursor mcHand
    strFile = xp.ShowOpenDialog(cfg.CompendiumPath, "Compendium Files (*.compendium)|*.compendium", "*.compendium")
    If Len(strFile) Then SwitchSavePath strFile
End Sub

Private Sub SwitchSavePath(pstrFile As String)
    Dim strPath As String
    
    If Not xp.File.Exists(pstrFile) Then
        Notice pstrFile & " not found"
        Exit Sub
    End If
    DeleteBackups
    strPath = GetPathFromFilespec(pstrFile)
    cfg.CompendiumPath = strPath
    cfg.DataFile = GetNameFromFilespec(pstrFile)
    BackupFiles
    OpenCompendiumFile False
    frmCompendium.ReloadLinkLists
    frmCompendium.usrtxtNotes.Text = LoadNotes()
    LoadData
    Me.txtSavePath.Text = cfg.CompendiumPath
End Sub

Private Sub lnkFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strPath As String
    Dim strFile As String
    
    xp.SetMouseCursor mcHand
    strFile = xp.ShowSaveAsDialog(cfg.CompendiumPath, cfg.DataFile & ".compendium", "Compendium Files|*.compendium", "*.compendium")
    If Len(strFile) = 0 Then Exit Sub
    strPath = GetPathFromFilespec(strFile)
    MoveSavePath strPath
End Sub

Private Sub lnkMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub MoveSavePath(pstrNew As String)
    Const TwoLines As String = vbNewLine & vbNewLine
    Dim strNew As String
    Dim strOld As String
    Dim strFound() As String
    Dim lngFound As Long
    Dim strSource As String
    Dim strDest As String
    Dim i As Long
    
    strNew = pstrNew
    strOld = cfg.CompendiumPath
    If strNew = strOld Then
        Notice "New path same as old path. No action taken."
        Exit Sub
    End If
    If Not ValidPath(strOld, "Current") Then Exit Sub
    If Not ValidPath(strNew, "New") Then Exit Sub
    lngFound = FindExistingFiles(strNew, strFound)
    If lngFound = 0 Then
        If Not AskAlways("Move all Compendium save files to: " & TwoLines & strNew & TwoLines & "Are you sure?") Then Exit Sub
    Else
        If MsgBox("The following save files already exist in new folder:" & vbNewLine & Join(strFound, vbNewLine & "     ") & TwoLines & "Choosing OK will overwrite them with your current files.", vbInformation + vbOKCancel + vbDefaultButton2, "Notice") = vbCancel Then
            Notice "No action taken."
            Exit Sub
        End If
    End If
    AutoSave
    lngFound = FindExistingFiles(strOld, strFound)
    If lngFound = 0 Then
        Notice "No files found"
        Exit Sub
    End If
    For i = 1 To lngFound
        strSource = strOld & "\" & strFound(i)
        strDest = strNew & "\" & strFound(i)
        If xp.File.Exists(strDest) Then xp.File.Delete strDest
        xp.File.Move strSource, strDest
    Next
    cfg.CompendiumPath = strNew
    Me.txtSavePath.Text = cfg.CompendiumPath
    Notice "Save Path changed to: " & vbNewLine & vbNewLine & pstrNew
End Sub

Private Function ValidPath(pstrPath As String, pstrDescrip As String) As Boolean
    If xp.Folder.Exists(pstrPath) Then
        ValidPath = True
    Else
        Notice pstrDescrip & " path not found: " & vbNewLine & vbNewLine & pstrPath
    End If
End Function

Private Function FindExistingFiles(pstrPath As String, pstrFile() As String) As Long
    Dim strFile As String
    Dim strList As String
    Dim strCheck As String
    
    strFile = Dir(pstrPath & "\*.compendium")
    Do While Len(strFile)
        strList = strList & vbTab & strFile
        strFile = Dir
    Loop
    AddExistingFile pstrPath, "Compendium.linklists", strList
    AddExistingFile pstrPath, LinkListsBackupFile(True), strList
    AddExistingFile pstrPath, "Notes.txt", strList
    AddExistingFile pstrPath, NotesBackupFile(True), strList
    If Len(strList) = 0 Then Exit Function
    pstrFile = Split(strList, vbTab)
    FindExistingFiles = UBound(pstrFile)
End Function

Private Sub AddExistingFile(pstrPath As String, pstrFile As String, pstrList As String)
    If xp.File.Exists(pstrPath & "\" & pstrFile) Then pstrList = pstrList & vbTab & pstrFile
End Sub


' ************* OPTIONS *************


Private Sub cboWindow_Click()
    If mblnOverride Then Exit Sub
    cfg.WindowSize = Me.cboWindow.ListIndex
    cfg.SizeWindow
    DirtyFlag dfeSettings
End Sub

Private Sub usrspnSides_Change()
    If mblnOverride Then Exit Sub
    cfg.Sides = Me.usrspnSides.Value
    frmCompendium.RedrawQuests
    DirtyFlag dfeSettings
End Sub

Private Sub usrspnBottom_Change()
    If mblnOverride Then Exit Sub
    cfg.Bottom = Me.usrspnBottom.Value
    frmCompendium.RedrawQuests
    DirtyFlag dfeSettings
End Sub

Private Sub usrchkChildWindows_UserChange()
    If mblnOverride Then Exit Sub
    If Not Me.usrchkChildWindows.Value Then
        If Not Ask("Child Windows are strongly recommended. Are you sure?") Then
            Me.usrchkChildWindows.Value = True
            Exit Sub
        End If
    End If
    cfg.ChildWindows = Me.usrchkChildWindows.Value
End Sub

Private Sub cboLevelSort_Click()
    If mblnOverride Then Exit Sub
    If Me.cboLevelSort.ListIndex = -1 Then cfg.LevelSort = 0 Else cfg.LevelSort = Me.cboLevelSort.ListIndex
    cfg.CompendiumOrder = coeLevel
    frmCompendium.usrQuests.ReQuery
    DirtyFlag dfeSettings
End Sub

Private Sub usrspnHorizontal_Change()
    If mblnOverride Then Exit Sub
    cfg.MarginX = Me.usrspnHorizontal.Value
    ReDrawGrids
    DirtyFlag dfeSettings
End Sub

Private Sub usrspnVertical_Change()
    If mblnOverride Then Exit Sub
    cfg.MarginY = Me.usrspnVertical.Value
    ReDrawGrids
    DirtyFlag dfeSettings
End Sub

Private Sub ReDrawGrids()
    Dim frm As Form
    
    frmCompendium.RedrawQuests
    If GetForm(frm, "frmPatrons") Then frm.ReDrawForm
    If GetForm(frm, "frmWilderness") Then frm.ReDrawForm
    If GetForm(frm, "frmSagas") Then frm.Redraw
End Sub

Private Sub usrspnWheel_Change()
    If mblnOverride Then Exit Sub
    cfg.WheelStep = Me.usrspnWheel.Value
    DirtyFlag dfeSettings
End Sub

Private Sub usrchkAbbrev_UserChange(Index As Integer)
    Dim blnValue As Boolean
    
    If mblnOverride Then Exit Sub
    blnValue = Me.usrchkAbbrev(Index).Value
    Select Case Index
        Case 0
            cfg.AbbreviateColumns = blnValue
            Me.usrchkAbbrev(1).Enabled = blnValue
            Me.usrchkAbbrev(2).Enabled = blnValue
        Case 1
            cfg.AbbreviatePacks = blnValue
        Case 2
            cfg.AbbreviatePatrons = blnValue
    End Select
    frmCompendium.RedrawQuests
    DirtyFlag dfeSettings
End Sub


' ************* FONT *************


Private Sub cboFont_Click()
    Dim dblSize As Double
    
    If mblnOverride Then Exit Sub
    dblSize = frmCompendium.SetQuestsFont(Me.cboFont.Text)
    PopulateFontSizes dblSize
End Sub

Private Sub cboSize_Click()
    If mblnOverride Then Exit Sub
    frmCompendium.SetQuestsFont Me.cboFont.Text, Val(Me.cboSize.Text)
End Sub


' ************* PLAY BUTTON *************


Private Sub usrchkPlay_UserChange()
    cfg.PlayButton = Me.usrchkPlay.Value
    frmCompendium.ShowPlayButton
    DirtyFlag dfeSettings
End Sub

Private Sub txtPlay_Change()
    If mblnOverride Then Exit Sub
    cfg.PlayEXE = Me.txtPlay.Text
End Sub

Private Sub chkPlay_Click()
    If UncheckButton(Me.chkPlay, mblnOverride) Then Exit Sub
    If ChooseEXE() Then Me.txtPlay.Text = cfg.PlayEXE
End Sub
