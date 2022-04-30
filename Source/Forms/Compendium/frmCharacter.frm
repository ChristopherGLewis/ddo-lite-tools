VERSION 5.00
Begin VB.Form frmCharacter 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Characters"
   ClientHeight    =   8805
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraImport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3972
      Left            =   2640
      TabIndex        =   237
      Top             =   240
      Visible         =   0   'False
      Width           =   2712
      Begin VB.CheckBox chkImport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Done"
         ForeColor       =   &H80000008&
         Height          =   432
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   241
         Top             =   3300
         Width           =   1332
      End
      Begin VB.CheckBox chkImport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Import"
         ForeColor       =   &H80000008&
         Height          =   432
         Index           =   0
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   2820
         Width           =   1332
      End
      Begin VB.ListBox lstImport 
         Height          =   2160
         ItemData        =   "frmCharacter.frx":000C
         Left            =   600
         List            =   "frmCharacter.frx":000E
         TabIndex        =   238
         Top             =   420
         Width           =   1572
      End
      Begin VB.Label lblImport 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Import Character"
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   120
         TabIndex        =   239
         Top             =   0
         Width           =   1548
      End
      Begin VB.Shape shpImport 
         Height          =   3852
         Left            =   0
         Top             =   120
         Width           =   2712
      End
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Export"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   8
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3780
      Width           =   1392
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Import"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   7
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3300
      Width           =   1392
   End
   Begin Compendium.userTab usrTab 
      Height          =   312
      Left            =   360
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4440
      Width           =   3552
      _ExtentX        =   6271
      _ExtentY        =   556
      Captions        =   "Notes,Tomes,Past Lives"
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sagas"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   6
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   235
      Top             =   3300
      Width           =   1452
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reincarnate"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   5
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   236
      Top             =   3780
      Width           =   1452
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Help"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   4
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   234
      Top             =   1920
      Width           =   1212
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Apply"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   3
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   233
      Top             =   1320
      Width           =   1212
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   2
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   840
      Width           =   1212
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   1
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   231
      Top             =   360
      Width           =   1212
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2352
      Left            =   5580
      TabIndex        =   20
      Top             =   1860
      Width           =   2892
      Begin Compendium.userCheckBox usrchkCustomColors 
         Height          =   252
         Left            =   300
         TabIndex        =   24
         Top             =   960
         Width           =   1992
         _ExtentX        =   3519
         _ExtentY        =   450
         Value           =   0   'False
         Caption         =   "Custom Colors"
      End
      Begin VB.ComboBox cboColor 
         Height          =   312
         ItemData        =   "frmCharacter.frx":0010
         Left            =   1080
         List            =   "frmCharacter.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   420
         Width           =   1512
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   1
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1800
         Width           =   492
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   0
         Left            =   600
         ScaleHeight     =   285
         ScaleWidth      =   465
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1440
         Width           =   492
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Colors"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   576
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Skipped"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   8
         Left            =   1260
         TabIndex        =   28
         Top             =   1860
         Width           =   696
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Default"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   7
         Left            =   1260
         TabIndex        =   26
         Top             =   1500
         Width           =   624
      End
      Begin VB.Label lnkLink 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Named"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   264
         TabIndex        =   22
         Top             =   468
         Width           =   636
      End
      Begin VB.Shape shpColors 
         Height          =   2232
         Left            =   0
         Top             =   120
         Width           =   2892
      End
   End
   Begin VB.Frame fraDifficulty 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1452
      Left            =   5580
      TabIndex        =   14
      Top             =   240
      Width           =   2892
      Begin VB.TextBox txtTotalFavor 
         Alignment       =   2  'Center
         Height          =   324
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   19
         Top             =   840
         Width           =   912
      End
      Begin VB.TextBox txtChallengeFavor 
         Alignment       =   2  'Center
         Height          =   324
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   17
         Top             =   420
         Width           =   912
      End
      Begin VB.Label lnkLink 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   756
         TabIndex        =   18
         Top             =   876
         Width           =   444
      End
      Begin VB.Label lnkLink 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Challenges"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   0
         Left            =   252
         TabIndex        =   16
         Top             =   456
         Width           =   948
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Favor"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   516
      End
      Begin VB.Shape shpDifficulty 
         Height          =   1332
         Left            =   0
         Top             =   120
         Width           =   2892
      End
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Add New"
      ForeColor       =   &H80000008&
      Height          =   432
      Index           =   0
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1392
   End
   Begin VB.ListBox lstCharacter 
      Height          =   1740
      ItemData        =   "frmCharacter.frx":0014
      Left            =   360
      List            =   "frmCharacter.frx":0016
      TabIndex        =   1
      Top             =   480
      Width           =   1512
   End
   Begin VB.Frame fraHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2352
      Left            =   2640
      TabIndex        =   8
      Top             =   1860
      Width           =   2712
      Begin VB.CheckBox chkContext 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Customize"
         ForeColor       =   &H80000008&
         Height          =   432
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1992
      End
      Begin VB.ComboBox cboLeftClick 
         Height          =   312
         ItemData        =   "frmCharacter.frx":0018
         Left            =   360
         List            =   "frmCharacter.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   1992
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Right Click Menu"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   420
         Width           =   1992
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Column Header"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   1380
      End
      Begin VB.Shape shpHeader 
         Height          =   2232
         Left            =   0
         Top             =   120
         Width           =   2712
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Left Click Action"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   4
         Left            =   360
         TabIndex        =   12
         Top             =   1380
         Width           =   1992
      End
   End
   Begin VB.Frame fraName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1452
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   2712
      Begin VB.TextBox txtName 
         Height          =   324
         Left            =   360
         MaxLength       =   12
         TabIndex        =   7
         Top             =   420
         Width           =   1992
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Character Name"
         ForeColor       =   &H80000008&
         Height          =   216
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1464
      End
      Begin VB.Shape shpName 
         Height          =   1332
         Left            =   0
         Top             =   120
         Width           =   2712
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   1
      Left            =   360
      ScaleHeight     =   3855
      ScaleWidth      =   9735
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "ctl"
      Top             =   4740
      Visible         =   0   'False
      Width           =   9732
      Begin VB.Frame fraMisc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   2580
         TabIndex        =   53
         Tag             =   "ctl"
         Top             =   300
         Width           =   2232
         Begin Compendium.userSpinner usrspnFate 
            Height          =   255
            Left            =   1260
            TabIndex        =   59
            Top             =   660
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnPower 
            Height          =   255
            Index           =   1
            Left            =   1260
            TabIndex        =   61
            Top             =   1680
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   2
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnPower 
            Height          =   255
            Index           =   2
            Left            =   1260
            TabIndex        =   62
            Top             =   1920
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   2
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnPower 
            Height          =   255
            Index           =   3
            Left            =   1260
            TabIndex        =   64
            Top             =   2160
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   2
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRR 
            Height          =   255
            Index           =   1
            Left            =   1260
            TabIndex        =   67
            Top             =   2580
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   2
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRR 
            Height          =   255
            Index           =   2
            Left            =   1260
            TabIndex        =   69
            Top             =   2820
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   2
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRacialAP 
            Height          =   252
            Left            =   1260
            TabIndex        =   55
            Top             =   0
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   1
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnUniversalAP 
            Height          =   255
            Left            =   1260
            TabIndex        =   57
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnDestiny 
            Height          =   255
            Left            =   1260
            TabIndex        =   60
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   1
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Destiny"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   106
            Left            =   0
            TabIndex        =   247
            Tag             =   "ctl"
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Univ AP"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   105
            Left            =   0
            TabIndex        =   246
            Tag             =   "ctl"
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Racial AP"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   15
            Left            =   0
            TabIndex        =   54
            Tag             =   "ctl"
            Top             =   12
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "MRR"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   92
            Left            =   0
            TabIndex        =   68
            Tag             =   "ctl"
            Top             =   2835
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "PRR"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   91
            Left            =   0
            TabIndex        =   66
            Tag             =   "ctl"
            Top             =   2595
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Spellpower"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   90
            Left            =   0
            TabIndex        =   65
            Tag             =   "ctl"
            Top             =   2175
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ranged"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   89
            Left            =   0
            TabIndex        =   63
            Tag             =   "ctl"
            Top             =   1935
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Melee"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   88
            Left            =   0
            TabIndex        =   58
            Tag             =   "ctl"
            Top             =   1695
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fate"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   87
            Left            =   0
            TabIndex        =   56
            Tag             =   "ctl"
            Top             =   675
            Width           =   1140
         End
      End
      Begin VB.Frame fraXPTomes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1032
         Left            =   180
         TabIndex        =   47
         Tag             =   "ctl"
         Top             =   2100
         Width           =   2232
         Begin VB.ComboBox cboEpicXP 
            Height          =   330
            ItemData        =   "frmCharacter.frx":0049
            Left            =   840
            List            =   "frmCharacter.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   600
            Width           =   1332
         End
         Begin VB.ComboBox cboHeroicXP 
            Height          =   330
            ItemData        =   "frmCharacter.frx":006D
            Left            =   840
            List            =   "frmCharacter.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   300
            Width           =   1332
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "XP Tomes"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   86
            Left            =   924
            TabIndex        =   48
            Tag             =   "ctl"
            Top             =   0
            Width           =   912
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Epic"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   85
            Left            =   72
            TabIndex        =   51
            Tag             =   "ctl"
            Top             =   648
            Width           =   648
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Heroic"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   84
            Left            =   72
            TabIndex        =   49
            Tag             =   "ctl"
            Top             =   348
            Width           =   648
         End
      End
      Begin VB.Frame fraSkill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3252
         Left            =   4920
         TabIndex        =   70
         Tag             =   "ctl"
         Top             =   300
         Width           =   4692
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   1
            Left            =   1164
            TabIndex        =   72
            Top             =   0
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   10
            Left            =   1164
            TabIndex        =   74
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   19
            Left            =   1164
            TabIndex        =   76
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   20
            Left            =   1164
            TabIndex        =   78
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   2
            Left            =   1164
            TabIndex        =   80
            Top             =   1200
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   4
            Left            =   1164
            TabIndex        =   82
            Top             =   1440
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   9
            Left            =   1164
            TabIndex        =   84
            Top             =   1680
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   8
            Left            =   3684
            TabIndex        =   112
            Top             =   2400
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   3
            Left            =   3684
            TabIndex        =   94
            Top             =   0
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   17
            Left            =   3684
            TabIndex        =   96
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   7
            Left            =   3684
            TabIndex        =   98
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   15
            Left            =   3684
            TabIndex        =   100
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   16
            Left            =   3684
            TabIndex        =   102
            Top             =   1200
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   18
            Left            =   3684
            TabIndex        =   104
            Top             =   1440
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   5
            Left            =   3684
            TabIndex        =   108
            Top             =   1680
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   13
            Left            =   3684
            TabIndex        =   110
            Top             =   1920
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   21
            Left            =   1164
            TabIndex        =   86
            Top             =   2160
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   14
            Left            =   1164
            TabIndex        =   88
            Top             =   2400
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   6
            Left            =   1164
            TabIndex        =   90
            Top             =   2640
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   11
            Left            =   1164
            TabIndex        =   92
            Top             =   2880
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnSkill 
            Height          =   252
            Index           =   12
            Left            =   3684
            TabIndex        =   114
            Top             =   2640
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   5
            Value           =   0
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Repair"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   30
            Left            =   2160
            TabIndex        =   99
            Tag             =   "ctl"
            Top             =   732
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Heal"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   29
            Left            =   2160
            TabIndex        =   97
            Tag             =   "ctl"
            Top             =   492
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Spellcraft"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   28
            Left            =   2160
            TabIndex        =   95
            Tag             =   "ctl"
            Top             =   252
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Concentration"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   27
            Left            =   2160
            TabIndex        =   93
            Tag             =   "ctl"
            Top             =   12
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Search"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   31
            Left            =   2160
            TabIndex        =   101
            Tag             =   "ctl"
            Top             =   1212
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Spot"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   32
            Left            =   2160
            TabIndex        =   103
            Tag             =   "ctl"
            Top             =   1452
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Disable Device"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   33
            Left            =   2160
            TabIndex        =   107
            Tag             =   "ctl"
            Top             =   1692
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Open Lock"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   34
            Left            =   2160
            TabIndex        =   109
            Tag             =   "ctl"
            Top             =   1932
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Hide"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   35
            Left            =   2160
            TabIndex        =   111
            Tag             =   "ctl"
            Top             =   2412
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Move Silently"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   36
            Left            =   2160
            TabIndex        =   113
            Tag             =   "ctl"
            Top             =   2652
            Width           =   1404
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Listen"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   26
            Left            =   0
            TabIndex        =   91
            Tag             =   "ctl"
            Top             =   2892
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Haggle"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   25
            Left            =   0
            TabIndex        =   89
            Tag             =   "ctl"
            Top             =   2652
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Perform"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   24
            Left            =   0
            TabIndex        =   87
            Tag             =   "ctl"
            Top             =   2412
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "UMD"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   23
            Left            =   0
            TabIndex        =   85
            Tag             =   "ctl"
            Top             =   2172
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Intimidate"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   22
            Left            =   0
            TabIndex        =   83
            Tag             =   "ctl"
            Top             =   1692
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Diplomacy"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   21
            Left            =   0
            TabIndex        =   81
            Tag             =   "ctl"
            Top             =   1452
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Bluff"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   20
            Left            =   0
            TabIndex        =   79
            Tag             =   "ctl"
            Top             =   1212
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Balance"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   16
            Left            =   0
            TabIndex        =   71
            Tag             =   "ctl"
            Top             =   12
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Jump"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   17
            Left            =   0
            TabIndex        =   73
            Tag             =   "ctl"
            Top             =   252
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Swim"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   18
            Left            =   0
            TabIndex        =   75
            Tag             =   "ctl"
            Top             =   492
            Width           =   1044
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tumble"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   19
            Left            =   0
            TabIndex        =   77
            Tag             =   "ctl"
            Top             =   732
            Width           =   1044
         End
      End
      Begin VB.Frame fraStatTomes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1632
         Left            =   180
         TabIndex        =   34
         Tag             =   "ctl"
         Top             =   300
         Width           =   2232
         Begin Compendium.userSpinner usrspnStat 
            Height          =   252
            Index           =   1
            Left            =   1260
            TabIndex        =   36
            Top             =   0
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   8
            Value           =   0
            StepLarge       =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnStat 
            Height          =   252
            Index           =   2
            Left            =   1260
            TabIndex        =   38
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   8
            Value           =   0
            StepLarge       =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnStat 
            Height          =   252
            Index           =   3
            Left            =   1260
            TabIndex        =   40
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   8
            Value           =   0
            StepLarge       =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnStat 
            Height          =   252
            Index           =   4
            Left            =   1260
            TabIndex        =   42
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   8
            Value           =   0
            StepLarge       =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnStat 
            Height          =   252
            Index           =   5
            Left            =   1260
            TabIndex        =   44
            Top             =   960
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   8
            Value           =   0
            StepLarge       =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnStat 
            Height          =   252
            Index           =   6
            Left            =   1260
            TabIndex        =   46
            Top             =   1200
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   8
            Value           =   0
            StepLarge       =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Charisma"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   14
            Left            =   0
            TabIndex        =   45
            Tag             =   "ctl"
            Top             =   1212
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Wisdom"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   13
            Left            =   0
            TabIndex        =   43
            Tag             =   "ctl"
            Top             =   972
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Intelligence"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   12
            Left            =   0
            TabIndex        =   41
            Tag             =   "ctl"
            Top             =   732
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Constitution"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   11
            Left            =   0
            TabIndex        =   39
            Tag             =   "ctl"
            Top             =   492
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Dexterity"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   10
            Left            =   0
            TabIndex        =   37
            Tag             =   "ctl"
            Top             =   252
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Strength"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   0
            TabIndex        =   35
            Tag             =   "ctl"
            Top             =   12
            Width           =   1140
         End
      End
      Begin VB.Shape shpTab 
         Height          =   1332
         Index           =   1
         Left            =   8220
         Top             =   120
         Width           =   1392
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Index           =   2
      Left            =   360
      ScaleHeight     =   3975
      ScaleWidth      =   9735
      TabIndex        =   115
      TabStop         =   0   'False
      Tag             =   "ctl"
      Top             =   4740
      Visible         =   0   'False
      Width           =   9732
      Begin VB.Frame fraPastLives 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3675
         Left            =   120
         TabIndex        =   116
         Tag             =   "ctl"
         Top             =   120
         Width           =   2055
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   12
            Left            =   1080
            TabIndex        =   106
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   119
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   121
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   3
            Left            =   1080
            TabIndex        =   123
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   13
            Left            =   1080
            TabIndex        =   125
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   11
            Left            =   1080
            TabIndex        =   127
            Top             =   1440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   129
            Top             =   1680
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   10
            Left            =   1080
            TabIndex        =   131
            Top             =   1920
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   5
            Left            =   1080
            TabIndex        =   133
            Top             =   2160
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   135
            Top             =   2400
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   7
            Left            =   1080
            TabIndex        =   137
            Top             =   2640
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   8
            Left            =   1080
            TabIndex        =   139
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   14
            Left            =   1080
            TabIndex        =   141
            Top             =   3120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   9
            Left            =   1080
            TabIndex        =   143
            Top             =   3360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnClass 
            Height          =   255
            Index           =   15
            Left            =   1080
            TabIndex        =   105
            Top             =   0
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Alchemist"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   93
            Left            =   0
            TabIndex        =   242
            Tag             =   "ctl"
            Top             =   0
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Wizard"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   50
            Left            =   0
            TabIndex        =   142
            Tag             =   "ctl"
            Top             =   3375
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Warlock"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   49
            Left            =   0
            TabIndex        =   140
            Tag             =   "ctl"
            Top             =   3135
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sorcerer"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   48
            Left            =   0
            TabIndex        =   138
            Tag             =   "ctl"
            Top             =   2895
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Rogue"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   47
            Left            =   0
            TabIndex        =   136
            Tag             =   "ctl"
            Top             =   2655
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ranger"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   46
            Left            =   0
            TabIndex        =   134
            Tag             =   "ctl"
            Top             =   2415
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Paladin"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   45
            Left            =   0
            TabIndex        =   132
            Tag             =   "ctl"
            Top             =   2175
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Monk"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   44
            Left            =   0
            TabIndex        =   130
            Tag             =   "ctl"
            Top             =   1935
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fighter"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   43
            Left            =   0
            TabIndex        =   128
            Tag             =   "ctl"
            Top             =   1695
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fav Soul"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   42
            Left            =   0
            TabIndex        =   126
            Tag             =   "ctl"
            Top             =   1455
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Druid"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   41
            Left            =   0
            TabIndex        =   124
            Tag             =   "ctl"
            Top             =   1215
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cleric"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   40
            Left            =   0
            TabIndex        =   122
            Tag             =   "ctl"
            Top             =   975
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Bard"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   39
            Left            =   0
            TabIndex        =   120
            Tag             =   "ctl"
            Top             =   735
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Barbarian"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   38
            Left            =   0
            TabIndex        =   118
            Tag             =   "ctl"
            Top             =   495
            Width           =   960
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Artificer"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   37
            Left            =   0
            TabIndex        =   117
            Tag             =   "ctl"
            Top             =   255
            Width           =   960
         End
      End
      Begin VB.Frame fraEpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   3
         Left            =   4680
         TabIndex        =   189
         Tag             =   "ctl"
         Top             =   2460
         Width           =   2352
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   255
            Index           =   1
            Left            =   1380
            TabIndex        =   193
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   2
            Left            =   1380
            TabIndex        =   195
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   3
            Left            =   1380
            TabIndex        =   197
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   255
            Index           =   13
            Left            =   1380
            TabIndex        =   199
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   255
            Index           =   17
            Left            =   1380
            TabIndex        =   201
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Eclipse"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   102
            Left            =   0
            TabIndex        =   243
            Tag             =   "ctl"
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Knowlege"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   96
            Left            =   0
            TabIndex        =   198
            Tag             =   "ctl"
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Arcane"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   68
            Left            =   1380
            TabIndex        =   191
            Tag             =   "ctl"
            Top             =   0
            Width           =   840
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Alacrity"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   69
            Left            =   0
            TabIndex        =   192
            Tag             =   "ctl"
            Top             =   252
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Criticals"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   70
            Left            =   0
            TabIndex        =   194
            Tag             =   "ctl"
            Top             =   492
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Enchant"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   71
            Left            =   0
            TabIndex        =   196
            Tag             =   "ctl"
            Top             =   732
            Width           =   1260
         End
      End
      Begin VB.Frame fraEpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   2
         Left            =   7200
         TabIndex        =   221
         Tag             =   "ctl"
         Top             =   2700
         Width           =   2352
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   10
            Left            =   1380
            TabIndex        =   224
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   11
            Left            =   1380
            TabIndex        =   226
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   12
            Left            =   1380
            TabIndex        =   228
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   255
            Index           =   16
            Left            =   1380
            TabIndex        =   230
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Power"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   99
            Left            =   0
            TabIndex        =   229
            Tag             =   "ctl"
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Primal"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   80
            Left            =   1380
            TabIndex        =   222
            Tag             =   "ctl"
            Top             =   0
            Width           =   840
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Doubleshot"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   81
            Left            =   0
            TabIndex        =   223
            Tag             =   "ctl"
            Top             =   252
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fast Healing"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   82
            Left            =   0
            TabIndex        =   225
            Tag             =   "ctl"
            Top             =   492
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Queen"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   83
            Left            =   0
            TabIndex        =   227
            Tag             =   "ctl"
            Top             =   732
            Width           =   1260
         End
      End
      Begin VB.Frame fraEpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   1
         Left            =   7200
         TabIndex        =   211
         Tag             =   "ctl"
         Top             =   1440
         Width           =   2352
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   7
            Left            =   1380
            TabIndex        =   214
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   8
            Left            =   1380
            TabIndex        =   216
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   9
            Left            =   1380
            TabIndex        =   218
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   255
            Index           =   15
            Left            =   1380
            TabIndex        =   220
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tactics"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   98
            Left            =   0
            TabIndex        =   219
            Tag             =   "ctl"
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Martial"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   76
            Left            =   1380
            TabIndex        =   212
            Tag             =   "ctl"
            Top             =   0
            Width           =   840
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Doublestrike"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   77
            Left            =   0
            TabIndex        =   213
            Tag             =   "ctl"
            Top             =   252
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Fortification"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   78
            Left            =   0
            TabIndex        =   215
            Tag             =   "ctl"
            Top             =   492
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Skill Mastery"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   79
            Left            =   0
            TabIndex        =   217
            Tag             =   "ctl"
            Top             =   732
            Width           =   1260
         End
      End
      Begin VB.Frame fraEpic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   7200
         TabIndex        =   200
         Tag             =   "ctl"
         Top             =   180
         Width           =   2352
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   4
            Left            =   1380
            TabIndex        =   204
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   5
            Left            =   1380
            TabIndex        =   206
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   252
            Index           =   6
            Left            =   1380
            TabIndex        =   208
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnEpic 
            Height          =   255
            Index           =   14
            Left            =   1380
            TabIndex        =   210
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Blessings"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   97
            Left            =   0
            TabIndex        =   209
            Tag             =   "ctl"
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Block Energy"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   75
            Left            =   0
            TabIndex        =   207
            Tag             =   "ctl"
            Top             =   732
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Life/Death"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   74
            Left            =   0
            TabIndex        =   205
            Tag             =   "ctl"
            Top             =   492
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Brace"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   73
            Left            =   0
            TabIndex        =   203
            Tag             =   "ctl"
            Top             =   252
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Divine"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   72
            Left            =   1380
            TabIndex        =   202
            Tag             =   "ctl"
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Frame fraIconic 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2235
         Left            =   4680
         TabIndex        =   171
         Tag             =   "ctl"
         Top             =   180
         Width           =   2352
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   252
            Index           =   1
            Left            =   1380
            TabIndex        =   174
            Top             =   0
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   252
            Index           =   2
            Left            =   1380
            TabIndex        =   176
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   252
            Index           =   3
            Left            =   1380
            TabIndex        =   178
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   252
            Index           =   4
            Left            =   1380
            TabIndex        =   180
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   255
            Index           =   5
            Left            =   1380
            TabIndex        =   184
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   255
            Index           =   6
            Left            =   1380
            TabIndex        =   186
            Top             =   1440
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   255
            Index           =   7
            Left            =   1380
            TabIndex        =   188
            Top             =   1680
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   255
            Index           =   8
            Left            =   1380
            TabIndex        =   182
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnIconic 
            Height          =   255
            Index           =   9
            Left            =   1380
            TabIndex        =   190
            Top             =   1920
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Trailblazer"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   104
            Left            =   0
            TabIndex        =   245
            Tag             =   "ctl"
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Scoundrel"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   101
            Left            =   0
            TabIndex        =   181
            Tag             =   "ctl"
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Razorclaw"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   95
            Left            =   0
            TabIndex        =   187
            Tag             =   "ctl"
            Top             =   1680
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Shadar-kai"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   67
            Left            =   0
            TabIndex        =   185
            Tag             =   "ctl"
            Top             =   1455
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Scourge"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   66
            Left            =   0
            TabIndex        =   183
            Tag             =   "ctl"
            Top             =   1215
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "PDK"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   65
            Left            =   0
            TabIndex        =   179
            Tag             =   "ctl"
            Top             =   732
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Morninglord"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   64
            Left            =   0
            TabIndex        =   177
            Tag             =   "ctl"
            Top             =   492
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Deep Gnome"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   63
            Left            =   0
            TabIndex        =   175
            Tag             =   "ctl"
            Top             =   252
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Bladeforged"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   62
            Left            =   0
            TabIndex        =   173
            Tag             =   "ctl"
            Top             =   12
            Width           =   1260
         End
      End
      Begin VB.Frame fraRacial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3432
         Left            =   2280
         TabIndex        =   144
         Tag             =   "ctl"
         Top             =   180
         Width           =   2292
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   1
            Left            =   1320
            TabIndex        =   146
            Top             =   0
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   1
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   2
            Left            =   1320
            TabIndex        =   148
            Top             =   240
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   3
            Left            =   1320
            TabIndex        =   150
            Top             =   480
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   4
            Left            =   1320
            TabIndex        =   152
            Top             =   720
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   5
            Left            =   1320
            TabIndex        =   154
            Top             =   960
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   6
            Left            =   1320
            TabIndex        =   156
            Top             =   1200
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   7
            Left            =   1320
            TabIndex        =   158
            Top             =   1440
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   8
            Left            =   1320
            TabIndex        =   160
            Top             =   1680
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   9
            Left            =   1320
            TabIndex        =   162
            Top             =   1920
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   252
            Index           =   10
            Left            =   1320
            TabIndex        =   164
            Top             =   2160
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   255
            Index           =   11
            Left            =   1320
            TabIndex        =   168
            Top             =   2640
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   255
            Index           =   12
            Left            =   1320
            TabIndex        =   170
            Top             =   2880
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   255
            Index           =   13
            Left            =   1320
            TabIndex        =   166
            Top             =   2400
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   2
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Compendium.userSpinner usrspnRace 
            Height          =   255
            Index           =   14
            Left            =   1320
            TabIndex        =   172
            Top             =   3120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            Min             =   0
            Max             =   3
            Value           =   0
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   3
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tabaxi"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   103
            Left            =   0
            TabIndex        =   244
            Tag             =   "ctl"
            Top             =   3120
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Tiefling"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   100
            Left            =   0
            TabIndex        =   165
            Tag             =   "ctl"
            Top             =   2400
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Shifter"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   94
            Left            =   0
            TabIndex        =   169
            Tag             =   "ctl"
            Top             =   2880
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Warforged"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   61
            Left            =   0
            TabIndex        =   167
            Tag             =   "ctl"
            Top             =   2655
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Human"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   60
            Left            =   0
            TabIndex        =   163
            Tag             =   "ctl"
            Top             =   2172
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Half-Orc"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   59
            Left            =   0
            TabIndex        =   161
            Tag             =   "ctl"
            Top             =   1932
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Half-Elf"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   58
            Left            =   0
            TabIndex        =   159
            Tag             =   "ctl"
            Top             =   1692
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Halfling"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   57
            Left            =   0
            TabIndex        =   157
            Tag             =   "ctl"
            Top             =   1452
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Gnome"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   56
            Left            =   0
            TabIndex        =   155
            Tag             =   "ctl"
            Top             =   1212
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Elf"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   55
            Left            =   0
            TabIndex        =   153
            Tag             =   "ctl"
            Top             =   972
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Dwarf"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   54
            Left            =   0
            TabIndex        =   151
            Tag             =   "ctl"
            Top             =   732
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Drow"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   53
            Left            =   0
            TabIndex        =   149
            Tag             =   "ctl"
            Top             =   492
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Dragonborn"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   52
            Left            =   0
            TabIndex        =   147
            Tag             =   "ctl"
            Top             =   252
            Width           =   1200
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Aasimar"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   51
            Left            =   0
            TabIndex        =   145
            Tag             =   "ctl"
            Top             =   12
            Width           =   1200
         End
      End
      Begin VB.Shape shpTab 
         Height          =   972
         Index           =   2
         Left            =   8940
         Top             =   60
         Width           =   672
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   0
      Left            =   360
      ScaleHeight     =   3855
      ScaleWidth      =   9735
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "ctl"
      Top             =   4740
      Width           =   9732
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1692
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   600
         Width           =   4572
      End
      Begin VB.Shape shpTab 
         Height          =   2352
         Index           =   0
         Left            =   360
         Top             =   300
         Width           =   5892
      End
   End
   Begin VB.Label lblTab 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Insert tab characters with Ctrl+Tab"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4020
      TabIndex        =   30
      Top             =   4500
      Width           =   5700
   End
   Begin VB.Image imgArrow 
      Height          =   312
      Index           =   1
      Left            =   1920
      Picture         =   "frmCharacter.frx":0091
      Stretch         =   -1  'True
      ToolTipText     =   "HIgher Priority"
      Top             =   540
      Width           =   300
   End
   Begin VB.Image imgArrow 
      Height          =   312
      Index           =   2
      Left            =   1920
      Picture         =   "frmCharacter.frx":088B
      Stretch         =   -1  'True
      ToolTipText     =   "Lower Priority"
      Top             =   900
      Width           =   300
   End
   Begin VB.Image imgArrow 
      Height          =   312
      Index           =   0
      Left            =   1920
      Picture         =   "frmCharacter.frx":1085
      Stretch         =   -1  'True
      ToolTipText     =   "Clear All"
      Top             =   1620
      Width           =   300
   End
   Begin VB.Label lblLabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Characters"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   972
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuExport 
         Caption         =   "Character Builder Lite"
         Index           =   0
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Ron's Character Planner"
         Index           =   1
      End
      Begin VB.Menu mnuExport 
         Caption         =   "DDO Builder"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngCharacter As Long

Private mblnOverride As Boolean
Private mblnDirty As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    LoadData
    ShowCharacter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
    If mblnDirty Then
        Select Case MsgBox("Save changes to character " & Me.txtName.Text & "?", vbYesNoCancel + vbQuestion, "Notice")
            Case vbYes
                Cancel = SaveChanges()
            Case vbCancel
                Cancel = 1
                Exit Sub
        End Select
    End If
    mlngCharacter = 0
End Sub

Public Property Let Character(plngCharacter As Long)
    mlngCharacter = plngCharacter
End Property

Public Sub FavorTotals()
    If mlngCharacter = 0 Then
        Me.txtChallengeFavor.Text = vbNullString
        Me.txtTotalFavor.Text = vbNullString
    Else
        With db.Character(mlngCharacter)
            Me.txtChallengeFavor.Text = .ChallengeFavor
            Me.txtTotalFavor.Text = .TotalFavor
        End With
    End If
End Sub

Public Sub DataFileChanged()
    mlngCharacter = 0
    LoadData
    ShowCharacter
End Sub


' ************* DATA *************


Private Sub LoadData()
    Dim i As Long
    
    mblnOverride = True
    ListboxClear Me.lstCharacter
    For i = 1 To db.Characters
        ListboxAddItem Me.lstCharacter, db.Character(i).Character, i
    Next
    ComboClear Me.cboColor
    For i = 0 To gceColors - 1
        ComboAddItem Me.cboColor, GetColorName(i), i
    Next
    Me.usrTab.Width = Me.usrTab.TabsWidth
    For i = Me.shpTab.LBound To Me.shpTab.UBound
        With Me.picTab(i)
            Me.shpTab(i).Move 0, 0, .ScaleWidth, .ScaleHeight
        End With
    Next
    With Me.picTab(0)
        Me.txtNotes.Move PixelX, PixelY, .ScaleWidth - PixelX * 2, .ScaleHeight - PixelY * 2
    End With
    ' Tomes
    For i = 1 To 6
        Me.usrspnStat(i).Max = tomes.Stat.Max
    Next
    For i = 1 To 21
        Me.usrspnSkill(i).Max = tomes.Skill.Max
    Next
    For i = 1 To 3
        Me.usrspnPower(i).Max = tomes.PowerMax
    Next
    Me.usrspnRacialAP.Max = tomes.RacialAPMax
    Me.usrspnUniversalAP.Max = tomes.UniversalAPMax
    Me.usrspnFate.Max = tomes.FateTomeMax
    Me.usrspnDestiny.Max = tomes.DestinyTomeMax
    Me.usrspnRR(1).Max = tomes.RRMax
    Me.usrspnRR(2).Max = tomes.RRMax
    mblnOverride = False
End Sub

Private Sub ShowCharacter()
    Dim typCharacter As CharacterType
    Dim blnEnabled As Boolean
    Dim i As Long
    
    mblnOverride = True
    blnEnabled = (mlngCharacter > 0)
    For i = 1 To Me.lblLabel.UBound
        Me.lblLabel(i).Enabled = blnEnabled
    Next
    If blnEnabled Then typCharacter = db.Character(mlngCharacter)
    With typCharacter
        ' List
        Me.lblLabel(0).Enabled = (db.Characters > 0)
        Me.lstCharacter.Enabled = (db.Characters > 0)
        If blnEnabled Then Me.lstCharacter.ListIndex = mlngCharacter - 1
        ' Arrows
        EnableArrows
        ' Export
        Me.chkButton(8).Enabled = blnEnabled
        ' Name
        Me.txtName.Text = .Character
        Me.txtName.Enabled = blnEnabled
        ' Context menu
        Me.chkContext.Enabled = blnEnabled
        ' LeftClick
        ComboClear Me.cboLeftClick
        ComboAddItem Me.cboLeftClick, "Edit Characters", 0
        With .ContextMenu
            For i = 1 To .Commands
                If .Command(i).Style <> mceSeparator Then ComboAddItem Me.cboLeftClick, .Command(i).Caption, i
            Next
        End With
        ComboSetText Me.cboLeftClick, .LeftClick
        If Me.cboLeftClick.ListIndex = -1 Then Me.cboLeftClick.ListIndex = 0
        Me.cboLeftClick.Enabled = blnEnabled
        ' Favor
        Me.txtChallengeFavor.Text = .ChallengeFavor
        Me.txtTotalFavor.Text = .TotalFavor
        ' Generated
        ComboSetValue Me.cboColor, .GeneratedColor
        Me.cboColor.Enabled = blnEnabled
        ' Custom
        Me.usrchkCustomColors.Value = .CustomColor
        Me.usrchkCustomColors.Enabled = blnEnabled
        Me.picColor(0).BackColor = .BackColor
        Me.picColor(1).BackColor = .DimColor
        ' Notes
        Me.txtNotes.Text = .Notes
        Me.txtNotes.Enabled = blnEnabled
        ' Tomes
        For i = 1 To UBound(.Tome.Stat)
            Me.usrspnStat(i).Value = .Tome.Stat(i)
            Me.usrspnStat(i).Enabled = blnEnabled
        Next
        For i = 1 To UBound(.Tome.Skill)
            Me.usrspnSkill(i).Value = .Tome.Skill(i)
            Me.usrspnSkill(i).Enabled = blnEnabled
        Next
        ComboSetText Me.cboHeroicXP, .Tome.HeroicXP
        Me.cboHeroicXP.Enabled = blnEnabled
        ComboSetText Me.cboEpicXP, .Tome.EpicXP
        Me.cboEpicXP.Enabled = blnEnabled
        Me.usrspnRacialAP.Value = .Tome.RacialAP
        Me.usrspnRacialAP.Enabled = blnEnabled
        Me.usrspnUniversalAP.Value = .Tome.UniversalAP
        Me.usrspnUniversalAP.Enabled = blnEnabled
        Me.usrspnFate.Value = .Tome.Fate
        Me.usrspnFate.Enabled = blnEnabled
        Me.usrspnDestiny.Value = .Tome.Destiny
        Me.usrspnDestiny.Enabled = blnEnabled
        For i = 1 To UBound(.Tome.Power)
            Me.usrspnPower(i).Value = .Tome.Power(i)
            Me.usrspnPower(i).Enabled = blnEnabled
        Next
        For i = 1 To UBound(.Tome.RR)
            Me.usrspnRR(i).Value = .Tome.RR(i)
            Me.usrspnRR(i).Enabled = blnEnabled
        Next
        ' Past Lives
        For i = 1 To UBound(.PastLife.Class)
            Me.usrspnClass(i).Value = .PastLife.Class(i)
            Me.usrspnClass(i).Enabled = blnEnabled
        Next
        For i = 1 To UBound(.PastLife.Racial)
            Me.usrspnRace(i).Value = .PastLife.Racial(i)
            Me.usrspnRace(i).Enabled = blnEnabled
        Next
        For i = 1 To UBound(.PastLife.Iconic)
            Me.usrspnIconic(i).Value = .PastLife.Iconic(i)
            Me.usrspnIconic(i).Enabled = blnEnabled
        Next
        For i = 1 To UBound(.PastLife.Epic)
            Me.usrspnEpic(i).Value = .PastLife.Epic(i)
            Me.usrspnEpic(i).Enabled = blnEnabled
        Next
    End With
    mblnOverride = False
    mblnDirty = False
End Sub

Private Function DefaultCharacter() As CharacterType
    Dim i As Long
    
    With DefaultCharacter
        .Character = NewName()
        .CustomColor = False
        .GeneratedColor = NewColor()
        .BackColor = GetColorValue(.GeneratedColor)
        .DimColor = GetColorDim(.BackColor)
        If db.Sagas Then
            ReDim .Saga(1 To db.Sagas)
            For i = 1 To db.Sagas
                ReDim .Saga(i).Progress(1 To db.Saga(i).Quests)
            Next
        End If
    End With
End Function

Private Function NewName() As String
    Dim lngNew As Long
    Dim blnFound As Boolean
    Dim i As Long
    
    Do
        blnFound = False
        lngNew = lngNew + 1
        For i = 1 To db.Characters
            If db.Character(i).Character = "New" & lngNew Then
                blnFound = True
                Exit For
            End If
        Next
    Loop Until Not blnFound
    NewName = "New" & lngNew
End Function

Private Function NewColor() As GeneratedColorEnum
    Dim blnTaken() As Boolean
    Dim lngNew As Long
    Dim i As Long
    
    lngNew = RandomNumber(gceColors - 1)
    If db.Characters < gceColors - 1 Then
        ReDim blnTaken(gceColors)
        For i = 1 To db.Characters
            blnTaken(db.Character(i).GeneratedColor) = True
        Next
        Do While blnTaken(lngNew)
            lngNew = RandomNumber(gceColors - 1)
        Loop
    End If
    NewColor = lngNew
End Function

Private Function SaveChanges() As Integer
    Dim frm As Form
    Dim blnRenamed As Boolean
    Dim blnColorChanged As Boolean
    Dim i As Long
    
    If Not mblnDirty Then Exit Function
    For i = 1 To db.Characters
        If i <> mlngCharacter And db.Character(i).Character = Me.txtName.Text Then
            MsgBox Me.txtName.Text & " already exists.", vbInformation, "Notice"
            SaveChanges = 1
            Exit Function
        End If
    Next
    With db.Character(mlngCharacter)
        If .Character <> Me.txtName.Text Then blnRenamed = True
        If .CustomColor <> Me.usrchkCustomColors.Value Then
            blnColorChanged = True
        ElseIf Me.usrchkCustomColors.Value = True And (.BackColor <> Me.picColor(0).BackColor Or .DimColor <> Me.picColor(1).BackColor) Then
            blnColorChanged = True
        ElseIf Me.usrchkCustomColors.Value = False And .GeneratedColor <> ComboGetValue(Me.cboColor) Then
            blnColorChanged = True
        End If
        .Character = Me.txtName.Text
        .LeftClick = ComboGetText(Me.cboLeftClick)
        .GeneratedColor = ComboGetValue(Me.cboColor)
        .CustomColor = Me.usrchkCustomColors.Value
        .BackColor = Me.picColor(0).BackColor
        .DimColor = Me.picColor(1).BackColor
        .Notes = Me.txtNotes.Text
        For i = 1 To UBound(.Tome.Stat)
            .Tome.Stat(i) = Me.usrspnStat(i).Value
        Next
        For i = 1 To UBound(.Tome.Skill)
            .Tome.Skill(i) = Me.usrspnSkill(i).Value
        Next
        .Tome.HeroicXP = Me.cboHeroicXP.Text
        .Tome.EpicXP = Me.cboEpicXP.Text
        .Tome.RacialAP = Me.usrspnRacialAP.Value
        .Tome.UniversalAP = Me.usrspnUniversalAP.Value
        .Tome.Fate = Me.usrspnFate.Value
        .Tome.Destiny = Me.usrspnDestiny.Value
        For i = 1 To UBound(.Tome.Power)
            .Tome.Power(i) = Me.usrspnPower(i).Value
        Next
        For i = 1 To UBound(.Tome.RR)
            .Tome.RR(i) = Me.usrspnRR(i).Value
        Next
        For i = 1 To UBound(.PastLife.Class)
            .PastLife.Class(i) = Me.usrspnClass(i).Value
        Next
        For i = 1 To UBound(.PastLife.Racial)
            .PastLife.Racial(i) = Me.usrspnRace(i).Value
        Next
        For i = 1 To UBound(.PastLife.Iconic)
            .PastLife.Iconic(i) = Me.usrspnIconic(i).Value
        Next
        For i = 1 To UBound(.PastLife.Epic)
            .PastLife.Epic(i) = Me.usrspnEpic(i).Value
        Next
    End With
    With Me.lstCharacter
        If .ListIndex <> -1 Then .List(.ListIndex) = Me.txtName.Text
    End With
    ShowChanges blnRenamed, blnColorChanged
    Dirty = False
    DirtyFlag dfeData
End Function

Private Sub ShowChanges(pblnRenamed As Boolean, pblnColorChanged As Boolean)
    Dim frm As Form
    
    ' Compendium and Patrons
    If pblnRenamed Or pblnColorChanged Then
        frmCompendium.RedrawQuests
        If GetForm(frm, "frmPatrons") Then frm.ReDrawForm
    End If
    ' Challenges
    If pblnRenamed Then
        If GetForm(frm, "frmChallenges") Then
            frm.CharacterListChanged
            frm.Character = mlngCharacter
        End If
    End If
    ' Sagas
    If pblnRenamed Or pblnColorChanged Then
        If GetForm(frm, "frmSagas") Then
            frm.Redraw
            If pblnRenamed Then frm.Character = mlngCharacter
        End If
    End If
End Sub


' ************* CHARACTER LIST *************


Private Sub CharacterListChanged()
    Dim frm As Form
    
    frmCompendium.RedrawQuests
    DirtyFlag dfeData
    If GetForm(frm, "frmPatrons") Then frm.ReDrawForm
    If GetForm(frm, "frmChallenges") Then frm.CharacterListChanged
    If GetForm(frm, "frmSagas") Then frm.CharacterListChanged
'    CharacterChanged
End Sub

Private Sub CharacterChanged()
    Dim frm As Form
    
    cfg.Character = mlngCharacter
    If GetForm(frm, "frmChallenges") Then frm.Character = mlngCharacter
    If GetForm(frm, "frmSagas") Then frm.Character = mlngCharacter
    frmCompendium.CharacterChanged
End Sub



Private Sub lstCharacter_Click()
    If mblnOverride Then Exit Sub
    mlngCharacter = Me.lstCharacter.ListIndex + 1
    If mlngCharacter = 0 Then Exit Sub
    ShowCharacter
    CharacterChanged
End Sub

Private Sub EnableArrows(Optional pblnDisableAll As Boolean)
    With Me.lstCharacter
        EnableArrow Me.imgArrow(0), (.ListIndex <> -1) And Not pblnDisableAll
        EnableArrow Me.imgArrow(1), (.ListIndex > 0) And Not pblnDisableAll
        EnableArrow Me.imgArrow(2), (.ListIndex > -1 And .ListIndex < .ListCount - 1) And Not pblnDisableAll
    End With
End Sub

Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetArrowIcon Me.imgArrow(Index), asePressed
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetArrowIcon Me.imgArrow(Index), aseEnabled
    If Index = 0 Then
        If MsgBox("Delete " & db.Character(mlngCharacter).Character & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") = vbYes Then DeleteCharacter
    Else
        If Dirty Then
            MsgBox "Save changes before reordering.", vbInformation, "Notice"
            Exit Sub
        End If
        Select Case Index
            Case 1: MoveCharacter -1
            Case 2: MoveCharacter 1
        End Select
    End If
End Sub

Private Sub DeleteCharacter()
    Dim lngIndex As Long
    Dim i As Long
    Dim c As Long
    
    If mlngCharacter = 0 Then Exit Sub
    Dirty = False
    lngIndex = Me.lstCharacter.ListIndex
    For i = 1 To db.Quests
        With db.Quest(i)
            If db.Characters - 1 = 0 Then
                Erase .Progress
            Else
                For c = mlngCharacter To db.Characters - 1
                    .Progress(c) = .Progress(c + 1)
                Next
                ReDim Preserve .Progress(1 To db.Characters - 1)
            End If
        End With
    Next
    For i = 1 To db.Challenges
        With db.Challenge(i)
            If db.Characters - 1 = 0 Then
                Erase .Stars
            Else
                For c = mlngCharacter To db.Characters - 1
                    .Stars(c) = .Stars(c + 1)
                Next
                ReDim Preserve .Stars(1 To db.Characters - 1)
            End If
        End With
    Next
    db.Characters = db.Characters - 1
    If db.Characters = 0 Then
        Erase db.Character
        ListboxClear Me.lstCharacter
    Else
        For c = mlngCharacter To db.Characters
            db.Character(c) = db.Character(c + 1)
        Next
        ReDim Preserve db.Character(1 To db.Characters)
        Me.lstCharacter.RemoveItem lngIndex
        If lngIndex > Me.lstCharacter.ListCount - 1 Then lngIndex = Me.lstCharacter.ListCount - 1
        Me.lstCharacter.ListIndex = lngIndex
    End If
    CharacterListChanged
End Sub

Private Sub MoveCharacter(plngIncrement As Long)
    Dim lngOld As Long
    Dim lngNew As Long
    Dim typSwap As CharacterType
    Dim enSwap As ProgressEnum
    Dim lngSwap As Long
    Dim i As Long
    Dim c As Long
    
    lngOld = mlngCharacter
    lngNew = mlngCharacter + plngIncrement
    If lngNew < 1 Or lngNew > db.Characters Then Exit Sub
    For i = 1 To db.Quests
        With db.Quest(i)
            enSwap = .Progress(lngOld)
            .Progress(lngOld) = .Progress(lngNew)
            .Progress(lngNew) = enSwap
        End With
    Next
    For i = 1 To db.Challenges
        With db.Challenge(i)
            lngSwap = .Stars(lngOld)
            .Stars(lngOld) = .Stars(lngNew)
            .Stars(lngNew) = lngSwap
        End With
    Next
    typSwap = db.Character(lngOld)
    db.Character(lngOld) = db.Character(lngNew)
    db.Character(lngNew) = typSwap
    mblnOverride = True
    With Me.lstCharacter
        .List(lngOld - 1) = db.Character(lngOld).Character
        .List(lngNew - 1) = db.Character(lngNew).Character
        .ListIndex = lngNew - 1
    End With
    mblnOverride = False
    mlngCharacter = lngNew
    CharacterListChanged
    EnableArrows
End Sub


' ************* EDIT *************


Private Property Get Dirty() As Boolean
    Dirty = mblnDirty
End Property

Private Property Let Dirty(ByVal pblnDirty As Boolean)
    If mblnOverride = True Or mblnDirty = pblnDirty Then Exit Property
    mblnDirty = pblnDirty
    EnableArrow Me.imgArrow(1), Not mblnDirty
    EnableArrow Me.imgArrow(2), Not mblnDirty
    Me.lstCharacter.Enabled = Not mblnDirty
    Me.chkButton(0).Enabled = Not mblnDirty
    Me.chkButton(3).Enabled = mblnDirty
    Me.chkButton(5).Enabled = Not mblnDirty
    Me.chkButton(7).Enabled = Not mblnDirty
    Me.chkButton(8).Enabled = Not mblnDirty
End Property

Private Sub txtName_GotFocus()
    TextboxGotFocus Me.txtName
End Sub

Private Sub txtName_Change()
    Dirty = True
End Sub

Private Sub chkContext_Click()
    Dim i As Long
    
    If UncheckButton(Me.chkContext, mblnOverride) Then Exit Sub
    AutoSave
    gtypMenu = db.Character(mlngCharacter).ContextMenu
    gtypMenu.Title = db.Character(mlngCharacter).Character & " Context Menu"
    gtypMenu.Accepted = False
    gtypMenu.LinkList = False
    frmMenuEditor.Show vbModal, Me
    If gtypMenu.Accepted Then
        mblnOverride = True
        With db.Character(mlngCharacter)
            .ContextMenu = gtypMenu
            ComboClear Me.cboLeftClick
            With .ContextMenu
                ComboAddItem Me.cboLeftClick, "Edit Characters", 0
                For i = 1 To .Commands
                    If .Command(i).Style <> mceSeparator Then ComboAddItem Me.cboLeftClick, .Command(i).Caption, i
                Next
            End With
            ComboSetText Me.cboLeftClick, .LeftClick
            If Me.cboLeftClick.ListIndex = -1 Then Me.cboLeftClick.ListIndex = 0
        End With
        mblnOverride = False
    End If
End Sub

Private Sub cboLeftClick_Click()
    If mblnOverride Then Exit Sub
    Dirty = True
End Sub

Private Sub cboHeroic_Click()
    If mblnOverride Then Exit Sub
    Dirty = True
End Sub

Private Sub cboEpic_Click()
    If mblnOverride Then Exit Sub
    Dirty = True
End Sub

Private Sub usrchkCustomColors_UserChange()
    ColorChange
End Sub

Private Sub lnkLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    Select Case Me.lnkLink(Index).Caption
        Case "Named"
            OpenForm "frmColorPreview"
        Case "Challenges"
            frmChallenges.Character = mlngCharacter
            OpenForm "frmChallenges"
        Case "Total"
            OpenForm "frmPatrons"
    End Select
End Sub

Private Sub cboColor_Click()
    If mblnOverride Then Exit Sub
    ColorChange
End Sub

Private Sub picColor_Click(Index As Integer)
    glngActiveColor = Me.picColor(Index).BackColor
    frmColor.Show vbModal, Me
    If Me.picColor(Index).BackColor = glngActiveColor Then Exit Sub
    Me.picColor(Index).BackColor = glngActiveColor
    Me.usrchkCustomColors.Value = True
    If Index = 0 Then ColorChange Else Dirty = True
End Sub

Private Sub ColorChange()
    If Not Me.usrchkCustomColors.Value Then Me.picColor(0).BackColor = GetColorValue(ComboGetValue(Me.cboColor))
    Me.picColor(1).BackColor = GetColorDim(Me.picColor(0).BackColor)
    Dirty = True
End Sub

Private Sub txtNotes_Change()
    Dirty = True
End Sub


' ************* TABS *************


Private Sub usrTab_Click(pstrCaption As String)
    ShowTab pstrCaption
End Sub

Private Sub ShowTab(pstrCaption As String)
    Dim i As Long
    
    Me.lblTab.Visible = (pstrCaption = "Notes")
    For i = 0 To Me.picTab.UBound
        Me.picTab(i).Visible = (i = Me.usrTab.ActiveTabIndex)
    Next
End Sub

Private Sub usrspnStat_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnRacialAP_Change()
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnUniversalAP_Change()
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnFate_Change()
    If Not mblnOverride Then Dirty = True
End Sub
Private Sub usrspnDestiny_Change()
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnPower_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnRR_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub cboHeroicXP_Click()
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub cboEpicXP_Click()
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnSkill_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnClass_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnRace_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnIconic_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub

Private Sub usrspnEpic_Change(Index As Integer)
    If Not mblnOverride Then Dirty = True
End Sub


' ************* BUTTONS *************


Private Sub chkButton_Click(Index As Integer)
    If UncheckButton(Me.chkButton(Index), mblnOverride) Then Exit Sub
    Me.chkButton(Index).Refresh
    Select Case Me.chkButton(Index).Caption
        Case "Add New"
            AddNew
        Case "OK"
            If SaveChanges() = 0 Then Unload Me
        Case "Apply"
            SaveChanges
            Me.chkButton(1).SetFocus
        Case "Cancel"
            mblnDirty = False
            Unload Me
        Case "Help"
            ShowHelp "Characters"
        Case "Reincarnate"
            Reincarnate mlngCharacter
        Case "Sagas"
            frmSagas.Character = mlngCharacter
            OpenForm "frmSagas"
        Case "Import"
            ImportCharacters
        Case "Export"
            PopupMenu Me.mnuMain(0)
    End Select
End Sub

Private Sub AddNew()
    Dim i As Long
    
    db.Characters = db.Characters + 1
    ReDim Preserve db.Character(1 To db.Characters)
    db.Character(db.Characters) = DefaultCharacter()
    InitCharacterSagas db.Character(db.Characters)
    For i = 1 To db.Quests
        ReDim Preserve db.Quest(i).Progress(1 To db.Characters)
    Next
    For i = 1 To db.Challenges
        ReDim Preserve db.Challenge(i).Stars(1 To db.Characters)
    Next
    mlngCharacter = db.Characters
    ListboxAddItem Me.lstCharacter, db.Character(db.Characters).Character, db.Characters
    ShowCharacter
    CharacterListChanged
    Me.txtName.SetFocus
End Sub


' ************* IMPORT *************


Private Sub ImportCharacters()
    Dim strFile As String
    Dim i As Long
    
    strFile = xp.ShowOpenDialog(cfg.CompendiumPath, "Compendium Files|*.compendium", "*.compendium")
    If Len(strFile) = 0 Then Exit Sub
    If strFile = CompendiumFile() Then
        Notice "Cannot import from active compendium file"
        Exit Sub
    End If
    EnableImport True
    ImportCompendium strFile
    ListboxClear Me.lstImport
    For i = 1 To idb.Characters
        ListboxAddItem Me.lstImport, idb.Character(i).Character, i
    Next
    If Me.lstImport.ListCount > 0 Then
        Me.lstImport.SetFocus
        Me.lstImport.ListIndex = 0
    Else
        Me.chkImport(1).SetFocus
    End If
End Sub

Private Sub EnableImport(pblnEnabled As Boolean)
    Dim i As Long
    
    Me.fraImport.ZOrder vbBringToFront
    Me.fraImport.Visible = pblnEnabled
    Me.fraColors.Enabled = Not pblnEnabled
    Me.chkButton(0).Enabled = Not pblnEnabled
    Me.chkButton(5).Enabled = Not pblnEnabled
    Me.chkButton(7).Enabled = Not pblnEnabled
    Me.chkButton(8).Enabled = Not pblnEnabled
    For i = 0 To 2
        Me.picTab(i).Enabled = Not pblnEnabled
    Next
    EnableArrows pblnEnabled
End Sub

Private Sub chkImport_Click(Index As Integer)
    If UncheckButton(Me.chkImport(Index), mblnOverride) Then Exit Sub
    Select Case Me.chkImport(Index).Caption
        Case "Import"
            ImportCharacter
        Case "Done", "Finished"
            ClearImportData
            EnableImport False
    End Select
End Sub

Private Sub lstImport_DblClick()
    ImportCharacter
End Sub

Private Sub ImportCharacter()
    Dim lngCharacter As Long
    Dim strCharacter As String
    Dim i As Long
    
    lngCharacter = ListboxGetValue(Me.lstImport)
    If lngCharacter = 0 Then
        Notice "Select a character to import."
        Exit Sub
    End If
    strCharacter = idb.Character(lngCharacter).Character
    For i = 1 To db.Characters
        If db.Character(i).Character = strCharacter Then
            Notice strCharacter & " already exists"
            Exit Sub
        End If
    Next
    db.Characters = db.Characters + 1
    ReDim Preserve db.Character(1 To db.Characters)
    db.Character(db.Characters) = idb.Character(lngCharacter)
    ' Quests
    For i = 1 To db.Quests
        With db.Quest(i)
            ReDim Preserve .Progress(1 To db.Characters)
            .Progress(db.Characters) = idb.Quest(i).Progress(lngCharacter)
        End With
    Next
    ' Challenges
    For i = 1 To db.Challenges
        With db.Challenge(i)
            ReDim Preserve .Stars(1 To db.Characters)
            .Stars(db.Characters) = idb.Challenge(i).Stars(lngCharacter)
        End With
    Next
    Me.lstImport.RemoveItem Me.lstImport.ListIndex
    ListboxAddItem Me.lstCharacter, strCharacter, db.Characters
    CharacterListChanged
    Me.lstCharacter.ListIndex = Me.lstCharacter.NewIndex
End Sub


' ************* EXPORT *************


Private Sub mnuExport_Click(Index As Integer)
    Select Case Index
        Case 0: ExportCharacterLite mlngCharacter
        Case 1: ExportCharacterRon mlngCharacter
        Case 2: ExportDDOBuilder mlngCharacter
    End Select
End Sub
