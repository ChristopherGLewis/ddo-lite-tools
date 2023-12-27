VERSION 5.00
Begin VB.Form frmGearset 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Gearset"
   ClientHeight    =   9015
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   13560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGearset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userAugment usrAugPicker 
      Height          =   4512
      Left            =   4440
      TabIndex        =   87
      Top             =   2400
      Visible         =   0   'False
      Width           =   2832
      _ExtentX        =   5001
      _ExtentY        =   7779
      VariantOnly     =   -1  'True
   End
   Begin CannithCrafting.userAugSlots usrAugSlot 
      Height          =   2592
      Left            =   720
      TabIndex        =   86
      Top             =   3900
      Visible         =   0   'False
      Width           =   2952
      _ExtentX        =   5212
      _ExtentY        =   4577
   End
   Begin CannithCrafting.userCheckBox usrchkFiltered 
      Height          =   252
      Left            =   8340
      TabIndex        =   38
      Tag             =   "nav"
      Top             =   60
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Filtered"
   End
   Begin CannithCrafting.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8640
      Width           =   13548
      _ExtentX        =   23892
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      LeftLinks       =   "New Item;New Gearset;Load Gearset"
      RightLinks      =   "Effects;Materials;Augments;Scaling"
   End
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13548
      _ExtentX        =   23892
      _ExtentY        =   688
      Spacing         =   264
      BorderColor     =   -2147483640
      LeftLinks       =   "Items;Effects;Slotting;Review"
      RightLinks      =   "Help"
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8052
      Index           =   0
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   13275
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   13272
      Begin VB.PictureBox picTooltip 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000017&
         Height          =   252
         Left            =   6660
         ScaleHeight     =   225
         ScaleWidth      =   1125
         TabIndex        =   85
         Tag             =   "tip"
         Top             =   240
         Visible         =   0   'False
         Width           =   1152
      End
      Begin VB.Frame fraNotes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Weapons"
         ForeColor       =   &H80000008&
         Height          =   4872
         Left            =   8280
         TabIndex        =   6
         Top             =   60
         Width           =   4812
         Begin VB.TextBox txtNotes 
            BorderStyle     =   0  'None
            Height          =   4440
            Left            =   132
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   83
            Tag             =   "wrk"
            Top             =   360
            Width           =   4668
         End
         Begin VB.Label lblNotes 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Notes"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   528
         End
         Begin VB.Shape shpWeapons 
            Height          =   4752
            Left            =   0
            Top             =   120
            Width           =   4812
         End
      End
      Begin VB.Frame fraGearset 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Gearset"
         ForeColor       =   &H80000008&
         Height          =   2712
         Left            =   8280
         TabIndex        =   8
         Top             =   5160
         Width           =   4812
         Begin VB.CheckBox chkLoadGearset 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Load Gearset"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   660
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2040
            Width           =   1752
         End
         Begin VB.CheckBox chkSaveGearset 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Save Gearset"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2040
            Width           =   1752
         End
         Begin CannithCrafting.userSpinner usrspnBaseML 
            Height          =   312
            Index           =   0
            Left            =   3120
            TabIndex        =   11
            Top             =   360
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   556
            Max             =   34
            Value           =   34
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblGearset 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Gearset"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   9
            Top             =   0
            Width           =   708
         End
         Begin VB.Shape shpGearset 
            Height          =   2592
            Left            =   0
            Top             =   120
            Width           =   4812
         End
         Begin VB.Label lblML 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   $"frmGearset.frx":08CA
            ForeColor       =   &H80000008&
            Height          =   1128
            Index           =   1
            Left            =   300
            TabIndex        =   12
            Top             =   900
            Width           =   4176
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblML 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Gearset Minimum Level:"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   780
            TabIndex        =   10
            Top             =   420
            Width           =   2172
         End
      End
      Begin VB.Frame fraEquipment 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Equipment"
         ForeColor       =   &H80000008&
         Height          =   7812
         Left            =   180
         TabIndex        =   2
         Top             =   60
         Width           =   7752
         Begin VB.CheckBox chkLoadItemList 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Load Item List"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   7140
            Width           =   1752
         End
         Begin VB.CheckBox chkSaveItemList 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Save Item List"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   7140
            Width           =   1752
         End
         Begin VB.CheckBox chkDefaultItemList 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Set As Default Item List"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   4380
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   7140
            Width           =   3012
         End
         Begin VB.CheckBox chkNoAccessories 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "None"
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   4488
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   132
            Visible         =   0   'False
            Width           =   792
         End
         Begin VB.CheckBox chkAllAccessories 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "All"
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   3660
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   132
            Visible         =   0   'False
            Width           =   792
         End
         Begin CannithCrafting.userGearSlots usrGearSlots 
            Height          =   5844
            Left            =   360
            TabIndex        =   73
            Top             =   720
            Width           =   7032
            _ExtentX        =   12409
            _ExtentY        =   10663
         End
         Begin VB.Label lblHeader 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Crafted Effects or Named Item"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   1260
            TabIndex        =   79
            Top             =   420
            Width           =   2748
         End
         Begin VB.Label lblHeader 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Swap"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   7020
            TabIndex        =   82
            Top             =   420
            Width           =   492
         End
         Begin VB.Label lblHeader 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Ritual"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   6360
            TabIndex        =   81
            Top             =   420
            Width           =   480
         End
         Begin VB.Label lblHeader 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Augments"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   5100
            TabIndex        =   80
            Top             =   420
            Width           =   900
         End
         Begin VB.Label lblHeader 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Crafted?"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   360
            TabIndex        =   78
            Top             =   420
            Width           =   780
         End
         Begin VB.Label lblItemList 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Item List"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   77
            Top             =   6792
            Width           =   780
         End
         Begin VB.Label lblEquipment 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Equipment"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   936
         End
         Begin VB.Line linItemList 
            X1              =   0
            X2              =   7740
            Y1              =   6912
            Y2              =   6912
         End
         Begin VB.Shape shpEquipment 
            Height          =   7692
            Left            =   0
            Top             =   120
            Width           =   7752
         End
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8052
      Index           =   1
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   13275
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   13272
      Begin VB.PictureBox picAnalyze 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   972
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   12975
         TabIndex        =   71
         Top             =   6960
         Width           =   12972
         Begin VB.PictureBox picProgress 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            ForeColor       =   &H8000000D&
            Height          =   192
            Left            =   4380
            ScaleHeight     =   165
            ScaleWidth      =   8385
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   600
            Visible         =   0   'False
            Width           =   8412
         End
         Begin VB.Timer tmrFinish 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   1320
            Top             =   480
         End
         Begin VB.Timer tmrAnalyze 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1860
            Top             =   480
         End
         Begin VB.CheckBox chkCheck 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   0
            Width           =   1392
         End
         Begin VB.CheckBox chkAnalyze 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Analyze"
            ForeColor       =   &H80000008&
            Height          =   432
            Left            =   2820
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   480
            Width           =   1392
         End
         Begin CannithCrafting.userSpinner usrspnBaseML 
            Height          =   312
            Index           =   1
            Left            =   1740
            TabIndex        =   30
            Top             =   36
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   556
            Max             =   34
            Value           =   34
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblCheck 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Check if all effects can fit using QuickMatch (courtesy of Morten Michael Lindahl)"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   4380
            TabIndex        =   33
            Top             =   108
            Width           =   7128
         End
         Begin VB.Label lblMatches 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Matches"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   4440
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   744
         End
         Begin VB.Label lblTime 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Time"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   12300
            TabIndex        =   36
            Top             =   360
            Visible         =   0   'False
            Width           =   432
         End
         Begin VB.Label lblAnalyze 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Exhaustively check all combinations (based on priority) to see where all effects can go"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   4380
            TabIndex        =   72
            Top             =   588
            Width           =   7692
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Mnimum Level:"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Top             =   72
            Width           =   1368
         End
      End
      Begin VB.PictureBox picEffects 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6732
         Left            =   120
         ScaleHeight     =   6735
         ScaleWidth      =   12975
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   12972
         Begin CannithCrafting.userInfo usrDetails 
            Height          =   2772
            Left            =   9600
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   3720
            Width           =   3192
            _ExtentX        =   5636
            _ExtentY        =   4895
            TitleSize       =   2
         End
         Begin VB.TextBox txtSearch 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   4200
            TabIndex        =   21
            Top             =   60
            Width           =   1272
         End
         Begin VB.ListBox lstChosen 
            Appearance      =   0  'Flat
            Height          =   5910
            Left            =   5760
            TabIndex        =   25
            Top             =   420
            Width           =   3372
         End
         Begin VB.ListBox lstAvailable 
            Appearance      =   0  'Flat
            Height          =   5910
            Left            =   2100
            TabIndex        =   22
            Top             =   420
            Width           =   3372
         End
         Begin VB.ListBox lstGroup 
            Appearance      =   0  'Flat
            Height          =   5910
            Left            =   240
            TabIndex        =   18
            Top             =   420
            Width           =   1572
         End
         Begin CannithCrafting.userInfo usrInfo 
            Height          =   3012
            Left            =   9600
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   420
            Width           =   3192
            _ExtentX        =   5636
            _ExtentY        =   4895
            TitleSize       =   2
         End
         Begin VB.Image imgArrow 
            Enabled         =   0   'False
            Height          =   312
            Index           =   0
            Left            =   9180
            Picture         =   "frmGearset.frx":096A
            Stretch         =   -1  'True
            ToolTipText     =   "Clear All"
            Top             =   1500
            Width           =   300
         End
         Begin VB.Image imgArrow 
            Enabled         =   0   'False
            Height          =   312
            Index           =   2
            Left            =   9180
            Picture         =   "frmGearset.frx":1164
            Stretch         =   -1  'True
            ToolTipText     =   "Lower Priority"
            Top             =   780
            Width           =   300
         End
         Begin VB.Image imgArrow 
            Enabled         =   0   'False
            Height          =   312
            Index           =   1
            Left            =   9180
            Picture         =   "frmGearset.frx":195E
            Stretch         =   -1  'True
            ToolTipText     =   "HIgher Priority"
            Top             =   420
            Width           =   300
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Information"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   9600
            TabIndex        =   26
            Top             =   84
            Width           =   1044
         End
         Begin VB.Label lblSlots 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Slots: 0 of 0"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   8004
            TabIndex        =   24
            Top             =   84
            Width           =   1128
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Selected"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   5760
            TabIndex        =   23
            Top             =   84
            Width           =   756
         End
         Begin VB.Label lblLabel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Search:"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   3396
            TabIndex        =   20
            Top             =   84
            Width           =   696
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Effects"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   2100
            TabIndex        =   19
            Top             =   84
            Width           =   600
         End
         Begin VB.Label lblLabel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Group"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   84
            Width           =   552
         End
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8052
      Index           =   2
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   13275
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   13272
      Begin VB.PictureBox picHeader 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   552
         Left            =   420
         ScaleHeight     =   555
         ScaleWidth      =   7575
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   7572
         Begin CannithCrafting.userIcon usrIconSlot 
            Height          =   408
            Index           =   0
            Left            =   3960
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   408
            _ExtentX        =   714
            _ExtentY        =   714
            Style           =   1
         End
         Begin VB.CheckBox chkQuickMatch 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "QuickMatch"
            ForeColor       =   &H80000008&
            Height          =   372
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   60
            Width           =   1392
         End
      End
      Begin VB.VScrollBar scrollVertical 
         Height          =   3972
         Left            =   12480
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2100
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5412
         Left            =   1200
         ScaleHeight     =   5415
         ScaleWidth      =   8475
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1680
         Width           =   8472
         Begin VB.PictureBox picGrid 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4452
            Left            =   360
            ScaleHeight     =   4455
            ScaleWidth      =   6915
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   540
            Width           =   6912
         End
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8052
      Index           =   3
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   13275
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   13272
      Begin VB.Frame fraSaveLoad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   1632
         Left            =   180
         TabIndex        =   57
         Top             =   5040
         Width           =   1872
         Begin VB.CheckBox chkOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Load"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   5
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   900
            Width           =   1512
         End
         Begin VB.CheckBox chkOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Save"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   4
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   420
            Width           =   1512
         End
         Begin VB.Label lblFiles 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Gearset"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   60
            TabIndex        =   58
            Top             =   0
            Width           =   708
         End
         Begin VB.Shape shpFiles 
            Height          =   1512
            Left            =   0
            Top             =   120
            Width           =   1872
         End
      End
      Begin VB.Frame fraOutput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Output"
         ForeColor       =   &H80000008&
         Height          =   1632
         Left            =   180
         TabIndex        =   53
         Top             =   3300
         Width           =   1872
         Begin VB.CheckBox chkOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Forums"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   3
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   420
            Width           =   1512
         End
         Begin VB.CheckBox chkOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Plain Text"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   2
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   900
            Width           =   1512
         End
         Begin VB.Label lblOutput 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Output"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   60
            TabIndex        =   54
            Top             =   0
            Width           =   612
         End
         Begin VB.Shape Shape1 
            Height          =   1512
            Left            =   0
            Top             =   120
            Width           =   1872
         End
      End
      Begin VB.Frame fraDisplay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   1992
         Left            =   180
         TabIndex        =   48
         Top             =   1200
         Width           =   1872
         Begin CannithCrafting.userCheckBox usrchkBound 
            Height          =   252
            Left            =   420
            TabIndex        =   52
            Top             =   1440
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "Bound"
            Enabled         =   0   'False
         End
         Begin VB.CheckBox chkOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Item List"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   0
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   420
            Width           =   1512
         End
         Begin VB.CheckBox chkOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Ingredients"
            ForeColor       =   &H80000008&
            Height          =   432
            Index           =   1
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   900
            Width           =   1512
         End
         Begin VB.Label lblDisplay 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Display"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   60
            TabIndex        =   49
            Top             =   0
            Width           =   624
         End
         Begin VB.Shape shpDisplay 
            Height          =   1872
            Left            =   0
            Top             =   120
            Width           =   1872
         End
      End
      Begin VB.Frame fraML 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Minimum Level"
         ForeColor       =   &H80000008&
         Height          =   1032
         Left            =   180
         TabIndex        =   45
         Top             =   60
         Width           =   1872
         Begin CannithCrafting.userSpinner usrspnBaseML 
            Height          =   312
            Index           =   2
            Left            =   480
            TabIndex        =   47
            Top             =   420
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   556
            Max             =   34
            Value           =   34
            StepLarge       =   3
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Label lblReviewML 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Minimum Level"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   60
            TabIndex        =   46
            Top             =   0
            Width           =   1320
         End
         Begin VB.Shape shpML 
            Height          =   912
            Left            =   0
            Top             =   120
            Width           =   1872
         End
      End
      Begin VB.PictureBox picReviewItemList 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7704
         Left            =   2100
         ScaleHeight     =   7710
         ScaleWidth      =   11175
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   11172
         Begin VB.VScrollBar scrollVerticalReview 
            Height          =   7704
            Left            =   10920
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.PictureBox picReview 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   7704
            Left            =   180
            ScaleHeight     =   7710
            ScaleWidth      =   10755
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   0
            Width           =   10752
            Begin CannithCrafting.userItemSlot usrSlot 
               Height          =   996
               Index           =   0
               Left            =   0
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   0
               Width           =   10740
               _ExtentX        =   18944
               _ExtentY        =   1746
            End
         End
      End
      Begin VB.PictureBox picReviewOutput 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7992
         Left            =   2100
         ScaleHeight     =   7995
         ScaleWidth      =   11175
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   11172
         Begin CannithCrafting.userInfo usrReviewOutput 
            Height          =   7704
            Left            =   180
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   180
            Width           =   10932
            _ExtentX        =   19288
            _ExtentY        =   13600
            TitleSize       =   2
            TitleIcon       =   0   'False
            CanScroll       =   0   'False
         End
      End
      Begin VB.PictureBox picReviewIngredients 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7992
         Left            =   2100
         ScaleHeight     =   7995
         ScaleWidth      =   11175
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   11172
         Begin VB.ListBox lstReviewShards 
            Appearance      =   0  'Flat
            Height          =   7704
            IntegralHeight  =   0   'False
            Left            =   180
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   180
            Width           =   3492
         End
         Begin CannithCrafting.userInfo usrReviewIngredients 
            Height          =   7704
            Left            =   3900
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   180
            Width           =   7152
            _ExtentX        =   12621
            _ExtentY        =   13600
            TitleSize       =   2
            TitleIcon       =   0   'False
         End
      End
   End
End
Attribute VB_Name = "frmGearset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const AllowFixed As Boolean = False

Private Const ItemListExt As String = "itemlist"
Private Const GearsetExt As String = "gearset"

Private Enum WeaponSetEnum
    wseNone
    wseTHF
    wseTWF
    wseSWF
    wseSB
    wseOrb
    wseHandwraps
    wseBow
    wseThrow
    wseRunearm
End Enum

Private Enum ReviewEnum
    reItems
    reIngredients
    reOutputText
    reOutputForums
End Enum

Private Enum RowHeaderStateEnum
    rhseStandard
    rhseActive
    rhseLink
End Enum

Private Enum IngredientsEnum
    ieEverything = 900
    ieShards
    ieEffects
    ieML
    ieAugments
    ieEldritch
End Enum

Private Type IngredientAugmentType
    Color As AugmentColorEnum
    ColorOrder As Long
    FullName As String
    Augment As Long
    Variation As Long
    Scaling As Long
End Type

Private menReview As ReviewEnum

Private mstrFile As String

Private grid As GridType
Private gs As GearsetType
Private anal As AnalysisType ' Because I have the sense of humor of a 12-year-old

Private mblnDirty As Boolean
Private mblnOverride As Boolean
Private mlngTab As Long

Private mlngSlots As Long


' ************* FORM *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    Me.usrAugSlot.Init
    Me.usrAugPicker.Init aceAny
    InitItemList
    DefaultGearset
    InitEffects
    InitGroups
    If Len(mstrFile) Then OpenGearsetFile mstrFile, 0
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Resize()
    With Me.picTab(2)
        Me.picGrid.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.chkAnalyze.Caption = "Stop" Then
        Cancel = True
        Exit Sub
    End If
    If mblnDirty Then
        Select Case MsgBox("Save gearset before closing?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Notice")
            Case vbYes
                SaveFile
                Cancel = True
                Exit Sub
            Case vbNo
            Case vbCancel
                Cancel = True
                Exit Sub
        End Select
    End If
    If Not xp.DebugMode Then
        Call WheelUnHook(Me.hwnd)
    End If
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    Dim i As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    Select Case mlngTab
        Case 0
            If IsOver(Me.usrspnBaseML(0).hwnd, Xpos, Ypos) Then Me.usrspnBaseML(0).WheelScroll lngValue
        Case 1
            Select Case True
                Case IsOver(Me.usrspnBaseML(1).hwnd, Xpos, Ypos): Me.usrspnBaseML(1).WheelScroll lngValue
                Case IsOver(Me.usrInfo.hwnd, Xpos, Ypos): Me.usrInfo.Scroll -lngValue
                Case IsOver(Me.usrDetails.hwnd, Xpos, Ypos): Me.usrDetails.Scroll -lngValue
            End Select
        Case 2
            If IsOver(Me.picContainer.hwnd, Xpos, Ypos) Then ScrollGridWheel lngValue
        Case 3
            If IsOver(Me.usrspnBaseML(2).hwnd, Xpos, Ypos) Then
                Me.usrspnBaseML(2).WheelScroll lngValue
                Exit Sub
            End If
            Select Case menReview
                Case reItems
                    For i = 0 To mlngSlots
                        If IsOver(Me.usrSlot(i).Spinhwnd, Xpos, Ypos) Then
                            Me.usrSlot(i).SpinWheel lngValue
                            Exit Sub
                        End If
                    Next
                    If IsOver(Me.picReviewItemList.hwnd, Xpos, Ypos) Then ScrollItemList lngValue
                Case reIngredients
                    If IsOver(Me.usrReviewIngredients.hwnd, Xpos, Ypos) Then Me.usrReviewIngredients.Scroll -lngValue
            End Select
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If mlngTab <> 1 Then Exit Sub
    Select Case KeyAscii
        Case 43 ' +
            Me.BaseLevel = Me.BaseLevel + 1
        Case 45 ' -
            Me.BaseLevel = Me.BaseLevel - 1
        Case Else: Exit Sub
    End Select
    KeyAscii = 0
End Sub

Public Property Let BaseLevel(ByVal plngML As Long)
    Dim i As Long
    
    If plngML < 1 Then plngML = 1
    If plngML > MAX_CRAFTING_LEVEL Then plngML = MAX_CRAFTING_LEVEL
    If gs.BaseLevel = plngML Then Exit Property
    gs.BaseLevel = plngML
    mblnOverride = True
    For i = 0 To 2
        Me.usrspnBaseML(i).Value = gs.BaseLevel
    Next
    mblnOverride = False
    For i = 0 To seSlotCount - 1
        gs.Item(i).ML = gs.BaseLevel
    Next
    Select Case mlngTab
        Case 1
            RefreshEffects
        Case 3
            If menReview = reItems Then
                For i = 0 To mlngSlots
                    Me.usrSlot(i).ML = gs.BaseLevel
                Next
            Else
                Review menReview
            End If
    End Select
    gs.Analyzed = False
    Dirty True, False
End Property

Public Property Get BaseLevel() As Long
    BaseLevel = gs.BaseLevel
End Property

Public Sub Dirty(pblnDirty As Boolean, pblnReset As Boolean)
    Dim strCaption As String
    
    If pblnReset Then
        gs.Analyzed = False
        grid.Initialized = False
    End If
    mblnDirty = pblnDirty
    If Len(mstrFile) = 0 Then strCaption = "New Gearset" Else strCaption = mstrFile
    If mblnDirty Then strCaption = strCaption & "*"
    If Me.Caption <> strCaption Then Me.Caption = strCaption
End Sub

Public Sub Redraw()
    If mlngTab = 2 Or mlngTab = 3 Then ShowTab mlngTab
End Sub

Private Sub txtNotes_Change()
    If mblnOverride Then Exit Sub
    gs.Notes = Me.txtNotes.Text
    Dirty True, False
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Items": ShowTab 0
        Case "Effects": ShowTab 1
        Case "Slotting": ShowTab 2
        Case "Review": ShowTab 3
        Case "Help"
            Select Case mlngTab
                Case 0: ShowHelp "Gearset_-_Items"
                Case 1: ShowHelp "Gearset_-_Effects"
                Case 2: ShowHelp "Gearset_-_Slotting"
                Case 3: ShowHelp "Gearset_-_Review"
            End Select
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    FooterClick pstrCaption
End Sub

Private Sub ShowTab(plngTab As Long)
    Dim i As Long
    
    Me.usrAugSlot.Visible = False
    Me.usrAugPicker.Visible = False
    mlngTab = plngTab
    Me.usrchkFiltered.Visible = False
    For i = 0 To 3
        Me.picTab(i).Visible = False
    Next
    Select Case mlngTab
        Case 0
            ShowGearsetSlots
            Me.picTab(plngTab).Visible = True
            Me.picTab(plngTab).SetFocus
        Case 1
            RefreshEffects
            Me.picTab(plngTab).Visible = True
            If Me.lstGroup.Visible Then Me.lstGroup.SetFocus
        Case 2
            InitSlotting
            Me.usrchkFiltered.Visible = True
        Case 3
            Review reItems
            Me.picTab(plngTab).Visible = True
    End Select
End Sub


' ************* INITIALIZE *************


Private Sub InitItemList()
    Dim lngTop As Long
    
    Me.usrGearSlots.Init
    MoveHeader 0, vbCenter ' Crafted?
    MoveHeader 1, vbLeftJustify ' Named Item
    MoveHeader 2, vbLeftJustify ' Augments
    MoveHeader 3, vbRightJustify ' Ritual
    MoveHeader 4, vbLeftJustify ' Swaps
    With Me.usrGearSlots
        lngTop = .Top + .Height + (Me.chkLoadItemList.Top - .Top - .Height) \ 2
    End With
    With Me.linItemList
        .X1 = 0
        .X2 = Me.fraEquipment.Width
        .Y1 = lngTop
        .Y2 = lngTop
    End With
    Me.lblItemList.Top = Me.linItemList.Y2 - (Me.shpEquipment.Top - Me.lblEquipment.Top)
End Sub

Private Sub MoveHeader(plngIndex As Long, penAlign As AlignmentConstants)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngSpaces As Long
    Dim X1 As Long
    Dim X2 As Long
    
    ' ItemList header labels
    With Me.usrGearSlots
        lngHeight = Me.TextHeight("Q")
        lngTop = .Top - (lngHeight * 1.25)
        lngSpaces = Me.TextWidth("  ")
        lngWidth = Me.TextWidth(Me.lblHeader(plngIndex).Caption) + lngSpaces
        .GetCoords plngIndex + 1, X1, X2
        Select Case penAlign
            Case vbLeftJustify: lngLeft = .Left + X1
            Case vbCenter: lngLeft = .Left + X1 + (X2 - X1 - lngWidth) \ 2
            Case vbRightJustify: lngLeft = .Left + X2 - lngWidth
        End Select
        Me.lblHeader(plngIndex).Move lngLeft, lngTop, lngWidth, lngHeight
    End With
End Sub

Private Sub DefaultGearset()
    Dim typBlank As GearsetType
    Dim i As Long
    
    gs = typBlank
    gs.Mainhand = mheMelee
    gs.Offhand = oheMelee
    gs.TwoHanded = False
    ReDim gs.Item(seSlotCount - 1)
    For i = seHelmet To seArmor
        gs.Item(i).Crafted = True
    Next
    Me.BaseLevel = MAX_CRAFTING_LEVEL
    LoadItemList cfg.CraftingPath & "\Default." & ItemListExt, gs
    ShowGearsetSlots
    MapSlotsToGear gs
    Dirty False, True
End Sub

Private Sub InitEffects()
    Me.usrInfo.Clear
    Me.usrDetails.Clear
    Erase gs.Effect
    gs.Effects = 0
End Sub

Private Sub InitGroups()
    Dim i As Long
    
    Me.lstGroup.ListIndex = -1
    Me.lstGroup.Clear
    ListboxAddItem Me.lstGroup, "Show All", 0
    For i = 1 To db.Groups
        ListboxAddItem Me.lstGroup, db.Group(i), i
    Next
    Me.lstGroup.ListIndex = 0
End Sub


' ************* ITEMS *************


Private Sub chkAllAccessories_Click()
    If UncheckButton(Me.chkAllAccessories, mblnOverride) Then Exit Sub
    CheckAll True
End Sub

Private Sub chkNoAccessories_Click()
    If UncheckButton(Me.chkNoAccessories, mblnOverride) Then Exit Sub
    CheckAll False
End Sub

Private Sub CheckAll(pblnChecked As Boolean)
    Dim i As Long
    
    For i = seHelmet To seArmor
        Me.usrGearSlots.SetCrafted i, pblnChecked
        GearClick i
    Next
    Me.picTab(0).SetFocus
End Sub

Private Sub GearClick(ByVal penSlot As SlotEnum)
    If gs.Item(penSlot).Crafted = Me.usrGearSlots.GetCrafted(penSlot) Then Exit Sub
    gs.Item(penSlot).Crafted = Me.usrGearSlots.GetCrafted(penSlot)
    ShowGearSlot penSlot
    Dirty True, True
End Sub

Private Sub ShowGearsetSlots()
    Dim i As Long
    
    For i = seHelmet To seOffHand
        ShowGearSlot i
    Next
End Sub

Private Sub ShowGearSlot(ByVal penSlot As SlotEnum)
    If mblnOverride Then Exit Sub
    mblnOverride = True
    With gs.Item(penSlot)
        Me.usrGearSlots.SetCrafted penSlot, .Crafted
        SetEffects penSlot
        If Len(.ItemStyle) Then Me.usrGearSlots.SetItemStyle penSlot, .ItemStyle
        Me.usrGearSlots.SetNamedItem penSlot, .Named
        Me.usrGearSlots.SetAugmentSlots penSlot, GearsetAugmentToString(.Augment)
        Me.usrGearSlots.SetEldritchRitual penSlot, .EldritchRitual
    End With
    mblnOverride = False
End Sub

Private Sub SetEffects(penSlot As SlotEnum)
    Dim i As Long
    
    With gs.Item(penSlot)
        If .Crafted Then
            Me.usrGearSlots.SetEffects penSlot, GetEffect(.Effect(0)), GetEffect(.Effect(1)), GetEffect(.Effect(2))
        Else
            Me.usrGearSlots.SetEffects penSlot, vbNullString, vbNullString, vbNullString
        End If
    End With
End Sub

Private Function GetEffect(plngShard As Long) As String
    If plngShard = 0 Then Exit Function
    GetEffect = db.Shard(plngShard).ShortName
End Function

Private Sub usrGearSlots_CraftedChange(ByVal Slot As SlotEnum, ByVal Crafted As Boolean)
    GearClick Slot
End Sub

' Logic for mainhand changing between One-Hand and Two-Hand is a bit of a mess.
' It's handled internally in usrGearSlots, and also again here. Not sure it's needed in both.
Private Sub usrGearSlots_ItemTypeChange(ByVal Slot As SlotEnum, ByVal ItemType As String, HideOffhand As Boolean)
    Dim lngIndex As Long
    Dim enArmor As ArmorMaterialEnum
    Dim blnTwoHand As Boolean
    Dim enMainhand As MainHandEnum
    Dim blnReset As Boolean
    
    Select Case Slot
        Case seArmor
            enArmor = GetArmorMaterial(ItemType)
            If gs.Armor <> enArmor Then
                gs.Armor = enArmor
                blnReset = True
            End If
        Case seMainHand
            GetMainhandInfo ItemType, enMainhand, blnTwoHand
            gs.Mainhand = enMainhand
            gs.TwoHanded = blnTwoHand
            HideOffhand = blnTwoHand
            blnReset = True
        Case seOffHand
            blnReset = ChangeOffhand(GetOffhandStyle(ItemType))
    End Select
    gs.Item(Slot).ItemStyle = ItemType
    If blnReset Then MapSlotsToGear gs
    Dirty True, blnReset
End Sub

'~ Need to also reset slotted effects, eldritch rituals and augment slot colors
Private Function ChangeOffhand(penOffHand As OffHandEnum) As Boolean
    If gs.Offhand = penOffHand Then Exit Function
    gs.Offhand = penOffHand
    ChangeOffhand = True
End Function

Private Sub usrGearSlots_NamedItemChange(ByVal Slot As SlotEnum, ByVal NamedItem As String)
    gs.Item(Slot).Named = NamedItem
    Dirty True, False
End Sub

Private Sub chkLoadItemList_Click()
    Dim strFile As String
    
    If UncheckButton(Me.chkLoadItemList, mblnOverride) Then Exit Sub
    Me.picTab(0).SetFocus
    strFile = xp.ShowOpenDialog(cfg.CraftingPath, "Item Lists|*." & ItemListExt, ItemListExt)
    If Len(strFile) Then
        cfg.CraftingPath = GetPathFromFilespec(strFile)
        LoadItemList strFile, gs
        gs.Analyzed = False
        grid.Initialized = False
        Dirty True, True
        ShowGearsetSlots
        MapSlotsToGear gs
    End If
End Sub

Private Sub chkSaveItemList_Click()
    Dim strFile As String
    
    If UncheckButton(Me.chkSaveItemList, mblnOverride) Then Exit Sub
    Me.picTab(0).SetFocus
    strFile = xp.ShowSaveAsDialog(cfg.CraftingPath, strFile, "Item Lists|*." & ItemListExt, ItemListExt)
    If Len(strFile) = 0 Then Exit Sub
    If xp.File.Exists(strFile) Then
        If MsgBox(GetFileFromFilespec(strFile) & " exists. Overwrite?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    End If
    cfg.CraftingPath = GetPathFromFilespec(strFile)
    SaveItemList strFile, gs
End Sub

Private Sub chkDefaultItemList_Click()
    If UncheckButton(Me.chkDefaultItemList, mblnOverride) Then Exit Sub
    Me.picTab(0).SetFocus
    SaveItemList cfg.CraftingPath & "\Default." & ItemListExt, gs
End Sub

Private Sub usrspnBaseML_Change(Index As Integer)
    If Not mblnOverride Then Me.BaseLevel = Me.usrspnBaseML(Index).Value
End Sub

Private Sub chkLoadGearset_Click()
    If UncheckButton(Me.chkLoadGearset, mblnOverride) Then Exit Sub
    Me.picTab(0).SetFocus
    LoadFile
    On Error Resume Next
    Me.picTab(0).SetFocus
End Sub

Private Sub chkSaveGearset_Click()
    If UncheckButton(Me.chkSaveGearset, mblnOverride) Then Exit Sub
    Me.picTab(0).SetFocus
    SaveFile
End Sub

Private Sub usrGearSlots_Tooltip(TooltipText As String, Left As Long, Top As Long, Height As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngHeight As Long
    Dim strText As String
    
    If Len(TooltipText) = 0 Then
        If Me.picTooltip.Visible Then Me.picTooltip.Visible = False
        Me.picTooltip.Tag = vbNullString
    Else
        With Me.picTooltip
            If .Tag <> TooltipText Then
                .Tag = TooltipText
                strText = " " & TooltipText & " "
                .Width = .TextWidth(strText)
                lngHeight = .TextHeight(strText)
                .Height = lngHeight * 1.25
                .Cls
                .CurrentY = lngHeight \ 8
                Me.picTooltip.Print strText
                lngLeft = Me.fraEquipment.Left + Me.usrGearSlots.Left + Left
                lngTop = Me.fraEquipment.Top + Me.usrGearSlots.Top + Top + (Height - .Height) \ 2
                If .Left <> lngLeft Or .Top <> lngTop Then .Move lngLeft, lngTop
            End If
            If Not .Visible Then .Visible = True
        End With
    End If
End Sub

Private Sub picTooltip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picTooltip.Visible = False
End Sub

Private Sub fraEquipment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.picTooltip.Visible Then Me.picTooltip.Visible = False
End Sub


' ************* AUGMENTS *************


Private Sub usrGearSlots_AugmentClick(ByVal Slot As SlotEnum, Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    Me.usrAugPicker.Visible = False
    lngLeft = Me.picTab(0).Left + Me.fraEquipment.Left + Me.usrGearSlots.Left + Right + Me.TextWidth(" ")
    lngTop = Me.picTab(0).Top + Me.fraEquipment.Top + Me.usrGearSlots.Top + Top - Me.usrAugSlot.IconTop
    If lngTop + Me.usrAugSlot.Height > Me.ScaleHeight Then lngTop = Me.ScaleHeight - Me.usrAugSlot.Height
    Me.usrAugSlot.Move lngLeft, lngTop
    Me.usrAugSlot.Slot = Slot
    Me.usrAugSlot.SlotData = GearsetAugmentToString(gs.Item(Slot).Augment)
    Me.usrAugSlot.Visible = True
    On Error Resume Next ' I don't trust SetFocus not to throw an error just on general principle
    Me.usrAugSlot.SetFocus
End Sub

Private Sub usrAugSlot_DataChanged(Slot As SlotEnum, Text As String)
    StringToGearsetAugment gs.Item(Slot).Augment, Text
    Me.usrGearSlots.SetAugmentSlots Slot, Text
    Dirty True, False
End Sub

Private Sub usrAugSlot_ChooseAugment(Slot As SlotEnum, Color As AugmentColorEnum)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngAugment As Long
    Dim lngVariant As Long
    
    lngLeft = Me.usrAugSlot.Left + Me.usrAugSlot.Width
    lngTop = Me.usrAugSlot.Top - (Me.usrAugPicker.Height - Me.usrAugSlot.Height) \ 2
    If lngTop < Me.usrHeader.Top + Me.usrHeader.Height Then lngTop = Me.usrHeader.Top + Me.usrHeader.Height
    If lngTop + Me.usrAugPicker.Height > Me.ScaleHeight Then lngTop = Me.ScaleHeight - Me.usrAugPicker.Height
    With gs.Item(Slot).Augment(Color)
        lngAugment = .Augment
        lngVariant = .Variation
    End With
    With Me.usrAugPicker
        .Move lngLeft, lngTop
        .ClearSelection
        .SlotColor = Color
        .GearSlot = Slot
        If lngAugment > 0 And lngVariant > 0 Then .SetSelected lngAugment, lngVariant, 0
        .Visible = True
    End With
End Sub

Private Sub usrAugSlot_Hotkey(Slot As SlotEnum, KeyCode As Integer)
    Dim enSlot As Long
    Dim enMax As SlotEnum
    
    Select Case KeyCode
        Case vbKeyUp, vbKeySubtract: enSlot = Slot - 1
        Case vbKeyDown, vbKeyAdd: enSlot = Slot + 1
        Case vbKeyHome, vbKeyPageUp: enSlot = seHelmet
        Case vbKeyEnd, vbKeyPageDown: enSlot = seOffHand
    End Select
    If gs.TwoHanded Or gs.Item(seOffHand).ItemStyle = "Empty" Then enMax = seMainHand Else enMax = seOffHand
    Select Case enSlot
        Case Is < seHelmet: enSlot = seHelmet
        Case Is > enMax: enSlot = enMax
    End Select
    Me.usrGearSlots.SimulateAugmentClick enSlot
End Sub

Private Sub usrAugSlot_CloseControl()
    Me.usrAugPicker.Visible = False
    Me.usrAugSlot.Visible = False
End Sub

Private Sub usrAugPicker_AugmentSlotted(Augment As Long, Variation As Long, GearSlot As SlotEnum, AugmentSlot As AugmentColorEnum)
    Dim strText As String
    
    If GearSlot = seUnknown Or AugmentSlot = aceAny Then Exit Sub
    With gs.Item(GearSlot).Augment(AugmentSlot)
        .Augment = Augment
        .Variation = Variation
    End With
    strText = GearsetAugmentToString(gs.Item(GearSlot).Augment)
    Me.usrGearSlots.SetAugmentSlots GearSlot, strText
    Me.usrAugSlot.SlotData = strText
    Me.usrAugPicker.Visible = False
End Sub


' ************* ELDRITCH RITUAL *************


Private Sub usrGearSlots_EldritchChange(ByVal Slot As SlotEnum, ByVal Ritual As Long)
    If mblnOverride Then Exit Sub
    gs.Item(Slot).EldritchRitual = Ritual
    Dirty True, False
End Sub


' ************* EFFECTS *************


Private Sub lstGroup_Click()
    ShowMatches
End Sub

Private Sub lstGroup_GotFocus()
    Me.lstChosen.ListIndex = -1
End Sub

Private Sub lstGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            KeyCode = 0
        Case vbKeyRight
            KeyCode = 0
            Me.lstAvailable.SetFocus
    End Select
End Sub

Private Sub txtSearch_GotFocus()
    With Me.txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyReturn
            If Me.lstAvailable.ListCount > 0 Then
                Me.lstAvailable.ListIndex = 0
                Me.lstAvailable.SetFocus
            End If
    End Select
End Sub

Private Sub txtSearch_Change()
    ShowMatches
End Sub

Private Sub ShowMatches()
    Dim strGroup As String
    Dim strSearch As String
    Dim blnShow As Boolean
    Dim blnLookup() As Boolean
    Dim i As Long
    
    ListboxClear Me.lstAvailable
    ReDim blnLookup(1 To db.Shards)
    For i = 1 To gs.Effects
        blnLookup(gs.Effect(i)) = True
    Next
    strGroup = Me.lstGroup.Text
    If strGroup = "Show All" Then strGroup = vbNullString
    strSearch = LCase(Me.txtSearch.Text)
    For i = 1 To db.Shards
        With db.Shard(i)
            blnShow = Not blnLookup(i)
            If blnShow And Len(strGroup) <> 0 Then
                If .Group <> strGroup Then blnShow = False
            End If
            If blnShow And Len(strSearch) <> 0 Then
                If InStr(LCase$(.ShardName), strSearch) = 0 Then blnShow = False
            End If
            If blnShow Then
                If .ML > gs.BaseLevel Then blnShow = False
            End If
            If blnShow Then ListboxAddItem Me.lstAvailable, .ShardName, i
        End With
    Next
End Sub

Private Sub ShowChosen()
    Dim i As Long
    
    ListboxClear Me.lstChosen
    For i = 1 To gs.Effects
        ListboxAddItem Me.lstChosen, db.Shard(gs.Effect(i)).ShardName, gs.Effect(i)
    Next
    EnableArrows True
End Sub

Private Sub lstAvailable_GotFocus()
    Me.lstChosen.ListIndex = -1
End Sub

Private Sub lstAvailable_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If Me.lstAvailable.ListIndex = 0 Then Me.txtSearch.SetFocus
        Case vbKeyLeft
            KeyCode = 0
            Me.lstGroup.SetFocus
        Case vbKeyRight
            KeyCode = 0
            Me.lstChosen.SetFocus
        Case vbKeyReturn, vbKeyInsert
            KeyCode = 0
            AddAvailable
    End Select
End Sub

Private Sub lstAvailable_Click()
    ShardDetails Me.lstAvailable
End Sub

Private Sub lstAvailable_DblClick()
    AddAvailable
End Sub

Private Sub AddAvailable()
    If Me.lstAvailable.ListIndex = -1 Then Exit Sub
    gs.Effects = gs.Effects + 1
    ReDim Preserve gs.Effect(1 To gs.Effects)
    gs.Effect(gs.Effects) = ListboxGetValue(Me.lstAvailable)
    RefreshEffects
    Dirty True, True
End Sub

Private Sub lstChosen_GotFocus()
    Me.lstAvailable.ListIndex = -1
    EnableArrows True
End Sub

Private Sub lstChosen_LostFocus()
    EnableArrows False
End Sub

Private Sub EnableArrows(pblnEnabled As Boolean)
    With Me.lstChosen
        EnableArrow 1, (pblnEnabled = True And .ListIndex > 0)
        EnableArrow 2, (pblnEnabled = True And .ListIndex > -1 And .ListIndex < .ListCount - 1)
        EnableArrow 0, (Me.lstChosen.ListCount > 0)
    End With
End Sub

Private Sub EnableArrow(plngIndex As Long, pblnEnabled As Boolean)
    Dim enState As ArrowStateEnum
    Dim strID As String
    
    If pblnEnabled Then enState = aseEnabled Else enState = aseDisabled
    strID = GetArrowResource(plngIndex, enState)
    If Me.imgArrow(plngIndex).Enabled <> pblnEnabled Then
        Me.imgArrow(plngIndex).Enabled = pblnEnabled
        Me.imgArrow(plngIndex).Picture = LoadResPicture(strID, vbResBitmap)
    End If
End Sub

Private Sub lstChosen_Click()
    ShardDetails Me.lstChosen
End Sub

Private Sub lstChosen_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDelete
            RemoveChosen
        Case vbKeyLeft
            KeyCode = 0
            Me.lstAvailable.SetFocus
        Case vbKeyRight
            KeyCode = 0
        Case vbKeyUp
            Select Case Shift
                Case vbCtrlMask: MoveChosen -1
                Case vbAltMask: MoveChosen 0
                Case Else: Exit Sub
            End Select
            KeyCode = 0
        Case vbKeyDown
            Select Case Shift
                Case vbCtrlMask: MoveChosen 1
                Case vbAltMask: MoveChosen Me.lstChosen.ListCount - 1
                Case Else: Exit Sub
            End Select
            KeyCode = 0
    End Select
End Sub

Private Sub MoveChosen(plngNew As Long)
    Dim lngPos As Long
    Dim lngStep As Long
    Dim strSwap As String
    Dim lngSwap As Long
    Dim lngIndex As Long
    Dim i As Long
    
    If Me.lstChosen.ListIndex = -1 Then Exit Sub
    With Me.lstChosen
        Select Case plngNew
            Case -1, 1: lngPos = .ListIndex + plngNew
            Case Else: lngPos = plngNew
        End Select
        If lngPos < 0 Then lngPos = 0
        If lngPos > .ListCount - 1 Then lngPos = .ListCount - 1
        If .ListIndex <> lngPos Then
            If lngPos < .ListIndex Then lngStep = -1 Else lngStep = 1
            For i = .ListIndex To lngPos - lngStep Step lngStep
                lngIndex = i + 1
                lngSwap = gs.Effect(lngIndex)
                gs.Effect(lngIndex) = gs.Effect(lngIndex + lngStep)
                gs.Effect(lngIndex + lngStep) = lngSwap
                strSwap = .List(i)
                .List(i) = .List(i + lngStep)
                .List(i + lngStep) = strSwap
                lngSwap = .ItemData(i)
                .ItemData(i) = .ItemData(i + lngStep)
                .ItemData(i + lngStep) = lngSwap
            Next
            .ListIndex = lngPos
        End If
    End With
    Dirty True, True
End Sub

Private Sub lstChosen_KeyUp(KeyCode As Integer, Shift As Integer)
    EnableArrows True
End Sub

Private Sub lstChosen_DblClick()
    RemoveChosen
End Sub

Private Sub RemoveChosen()
    Dim i As Long
    
    If Me.lstChosen.ListIndex = -1 Then Exit Sub
    If gs.Effects = 1 Then
        Erase gs.Effect
        gs.Effects = 0
    Else
        For i = Me.lstChosen.ListIndex + 1 To gs.Effects - 1
            gs.Effect(i) = gs.Effect(i + 1)
        Next
        gs.Effects = gs.Effects - 1
        ReDim Preserve gs.Effect(1 To gs.Effects)
    End If
    RefreshEffects
    Dirty True, True
End Sub

Private Sub ShardDetails(plst As ListBox)
    Dim lngShard As Long
    
    lngShard = ListboxGetValue(plst)
    If lngShard = 0 Then Exit Sub
    With db.Shard(lngShard)
        ' Header
        Me.usrInfo.ClearContents
        Me.usrInfo.AddLink .ShardName, lseShard, .ShardName, 2, False
        Me.usrInfo.AddText "Group: " & .Group, 2
        ' Slots
        ShardSlots .Prefix, "Prefix: "
        ShardSlots .Suffix, "Suffix: "
        ShardSlots .Extra, "Extra: "
        Me.usrInfo.AddText vbNullString
        ' Bound
        Me.usrInfo.AddText "Bound Crafting:"
        Me.usrInfo.AddText "Level " & .Bound.Level, 1, "   "
        AddRecipeToInfo .Bound, Me.usrInfo, "   "
        Me.usrInfo.AddText vbNullString, 2
'        ' Unbound
'        Me.usrInfo.AddText "Unbound Crafting:"
'        Me.usrInfo.AddText "Level " & .Unbound.Level, 1, "   "
'        AddRecipeToInfo .Unbound, Me.usrInfo, "   "
'        Me.usrInfo.AddClipboard vbNullString
    End With
End Sub

Private Function ShardSlots(pblnSlot() As Boolean, pstrTitle As String) As String
    Dim strText As String
    Dim i As Long
    
    For i = 0 To geGearCount - 1
        If pblnSlot(i) Then
            If Len(strText) Then strText = strText & ", "
            strText = strText & GetGearName(i)
        End If
    Next
    If Len(strText) Then Me.usrInfo.AddText pstrTitle & strText, 1, vbNullString, pstrTitle
End Function

Private Sub UpdateSlotCounter()
    Dim lngSlots As Long
    Dim lngChosen As Long
    Dim blnLookup() As Boolean
    Dim lngColor As Long
    Dim lngItemSlots As Long
    Dim i As Long
    Dim j As Long
    
    Me.usrInfo.ClearContents
    Me.usrDetails.ClearContents
    lngColor = cfg.GetColor(cgeWorkspace, cveText)
    ReDim blnLookup(seSlotCount - 1)
    If gs.BaseLevel < 10 Then lngItemSlots = 2 Else lngItemSlots = 3
    For i = 0 To seSlotCount - 1
        blnLookup(i) = gs.Item(i).Crafted
        If gs.Item(i).Crafted Then lngSlots = lngSlots + lngItemSlots
    Next
    lngChosen = gs.Effects
    If lngChosen > lngSlots Then
        lngColor = cfg.GetColor(cgeWorkspace, cveTextError)
        Me.usrDetails.AddError "Error: Too many effects chosen", 2
    End If
    With Me.lblSlots
        .Caption = vbNullString
        If .ForeColor <> lngColor Then .ForeColor = lngColor
        .Caption = "Slots: " & lngChosen & " of " & lngSlots
    End With
End Sub

Private Sub RefreshEffects()
    Dim lngAvailableTop As Long
    Dim lngAvailableIndex As Long
    Dim lngChosenTop As Long
    Dim lngChosenIndex As Long
    Dim blnEnabled As Boolean
    
    If Not xp.DebugMode Then xp.LockWindow Me.hwnd
    lngAvailableTop = Me.lstAvailable.TopIndex
    lngAvailableIndex = Me.lstAvailable.ListIndex
    lngChosenTop = Me.lstChosen.TopIndex
    lngChosenIndex = Me.lstChosen.ListIndex
    ShowMatches
    ShowChosen
    UpdateSlotCounter
    CheckForErrors gs, Me.usrDetails
    With Me.lstAvailable
        If lngAvailableIndex > .ListCount - 1 Then lngAvailableIndex = .ListCount - 1
        .TopIndex = lngAvailableTop
        .ListIndex = lngAvailableIndex
    End With
    With Me.lstChosen
        If lngChosenIndex > .ListCount - 1 Then lngChosenIndex = .ListCount - 1
        .TopIndex = lngChosenTop
        .ListIndex = lngChosenIndex
    End With
    blnEnabled = (gs.Effects > 0 And TotalSlots() > 0)
    Me.chkCheck.Enabled = blnEnabled
    Me.chkAnalyze.Enabled = blnEnabled
    If Not xp.DebugMode Then xp.UnlockWindow
End Sub

Private Function TotalSlots() As Long
    Dim lngCount As Long
    Dim i As Long
    
    For i = 0 To seSlotCount - 1
        If gs.Item(i).Crafted Then lngCount = lngCount + 1
    Next
    If gs.BaseLevel < 10 Then TotalSlots = lngCount * 2 Else TotalSlots = lngCount * 3
End Function

Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strID As String
    
    strID = GetArrowResource(CLng(Index), asePressed)
    Me.imgArrow(Index).Picture = LoadResPicture(strID, vbResBitmap)
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strID As String
    
    strID = GetArrowResource(CLng(Index), aseEnabled)
    Me.imgArrow(Index).Picture = LoadResPicture(strID, vbResBitmap)
End Sub

Private Sub imgArrow_Click(Index As Integer)
    Select Case Index
        Case aeUp
            MoveChosen -1
        Case aeDown
            MoveChosen 1
        Case aeDelete
            If MsgBox("Clear all chosen effects?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
            Erase gs.Effect
            gs.Effects = 0
            RefreshEffects
            grid.Initialized = False
    End Select
End Sub

Private Sub imgArrow_DblClick(Index As Integer)
    Select Case Index
        Case aeUp: MoveChosen -1
        Case aeDown: MoveChosen 1
    End Select
End Sub

Private Sub chkCheck_Click()
    Dim lngFailed As Long
    Dim lngIterations As Long
    Dim dblTime As Double
    Dim strTries As String
    Dim strFormat As String
    
    If UncheckButton(Me.chkCheck, mblnOverride) Then Exit Sub
    Me.picAnalyze.SetFocus
    If Not AnalysisCanRun() Then Exit Sub
    QuickMatchCheck gs, lngIterations, dblTime, lngFailed
    If lngIterations = 1 Then strTries = "1 try, " Else strTries = lngIterations & " tries, "
    If lngFailed = 0 Then strFormat = "0.00000" Else strFormat = "0.00"
    Me.usrDetails.AddText "QuickMatch"
    Me.usrDetails.AddText strTries & Format(dblTime, strFormat) & " seconds", 2
    If lngFailed = 0 Then
        Me.usrDetails.AddText "Success! All effects can fit."
    Else
        Me.usrDetails.AddError "Unable to fit " & db.Shard(gs.Effect(lngFailed)).Abbreviation
    End If
End Sub

Private Sub chkAnalyze_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcArrow
End Sub

Private Sub chkAnalyze_Click()
    Dim i As Long
    
    If UncheckButton(Me.chkAnalyze, mblnOverride) Then Exit Sub
    Me.picAnalyze.SetFocus
    If Me.chkAnalyze.Caption = "Stop" Then
        Me.tmrAnalyze.Enabled = False
        FinishProcessing
    Else
        If gs.Analyzed Then
            If MsgBox("Analysis has already been completed." & vbNewLine & vbNewLine & "Run it again anyway?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
        End If
        If Not AnalysisCanRun() Then Exit Sub
        ClearSelections gs
        EnableForms False
        xp.Mouse = msAppWait
        Me.picProgress.Cls
        Me.lblCheck.Visible = False
        Me.lblAnalyze.Visible = False
        Me.lblMatches.Caption = vbNullString
        Me.lblMatches.Visible = True
        Me.lblTime.Caption = vbNullString
        Me.lblTime.Visible = True
        Me.picProgress.Visible = True
        Me.chkAnalyze.Caption = "Stop"
        InitProcessing gs, anal
        Me.tmrAnalyze.Enabled = True
    End If
End Sub

Private Function AnalysisCanRun() As Boolean
    Me.lstAvailable.ListIndex = -1
    Me.lstChosen.ListIndex = -1
    Me.usrInfo.Clear
    Me.usrDetails.Clear
    If gs.Effects > TotalSlots() Then
        Me.usrDetails.AddError "Error: Too many effects chosen", 2
        Me.usrDetails.AddText "Resolve errors and try again"
        Exit Function
    End If
    If CheckForErrors(gs, Me.usrDetails) Then
        Me.usrDetails.AddText "Resolve errors and try again"
        Exit Function
    End If
    AnalysisCanRun = True
End Function

Private Sub tmrAnalyze_Timer()
    Me.tmrAnalyze.Enabled = False
    UpdateProgressbar
    If ProcessingFinished() Then
        ShowResults
        Me.tmrFinish.Enabled = True
    Else
        Me.tmrAnalyze.Enabled = True
        ProcessChunk gs, anal
    End If
End Sub

Private Sub UpdateProgressbar()
    Me.lblMatches.Caption = Format(GetValid(), "#,##0") & " valid combinations"
    Me.lblTime.Caption = StopwatchStopTime()
    Me.picProgress.Line (0, 0)-(Me.picProgress.ScaleWidth * GetProgress(gs, anal), Me.picProgress.ScaleHeight), vbHighlight, BF
    Me.picProgress.Refresh
End Sub

Private Sub FinishProcessing()
    Me.tmrFinish.Enabled = False
    Me.lblMatches.Visible = False
    Me.lblTime.Visible = False
    Me.picProgress.Visible = False
    Me.lblCheck.Visible = True
    Me.lblAnalyze.Visible = True
    Me.chkAnalyze.Caption = "Analyze"
    EnableForms True
    xp.Mouse = msNormal
End Sub

Private Sub tmrFinish_Timer()
    FinishProcessing
End Sub

Private Sub EnableForms(pblnEnabled As Boolean)
    Dim frm As Form

    For Each frm In Forms
        frm.Enabled = pblnEnabled
    Next
    Me.Enabled = True
    Me.picEffects.Enabled = pblnEnabled
    Me.usrspnBaseML(1).Enabled = pblnEnabled
    Me.chkCheck.Enabled = pblnEnabled
    Me.usrHeader.Enabled = pblnEnabled
    Me.usrFooter.Enabled = pblnEnabled
End Sub

Private Sub ShowResults()
    Dim lngShard As Long
    
    Me.usrDetails.AddText "Analyzed " & FormattedCombinations() & " combinations in " & StopwatchStopFormatted(), 2
    Me.usrDetails.AddText Format(GetValid(), "#,##0") & " valid combinations found"
    If GetValid() = 0 Then
        Me.usrDetails.AddText vbNullString
        lngShard = gs.Effect(GetFailedOn())
        Me.usrDetails.AddError "Failed on: " & db.Shard(lngShard).Abbreviation
        If lngShard <> Me.lstChosen.ItemData(Me.lstChosen.ListCount - 1) Then
            Me.usrDetails.AddText vbNullString
            Me.usrDetails.AddText "Effects after " & db.Shard(lngShard).Abbreviation & " were not analyzed"
        End If
        Dirty True, True
    Else
        Dirty True, False
        gs.Analyzed = True
        grid.Initialized = False
    End If
End Sub


' ************* SLOTTING *************


Private Sub InitSlotting()
    Me.picGrid.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
    Me.usrchkFiltered.Value = gs.Analyzed
    Me.usrchkFiltered.Enabled = gs.Analyzed
    Me.chkQuickMatch.Enabled = (gs.Effects > 0 And TotalSlots() > 0)
    RefreshGrid
End Sub

Private Sub usrchkFiltered_UserChange()
    grid.Initialized = False
    RefreshGrid
End Sub

Private Sub RefreshGrid()
    Dim i As Long
    
    Me.picTab(2).Visible = False
    Me.picHeader.Cls
    Me.picGrid.Cls
    If Not grid.Initialized Then
        InitGrid grid, gs, anal, Me.usrchkFiltered.Value
    End If
    CalculateDimensions
    DrawRowHeaders
    DrawColumnHeaders
    DrawGrid
    Me.picTab(2).Visible = True
End Sub

' This started off simple, then more and more stuff got added
' So the entire control sizing code is now kind of a jumbled mess
Private Sub CalculateDimensions()
    Dim lngShard As Long
    Dim lngMax As Long
    Dim lngSpots As Long
    Dim lngSlots As Long
    Dim i As Long
    
    Me.picGrid.FontBold = True
    grid.RowHeight = ((Me.picGrid.TextHeight("Q") * 1.1) \ PixelY) * PixelY
    grid.TextOffsetY = (grid.RowHeight - Me.picGrid.TextHeight("Q")) \ 2
    grid.IconTop = 0
    grid.IconHeight = Me.usrIconSlot(0).Height
    grid.IconWidth = Me.usrIconSlot(0).Width
    With Me.chkQuickMatch
        .Move .Left, grid.IconTop, .Width, grid.IconHeight
    End With
    grid.HeaderHeight = grid.IconTop + grid.IconHeight + grid.RowHeight
    ResizePictureboxes
    grid.EffectWidth = Me.chkQuickMatch.Left * 2 + Me.chkQuickMatch.Width
    For i = 1 To gs.Effects
        lngShard = gs.Effect(i)
        If grid.EffectWidth < Me.picGrid.TextWidth(db.Shard(lngShard).GridName) Then
            grid.EffectWidth = Me.picGrid.TextWidth(db.Shard(lngShard).GridName)
        End If
    Next
    grid.EffectWidth = grid.EffectWidth + PixelX
    If gs.BaseLevel < 10 Then lngSpots = 2 Else lngSpots = 3
    lngMax = Me.picGrid.ScaleWidth - grid.EffectWidth
    lngSlots = grid.Slots
    If lngSlots < 6 Then lngSlots = 6
    grid.SlotWidth = ((lngMax \ lngSlots) \ PixelX) * PixelX
    grid.ColWidth = ((grid.SlotWidth \ lngSpots) \ PixelX) * PixelX
    grid.SlotWidth = grid.ColWidth * lngSpots
    grid.EffectWidth = Me.picGrid.ScaleWidth - (grid.SlotWidth * lngSlots) - PixelX
    Do
        If ApplyAffixes("Prefix", "Suffix", "Extra") Then Exit Do
        If ApplyAffixes("Pre", "Suf", "Ext") Then Exit Do
        If ApplyAffixes("Pr", "Su", "Ex") Then Exit Do
        If ApplyAffixes("P", "S", "X", True) Then Exit Do
    Loop Until True
    If grid.Slots < lngSlots Then
        Me.picGrid.Width = grid.EffectWidth + grid.SlotWidth * grid.Slots + PixelX
        Me.picContainer.Width = Me.picGrid.Width + PixelX
        Me.scrollVertical.Left = Me.picContainer.Left + Me.picContainer.Width
    End If
    For i = 1 To grid.Slots
        grid.Slot(i).IconLeft = grid.EffectWidth + grid.SlotWidth * (i - 1) + (grid.SlotWidth - grid.IconWidth) \ 2
    Next
    Me.picGrid.FontBold = False
End Sub

Private Sub ResizePictureboxes()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim blnScroll As Boolean
    Dim lngClientHeight As Long
    
    Me.scrollVertical.Visible = False
    ' Tab size starts same as other tabs
    With Me.picTab(0)
        lngLeft = .Left
        lngTop = .Top
        lngWidth = .Width
        lngHeight = .Height
    End With
    lngClientHeight = grid.RowHeight * grid.Rows + PixelY
    If grid.HeaderHeight + lngClientHeight > lngHeight Then
        lngHeight = Me.ScaleHeight - Me.usrHeader.Height - Me.usrFooter.Height - (PixelY * 2)
        lngHeight = (lngHeight \ grid.RowHeight) * grid.RowHeight
        lngTop = Me.usrHeader.Height + (Me.ScaleHeight - Me.usrHeader.Height - Me.usrFooter.Height - lngHeight) \ 2
        lngWidth = Me.ScaleWidth - lngLeft
    End If
    Me.picTab(2).Move lngLeft, lngTop, lngWidth, lngHeight
    If grid.HeaderHeight + lngClientHeight > lngHeight Then
        blnScroll = True
        lngWidth = lngWidth - Me.scrollVertical.Width
    End If
    Me.picHeader.Move 0, 0, lngWidth, grid.HeaderHeight '+ PixelY
    lngHeight = Me.picTab(2).Height - grid.HeaderHeight
    lngHeight = (lngHeight \ grid.RowHeight) * grid.RowHeight + PixelY
    Me.picContainer.Move 0, grid.HeaderHeight, lngWidth, lngHeight
    Me.picGrid.Move 0, 0, Me.picContainer.ScaleWidth, lngClientHeight
    With Me.picContainer
        Me.scrollVertical.Move .Left + .Width, .Top, Me.scrollVertical.Width, .Height
    End With
    If blnScroll Then
        With Me.scrollVertical
            mblnOverride = True
            .Value = 0
            .Max = (Me.picGrid.Height - Me.picContainer.Height) \ grid.RowHeight
            .SmallChange = 1
            .LargeChange = Me.picContainer.Height \ grid.RowHeight
            .Visible = True
            mblnOverride = False
        End With
    End If
End Sub

Private Function ApplyAffixes(pstrPrefix As String, pstrSuffix As String, pstrExtra As String, Optional pblnForce As Boolean) As Boolean
    Dim blnApply As Boolean
    Dim i As Long
    
    blnApply = True
    If Me.picGrid.TextWidth(pstrPrefix & "  ") > grid.ColWidth Then
        blnApply = False
    ElseIf Me.picGrid.TextWidth(pstrSuffix & "  ") > grid.ColWidth Then
        blnApply = False
    ElseIf Me.picGrid.TextWidth(pstrExtra & "  ") > grid.ColWidth Then
        blnApply = False
    End If
    If pblnForce Then blnApply = True
    If Not blnApply Then Exit Function
    For i = 1 To grid.Cols
        Select Case grid.Col(i).Affix
            Case aePrefix: grid.Col(i).Header = pstrPrefix
            Case aeSuffix: grid.Col(i).Header = pstrSuffix
            Case aeExtra: grid.Col(i).Header = pstrExtra
        End Select
    Next
    ApplyAffixes = True
End Function

Private Sub DrawRowHeaders()
    Dim i As Long
    
    For i = 1 To gs.Effects
        DrawRowHeader i, False
    Next
End Sub

Private Sub DrawRowHeader(plngRow As Long, penState As RowHeaderStateEnum)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    
    If penState = rhseActive Then Me.picGrid.FontBold = True
    lngLeft = 0
    lngTop = grid.RowHeight * (plngRow - 1)
    lngRight = grid.EffectWidth - PixelX
    lngBottom = lngTop + grid.RowHeight
    If grid.Row(plngRow).ColSelected Then
        Me.picGrid.ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        lngColor = cfg.GetColor(cgeWorkspace, cveBackground)
    Else
        Me.picGrid.ForeColor = cfg.GetColor(cgeControls, cveText)
        lngColor = cfg.GetColor(cgeControls, cveBackground)
    End If
    If penState = rhseLink Then Me.picGrid.ForeColor = cfg.GetColor(cgeWorkspace, cveTextLink)
    Me.picGrid.Line (lngLeft, lngTop + PixelY)-(lngRight, lngBottom - PixelY), lngColor, BF
    Me.picGrid.CurrentX = lngLeft
    Me.picGrid.CurrentY = lngTop + grid.TextOffsetY
    Me.picGrid.Print grid.Row(plngRow).Caption
    If penState = rhseActive Then Me.picGrid.FontBold = False
End Sub

Private Sub DrawColumnHeaders()
    Dim i As Long
    
    For i = 1 To grid.Slots
        If i > Me.usrIconSlot.UBound Then Load Me.usrIconSlot(i)
        DrawSlotIcon i
    Next
    For i = Me.usrIconSlot.UBound To grid.Slots + 1 Step -1
        Unload Me.usrIconSlot(i)
    Next
    For i = 1 To grid.Cols
        DrawColumnHeader i, False
    Next
    DrawHeaderLine
End Sub

Private Sub DrawHeaderLine()
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    
    lngLeft = grid.EffectWidth
    lngRight = lngLeft + grid.ColWidth * (grid.Cols)
    lngBottom = grid.HeaderHeight
    Me.picHeader.Line (lngLeft, lngBottom)-(lngRight, lngBottom), cfg.GetColor(cgeControls, cveBorderExterior)
End Sub

Private Sub DrawSlotIcon(plngSlot As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    lngLeft = grid.Slot(plngSlot).IconLeft
    lngTop = grid.IconTop
    With Me.usrIconSlot(plngSlot)
        .Init grid.Slot(plngSlot).Gear, uiseLink, True
        .IconName = gs.Item(grid.Slot(plngSlot).GearsetSlot).ItemStyle
        .AllowMenu = False
        .Move lngLeft, lngTop
        .Visible = True
    End With
End Sub

Private Sub DrawColumnHeader(plngCol As Long, pblnActive As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    
    If pblnActive Then Me.picHeader.FontBold = True
    lngLeft = grid.EffectWidth + grid.ColWidth * (plngCol - 1)
    lngTop = grid.HeaderHeight - grid.RowHeight
    lngRight = lngLeft + grid.ColWidth
    lngBottom = lngTop + grid.RowHeight - PixelY
    If grid.Col(plngCol).RowSelected Then
        Me.picHeader.ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        lngColor = cfg.GetColor(cgeWorkspace, cveBackground)
    Else
        Me.picHeader.ForeColor = cfg.GetColor(cgeControls, cveText)
        lngColor = cfg.GetColor(cgeControls, cveBackground)
    End If
    Me.picHeader.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngColor, BF
    Me.picHeader.CurrentX = lngLeft + (grid.ColWidth - Me.picHeader.TextWidth(grid.Col(plngCol).Header)) \ 2
    Me.picHeader.CurrentY = lngTop + grid.TextOffsetY
    Me.picHeader.Print grid.Col(plngCol).Header
    If pblnActive Then Me.picHeader.FontBold = False
End Sub

Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    For lngRow = 1 To gs.Effects
        For lngCol = 1 To grid.Cols
            DrawCell lngRow, lngCol, False
        Next
    Next
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    Dim strText As String
    
    lngLeft = grid.EffectWidth + grid.ColWidth * (plngCol - 1)
    lngTop = grid.RowHeight * (plngRow - 1)
    lngRight = lngLeft + grid.ColWidth
    lngBottom = lngTop + grid.RowHeight
    If grid.Cell(plngRow, plngCol).Selected Then
        Me.picGrid.ForeColor = cfg.GetColor(cgeControls, cveText)
        lngColor = cfg.GetColor(cgeControls, cveBackHighlight)
    ElseIf grid.Col(plngCol).RowSelected <> 0 Or grid.Row(plngRow).ColSelected <> 0 Then
        Me.picGrid.ForeColor = cfg.GetColor(cgeControls, cveTextDim)
        lngColor = cfg.GetColor(cgeControls, cveBackground)
    Else
        Me.picGrid.ForeColor = cfg.GetColor(cgeControls, cveText)
        lngColor = cfg.GetColor(cgeControls, cveBackground)
    End If
    Me.picGrid.Line (lngLeft + PixelX, lngTop + PixelY)-(lngRight - PixelX, lngBottom - PixelY), lngColor, BF
    If grid.Cell(plngRow, plngCol).Active Then
        strText = grid.Col(plngCol).Header
        If Len(strText) Then
            Me.picGrid.CurrentX = lngLeft + (grid.ColWidth - Me.picGrid.TextWidth(strText)) \ 2
            Me.picGrid.CurrentY = lngTop + grid.TextOffsetY
            Me.picGrid.Print strText
        End If
    End If
    DrawBorder plngRow, plngCol, pblnActive
End Sub

Private Sub DrawBorder(plngRow As Long, plngCol As Long, pblnActive As Boolean)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngColor As Long
    
    lngLeft = grid.EffectWidth + grid.ColWidth * (plngCol - 1)
    lngTop = grid.RowHeight * (plngRow - 1)
    lngRight = lngLeft + grid.ColWidth
    lngBottom = lngTop + grid.RowHeight
    If pblnActive Then
        lngColor = cfg.GetColor(cgeControls, cveBorderHighlight)
    Else
        lngColor = cfg.GetColor(cgeControls, cveBorderInterior)
    End If
    Me.picGrid.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngColor, B
    If pblnActive Then Exit Sub
    lngColor = cfg.GetColor(cgeControls, cveBorderExterior)
    If grid.Row(plngRow).TopThick Then Me.picGrid.Line (lngLeft, lngTop)-(lngRight + PixelX, lngTop), lngColor
    If grid.Col(plngCol).LeftThick Then Me.picGrid.Line (lngLeft, lngTop)-(lngLeft, lngBottom + PixelY), lngColor
    If grid.Col(plngCol).RightThick Then Me.picGrid.Line (lngRight, lngTop)-(lngRight, lngBottom + PixelY), lngColor
    If grid.Row(plngRow).BottomThick Then Me.picGrid.Line (lngLeft, lngBottom)-(lngRight + PixelX, lngBottom), lngColor
End Sub

Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
    If grid.CurrentRow <> 0 And grid.CurrentCol = 0 Then
        OpenShard db.Shard(grid.Row(grid.CurrentRow).Shard).ShardName
    Else
        ToggleEffect grid.CurrentRow, grid.CurrentCol
    End If
End Sub

Private Sub usrIconSlot_ActiveChange(Index As Integer, pblnActive As Boolean)
    If pblnActive Then grid.CurrentSlot = Index
End Sub

Private Sub usrIconSlot_Click(Index As Integer)
    If grid.CurrentSlot <> Index Then grid.CurrentSlot = Index ' Unnecessary error correction
    OpenItem grid.Slot(grid.CurrentSlot).GearsetSlot
End Sub

Private Sub ActiveCell(X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngSlot As Long
    
    If X = 0 And Y = 0 Then lngRow = 0 Else lngRow = (Y \ grid.RowHeight) + 1
    If X < grid.EffectWidth Then lngCol = 0 Else lngCol = (X - grid.EffectWidth) \ grid.ColWidth + 1
    If lngRow < 1 Or lngRow > gs.Effects Then
        lngRow = 0
    ElseIf Me.scrollVertical.Visible Then
        If lngRow > Me.scrollVertical.Value + Me.picContainer.Height \ grid.RowHeight Then lngRow = 0
    End If
    If lngCol < 1 Or lngCol > grid.Cols Then lngCol = 0
    If lngRow = 0 Then lngCol = 0
    If lngCol = 0 Then
        If lngRow Then xp.SetMouseCursor mcHand
        lngSlot = 0
    Else
        lngSlot = grid.Col(lngCol).Slot
    End If
    If lngRow = grid.CurrentRow And lngCol = grid.CurrentCol Then
        If lngSlot <> grid.CurrentSlot Then ActiveSlot 0, 0
        Exit Sub
    End If
    If grid.CurrentRow <> 0 Then
        If grid.CurrentCol = 0 Then
            DrawRowHeader grid.CurrentRow, rhseStandard
        Else
            If grid.CurrentRow <> lngRow Then DrawRowHeader grid.CurrentRow, False
            If grid.CurrentCol <> lngCol Then DrawColumnHeader grid.CurrentCol, False
            DrawBorder grid.CurrentRow, grid.CurrentCol, False
        End If
    End If
    grid.CurrentRow = lngRow
    grid.CurrentCol = lngCol
    If grid.CurrentRow <> 0 Then
        If grid.CurrentCol = 0 Then
            DrawRowHeader grid.CurrentRow, rhseLink
        Else
            DrawRowHeader grid.CurrentRow, rhseActive
            DrawColumnHeader grid.CurrentCol, True
            DrawBorder grid.CurrentRow, grid.CurrentCol, True
        End If
    End If
    If grid.CurrentSlot <> lngSlot Then
        If grid.CurrentSlot Then Me.usrIconSlot(grid.CurrentSlot).Active = False
        grid.CurrentSlot = lngSlot
        If grid.CurrentSlot Then Me.usrIconSlot(grid.CurrentSlot).Active = True
    End If
End Sub

Private Sub ToggleEffect(plngRow As Long, plngCol As Long)
    Dim enGear As GearEnum
    Dim blnSelected As Boolean
    Dim i As Long
    
    If plngRow = 0 Or plngCol = 0 Then Exit Sub
    If Not grid.Cell(plngRow, plngCol).Active Then Exit Sub
    blnSelected = Not grid.Cell(plngRow, plngCol).Selected
    If blnSelected Then
        For i = 1 To grid.Rows
            If i <> plngRow And grid.Row(i).ColSelected = plngCol Then
                grid.Row(i).ColSelected = 0
                grid.Cell(i, plngCol).Selected = False
                ShowRow i, False
            End If
        Next
        For i = 1 To grid.Cols
            If i <> plngCol And grid.Col(i).RowSelected = plngRow Then
                grid.Col(i).RowSelected = 0
                grid.Cell(plngRow, i).Selected = False
                ShowCol i, False
            End If
        Next
    End If
    If blnSelected Then
        grid.Row(plngRow).ColSelected = plngCol
        grid.Col(plngCol).RowSelected = plngRow
        grid.Cell(plngRow, plngCol).Selected = True
    Else
        grid.Row(plngRow).ColSelected = 0
        grid.Col(plngCol).RowSelected = 0
        grid.Cell(plngRow, plngCol).Selected = False
    End If
    GridSelection gs, grid, plngRow, plngCol, blnSelected
    Dirty True, False
    ShowRow plngRow, True
    ShowCol plngCol, True
    DrawCell plngRow, plngCol, True
End Sub

Private Sub ShowRow(plngRow As Long, pblnActive As Boolean)
    Dim enState As RowHeaderStateEnum
    Dim lngCol As Long
    
    If pblnActive Then enState = rhseActive Else enState = rhseStandard
    DrawRowHeader plngRow, enState
    For lngCol = 1 To grid.Cols
        DrawCell plngRow, lngCol, False
    Next
End Sub

Private Sub ShowCol(plngCol As Long, pblnActive As Boolean)
    Dim lngRow As Long
    
    DrawColumnHeader plngCol, pblnActive
    For lngRow = 1 To gs.Effects
        DrawCell lngRow, plngCol, False
    Next
End Sub

Private Sub chkQuickMatch_Click()
    If UncheckButton(Me.chkQuickMatch, mblnOverride) Then Exit Sub
    Me.picHeader.SetFocus
    QuickMatchCommit gs
    grid.Initialized = False
    RefreshGrid
    Dirty True, False
End Sub

Private Sub scrollVertical_Change()
    If Not mblnOverride Then ScrollGrid
End Sub

Private Sub scrollVertical_GotFocus()
    ActiveCell 0, 0
    Me.picGrid.SetFocus
End Sub

Private Sub scrollVertical_Scroll()
    If Not mblnOverride Then ScrollGrid
End Sub

Private Sub ScrollGrid()
    Me.picGrid.Top = 0 - Me.scrollVertical.Value * grid.RowHeight
End Sub

Private Sub ScrollGridWheel(plngValue As Long)
    Dim lngValue As Long
    
    With Me.scrollVertical
        If .Visible Then
            lngValue = .Value - (plngValue * .SmallChange)
            If lngValue < .Min Then lngValue = .Min
            If lngValue > .Max Then lngValue = .Max
            If .Value <> lngValue Then .Value = lngValue
        End If
    End With
End Sub

Private Sub picHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If grid.CurrentRow <> 0 Or grid.CurrentCol <> 0 Then ActiveCell 0, 0
    If grid.CurrentSlot <> 0 Then Me.usrIconSlot(grid.CurrentSlot).Active = False
    ActiveSlot X, Y
End Sub

Private Sub picHeader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveSlot X, Y
End Sub

Private Sub ActiveSlot(X As Single, Y As Single)
    Dim lngSlot As Long
    Dim i As Long
    
    If Y <= grid.IconTop Or Y >= grid.IconTop + grid.IconHeight Then
        lngSlot = 0
    Else
        lngSlot = (X - grid.EffectWidth) \ grid.SlotWidth + 1
        If lngSlot < 1 Then
            lngSlot = 0
        ElseIf lngSlot > grid.Slots Then
            lngSlot = grid.Slots
        ElseIf X < grid.Slot(lngSlot).IconLeft Or X > grid.Slot(lngSlot).IconLeft + grid.IconWidth Then
            lngSlot = 0
        End If
    End If
    If lngSlot Then xp.SetMouseCursor mcHand
    If grid.CurrentSlot <> 0 And grid.CurrentSlot <> lngSlot Then Me.usrIconSlot(grid.CurrentSlot).Active = False
    grid.CurrentSlot = lngSlot
    If lngSlot Then Me.usrIconSlot(grid.CurrentSlot).Active = True
End Sub

Private Sub picTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.picTooltip.Visible Then Me.picTooltip.Visible = False
    If mlngTab = 2 Then ActiveCell 0, 0
End Sub

Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell 0, 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngTab = 2 Then ActiveCell 0, 0
End Sub

Private Sub usrHeader_MouseOver()
    If mlngTab = 2 Then ActiveCell 0, 0
End Sub

Private Sub usrFooter_MouseOver()
    If mlngTab = 2 Then ActiveCell 0, 0
End Sub


' ************* REVIEW *************


Private Sub chkOutput_Click(Index As Integer)
    Dim enReview As ReviewEnum
    
    If UncheckButton(Me.chkOutput(Index), mblnOverride) Then Exit Sub
    Me.picTab(3).SetFocus
    Me.usrchkBound.Enabled = False
    Select Case Me.chkOutput(Index).Caption
        Case "Save": SaveFile
        Case "Load": LoadFile
        Case Else: Review Index
    End Select
End Sub

Private Sub usrchkBound_UserChange()
    ShowRecipe
End Sub

Private Sub Review(ByVal penReview As ReviewEnum)
    InitLevels
    Me.picReviewItemList.Visible = False
    Me.picReviewIngredients.Visible = False
    Me.picReviewOutput.Visible = False
    menReview = penReview
    Select Case menReview
        Case reItems: ReviewItemList
        Case reIngredients: ReviewIngredients
        Case reOutputText: OutputText
        Case reOutputForums: OutputForums
    End Select
End Sub

Private Sub InitLevels()
    Dim i As Long
    
    For i = 0 To seSlotCount - 1
        If gs.Item(i).Crafted And gs.Item(i).ML = 0 Then gs.Item(i).ML = gs.BaseLevel
    Next
End Sub

Private Sub ReviewItemList()
    Dim enSlot As SlotEnum
    Dim lngTop As Long
    Dim blnEldritch As Boolean
    Dim i As Long
    
    xp.Mouse = msAppWait
    For i = 0 To Me.usrSlot.UBound
        Me.usrSlot(i).Visible = False
    Next
    For enSlot = 0 To seSlotCount - 1
        If gs.Item(enSlot).Crafted Or Len(gs.Item(enSlot).Named) <> 0 Then
            If gs.Item(enSlot).EldritchRitual <> 0 Then
                blnEldritch = True
                Exit For
            End If
        End If
    Next
    mlngSlots = -1
    For enSlot = 0 To seSlotCount - 1
        If gs.Item(enSlot).Crafted Or Len(gs.Item(enSlot).Named) <> 0 Then
            mlngSlots = mlngSlots + 1
            If mlngSlots > Me.usrSlot.UBound Then Load Me.usrSlot(mlngSlots)
            Me.usrSlot(mlngSlots).Top = lngTop
            lngTop = lngTop + Me.usrSlot(0).Height - PixelY
            Me.usrSlot(mlngSlots).EldritchLeaveSpace = blnEldritch
            ReviewItem Me.usrSlot(mlngSlots), gs.Item(enSlot), enSlot
            Me.usrSlot(mlngSlots).Visible = True
        End If
    Next
    If mlngSlots <> -1 Then
        With Me.usrSlot(mlngSlots)
            Me.picReview.Height = .Top + .Height + PixelY
        End With
        With Me.scrollVerticalReview
            If Me.picReview.Height > Me.picReviewItemList.Height Then
                .Max = Me.picReview.Height - Me.picReviewItemList.Height
                .SmallChange = Me.usrSlot(0).Height - PixelY
                .LargeChange = Me.picReviewItemList.Height
                .Value = 0
                .Visible = True
            Else
                .Visible = False
            End If
        End With
    End If
    Me.picReviewItemList.Visible = True
    xp.Mouse = msNormal
End Sub

Private Sub ReviewItem(pctl As userItemSlot, ptypItemSlot As ItemSlotType, penSlot As SlotEnum)
    With ptypItemSlot
        pctl.Clear
        pctl.RefreshColors
        pctl.Slot = penSlot
        pctl.Gear = .Gear
        pctl.ItemStyle = .ItemStyle
        pctl.Crafted = .Crafted
        pctl.Named = .Named
        pctl.SetML .ML, .MLDone
        pctl.SetEffects .Effect, .EffectDone
        pctl.SetAugments GearsetAugmentToString(.Augment)
        pctl.SetEldritch .EldritchRitual, .EldritchDone
        pctl.Refresh
    End With
End Sub

Private Sub usrSlot_LevelChange(Index As Integer, ML As Long)
    With gs.Item(Me.usrSlot(Index).Slot)
        If .ML <> ML Then
            .ML = ML
            Dirty True, False
        End If
    End With
End Sub

Private Sub usrSlot_MLDone(Index As Integer, Value As Boolean)
    With gs.Item(Me.usrSlot(Index).Slot)
        If .MLDone <> Value Then
            .MLDone = Value
            Dirty True, False
        End If
    End With
End Sub

Private Sub usrSlot_EffectDone(Index As Integer, ByVal Effect As Long, Value As Boolean)
    With gs.Item(Me.usrSlot(Index).Slot)
        If .EffectDone(Effect) <> Value Then
            .EffectDone(Effect) = Value
            Dirty True, False
        End If
    End With
End Sub

Private Sub usrSlot_AugmentDone(Index As Integer, ByVal Color As AugmentColorEnum, Value As Boolean)
    With gs.Item(Me.usrSlot(Index).Slot)
        If .Augment(Color).Done <> Value Then
            .Augment(Color).Done = Value
            Dirty True, False
        End If
    End With
End Sub

Private Sub usrSlot_EldritchDone(Index As Integer, Value As Boolean)
    With gs.Item(Me.usrSlot(Index).Slot)
        If .EldritchDone <> Value Then
            .EldritchDone = Value
            Dirty True, False
        End If
    End With
End Sub

Private Sub usrSlot_GearClick(Index As Integer, Gear As GearEnum, ML As Long, Prefix As Long, Suffix As Long, Extra As Long, Augment As String, Eldritch As Long)
    OpenItem Me.usrSlot(Index).Slot
End Sub

Private Sub scrollVerticalReview_GotFocus()
    Me.picReviewItemList.SetFocus
End Sub

Private Sub scrollVerticalReview_Change()
    ScrollReview
End Sub

Private Sub scrollVerticalReview_Scroll()
    ScrollReview
End Sub

Private Sub ScrollReview()
    Me.picReview.Top = -Me.scrollVerticalReview.Value
End Sub

Private Sub ScrollItemList(ByVal plngValue As Long)
    Dim lngValue As Long
    
    plngValue = plngValue \ 3
    With Me.scrollVerticalReview
        If .Visible Then
            lngValue = .Value - (plngValue * .SmallChange)
            If lngValue < .Min Then lngValue = .Min
            If lngValue > .Max Then lngValue = .Max
            If .Value <> lngValue Then .Value = lngValue
        End If
    End With
End Sub

Private Sub ReviewIngredients()
    Dim blnEffect() As Boolean
    Dim enSlot As SlotEnum
    Dim i As Long
    
    ListboxClear Me.lstReviewShards
    ListboxAddItem Me.lstReviewShards, "Everything", ieEverything
    ListboxAddItem Me.lstReviewShards, "All Shards", ieShards
    ListboxAddItem Me.lstReviewShards, "All Effect Shards", ieEffects
    ListboxAddItem Me.lstReviewShards, "All Minimum Level Shards", ieML
    ListboxAddItem Me.lstReviewShards, "All Augments", ieAugments
    ListboxAddItem Me.lstReviewShards, "All Eldritch Rituals", ieEldritch
    GetEffects blnEffect, ieEffects
    For i = 1 To db.Shards
        If blnEffect(i) Then ListboxAddItem Me.lstReviewShards, db.Shard(i).ShardName, i
    Next
    For enSlot = 0 To seSlotCount - 1
        With gs.Item(enSlot)
            If (.Crafted = True Or Len(.Named) > 0) And .EldritchRitual > 0 And .EldritchDone = False Then
                ListboxAddItem Me.lstReviewShards, "Eldritch " & db.Ritual(.EldritchRitual).RitualName, 1000 + enSlot
            End If
        End With
    Next
    Me.lstReviewShards.ListIndex = 0
    Me.usrchkBound.Enabled = True
    Me.picReviewIngredients.Visible = True
End Sub

Private Sub lstReviewShards_Click()
    ShowRecipe
End Sub

Private Sub ShowRecipe()
    Dim lngItemData As Long
    
    Me.usrReviewIngredients.Clear
    lngItemData = ListboxGetValue(Me.lstReviewShards)
    If lngItemData = 0 Then Exit Sub
    Select Case lngItemData
        Case 1 To 899: ShowIngredientsShard lngItemData
        Case 1000 To 1999: ShowIngredientsRitual lngItemData - 1000
        Case ieEffects: ShowIngredientsEffects
        Case ieML: ShowIngredientsML
        Case ieShards: ShowIngredientsShards
        Case ieEldritch: ShowIngredientsRituals
        Case ieAugments: ShowIngredientsAugments
        Case ieEverything: ShowIngredientsEverything
    End Select
End Sub

' Specific effect shard
Private Sub ShowIngredientsShard(plngShard As Long)
    Dim enSlot As SlotEnum
    Dim strSlot As String
    Dim enAffix As AffixEnum
    Dim strAffix As String
    Dim blnEffect() As Boolean
    Dim typRecipe As RecipeType
    
    ' Shard name
    With db.Shard(plngShard)
        Me.usrReviewIngredients.AddLink .ShardName, lseShard, .ShardName, 2
    End With
    ' Where this shard is slotted into gearset
    Do
        For enSlot = 0 To seSlotCount - 1
            For enAffix = aePrefix To aeExtra
                If gs.Item(enSlot).Crafted And gs.Item(enSlot).Effect(enAffix) = plngShard Then Exit Do
            Next
        Next
        Exit Sub
    Loop Until True
    Select Case enAffix
        Case aePrefix: strAffix = "Prefix"
        Case aeSuffix: strAffix = "Suffix"
        Case aeExtra: strAffix = "Extra"
    End Select
    strSlot = GetSlotName(enSlot)
    Me.usrReviewIngredients.AddText strSlot & " " & strAffix, 2
    ' Shard recipe
    GetEffects blnEffect, plngShard
    GetShardRecipes typRecipe, blnEffect
    AggregateRecipeSort typRecipe
    ' Display
    Me.usrReviewIngredients.AddText BindingText(Me.usrchkBound.Value) & " Crafting:"
    Me.usrReviewIngredients.AddText "Crafting Level " & typRecipe.Level, 1, "   "
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
End Sub

' All effect shards
Private Sub ShowIngredientsEffects()
    Dim blnEffect() As Boolean
    Dim typRecipe As RecipeType
    Dim strClipboard As String
    
    ' Gearset name
    If Len(mstrFile) <> 0 Then Me.usrReviewIngredients.AddText mstrFile, 2
    ' Aggregate recipes
    GetEffects blnEffect, ieEffects
    GetShardRecipes typRecipe, blnEffect
    AggregateRecipeSort typRecipe
    ' Clipboard
    strClipboard = GetIngredientsClipboard(typRecipe)
    ' Display
    Me.usrReviewIngredients.AddText BindingText(Me.usrchkBound.Value) & " Crafting:", 0
    Me.usrReviewIngredients.AddClipboard strClipboard
    Me.usrReviewIngredients.AddText "Crafting Level " & typRecipe.Level, 1, "   "
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
End Sub

' All crafting shards (effects and MLs, but not rituals or augments)
Private Sub ShowIngredientsShards()
    Dim typRecipe As RecipeType
    Dim blnEffect() As Boolean
    Dim strClipboard As String
    
    ' Gearset name
    If Len(mstrFile) <> 0 Then Me.usrReviewIngredients.AddText mstrFile, 2
    ' Aggregate recipes
    GetEffects blnEffect, ieShards
    GetShardRecipes typRecipe, blnEffect
    GetMLRecipes typRecipe
    AggregateRecipeSort typRecipe
    ' Clipboard
    strClipboard = GetIngredientsClipboard(typRecipe)
    ' Display
    Me.usrReviewIngredients.AddText BindingText(Me.usrchkBound.Value) & " Crafting:", 0
    Me.usrReviewIngredients.AddClipboard strClipboard
    Me.usrReviewIngredients.AddText "Crafting Level " & typRecipe.Level, 1, "   "
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
End Sub

' All ML shards
Private Sub ShowIngredientsML()
    Dim lngML(1 To MAX_CRAFTING_LEVEL) As Long
    Dim blnShow As Boolean
    Dim lngPadding As Long
    Dim strPlural As String
    Dim typRecipe As RecipeType
    Dim i As Long
    
    ' Gearset name
    If Len(mstrFile) <> 0 Then Me.usrReviewIngredients.AddText mstrFile, 2
    ' Count ML shards
    For i = 0 To seSlotCount - 1
        With gs.Item(i)
            If .Crafted = True And .ML <> 0 And .MLDone = False Then
                lngML(.ML) = lngML(.ML) + 1
                blnShow = True
            End If
        End With
    Next
    If Not blnShow Then Exit Sub
    ' Formatting
    lngPadding = 1
    For i = 1 To MAX_CRAFTING_LEVEL
        If lngML(i) > 9 Then lngPadding = 2
    Next
    ' ML Shards
    For i = 1 To MAX_CRAFTING_LEVEL
        If lngML(i) <> 0 Then
            Me.usrReviewIngredients.AddNumber lngML(i), lngPadding
            If lngML(i) = 1 Then strPlural = vbNullString Else strPlural = "s"
            Me.usrReviewIngredients.AddText "ML" & i & " shard" & strPlural
        End If
    Next
    Me.usrReviewIngredients.AddText vbNullString
    ' Aggregate recipes
    GetMLRecipes typRecipe
    ' Display
    Me.usrReviewIngredients.AddText BindingText(Me.usrchkBound.Value) & " Crafting:"
    Me.usrReviewIngredients.AddText "Crafting Level " & typRecipe.Level, 1, "   "
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
End Sub

' Specific ritual (only one ritual per item, so if we know the slot we know the ritual)
Private Sub ShowIngredientsRitual(penSlot As SlotEnum)
    Dim lngRitual As Long
    Dim blnEffect() As Boolean
    Dim typRecipe As RecipeType
    
    lngRitual = gs.Item(penSlot).EldritchRitual
    Me.usrReviewIngredients.AddLink "Eldritch " & db.Ritual(lngRitual).RitualName, lseForm, "frmEldritch", 2
    Me.usrReviewIngredients.AddText GetSlotName(penSlot)
    ' Ritual recipe
    AggregateRecipe typRecipe, db.Ritual(lngRitual).Recipe
    AggregateRecipeSort typRecipe
    ' Display
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
End Sub

' All rituals
Private Sub ShowIngredientsRituals()
    Dim typRecipe As RecipeType
    Dim strClipboard As String
    
    ' Gearset name
    If Len(mstrFile) <> 0 Then Me.usrReviewIngredients.AddText mstrFile, 2
    GetRituals typRecipe
    AggregateRecipeSort typRecipe
    strClipboard = GetIngredientsClipboard(typRecipe, True)
    ' Display
    Me.usrReviewIngredients.AddText "Stone of Change recipes:", 0
    Me.usrReviewIngredients.AddClipboard strClipboard, 0
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
End Sub

Private Sub ShowIngredientsAugments()
    Dim typAugment() As IngredientAugmentType
    Dim lngAugments As Long
    Dim strClipboard As String
    Dim i As Long
    
    ' Gearset name
    If Len(mstrFile) <> 0 Then Me.usrReviewIngredients.AddText mstrFile, 2
    ' Get sorted augment list
    lngAugments = GetAugments(typAugment)
    ' Clipboard
    Me.usrReviewIngredients.AddText "Augments", 0
    strClipboard = GetAugmentsClipboard(typAugment, lngAugments, False)
    Me.usrReviewIngredients.AddClipboard strClipboard
    ' Augments
    For i = 1 To lngAugments
        With typAugment(i)
            Me.usrReviewIngredients.AddAugment .Augment, .Variation, .Scaling, , "   "
        End With
    Next
End Sub

Private Sub ShowIngredientsEverything()
    Dim blnEffect() As Boolean
    Dim typRecipe As RecipeType
    Dim typAugment() As IngredientAugmentType
    Dim lngAugments As Long
    Dim strClipboard As String
    Dim lngLeft As Long
    Dim i As Long
    
    ' Gearset name
    If Len(mstrFile) <> 0 Then Me.usrReviewIngredients.AddText mstrFile, 2
    ' Gather data
    GetEffects blnEffect, ieShards
    GetShardRecipes typRecipe, blnEffect
    GetMLRecipes typRecipe
    GetRituals typRecipe
    AggregateRecipeSort typRecipe
    lngAugments = GetAugments(typAugment)
    ' Clipboard
    strClipboard = GetIngredientsClipboard(typRecipe) & GetAugmentsClipboard(typAugment, lngAugments, True)
    ' Display
    Me.usrReviewIngredients.AddText BindingText(Me.usrchkBound.Value) & " Crafting:", 0
    Me.usrReviewIngredients.AddClipboard strClipboard
    Me.usrReviewIngredients.AddText "Crafting Level " & typRecipe.Level, 1, "   "
    AddRecipeToInfo typRecipe, Me.usrReviewIngredients, "   "
    Me.usrReviewIngredients.AddText vbNullString
    ' Augments
    lngLeft = Me.usrReviewIngredients.LastLinkLeft
    For i = 1 To lngAugments
        With typAugment(i)
            Me.usrReviewIngredients.AddAugment .Augment, .Variation, .Scaling, , , lngLeft
        End With
    Next
    Me.usrReviewIngredients.AddText vbNullString
End Sub

' Set boolean array mirroring db.Shards() where any shards to be included are flagged True
Private Sub GetEffects(pblnEffect() As Boolean, plngItemData As Long)
    Dim lngItem As Long
    Dim enAffix As AffixEnum
    Dim enLastAffix As AffixEnum
    Dim lngShard As Long
    Dim blnInclude As Boolean
    
    ReDim pblnEffect(1 To db.Shards)
    Select Case plngItemData
        Case ieEverything, ieShards, ieEffects: blnInclude = True
    End Select
    For lngItem = 0 To seSlotCount - 1
        If gs.Item(lngItem).Crafted Then
            If gs.Item(lngItem).ML < 10 Then enLastAffix = aeSuffix Else enLastAffix = aeExtra
            For enAffix = 0 To enLastAffix
                If Not gs.Item(lngItem).EffectDone(enAffix) Then
                    lngShard = gs.Item(lngItem).Effect(enAffix)
                    If lngShard Then
                        If blnInclude = True Or lngShard = plngItemData Then
                            If db.Shard(lngShard).ML <= gs.Item(lngItem).ML Then pblnEffect(lngShard) = True
                        End If
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub GetShardRecipes(ptypRecipe As RecipeType, pblnEffect() As Boolean)
    Dim blnBound As Boolean
    Dim i As Long
    
    blnBound = Me.usrchkBound.Value
    For i = 1 To db.Shards
        If pblnEffect(i) Then
            If blnBound Then
                AggregateRecipe ptypRecipe, db.Shard(i).Bound
            Else
                AggregateRecipe ptypRecipe, db.Shard(i).Unbound
            End If
        End If
    Next
End Sub

Private Sub GetMLRecipes(ptypRecipe As RecipeType)
    Dim enSlot As SlotEnum
    Dim blnBound As Boolean
    
    blnBound = Me.usrchkBound.Value
    For enSlot = 0 To seSlotCount - 1
        With gs.Item(enSlot)
            If .Crafted And Not .MLDone Then AggregateRecipeML ptypRecipe, .ML, blnBound
        End With
    Next
End Sub

Private Sub GetRituals(ptypRecipe As RecipeType)
    Dim enSlot As SlotEnum
    
    For enSlot = 0 To seSlotCount - 1
        With gs.Item(enSlot)
            If (.Crafted = True Or Len(.Named) > 0) And .EldritchRitual > 0 And .EldritchDone = False Then
                AggregateRecipe ptypRecipe, db.Ritual(.EldritchRitual).Recipe
            End If
        End With
    Next
End Sub

Private Function GetIngredientsClipboard(ptypRecipe As RecipeType, Optional pblnSkipEssences As Boolean = False) As String
    Dim strReturn As String
    Dim i As Long
    
'    If Len(mstrFile) Then
'        strReturn = mstrFile & " (" & BindingText(Me.usrchkBound.Value) & ")"
'    Else
'        strReturn = BindingText(Me.usrchkBound.Value) & " Crafting"
'    End If
'    strReturn = strReturn & vbNewLine
'    If Not pblnSkipEssences Then strReturn = strReturn & FormatEssences(ptypRecipe.Essences) & vbTab & "Essences" & vbNewLine
    If Not pblnSkipEssences Then strReturn = FormatEssences(ptypRecipe.Essences) & vbTab & "Essences" & vbNewLine
    For i = 1 To ptypRecipe.Ingredients
        strReturn = strReturn & ptypRecipe.Ingredient(i).Count & vbTab & Pluralized(ptypRecipe.Ingredient(i)) & vbNewLine
    Next
    GetIngredientsClipboard = strReturn
End Function

Private Function GetAugments(ptypAugment() As IngredientAugmentType) As Long
    Dim enSlot As Long
    Dim lngCount As Long
    Dim lngAugments As Long
    Dim lngColorOrder(1 To 7) As Long
    Dim lngScale As Long
    Dim i As Long
    
    ReDim ptypAugment(1 To 64)
    lngColorOrder(aceColorless) = 1
    lngColorOrder(aceYellow) = 2
    lngColorOrder(aceGreen) = 3
    lngColorOrder(aceBlue) = 4
    lngColorOrder(acePurple) = 5
    lngColorOrder(aceRed) = 6
    lngColorOrder(aceOrange) = 7
    For enSlot = 0 To seSlotCount - 1
        lngCount = 0
        For i = 1 To 7
            With gs.Item(enSlot).Augment(i)
                If .Exists = True And .Augment <> 0 And .Variation <> 0 And .Done = False And lngCount < 3 Then
                    For lngScale = db.Augment(.Augment).Scalings To 1 Step -1
                        If db.Augment(.Augment).Scaling(lngScale).ML <= gs.Item(enSlot).ML Then Exit For
                    Next
                    If lngScale > 0 Then
                        lngCount = lngCount + 1
                        lngAugments = lngAugments + 1
                        ptypAugment(lngAugments).Color = db.Augment(.Augment).Color
                        ptypAugment(lngAugments).ColorOrder = lngColorOrder(db.Augment(.Augment).Color)
                        ptypAugment(lngAugments).Augment = .Augment
                        ptypAugment(lngAugments).Variation = .Variation
                        ptypAugment(lngAugments).Scaling = lngScale
                        ptypAugment(lngAugments).FullName = AugmentFullName(.Augment, .Variation, lngScale)
                    End If
                End If
            End With
        Next
    Next
    If lngAugments = 0 Then
        Erase ptypAugment
    Else
        ReDim Preserve ptypAugment(1 To lngAugments)
        SortIngredientAugments ptypAugment, lngAugments
        GetAugments = lngAugments
    End If
End Function

' Comb sort because it's simpler for dropping in a udt array
Private Sub SortIngredientAugments(ptypAugment() As IngredientAugmentType, plngAugments As Long)
    Const ShrinkFactor = 1.3
    Dim lngGap As Long
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As IngredientAugmentType
    Dim blnSwapped As Boolean
    
    iMin = 1
    iMax = plngAugments
    lngGap = iMax - iMin + 1
    Do
        If lngGap > 1 Then
            lngGap = Int(lngGap / ShrinkFactor)
            If lngGap = 10 Or lngGap = 9 Then lngGap = 11
        End If
        blnSwapped = False
        For i = iMin To iMax - lngGap
            If (ptypAugment(i).ColorOrder > ptypAugment(i + lngGap).ColorOrder) Or (ptypAugment(i).ColorOrder = ptypAugment(i + lngGap).ColorOrder And ptypAugment(i).FullName > ptypAugment(i + lngGap).FullName) Then
                typSwap = ptypAugment(i)
                ptypAugment(i) = ptypAugment(i + lngGap)
                ptypAugment(i + lngGap) = typSwap
                blnSwapped = True
            End If
        Next
    Loop Until lngGap = 1 And Not blnSwapped
End Sub

Private Function GetAugmentsClipboard(ptypAugment() As IngredientAugmentType, plngAugments As Long, pblnTabs As Boolean) As String
    Dim strLine() As String
    Dim strTab As String
    Dim i As Long
    
    If pblnTabs Then strTab = vbTab
    ReDim strLine(1 To plngAugments)
    For i = 1 To plngAugments
        With ptypAugment(i)
            strLine(i) = AugmentFullName(.Augment, .Variation, .Scaling)
        End With
    Next
    GetAugmentsClipboard = strTab & Join(strLine, vbNewLine & strTab) & vbNewLine
End Function

Private Sub lstReviewShards_DblClick()
    If Me.lstReviewShards.ListIndex = -1 Then Exit Sub
    With Me.lstReviewShards
        Select Case .ItemData(.ListIndex)
            Case 1 To ieEverything - 1: OpenShard db.Shard(.ItemData(.ListIndex)).ShardName
        End Select
    End With
End Sub

Private Sub OutputText()
    Dim enSlot As SlotEnum
    Dim strTitle As String
    Dim strBody As String
    Dim strSlot As String
    Dim strPrefix As String
    Dim strSuffix As String
    Dim strExtra As String
    Dim strAugments As String
    Dim strLine As String
    
    SetColors cgeOutput
    If Len(mstrFile) Then strTitle = mstrFile Else strTitle = "New Gearset"
    For enSlot = 0 To seSlotCount - 1
        strLine = vbNullString
        With gs.Item(enSlot)
            Do
                If .Crafted Then
                    If .Effect(0) <> 0 Or .Effect(1) <> 0 Or .Effect(2) <> 0 Then
                        strLine = GetSlotDisplay(enSlot) & ": ML" & .ML & " "
                        strSlot = GetItemDisplay(enSlot, .ML)
                        strPrefix = GetEffectDisplay(enSlot, aePrefix)
                        strSuffix = GetEffectDisplay(enSlot, aeSuffix)
                        strExtra = GetEffectDisplay(enSlot, aeExtra)
                        If Len(strPrefix) Then strLine = strLine & strPrefix & " "
                        strLine = strLine & strSlot
                        If Len(strSuffix) Then strLine = strLine & " of " & strSuffix
                        If Len(strExtra) Then strLine = strLine & " w/" & strExtra
                    End If
                ElseIf Len(.Named) Then
                    strLine = GetSlotDisplay(enSlot) & ": " & .Named
                Else
                    Exit Do
                End If
                strLine = strLine & OutputTextAugments(.Augment, .ML, False)
            Loop Until True
        End With
        If Len(strLine) Then strBody = strBody & strLine & vbNewLine
    Next
    Me.usrReviewOutput.AddText strTitle, 0
    Me.usrReviewOutput.AddClipboard strBody, 2
    Me.usrReviewOutput.AddText strBody
    Me.picReviewOutput.Visible = True
End Sub

Private Sub SetColors(penGroup As ColorGroupEnum)
    With Me.usrReviewOutput
        .TextColor = cfg.GetColor(penGroup, cveText)
        .BackColor = cfg.GetColor(penGroup, cveBackground)
        .LinkColor = cfg.GetColor(penGroup, cveTextLink)
        .Clear
    End With
End Sub

Private Function OutputTextAugments(ptypAugment() As AugmentSlotType, plngML As Long, pblnBBCodes As Boolean) As String
    Dim strReturn As String
    Dim lngCount As Long
    Dim i As Long
    
    For i = 1 To 7
        With ptypAugment(i)
            If .Exists = True And .Augment <> 0 And .Variation <> 0 Then
                lngCount = lngCount + 1
                If lngCount = 1 Then strReturn = " (" Else strReturn = strReturn & ", "
                strReturn = strReturn & AugmentScaledName(.Augment, .Variation, plngML)
            End If
        End With
        If lngCount = 3 Then Exit For
    Next
    If lngCount Then OutputTextAugments = strReturn & ")"
End Function

Private Function GetSlotDisplay(penSlot As SlotEnum) As String
    Dim strReturn As String
    
    Select Case penSlot
        Case seHelmet: strReturn = "Head"
        Case seGoggles: strReturn = "Eyes"
        Case seNecklace: strReturn = "Neck"
        Case seCloak: strReturn = "Back"
        Case seBracers: strReturn = "Wrist"
        Case seGloves: strReturn = "Hand"
        Case seBelt: strReturn = "Waist"
        Case seBoots: strReturn = "Feet"
        Case seRing1: strReturn = "Ring"
        Case seRing2: strReturn = "Ring"
        Case seTrinket: strReturn = "Trinket"
        Case seArmor: strReturn = "Body"
        Case seMainHand: strReturn = "Main"
        Case seOffHand: strReturn = "Off"
    End Select
    GetSlotDisplay = strReturn
End Function

Private Function GetItemDisplay(penSlot As SlotEnum, plngML As Long)
    Dim lngIndex As Long
    
    Select Case penSlot
        Case seMainHand, seOffHand, seArmor
            GetItemDisplay = gs.Item(penSlot).ItemStyle
            lngIndex = SeekItem(gs.Item(penSlot).ItemStyle)
            If lngIndex = 0 Then Exit Function
            If Not db.Item(lngIndex).Scales Then Exit Function
            With db.Item(lngIndex)
                Select Case plngML
                    Case Is < 4: GetItemDisplay = .Scaling(0)
                    Case 4 To 9: GetItemDisplay = .Scaling(1)
                    Case 10 To 15: GetItemDisplay = .Scaling(2)
                    Case 16 To 21: GetItemDisplay = .Scaling(3)
                    Case Is > 21: GetItemDisplay = .Scaling(4)
                End Select
            End With
        Case seRing1, seRing2
            GetItemDisplay = "Ring"
        Case Else
            GetItemDisplay = GetSlotName(penSlot)
    End Select
End Function

Private Function GetEffectDisplay(penSlot As SlotEnum, penAffix As AffixEnum) As String
    Dim strShard As String
    Dim strScale As String
    Dim strReturn As String
    
    strShard = GetEffectName(penSlot, penAffix)
    strScale = GetEffectScale(penSlot, penAffix)
    If Len(strShard) Then
        If Len(strScale) Then
            strReturn = strShard & " " & strScale
        Else
            strReturn = strShard
        End If
    End If
    GetEffectDisplay = strReturn
End Function

Private Function GetEffectName(penSlot As SlotEnum, penAffix As AffixEnum) As String
    Dim lngShard As Long
    
    lngShard = gs.Item(penSlot).Effect(penAffix)
    If lngShard Then GetEffectName = db.Shard(lngShard).GridName
End Function

Private Function GetEffectScale(penSlot As SlotEnum, penAffix As AffixEnum) As String
    Dim lngShard As Long
    Dim lngML As Long
    Dim lngScale As Long
    
    lngShard = gs.Item(penSlot).Effect(penAffix)
    If lngShard = 0 Then Exit Function
    lngML = gs.Item(penSlot).ML
    With db.Shard(lngShard)
        If .ScaleName <> "None" Then
            lngScale = SeekScaling(.ScaleName)
            If lngScale Then
                With db.Scaling(lngScale)
                    GetEffectScale = .Table(lngML)
                End With
            End If
        End If
    End With
End Function

Private Sub OutputForums()
    Dim strTitle As String
    Dim enSlot As SlotEnum
    Dim strSlot As String
    Dim strML As String
    Dim strPrefix As String
    Dim strSuffix As String
    Dim strExtra As String
    Dim strItem As String
    Dim blnFixed As Boolean
    Dim blnInclude As Boolean
    Dim blnCrafted As Boolean
    Dim blnTextColor As Boolean
    Dim lngDim As Long
    
    SetColors cgeOutput
    If AllowFixed Then cfg.GetOutputFixed blnFixed, vbNullString, vbNullString
    cfg.GetOutputColor blnTextColor, vbNullString, vbNullString
    If blnTextColor Then blnTextColor = cfg.OutputTextColors
    If blnTextColor Then lngDim = cfg.GetColor(cgeOutput, cveTextDim) Else lngDim = cfg.GetColor(cgeOutput, cveText)
    If Len(mstrFile) Then strTitle = mstrFile Else strTitle = "New Gearset"
    Me.usrReviewOutput.AddText strTitle, 0
    Me.usrReviewOutput.AddClipboard ClipboardForums(), 2
    For enSlot = 0 To seSlotCount - 1
        blnInclude = False
        With gs.Item(enSlot)
            strSlot = GetSlotDisplay(enSlot)
            If blnFixed And Len(strSlot) = 4 Then strSlot = strSlot & " "
            strSlot = strSlot & ":"
            strItem = GetItemDisplay(enSlot, .ML)
            If .Crafted Then
                If .Effect(0) <> 0 Or .Effect(1) <> 0 Or .Effect(2) <> 0 Then
                    blnInclude = True
                    blnCrafted = True
                    strML = "ML" & .ML
                    strPrefix = GetEffectDisplay(enSlot, aePrefix)
                    strSuffix = GetEffectDisplay(enSlot, aeSuffix)
                    strExtra = GetEffectDisplay(enSlot, aeExtra)
                End If
            ElseIf Len(.Named) Then
                blnInclude = True
                blnCrafted = False
            End If
            If blnInclude Then
                Me.usrReviewOutput.AddTextFormatted strSlot, False, False, False, -1, 0, vbNullString, vbNullString, blnFixed
                If blnCrafted Then
                    Me.usrReviewOutput.AddTextFormatted strML, False, False, False, lngDim, 0
                    If Len(strPrefix) Then Me.usrReviewOutput.AddText strPrefix, 0
                    If Len(strSuffix) Then strItem = strItem & " of"
                    Me.usrReviewOutput.AddTextFormatted strItem, False, False, False, lngDim, 0
                    If Len(strSuffix) Then Me.usrReviewOutput.AddText strSuffix, 0
                    If Len(strExtra) Then
                        Me.usrReviewOutput.AddTextFormatted "w/", False, False, False, lngDim, 0
                        Me.usrReviewOutput.BackupOneSpace
                        Me.usrReviewOutput.AddText strExtra, 0
                    End If
                Else
                    Me.usrReviewOutput.AddText .Named, 0
                End If
                OutputForumsAugments .Augment, .ML, blnTextColor
                Me.usrReviewOutput.AddText vbNullString
            End If
        End With
    Next
    Me.picReviewOutput.Visible = True
End Sub

Private Sub OutputForumsAugments(ptypAugment() As AugmentSlotType, plngML As Long, pblnBBCodes As Boolean)
    Dim strText As String
    Dim lngCount As Long
    Dim lngLast As Long
    Dim i As Long
    
    For i = 1 To 7
        With ptypAugment(i)
            If .Exists = True And .Augment <> 0 And .Variation <> 0 Then lngLast = lngLast + 1
        End With
    Next
    If lngLast > 3 Then lngLast = 3
    For i = 1 To 7
        With ptypAugment(i)
            If .Exists = True And .Augment <> 0 And .Variation <> 0 Then
                lngCount = lngCount + 1
                If lngCount = 1 Then strText = "(" Else strText = vbNullString
                strText = strText & AugmentScaledName(.Augment, .Variation, plngML)
                If lngCount = lngLast Then strText = strText & ")" Else strText = strText & ","
                Me.usrReviewOutput.AddTextFormatted strText, False, , , cfg.GetColor(cgeOutput, GetAugmentColorValue(i)), 0
            End If
        End With
        If lngCount = 3 Then Exit For
    Next
End Sub

'Replace(mstrColorOpen, "$", xp.ColorToHex(cfg.GetColor(cgeOutput, cveBackground)))
Private Function ClipboardForums() As String
    Dim enSlot As SlotEnum
    Dim strSlot As String
    Dim strML As String
    Dim strPrefix As String
    Dim strSuffix As String
    Dim strExtra As String
    Dim strItem As String
    Dim blnFixed As Boolean
    Dim strFixedOpen As String
    Dim strFixedClose As String
    Dim blnInclude As Boolean
    Dim blnCrafted As Boolean
    Dim blnTextColor As Boolean
    Dim strColorOpen As String
    Dim strColorClose As String
    Dim strColor As String
    Dim strReturn As String
    
    If AllowFixed Then cfg.GetOutputFixed blnFixed, strFixedOpen, strFixedClose
    If Not blnFixed Then
        strFixedOpen = vbNullString
        strFixedClose = vbNullString
    End If
    cfg.GetOutputColor blnTextColor, strColorOpen, strColorClose
    If blnTextColor Then blnTextColor = cfg.OutputTextColors
    If Not blnTextColor Then
        strColorOpen = vbNullString
        strColorClose = vbNullString
    End If
    For enSlot = 0 To seSlotCount - 1
        blnInclude = False
        With gs.Item(enSlot)
            strSlot = GetSlotDisplay(enSlot)
            If blnFixed And Len(strSlot) = 4 Then strSlot = strSlot & " "
            strSlot = strSlot & ": "
            strItem = GetItemDisplay(enSlot, .ML)
            If .Crafted Then
                If .Effect(0) <> 0 Or .Effect(1) <> 0 Or .Effect(2) <> 0 Then
                    blnInclude = True
                    blnCrafted = True
                    strML = "ML" & .ML
                    strPrefix = GetEffectDisplay(enSlot, aePrefix)
                    strSuffix = GetEffectDisplay(enSlot, aeSuffix)
                    strExtra = GetEffectDisplay(enSlot, aeExtra)
                End If
            ElseIf Len(.Named) Then
                blnInclude = True
                blnCrafted = False
            End If
            If blnInclude Then
                strReturn = strReturn & strFixedOpen & strSlot & strFixedClose
                If blnCrafted Then
                    strReturn = strReturn & ColorCode(cveTextDim, strColorOpen) & strML & strColorClose & " "
                    If Len(strPrefix) Then strReturn = strReturn & strPrefix & " "
                    strReturn = strReturn & ColorCode(cveTextDim, strColorOpen) & strItem
                    If Len(strSuffix) Then
                        strReturn = strReturn & " of" & strColorClose & " " & strSuffix
                    Else
                        strReturn = strReturn & strColorClose
                    End If
                    If Len(strExtra) Then
                        strReturn = strReturn & " " & ColorCode(cveTextDim, strColorOpen) & "w/" & strColorClose & strExtra
                    End If
                Else
                    strReturn = strReturn & .Named
                End If
                strReturn = strReturn & ClipboardAugments(.Augment, .ML, blnTextColor) & vbNewLine
            End If
        End With
    Next
    ClipboardForums = strReturn
End Function

Private Function ClipboardAugments(ptypAugment() As AugmentSlotType, plngML As Long, pblnTextColor As Boolean) As String
    Dim strColorOpen As String
    Dim strColorClose As String
    Dim strText As String
    Dim lngCount As Long
    Dim lngLast As Long
    Dim i As Long
    
    If pblnTextColor Then cfg.GetOutputColor True, strColorOpen, strColorClose
    For i = 1 To 7
        With ptypAugment(i)
            If .Exists = True And .Augment <> 0 And .Variation <> 0 Then lngLast = lngLast + 1
        End With
    Next
    If lngLast > 3 Then lngLast = 3
    strText = " "
    For i = 1 To 7
        With ptypAugment(i)
            If .Exists = True And .Augment <> 0 And .Variation <> 0 Then
                lngCount = lngCount + 1
                If pblnTextColor Then strText = strText & ColorCode(GetAugmentColorValue(i), strColorOpen)
                If lngCount = 1 Then strText = strText & "("
                strText = strText & AugmentScaledName(.Augment, .Variation, plngML)
                If lngCount = lngLast Then strText = strText & ")" Else strText = strText & ","
                strText = strText & strColorClose
                If lngCount < lngLast Then strText = strText & " "
            End If
        End With
        If lngCount = 3 Then Exit For
    Next
    ClipboardAugments = strText
End Function

Private Function ColorCode(penColor As ColorValueEnum, pstrTagOpen As String) As String
    If InStr(pstrTagOpen, "$") Then ColorCode = Replace(pstrTagOpen, "$", xp.ColorToHex(cfg.GetColor(cgeOutput, penColor)))
End Function


' ************* FILES *************


Private Sub LoadFile()
    Dim strFile As String
    
    If mblnDirty Then
        Select Case MsgBox("Save gearset first?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Notice")
            Case vbYes
                SaveFile
                Exit Sub
            Case vbCancel
                Exit Sub
        End Select
    End If
    strFile = xp.ShowOpenDialog(cfg.CraftingPath, "Gearsets|*." & GearsetExt, GearsetExt)
    If Len(strFile) Then OpenGearsetFile strFile, mlngTab
End Sub

Private Sub OpenGearsetFile(ByVal pstrFile As String, plngTab As Long)
    Dim i As Long
    
    cfg.CraftingPath = GetPathFromFilespec(pstrFile)
    mstrFile = GetNameFromFilespec(pstrFile)
    If LoadGearset(pstrFile, gs, anal) Then
        For i = 0 To 2
            Me.usrspnBaseML(i).Value = gs.BaseLevel
        Next
        grid.Initialized = False
        ShowGearsetSlots
        MapSlotsToGear gs
        Me.txtNotes.Text = gs.Notes
        If mlngTab <> plngTab Then Me.usrHeader.SetTab plngTab
        If mlngTab = 3 Then Review reItems
        Dirty False, Not gs.Analyzed
    End If
End Sub

Private Sub SaveFile()
    Dim strFile As String
    
    If Len(mstrFile) Then strFile = mstrFile & "." & GearsetExt
    strFile = xp.ShowSaveAsDialog(cfg.CraftingPath, strFile, "Gearsets|*." & GearsetExt, GearsetExt)
    If Len(strFile) = 0 Then Exit Sub
    If xp.File.Exists(strFile) Then
        If MsgBox(GetFileFromFilespec(strFile) & " exists. Overwrite?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    End If
    cfg.CraftingPath = GetPathFromFilespec(strFile)
    mstrFile = GetNameFromFilespec(strFile)
    SaveGearset strFile, gs, anal
    Dirty False, False
End Sub

' Normally, mstrFile contains just the gearset name, not the full filespec
' When opening a new instance with a loaded file, mstrName gets prestuffed with full filespec
' The Form_Load calls OpenGearsetFile, which corrects the value of mstrFile
Public Sub OpenFile(pstrFile As String)
    mstrFile = pstrFile
End Sub


' ************* OPEN ITEM *************


Public Sub OpenItem(penSlot As SlotEnum)
    Dim typOpen As OpenItemType
    Dim frm As Form
    Dim i As Long
    
    With gs.Item(penSlot)
        typOpen.hwnd = Me.hwnd
        typOpen.ML = .ML
        typOpen.Crafted = .Crafted
        typOpen.Named = .Named
        If .ItemStyle = "Handwraps" Then typOpen.Gear = geHandwraps Else typOpen.Gear = .Gear
        typOpen.Slot = penSlot
        typOpen.Prefix = .Effect(aePrefix)
        typOpen.Suffix = .Effect(aeSuffix)
        typOpen.Extra = .Effect(aeExtra)
        For i = 1 To 7
            typOpen.Augment(i) = .Augment(i)
        Next
        typOpen.EldritchRitual = .EldritchRitual
        typOpen.Mainhand = gs.Item(seMainHand).ItemStyle
        typOpen.Offhand = gs.Item(seOffHand).ItemStyle
        typOpen.Armor = gs.Item(seArmor).ItemStyle
    End With
    gtypOpenItem = typOpen
    Set frm = New frmItem
    frm.OpenGearsetItem
    frm.Show
    Set frm = Nothing
End Sub

