VERSION 5.00
Begin VB.Form frmColors 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colors"
   ClientHeight    =   5976
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   8568
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   8568
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5292
      Index           =   0
      Left            =   120
      ScaleHeight     =   5292
      ScaleWidth      =   8292
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   8292
      Begin VB.ListBox lstOutput 
         Height          =   3936
         Left            =   4320
         TabIndex        =   8
         Top             =   720
         Width           =   2412
      End
      Begin VB.ListBox lstScreen 
         Height          =   3936
         Left            =   1380
         TabIndex        =   6
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label lblOutputColors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Output Colors"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4320
         TabIndex        =   7
         Top             =   480
         Width           =   1872
      End
      Begin VB.Label lblScreenColors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Screen Colors"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1380
         TabIndex        =   5
         Top             =   480
         Width           =   1872
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5292
      Index           =   1
      Left            =   120
      ScaleHeight     =   5292
      ScaleWidth      =   8292
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   8292
      Begin VB.ListBox lstMaterial 
         Appearance      =   0  'Flat
         Height          =   4128
         Left            =   420
         TabIndex        =   11
         Top             =   600
         Width           =   1992
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4128
         Left            =   2400
         ScaleHeight     =   4104
         ScaleWidth      =   1968
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   1992
      End
      Begin VB.Frame fraPalette 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2592
         Left            =   4800
         TabIndex        =   79
         Top             =   600
         Width           =   3132
         Begin Colors.userSpinner usrspnGroup 
            Height          =   300
            Index           =   1
            Left            =   1680
            TabIndex        =   17
            Top             =   780
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   529
            Appearance3D    =   -1  'True
            Value           =   6
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Colors.userSpinner usrspnGroup 
            Height          =   300
            Index           =   3
            Left            =   1680
            TabIndex        =   19
            Top             =   1200
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   529
            Appearance3D    =   -1  'True
            Value           =   7
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Colors.userSpinner usrspnGroup 
            Height          =   300
            Index           =   2
            Left            =   1680
            TabIndex        =   21
            Top             =   1620
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   529
            Appearance3D    =   -1  'True
            Value           =   9
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin Colors.userSpinner usrspnGroup 
            Height          =   300
            Index           =   0
            Left            =   1680
            TabIndex        =   15
            Top             =   360
            Width           =   972
            _ExtentX        =   1715
            _ExtentY        =   529
            Appearance3D    =   -1  'True
            Value           =   5
            ShowZero        =   0   'False
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderColor     =   0
            BorderInterior  =   -2147483631
            Position        =   0
            Enabled         =   -1  'True
            DisabledColor   =   -2147483631
         End
         Begin VB.Shape shpPalette 
            Height          =   2292
            Left            =   0
            Top             =   0
            Width           =   3132
         End
         Begin VB.Label lblPersonalize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Workspace:"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   396
            Width           =   1392
         End
         Begin VB.Label lblPersonalize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Controls:"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   816
            Width           =   1392
         End
         Begin VB.Label lblPersonalize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Navigation:"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Top             =   1656
            Width           =   1392
         End
         Begin VB.Label lblPersonalize 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Drop Slots:"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   1236
            Width           =   1392
         End
      End
      Begin VB.Label lblPersonalize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Personalize"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   1992
      End
      Begin VB.Label lblPalette 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Palette"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   420
         TabIndex        =   10
         Top             =   360
         Width           =   1992
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5292
      Index           =   2
      Left            =   120
      ScaleHeight     =   5292
      ScaleWidth      =   8292
      TabIndex        =   22
      Top             =   540
      Visible         =   0   'False
      Width           =   8292
      Begin VB.CheckBox chkFiles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Export"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   3
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   4260
         Width           =   1512
      End
      Begin VB.CheckBox chkFiles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Import"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   2
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3840
         Width           =   1512
      End
      Begin VB.CheckBox chkFiles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Save As"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3000
         Width           =   1512
      End
      Begin VB.CheckBox chkFiles 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Load"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   0
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2580
         Width           =   1512
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Header and Footer"
         ForeColor       =   &H80000008&
         Height          =   1692
         Index           =   0
         Left            =   5340
         TabIndex        =   50
         Top             =   120
         Width           =   2772
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   8
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   420
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   10
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1140
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   9
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   780
            Width           =   492
         End
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Borders"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   0
            Width           =   696
         End
         Begin VB.Shape shpFrame 
            BorderColor     =   &H80000006&
            Height          =   1572
            Index           =   0
            Left            =   0
            Top             =   120
            Width           =   2772
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Interior"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   8
            Left            =   1080
            TabIndex        =   53
            Top             =   468
            Width           =   1620
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Highlight"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   10
            Left            =   1080
            TabIndex        =   57
            Top             =   1188
            Width           =   1620
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Exterior"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   9
            Left            =   1080
            TabIndex        =   55
            Top             =   828
            Width           =   1620
         End
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Drop Slots"
         ForeColor       =   &H80000008&
         Height          =   3132
         Index           =   1
         Left            =   5340
         TabIndex        =   58
         Top             =   1980
         Width           =   2772
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   12
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   780
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   11
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   420
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   13
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1140
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   14
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   1500
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   16
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   2220
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   15
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   1860
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   17
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   2580
            Width           =   492
         End
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Literal Colors"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   3
            Left            =   120
            TabIndex        =   59
            Top             =   0
            Width           =   1164
         End
         Begin VB.Shape shpFrame 
            BorderColor     =   &H80000006&
            Height          =   3012
            Index           =   3
            Left            =   0
            Top             =   120
            Width           =   2772
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Red"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   12
            Left            =   1080
            TabIndex        =   63
            Top             =   828
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Light Gray"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   11
            Left            =   1080
            TabIndex        =   61
            Top             =   468
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Yellow"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   13
            Left            =   1080
            TabIndex        =   65
            Top             =   1188
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Blue"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   14
            Left            =   1080
            TabIndex        =   67
            Top             =   1548
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Green"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   16
            Left            =   1080
            TabIndex        =   71
            Top             =   2268
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Purple"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   15
            Left            =   1080
            TabIndex        =   69
            Top             =   1908
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Orange"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   17
            Left            =   1080
            TabIndex        =   73
            Top             =   2628
            Width           =   1560
         End
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Workspace"
         ForeColor       =   &H80000008&
         Height          =   2352
         Index           =   4
         Left            =   2340
         TabIndex        =   40
         Top             =   2760
         Width           =   2772
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   7
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1500
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   6
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1140
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   5
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   780
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   4
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   420
            Width           =   492
         End
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Text"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   4
            Left            =   120
            TabIndex        =   41
            Top             =   0
            Width           =   396
         End
         Begin VB.Shape shpFrame 
            BorderColor     =   &H80000006&
            Height          =   2232
            Index           =   4
            Left            =   0
            Top             =   120
            Width           =   2772
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Link"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   7
            Left            =   1080
            TabIndex        =   49
            Top             =   1548
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Dim"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   6
            Left            =   1080
            TabIndex        =   47
            Top             =   1188
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Error"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   5
            Left            =   1080
            TabIndex        =   45
            Top             =   828
            Width           =   1560
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Normal"
            ForeColor       =   &H80000008&
            Height          =   252
            Index           =   4
            Left            =   1080
            TabIndex        =   43
            Top             =   468
            Width           =   1560
         End
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Workspace"
         ForeColor       =   &H80000008&
         Height          =   2412
         Index           =   2
         Left            =   2340
         TabIndex        =   30
         Top             =   120
         Width           =   2772
         Begin VB.ComboBox cboBackground 
            Height          =   312
            ItemData        =   "frmColors.frx":08CA
            Left            =   1080
            List            =   "frmColors.frx":08DD
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   420
            Visible         =   0   'False
            Width           =   1572
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   3
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1500
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   2
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1140
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   0
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   420
            Width           =   492
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   1
            Left            =   360
            ScaleHeight     =   288
            ScaleWidth      =   468
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   780
            Width           =   492
         End
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Background"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   0
            Width           =   1056
         End
         Begin VB.Shape shpFrame 
            BorderColor     =   &H80000006&
            Height          =   2292
            Index           =   1
            Left            =   0
            Top             =   120
            Width           =   2772
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Related"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   216
            Index           =   3
            Left            =   1080
            TabIndex        =   39
            Top             =   1548
            Width           =   756
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Error"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   216
            Index           =   2
            Left            =   1080
            TabIndex        =   37
            Top             =   1188
            Width           =   492
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   216
            Index           =   0
            Left            =   1080
            TabIndex        =   74
            Top             =   468
            Width           =   708
         End
         Begin VB.Label lblColor 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Highlight"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   216
            Index           =   1
            Left            =   1080
            TabIndex        =   35
            Top             =   828
            Width           =   852
         End
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2052
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1992
         Begin Colors.userCheckBox usrchkArea 
            Height          =   252
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   660
            Width           =   1752
            _ExtentX        =   3090
            _ExtentY        =   445
            Value           =   0   'False
            Caption         =   "Controls"
         End
         Begin Colors.userCheckBox usrchkArea 
            Height          =   252
            Index           =   2
            Left            =   180
            TabIndex        =   27
            Top             =   960
            Width           =   1752
            _ExtentX        =   3090
            _ExtentY        =   445
            Value           =   0   'False
            Caption         =   "Navigation"
         End
         Begin Colors.userCheckBox usrchkArea 
            Height          =   252
            Index           =   3
            Left            =   180
            TabIndex        =   28
            Top             =   1260
            Width           =   1752
            _ExtentX        =   3090
            _ExtentY        =   445
            Value           =   0   'False
            Caption         =   "Drop Slots"
         End
         Begin Colors.userCheckBox usrchkArea 
            Height          =   252
            Index           =   4
            Left            =   180
            TabIndex        =   29
            Top             =   1560
            Width           =   1752
            _ExtentX        =   3090
            _ExtentY        =   445
            Value           =   0   'False
            Caption         =   "Output"
         End
         Begin Colors.userCheckBox usrchkArea 
            Height          =   252
            Index           =   0
            Left            =   180
            TabIndex        =   25
            Top             =   360
            Width           =   1752
            _ExtentX        =   3090
            _ExtentY        =   445
            Caption         =   "Workspace"
         End
         Begin VB.Label lblFrame 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Area"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   0
            Width           =   408
         End
         Begin VB.Shape shpFrame 
            Height          =   1932
            Index           =   2
            Left            =   0
            Top             =   120
            Width           =   1992
         End
      End
   End
   Begin VB.Label lnkTab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Customize"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   2
      Left            =   2712
      TabIndex        =   2
      Tag             =   "nav"
      Top             =   84
      Width           =   1032
   End
   Begin VB.Label lnkTab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Palette"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Tag             =   "nav"
      Top             =   84
      Width           =   708
   End
   Begin VB.Label lnkTab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Saved"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   0
      Left            =   264
      TabIndex        =   0
      Tag             =   "nav"
      Top             =   84
      Width           =   612
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   216
      Index           =   0
      Left            =   7872
      TabIndex        =   3
      Tag             =   "nav"
      Top             =   84
      Width           =   432
   End
   Begin VB.Shape shpNav 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   384
      Left            =   0
      Tag             =   "nav"
      Top             =   0
      Width           =   8568
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Written by Ellis Dee
Option Explicit

Public Enum ImportExportEnum
    ieImport
    ieExport
End Enum

Public Enum RelativeColorEnum
    rceRed
    rceYellow
    rceBlue
End Enum

Private mlngTab As Long

Private menColorGroup As ColorGroupEnum
Private mlngIndex As Long
Private mblnOverride As Boolean
Private menImportExport As ImportExportEnum

Private clr As clsMaterialColor


' ************* FORM *************


Private Sub Form_Load()
    Set clr = New clsMaterialColor
    mblnOverride = True
    mlngTab = 0
    mlngIndex = 0
    menColorGroup = cgeWorkspace
    cfg.MoveForm Me
    LoadData
    RefreshColors
    mblnOverride = False
    If Not xp.DebugMode Then Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    Set clr = Nothing
    cfg.SavePosition Me
    If App.Title = "Character Builder Lite" And GetForm(frm, "frmMain") Then frm.UpdateWindowMenu
    If Not xp.DebugMode Then Call WheelUnHook(Me.hWnd)
    CloseApp
End Sub

Public Property Get ColorGroup() As ColorGroupEnum
    ColorGroup = menColorGroup
End Property

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    Dim i As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    For i = 0 To 3
        If IsOver(Me.usrspnGroup(i).hWnd, Xpos, Ypos) Then
            Me.usrspnGroup(i).WheelScroll lngValue
        End If
    Next
End Sub


' ************* NAVIGATION *************


Private Sub lnkTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    xp.SetMouseCursor mcHand
    mlngTab = Index
    For i = 0 To 2
        Me.lnkTab(i).FontUnderline = (mlngTab = i)
        Me.picTab(i).Visible = (mlngTab = i)
    Next
    RefreshColors
End Sub

Private Sub lnkNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    Select Case Me.lnkNav(Index).Caption
        Case "Help": ShowHelp "Colors"
    End Select
End Sub


' ************* GENERAL *************


Private Sub LoadData()
    Dim i As Long
    Dim j As Long
    
    InitListboxes
    Me.lstMaterial.Clear
    For i = 0 To 18
        Me.lstMaterial.AddItem clr.GetColorName(i)
    Next
End Sub

Private Sub InitListboxes()
    InitListBox Me.lstScreen, "screen", cfg.ScreenColors
    InitListBox Me.lstOutput, "output", cfg.OutputColors
End Sub

Private Sub InitListBox(plst As ListBox, pstrExt As String, pstrSelected As String)
    Dim strFile As String
    Dim i As Long
    
    mblnOverride = True
    plst.Clear
    strFile = Dir(SettingsPath() & "*." & pstrExt)
    Do While Len(strFile)
        plst.AddItem GetNameFromFilespec(strFile)
        strFile = Dir
    Loop
    For i = 0 To plst.ListCount - 1
        If plst.List(i) = pstrSelected Then
            plst.ListIndex = i
            Exit For
        End If
    Next
    mblnOverride = False
End Sub

Public Sub RefreshColors()
    Dim lngBackColor As Long
    Dim lngTextColor As Long
    Dim lngLinkColor As Long
    Dim blnEnabled As Boolean
    Dim lngColor As Long
    Dim i As Long
    
    mblnOverride = True
    xp.LockWindow Me.hWnd
    cfg.ApplyColors Me.shpNav, cgeNavigation
    For i = 0 To 2
        cfg.ApplyColors Me.lnkTab(i), cgeNavigation
    Next
    cfg.ApplyColors Me.lnkNav(0), cgeNavigation
    If mlngTab <> 2 Then
        Me.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        Me.picTab(mlngTab).BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
    End If
    Select Case mlngTab
        Case 0 ' Saved
            ' Listboxes
            cfg.ApplyColors Me.lblScreenColors, cgeWorkspace
            cfg.ApplyColors Me.lstScreen, cgeControls
            cfg.ApplyColors Me.lblOutputColors, cgeWorkspace
            cfg.ApplyColors Me.lstOutput, cgeControls
        Case 1 ' Palette
            cfg.ApplyColors Me.lblPalette, cgeWorkspace
            cfg.ApplyColors Me.lstMaterial, cgeControls
            cfg.ApplyColors Me.shpPalette, cgeWorkspace
            cfg.ApplyColors Me.fraPalette, cgeWorkspace
            blnEnabled = (Me.lstMaterial.ListIndex <> -1)
            Me.fraPalette.Enabled = blnEnabled
            Me.picPreview.Visible = blnEnabled
            If blnEnabled Then lngColor = cfg.GetColor(cgeWorkspace, cveText) Else lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
            For i = 0 To 4
                Me.lblPersonalize(i).ForeColor = lngColor
                Me.lblPersonalize(i).BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
            Next
            If blnEnabled Then lngColor = cfg.GetColor(cgeControls, cveText) Else lngColor = cfg.GetColor(cgeControls, cveTextDim)
            For i = 0 To 3
                Me.usrspnGroup(i).ForeColor = lngColor
                Me.usrspnGroup(i).BackColor = cfg.GetColor(cgeControls, cveBackground)
            Next
        Case 2 ' Customize
            Select Case menColorGroup
                Case cgeWorkspace
                    Me.lblColor(8).Caption = "Checkbox Int."
                    Me.lblColor(9).Caption = "Checkbox Ext."
                    Me.lblColor(10).Caption = "Frame Borders"
                Case cgeControls
                    Me.lblColor(8).Caption = "Grid Interior"
                    Me.lblColor(9).Caption = "Grid Exterior"
                    Me.lblColor(10).Caption = "Active Cell"
                Case cgeDropSlots
                    Me.lblColor(8).Caption = "Grid Interior"
                    Me.lblColor(9).Caption = "Slot Exterior"
                    Me.lblColor(10).Caption = "Active Slot/Cell"
                Case cgeNavigation
                    Me.lblColor(8).Caption = "(Not used)"
                    Me.lblColor(9).Caption = "(Not used)"
                    Me.lblColor(10).Caption = "Nav. Border"
                Case cgeOutput
                    Me.lblColor(8).Caption = "(Not used)"
                    Me.lblColor(9).Caption = "(Not used)"
                    Me.lblColor(10).Caption = "(Not used)"
            End Select
            lngTextColor = cfg.GetColor(menColorGroup, cveText)
            lngLinkColor = cfg.GetColor(menColorGroup, cveTextLink)
            lngBackColor = cfg.GetColor(menColorGroup, mlngIndex)
            Me.BackColor = lngBackColor
            Me.picTab(2).BackColor = lngBackColor
            For i = 0 To 4
                ' Checkboxes
                With Me.usrchkArea(i)
                    .Value = (menColorGroup = i)
                    .CustomColors menColorGroup, mlngIndex
                End With
                ' Frames
                Me.fraGroup(i).BackColor = lngBackColor
                Me.shpFrame(i).BorderColor = cfg.GetColor(menColorGroup, cveBorderExterior)
                Me.lblFrame(i).ForeColor = lngTextColor
                Me.lblFrame(i).BackColor = lngBackColor
            Next
            For i = 0 To 3
                cfg.ApplyColors Me.chkFiles(i), cgeControls
            Next
            ' Combobox
            cfg.ApplyColors Me.cboBackground, cgeControls
            ComboSetValue Me.cboBackground, cfg.BackgroundStyle
            Me.cboBackground.Visible = (menColorGroup = cgeOutput)
            ' Colors
            For i = 0 To 17
                Me.picColor(i).BackColor = cfg.GetColor(menColorGroup, i)
                Me.lblColor(i).BackColor = lngBackColor
                Select Case i
                    Case 0 To 3: Me.lblColor(i).ForeColor = lngLinkColor
                    Case 4 To 7, 11 To 17: Me.lblColor(i).ForeColor = cfg.GetColor(menColorGroup, i)
                    Case Else: Me.lblColor(i).ForeColor = lngTextColor
                End Select
            Next
    End Select
    xp.UnlockWindow
    mblnOverride = False
End Sub

Private Sub ComboColors(pcbo As ComboBox, penGroup As ColorGroupEnum, penFore As ColorValueEnum, penBack As ColorValueEnum)
    pcbo.ForeColor = cfg.GetColor(penGroup, penFore)
    pcbo.BackColor = cfg.GetColor(penGroup, penBack)
End Sub

Private Sub ListboxColors(plst As ListBox, plngTextColor As Long, plngBackColor As Long)
    plst.ForeColor = plngTextColor
    plst.BackColor = plngBackColor
End Sub


' ************* SAVED *************


Private Sub lstScreen_Click()
    If mblnOverride Then Exit Sub
    LoadFile Me.lstScreen.Text, "screen"
    Me.lstMaterial.ListIndex = -1
End Sub

Private Sub lstOutput_Click()
    If mblnOverride Then Exit Sub
    LoadFile Me.lstOutput.Text, "output"
End Sub

Private Sub LoadFile(pstrName As String, pstrExt As String)
    Dim strFile As String
    
    strFile = SettingsPath() & pstrName & "." & pstrExt
    If xp.File.Exists(strFile) Then
        cfg.LoadColorFile strFile
        If Not mblnOverride Then ColorChange
    End If
End Sub


' ************* PALETTE *************


Private Sub lstMaterial_Click()
    Dim enPalette As MaterialColorEnum
    Dim blnEnabled As Boolean
    Dim lngColor As Long
    Dim i As Long
    
    blnEnabled = (Me.lstMaterial.ListIndex <> -1)
    Me.fraPalette.Enabled = blnEnabled
    If blnEnabled Then lngColor = cfg.GetColor(cgeWorkspace, cveText) Else lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
    For i = 0 To 4
        Me.lblPersonalize(i).ForeColor = lngColor
    Next
    If blnEnabled Then lngColor = cfg.GetColor(cgeControls, cveText) Else lngColor = cfg.GetColor(cgeControls, cveTextDim)
    For i = 0 To 3
        Me.usrspnGroup(i).ForeColor = lngColor
    Next
    If blnEnabled Then
        enPalette = Me.lstMaterial.ListIndex
        ApplyMaterialColors
        PreviewColors
    End If
End Sub

Private Sub usrspnGroup_Change(Index As Integer)
    ApplyMaterialColors
End Sub

Private Sub PreviewColors()
    Dim enPalette As MaterialColorEnum
    Dim strCaption As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngForeColor As Long
    Dim lngBackColor As Long
    Dim i As Long

    enPalette = Me.lstMaterial.ListIndex
    Me.picPreview.Cls
    lngHeight = Me.picPreview.ScaleHeight \ 10
    For i = 0 To 9
        lngForeColor = clr.GetTextColor(enPalette, i)
        lngBackColor = clr.GetColor(enPalette, i)
        lngTop = lngHeight * (9 - i)
        strCaption = i + 1 & ": #" & clr.GetColorHex(enPalette, i)
        Me.picPreview.Line (0, lngTop)-(Me.picPreview.ScaleWidth, lngTop + lngHeight), lngBackColor, BF
        With Me.picPreview
            .ForeColor = lngForeColor
            lngWidth = .TextWidth(strCaption)
            .CurrentX = (.ScaleWidth - .TextWidth(strCaption)) \ 2
            .CurrentY = lngTop + (lngHeight - .TextHeight(strCaption)) \ 2
        End With
        Me.picPreview.Print strCaption
    Next
    Me.picPreview.Visible = True
End Sub

Private Sub ApplyMaterialColors()
    Dim enGroup As ColorGroupEnum
    Dim enPalette As MaterialColorEnum
    Dim lngValue As Long
    Dim lngColor As Long
    
    If frmColors.lstMaterial.ListIndex = -1 Then Exit Sub
    enPalette = frmColors.lstMaterial.ListIndex
    ' Workspace, Controls, Navigation, Drop Slots
    For enGroup = 0 To 3
        lngValue = frmColors.usrspnGroup(enGroup).Value - 1
        lngColor = clr.GetColor(enPalette, lngValue)
        ' Background
        cfg.SetColor enGroup, cveBackground, clr.GetColor(enPalette, lngValue)
        cfg.SetColor enGroup, cveBackground, clr.GetColor(enPalette, lngValue)
        cfg.SetColor enGroup, cveBackHighlight, RelativeColor(enPalette, lngValue, rceYellow)
        cfg.SetColor enGroup, cveBackError, RelativeColor(enPalette, lngValue, rceRed)
        cfg.SetColor enGroup, cveBackRelated, RelativeColor(enPalette, lngValue, rceBlue)
        ' Text
        ApplyMaterialTextColor enPalette, lngValue, enGroup
        ' Literal
        cfg.SetColor enGroup, cveLightGray, clr.GetColor(mceGrey, lngValue)
        cfg.SetColor enGroup, cveRed, clr.GetColor(mceRed, lngValue)
        cfg.SetColor enGroup, cveGreen, clr.GetColor(mceGreen, lngValue)
        cfg.SetColor enGroup, cveBlue, clr.GetColor(mceBlue, lngValue)
        cfg.SetColor enGroup, cveYellow, clr.GetColor(mceYellow, lngValue)
        cfg.SetColor enGroup, cveOrange, clr.GetColor(mceOrange, lngValue)
        cfg.SetColor enGroup, cvePurple, clr.GetColor(mcePurple, lngValue)
    Next
    ApplyMaterialBorderColors
    cfg.RefreshAllColors
    cfg.ScreenColors = vbNullString
    InitListboxes
    If Not mblnOverride Then ColorChange
End Sub

Public Function RelativeColor(penPalette As MaterialColorEnum, plngValue As Long, penRelative As RelativeColorEnum) As Long
    Dim enRelative As MaterialColorEnum
    Dim enText As MaterialColorTextEnum
    Dim lngStep As Long
    Dim lngLast As Long
    Dim i As Long
    
    Select Case penRelative
        Case rceYellow
            Select Case penPalette
                Case mceBrown: enRelative = mceLime
                Case mceYellow: enRelative = mceLightGreen
                Case Else: enRelative = mceYellow
            End Select
        Case rceRed
            Select Case penPalette
                Case mceAmber: enRelative = mcePink
                Case mceDeepOrange: enRelative = mceDeepPurple
                Case mcePink: enRelative = mceDeepOrange
                Case mcePurple: enRelative = mceDeepOrange
                Case mceRed: enRelative = mceDeepOrange
                Case mceYellow: enRelative = mceLightGreen
                Case Else: enRelative = mceRed
            End Select
        Case rceBlue
            Select Case penPalette
                Case mceAmber: enRelative = mceCyan
                Case mceBlue: enRelative = mceGreen
                Case mceBlueGrey: enRelative = mceBlue
                Case mceCyan: enRelative = mceDeepPurple
                Case mceIndigo: enRelative = mceLightBlue
                Case mceLightBlue: enRelative = mcePurple
                Case mceLime: enRelative = mceCyan
                Case Else: enRelative = mceLightBlue
            End Select
    End Select
    enText = clr.GetTextColor(penPalette, plngValue)
    Select Case enText
        Case mcteBlack
            lngStep = 1
            lngLast = 9
        Case mcteWhite
            lngStep = -1
            lngLast = 0
    End Select
    For i = plngValue To lngLast Step lngStep
        If clr.GetTextColor(enRelative, i) = enText Then
            RelativeColor = clr.GetColor(enRelative, i)
            Exit Function
        End If
    Next
    RelativeColor = clr.GetColor(enRelative, i - lngStep)
End Function

Public Sub ApplyMaterialTextColor(penPalette As MaterialColorEnum, plngValue As Long, penGroup As ColorGroupEnum)
    Dim enRelative As MaterialColorEnum
    Dim lngDefault As Long
    Dim lngDim As Long
    Dim lngText As Long
    Dim lngLink As Long
    Dim enError As MaterialColorEnum
    Dim enLink As MaterialColorEnum
    Dim i As Long
    
    Select Case penPalette
        Case mceAmber: enError = mceRed
        Case mceDeepOrange: enError = mcePink
        Case mcePink: enError = mceLightGreen
        Case mcePurple: enError = mcePink
        Case mceRed: enError = mcePink
        Case mceYellow: enError = mceRed
        Case Else: enError = mceRed
    End Select
    Select Case penPalette
        Case mceAmber: enLink = mceBlue
        Case mceBlue: enLink = mceIndigo
        Case mceBlueGrey: enLink = mceBlue
        Case mceCyan: enLink = mceBlue
        Case mceIndigo: enLink = mceYellow
        Case mceLightBlue: enLink = mcePurple
        Case mceLime: enLink = mceBlue
        Case Else: enLink = mceBlue
    End Select
    Select Case clr.GetTextColor(penPalette, plngValue)
        Case mcteBlack
            lngText = vbBlack
            lngDefault = 1
            lngDim = 3
        Case mcteWhite
            lngText = vbWhite
            lngDefault = 8
            lngDim = 6
    End Select
    cfg.SetColor penGroup, cveText, lngText
    cfg.SetColor penGroup, cveTextError, clr.GetColor(enError, lngDefault)
    cfg.SetColor penGroup, cveTextDim, clr.GetColor(mceGrey, lngDim)
    If penGroup = cgeNavigation Then
        If plngValue > 5 Then lngLink = 1 Else lngLink = 8
        cfg.SetColor penGroup, cveTextLink, clr.GetColor(penPalette, lngLink)
    Else
        cfg.SetColor penGroup, cveTextLink, clr.GetColor(enLink, lngDefault)
    End If
End Sub

Private Sub ApplyMaterialBorderColors()
    Dim enPalette As MaterialColorEnum
    Dim lngValue As Long

    enPalette = frmColors.lstMaterial.ListIndex
    ' Workspace
    lngValue = frmColors.usrspnGroup(cgeWorkspace).Value - 1
    cfg.SetColor cgeWorkspace, cveBorderInterior, vbGrayText
    cfg.SetColor cgeWorkspace, cveBorderExterior, vbBlack
    cfg.SetColor cgeWorkspace, cveBorderHighlight, clr.StepColor(enPalette, lngValue, 4)
    ' Controls
    lngValue = frmColors.usrspnGroup(cgeControls).Value - 1
    cfg.SetColor cgeControls, cveBorderInterior, clr.StepColor(enPalette, lngValue, 2)
    cfg.SetColor cgeControls, cveBorderExterior, clr.StepColor(enPalette, lngValue, 4)
    Select Case clr.GetTextColor(enPalette, lngValue)
        Case mcteBlack: cfg.SetColor cgeControls, cveBorderHighlight, clr.GetColor(enPalette, 0)
        Case mcteWhite: cfg.SetColor cgeControls, cveBorderHighlight, clr.GetColor(enPalette, 9)
    End Select
    ' Drop slots
    lngValue = frmColors.usrspnGroup(cgeDropSlots).Value - 1
    cfg.SetColor cgeDropSlots, cveBorderInterior, clr.StepColor(enPalette, lngValue, 2)
    cfg.SetColor cgeDropSlots, cveBorderExterior, clr.StepColor(enPalette, lngValue, 4)
    Select Case clr.GetTextColor(enPalette, lngValue)
        Case mcteBlack: cfg.SetColor cgeDropSlots, cveBorderHighlight, clr.GetColor(enPalette, 0)
        Case mcteWhite: cfg.SetColor cgeDropSlots, cveBorderHighlight, clr.GetColor(enPalette, 9)
    End Select
    ' Navigation
    lngValue = frmColors.usrspnGroup(cgeNavigation).Value - 1
    cfg.SetColor cgeNavigation, cveBorderInterior, vbBlack
    cfg.SetColor cgeNavigation, cveBorderExterior, vbBlack
    cfg.SetColor cgeNavigation, cveBorderHighlight, vbBlack
End Sub

Private Function GetHighLow(plngColor As Long, plngHigh As Long, plngLow As Long) As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    xp.ColorToRGB plngColor, lngRed, lngGreen, lngBlue
    ' High color
    plngHigh = lngRed
    If plngHigh < lngGreen Then plngHigh = lngGreen
    If plngHigh < lngBlue Then plngHigh = lngBlue
    ' Low color
    plngLow = lngRed
    If plngLow > lngGreen Then plngLow = lngGreen
    If plngLow > lngBlue Then plngLow = lngBlue
End Function


' ************* CUSTOM *************


Private Sub usrchkArea_UserChange(Index As Integer)
    mlngIndex = 0
    menColorGroup = Index
    RefreshColors
End Sub

Private Sub chkFiles_Click(Index As Integer)
    If UncheckButton(Me.chkFiles(Index), mblnOverride) Then Exit Sub
    Select Case Me.chkFiles(Index).Caption
        Case "Load": LoadColors
        Case "Save As": SaveColors
        Case "Import": ImportColors
        Case "Export": ExportColors
    End Select
End Sub

Private Sub lblColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index < 4 Then xp.SetMouseCursor mcHand
End Sub

Private Sub lblColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index < 4 Then
        xp.SetMouseCursor mcHand
        mlngIndex = Index
        RefreshColors
    End If
End Sub

Private Sub lblColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index < 4 Then xp.SetMouseCursor mcHand
End Sub

Private Sub cboBackground_GotFocus()
    If mlngIndex <> 0 Then
        mlngIndex = 0
        RefreshColors
    End If
End Sub

Private Sub cboBackground_Click()
    Dim enBack As BackColorEnum
    
    If mblnOverride Then Exit Sub
    With Me.cboBackground
        If .ListIndex = -1 Then enBack = bceCustom Else enBack = .ItemData(.ListIndex)
    End With
    mlngIndex = 0
    cfg.BackgroundStyle = enBack
    ApplyChanges
    cfg.OutputColors = vbNullString
 End Sub

Private Sub lblFrame_Click(Index As Integer)
    Dim lngValue As Long
    Dim lngHigh As Long
    Dim lngLow As Long
    Dim i As Long
    
    If Me.lblFrame(Index).Caption = "Area" Then Exit Sub
    If menColorGroup = cgeOutput Then
        If Not Ask("It is not recommended to change output colors in this manner. Do it anyway?") Then Exit Sub
    End If
    Select Case Me.lblFrame(Index).Caption
        Case "Text"
            If Not Ask("Toggle text colors?") Then Exit Sub
            GetHighLow cfg.GetColor(menColorGroup, cveText), lngHigh, lngLow
            If lngHigh < 128 Then
                cfg.SetColor menColorGroup, cveText, vbWhite
                cfg.SetColor menColorGroup, cveTextError, clr.GetColor(mceRed, 7)
                cfg.SetColor menColorGroup, cveTextDim, clr.GetColor(mceGrey, 5)
                cfg.SetColor menColorGroup, cveTextLink, clr.GetColor(mceBlue, 7)
            Else
                cfg.SetColor menColorGroup, cveText, vbBlack
                cfg.SetColor menColorGroup, cveTextError, RGB(255, 0, 0)
                cfg.SetColor menColorGroup, cveTextDim, clr.GetColor(mceGrey, 4)
                cfg.SetColor menColorGroup, cveTextLink, RGB(0, 0, 255)
            End If
        Case "Background"
            If Not Ask("Calculate accent colors based on standard background?") Then Exit Sub
            GetHighLow cfg.GetColor(menColorGroup, cveBackground), lngHigh, lngLow
            cfg.SetColor menColorGroup, cveBackHighlight, RGB(lngHigh, lngHigh, lngLow)
            cfg.SetColor menColorGroup, cveBackError, RGB(lngHigh, lngLow, lngLow)
            cfg.SetColor menColorGroup, cveBackRelated, RGB(lngLow, lngLow, lngHigh)
        Case "Borders"
            If menColorGroup = cgeOutput Then Exit Sub
            If Not Ask("Set border colors for all groups?") Then Exit Sub
            cfg.SetColor cgeWorkspace, cveBorderInterior, cfg.GetColor(cgeWorkspace, cveTextDim)
            cfg.SetColor cgeWorkspace, cveBorderExterior, cfg.GetColor(cgeWorkspace, cveText)
            cfg.SetColor cgeWorkspace, cveBorderHighlight, cfg.GetColor(cgeWorkspace, cveText)
            cfg.SetColor cgeControls, cveBorderInterior, cfg.GetColor(cgeWorkspace, cveBackground)
            cfg.SetColor cgeControls, cveBorderHighlight, clr.GetColor(mceGrey, 9)
        Case "Literal Colors"
            lngValue = Val(InputBox("Set literal colors to Material Design values" & vbNewLine & vbNewLine & "Select 1-10, or 0 to Cancel", "Apply Material Design Colors", 0))
            If lngValue < 1 Or lngValue > 10 Then Exit Sub
            lngValue = lngValue - 1
            cfg.SetColor menColorGroup, cveLightGray, clr.GetColor(mceGrey, lngValue)
            cfg.SetColor menColorGroup, cveRed, clr.GetColor(mceRed, lngValue)
            cfg.SetColor menColorGroup, cveYellow, clr.GetColor(mceYellow, lngValue)
            cfg.SetColor menColorGroup, cveBlue, clr.GetColor(mceBlue, lngValue)
            cfg.SetColor menColorGroup, cveGreen, clr.GetColor(mceGreen, lngValue)
            cfg.SetColor menColorGroup, cvePurple, clr.GetColor(mcePurple, lngValue)
            cfg.SetColor menColorGroup, cveOrange, clr.GetColor(mceOrange, lngValue)
    End Select
    ApplyChanges
    cfg.OutputColors = vbNullString
End Sub

Private Sub picColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub picColor_Click(Index As Integer)
    If mlngIndex <> Index And Index < 4 Then
        mlngIndex = Index
        RefreshColors
    End If
    glngActiveColor = cfg.GetColor(menColorGroup, Index)
    frmColor.Show vbModal, Me
    If cfg.GetColor(menColorGroup, Index) = glngActiveColor Then Exit Sub
    If menColorGroup = cgeOutput And Index = cveBackground Then
        mblnOverride = True
        ComboSetValue Me.cboBackground, cfg.BackgroundStyle
        mblnOverride = False
    End If
    cfg.SetColor menColorGroup, Index, glngActiveColor
    ApplyChanges
    Me.lstMaterial.ListIndex = -1
    If menColorGroup = cgeOutput Then cfg.OutputColors = vbNullString Else cfg.ScreenColors = vbNullString
    InitListboxes
End Sub

Private Sub ApplyChanges()
    Dim frm As Form
    
    RefreshColors
    For Each frm In Forms
        Select Case frm.Name
            Case "frmColors"
            Case "frmMain": If App.Title = "Character Builder Lite" And menColorGroup = cgeOutput Then cfg.RefreshColors frm
            Case Else: cfg.RefreshColors frm
        End Select
    Next
    Set frm = Nothing
    If Not mblnOverride Then
        ColorChange
    End If
End Sub

Private Sub LoadColors()
    Dim strFile As String
    Dim strText As String
    Dim strFilter As String
    
    If menColorGroup = cgeOutput Then
        strFile = xp.ShowOpenDialog(SettingsPath(), "Output Colors (*.output)|*.output|Screen Colors (*.screen)|*.screen", "*.output")
    Else
        strFile = xp.ShowOpenDialog(SettingsPath(), "Screen Colors (*.screen)|*.screen|Output Colors (*.output)|*.output", "*.screen")
    End If
    If Len(strFile) Then
        cfg.LoadColorFile strFile
        ColorChange
    End If
    InitListboxes
End Sub

Private Sub SaveColors()
    Dim strFile As String
    Dim strText As String
    
    If menColorGroup = cgeOutput Then
        strFile = xp.ShowSaveAsDialog(SettingsPath(), vbNullString, "Output Colors (*.output)|*.output", "*.output")
    Else
        strFile = xp.ShowSaveAsDialog(SettingsPath(), vbNullString, "Screen Colors (*.screen)|*.screen", "*.screen")
    End If
    If Len(strFile) = 0 Then
        InitListboxes
        Exit Sub
    End If
    If menColorGroup = cgeOutput Then
        strText = cfg.MakeColorSettings(cfeOutput)
    Else
        strText = cfg.MakeColorSettings(cfeScreen)
    End If
    If xp.File.Exists(strFile) Then
        If MsgBox(GetFileFromFilespec(strFile) & " exists. Overwrite?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
        xp.File.Delete strFile
    End If
    xp.File.SaveStringAs strFile, strText
    Select Case LCase$(GetExtFromFilespec(strFile))
        Case cfg.GetExtension(eeScreenColor): cfg.ScreenColors = GetNameFromFilespec(strFile)
        Case cfg.GetExtension(eeOutputColor): cfg.OutputColors = GetNameFromFilespec(strFile)
    End Select
    InitListboxes
End Sub

Public Property Get ImportExport() As ImportExportEnum
    ImportExport = menImportExport
End Property

Private Sub ImportColors()
    menImportExport = ieImport
    frmColorFile.Show vbModeless, Me
End Sub

Private Sub ExportColors()
    menImportExport = ieExport
    frmColorFile.Show vbModeless, Me
End Sub

