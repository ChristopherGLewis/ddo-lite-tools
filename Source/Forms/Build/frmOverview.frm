VERSION 5.00
Begin VB.Form frmOverview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Overview"
   ClientHeight    =   7920
   ClientLeft      =   30
   ClientTop       =   405
   ClientWidth     =   12990
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOverview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   12990
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6252
      Left            =   8100
      ScaleHeight     =   6255
      ScaleWidth      =   4635
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   540
      Width           =   4635
      Begin VB.PictureBox picBuildClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   672
         Index           =   2
         Left            =   2580
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   672
      End
      Begin VB.PictureBox picBuildClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   672
         Index           =   1
         Left            =   1620
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   672
      End
      Begin VB.PictureBox picBuildClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   672
         Index           =   0
         Left            =   660
         ScaleHeight     =   675
         ScaleWidth      =   675
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   672
      End
   End
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   390
      Left            =   0
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   7560
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   688
      Spacing         =   264
      UseTabs         =   0   'False
      BorderColor     =   -2147483640
      RightLinks      =   "Stats >"
   End
   Begin VB.TextBox txtRace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   204
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1896
   End
   Begin VB.ComboBox cboBuildClass 
      Height          =   330
      Index           =   0
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2460
      Width           =   2532
   End
   Begin VB.ComboBox cboBuildClass 
      Height          =   330
      Index           =   2
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3180
      Width           =   2532
   End
   Begin VB.ComboBox cboBuildClass 
      Height          =   330
      Index           =   1
      Left            =   1800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2820
      Width           =   2532
   End
   Begin VB.TextBox txtName 
      Height          =   324
      Left            =   1800
      TabIndex        =   2
      Top             =   660
      Width           =   3552
   End
   Begin CharacterBuilderLite.userSpinner usrSpinner 
      Height          =   300
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.ComboBox cboAlignment 
      Height          =   330
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1380
      Width           =   2412
   End
   Begin CharacterBuilderLite.userRaceCombo usrRace 
      Height          =   195
      Left            =   1800
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   344
   End
   Begin VB.ComboBox cboRace 
      Height          =   330
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2412
   End
   Begin CharacterBuilderLite.userHeader usrHeader 
      Height          =   390
      Left            =   -240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   688
      Spacing         =   264
      BorderColor     =   -2147483640
      LeftLinks       =   "Overview;Notes"
      RightLinks      =   "Help"
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   7575
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   7575
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":000C
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   23
         Left            =   6870
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":08D6
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   22
         Left            =   6075
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":11A0
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   21
         Left            =   5205
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":1A6A
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   20
         Left            =   4305
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":2334
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   19
         Left            =   3375
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":2BFE
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   18
         Left            =   2325
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":34C8
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   17
         Left            =   1275
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":3D92
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   16
         Left            =   270
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2520
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":465C
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   13
         Left            =   4305
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":4F26
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   15
         Left            =   6075
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":5FF0
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   14
         Left            =   5205
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":68BA
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   12
         Left            =   3375
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":7184
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   11
         Left            =   2325
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":7A4E
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   10
         Left            =   1275
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":8318
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   9
         Left            =   270
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1140
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":8BE2
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   8
         Left            =   6870
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":94AC
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   7
         Left            =   6075
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":9D76
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   6
         Left            =   5205
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":A640
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   5
         Left            =   4305
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":AF0A
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   4
         Left            =   3375
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":B7D4
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   3
         Left            =   2325
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":C09E
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   2
         Left            =   1275
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.PictureBox picClass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragIcon        =   "frmOverview.frx":C968
         DragMode        =   1  'Automatic
         ForeColor       =   &H80000008&
         Height          =   624
         Index           =   1
         Left            =   270
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   624
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Wild Mage"
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   23
         Left            =   6840
         TabIndex        =   71
         Top             =   3180
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Dragon Lord"
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   22
         Left            =   6000
         TabIndex        =   68
         Top             =   3180
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Acolyte"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   21
         Left            =   5130
         TabIndex        =   66
         Top             =   3180
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Dark Hunter"
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   20
         Left            =   4230
         TabIndex        =   64
         Top             =   3180
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Sacred Fist"
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   19
         Left            =   3300
         TabIndex        =   62
         Top             =   3180
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Blightcaster"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   18
         Left            =   2070
         TabIndex        =   60
         Top             =   3180
         Width           =   1140
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Dark Apostate"
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   17
         Left            =   1155
         TabIndex        =   58
         Top             =   3180
         Width           =   870
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Stormsinger"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   16
         Left            =   15
         TabIndex        =   55
         Top             =   3180
         Width           =   1140
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Alchemist"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   15
         Left            =   5940
         TabIndex        =   54
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Warlock"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   14
         Left            =   5160
         TabIndex        =   44
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Druid"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   13
         Left            =   4392
         TabIndex        =   42
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Artificer"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   12
         Left            =   3345
         TabIndex        =   41
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Fav Soul"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   11
         Left            =   2250
         TabIndex        =   39
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Monk"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   10
         Left            =   1332
         TabIndex        =   37
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Wizard"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   267
         TabIndex        =   35
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Sorcerer"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   6795
         TabIndex        =   33
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Rogue"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   7
         Left            =   6105
         TabIndex        =   31
         Top             =   720
         Width           =   570
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Ranger"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   5205
         TabIndex        =   29
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Paladin"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   5
         Left            =   4302
         TabIndex        =   27
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Fighter"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   4
         Left            =   3375
         TabIndex        =   25
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cleric"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   2400
         TabIndex        =   23
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Bard"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   1385
         TabIndex        =   21
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Barbarian"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   19
         Top             =   720
         Width           =   840
      End
   End
   Begin VB.TextBox txtNotes 
      Height          =   2640
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4140
      Width           =   4932
   End
   Begin VB.Label lnkSubRace 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Sub-Race"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4380
      TabIndex        =   52
      Top             =   1065
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpRight 
      Height          =   6555
      Left            =   8040
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblNotes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Notes:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   420
      TabIndex        =   15
      Top             =   3885
      Width           =   615
   End
   Begin VB.Label lblMaxLevels 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Max Levels"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      TabIndex        =   7
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Label lblBuildClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Splash 2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   13
      Top             =   3225
      Width           =   1335
   End
   Begin VB.Label lblBuildClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Splash 1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   11
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Label lblBuildClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Main Class"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   9
      Top             =   2505
      Width           =   1335
   End
   Begin VB.Label lblAlignment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alignment"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      TabIndex        =   5
      Top             =   1425
      Width           =   1335
   End
   Begin VB.Label lblRace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Race"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      TabIndex        =   3
      Top             =   1065
      Width           =   1335
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   420
      TabIndex        =   1
      Top             =   690
      Width           =   1335
   End
   Begin VB.Shape shpLeft 
      Height          =   7035
      Left            =   120
      Top             =   480
      Width           =   7815
   End
End
Attribute VB_Name = "frmOverview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Enum ClassStateEnum
    cseNormal
    cseDisabled
    cseEmpty
End Enum

Private Enum GridStyleEnum
    gseWorkspace
    gseControl
    gseDropSlot
End Enum

Private Type ColumnType
    Header As String
    Left As Long
    Width As Long
    Right As Long
    Align As AlignmentConstants
End Type

Private mblnOverride As Boolean
Private mblnLoading As Boolean

Private mlngIncrement As Long
Private mlngActiveLevel As Long

Private Col(5) As ColumnType
Private grid(20, 5) As String
Private mlngBuildClass() As Long

Private mlngTop As Long
Private mlngWidth As Long
Private mlngHeight As Long
Private mlngOffsetX As Long
Private mlngOffsetY As Long

Private mblnNotes As Boolean

Private mlngNotesLabelOffsetY As Long
Private menDragState As DragEnum

Private mlngRow As Long
Private mlngCol As Long
Private mblnDrop As Boolean

Private mblnInclude32 As Boolean
Private mblnIconicClass As Boolean ' True if choosing an iconic race should auto-select its associated class

Private mblnDragging As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnNotes = False
    mblnDrop = False
    mlngNotesLabelOffsetY = Me.txtNotes.Top - Me.lblNotes.Top
    mblnLoading = True
    mblnInclude32 = (build.IncludePoints(beChampion) <> 0)
    mblnIconicClass = (build.BuildClass(0) = ceAny)
    InitComboboxes
    InitGrid
    mblnOverride = True
    cfg.Configure Me
    mblnOverride = False
    LoadData
    ShowBuild
    ShowControls
    If cfg.UseIcons And cfg.IconOverview Then Me.Show
    UpdateChoices
    mblnLoading = False
    If Not xp.DebugMode Then
        Call WheelHook(Me.hwnd)
        HookRaceCombo Me.cboRace.hwnd
    End If
End Sub

Private Sub Form_Activate()
    ActivateForm oeOverview
End Sub

Private Sub Form_Click()
    RaceHide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then
        Call WheelUnHook(Me.hwnd)
        UnhookRaceCombo Me.cboRace.hwnd
    End If
    UnloadForm Me, mblnOverride
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    If IsOver(Me.usrSpinner.hwnd, Xpos, Ypos) Then Me.usrSpinner.WheelScroll lngValue
End Sub

Public Sub RefreshColors()
    If cfg.UseIcons And cfg.IconOverview And Not mblnOverride Then DrawIcons
    DrawGrid
End Sub

Public Sub Cascade()
    mblnLoading = True
    LoadData
    ShowBuild
    ShowControls
    UpdateChoices
    InitGrid
    mblnIconicClass = (build.BuildClass(0) = ceAny)
    mblnLoading = False
End Sub

Private Function CheckGuide() As Boolean
    If build.Guide.Enhancements > 0 Then CheckGuide = Not AskAlways("This will erase the Leveling Guide. Are you sure?")
End Function


' ************* DISPLAY *************


Public Sub RefreshIcons()
    Dim i As Long
    
    xp.LockWindow Me.hwnd
    Me.picIcons.Visible = cfg.UseIcons And cfg.IconOverview
    If Not (cfg.UseIcons And cfg.IconOverview) Then
        For i = 0 To 2
            Me.picBuildClass(i).Visible = False
        Next
    End If
    InitGrid
    ShowControls
    Cascade
    xp.UnlockWindow
End Sub

Private Sub ShowControls()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngBottom As Long
    Dim blnIcons As Boolean
    Dim lngSpacing As Long
    Dim ctl As Control
    Dim i As Long
    
    xp.LockWindow Me.hwnd
    blnIcons = cfg.UseIcons And cfg.IconOverview
    ' Resize edit area of race combo
    ComboGetEditRect Me.cboRace, lngLeft, lngTop, lngWidth, lngHeight
    Me.txtRace.Move lngLeft, lngTop, lngWidth, lngHeight
    ' Resize txtNotes and lblNotes
    If mblnNotes Then
        lngLeft = Me.shpLeft.Left
        lngTop = Me.shpLeft.Top
        lngWidth = Me.shpRight.Left + Me.shpRight.Width - Me.shpLeft.Left
        lngHeight = Me.shpLeft.Height
        Me.txtNotes.ZOrder vbBringToFront
    ElseIf blnIcons Then
        lngLeft = Me.lblBuildClass(0).Left
        lngTop = Me.lblBuildClass(0).Top
        Me.lblNotes.Move lngLeft, lngTop
        Me.lblNotes.Caption = "Notes"
        lngLeft = Me.txtName.Left
        lngTop = Me.cboBuildClass(0).Top
        lngWidth = Me.txtName.Width
        lngHeight = Me.picIcons.Top - lngTop - Me.TextHeight("X")
    Else
        lngLeft = Me.lblBuildClass(0).Left
        lngTop = Me.lblBuildClass(2).Top + (Me.lblBuildClass(0).Top - Me.lblAlignment.Top) \ 2
        Me.lblNotes.Move lngLeft, lngTop
        Me.lblNotes.Caption = "Notes:"
        With Me.shpLeft
            lngSpacing = .Left + .Width - (Me.txtName.Left + Me.txtName.Width)
            lngLeft = .Left + lngSpacing
            lngWidth = .Width - lngSpacing * 2
        End With
        lngTop = lngTop + mlngNotesLabelOffsetY
        lngBottom = Me.picGrid.Top + mlngTop + mlngHeight * 20
        lngHeight = lngBottom - lngTop
    End If
    Me.txtNotes.Move lngLeft, lngTop, lngWidth, lngHeight
    ' Visible
    Me.shpLeft.Visible = Not mblnNotes
    Me.lblName.Visible = Not mblnNotes
    Me.txtName.Visible = Not mblnNotes
    Me.lblRace.Visible = Not mblnNotes
    Me.cboRace.Visible = Not mblnNotes
    Me.txtRace.Visible = Not mblnNotes
    Me.usrRace.Visible = False
    Me.lblAlignment.Visible = Not mblnNotes
    Me.cboAlignment.Visible = Not mblnNotes
    Me.lblMaxLevels.Visible = Not mblnNotes
    Me.usrSpinner.Visible = Not mblnNotes
    For i = 0 To 2
        Me.lblBuildClass(i).Visible = Not (blnIcons Or mblnNotes)
        Me.cboBuildClass(i).Visible = Not (blnIcons Or mblnNotes)
    Next
    Me.lblNotes.Visible = Not mblnNotes
    Me.picIcons.Visible = blnIcons And Not mblnNotes
    Me.shpRight.Visible = Not mblnNotes
    Me.picGrid.Visible = Not mblnNotes
    xp.UnlockWindow
End Sub




' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help"
            ShowHelp "Overview"
            Exit Sub
        Case "Overview": mblnNotes = False
        Case "Notes": mblnNotes = True
    End Select
    ShowControls
    If mblnNotes Then Me.txtNotes.SetFocus
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    cfg.SavePosition Me
    Select Case pstrCaption
        Case "Stats >"
            If Not OpenForm("frmStats") Then Exit Sub
    End Select
    mblnOverride = True
    Unload Me
End Sub


' ************* INITIALIZE *************


Private Sub LoadData()
    Dim i As Long
    
    With db
        mblnOverride = True
        ' Alignment
        ComboClear Me.cboAlignment
        For i = 0 To 5
            ComboAddItem Me.cboAlignment, GetAlignmentName(i), i
        Next
        ' Max levels
        Me.usrSpinner.Max = MAX_LEVEL
        If Not (cfg.UseIcons And cfg.IconOverview) Then
            ' Build Classes
            ComboClear Me.cboBuildClass(0)
            For i = 0 To ceClasses - 1
                ComboAddItem Me.cboBuildClass(0), db.Class(i).ClassName, i
            Next
            RaceChange
            For i = 0 To 2
                ComboSetValue Me.cboBuildClass(i), build.BuildClass(i)
                ClassChange i
                mblnOverride = True
            Next
        End If
        mblnOverride = False
    End With
End Sub

Private Sub ShowBuild()
    Dim i As Long
    
    mblnOverride = True
    Me.txtName.Text = build.BuildName
    Me.txtRace.Text = GetRaceName(build.Race)
    ShowSubRaceLink
    ComboSetValue Me.cboAlignment, build.Alignment
    Me.usrSpinner.Value = build.MaxLevels
    For i = 0 To 2
        ComboSetValue Me.cboBuildClass(i), build.BuildClass(i)
    Next
    Me.txtNotes.Text = build.Notes
    If build.BuildClass(0) = ceAny Then mblnIconicClass = True
    mblnOverride = False
End Sub

Private Sub InitComboboxes()
    Dim i As Long
    
    For i = 0 To 2
        ComboListHeight Me.cboBuildClass(i), 20
    Next
End Sub


' ************* RACECOMBO *************


Private Sub cboRace_GotFocus()
    On Error Resume Next
    Me.txtRace.SetFocus
End Sub

Private Sub cboRace_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
    Shift = 0
End Sub

Private Sub txtRace_GotFocus()
    With Me.txtRace
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtRace_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown, vbKeyReturn
            If Me.usrRace.Visible Then Me.usrRace.KeyDown KeyCode Else RaceShow
        Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
            Me.usrRace.KeyDown KeyCode
        Case vbKeyF4
            RaceToggle
        Case vbKeyEscape
            RaceHide
            Exit Sub
        Case Else
            Exit Sub
    End Select
    KeyCode = 0
End Sub

Private Sub txtRace_LostFocus()
    RaceHide
End Sub

Private Sub txtRace_Click()
    If Me.usrRace.Visible Then RaceHide Else RaceShow
End Sub

Public Sub RaceShow(Optional pblnSetFocus As Boolean = False)
    If mblnNotes Then Exit Sub
    If pblnSetFocus Then
        On Error Resume Next
        Me.txtRace.SetFocus
        On Error GoTo 0
    End If
    Me.usrRace.ZOrder vbBringToFront
    With Me.cboRace
        Me.usrRace.Move .Left, .Top + .Height + PixelY
    End With
    Me.usrRace.DropDown build.Race
    Me.usrRace.Visible = True
End Sub

Private Sub RaceHide()
    Me.usrRace.Visible = False
End Sub

Public Sub RaceToggle()
    If Me.usrRace.Visible Then RaceHide Else RaceShow
End Sub

Private Sub usrRace_Click(Race As Long)
    If mblnOverride Then Exit Sub
    If CheckGuide() Then
        RaceHide
    Else
        build.Race = Race
        Me.txtRace.Text = GetRaceName(build.Race)
        RaceChange
        RaceHide
        SetDirty
    End If
    With Me.txtRace
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub RaceChange()
    If build.Race = reDrow Then
        build.IncludePoints(beChampion) = 0
        If build.BuildPoints = beChampion Then build.BuildPoints = beAdventurer
    ElseIf build.IncludePoints(beChampion) = 0 And mblnInclude32 Then
        build.IncludePoints(beChampion) = 1
    End If
    If Not mblnLoading Then
        ClearLevelingGuide
        CascadeChanges cceRace
    End If
    IconicClass
    ShowSubRaceLink
End Sub

Private Sub IconicClass()
    If Not mblnIconicClass Then Exit Sub
    ChangeBuildClass 0, db.Race(build.Race).IconicClass
    ShowBuild
    UpdateChoices
    DataChanged
    DrawGrid
End Sub

Private Sub ShowSubRaceLink()
    Dim blnSubRace As Boolean
    
    If build.Race <> reAny Then
        If db.Race(build.Race).SubRace <> reAny Then blnSubRace = True
    End If
    Me.lnkSubRace.Visible = blnSubRace
End Sub

Private Sub lnkSubRace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkSubRace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    RaceHide
End Sub

Private Sub lnkSubRace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim enRace As RaceEnum
    Dim enParent As RaceEnum
    Dim strMessage As String
    
    xp.SetMouseCursor mcHand
    enRace = build.Race
    If enRace = reAny Then Exit Sub
    enParent = db.Race(enRace).SubRace
    If enParent = reAny Then Exit Sub
    strMessage = GetRaceName(enRace) & " is a racial variant of " & GetRaceName(enParent) & " and for the purpose of Racial Reincarnation will award " & GetRaceName(enParent) & " Racial Past Lives."
    MsgBox strMessage, vbInformation, "Notice"
End Sub


' ************* BUILD *************


Private Sub DataChanged()
    CalculateBAB
    InitGrid
    If Not mblnLoading Then
        ClearLevelingGuide
        CascadeChanges cceClass
        SetDirty
    End If
End Sub

Private Sub ClearLevelingGuide()
    Dim typBuild As BuildGuideType
    Dim typGuide As GuideType
    
    build.Guide = typBuild
    Guide = typGuide
End Sub

Private Sub txtName_GotFocus()
    TextboxGotFocus Me.txtName
End Sub

Private Sub txtName_Change()
    If mblnOverride Then Exit Sub
    build.BuildName = Me.txtName.Text
    SetAppCaption
    SetDirty
End Sub

Private Sub cboAlignment_Click()
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    build.Alignment = ComboGetValue(Me.cboAlignment)
    UpdateChoices
    If Not mblnLoading Then
        CascadeChanges cceAlignment
        SetDirty
    End If
End Sub

Private Sub usrSpinner_Change()
    If mblnOverride Then Exit Sub
    build.MaxLevels = Me.usrSpinner.Value
    frmMain.mnuEdit(6).Enabled = (build.MaxLevels > 19)
    ShowBuild
    DrawGrid
    If Not mblnLoading Then
        CascadeChanges cceMaxLevel
        SetDirty
    End If
End Sub

Private Sub txtNotes_Change()
    If mblnOverride Then Exit Sub
    build.Notes = Me.txtNotes.Text
    SetDirty
End Sub

Private Sub cboBuildClass_Click(Index As Integer)
    Dim i As Long
    
    If mblnOverride Then Exit Sub
    If CheckGuide() Then
        mblnOverride = True
        ComboSetValue Me.cboBuildClass(Index), build.BuildClass(Index)
        mblnOverride = False
        Exit Sub
    End If
    ClassChange Index
    mblnIconicClass = (build.BuildClass(0) = ceAny)
End Sub

Private Sub ClassChange(ByVal plngIndex As Long)
    Dim enNewClass As ClassEnum
    Dim lngLast As Long
    Dim i As Long
    
    enNewClass = ComboGetValue(Me.cboBuildClass(plngIndex))
    If enNewClass = -1 Then enNewClass = ceAny
    If enNewClass = ceAny Then lngLast = 2 Else lngLast = plngIndex
    For i = plngIndex To lngLast
        ChangeBuildClass i, enNewClass
    Next
    RemoveDuplicates
    ShowBuild
    UpdateChoices
    DataChanged
End Sub

Private Sub ChangeBuildClass(plngIndex As Long, ByVal penNewClass As ClassEnum)
    Dim enOldClass As ClassEnum
    Dim blnEnableMenu As Boolean
    Dim i As Long
    
    enOldClass = build.BuildClass(plngIndex)
    build.BuildClass(plngIndex) = penNewClass
    If penNewClass = ceAny Then penNewClass = build.BuildClass(0)
    For i = 1 To 20
        If build.Class(i) = enOldClass Then build.Class(i) = penNewClass
    Next
End Sub

Private Sub RemoveDuplicates()
    With build
        If .BuildClass(2) = .BuildClass(1) Or .BuildClass(2) = .BuildClass(0) Then .BuildClass(2) = ceAny
        If .BuildClass(1) = .BuildClass(0) Then .BuildClass(1) = ceAny
        If .BuildClass(1) = ceAny And .BuildClass(2) <> ceAny Then .BuildClass(1) = .BuildClass(2)
    End With
End Sub

Private Sub UpdateChoices()
    If cfg.UseIcons And cfg.IconOverview Then DrawIcons Else FilterCombos
End Sub

' The more confused I get, the more comments I add. This function confused the hell out of me.
Private Sub FilterCombos()
    Dim blnFilterAlign(6) As Boolean
    Dim blnFilterClass() As Boolean
    Dim enClass As ClassEnum
    Dim enAlignment As AlignmentEnum
    Dim lngClass As Long
    Dim lngAlign As Long
    Dim lngAlignments As Long
    Dim i As Long
    
    ' Initially, all alignments are allowed
    For lngAlign = 0 To 6
        blnFilterAlign(lngAlign) = True
    Next
    ' Keep track of valid alignments in case only one is allowed
    lngAlignments = 6
    ' Filter out incompatible alignments based on classes chosen
    For lngClass = 0 To 2
        enClass = build.BuildClass(lngClass)
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) And Not db.Class(enClass).Alignment(lngAlign) Then
                blnFilterAlign(lngAlign) = False
                lngAlignments = lngAlignments - 1
            End If
        Next
    Next
    ' If only one alignment is allowed, go ahead and select it
    If lngAlignments = 1 Then
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) Then build.Alignment = lngAlign: Exit For
        Next
    End If
    ' Filter Main Class based on alignments possible...
    ReDim blnFilterClass(ceClasses - 1)
    If build.Alignment = aleAny Then
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) Then
                For lngClass = 0 To ceClasses - 1
                    If db.Class(lngClass).Alignment(lngAlign) Then blnFilterClass(lngClass) = True
                Next
            End If
        Next
    Else ' ...or the one alignment actually chosen
        For lngClass = 0 To ceClasses - 1
            blnFilterClass(lngClass) = db.Class(lngClass).Alignment(build.Alignment)
        Next
    End If
    ' Refresh alignment combobox
    mblnOverride = True
    ComboClear Me.cboAlignment
    For lngAlign = 0 To 6
        If blnFilterAlign(lngAlign) Then ComboAddItem Me.cboAlignment, GetAlignmentName(lngAlign), lngAlign
    Next
    ComboSetValue Me.cboAlignment, build.Alignment
    ' Refresh class comboboxes
    For i = 0 To 2
        ComboClear Me.cboBuildClass(i)
        ' Enable controls
        If i > 0 Then Me.cboBuildClass(i).Enabled = (build.BuildClass(i - 1) <> ceAny)
        ' Refresh dropdown list
        If Me.cboBuildClass(i).Enabled Then
            For lngClass = 0 To ceClasses - 1
                If blnFilterClass(lngClass) Then ComboAddItem Me.cboBuildClass(i), db.Class(lngClass).ClassName, lngClass
            Next
        End If
        ComboSetValue Me.cboBuildClass(i), build.BuildClass(i)
        ' Remove already chosen classes from subsequent comboboxes (Main class chosen removed from first splash, etc...)
        If build.BuildClass(i) <> ceAny Then blnFilterClass(build.BuildClass(i)) = False
    Next
    mblnOverride = False
End Sub


' ************* CLASSES *************


Private Sub DrawIcons()
    Dim blnFilterAlign(6) As Boolean
    Dim blnFilterClass() As Boolean
    Dim enClass As ClassEnum
    Dim enAlignment As AlignmentEnum
    Dim lngClass As Long
    Dim lngAlign As Long
    Dim lngAlignments As Long
    Dim enClassState As ClassStateEnum
'    Dim strFile As String
    Dim strResource As String
    Dim i As Long
    
    
    ' Initially, all alignments are allowed
    For lngAlign = 0 To 6
        blnFilterAlign(lngAlign) = True
    Next
    ' Keep track of valid alignments in case only one is allowed
    lngAlignments = 6
    ' Filter out incompatible alignments based on classes chosen
    For lngClass = 0 To 2
        enClass = build.BuildClass(lngClass)
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) And Not db.Class(enClass).Alignment(lngAlign) Then
                blnFilterAlign(lngAlign) = False
                lngAlignments = lngAlignments - 1
            End If
        Next
    Next
    ' If only one alignment is allowed, go ahead and select it
    If lngAlignments = 1 Then
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) Then
                build.Alignment = lngAlign
                Exit For
            End If
        Next
    End If
    ' Filter Classes based on alignments possible...
    ReDim blnFilterClass(ceClasses - 1)
    If build.Alignment = aleAny Then
        For lngAlign = 1 To 6
            If blnFilterAlign(lngAlign) Then
                For lngClass = 0 To ceClasses - 1
                    If db.Class(lngClass).Alignment(lngAlign) Then blnFilterClass(lngClass) = True
                Next
            End If
        Next
    Else ' ...or the one alignment actually chosen
        For lngClass = 0 To ceClasses - 1
            blnFilterClass(lngClass) = db.Class(lngClass).Alignment(build.Alignment)
        Next
    End If
    
    'Filter class based off of disallowed
    For lngClass = 0 To 2
        If build.BuildClass(lngClass) Then
        End If
        'filter disallowedclasses
        If db.Class(build.BuildClass(lngClass)).DisallowedClasses > 0 Then
            For i = 1 To db.Class(build.BuildClass(lngClass)).DisallowedClasses
                blnFilterClass(db.Class(build.BuildClass(lngClass)).DisallowedClassID(i)) = False
            Next
        End If
    Next
    
    ' Refresh alignment combobox
    mblnOverride = True
    ComboClear Me.cboAlignment
    For lngAlign = 0 To 6
        If blnFilterAlign(lngAlign) Then ComboAddItem Me.cboAlignment, GetAlignmentName(lngAlign), lngAlign
    Next
    ComboSetValue Me.cboAlignment, build.Alignment
    ' Draw Icons
    For i = 1 To ceClasses - 1
        Select Case i
            Case build.BuildClass(0), build.BuildClass(1), build.BuildClass(2)
                enClassState = cseEmpty
            Case Else
                If blnFilterClass(i) Then
                    enClassState = cseNormal
                Else
                    enClassState = cseDisabled
                End If
        End Select
        If enClassState = cseEmpty Then
            strResource = GetClassResourceID(ceEmpty)
        Else
            strResource = GetClassResourceID(i)
        End If
        With Me.picClass(i)
            Me.picClass(i).PaintPicture LoadResPicture(strResource, vbResBitmap), 0, 0, .Width, .Height
            'Me.picClass(i).Picture LoadResPicture(strResource, vbResBitmap)
            'copy the pic to the drag icon  - BUT you can't because they're different formats
            '  See CreateIconFromResourceEx API
            'Me.picClass(i).Refresh
            'Me.picClass(i).DragIcon = Me.picClass(i).Picture
        End With
        If enClassState = cseDisabled Then
            GrayScale Me.picClass(i)
            Me.picClass(i).Enabled = False
            Me.lblClass(i).ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
        Else
            Me.picClass(i).Enabled = True
            Me.lblClass(i).ForeColor = cfg.GetColor(cgeWorkspace, cveText)
            If enClassState = cseNormal Then Me.picClass(i).DragMode = vbAutomatic Else Me.picClass(i).DragMode = vbManual
        End If
    Next
    mblnOverride = False
End Sub

Private Sub DrawIconHeader(plngCol As Long)
    Dim lngBuildClass As Long
    Dim enClass As ClassEnum
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
'    Dim strFile As String
    Dim strResource As String
    
    lngBuildClass = plngCol - 2
    ' Clear area
    lngHeight = (mlngTop - PixelY) * 0.9
    lngWidth = mlngWidth * 0.9
    lngHeight = Me.picGrid.ScaleY(Me.picGrid.ScaleX(lngWidth, vbTwips, vbPixels), vbPixels, vbTwips)
    lngLeft = Col(plngCol).Left + (mlngWidth - lngWidth) \ 2
    lngTop = (mlngTop - lngHeight) \ 2
    ' Find bitmap
    enClass = build.BuildClass(lngBuildClass)
    If enClass = ceAny Then
'        strFile = App.Path & "\Images\Empty.bmp"
        strResource = GetClassResourceID(ceEmpty)
        Me.picBuildClass(lngBuildClass).ToolTipText = vbNullString
        Me.picBuildClass(lngBuildClass).DragMode = vbManual
    Else
'        strFile = App.Path & "\Images\" & db.Class(enClass).ClassName & ".bmp"
        strResource = GetClassResourceID(enClass)
        Me.picBuildClass(lngBuildClass).ToolTipText = db.Class(enClass).ClassName
        Me.picBuildClass(lngBuildClass).DragIcon = Me.picClass(enClass).DragIcon
        Me.picBuildClass(lngBuildClass).DragMode = vbAutomatic
    End If
'    If Not xp.File.Exists(strFile) Then Stop
    ' Draw slot
    With Me.picBuildClass(lngBuildClass)
        .Move lngLeft, lngTop, lngWidth, lngHeight
        .Enabled = True
'        .PaintPicture LoadPicture(strFile), 0, 0, lngWidth, lngHeight
        .PaintPicture LoadResPicture(strResource, vbResBitmap), 0, 0, .Width, .Height
        If lngBuildClass > 0 Then
            If build.BuildClass(lngBuildClass - 1) = ceAny Then
                GrayScale Me.picBuildClass(lngBuildClass)
                .Enabled = False
            End If
        End If
        .Visible = True
    End With
End Sub


' ************* DRAG CLASSES *************


Private Sub picClass_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picClass(Index).DragMode = vbAutomatic Then xp.SetMouseCursor mcHand
    RaceHide
End Sub

Private Sub picClass_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picClass(Index).DragMode = vbAutomatic Then xp.SetMouseCursor mcHand
End Sub

Private Sub picClass_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    RaceHide
End Sub

Private Sub picClass_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picClass(Index).DragMode = vbAutomatic Then xp.SetMouseCursor mcHand
End Sub

Private Sub picBuildClass_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picBuildClass(Index).DragMode = vbAutomatic Then xp.SetMouseCursor mcHand
    RaceHide
End Sub

Private Sub picBuildClass_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell 0, 0
    If picBuildClass(Index).DragMode = vbAutomatic Then xp.SetMouseCursor mcHand
End Sub

Private Sub picBuildClass_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    RaceHide
End Sub

Private Sub picBuildClass_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picBuildClass(Index).DragMode = vbAutomatic Then xp.SetMouseCursor mcHand
End Sub

Private Sub picBuildClass_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Dim enClass As ClassEnum
    
    Select Case Source.Name
        Case "picClass"
            If CheckGuide() Then Exit Sub
            enClass = Source.Index
            ChangeBuildClass CLng(Index), Source.Index
            DragComplete
        Case "picBuildClass"
            If build.BuildClass(Index) <> ceAny And build.BuildClass(Source.Index) <> ceAny And Index <> Source.Index Then
                enClass = build.BuildClass(Index)
                build.BuildClass(Index) = build.BuildClass(Source.Index)
                build.BuildClass(Source.Index) = enClass
                ShowBuild
                UpdateChoices
                InitGrid
                InitLevelingGuide
                SetDirty
'                DragComplete
            End If
    End Select
    mblnIconicClass = False
End Sub

Private Sub picClass_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    If Source.Name = "picBuildClass" Then ClearBuildClass Source.Index
End Sub

Private Sub ClearBuildClass(Index As Integer)
    Dim i As Long
    
    If CheckGuide() Then Exit Sub
    For i = Index To 2
        ChangeBuildClass i, ceAny
    Next
    DragComplete
End Sub

Private Sub DragComplete()
    ShowBuild
    UpdateChoices
    DataChanged
End Sub

Private Sub picGrid_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "picBuildClass" Then ClearBuildClass Source.Index
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "picBuildClass" Then ClearBuildClass Source.Index
End Sub

Private Sub picIcons_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Name = "picBuildClass" Then ClearBuildClass Source.Index
End Sub


' ************* DRAW GRID *************


Private Sub InitGrid()
    Dim lngLeft As Long
    Dim lngBABWidth As Long
    Dim lngClassLevel() As Long
    Dim lngExtra As Long
    Dim i As Long
    
    Erase grid
    With Me.picGrid
        .FontBold = False
        ' Width
        mlngOffsetX = .TextWidth(" ")
        .FontBold = True
        lngExtra = .TextWidth("   ")
        lngLeft = InitColumn(0, "Level", 0, .TextWidth("Level"), vbCenter)
        lngLeft = InitColumn(1, "Class", lngLeft + lngExtra, .TextWidth("  Favored Soul  "), vbLeftJustify)
        lngBABWidth = .TextWidth(" BAB ")
        mlngWidth = (.ScaleWidth - lngLeft - lngBABWidth - lngExtra - PixelX) \ 3
        ' round off width to whole pixel
        mlngWidth = .ScaleX(.ScaleX(mlngWidth, vbTwips, vbPixels), vbPixels, vbTwips)
        For i = 0 To 2
            lngLeft = InitColumn(i + 2, ClassText(i, True), lngLeft, mlngWidth, vbCenter)
        Next
        lngLeft = InitColumn(5, "BAB", lngLeft + lngExtra, lngBABWidth, vbCenter)
        ' Height
        If cfg.UseIcons And cfg.IconOverview Then
            mlngTop = .ScaleY(.ScaleX(mlngWidth, vbTwips, vbPixels), vbPixels, vbTwips)
        Else
            mlngTop = Me.picGrid.TextHeight("X") * 2
        End If
        mlngHeight = (.ScaleHeight - mlngTop - PixelY) \ 20
        ' round of height to whole pixel
        mlngHeight = .ScaleY(.ScaleY(mlngHeight, vbTwips, vbPixels), vbPixels, vbTwips)
        mlngOffsetY = (mlngHeight - .TextHeight("X")) \ 2
        .FontBold = False
    End With
    ReDim mlngBuildClass(ceClasses - 1)
    For i = 0 To 2
        mlngBuildClass(build.BuildClass(i)) = i + 1
    Next
    ReDim lngClassLevel(ceClasses - 1)
    For i = 1 To HeroicLevels()
        lngClassLevel(build.Class(i)) = lngClassLevel(build.Class(i)) + 1
        grid(i, 0) = i
        grid(i, 1) = ClassText(mlngBuildClass(build.Class(i)) - 1, False)
        grid(i, 2) = lngClassLevel(build.BuildClass(0))
        grid(i, 3) = lngClassLevel(build.BuildClass(1))
        grid(i, 4) = lngClassLevel(build.BuildClass(2))
        grid(i, 5) = build.BAB(i)
    Next
    DrawGrid
End Sub

Private Function InitColumn(plngCol As Long, pstrHeader As String, plngLeft As Long, plngWidth As Long, penAlign As AlignmentConstants) As Long
    With Col(plngCol)
        .Header = pstrHeader
        .Left = plngLeft
        .Width = plngWidth
        .Right = .Left + .Width
        .Align = penAlign
        InitColumn = .Right
    End With
    grid(0, plngCol) = pstrHeader
End Function

Private Function ClassText(plngBuildClass As Long, pblnHeader As Boolean) As String
    If build.BuildClass(plngBuildClass) <> ceAny Then
        If pblnHeader Then
            ClassText = db.Class(build.BuildClass(plngBuildClass)).Initial(3)
        Else
            ClassText = db.Class(build.BuildClass(plngBuildClass)).ClassName
        End If
    End If
End Function

Private Sub DrawGrid()
    Dim lngRow As Long
    Dim lngCol As Long
    
    With Me.picGrid
        .Cls
        For lngRow = 0 To HeroicLevels()
            For lngCol = 0 To 5
                DrawCell lngRow, lngCol, (lngRow = 0)
            Next
        Next
    End With
End Sub

Private Sub DrawCell(plngRow As Long, plngCol As Long, pblnBold As Boolean, Optional pblnDrop As Boolean = False)
    Dim enGroup As ColorGroupEnum
    Dim strText As String
    Dim lngBackColor As Long
    Dim lngBorderColor As Long
    Dim enStyle As GridStyleEnum
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngTextLeft As Long
    Dim lngTextTop As Long
    Dim lngWidth As Long
    
    If cfg.UseIcons And cfg.IconOverview And plngRow = 0 And plngCol > 1 And plngCol < 5 Then
        DrawIconHeader plngCol
        Exit Sub
    End If
    If plngRow > 0 And plngCol = 5 And build.BuildClass(0) = ceAny Then Exit Sub
    If plngRow > 0 And plngCol > 1 And plngCol < 5 Then
        If build.Class(plngRow) <> ceAny And plngCol - 1 = mlngBuildClass(build.Class(plngRow)) Then enStyle = gseDropSlot Else enStyle = gseControl
        If build.BuildClass(plngCol - 2) <> ceAny Then strText = grid(plngRow, plngCol)
    ElseIf plngRow > 0 And plngCol = 1 And build.BuildClass(1) <> ceAny Then
        enStyle = gseDropSlot
        strText = grid(plngRow, plngCol)
    Else
        enStyle = gseWorkspace
        strText = grid(plngRow, plngCol)
    End If
    With Me.picGrid
        If pblnBold Then .FontBold = True
        ' Coordinates
        lngWidth = .TextWidth(strText)
        With Col(plngCol)
            lngLeft = .Left
            lngTop = mlngTop + (plngRow - 1) * mlngHeight
            lngRight = .Right
            lngBottom = lngTop + mlngHeight
            Select Case .Align
                Case vbLeftJustify: lngTextLeft = .Left + Me.picGrid.TextWidth("  ")
                Case vbCenter: lngTextLeft = .Left + (.Width - lngWidth) \ 2
                Case vbRightJustify: lngTextLeft = .Right - lngWidth
            End Select
            lngTextTop = lngTop + mlngOffsetY
        End With
        ' Colors
        Select Case enStyle
            Case gseWorkspace
                .ForeColor = cfg.GetColor(cgeWorkspace, cveText)
                lngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
                lngBorderColor = cfg.GetColor(cgeWorkspace, cveBackground)
            Case gseControl
                .ForeColor = cfg.GetColor(cgeControls, cveTextDim)
                lngBackColor = cfg.GetColor(cgeControls, cveBackground)
                lngBorderColor = cfg.GetColor(cgeDropSlots, cveBorderInterior)
            Case gseDropSlot
                .ForeColor = cfg.GetColor(cgeDropSlots, cveText)
                If pblnDrop Then
                    lngBackColor = cfg.GetColor(cgeDropSlots, cveBackHighlight)
                Else
                    lngBackColor = cfg.GetColor(cgeDropSlots, cveBackground)
                End If
                lngBorderColor = cfg.GetColor(cgeDropSlots, cveBorderInterior)
        End Select
        Me.picGrid.Line (lngLeft, lngTop)-(lngRight, lngBottom), lngBorderColor, B
        Me.picGrid.Line (lngLeft + PixelX, lngTop + PixelY)-(lngRight - PixelX, lngBottom - PixelY), lngBackColor, BF
        Me.picGrid.CurrentX = lngTextLeft
        Me.picGrid.CurrentY = lngTextTop
        Me.picGrid.Print strText
        If pblnBold Then .FontBold = False
    End With
End Sub


' ************* USE GRID *************


Private Sub picGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
End Sub

Private Sub picGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell X, Y
    If mlngRow = 0 Then Exit Sub
    Select Case mlngCol
        Case 1: DragClass
        Case 2 To 4: ChangeLevelSplit
    End Select
End Sub

Private Sub DragClass()
    Dim lngCount As Long
    Dim i As Long
    
    For i = 1 To HeroicLevels()
        If build.Class(i) <> build.Class(mlngRow) Then
            DrawCell i, 1, False, True
            lngCount = lngCount + 1
        End If
    Next
    If lngCount Then
        DrawBorders cfg.GetColor(cgeDropSlots, cveBorderHighlight)
        Me.picGrid.OLEDropMode = vbOLEDropManual
        Me.picGrid.OLEDrag
    End If
End Sub

Private Sub ChangeLevelSplit()
    Dim lngLevel As Long
    Dim enClass As ClassEnum
    
    enClass = build.BuildClass(mlngCol - 2)
    lngLevel = mlngRow
    If enClass = ceAny Or build.Class(lngLevel) = enClass Then Exit Sub
    If CheckGuide() Then Exit Sub
    build.Class(lngLevel) = enClass
    DataChanged
    ActiveBorders mlngRow, mlngCol
End Sub

Private Sub picGrid_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove
    Data.SetData "List"
End Sub

Private Sub picGrid_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
    
    CoordsToCell X, Y, lngRow, lngCol
    If lngRow = 0 Or lngCol <> 1 Then
        Effect = vbDropEffectNone
    ElseIf build.Class(lngRow) = build.Class(mlngRow) Then
        Effect = vbDropEffectNone
    Else
        Effect = vbDropEffectMove
    End If
End Sub

Private Sub picGrid_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim enSwap As ClassEnum
    
    CoordsToCell X, Y, lngRow, lngCol
    If lngRow = 0 Or lngCol <> 1 Then Exit Sub
    If build.Class(lngRow) = build.Class(mlngRow) Then Exit Sub
    If CheckGuide() Then Exit Sub
    enSwap = build.Class(lngRow)
    build.Class(lngRow) = build.Class(mlngRow)
    build.Class(mlngRow) = enSwap
    mblnDrop = True
End Sub

Private Sub picGrid_OLECompleteDrag(Effect As Long)
    If mblnDrop Then DataChanged
    mblnDrop = False
    DrawGrid
End Sub

Private Sub ActiveCell(X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    CoordsToCell X, Y, lngRow, lngCol
    If lngRow <> mlngRow Then
        ActiveBorders lngRow, lngCol
    ElseIf lngCol <> mlngCol Then
        ActiveBorders lngRow, lngCol
    End If
End Sub

Private Sub CoordsToCell(X As Single, Y As Single, plngRow As Long, plngCol As Long)
    Dim i As Long
    
    If Y < mlngTop Then plngRow = 0 Else plngRow = (Y - mlngTop) \ mlngHeight + 1
    If plngRow < 1 Or plngRow > HeroicLevels() Then plngRow = 0
    plngCol = 0
    For i = 1 To 4
        If X >= Col(i).Left And X <= Col(i).Right Then
            plngCol = i
            Exit For
        End If
    Next
End Sub

Private Sub ActiveBorders(plngRow As Long, plngCol As Long)
    DrawBorders cfg.GetColor(cgeDropSlots, cveBorderInterior)
    HighlightRow False
    mlngRow = plngRow
    mlngCol = plngCol
    HighlightRow True
    DrawBorders cfg.GetColor(cgeDropSlots, cveBorderHighlight)
End Sub

Private Sub DrawBorders(plngColor As Long)
    Dim lngTop As Long
    
    If mlngCol = 1 And build.BuildClass(1) = ceAny Then Exit Sub
    If mlngRow > 0 And mlngRow < 21 And mlngCol > 0 And mlngCol < 5 Then
        lngTop = mlngTop + (mlngRow - 1) * mlngHeight
        Me.picGrid.Line (Col(mlngCol).Left, lngTop)-(Col(mlngCol).Right, lngTop + mlngHeight), plngColor, B
    End If
End Sub

Private Sub HighlightRow(pblnBold As Boolean)
    If mlngRow > 0 And mlngRow < 21 Then
        DrawCell mlngRow, 0, pblnBold
        DrawCell mlngRow, 5, pblnBold
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ActiveCell 0, 0
End Sub
