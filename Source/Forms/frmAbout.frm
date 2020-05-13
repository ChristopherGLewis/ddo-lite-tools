VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2085
   ClientLeft      =   30
   ClientTop       =   405
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Updates by Chris Lewis"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   1700
      Width           =   3972
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Written by Ellis Dee"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   3972
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Version"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3972
   End
   Begin VB.Label lblAppName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Application Name"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3972
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private Sub Form_Load()
    cfg.RefreshColors Me
    Me.lblAppName.Caption = App.ProductName
    With Me.lblVersion
        .Caption = "v" & App.Major & "." & App.Minor
        If App.Revision <> 0 Then .Caption = .Caption & "." & App.Revision
    End With
End Sub

