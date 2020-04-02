VERSION 5.00
Begin VB.Form frmPatrons 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Patrons"
   ClientHeight    =   3840
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5808
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatrons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5808
   ShowInTaskbar   =   0   'False
   Begin Compendium.userFavor usrFavor 
      Height          =   2292
      Left            =   480
      TabIndex        =   0
      Top             =   900
      Width           =   4572
      _ExtentX        =   8065
      _ExtentY        =   4043
   End
End
Attribute VB_Name = "frmPatrons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
' All code moved into usercontrol userFavor
Option Explicit


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    Me.usrFavor.Init 0
    SizeControl
End Sub

Private Sub SizeControl()
    With Me.usrFavor
        Me.Width = .FitWidth + Me.Width - Me.ScaleWidth
        Me.Height = .FitHeight + Me.Height - Me.ScaleHeight + Me.TextHeight("Q") \ 2
        .Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
End Sub

Public Sub ReDrawForm()
    Me.usrFavor.ReDrawControl
    SizeControl
End Sub

Public Sub Recalculate()
    Me.usrFavor.Recalculate
End Sub
