VERSION 5.00
Begin VB.Form frmMessages 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EllisoftLiteMessages"
   ClientHeight    =   2892
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4368
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMessages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   4368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtMessage 
      Height          =   2172
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   3492
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is how the various "Lite" programs communicate with each other
' Whenever one program needs to notify another of a change it should react to,
' it posts a message in the textbox of that app's frmMessages
' Note that there can be multiple instances of any or all of the apps
Option Explicit

Private Enum MessageEnum
    meNotes
    meLinkLists
End Enum

Private Sub txtMessage_Change()
    Dim strMessage As String
    
    If Len(Me.txtMessage.Text) = 0 Then Exit Sub
    strMessage = Me.txtMessage.Text
    Me.txtMessage.Text = vbNullString
    Select Case strMessage
        Case "Colors": cfg.ReQuery seColors
        Case "LinkLists": LinkLists
        Case "Dimensions"
        Case "Notes": Notes
    End Select
End Sub

Private Sub LinkLists()
'    Dim frm As Form
'
'    If CountCompendiums() > 1 Then
'        If GetForm(frm, "frmCompendium") Then frm.LinkLists
'    End If
End Sub

Private Sub Notes()
    Dim frm As Form
    
    If GetForm(frm, "frmCompendium") Then frm.Notes
End Sub

