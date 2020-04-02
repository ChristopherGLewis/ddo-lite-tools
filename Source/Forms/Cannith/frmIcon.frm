VERSION 5.00
Begin VB.Form frmIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFFF&
   Caption         =   "Farms"
   ClientHeight    =   8160
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8484
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   8484
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    CloseApp
End Sub
