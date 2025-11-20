VERSION 5.00
Begin VB.Form frmColorFile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Colors"
   ClientHeight    =   5700
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   7344
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColorFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7344
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   60
      ScaleHeight     =   252
      ScaleWidth      =   132
      TabIndex        =   0
      Tag             =   "nav"
      Top             =   60
      Width           =   132
   End
   Begin VB.PictureBox picTooltip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   708
      ScaleHeight     =   240
      ScaleWidth      =   852
      TabIndex        =   8
      Tag             =   "tip"
      Top             =   324
      Visible         =   0   'False
      Width           =   852
      Begin VB.Label tipClipboard 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Copied! "
         ForeColor       =   &H80000017&
         Height          =   240
         Left            =   0
         TabIndex        =   9
         Tag             =   "tip"
         Top             =   0
         Width           =   852
      End
   End
   Begin VB.Timer tmrTooltip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6240
      Top             =   0
   End
   Begin VB.TextBox txtColors 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2472
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   3552
   End
   Begin VB.Shape shpBorder 
      Height          =   552
      Left            =   2220
      Top             =   1080
      Visible         =   0   'False
      Width           =   672
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
      Left            =   6648
      TabIndex        =   6
      Tag             =   "nav"
      Top             =   84
      Width           =   432
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Save to File"
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
      Left            =   2556
      TabIndex        =   4
      Tag             =   "nav"
      Top             =   84
      Width           =   1152
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Copy to Clipboard"
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
      Left            =   264
      TabIndex        =   5
      Tag             =   "nav"
      Top             =   84
      Width           =   1728
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Apply Colors"
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
      Index           =   5
      Left            =   4860
      TabIndex        =   7
      Tag             =   "nav"
      Top             =   84
      Visible         =   0   'False
      Width           =   1224
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Paste from Clipboard"
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
      Left            =   264
      TabIndex        =   3
      Tag             =   "nav"
      Top             =   84
      Visible         =   0   'False
      Width           =   2064
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Load from File"
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
      Index           =   4
      Left            =   2892
      TabIndex        =   2
      Tag             =   "nav"
      Top             =   84
      Visible         =   0   'False
      Width           =   1404
   End
   Begin VB.Shape shpNav 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   384
      Left            =   0
      Tag             =   "nav"
      Top             =   0
      Width           =   7344
   End
End
Attribute VB_Name = "frmColorFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private menColorFile As ColorFileEnum

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    If frmColors.usrchkArea(cgeOutput).Value Then menColorFile = cfeOutput Else menColorFile = cfeScreen
    Select Case frmColors.ImportExport
        Case ieImport: Import
        Case ieExport: Export
    End Select
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    If cfg.GetColor(cgeNavigation, cveBackground) = cfg.GetColor(cgeControls, cveBackground) Then lngTop = (Me.shpNav.Height * 3) \ 2 Else lngTop = Me.shpNav.Height
    lngHeight = Me.ScaleHeight - lngTop
    With Me.shpBorder
        .Move 0, lngTop, Me.ScaleWidth, lngHeight
        Me.txtColors.Move PixelX, .Top + PixelY, .Width - PixelX * 2, .Height - PixelY * 2
        .Visible = True
    End With
    Me.txtColors.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    If App.Title = "Character Builder Lite" And GetForm(frm, "frmMain") Then frm.UpdateWindowMenu
End Sub

Private Sub Import()
    Dim i As Long
    
    Me.Caption = "Import Colors"
    For i = 1 To 5
        Me.lnkNav(i).Visible = (i > 2)
    Next
End Sub

Private Sub Export()
    Dim i As Long
    
    Me.Caption = "Export Colors"
    For i = 1 To 5
        Me.lnkNav(i).Visible = (i < 3)
    Next
    Me.txtColors.Text = cfg.MakeColorSettings(menColorFile)
End Sub

Public Sub RefreshColors()
    Form_Resize
End Sub


' ************* NAVIGATION *************


Private Sub lnkNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strFile As String
    Dim strFilters As String
    Dim strDefault As String
    
    xp.SetMouseCursor mcHand
    Select Case menColorFile
        Case cfeScreen
            strFilters = "Screen Colors (*.screen)|*.screen"
            strDefault = "*.screen"
        Case cfeOutput
            strFilters = "Output Colors (*.output)|*.output"
            strDefault = "*.output"
    End Select
    Select Case Me.lnkNav(Index).Caption
        Case "Copy to Clipboard"
            Clipboard.Clear
            Clipboard.SetText Me.txtColors.Text
            Me.picTooltip.Visible = True
            Me.tmrTooltip.Enabled = True
        Case "Paste from Clipboard"
            Me.txtColors.Text = Clipboard.GetText(vbCFText)
        Case "Load from File"
            strFile = xp.ShowOpenDialog(SettingsPath(), strFilters, strDefault)
            If Len(strFile) = 0 Then Exit Sub
            Me.txtColors.Text = xp.File.LoadToString(strFile)
        Case "Save to File"
            If Len(Me.txtColors.Text) = 0 Then
                MsgBox "No color info to save.", vbInformation, "Notice"
                Exit Sub
            End If
            strFile = xp.ShowSaveAsDialog(SettingsPath(), vbNullString, strFilters, strDefault)
            If Len(strFile) = 0 Then Exit Sub
            If xp.File.Exists(strFile) Then
                If MsgBox(GetFileFromFilespec(strFile) & " exists. Overwrite?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
            End If
            xp.File.SaveStringAs strFile, Me.txtColors.Text
        Case "Apply Colors"
            If Len(Me.txtColors.Text) = 0 Then
                MsgBox "No color info to apply.", vbInformation, "Notice"
                Exit Sub
            End If
'            xp.LockWindow frmOptions.hwnd
            cfg.ImportColors Me.txtColors.Text, cfeAll
            ColorChange
'            xp.UnlockWindow
        Case "Help"
            ShowHelp Me.Caption
    End Select
End Sub

Private Sub tmrTooltip_Timer()
    Me.tmrTooltip.Enabled = False
    Me.picTooltip.Visible = False
End Sub
