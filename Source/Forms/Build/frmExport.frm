VERSION 5.00
Begin VB.Form frmExport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export"
   ClientHeight    =   7764
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   12216
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7764
   ScaleWidth      =   12216
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9180
      Top             =   1380
   End
   Begin VB.PictureBox picFocus 
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   60
      ScaleHeight     =   252
      ScaleWidth      =   192
      TabIndex        =   0
      Tag             =   "nav"
      Top             =   60
      Width           =   192
   End
   Begin VB.PictureBox picTooltip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   756
      ScaleHeight     =   240
      ScaleWidth      =   840
      TabIndex        =   9
      Top             =   324
      Visible         =   0   'False
      Width           =   840
      Begin VB.Label lblTooltip 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Copied! "
         ForeColor       =   &H80000017&
         Height          =   240
         Left            =   0
         TabIndex        =   10
         Tag             =   "tip"
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.ComboBox cboSection 
      Height          =   312
      ItemData        =   "frmExport.frx":000C
      Left            =   8640
      List            =   "frmExport.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "nav"
      Top             =   30
      Width           =   1932
   End
   Begin VB.Timer tmrTooltip 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2100
      Top             =   0
   End
   Begin VB.ComboBox cboFormat 
      Height          =   312
      ItemData        =   "frmExport.frx":0079
      Left            =   5160
      List            =   "frmExport.frx":007B
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "nav"
      Top             =   30
      Width           =   2112
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1968
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   3192
   End
   Begin VB.Shape shpBorder 
      Height          =   672
      Left            =   1200
      Top             =   1260
      Visible         =   0   'False
      Width           =   1272
   End
   Begin VB.Label lblSection 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Section:"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   7836
      TabIndex        =   5
      Tag             =   "nav"
      Top             =   84
      Width           =   744
   End
   Begin VB.Label lnkNav 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Format:"
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
      Left            =   4272
      TabIndex        =   3
      Tag             =   "nav"
      Top             =   84
      Width           =   792
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
      TabIndex        =   2
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
      TabIndex        =   1
      Tag             =   "nav"
      Top             =   84
      Width           =   1728
   End
   Begin VB.Label lnkNav 
      Alignment       =   1  'Right Justify
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
      Left            =   11520
      TabIndex        =   7
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
      Width           =   12216
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    mblnOverride = True
    Me.Visible = True
    Me.Refresh
    UpdateCombo
    cfg.OutputSection = oeAll
    ComboSetValue Me.cboSection, cfg.OutputSection
    mblnOverride = False
    Me.Refresh
    Me.tmrLoad.Enabled = True
End Sub

Private Sub Form_Activate()
    ActivateForm
    Me.tmrLoad.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
    cfg.SaveSettings
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    If cfg.GetColor(cgeNavigation, cveBackground) = cfg.GetColor(cgeControls, cveBackground) Then lngTop = (Me.shpNav.Height * 3) \ 2 Else lngTop = Me.shpNav.Height
    lngHeight = Me.ScaleHeight - lngTop
    With Me.shpBorder
        .Move 0, lngTop, Me.ScaleWidth, lngHeight
        Me.txtOutput.Move PixelX, .Top + PixelY, .Width - PixelX * 2, .Height - PixelY * 2
        .Visible = True
    End With
    Me.txtOutput.Visible = True
End Sub

Private Sub tmrLoad_Timer()
    Me.tmrLoad.Enabled = False
    Me.txtOutput.Text = GenerateOutput(oeExport)
    GenerateOutput oeAll
End Sub

Public Sub Cascade()
    Me.txtOutput.Text = GenerateOutput(oeExport)
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
    
    xp.SetMouseCursor mcHand
    If Button <> vbLeftButton Then Exit Sub
    Select Case Me.lnkNav(Index).Caption
        Case "Copy to Clipboard"
            Clipboard.Clear
            Clipboard.SetText Me.txtOutput.Text
            Me.picTooltip.Visible = True
            Me.tmrTooltip.Enabled = True
        Case "Save to File"
            If Len(Me.txtOutput.Text) = 0 Then
                Notice "Nothing to export."
                Exit Sub
            End If
            strFile = xp.File.MakeNameDOS(build.BuildName) & ".txt"
            strFile = xp.ShowSaveAsDialog(App.Path & "\Save", strFile, "Text Files (*.txt)|*.txt", "*.txt")
            If Len(strFile) Then xp.File.SaveStringAs strFile, Me.txtOutput.Text
        Case "Format:"
            frmFormat.Show vbModeless, Me
        Case "Help"
            ShowHelp "Export"
    End Select
End Sub

Private Sub tmrTooltip_Timer()
    Me.tmrTooltip.Enabled = False
    Me.picTooltip.Visible = False
End Sub


' ************* GENERAL *************


Public Sub UpdateCombo()
    Dim strName() As String
    Dim i As Long
    
    mblnOverride = True
    cfg.GetFormatNames strName
    ComboClear Me.cboFormat
    Me.cboFormat.AddItem vbNullString
    For i = 0 To UBound(strName)
        Me.cboFormat.AddItem strName(i)
    Next
    ComboSetText Me.cboFormat, cfg.IdentifyOutputFormat()
    mblnOverride = False
End Sub

Private Sub cboFormat_Click()
    If mblnOverride Then Exit Sub
    cfg.SetOutputFormat Me.cboFormat.Text
    Me.tmrLoad.Enabled = True
End Sub

Private Sub cboSection_Click()
    If mblnOverride Then Exit Sub
    cfg.OutputSection = ComboGetValue(Me.cboSection)
    Me.txtOutput.Text = GenerateOutput(oeExport)
End Sub
