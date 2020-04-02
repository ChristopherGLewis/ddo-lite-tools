VERSION 5.00
Begin VB.Form frmFormat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Output Format"
   ClientHeight    =   6540
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFormat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   900
      ScaleHeight     =   252
      ScaleWidth      =   312
      TabIndex        =   0
      Tag             =   "nav"
      Top             =   60
      Width           =   312
   End
   Begin CharacterBuilderLite.userCheckBox usrchkBBCodes 
      Height          =   252
      Left            =   420
      TabIndex        =   3
      Top             =   1380
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   445
      Caption         =   "BBCodes"
   End
   Begin CharacterBuilderLite.userCheckBox usrchkUseDots 
      Height          =   252
      Left            =   420
      TabIndex        =   2
      Top             =   1020
      Width           =   1812
      _ExtentX        =   2985
      _ExtentY        =   656
      Caption         =   "Use Dots"
   End
   Begin CharacterBuilderLite.userCheckBox usrchkColoredText 
      Height          =   252
      Left            =   420
      TabIndex        =   1
      Top             =   660
      Width           =   1812
      _ExtentX        =   3196
      _ExtentY        =   445
      Caption         =   "Text Colors"
   End
   Begin VB.Frame fraFormat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4872
      Left            =   300
      TabIndex        =   24
      Top             =   1380
      Width           =   4932
      Begin VB.TextBox txtWrapperOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   22
         Text            =   "[code]"
         Top             =   3960
         Width           =   1452
      End
      Begin VB.TextBox txtWrapperClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   23
         Text            =   "[/code]"
         Top             =   3960
         Width           =   1092
      End
      Begin VB.TextBox txtNumberedOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   16
         Text            =   "[list=1]"
         Top             =   2400
         Width           =   1452
      End
      Begin VB.TextBox txtNumberedClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   17
         Text            =   "[/list]"
         Top             =   2400
         Width           =   1092
      End
      Begin VB.TextBox txtColorClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   20
         Text            =   "[/color]"
         Top             =   2880
         Width           =   1092
      End
      Begin VB.TextBox txtBulletClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   15
         Text            =   "[/list]"
         Top             =   2040
         Width           =   1092
      End
      Begin VB.TextBox txtFixedClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   12
         Text            =   "[/font]"
         Top             =   1320
         Width           =   1092
      End
      Begin VB.TextBox txtUnderlineClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   9
         Text            =   "[/u]"
         Top             =   960
         Width           =   1092
      End
      Begin VB.TextBox txtBoldClose 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   3420
         TabIndex        =   6
         Text            =   "[/b]"
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox txtColorOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   19
         Text            =   "[color=#$]"
         Top             =   2880
         Width           =   1452
      End
      Begin VB.TextBox txtBulletOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   14
         Text            =   "[list]"
         Top             =   2040
         Width           =   1452
      End
      Begin VB.TextBox txtFixedOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   11
         Text            =   "[font=courier]"
         Top             =   1320
         Width           =   1452
      End
      Begin VB.TextBox txtUnderlineOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   8
         Text            =   "[u]"
         Top             =   960
         Width           =   1452
      End
      Begin VB.TextBox txtBoldOpen 
         Appearance      =   0  'Flat
         Height          =   324
         Left            =   1860
         TabIndex        =   5
         Text            =   "[b]"
         Top             =   600
         Width           =   1452
      End
      Begin CharacterBuilderLite.userCheckBox usrchkBold 
         Height          =   276
         Left            =   300
         TabIndex        =   4
         Top             =   600
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   572
         Caption         =   "Bold"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkUnderline 
         Height          =   276
         Left            =   300
         TabIndex        =   7
         Top             =   960
         Width           =   1452
         _ExtentX        =   2350
         _ExtentY        =   445
         Caption         =   "Underline"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkFixed 
         Height          =   276
         Left            =   300
         TabIndex        =   10
         Top             =   1320
         Width           =   1452
         _ExtentX        =   2350
         _ExtentY        =   445
         Caption         =   "Fixed"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkLists 
         Height          =   276
         Left            =   300
         TabIndex        =   13
         Top             =   1680
         Width           =   1452
         _ExtentX        =   2350
         _ExtentY        =   445
         Caption         =   "Lists"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkColor 
         Height          =   276
         Left            =   300
         TabIndex        =   18
         Top             =   2880
         Width           =   1452
         _ExtentX        =   2350
         _ExtentY        =   445
         Caption         =   "Color"
      End
      Begin CharacterBuilderLite.userCheckBox usrchkWrapper 
         Height          =   276
         Left            =   300
         TabIndex        =   21
         Top             =   3960
         Width           =   1452
         _ExtentX        =   2350
         _ExtentY        =   445
         Value           =   0   'False
         Caption         =   "Wrapper"
      End
      Begin VB.Shape Shape1 
         Height          =   4752
         Left            =   0
         Top             =   120
         Width           =   4932
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wrap the entire build in tags?"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   5
         Left            =   660
         TabIndex        =   30
         Top             =   4380
         Width           =   3852
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Numbered:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   300
         TabIndex        =   29
         Top             =   2424
         Width           =   1452
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bullet Points:"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   300
         TabIndex        =   28
         Top             =   2064
         Width           =   1452
      End
      Begin VB.Label lblLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Use $ as the placeholder for the numeric color value in the Color Open tag."
         ForeColor       =   &H80000008&
         Height          =   492
         Index           =   4
         Left            =   660
         TabIndex        =   27
         Top             =   3300
         Width           =   3852
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   3420
         TabIndex        =   26
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   1860
         TabIndex        =   25
         Top             =   360
         Width           =   1452
      End
   End
   Begin CharacterBuilderLite.userCheckBox usrchkReddit 
      Height          =   252
      Left            =   3660
      TabIndex        =   35
      Top             =   660
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   445
      Value           =   0   'False
      Caption         =   "Reddit"
      CheckPosition   =   1
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Apply Changes"
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
      Left            =   2676
      TabIndex        =   34
      Tag             =   "nav"
      Top             =   84
      Width           =   1452
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
      Left            =   4824
      TabIndex        =   33
      Tag             =   "nav"
      Top             =   84
      Width           =   432
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Load"
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
      TabIndex        =   32
      Tag             =   "nav"
      Top             =   84
      Width           =   480
   End
   Begin VB.Label lnkNav 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Save As"
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
      Left            =   1308
      TabIndex        =   31
      Tag             =   "nav"
      Top             =   84
      Width           =   804
   End
   Begin VB.Shape shpNav 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   384
      Left            =   0
      Tag             =   "nav"
      Top             =   0
      Width           =   5520
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOverride As Boolean
Private mblnDirty As Boolean


' ************* FORM *************


Private Sub Form_Load()
    mblnDirty = False
    LoadData
    cfg.Configure Me
End Sub

Private Sub Form_Activate()
    ActivateForm
End Sub

Private Sub Form_Resize()
    Me.usrchkBBCodes.Width = Me.usrchkBBCodes.FitWidth
End Sub

Public Sub RefreshColors()
    Me.Dirty = Me.Dirty ' Update link color
    EnableControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    If Me.Dirty Then
        If Not Ask("Apply changes?") Then Exit Sub
        ApplyChanges
    End If
    If App.Title = "Character Builder Lite" And GetForm(frm, "frmMain") Then frm.UpdateWindowMenu
End Sub


' ************* NAVIGATION *************


Private Sub lnkNav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 3 And Not Me.Dirty Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 3 And Not Me.Dirty Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkNav_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 3 And Not Me.Dirty Then Exit Sub
    xp.SetMouseCursor mcHand
    Select Case Me.lnkNav(Index).Caption
        Case "Load": FormatLoad
        Case "Save As": FormatSave
        Case "Apply Changes": ApplyChanges
        Case "Help": ShowHelp "Output Format"
    End Select
End Sub

Private Sub FormatLoad()
    Dim strFormat As String
    Dim strFile As String
    Dim frm As Form
    
    strFile = xp.ShowOpenDialog(App.Path & "\Settings", "Output Formats|*.bbcodes", "*.bbcodes")
    If GetForm(frm, "frmExport") Then frm.UpdateCombo
    If Len(strFile) = 0 Then Exit Sub
    strFormat = GetNameFromFilespec(strFile)
    If Not cfg.SetOutputFormat(strFormat) Then
        Notice "Unrecognized file format."
    Else
        If Not Me.Dirty Then Me.Dirty = True
    End If
End Sub

Private Sub FormatSave()
'    Dim strFormat As String
    Dim strFile As String
    Dim frm As Form
    
'    ApplyChanges
'    strFormat = cfg.IdentifyOutputFormat()
'    If Len(strFormat) Then
'        Notice "Already saved as " & GetFileFromFilespec(strFormat)
'        Exit Sub
'    End If
    strFile = xp.ShowSaveAsDialog(App.Path & "\Settings", "", "Output Formats|*.bbcodes", "*.bbcodes")
    If GetForm(frm, "frmExport") Then frm.UpdateCombo
    If Len(strFile) = 0 Then Exit Sub
    If xp.File.Exists(strFile) Then
        If Not AskAlways(GetFileFromFilespec(strFile) & " exists. Overwrite?") Then Exit Sub
    End If
    ApplyChanges
    cfg.SaveFormat strFile
    If GetForm(frm, "frmExport") Then frm.UpdateCombo
    Me.Dirty = False
End Sub

Private Sub ApplyChanges()
    Dim frm As Form
    
    cfg.WriteOutputFormat
    If GetForm(frm, "frmExport") Then
        frm.UpdateCombo
        frm.tmrLoad.Enabled = True
    Else
        cfg.IdentifyOutputFormat
        GenerateOutput oeRemember
    End If
    Me.Dirty = False
End Sub


' ************* INITIALIZE *************


Private Sub LoadData()
    mblnOverride = True
    cfg.ReadOutputFormat
    mblnOverride = False
End Sub


' ************* GENERAL *************


Public Property Get Dirty() As Boolean
    Dirty = mblnDirty
End Property

Public Property Let Dirty(ByVal pblnDirty As Boolean)
    mblnDirty = pblnDirty
    If mblnDirty Then
        Me.lnkNav(3).ForeColor = cfg.GetColor(cgeNavigation, cveTextLink)
    Else
        Me.lnkNav(3).ForeColor = cfg.GetColor(cgeNavigation, cveTextDim)
    End If
End Property

Public Sub EnableControls()
    Dim blnEnabled As Boolean
    Dim ctl As Control
    
    blnEnabled = Me.usrchkBBCodes.Value = True And Me.usrchkReddit.Value = False
    For Each ctl In Me.Controls
        Select Case ctl.Name
            Case "usrchkBold", "usrchkUnderline", "usrchkFixed", "usrchkLists", "usrchkColor", "usrchkWrapper", "lblLabel": EnableControl ctl, blnEnabled
            Case "fraFormat": ctl.Enabled = blnEnabled
            Case Else: If TypeOf ctl Is TextBox Then EnableControl ctl, blnEnabled
        End Select
    Next
    blnEnabled = Not Me.usrchkReddit.Value
    Me.usrchkUseDots.Enabled = blnEnabled
    Me.usrchkBBCodes.Enabled = blnEnabled
End Sub

Private Sub EnableControl(pctl As Control, pblnEnabled As Boolean)
    Select Case TypeName(pctl)
        Case "userCheckBox"
            pctl.Enabled = pblnEnabled
        Case "Label"
            If pblnEnabled Then
                pctl.ForeColor = cfg.GetColor(cgeWorkspace, cveText)
            Else
                pctl.ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
            End If
        Case "TextBox"
            pctl.Locked = Not pblnEnabled
            If pblnEnabled Then
                pctl.ForeColor = cfg.GetColor(cgeControls, cveText)
            Else
                pctl.ForeColor = cfg.GetColor(cgeControls, cveTextDim)
            End If
    End Select
End Sub

Private Sub UserChange()
    If mblnOverride Then Exit Sub
    Me.Dirty = True
End Sub

Private Sub txtBoldClose_Change()
    UserChange
End Sub

Private Sub txtBoldOpen_Change()
    UserChange
End Sub

Private Sub txtBulletClose_Change()
    UserChange
End Sub

Private Sub txtBulletOpen_Change()
    UserChange
End Sub

Private Sub txtColorClose_Change()
    UserChange
End Sub

Private Sub txtColorOpen_Change()
    UserChange
End Sub

Private Sub txtFixedClose_Change()
    UserChange
End Sub

Private Sub txtFixedOpen_Change()
    UserChange
End Sub

Private Sub txtNumberedClose_Change()
    UserChange
End Sub

Private Sub txtNumberedOpen_Change()
    UserChange
End Sub

Private Sub txtUnderlineClose_Change()
    UserChange
End Sub

Private Sub txtUnderlineOpen_Change()
    UserChange
End Sub

Private Sub txtWrapperClose_Change()
    UserChange
End Sub

Private Sub txtWrapperOpen_Change()
    UserChange
End Sub

Private Sub usrchkBBCodes_UserChange()
    EnableControls
    UserChange
End Sub

Private Sub usrchkBold_UserChange()
    UserChange
End Sub

Private Sub usrchkColor_UserChange()
    UserChange
End Sub

Private Sub usrchkFixed_UserChange()
    UserChange
End Sub

Private Sub usrchkLists_UserChange()
    UserChange
End Sub

Private Sub usrchkReddit_UserChange()
    EnableControls
    UserChange
End Sub

Private Sub usrchkUnderline_UserChange()
    UserChange
End Sub

Private Sub usrchkUseDots_UserChange()
    UserChange
End Sub

Private Sub usrchkNumbered_UserChange()
    UserChange
End Sub

Private Sub usrchkColoredText_UserChange()
    UserChange
End Sub

Private Sub usrchkWrapper_UserChange()
    UserChange
End Sub

