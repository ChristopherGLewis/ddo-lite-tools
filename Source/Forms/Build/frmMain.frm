VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Character Builder Lite"
   ClientHeight    =   8160
   ClientLeft      =   135
   ClientTop       =   510
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   13530
   Begin VB.Timer tmrDeprecate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   900
      Top             =   5760
   End
   Begin VB.Timer tmrWindowMenu 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   900
      Top             =   5220
   End
   Begin VB.Timer tmrData 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   900
      Top             =   4140
   End
   Begin VB.Timer tmrOutput 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   900
      Top             =   4680
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1932
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   0
      Width           =   2052
      Begin VB.PictureBox picBuild 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1452
         Left            =   240
         ScaleHeight     =   1455
         ScaleWidth      =   1575
         TabIndex        =   3
         Top             =   240
         Width           =   1572
      End
   End
   Begin VB.HScrollBar scrollHorizontal 
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   1932
      Left            =   2220
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lblTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "tmrDeprecate: Used on build open: Load build and show output, *then* show deprecation screen."
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Index           =   3
      Left            =   1380
      TabIndex        =   8
      Top             =   5820
      Visible         =   0   'False
      Width           =   11652
   End
   Begin VB.Label lblTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "tmrData: Used only on startup: This main form is shown to user before any data files are read."
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Index           =   2
      Left            =   1380
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   11652
   End
   Begin VB.Label lblTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "tmrWindowMenu: Used to update the Window top menu whenever any form is opened or closed."
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Index           =   1
      Left            =   1380
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   11652
   End
   Begin VB.Label lblTimer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "tmrOutput: Data Screen forms use this to redraw complete output when (after) they close."
      ForeColor       =   &H00E0E0E0&
      Height          =   252
      Index           =   0
      Left            =   1380
      TabIndex        =   5
      Top             =   4740
      Visible         =   0   'False
      Width           =   11652
   End
   Begin VB.Label lblLoading 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   " Loading Data... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   300
      Left            =   5280
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   2004
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&New..."
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Import"
         Index           =   7
         Begin VB.Menu mnuImport 
            Caption         =   "from Forums..."
            Index           =   0
         End
         Begin VB.Menu mnuImport 
            Caption         =   "from Ron's Planner..."
            Index           =   1
         End
         Begin VB.Menu mnuImport 
            Caption         =   "from DDO Builder..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Export"
         Enabled         =   0   'False
         Index           =   8
         Begin VB.Menu mnuExport 
            Caption         =   "to Forums..."
            Index           =   0
         End
         Begin VB.Menu mnuExport 
            Caption         =   "to Ron's Planner..."
            Index           =   1
         End
         Begin VB.Menu mnuExport 
            Caption         =   "to DDO Builder..."
            Index           =   2
         End
         Begin VB.Menu mnuExport 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuExport 
            Caption         =   "Leveling &Guide"
            Enabled         =   0   'False
            Index           =   4
            Begin VB.Menu mnuExportGuide 
               Caption         =   "&csv"
               Index           =   0
            End
            Begin VB.Menu mnuExportGuide 
               Caption         =   "&txt"
               Index           =   1
            End
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU1"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU2"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU3"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU4"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU5"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU6"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU7"
         Index           =   18
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU8"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "MRU9"
         Index           =   20
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuEdit 
         Caption         =   "&Overview..."
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "St&ats..."
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "S&kills..."
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Feats..."
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "S&pells..."
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Enhancements..."
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Destiny..."
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Gear..."
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Tools"
      Index           =   2
      Begin VB.Menu mnuTools 
         Caption         =   "&Messages..."
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Con&vert Builds..."
         Index           =   1
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Refresh Data..."
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTools 
         Caption         =   "E&xceptions"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Error Log"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Colors..."
         Index           =   6
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Options..."
         Index           =   7
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Window"
      Index           =   3
      Begin VB.Menu mnuWindow 
         Caption         =   "Cascade Windows"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmOverview"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmStats"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmSkills"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmFeats"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmSpells"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmEnhancements"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmDestiny"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmGear"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmExport"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmImport"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmConvert"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmDeprecate"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmOptions"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmColors"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmColorFile"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmFormat"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "frmHelp"
         Index           =   18
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Help"
      Index           =   4
      Begin VB.Menu mnuHelp 
         Caption         =   "What's New?"
         Index           =   0
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Frequently Asked &Questions"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Save File Specifications"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&DDO Forums"
         Index           =   4
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&ddowiki Dashboard"
         Index           =   5
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
Option Explicit

Private mblnCanDrag As Boolean
Private mblnDragging As Boolean
Private mlngX As Long
Private mlngY As Long

Private mstrCaption() As String
Private mblnLoaded As Boolean

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.SizeWindow
    cfg.ShowMRU
    RefreshColors
    UpdateToolsMenu
    InitFonts Me.picBuild
    ' If it took longer than 400 milliseconds to load and process data last time program ran, display "Loading Data..."
    If cfg.ProcessTime > 400 Then Me.lblLoading.Visible = True
    Me.tmrData.Enabled = True
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    ResizeOutput
    With Me.lblLoading
        .Move (Me.ScaleWidth - .Width) \ 2, (Me.ScaleHeight - .Height) \ 2
    End With
    If cfg.OutputRefresh Then
        cfg.OutputRefresh = False
        GenerateOutput oeRemember
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CheckDirty() Then
        Cancel = True
        Exit Sub
    End If
    cfg.SaveWindowSize
    DeleteBackup
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    CloseApp
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    If IsOver(Me.picContainer.hwnd, Xpos, Ypos) Then WheelScroll lngValue
End Sub

Public Sub UpdateToolsMenu()
    Dim i As Long
    
    For i = 0 To Me.mnuTools.UBound
        Select Case StripMenuChars(Me.mnuTools(i).Caption)
            Case "Refresh Data":
                
                Me.mnuTools(i).Visible = xp.DebugMode
                
            Case "Messages":
                Me.mnuTools(i).Visible = gtypDeprecate.Deprecated
        End Select
    Next
End Sub

Public Sub RefreshColors()
    Me.BackColor = cfg.GetColor(cgeOutput, cveBackground)
    Me.picContainer.BackColor = cfg.GetColor(cgeOutput, cveBackground)
    Me.picContainer.Cls
    Me.picBuild.BackColor = cfg.GetColor(cgeOutput, cveBackground)
    Me.picBuild.Cls
    Me.Refresh
    GenerateOutput oeRemember
End Sub

Private Sub tmrData_Timer()
    Dim sngStart As Single
    Dim sngStop As Single
    Dim blnVisible As Boolean
    Dim i As Long
    
    Me.tmrData.Enabled = False
    sngStart = Timer
    xp.Mouse = msSystemWait
    ' Load and process data
    ClearLog
    InitHelp
    InitData
    ProcessData
    blnVisible = CreateErrorLog()
    ClearLog
    ' Show "Error Log" menu option if there were errors
    For i = 0 To Me.mnuTools.UBound
        If StripMenuChars(Me.mnuTools(i).Caption) = "Error Log" Then Me.mnuTools(i).Visible = blnVisible
    Next
    UpdateToolsMenu
    DeleteOldFiles
    Me.lblLoading.Visible = False
    xp.Mouse = msNormal
    sngStop = Timer
    ' Be careful we aren't crossing midnight, in which case sngStart will be very large and sngStop will be very small
    If sngStop > sngStart Then cfg.ProcessTime = Int((sngStop - sngStart) * 1000)
    If Not LoadBackup() Then
        If Not CheckCommandLine() Then
            If Not cfg.ConvertOnStartup Then
                If Dir(App.Path & "\Save\*.bld") <> vbNullString Then
                    OpenForm "frmConvert"
                    ShowHelp "Convert"
                    Exit Sub
                End If
            End If
        End If
    End If
    If UserTemplateOlder() Then ShowHelp "Stat Template Change"
End Sub

Private Sub tmrDeprecate_Timer()
    Me.tmrDeprecate.Enabled = False
    ShowDeprecated
End Sub

Private Sub tmrOutput_Timer()
    Me.tmrOutput.Enabled = False
    GenerateOutput oeAll
End Sub


' ************* MENUS *************


Private Sub mnuFile_Click(Index As Integer)
    Dim strCaption As String
    
    strCaption = StripMenuChars(Me.mnuFile(Index).Caption)
    Select Case strCaption
        Case "New": BuildNew
        Case "Open": BuildOpen
        Case "Close": BuildClose
        Case "Save": BuildSave
        Case "Save As": BuildSaveAs
        Case "Exit": Unload Me
        Case "Import", "Export"
        Case Else: MRUOpen strCaption
    End Select
End Sub

Private Sub MRUOpen(ByVal pstrMRU As String)
    Dim strMRU As String
    Dim strFile As String
    Dim strMessage As String
    
    strFile = cfg.GetMRUTarget(pstrMRU)
    Select Case BuildOpenMRU(strFile)
        Case leeNoError, leeUnexpectedError: Exit Sub
        Case leeFileNotFound: strMessage = "File not found."
        Case leeUnrecognized: strMessage = "Not a valid build file."
        Case leeUnsupported: strMessage = "Unsupported version."
    End Select
    If AskAlways(strMessage & " Remove from MRU list?", True) Then cfg.RemoveMRU pstrMRU
End Sub

Private Sub mnuImport_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuImport(Index).Caption)
        Case "from Forums": OpenForm "frmImport"
        Case "from Ron's Planner": ImportFileRon
        Case "from DDO Builder": ImportFileBuilder
    End Select
End Sub

Private Sub mnuExport_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuExport(Index).Caption)
        Case "to Forums": OpenForm "frmExport"
        Case "to Ron's Planner": ExportFileRon
        Case "to DDO Builder": ExportFileBuilder
    End Select
End Sub

Private Sub mnuExportGuide_Click(Index As Integer)
    ExportGuide StripMenuChars(Me.mnuExportGuide(Index).Caption)
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuEdit(Index).Caption)
        Case "Overview": OpenForm "frmOverview"
        Case "Stats": OpenForm "frmStats"
        Case "Skills": OpenForm "frmSkills"
        Case "Feats": OpenForm "frmFeats"
        Case "Enhancements": OpenForm "frmEnhancements"
        Case "Spells": OpenForm "frmSpells"
        Case "Destiny": OpenForm "frmDestiny"
        Case "Gear": OpenForm "frmGear"
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If Not BuildIsOpen() Then Exit Sub
    Select Case KeyCode
        Case vbKeyO: OpenForm "frmOverview"
        Case vbKeyT: OpenForm "frmStats"
        Case vbKeyK: OpenForm "frmSkills"
        Case vbKeyF: OpenForm "frmFeats"
        Case vbKeyE: OpenForm "frmEnhancements"
        Case vbKeyP: OpenForm "frmSpells"
        Case vbKeyD: OpenForm "frmDestiny"
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub mnuTools_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuTools(Index).Caption)
        Case "Convert Builds": If BuildClose() Then OpenForm "frmConvert"
        Case "Messages": ShowDeprecated
        Case "Refresh Data": OpenForm "frmRefreshData"
        Case "Error Log": ViewErrorLog True
        Case "Exceptions": ViewExceptions
        Case "Colors": cfg.RunUtil ueColors
        Case "Options": OpenForm "frmOptions"
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Select Case StripMenuChars(Me.mnuHelp(Index).Caption)
        Case "DDO Forums": xp.OpenURL "https://www.ddo.com/forums/forum.php"
        Case "ddowiki Dashboard": xp.OpenURL "https://ddowiki.com/page/User:EllisDee37/Dashboard"
        Case "About": frmAbout.Show vbModal, Me
        Case "What's New?": ShowHelp "What's New?"
        Case "Save File Specifications": ShowHelp "Save_File_Specifications"
        Case "F.A.Q.", "FAQ", "Frequently Asked Questions": ShowHelp "Frequently Asked Questions"
    End Select
End Sub

Private Sub mnuWindow_Click(Index As Integer)
    Dim frm As Form
    
    Select Case StripMenuChars(Me.mnuWindow(Index).Caption)
        Case "Cascade Windows": If BuildIsOpen() Then cfg.CascadeWindows
        Case Else: If GetForm(frm, mstrCaption(Index)) Then frm.SetFocus
    End Select
End Sub

Public Sub UpdateWindowMenu()
    Me.tmrWindowMenu.Enabled = True
End Sub

Public Sub WindowActivate()
On Error GoTo WindowActivateSkip
    Dim i As Long
    
    For i = 2 To Me.mnuWindow.UBound
        With Me.mnuWindow(i)
            If .Visible Then .Checked = (mstrCaption(i) = Screen.ActiveForm.Name) Else .Checked = False
        End With
    Next
    
WindowActivateExit:
    Exit Sub
    
WindowActivateSkip:
    Resume WindowActivateExit
End Sub

Private Sub tmrWindowMenu_Timer()
    Dim lngCount As Long
    Dim frm As Form
    Dim i As Long
    
    Me.tmrWindowMenu.Enabled = False
    If Not mblnLoaded Then InitWindowCaptions
    With frmMain
        For i = 2 To .mnuWindow.UBound
            If GetForm(frm, mstrCaption(i)) Then
                .mnuWindow(i).Caption = frm.Caption
                .mnuWindow(i).Visible = True
                lngCount = lngCount + 1
            Else
                .mnuWindow(i).Visible = False
            End If
        Next
        .mnuWindow(0).Enabled = (lngCount > 1)
        .mnuWindow(1).Visible = (lngCount > 0)
    End With
    Me.WindowActivate
End Sub

Private Sub InitWindowCaptions()
    Dim i As Long
    
    With frmMain
        ReDim mstrCaption(.mnuWindow.UBound)
        For i = 2 To .mnuWindow.UBound
            mstrCaption(i) = .mnuWindow(i).Caption
        Next
    End With
    mblnLoaded = True
End Sub


' ************* OUTPUT *************


Public Property Let CanDrag(ByVal pblnCanDrag As Boolean)
    mblnCanDrag = pblnCanDrag
End Property

Private Sub picBuild_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnCanDrag Then
        xp.SetMouseCursor mcHand
        mblnDragging = True
        mlngX = X
        mlngY = Y
    End If
End Sub

Private Sub picBuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long
    Dim lngTop As Long
    
    If mblnCanDrag Then xp.SetMouseCursor mcHand
    If Not (Button = vbLeftButton And mblnDragging) Then Exit Sub
    With Me.picBuild
        lngLeft = .ScaleX(.Left + X - mlngX, vbTwips, vbPixels)
        lngTop = .ScaleY(.Top + Y - mlngY, vbTwips, vbPixels)
        ValidCoords lngLeft, lngTop
        .Move Me.ScaleX(lngLeft, vbPixels, vbTwips), Me.ScaleY(lngTop, vbPixels, vbTwips)
    End With
End Sub

Private Sub picBuild_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnDragging = False
    If mblnCanDrag Then xp.SetMouseCursor mcHand
End Sub

Private Sub ValidCoords(plngLeft As Long, plngTop As Long)
    mblnOverride = True
    With Me.scrollHorizontal
        If .Visible Then
            If plngLeft < 0 - .Max Then plngLeft = 0 - .Max
            If plngLeft > 0 Then plngLeft = 0
            .Value = 0 - plngLeft
        Else
            plngLeft = Me.picBuild.Left
        End If
    End With
    With Me.scrollVertical
        If .Visible Then
            If plngTop < 0 - .Max Then plngTop = 0 - .Max
            If plngTop > 0 Then plngTop = 0
            .Value = 0 - plngTop
        Else
            plngTop = Me.picBuild.Top
        End If
    End With
    mblnOverride = False
End Sub


' ************* SCROLLBARS *************


Private Sub scrollHorizontal_GotFocus()
    Me.picBuild.SetFocus
End Sub

Private Sub scrollHorizontal_Scroll()
    HorizontalScroll
End Sub

Private Sub scrollHorizontal_Change()
    HorizontalScroll
End Sub

Private Sub HorizontalScroll()
    If mblnOverride Then Exit Sub
    Me.picBuild.Left = 0 - Me.ScaleX(Me.scrollHorizontal.Value, vbPixels, vbTwips)
End Sub

Private Sub scrollVertical_GotFocus()
    Me.picBuild.SetFocus
End Sub

Private Sub scrollVertical_Scroll()
    VerticalScroll
End Sub

Private Sub scrollVertical_Change()
    If mblnOverride Then Exit Sub
    VerticalScroll
End Sub

Private Sub VerticalScroll()
    Me.picBuild.Top = 0 - Me.ScaleY(Me.scrollVertical.Value, vbPixels, vbTwips)
End Sub

Private Sub WheelScroll(plngValue As Long)
    Dim lngValue As Long
    
    If Not Me.scrollVertical.Visible Then Exit Sub
    With Me.scrollVertical
        lngValue = .Value - plngValue * Me.ScaleY(Me.picBuild.TextHeight("Q"), vbTwips, vbPixels)
        If lngValue < .Min Then lngValue = .Min
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub
