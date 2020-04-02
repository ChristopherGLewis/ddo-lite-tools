VERSION 5.00
Begin VB.Form frmConvert 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Convert Builds"
   ClientHeight    =   7452
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9708
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7452
   ScaleWidth      =   9708
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "..."
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   8952
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   900
      Width           =   312
   End
   Begin VB.Frame fraList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4212
      Left            =   420
      TabIndex        =   16
      Top             =   1620
      Width           =   4272
      Begin VB.ListBox lstBuilds 
         Appearance      =   0  'Flat
         Height          =   3912
         Left            =   0
         TabIndex        =   5
         Top             =   300
         Width           =   4272
      End
      Begin VB.Label lblBuilds 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No Binary Builds Found"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4272
      End
   End
   Begin CharacterBuilderLite.userCheckBox usrchkStartup 
      Height          =   252
      Left            =   420
      TabIndex        =   14
      Tag             =   "nav"
      Top             =   7128
      Width           =   3852
      _ExtentX        =   6795
      _ExtentY        =   445
      Value           =   0   'False
      Caption         =   "Don't show this screen on startup"
   End
   Begin CharacterBuilderLite.userHeader usrFooter 
      Height          =   384
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7068
      Width           =   9708
      _ExtentX        =   17124
      _ExtentY        =   677
      Spacing         =   264
      BorderColor     =   -2147483640
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   5940
   End
   Begin VB.Timer tmrStop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6420
      Top             =   5940
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   672
      Left            =   4920
      ScaleHeight     =   672
      ScaleWidth      =   4392
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1860
      Width           =   4392
      Begin VB.CheckBox chkConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Convert All"
         ForeColor       =   &H80000008&
         Height          =   552
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   2052
      End
      Begin VB.CheckBox chkConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Convert Selected"
         ForeColor       =   &H80000008&
         Height          =   552
         Index           =   1
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   2052
      End
   End
   Begin CharacterBuilderLite.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9708
      _ExtentX        =   17124
      _ExtentY        =   677
      Spacing         =   264
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   420
      TabIndex        =   2
      Text            =   "Path"
      Top             =   900
      Width           =   8532
   End
   Begin CharacterBuilderLite.userDetails usrDetails 
      Height          =   2772
      Left            =   4992
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3060
      Width           =   4272
      _ExtentX        =   7535
      _ExtentY        =   4890
   End
   Begin VB.Label lblMessages 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Messages"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4980
      TabIndex        =   9
      Top             =   2760
      Width           =   1692
   End
   Begin VB.Label lnkStop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Stop"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Left            =   2580
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblConverting 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Converting Build: "
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   420
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   1584
   End
   Begin VB.Label lblProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   420
      TabIndex        =   13
      Top             =   6420
      Visible         =   0   'False
      Width           =   8832
   End
   Begin VB.Label lblFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Folder"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   420
      TabIndex        =   1
      Top             =   636
      Width           =   972
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ConvertEnum
    cveStop
    cveAll
    cveSelected
End Enum

Private menAction As ConvertEnum
Private mstrFile As String

Private mblnOverride As Boolean

Private Sub Form_Load()
    cfg.Configure Me
    Me.usrchkStartup.Value = cfg.ConvertOnStartup
    Me.txtPath.Text = App.Path & "\Save"
    FindFiles
End Sub

Private Sub Form_Activate()
    ActivateForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
    frmMain.UpdateWindowMenu
End Sub

Private Sub usrHeader_Click(pstrCaption As String)
    If pstrCaption = "Help" Then ShowHelp "Convert_Builds"
End Sub

Private Sub usrchkStartup_UserChange()
    cfg.ConvertOnStartup = Me.usrchkStartup.Value
End Sub

Private Sub FindFiles()
    Dim strRoot As String
    Dim strFile As String
    Dim blnError As Boolean
    Dim lngCount As Long
    
    xp.LockWindow Me.lstBuilds.hwnd
    ListboxClear Me.lstBuilds
    strRoot = GetRoot()
    If LCase$(strRoot) = LCase$(GetStorage()) Then
        blnError = True
    Else
        strFile = Dir(strRoot & "*.bld")
        Do While Len(strFile)
            lngCount = lngCount + 1
            Me.lstBuilds.AddItem strFile
            strFile = Dir
        Loop
    End If
    EnableControls
    If blnError Then
        Me.lblBuilds.Caption = "Invalid Source Path"
    Else
        Select Case lngCount
            Case 0: Me.lblBuilds.Caption = "No Binary Builds Found"
            Case 1: Me.lblBuilds.Caption = "1 Binary Build Found"
            Case Else: Me.lblBuilds.Caption = lngCount & " Binary Builds Found"
        End Select
    End If
    xp.UnlockWindow
End Sub

Private Sub EnableControls()
    Dim blnEnabled As Boolean
    
    If Me.Visible Then Me.picFocus.SetFocus
    blnEnabled = (menAction = cveStop)
    If blnEnabled Then Me.lblFolder.ForeColor = cfg.GetColor(cgeWorkspace, cveText) Else Me.lblFolder.ForeColor = cfg.GetColor(cgeWorkspace, cveTextDim)
    EnableControl Me.txtPath, blnEnabled
    EnableControl Me.txtPath, blnEnabled
    EnableControl Me.chkPath, blnEnabled
    EnableControl Me.fraList, blnEnabled
    EnableControl Me.chkConvert(0), (Me.lstBuilds.ListCount > 0 And blnEnabled = True)
    EnableControl Me.chkConvert(1), (Me.lstBuilds.ListIndex <> -1 And blnEnabled = True)
End Sub

Private Sub EnableControl(pctl As Control, pblnEnabled As Boolean)
    If pctl.Enabled <> pblnEnabled Then pctl.Enabled = pblnEnabled
End Sub

Private Sub txtPath_GotFocus()
    TextboxGotFocus Me.txtPath
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FindFiles
    End If
End Sub

Private Function GetRoot() As String
    Dim strRoot As String
    
    strRoot = Me.txtPath.Text
    If Right$(Me.txtPath.Text, 1) <> "\" Then strRoot = strRoot & "\"
    GetRoot = strRoot
End Function

Private Function GetStorage() As String
    GetStorage = App.Path & "\Save\Binary\"
End Function

Private Sub chkPath_Click()
    Dim strPath As String
    
    If UncheckButton(Me.chkPath, mblnOverride) Then Exit Sub
'    strPath = xp.Folder.Browse("Folder with *.bld files:")
    strPath = xp.ShowOpenDialog(cfg.BuilderPath, "Build Files|*.bld", "*.bld")
    If Len(strPath) Then
        strPath = GetPathFromFilespec(strPath)
        Me.txtPath.Text = strPath
        FindFiles
    End If
    Me.picFocus.SetFocus
End Sub

Private Sub lstBuilds_Click()
    Me.chkConvert(1).Enabled = (Me.lstBuilds.ListIndex <> -1)
End Sub

Private Sub chkConvert_Click(Index As Integer)
    If UncheckButton(Me.chkConvert(Index), mblnOverride) Then Exit Sub
    Me.usrDetails.Clear
    EnableControls
    menAction = Index + 1
    GetNextFile
End Sub

Private Sub lnkStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    menAction = cveStop
End Sub

Private Sub lnkStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub GetNextFile()
    Dim lngIndex As Long
    
    Select Case menAction
        Case cveStop: lngIndex = -1
        Case cveAll: lngIndex = 0
        Case cveSelected: lngIndex = Me.lstBuilds.ListIndex
    End Select
    If Me.lstBuilds.ListCount = 0 Then lngIndex = -1
    If lngIndex = -1 Then
        menAction = cveStop
        mstrFile = vbNullString
        Me.lblConverting.Visible = False
        Me.lblProgress.Caption = vbNullString
        Me.lblProgress.Visible = False
        Me.lnkStop.Visible = False
        EnableControls
    Else
        Me.picFocus.SetFocus
        mstrFile = Me.lstBuilds.List(lngIndex)
        Me.lblConverting.Visible = True
        Me.lblProgress.Caption = mstrFile
        Me.lblProgress.Visible = True
        Me.lnkStop.Visible = True
        Me.tmrStart.Enabled = True
    End If
End Sub

Private Sub tmrStart_Timer()
    Me.tmrStart.Enabled = False
    ConvertFile
End Sub

Private Sub ConvertFile()
    Dim strSource As String
    Dim strBinary As String
    Dim strText As String
    
    ClearBuild True
    xp.Folder.Create GetStorage()
    strSource = GetRoot() & mstrFile
    strBinary = xp.File.MakeNameUnique(GetStorage() & mstrFile)
    xp.File.Move strSource, strBinary
    If Not ShowError(LoadBuild(strBinary, False, True)) Then
        strText = xp.File.MakeNameUnique(GetRoot() & GetNameFromFilespec(mstrFile) & ".build")
        SaveFileLite strText
    End If
    Me.tmrStop.Enabled = True
End Sub

Private Function ShowError(penError As LoadErrorEnum) As Boolean
    Dim lngError As Long
    
    Select Case penError
        Case leeNoError: Exit Function
        Case leeFileNotFound: Me.usrDetails.AddText "File not found:"
        Case leeUnrecognized: Me.usrDetails.AddText "Not a valid build file:"
        Case leeUnsupported: Me.usrDetails.AddText "Unsupported version:"
        Case Else
            Me.usrDetails.AddText "Unexpected error loading file:"
            lngError = penError
    End Select
    Me.usrDetails.AddText mstrFile
    If lngError Then
        On Error Resume Next
        Err.Raise penError
        Me.usrDetails.AddText "Error: " & Err.Description
        On Error GoTo 0
    End If
    Me.usrDetails.AddText " "
    Me.usrDetails.Refresh
    ShowError = True
End Function

Private Sub tmrStop_Timer()
    Me.tmrStop.Enabled = False
    FindFiles
    GetNextFile
End Sub

