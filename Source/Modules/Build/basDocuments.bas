Attribute VB_Name = "basDocuments"
' Written by Ellis Dee
' Fairly generic document management routines. New, Open, Close, Save, EnableMenus, etc...
' Also code for deprecating builds on load
Option Explicit

Public Enum LoadErrorEnum
    leeNoError
    leeFileNotFound = 37371
    leeUnsupported = 37372
    leeUnrecognized = 37373
    leeUnexpectedError = 37374
End Enum

Private Const Signature As Long = 735346
Private Const Version As Byte = buildVersion  'New version

Private Const OpenFilter As String = "Build Files|*.build;*.bld|Text Build Files (*.build)|*.build|Binary Build Files (*.bld)|*.bld"
Private Const LiteExt As String = "*.build;*.bld"
Private Const RonExt As String = "*.txt"
Private Const BuilderExt As String = "*.ddocp"

Private mstrFile As String
Private mblnOpen As Boolean

Private msngLastBackup As Long


' ************* DOCUMENTS *************


Public Sub BuildNew()
    If CheckDirty() Then Exit Sub
    CloseForm "All"
    ClearBuild
    SetBuildDefaults
    SetDirty False
    SetAppCaption
    mblnOpen = True
    EnableMenus
    GenerateOutput oeAll
    OpenForm "frmOverview"
End Sub

Public Sub BuildOpen(Optional pstrFile As String = vbNullString)
    Dim strFile As String
    
    If Len(pstrFile) Then
        If CheckDirty() Then Exit Sub
        mstrFile = pstrFile
    Else
        mstrFile = OpenDialog(peLite)
    End If
    If Len(mstrFile) Then LoadBuild mstrFile, True
End Sub

Public Function OpenDialog(penPlanner As PlannerEnum) As String
    Dim strFile As String
    Dim strPath As String
    
    If CheckDirty() Then Exit Function
    Select Case penPlanner
        Case peLite
            strFile = xp.ShowOpenDialog(cfg.LitePath, OpenFilter, LiteExt)
            If Len(strFile) Then cfg.LitePath = GetPathFromFilespec(strFile)
        Case peRon
            strFile = xp.ShowOpenDialog(cfg.RonPath, "Ron's Planner Files|" & RonExt, RonExt)
            If Len(strFile) Then cfg.RonPath = GetPathFromFilespec(strFile)
        Case peBuilder
            strFile = xp.ShowOpenDialog(cfg.BuilderPath, "DDO Builder Files|" & BuilderExt, BuilderExt)
            If Len(strFile) Then cfg.BuilderPath = GetPathFromFilespec(strFile)
    End Select
    OpenDialog = strFile
End Function

Public Function BuildOpenMRU(pstrFile As String) As LoadErrorEnum
    Dim enStatus As LoadErrorEnum
    
    If CheckDirty() Then Exit Function
    mstrFile = pstrFile
    If Len(mstrFile) Then
        enStatus = LoadBuild(mstrFile, True, True)
        If enStatus = leeNoError Then
            mblnOpen = True
            cfg.AddMRU pstrFile
            EnableMenus
            SetAppCaption
            CascadeChanges cceAll
            SetDirty False
        Else
            mstrFile = vbNullString
        End If
    End If
    BuildOpenMRU = enStatus
End Function

Public Function BuildClose() As Boolean
    If CheckDirty() Then Exit Function
    ClearBuild
    CloseForm "All"
    SetDirty False
    SetAppCaption
    mblnOpen = False
    EnableMenus
    GenerateOutput oeAll
    BuildClose = True
End Function

' Return TRUE if cancelled
Public Function BuildSave() As Boolean
    If Len(mstrFile) = 0 Or LCase(GetExtFromFilespec(mstrFile)) = "bld" Then
        BuildSave = BuildSaveAs()
    ElseIf ReservedFolder(mstrFile) Then
        BuildSave = True
    Else
        SaveBuild mstrFile
        SetDirty False
        cfg.AddMRU mstrFile
        EnableMenus
    End If
End Function

' Return TRUE if cancelled
Public Function BuildSaveAs() As Boolean
    Dim strFile As String
    
    strFile = SaveAsDialog(peLite)
    If Len(strFile) = 0 Then
        BuildSaveAs = True
    ElseIf ReservedFolder(strFile) Then
        BuildSaveAs = True
    Else
        mstrFile = strFile
        SaveBuild mstrFile
        SetDirty False
        cfg.AddMRU mstrFile
        EnableMenus
    End If
End Function

Private Function ReservedFolder(pstrFile As String) As Boolean
    If LCase$(GetPathFromFilespec(pstrFile)) = LCase$(App.Path & "\Save\Binary") Then
        Notice "\Save\Binary is a reserved folder." & vbNewLine & vbNewLine & "Use the \Save folder instead.", True
        ReservedFolder = True
    End If
End Function

Public Function SaveAsDialog(penFormat As PlannerEnum) As String
    Dim strPath As String
    Dim strFile As String
    
    Select Case penFormat
        Case peLite
            strFile = xp.ShowSaveAsDialog(cfg.LitePath, xp.File.MakeNameDOS(build.BuildName), "Build Files|*.build", "*.build")
            If Len(strFile) Then cfg.LitePath = GetPathFromFilespec(strFile)
        Case peRon
            strFile = xp.ShowSaveAsDialog(cfg.RonPath, xp.File.MakeNameDOS(build.BuildName), "Ron's Planner Files|" & RonExt, RonExt)
            If Len(strFile) Then cfg.RonPath = GetPathFromFilespec(strFile)
        Case peBuilder
            strFile = xp.ShowSaveAsDialog(cfg.BuilderPath, xp.File.MakeNameDOS(build.BuildName), "DDO Builder Files|" & BuilderExt, BuilderExt)
            If Len(strFile) Then cfg.BuilderPath = GetPathFromFilespec(strFile)
    End Select
    If Len(strFile) = 0 Then Exit Function
    If xp.File.Exists(strFile) Then
        If Not AskAlways(GetFileFromFilespec(strFile) & " exists. Overwrite?") Then strFile = vbNullString
    End If
    SaveAsDialog = strFile
End Function

Public Sub SetAppCaption()
    Dim strCaption As String
    Dim strFile As String
    
    If Len(build.BuildName) Then
        strCaption = build.BuildName
        If Len(mstrFile) Then
            strFile = GetNameFromFilespec(mstrFile)
            If strCaption <> strFile Then strCaption = strCaption & " (" & GetFileFromFilespec(mstrFile) & ")"
        End If
        If cfg.Dirty Then strCaption = strCaption & "*"
    Else
        strCaption = App.ProductName
    End If
    If frmMain.Caption <> strCaption Then frmMain.Caption = strCaption
End Sub

Public Sub ClearBuild(Optional pblnClearFile As Boolean = True)
    Dim typBlank As BuildType6
    
    build = typBlank
    If pblnClearFile Then mstrFile = vbNullString
End Sub

Public Sub SetBuildDefaults()
    Dim i As Long
    
    build.BuildName = "New Build"
    build.MaxLevels = MAX_LEVEL
    build.BuildPoints = cfg.BuildPoints
    ReDim build.Feat(bftFeatTypes - 1)
    ReDim build.BAB(1 To MAX_LEVEL) ' Note: BAB needs to be redimmed to system max levels, not build max levels
    For i = 0 To 3
        build.IncludePoints(i) = 1
    Next
End Sub

' Used for importing
Public Sub SetBuildOpen(pblnOpen As Boolean)
    mblnOpen = pblnOpen
End Sub

Public Sub SetDirty(Optional pblnDirty As Boolean = True)
    Dim frm As Form
    
    cfg.Dirty = pblnDirty
    SetAppCaption
    GenerateOutput oeRemember
    If pblnDirty Then
        If GetForm(frm, "frmExport") Then frm.Cascade
    Else
        DeleteBackup
    End If
End Sub

' Returns TRUE if cancelled
Public Function CheckDirty() As Boolean
    If Not cfg.Dirty Then Exit Function
    Select Case AskCancel("Save changes to " & PromptName() & "?", True)
        Case vbYes: CheckDirty = BuildSave()
        Case vbCancel: CheckDirty = True
    End Select
End Function

Private Function PromptName() As String
    Dim strName As String
    
    If Len(mstrFile) = 0 Then
        strName = build.BuildName
    ElseIf GetNameFromFilespec(mstrFile) = build.BuildName Then
        strName = build.BuildName
    Else
        strName = GetFileFromFilespec(mstrFile)
    End If
    PromptName = strName
End Function

Public Sub EnableMenus(Optional pstrMostRecent As String = vbNullString)
    Dim i As Long
    
    With frmMain
        For i = 0 To .mnuFile.UBound
            With .mnuFile(i)
                Select Case StripMenuChars(.Caption)
                    Case "Close": .Enabled = mblnOpen
                    Case "Save": .Enabled = mblnOpen
                    Case "Save As": .Enabled = mblnOpen
                    Case "Export": .Enabled = mblnOpen
                End Select
            End With
        Next
        For i = 0 To .mnuExport.UBound
            With .mnuExport(i)
                Select Case StripMenuChars(.Caption)
                    Case "Leveling Guide": .Enabled = (mblnOpen And build.Guide.Enhancements > 0)
                End Select
            End With
        Next
        For i = 0 To .mnuEdit.UBound
            With .mnuEdit(i)
                Select Case StripMenuChars(.Caption)
                    Case "-"
                    Case "Spells": If mblnOpen Then .Enabled = (build.CanCastSpell(1) > 0) Else .Enabled = False
                    Case "Destiny": If mblnOpen Then .Enabled = (build.MaxLevels > 19) Else .Enabled = False
                    Case Else: .Enabled = mblnOpen
                End Select
            End With
        Next
    End With
    frmMain.UpdateToolsMenu
    frmMain.UpdateWindowMenu
    cfg.ShowMRU
End Sub

Public Function BuildIsOpen() As Boolean
    BuildIsOpen = mblnOpen
End Function

Public Sub BuildWasImported()
    mblnOpen = True
    EnableMenus
    SetAppCaption
    CascadeChanges cceAll
    SetDirty True
End Sub


' ************* BACKUP *************


Public Sub SaveBackup()
    Dim strFile As String
    
    If Not cfg.Dirty Then Exit Sub
    If Timer >= msngLastBackup And Timer - msngLastBackup < 1 Then Exit Sub
    msngLastBackup = Timer
    strFile = BackupName()
    SaveBuild strFile
End Sub

Public Function LoadBackup() As Boolean
    Dim strFile As String
    
    strFile = BackupName()
    If xp.File.Exists(strFile) Then
        If AskAlways("A backup build was autosaved. Open it now?" & vbNewLine & vbNewLine & "If you select No, the backup build will be deleted.", True) Then
            LoadBuild strFile, False
            SetDirty
            LoadBackup = True
        Else
            xp.File.Delete strFile
        End If
    End If
End Function

Public Function CheckCommandLine() As Boolean
    Dim strFile As String
    Dim strPath As String
    Dim strCheck As String
    
    If Len(gstrCommand) < 5 Then Exit Function
    strFile = gstrCommand
    If LCase$(Right$(strFile, 4)) <> ".bld" And LCase$(Right$(strFile, 6)) <> ".build" Then Exit Function
    If InStr(strFile, ":") = 0 And Left$(strFile, 1) <> "\" Then
        strPath = App.Path & "\"
        strCheck = cfg.CheckExtension(strPath & strFile)
        If Not xp.File.Exists(strCheck) Then strPath = strPath & "Save\"
        strFile = strPath & strFile
    End If
    strFile = cfg.CheckExtension(strFile)
    If xp.File.Exists(strFile) Then
        BuildOpen strFile
        CheckCommandLine = True
    End If
End Function

Public Sub DeleteBackup()
    Dim strFile As String
    
    strFile = BackupName()
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
End Sub

Private Function BackupName() As String
    BackupName = App.Path & "\Backup.bld"
End Function


' ************* EXPORT GUIDE *************


Public Sub ExportGuide(pstrExt As String)
    Dim strMask As String
    Dim strFile As String
    Dim strText As String
    
    If Guide.Enhancements = 0 Then
        Notice "Nothing to Export"
        Exit Sub
    End If
    If Len(mstrFile) Then
        strFile = GetNameFromFilespec(mstrFile) & "." & pstrExt
    ElseIf Len(build.BuildName) Then
        strFile = build.BuildName & "." & pstrExt
    Else
        strFile = "Leveling Guide." & pstrExt
    End If
    Select Case pstrExt
        Case "csv": strMask = "CSV File (*.csv)|*.csv|Text File (*.txt)|*.txt"
        Case "txt": strMask = "Text File (*.txt)|*.txt|CSV File (*.csv)|*.csv"
    End Select
    strFile = xp.ShowSaveAsDialog(App.Path & "\Save", strFile, strMask, "*." & pstrExt)
    If Len(strFile) = 0 Then Exit Sub
    If xp.File.Exists(strFile) Then
        If Not AskAlways(GetFileFromFilespec(strFile) & " exists. Overwrite?", True) Then Exit Sub
    End If
    Select Case LCase(GetExtFromFilespec(strFile))
        Case "txt": strText = ExportGuideTXT()
        Case "csv": strText = ExportGuideCSV()
    End Select
    If xp.File.SaveStringAs(strFile, strText) Then xp.File.Run strFile
End Sub

Public Function ExportGuideCSV() As String
    Dim strLine() As String
    Dim lngIndex As Long
    Dim lngLevel As Long
    Dim i As Long
    
    ReDim strLine(Guide.Enhancements + build.MaxLevels)
    strLine(0) = Quotes("Level") & "," & Quotes("Tree") & "," & Quotes("Tier") & "," & Quotes("Enhancement") & "," & Quotes("AP") & "," & Quotes("Prog") & "," & Quotes("Order")
    lngLevel = 1
    For i = 1 To Guide.Enhancements
        lngIndex = lngIndex + 1
        With Guide.Enhancement(i)
            If .Level > lngLevel Then
                strLine(lngIndex) = ",,,,,," & lngIndex
                lngLevel = .Level
                lngIndex = lngIndex + 1
            End If
            strLine(lngIndex) = strLine(lngIndex) & .Level & ","
            If .Style = geEnhancement Then
                With Guide.Tree(.GuideTreeID)
                    strLine(lngIndex) = strLine(lngIndex) & Quotes(.Display) & ","
                End With
                If .Tier = 0 Then
                    strLine(lngIndex) = strLine(lngIndex) & Quotes("Core " & .Ability) & ","
                Else
                    strLine(lngIndex) = strLine(lngIndex) & Quotes("Tier " & .Tier) & ","
                End If
                strLine(lngIndex) = strLine(lngIndex) & Quotes(.Display & .RankText) & "," & .Cost & "," & .SpentInTree & ","
            Else
                strLine(lngIndex) = strLine(lngIndex) & ",," & Quotes(.Display) & ",,,"
            End If
        End With
        strLine(lngIndex) = strLine(lngIndex) & lngIndex
    Next
    ReDim Preserve strLine(lngIndex)
    ExportGuideCSV = Join(strLine, vbNewLine)
End Function

Private Function Quotes(pstrText As String) As String
    Quotes = """" & pstrText & """"
End Function

Public Function ExportGuideTXT() As String
    Const Margin As String = "  "
    Dim lngTreeWidth As Long
    Dim lngEnhancementWidth As Long
    Dim strLine() As String
    Dim lngIndex As Long
    Dim lngLevel As Long
    Dim strCol() As String
    Dim i As Long
    
    ReDim strLine(Guide.Enhancements + build.MaxLevels + 1)
    ' Get widths
    lngTreeWidth = Len("Tree")
    lngEnhancementWidth = Len("Enhancement")
    For i = 1 To Guide.Enhancements
        With Guide.Enhancement(i)
            If lngEnhancementWidth < Len(.Display & .RankText) Then lngEnhancementWidth = Len(.Display & .RankText)
            If .Style = geEnhancement Then
                With Guide.Tree(.GuideTreeID)
                    If lngTreeWidth < Len(.Display) Then lngTreeWidth = Len(.Display)
                End With
            End If
        End With
    Next
    ReDim strCol(5)
    strCol(0) = "Level"
    strCol(1) = AlignText("Tree", lngTreeWidth, vbCenter)
    strCol(2) = " Tier "
    strCol(3) = AlignText("Enhancement", lngEnhancementWidth, vbCenter)
    strCol(4) = "AP"
    strCol(5) = "Prog"
    strLine(0) = Join(strCol, Margin)
    InitCol strCol, "-", lngTreeWidth, lngEnhancementWidth
    strLine(1) = Join(strCol, Margin)
    lngLevel = 1
    lngIndex = 1
    For i = 1 To Guide.Enhancements
        InitCol strCol, " ", lngTreeWidth, lngEnhancementWidth
        lngIndex = lngIndex + 1
        With Guide.Enhancement(i)
            If .Level > lngLevel Then
                strLine(lngIndex) = Join(strCol, Margin)
                lngLevel = .Level
                lngIndex = lngIndex + 1
            End If
            strCol(0) = AlignText(CStr(.Level), 4, vbRightJustify) & " "
            If .Style = geEnhancement Then
                With Guide.Tree(.GuideTreeID)
                    strCol(1) = AlignText(.Display, lngTreeWidth, vbLeftJustify)
                End With
                If .Tier = 0 Then
                    strCol(2) = "Core " & .Ability
                Else
                    strCol(2) = "Tier " & .Tier
                End If
                strCol(3) = AlignText(.Display & .RankText, lngEnhancementWidth, vbLeftJustify)
                strCol(4) = " " & .Cost
                strCol(5) = AlignText("  " & .SpentInTree, 3, vbRightJustify) & " "
            Else
                strCol(3) = AlignText(.Display, lngEnhancementWidth, vbLeftJustify)
            End If
        End With
        strLine(lngIndex) = Join(strCol, Margin)
    Next
    ReDim Preserve strLine(lngIndex)
    ExportGuideTXT = Join(strLine, vbNewLine)
End Function

Private Sub InitCol(plngCol() As String, pstrChar As String, plngTreeWidth As Long, plngEnhancementWidth As Long)
    Dim bytChar As Byte
    
    bytChar = Asc(pstrChar)
    plngCol(0) = String(5, bytChar)
    plngCol(1) = String(plngTreeWidth, bytChar)
    plngCol(2) = String(6, bytChar)
    plngCol(3) = String(plngEnhancementWidth, bytChar)
    plngCol(4) = String(2, bytChar)
    plngCol(5) = String(4, bytChar)
End Sub


' ************* FILE I/O *************


Public Function LoadBuild(pstrFile As String, pblnMRU As Boolean, Optional pblnSilent As Boolean = False) As LoadErrorEnum
On Error GoTo LoadBuildError
    Dim lngSignature As Long
    Dim bytVersion As Byte
    Dim FileNumber As Long
    Dim typVersion1 As BuildType1
    Dim typVersion2 As BuildType2
    Dim typVersion3 As BuildType3
    Dim typVersion4 As BuildType4
    Dim typVersion5 As BuildType5
    
    CloseForm "frmDeprecate"
    If Not xp.File.Exists(pstrFile) Then
        LoadBuild = leeFileNotFound
        If Not pblnSilent Then Notice "File not found:" & vbNewLine & vbNewLine & pstrFile
        GoTo LoadBuildClose
    Else
        LoadBuild = leeNoError
    End If
    Select Case LCase$(GetExtFromFilespec(pstrFile))
        ' Text Build
        Case "build"
            DeprecateInit
            If Not LoadFileLite(pstrFile) Then
                LoadBuild = leeUnexpectedError
                GoTo LoadBuildClose
            End If
        ' Binary Build
        Case "bld"
            ' Physically read the file
            FileNumber = FreeFile
            Open pstrFile For Binary As #FileNumber
            ' Check first 4 bytes for valid signature
            Get #FileNumber, 1, lngSignature
            If lngSignature = Signature Then
                ' Get version from next  byte
                Get #FileNumber, 5, bytVersion
                ' Load Data
                Select Case bytVersion
                    Case 1
                        Get #FileNumber, 6, typVersion1
                        Version1To2 typVersion1, typVersion2
                        Version2To3 typVersion2, typVersion3
                        Version3To4 typVersion3, typVersion4
                        Version4To5 typVersion4, typVersion5
                        Version5To6 typVersion5, build
                    Case 2
                        Get #FileNumber, 6, typVersion2
                        Version2To3 typVersion2, typVersion3
                        Version3To4 typVersion3, typVersion4
                        Version4To5 typVersion4, typVersion5
                        Version5To6 typVersion5, build
                    Case 3
                        Get #FileNumber, 6, typVersion3
                        Version3To4 typVersion3, typVersion4
                        Version4To5 typVersion4, typVersion5
                        Version5To6 typVersion5, build
                    Case 4
                        Get #FileNumber, 6, typVersion3
                        Version4To5 typVersion4, typVersion5
                        Version5To6 typVersion5, build
                    Case 5
                        Version5To6 typVersion5, build
                    Case 6
                        Get #FileNumber, 6, build
                    Case Else
                        LoadBuild = leeUnsupported
                        If Not pblnSilent Then Notice "Unsupported build version"
                        GoTo LoadBuildClose
                End Select
                DeprecateBinary
            Else
                LoadBuild = leeUnrecognized
                If Not pblnSilent Then Notice "Build file not recognized."
                GoTo LoadBuildClose
            End If
        Case Else
            LoadBuild = leeUnrecognized
            If Not pblnSilent Then Notice "Build file not recognized."
            GoTo LoadBuildClose
    End Select
    Close
    If pblnSilent Then
        InitBuildFeats
        InitBuildSpells
        InitBuildTrees
        InitLevelingGuide
    Else
        mblnOpen = True
        If pblnMRU Then cfg.AddMRU pstrFile
        EnableMenus
        SetAppCaption
        CascadeChanges cceAll
        frmMain.tmrDeprecate.Enabled = gtypDeprecate.Deprecated
        SetDirty False
    End If
    Exit Function
    
LoadBuildClose:
    Close
    BuildClose
    Exit Function
    
LoadBuildError:
    If pblnSilent Then
        LoadBuild = Err.Number
    Else
        MsgBox Err.Description, vbExclamation, "Error #" & Err.Number
    End If
    Resume LoadBuildClose
End Function

Private Sub SaveBuild(pstrFile As String)
    Dim FileNumber As Long
    
On Error GoTo DeleteBuildErr
    If xp.File.Exists(pstrFile) Then xp.File.Delete pstrFile

On Error GoTo SaveBuildErr
    Select Case LCase$(GetExtFromFilespec(pstrFile))
        Case "build"
            SaveFileLite pstrFile
        Case "bld"
            ' Save file
            FileNumber = FreeFile()
            Open pstrFile For Binary As #FileNumber
            Put #FileNumber, 1, Signature
            Put #FileNumber, 5, Version
            Put #FileNumber, 6, build
            Close
    End Select
    
SaveBuildExit:
    SetAppCaption
    Exit Sub
    
SaveBuildErr:
    MsgBox "Error saving backup build:" & vbNewLine & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Filename: " & pstrFile, vbInformation, "Error #" & Err.Number
    Close
    Resume SaveBuildExit
    
DeleteBuildErr:
    MsgBox "Error deleting previous backup build:" & vbNewLine & vbNewLine & Err.Description & vbNewLine & vbNewLine & "Filename: " & pstrFile, vbInformation, "Error #" & Err.Number
    Resume SaveBuildExit
End Sub
