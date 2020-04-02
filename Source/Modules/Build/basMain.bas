Attribute VB_Name = "basMain"
' Written by Ellis Dee
' Program entrypoint, global classes, window functions, developer tools, test functions
Option Explicit

Public Const MaxLevel As Long = 30

Public Const ErrorIgnore As Long = 37001

Public glngActiveColor As Long
Public glngLevel As Long ' Exchange feat at level
Public gstrError As String ' Error description to display in Details window (all screens)

Public gstrCommand As String

Public PixelX As Long
Public PixelY As Long

Public xp As clsWindowsXP
Public cfg As clsConfig


' ************* PROGRAM *************


Sub Main()
    Load frmMessages
    gstrCommand = Command
    Randomize
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    Set xp = New clsWindowsXP
    Set cfg = New clsConfig
    frmMain.Show
End Sub

Public Sub CloseApp()
    Dim frm As Form
    
    For Each frm In Forms
        Unload frm
    Next
    Set cfg = Nothing
    Set xp = Nothing
End Sub


' ************* CLEANUP *************


Public Sub DeleteOldFiles()
    Dim strVersion As String
    Dim lngCount As Long
    
    strVersion = App.Major & "." & App.Minor & "." & App.Revision
    If cfg.Version <> strVersion Then
        lngCount = lngCount + DeleteFolder("Images")
        lngCount = lngCount + DeleteFile("Data\Classes.txt", True)
        lngCount = lngCount + DeleteFile("Data\Destinies.txt", True)
        lngCount = lngCount + DeleteFile("Data\Enhancements.txt", True)
        lngCount = lngCount + DeleteFile("Data\Exceptions.txt", True)
        lngCount = lngCount + DeleteFile("Data\Feats.txt", True)
        lngCount = lngCount + DeleteFile("Data\Help.txt", True)
        lngCount = lngCount + DeleteFile("Data\Races.txt", True)
        lngCount = lngCount + DeleteFile("Data\ReadMe.txt", True)
        lngCount = lngCount + DeleteFile("Data\Spells.txt", True)
        lngCount = lngCount + DeleteFile("Kensei Warpriest.bld")
        lngCount = lngCount + DeleteFile("Default.bbcodes")
        lngCount = lngCount + DeleteFile("CharacterBuilderLite.cfg", True)
        lngCount = lngCount + DeleteFileMask("Save", "*.clr")
        lngCount = lngCount + DeleteFileMask("Save", "*.out")
        lngCount = lngCount + DeleteFileMask("Settings", "*.clr")
        lngCount = lngCount + DeleteFileMask("Settings", "*.out")
        lngCount = lngCount + DeleteFileMask("Save", "*.log")
        If lngCount = 0 Then cfg.Version = strVersion
    End If
End Sub

Private Function DeleteFile(pstrFile As String, Optional pblnRoot As Boolean = False) As Long
    Dim lngReturn As Long
    
    If pblnRoot Then
        lngReturn = lngReturn + Delete(App.Path & "\" & pstrFile)
    Else
        lngReturn = lngReturn + Delete(App.Path & "\Settings\" & pstrFile)
        lngReturn = lngReturn + Delete(App.Path & "\Save\" & pstrFile)
    End If
    DeleteFile = lngReturn
End Function

Private Function DeleteFileMask(pstrRelativePath As String, pstrMask As String) As Long
On Error Resume Next
    Dim strFile As String
    Dim strFolder As String
    Dim lngReturn As Long
    
    strFolder = App.Path & "\" & pstrRelativePath & "\"
    strFile = Dir(strFolder & pstrMask)
    Do While Len(strFile)
        Kill strFolder & strFile
        lngReturn = lngReturn + 1
        strFile = Dir
    Loop
    DeleteFileMask = lngReturn
End Function

Private Function Delete(pstrFile As String) As Long
    If Not xp.File.Exists(pstrFile) Then Exit Function
    Delete = 1
    On Error Resume Next
    Kill pstrFile
End Function

Private Function DeleteFolder(pstrRelative As String) As Long
    Dim strFolder As String
    
    ' Bitmaps (now wrapped into exe via resource file)
    strFolder = App.Path & "\" & pstrRelative & "\"
    If Not xp.Folder.Exists(strFolder) Then Exit Function
    DeleteFolder = 1
    xp.Folder.Delete strFolder
End Function


' ************* WINDOWS *************


' All forms are opened with this helper function, which returns TRUE on success
Public Function OpenForm(pstrForm As String) As Boolean
    Dim frm As Form
    
    ' Valid Race?
    If InStr("frmStats.frmSkills.frmFeats.frmEnhancements", pstrForm) Then
        If build.Race = reAny Then
            If Screen.ActiveForm.Name = "frmOverview" Then
                Notice "Select a Race first."
                frmOverview.RaceShow True
            Else
                Notice "Select a Race in the Overview screen first."
            End If
            Exit Function
        End If
    End If
    ' Valid Class?
    If InStr("frmStats.frmSkills.frmFeats.frmSpells.frmEnhancements", pstrForm) Then
        If build.Class(1) = ceAny Then
            If Screen.ActiveForm.Name = "frmOverview" Then
                Notice "Select a class split first."
            Else
                Notice "Select a class split in the Overview screen first."
            End If
            Exit Function
        End If
    End If
    ' Valid Stats?
    If InStr("frmSkills.frmFeats", pstrForm) Then
        If build.StatPoints(build.BuildPoints, 0) <> GetBuildPoints(build.BuildPoints) Then
            If Screen.ActiveForm.Name = "frmStats" Then
                Notice "Allocate base stats first."
            Else
                Notice "Allocate base stats in the Stats screen first."
            End If
            Exit Function
        End If
    End If
    ' Can cast spells?
    If pstrForm = "frmSpells" Then
        If build.CanCastSpell(1) = 0 Then
            Notice "This build cannot cast spells."
            Exit Function
        End If
    End If
    ' Epic levels?
    If pstrForm = "frmDestiny" Then
        If build.MaxLevels < 20 Then
            Notice "This build never reaches level 20."
            Exit Function
        End If
    End If
    ' Everything checks out
    Select Case pstrForm
        Case "frmOverview": Set frm = frmOverview
        Case "frmStats": Set frm = frmStats
        Case "frmSkills": Set frm = frmSkills
        Case "frmFeats": Set frm = frmFeats
        Case "frmSpells": Set frm = frmSpells
        Case "frmEnhancements": Set frm = frmEnhancements
        Case "frmDestiny": Set frm = frmDestiny
'        Case "frmGear": Set frm = frmGear
        Case "frmImport": Set frm = frmImport
        Case "frmExport": Set frm = frmExport
        Case "frmConvert": Set frm = frmConvert
        Case "frmOptions": Set frm = frmOptions
        Case "frmDeprecate": Set frm = frmDeprecate
'        Case "frmViewData": Set frm = frmViewData
        Case "frmRefreshData": Set frm = frmRefreshData
        Case "frmHelp": Set frm = frmHelp
        Case "frmFormat": Set frm = frmFormat
        Case Else: Exit Function
    End Select
    If cfg.ChildWindows Then frm.Show vbModeless, frmMain Else frm.Show
    Set frm = Nothing
    frmMain.UpdateWindowMenu
    OpenForm = True
End Function

' This is called when form A needs to remotely close form B
Public Sub CloseForm(pstrForm As String)
    Dim frm As Form
    
    For Each frm In Forms
        If InStr("frmMain.frmOptions.frmFormat.frmMessages.frmConvert", frm.Name) = 0 And pstrForm = "All" Then
            Unload frm
        ElseIf InStr("frmOverview.frmStats.frmSkills.frmFeats.frmSpells.frmEnhancements.frmDestiny.frmGear.frmDeprecate", frm.Name) Then
            If frm.Name = pstrForm Then Unload frm
        End If
    Next
    frmMain.UpdateWindowMenu
End Sub

' Called by all forms when closed by the user
Public Sub UnloadForm(pfrm As Form, pblnOverride As Boolean)
    Dim frm As Form
    Dim blnSkip As Boolean
    
    cfg.SavePosition pfrm
    SaveBackup
    For Each frm In Forms
        If frm.Name <> pfrm.Name And InStr("frmOverview.frmStats.frmSkills.frmFeats.frmSpells.frmEnhancements.frmDestiny.frmGear", frm.Name) <> 0 Then
            blnSkip = True
            Exit For
        End If
    Next
    If pblnOverride Then
        pblnOverride = False
    ElseIf Not blnSkip Then
        frmMain.tmrOutput.Enabled = True
    End If
    frmMain.UpdateWindowMenu
End Sub

Public Sub ActivateForm(Optional penOutput As OutputEnum = oeNone)
    frmMain.WindowActivate
    If penOutput <> oeNone Then GenerateOutput penOutput
End Sub

' Update all open forms to reflect data changes from other forms
Public Sub CascadeChanges(penChange As CascadeChangeEnum)
    Dim frm As Form
    
    Select Case penChange
        Case cceAll, cceMaxLevel
            If build.Class(1) <> ceAny Then
                CalculateBAB
                InitBuildSkills
                InitBuildSpells
                InitBuildFeats
                InitLevelingGuide
            End If
            InitBuildTrees
        Case cceClass
            If build.Class(1) <> ceAny Then
                InitBuildSkills
                InitBuildSpells
                InitBuildFeats
                InitLevelingGuide
            End If
            InitBuildTrees
        Case cceRace
            If build.Class(1) <> ceAny Then
                InitBuildSkills
                InitBuildFeats
            End If
            InitBuildTrees
            InitLevelingGuide
        Case cceAlignment
            If build.Class(1) <> ceAny Then CheckSlotErrors
        Case cceStats
            If build.Class(1) <> ceAny Then
                InitBuildSkills
                CheckSlotErrors
            End If
        Case cceSkill
            If build.Class(1) <> ceAny Then CheckSlotErrors
        Case cceFeat
            If build.Class(1) <> ceAny Then InitBuildSpells
            InitLevelingGuide
        Case cceEnhancements
            InitLevelingGuide
    End Select
    For Each frm In Forms
        Select Case frm.Name
            Case "frmOverview"
                If penChange = cceAll Then frm.Cascade
            Case "frmStats"
                Select Case penChange
                    Case cceAll, cceClass, cceRace, cceMaxLevel: If build.Race = reAny Then CloseForm "frmStats" Else frm.Cascade
                End Select
            Case "frmSkills"
                Select Case penChange
                    Case cceAll, cceRace, cceClass, cceStats, cceMaxLevel: If build.Race = reAny Or build.Class(1) = ceAny Then CloseForm "frmSkills" Else frm.Cascade
                End Select
            Case "frmFeats"
                If penChange <> cceFeat Then
                    If build.Class(1) = ceAny Then CloseForm "frmFeats" Else frm.Cascade
                End If
            Case "frmSpells"
                If build.CanCastSpell(1) = 0 Then CloseForm "frmSpells" Else frm.Cascade
            Case "frmEnhancements"
                Select Case penChange
                    Case cceAll, cceRace, cceClass, cceMaxLevel, cceFeat, cceEnhancements: If build.Race = reAny Or build.Class(1) = ceAny Then CloseForm "frmEnhancements" Else frm.Cascade
                End Select
            Case "frmDestiny"
                If build.MaxLevels < 20 Then CloseForm "frmDestiny" Else frm.Cascade
        End Select
    Next
End Sub


' ************* DEVELOPER TOOLS *************


' Generate source code for setting initial default colors
Public Sub DefaultColorCode()
    Dim strColor() As String
    Dim lngColor As Long
    Dim lngGroup As Long
    Dim strText As String
    Dim strFile As String
    
    strColor = Split("cveBackground,cveBackHighlight,cveBackError,cveBackRelated,cveText,cveTextError,cveTextDim,cveTextLink,cveBorderInterior,cveBorderExterior,cveBorderHighlight,cveLightGray,cveRed,cveYellow,cveBlue,cvePurple,cveGreen,cveOrange", ",")
    For lngColor = 0 To 17
        strText = strText & "    InitColors " & strColor(lngColor)
        For lngGroup = 0 To 4
            strText = strText & ", " & cfg.GetColor(lngGroup, lngColor)
        Next
        strText = strText & vbNewLine
    Next
    strFile = App.Path & "\Temp.txt"
    xp.File.SaveStringAs strFile, strText
    xp.File.Run strFile
End Sub

Public Sub CleanTrees()
    Dim i As Long
    
    For i = 1 To db.Trees
        CleanTree db.Tree(i)
    Next
    For i = 1 To db.Destinies
        CleanTree db.Destiny(i)
    Next
End Sub

Private Sub CleanTree(ptypTree As TreeType)
    Dim lngTier As Long
    Dim i As Long
    
    For lngTier = 1 To ptypTree.Tiers
        With ptypTree.Tier(lngTier)
            For i = 1 To .Abilities
                With .Ability(i)
                    Do While InStr(.Descrip, " }")
                        .Descrip = Replace(.Descrip, " }", "}")
                    Loop
                    Do While InStr(.Descrip, "}}")
                        .Descrip = Replace(.Descrip, "}}", "}")
                    Loop
                    Do While InStr(.Descrip, "{}")
                        .Descrip = Replace(.Descrip, "{}", "")
                    Loop
                    ' Final cleanup
                    Do
                        .Descrip = Trim$(.Descrip)
                        If Right$(.Descrip, 1) = "}" Then .Descrip = Left$(.Descrip, Len(.Descrip) - 1) Else Exit Do
                    Loop
                End With
            Next
        End With
    Next
End Sub

Public Sub FindPrereqsWithFewerRanks()
    Dim i As Long
    
    For i = 1 To db.Trees
        FindPrereqsWithFewerRanksTree db.Tree(i)
    Next
    For i = 1 To db.Destinies
        FindPrereqsWithFewerRanksTree db.Destiny(i)
    Next
End Sub

Private Sub FindPrereqsWithFewerRanksTree(ptypTree As TreeType)
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim enReq As ReqGroupEnum
    Dim lngReq As Long
    Dim lngPrereqRanks As Long
    Dim lngRank As Long
    
    For lngTier = 0 To ptypTree.Tiers
        For lngAbility = 1 To ptypTree.Tier(lngTier).Abilities
            With ptypTree.Tier(lngTier).Ability(lngAbility)
                For enReq = rgeAll To rgeNone
                    For lngReq = 1 To .Req(enReq).Reqs
                        If .Req(enReq).Req(lngReq).Style <> peFeat Then
                            lngPrereqRanks = ReqMaxRanks(.Req(enReq).Req(lngReq))
                            If lngPrereqRanks < .Ranks Then Debug.Print ptypTree.TreeName & " Tier " & lngTier & ": " & .AbilityName & " " & .Ranks & " requires " & PointerDisplay(.Req(enReq).Req(lngReq), True)
                        End If
                    Next
                Next
                If .RankReqs Then
                    Debug.Print "Rank reqs: " & ptypTree.TreeName & " Tier " & lngTier & ": " & .AbilityName & " (???)"
                End If
            End With
        Next
    Next
End Sub

Private Function ReqMaxRanks(ptypReq As PointerType) As Long
    With ptypReq
        Select Case .Style
            Case peDestiny
                ReqMaxRanks = db.Destiny(.Tree).Tier(.Tier).Ability(.Ability).Ranks
            Case peEnhancement
                ReqMaxRanks = db.Tree(.Tree).Tier(.Tier).Ability(.Ability).Ranks
        End Select
    End With
End Function

Public Sub TreeAbbreviations()
    Dim i As Long
    
    For i = 1 To db.Trees
        Debug.Print db.Tree(i).Abbreviation
    Next
End Sub


' ************* PRINTABLE DATA *************


Public Function PrintEnhancements()
    Dim lngTree As Long
    Dim strOutput As String
    Dim strFile As String
    
    strFile = App.Path & "\Print.txt"
    For lngTree = 1 To db.Trees
        strOutput = strOutput & PrintTree(db.Tree(lngTree))
    Next
    xp.File.SaveStringAs strFile, strOutput
    xp.File.Run strFile
End Function

Public Function PrintTree(ptypTree As TreeType) As String
    Dim strReturn As String
    Dim lngTier As Long
    Dim lngAbility As Long
    Dim lngFirst As Long
    Dim strLine As String
    
    With ptypTree
        For lngTier = 0 To .Tiers
            With .Tier(lngTier)
                For lngAbility = 1 To .Abilities
                    With .Ability(lngAbility)
                        strLine = ptypTree.TreeName & GetTier(lngTier, lngAbility, ptypTree.TreeType)
                        strReturn = strReturn & .AbilityName & " (" & .Cost & " AP"
                        If .Ranks > 1 Then strReturn = strReturn & " per rank, " & .Ranks & " ranks"
                        strReturn = strReturn & ")" & vbNewLine
                        strReturn = strReturn & .Descrip
                        strReturn = strReturn & vbNewLine & vbNewLine
                    End With
                Next
            End With
        Next
    End With
    PrintTree = strReturn & vbNewLine
End Function
