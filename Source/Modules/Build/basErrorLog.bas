Attribute VB_Name = "basErrorLog"
' Written by Ellis Dee
' I would prefer this be a class module, but the "ptr" UDT doesn't play nice with property get/let
' On the plus side, not being a class makes the code faster
Option Explicit

Public Enum ActivityEnum
    actOpenFile
    actReadFile
    actFindRace
    actFindClass
    actLoadSelector
    actLoadFeatMap
    actLoadTree
    actProcessRacialTrees
    actProcessClassTrees
    actProcessFeatSelectors
    actProcessEnhancementSelectors
    actProcessDestinySelectors
    actProcessRaceGrantedFeats
    actProcessClassGrantedFeats
    actProcessFeatReqs
    actProcessFeatMap
    actProcessEnhancementReqs
    actProcessDestinyReqs
    actProcessClassSpells
    actProcessTreeLockouts
    actTemplateStats
    actTemplatePoints
    actDeprecateFeat
    actDeprecateFeatSelector
    actDeprecateEnhancement
    actDeprecateLevelingGuide
    actDeprecateDestiny
    actDeprecateTwist
End Enum

Public Type LogType
    Activity As ActivityEnum
    Style As PointerEnum
    Raw As String
    Race As Long
    Class As Long
    Feat As Long
    Level As Long
    Tree As Long
    Tier As Long
    Ability As Long
    ptr As PointerType
    ReqGroup As ReqGroupEnum
    Req As Long
    HasError As Boolean
    Errors As Long
    Selector As Long
    Rank As Long
    Template As Long
    Stat As StatEnum
    Points As Long
    Total As Long
    LineNumber As Long
    ErrorDetail() As String
    ErrorDetails As Long
    LoadFile As String
    LoadItem As String
    LoadLine As String
    LoadSelector As String
    LoadTree As String
    LoadSpell As String
    LoadSpellType As String
End Type

Public log As LogType

Public Sub ClearLog()
    Dim typBlank As LogType
    
    log = typBlank
End Sub

Public Sub LoadError(pstrError As String)
    With log
        .HasError = True
        .Errors = .Errors + 1
        AddError log.LoadFile & ": " & pstrError
        AddError vbNullString
    End With
End Sub

Public Sub LogError()
    Dim strError As String
    
    With log
        .HasError = True
        .Errors = .Errors + 1
        Select Case .Activity
            ' General
            Case actOpenFile
                AddError "File not found: " & log.LoadFile
                AddError DataPath & log.LoadFile
            Case actReadFile
                AddError log.LoadFile & ": " & log.LoadItem & " has invalid line:"
                AddError log.LoadLine
            ' Races
            Case actFindRace
                AddError "Race not recognized: " & log.LoadItem
            Case actProcessRacialTrees
                AddError "Race: " & db.Race(log.Race).RaceName
                If log.Tree = 0 Then AddError "Racial tree not found" Else AddError "Invalid RaceClass tree: " & db.Race(log.Race).Tree(log.Tree)
            Case actProcessRaceGrantedFeats
                AddError "Race: " & db.Race(log.Race).RaceName
                AddError "Invalid granted feat (level " & .Level & "): " & db.Race(log.Race).GrantedFeat(.Feat).Raw
            ' Classes
            Case actFindClass
                AddError "Class not recognized: " & log.LoadItem
            Case actProcessClassSpells
                If log.Level = 0 Then strError = log.LoadSpellType Else strError = "level " & log.Level & " spell"
                AddError log.LoadFile & ": " & db.Class(log.Class).ClassName & " has invalid " & strError & ":"
                AddError log.LoadSpell
            Case actProcessClassGrantedFeats
                AddError "Class: " & db.Class(log.Class).ClassName
                AddError "Invalid granted feat (level " & .Level & "): " & db.Class(log.Class).GrantedFeat(.Feat).Raw
            Case actProcessClassTrees
                AddError "Class: " & db.Class(log.Class).ClassName
                AddError "Invalid class tree: " & db.Class(log.Class).Tree(log.Tree)
            ' Feats
            Case actLoadSelector
                AddError log.LoadFile & ": " & log.LoadItem & " SelectorName: " & log.LoadSelector & " has invalid line:"
                AddError log.LoadLine
            Case actProcessFeatSelectors
                AddError "Feat: " & db.Feat(log.Feat).FeatName
                AddError "Invalid parent selector: " & db.Feat(log.Feat).Parent.Raw
            Case actProcessFeatReqs
                If log.Selector = 0 Then
                    AddError "Feat: " & db.Feat(log.Feat).FeatName
                    AddError "Invalid requirement: " & .ptr.Raw
                Else
                    AddError "Feat: " & db.Feat(log.Feat).FeatName & ": " & db.Feat(log.Feat).Selector(log.Selector).SelectorName
                    AddError "Invalid requirement: " & .ptr.Raw
                End If
            ' FeatMap
            Case actLoadFeatMap
                If log.LineNumber = 0 Then
                    AddError log.LoadFile & ": File is empty"
                Else
                    AddError log.LoadFile & ": Line " & log.LineNumber + 1 & " is invalid:"
                    AddError log.LoadLine
                End If
            Case actProcessFeatMap
                If log.Feat = 0 Then
                    AddError log.LoadFile & ": Feat not found: " & log.LoadItem
                Else
                    AddError log.LoadFile & ": " & log.LoadItem & " selector not found: " & log.LoadSelector
                End If
            ' Templates
            Case actTemplateStats
                AddError log.LoadFile & ": " & LogTemplateName() & " has invalid stat point value"
                AddError 28 + (log.Points * 2) & "pt " & GetStatName(log.Stat) & ": " & db.Template(log.Template).StatPoints(log.Points, log.Stat)
            Case actTemplatePoints
                AddError log.LoadFile & ": " & LogTemplateName() & " has invalid points total"
                AddError 28 + (log.Points * 2) & "pt build totals " & log.Total
            ' Trees
            Case actLoadTree
                strError = log.LoadFile & ": " & log.LoadTree & " "
                If log.Tier = -1 Then
                    strError = strError & "header"
                Else
                    strError = strError & "Tier " & log.Tier & ": " & log.LoadItem
                End If
                strError = strError & " has invalid line:"
                AddError strError
                AddError log.LoadLine
            Case actProcessTreeLockouts
                AddError "Tree: " & db.Tree(log.Tree).TreeName
                AddError "Invalid lockout tree: " & db.Tree(log.Tree).Lockout
            ' Enhancements
            Case actProcessEnhancementSelectors
                AddError "Enhancement: " & TreeError(db.Tree(.Tree))
                AddError "Invalid parent selector: " & .ptr.Raw
            Case actProcessEnhancementReqs
                AddError "Enhancement: " & TreeError(db.Tree(.Tree))
                AddError "Invalid requirement: " & .ptr.Raw
            ' Destinies
            Case actProcessDestinySelectors
                AddError "Destiny: " & TreeError(db.Destiny(.Tree))
                AddError "Invalid parent selector: " & .ptr.Raw
            Case actProcessDestinyReqs
                AddError "Destiny: " & TreeError(db.Destiny(.Tree))
                AddError "Invalid requirement: " & .ptr.Raw
            ' Deprecate
            Case actDeprecateFeat
                AddError "Feat not found: " & log.LoadItem
            Case actDeprecateFeatSelector
                AddError "Invalid selector for Feat: " & log.LoadItem
            Case actDeprecateEnhancement
                AddError "Tree not found or tree has changed: " & log.LoadTree
            Case actDeprecateLevelingGuide
                AddError "Leveling Guide has been reset"
            Case actDeprecateDestiny
                AddError "Destiny not found or destiny has changed: " & log.LoadTree
            Case actDeprecateTwist
                strError = "Twist is invalid: " & log.LoadTree & " Tier " & log.Tier & ", Ability " & log.Ability
                If log.Selector <> 0 Then strError = strError & ", Selector " & log.Selector
                AddError strError
            Case Else
                AddError "Unknown error"
        End Select
    End With
    AddError vbNullString
End Sub

Private Function LogTemplateName() As String
    Dim strReturn As String
    
    With db.Template(log.Template)
        strReturn = db.Class(.Class).ClassName & ": " & .Caption
        If .Trapping Then strReturn = strReturn & " (Traps)"
    End With
    LogTemplateName = strReturn
End Function

Public Sub ErrorLoading(pstrRaw As String)
    Dim strError As String
    Dim strLine() As String
    
    With log
        .Errors = .Errors + 1
        strError = "Invalid section in " & .LoadFile
        Select Case .LoadFile
            Case "Races.txt": strError = strError & " ('Stats: ' not found):"
            Case "Classes.txt": strError = strError & " ('BAB: ' not found):"
            Case "Feats.txt": strError = strError & " ('Group: ' not found):"
            Case "Enhancements.txt", "Destinies.txt": strError = strError & " ('Type: ' not found):"
        End Select
    End With
    AddError strError
    strLine = Split(pstrRaw, vbNewLine)
    AddError strLine(0)
    AddError vbNullString
End Sub

Public Function TreeError(ptypTree As TreeType) As String
    Dim strError As String
    
    If log.Tree Then
        strError = ptypTree.TreeName & " "
        strError = strError & "Tier " & log.Tier & ": "
        If log.Ability Then strError = strError & ptypTree.Tier(log.Tier).Ability(log.Ability).AbilityName & " "
    End If
    TreeError = strError
End Function

Public Sub AddError(pstrError As String)
    With log
        .ErrorDetails = .ErrorDetails + 1
        ReDim Preserve .ErrorDetail(1 To .ErrorDetails)
        .ErrorDetail(.ErrorDetails) = pstrError
    End With
End Sub

Public Function CreateErrorLog(Optional pblnAlwaysShow As Boolean = False) As Boolean
    Dim strFile As String
    Dim strOutput As String
    
    strFile = ErrorLogFile()
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    If log.Errors Then
        strOutput = log.Errors & " Errors:" & vbCrLf & vbCrLf & Join(log.ErrorDetail, vbCrLf)
        xp.File.SaveStringAs strFile, strOutput
        CreateErrorLog = True
        If cfg.ShowErrors Or pblnAlwaysShow Then ViewErrorLog False
    End If
End Function

Public Function ErrorLogText() As String
    If log.Errors Then ErrorLogText = log.Errors & " Errors:" & vbCrLf & vbCrLf & Join(log.ErrorDetail, vbCrLf)
End Function

Public Sub ViewErrorLog(pblnNotifyWhenMissing As Boolean)
    Dim strFile As String
    
    strFile = ErrorLogFile()
    If xp.File.Exists(strFile) Then
        xp.File.Run strFile
    ElseIf pblnNotifyWhenMissing Then
        Notice "No errors at startup."
    End If
End Sub

Private Function ErrorLogFile() As String
    ErrorLogFile = App.Path & "\Error.log"
End Function

Public Sub ViewExceptions()
    Dim strFile As String
    
    strFile = DataPath() & "Exceptions.txt"
    If xp.File.Exists(strFile) Then
        xp.File.Run strFile
    Else
        Notice "No exceptions found."
    End If
End Sub
