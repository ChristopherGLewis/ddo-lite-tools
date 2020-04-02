VERSION 5.00
Begin VB.UserControl userCharacter 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2412
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4392
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   2412
   ScaleWidth      =   4392
   Begin VB.PictureBox picCharXP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1032
      Left            =   408
      ScaleHeight     =   1032
      ScaleWidth      =   2352
      TabIndex        =   1
      Top             =   1296
      Width           =   2352
   End
   Begin Compendium.userSpinner usrspnChar 
      Height          =   372
      Left            =   1212
      TabIndex        =   0
      Tag             =   "ctl"
      Top             =   852
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   656
      Max             =   30
      ShowZero        =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin Compendium.userDropdown udrpColor 
      Height          =   372
      Left            =   1212
      TabIndex        =   2
      Top             =   432
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   656
      Caption         =   "Color"
   End
   Begin Compendium.userDropdown udrpCharacter 
      Height          =   372
      Left            =   1212
      TabIndex        =   3
      Top             =   12
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   656
      Caption         =   "Character"
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   0
      Left            =   2928
      TabIndex        =   4
      Top             =   12
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   762
      Caption         =   "Epic Sagas"
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   1
      Left            =   2928
      TabIndex        =   5
      Top             =   432
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Heroic Sagas"
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   2
      Left            =   2928
      TabIndex        =   6
      Top             =   852
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Challenges"
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   4
      Left            =   2928
      TabIndex        =   7
      Top             =   2016
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Character..."
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   7
      Left            =   12
      TabIndex        =   8
      Top             =   852
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   656
      Caption         =   "Level"
      Style           =   1
      Appearance      =   3
      Alignment       =   1
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   5
      Left            =   12
      TabIndex        =   9
      Top             =   12
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   656
      Caption         =   "Character"
      Style           =   1
      Appearance      =   3
      Alignment       =   1
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   6
      Left            =   12
      TabIndex        =   10
      Top             =   432
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   656
      Caption         =   "Color"
      Style           =   1
      Appearance      =   3
      Alignment       =   1
   End
   Begin Compendium.userButton usrbtnChar 
      Height          =   372
      Index           =   3
      Left            =   2928
      TabIndex        =   11
      Top             =   1596
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   656
      Caption         =   "Reincarnate"
   End
   Begin VB.Shape shpCharacter 
      Height          =   672
      Left            =   48
      Top             =   1488
      Width           =   312
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Index           =   0
      Begin VB.Menu mnuCharacter 
         Caption         =   "Add New"
         Index           =   0
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Delete"
         Index           =   1
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Rename"
         Index           =   2
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Import"
         Index           =   4
      End
      Begin VB.Menu mnuCharacter 
         Caption         =   "Export"
         Index           =   5
         Begin VB.Menu mnuExport 
            Caption         =   "Character Builder Lite"
            Index           =   0
         End
         Begin VB.Menu mnuExport 
            Caption         =   "Ron's Character Planner"
            Index           =   1
         End
         Begin VB.Menu mnuExport 
            Caption         =   "DDO Builder"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "userCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event CharacterChanged()
Public Event CharacterListChanged()

Private Enum PastLifeEnum
    pleClass
    pleRace
    pleIconic
    pleEpic
    pleTotal
End Enum

Private mblnInit As Boolean
Private mblnOverride As Boolean

Private mlngLife As Long
Private mlngLives(4) As Long


' ************* INITIALIZE *************


Public Sub Init()
    RefreshCharacterList
    InitColorMenu
    RefreshColors
    mblnInit = True
    ShowCharacter
    RedrawXP
End Sub

Private Sub InitColorMenu()
    Dim i As Long
    
    With UserControl.udrpColor
        .ListClear False
        For i = 1 To gceColors - 1
            .AddItem GetColorName(i), i, GetColorValue(i)
        Next
        .SetData gceSky
    End With
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .shpCharacter.Move 0, 0, .ScaleWidth, .ScaleHeight
    End With
End Sub


' ************* PUBLIC *************


Public Sub CharacterChanged()
    ShowCharacter
End Sub

Public Sub RefreshCharacterList()
    Dim strMenu As String
    Dim enColor As GeneratedColorEnum
    Dim i As Long
    
    If cfg.Character > db.Characters Then cfg.Character = 0
    With UserControl.udrpCharacter
        .ListClear False
        .AddItem "Show All", 0, cfg.GetColor(cgeControls, cveBackground)
        For i = 1 To db.Characters
            .AddItem db.Character(i).Character, i, db.Character(i).BackColor
        Next
        .SetData cfg.Character
    End With
    With UserControl.udrpColor
        If cfg.Character = 0 Then .SetData 0 Else .SetData db.Character(cfg.Character).GeneratedColor
    End With
End Sub

Public Sub RefreshColors()
    Dim ctl As Control
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        For Each ctl In .Controls
            Select Case ctl.Name
                Case "usrFavor"
                    ctl.ReDrawControl
                    ctl.Move ctl.Left, ctl.Top, ctl.FitWidth, ctl.FitHeight
                Case "usrtabChar", "usrbtnChar", "udrpCharacter", "udrpColor"
                    ctl.RefreshColors
                Case Else
                    Select Case ctl.Tag
                        Case "ctl": cfg.ApplyColors ctl, cgeControls
                        Case Else: cfg.ApplyColors ctl, cgeWorkspace
                    End Select
            End Select
        Next
        .shpCharacter.BorderColor = cfg.GetColor(cgeControls, cveBorderInterior)
        If mblnInit Then RedrawXP
    End With
End Sub


' ************* CHARACTER *************


Private Sub ShowCharacter()
    Dim strText As String
    Dim lngLevel As Long
    Dim lngColor As Long
    Dim lngColorID As Long
    Dim strDropdown As String
    Dim i As Long
    
    UserControl.usrbtnChar(3).Enabled = (cfg.Character <> 0)
    If cfg.Character = 0 Then
        strText = "Show All"
        lngColor = cfg.GetColor(cgeControls, cveBackground)
        lngLevel = UserControl.usrspnChar.Value
    Else
        With db.Character(cfg.Character)
            strText = .Character
            lngColorID = .GeneratedColor
            lngColor = .BackColor
            lngLevel = .Level
            If lngLevel = 0 Then
                lngLevel = 1
                For i = 1 To db.Quests
                    If lngLevel < db.Quest(i).BaseLevel Then
                        If db.Quest(i).Progress(cfg.Character) <> peNone Then lngLevel = db.Quest(i).BaseLevel
                    End If
                Next
            End If
        End With
    End If
    With UserControl
        mblnOverride = True
        .udrpCharacter.SetData cfg.Character
        .udrpColor.SetData lngColorID
        .usrspnChar.Value = lngLevel
        mblnOverride = False
    End With
    RedrawXP
End Sub

Private Sub udrpCharacter_ListChange(Index As Long, Caption As String, ItemData As Long)
    If mblnOverride Then Exit Sub
    cfg.Character = ItemData
    ShowCharacter
    RaiseEvent CharacterChanged
End Sub

Private Sub udrpColor_ListChange(Index As Long, Caption As String, ItemData As Long)
    If mblnOverride = True Or cfg.Character = 0 Then Exit Sub
    With db.Character(cfg.Character)
        .GeneratedColor = Index
        .BackColor = GetColorValue(.GeneratedColor)
        .DimColor = GetColorValue(.GeneratedColor, True)
        If cfg.Character Then RefreshCharacterList
    End With
    ShowCharacter
    frmCompendium.RedrawQuests
    DirtyFlag dfeData
End Sub


' ************* XP TABLE *************


Private Sub usrspnChar_Change()
    If mblnOverride Then Exit Sub
    If cfg.Character Then
        db.Character(cfg.Character).Level = UserControl.usrspnChar.Value
        DirtyFlag dfeData
    End If
    RedrawXP
End Sub

Private Sub RedrawXP()
    Dim lngRowHeight As Long
    Dim lngLeft As Long
    Dim lngMiddle As Long
    Dim lngRight As Long
    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngTextColor As Long
    Dim lngBackColor As Long
    Dim lngBorderColor As Long
    Dim blnPastLives As Boolean
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngLines As Long
    Dim varArray As Variant
    Dim i As Long
    
    With UserControl
        ' Data
        CountPastLives
        blnPastLives = (.usrspnChar.Value = 30 And cfg.Character > 0 And mlngLife > 1)
        ' Coordinates
        lngRowHeight = .picCharXP.TextHeight("Q") + .picCharXP.ScaleY(2, vbPixels, vbTwips)
        lngLeft = .udrpCharacter.Left - .picCharXP.TextWidth(" Level ")
        lngLeft = .udrpCharacter.Left - .picCharXP.TextWidth(" Level ")
        lngTop = .picCharXP.Top
        lngWidth = .udrpCharacter.Left + .udrpCharacter.Width - lngLeft
        lngHeight = lngRowHeight * 4 + PixelY
        .picCharXP.Move lngLeft, lngTop, lngWidth, lngHeight
        .picCharXP.BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .picCharXP.Cls
        lngMiddle = .udrpCharacter.Left - .picCharXP.Left
        lngRight = .picCharXP.ScaleWidth - PixelX
        lngTop = 0
        lngBottom = lngRowHeight
        ' Draw table
        If cfg.Character = 0 Then
            ' No character selected, so show total XP total required per life (1st, 2nd, 3rd, Epic)
            lngTextColor = cfg.GetColor(cgeWorkspace, cveText)
            lngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
            ReDim varArray(1, 3)
            varArray(0, 0) = "1st"
            varArray(1, 0) = 1900000
            varArray(0, 1) = "2nd"
            varArray(1, 1) = 2850000
            varArray(0, 2) = "3rd"
            varArray(1, 2) = 3800000
            varArray(0, 3) = "Epic"
            varArray(1, 3) = 8250000
            For i = 0 To 3
                PrintText varArray(0, i), vbRightJustify, 0, lngTop, lngMiddle, lngBottom, lngTextColor, lngBackColor, False
                PrintText Format(varArray(1, i), "#,##0"), vbRightJustify, lngMiddle, lngTop, lngRight, lngBottom, lngTextColor, lngBackColor, False
                lngTop = lngBottom
                lngBottom = lngBottom + lngRowHeight
            Next
        ElseIf blnPastLives Then
            ' Selected character is capped and has past lives, show past life totals by type
            lngTextColor = cfg.GetColor(cgeWorkspace, cveText)
            lngBackColor = cfg.GetColor(cgeWorkspace, cveBackground)
            For i = 0 To 4
                If mlngLives(i) > 0 And lngLines < 4 Then ' Display max of 4 lines
                    lngLines = lngLines + 1
                    PrintText Format(mlngLives(i)), vbRightJustify, 0, lngTop, lngMiddle, lngBottom, lngTextColor, lngBackColor, False
                    PrintText GetPastLifeName(i, mlngLives(i)), vbLeftJustify, lngMiddle, lngTop, lngRight, lngBottom, lngTextColor, lngBackColor, False
                    lngTop = lngBottom
                    lngBottom = lngBottom + lngRowHeight
                End If
            Next
        Else
            ' Show XP table for next two levels (xp cap)
            lngTextColor = cfg.GetColor(cgeDropSlots, cveText)
            lngBackColor = cfg.GetColor(cgeDropSlots, cveBackground)
            PrintText "Level", vbCenter, 0, lngTop, lngMiddle, lngBottom, lngTextColor, lngBackColor
            PrintText "XP", vbCenter, lngMiddle, lngTop, lngRight, lngBottom, lngTextColor, lngBackColor
            ' Rows
            lngTextColor = cfg.GetColor(cgeControls, cveText)
            lngBackColor = cfg.GetColor(cgeControls, cveBackground)
            lngFirst = .usrspnChar.Value
            Select Case lngFirst
                Case 19
                    lngLast = 20
                Case 20
                    lngFirst = 21
                    lngLast = 23
                Case 28 To 30
                    lngLast = 30
                Case Else
                    lngLast = lngFirst + 2
            End Select
            For i = lngFirst To lngLast
                lngTop = lngBottom
                lngBottom = lngBottom + lngRowHeight
                PrintText Format(i), vbRightJustify, 0, lngTop, lngMiddle, lngBottom, lngTextColor, lngBackColor
                PrintText GetXP(i), vbRightJustify, lngMiddle, lngTop, lngRight, lngBottom, lngTextColor, lngBackColor
            Next
        End If
    End With
End Sub

Private Sub CountPastLives()
    Dim i As Long
    
    Erase mlngLives
    mlngLife = 0
    If cfg.Character = 0 Then Exit Sub
    With db.Character(cfg.Character).PastLife
        For i = 1 To UBound(.Class)
            IncrementPastLives pleClass, .Class(i)
        Next
        For i = 1 To UBound(.Racial)
            IncrementPastLives pleRace, .Racial(i)
        Next
        For i = 1 To UBound(.Iconic)
            IncrementPastLives pleIconic, .Iconic(i)
        Next
        For i = 1 To UBound(.Epic)
            IncrementPastLives pleEpic, .Epic(i)
        Next
    End With
    Select Case mlngLife
        Case 0: mlngLife = 1
        Case 1: mlngLife = 2
        Case Else: mlngLife = 3
    End Select
End Sub

Private Sub IncrementPastLives(penPastLife As PastLifeEnum, ByVal plngLives As Long)
    If plngLives > 3 Then plngLives = 3
    If penPastLife <> pleEpic Then mlngLife = mlngLife + plngLives
    mlngLives(penPastLife) = mlngLives(penPastLife) + plngLives
    mlngLives(pleTotal) = mlngLives(pleTotal) + plngLives
End Sub

Private Function GetPastLifeName(penPastLife As PastLifeEnum, plngLives As Long) As String
    Dim strPlural As String
    Dim strReturn As String
    
    If plngLives = 1 Then strPlural = " past life" Else strPlural = " past lives"
    Select Case penPastLife
        Case pleClass: strReturn = "class" & strPlural
        Case pleRace: strReturn = "racial" & strPlural
        Case pleEpic: strReturn = "epic" & strPlural
        Case pleIconic: strReturn = "iconic" & strPlural
        Case pleTotal: strReturn = "total" & strPlural
    End Select
    GetPastLifeName = strReturn
End Function

Private Function GetXP(plngLevel As Long) As String
    Dim lngRow As Long
    
    If plngLevel > 20 Then
        lngRow = FindTable("Epic")
        GetXP = Format(db.Table(lngRow).Row(plngLevel - 20).Value(2), "#,##0")
    Else
        lngRow = FindTable("Heroic")
        GetXP = Format(db.Table(lngRow).Row(plngLevel).Value(mlngLife + 1), "#,##0")
    End If
End Function

Private Function FindTable(pstrTable As String) As Long
    Dim i As Long
    
    For i = 1 To db.Tables
        If db.Table(i).TableID = pstrTable Then
            FindTable = i
            Exit For
        End If
    Next
End Function

Private Sub PrintText(ByVal pstrText As String, penAlign As AlignmentConstants, plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long, plngTextColor As Long, plngBackColor As Long, Optional pblnBorder As Boolean = True)
    Dim lngLeft As Long
    Dim lngWidth As Long
    Dim lngBorder As Long
    
    With UserControl.picCharXP
        .ForeColor = plngTextColor
        .FillColor = plngBackColor
        If pblnBorder Then lngBorder = cfg.GetColor(cgeControls, cveBorderExterior) Else lngBorder = plngBackColor
        UserControl.picCharXP.Line (plngLeft, plngTop)-(plngRight, plngBottom), lngBorder, B
        Select Case penAlign
            Case vbLeftJustify: lngLeft = plngLeft + .TextWidth(" ")
            Case vbCenter: lngLeft = plngLeft + (plngRight - plngLeft - .TextWidth(pstrText)) \ 2
            Case vbRightJustify: lngLeft = plngRight - .TextWidth(pstrText & " ")
        End Select
        .CurrentX = lngLeft
        .CurrentY = plngTop + PixelY
        UserControl.picCharXP.Print pstrText
    End With
End Sub


' ************* BUTTONS *************


Private Sub usrbtnChar_Click(Index As Integer, Caption As String)
    Select Case Caption
        Case "Epic Sagas": OpenSagas steEpic
        Case "Heroic Sagas": OpenSagas steHeroic
        Case "Challenges": OpenChallenges
        Case "Reincarnate": Reincarnate cfg.Character
        Case "Character...": CharactersMenu Index
    End Select
End Sub

Private Sub CharactersMenu(Index As Integer)
    Dim blnCharacter As Boolean
    
    blnCharacter = (cfg.Character > 0)
    With UserControl
        .mnuCharacter(1).Enabled = blnCharacter
        .mnuCharacter(2).Enabled = blnCharacter
        .mnuCharacter(5).Enabled = blnCharacter
        With .usrbtnChar(Index)
            PopupMenu UserControl.mnuMain(0), , .Left, .Top + .Height
        End With
    End With
End Sub

Private Sub mnuCharacter_Click(Index As Integer)
    Select Case UserControl.mnuCharacter(Index).Caption
        Case "Add New": AddCharacter
        Case "Delete": DeleteCharacter
        Case "Rename"
        Case "Import"
    End Select
End Sub

Private Sub mnuExport_Click(Index As Integer)
    Select Case UserControl.mnuExport(Index).Caption
        Case "Character Builder Lite"
        Case "Ron's Character Planner"
        Case "DDO Builder"
    End Select
End Sub

Private Sub OpenSagas(penSagaTier As SagaTierEnum)
    cfg.SagaTier = penSagaTier
    frmSagas.Character = cfg.Character
    OpenForm "frmSagas"
End Sub

Private Sub OpenChallenges()
    frmChallenges.Character = cfg.Character
    OpenForm "frmChallenges"
End Sub


' ************* ADD NEW *************


Private Sub AddCharacter()
    Dim i As Long
    
    db.Characters = db.Characters + 1
    ReDim Preserve db.Character(1 To db.Characters)
    db.Character(db.Characters) = DefaultCharacter()
    InitCharacterSagas db.Character(db.Characters)
    For i = 1 To db.Quests
        ReDim Preserve db.Quest(i).Progress(1 To db.Characters)
    Next
    For i = 1 To db.Challenges
        ReDim Preserve db.Challenge(i).Stars(1 To db.Characters)
    Next
    cfg.Character = db.Characters
    RefreshCharacterList
    ShowCharacter
    RaiseEvent CharacterListChanged
End Sub

Private Function DefaultCharacter() As CharacterType
    Dim i As Long
    
    With DefaultCharacter
        .Character = NewName()
        .CustomColor = False
        .GeneratedColor = NewColor()
        .BackColor = GetColorValue(.GeneratedColor)
        .DimColor = GetColorDim(.BackColor)
        If db.Sagas Then
            ReDim .Saga(1 To db.Sagas)
            For i = 1 To db.Sagas
                ReDim .Saga(i).Progress(1 To db.Saga(i).Quests)
            Next
        End If
    End With
End Function

Private Function NewName() As String
    Dim lngNew As Long
    Dim blnFound As Boolean
    Dim i As Long
    
    Do
        blnFound = False
        lngNew = lngNew + 1
        For i = 1 To db.Characters
            If db.Character(i).Character = "New" & lngNew Then
                blnFound = True
                Exit For
            End If
        Next
    Loop Until Not blnFound
    NewName = "New" & lngNew
End Function

Private Function NewColor() As GeneratedColorEnum
    Dim blnTaken() As Boolean
    Dim lngNew As Long
    Dim i As Long
    
    ' Create list of taken colors
    ReDim blnTaken(gceColors)
    If db.Characters < gceColors - 1 Then ' No dups: Choose an unused color at random
        For i = 1 To db.Characters
            blnTaken(db.Character(i).GeneratedColor) = True
        Next
    ElseIf db.Characters > 1 Then ' Dups: Choose any color except the one right next to this new character
        blnTaken(db.Character(db.Characters - 1).GeneratedColor) = True
    End If
    ' Randomly choose a color until a valid color found
    lngNew = RandomNumber(gceColors - 1)
    Do While blnTaken(lngNew)
        lngNew = RandomNumber(gceColors - 1)
    Loop
    NewColor = lngNew
End Function


' ************* DELETE *************


Private Sub DeleteCharacter()
    If cfg.Character = 0 Then Exit Sub
    If MsgBox("Delete " & db.Character(cfg.Character).Character & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
    DeleteChar
    RefreshCharacterList
    RaiseEvent CharacterListChanged
End Sub

Private Sub DeleteChar()
    Dim i As Long
    Dim c As Long
    
    If cfg.Character = 0 Then Exit Sub
    For i = 1 To db.Quests
        With db.Quest(i)
            If db.Characters - 1 = 0 Then
                Erase .Progress
            Else
                For c = cfg.Character To db.Characters - 1
                    .Progress(c) = .Progress(c + 1)
                Next
                ReDim Preserve .Progress(1 To db.Characters - 1)
            End If
        End With
    Next
    For i = 1 To db.Challenges
        With db.Challenge(i)
            If db.Characters - 1 = 0 Then
                Erase .Stars
            Else
                For c = cfg.Character To db.Characters - 1
                    .Stars(c) = .Stars(c + 1)
                Next
                ReDim Preserve .Stars(1 To db.Characters - 1)
            End If
        End With
    Next
    db.Characters = db.Characters - 1
    If db.Characters = 0 Then
        Erase db.Character
        cfg.Character = 0
    Else
        For c = cfg.Character To db.Characters
            db.Character(c) = db.Character(c + 1)
        Next
        ReDim Preserve db.Character(1 To db.Characters)
        cfg.Character = db.Characters
    End If
End Sub
