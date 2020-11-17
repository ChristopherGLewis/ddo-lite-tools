VERSION 5.00
Begin VB.Form frmRefreshData 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refresh Data"
   ClientHeight    =   4425
   ClientLeft      =   30
   ClientTop       =   405
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRefreshData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   3
      Left            =   2880
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   3240
      Width           =   1812
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   2
      Left            =   2880
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   2760
      Width           =   1812
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   1
      Left            =   2880
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   7
      Top             =   2280
      Width           =   1812
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   0
      Left            =   2880
      ScaleHeight     =   375
      ScaleWidth      =   1815
      TabIndex        =   5
      Top             =   1680
      Width           =   1812
   End
   Begin VB.CheckBox chkRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Refresh Trees"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   2052
   End
   Begin VB.CheckBox chkRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Destinies"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   2052
   End
   Begin VB.CheckBox chkRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enhancements"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2052
   End
   Begin VB.CheckBox chkRefresh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spells"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   1
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   2052
   End
   Begin VB.Timer tmrCrawl 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   3900
   End
   Begin VB.PictureBox picStatbar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   1020
      ScaleHeight     =   345
      ScaleWidth      =   1785
      TabIndex        =   4
      Top             =   3780
      Width           =   1812
   End
   Begin VB.Label lblCurrent 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "internet resources, including ddowiki."
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   1
      Left            =   600
      TabIndex        =   11
      Top             =   600
      Width           =   3972
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "WARNING: These routines crawl various"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   240
      Width           =   3972
   End
   Begin VB.Shape Shape1 
      Height          =   1212
      Left            =   360
      Top             =   120
      Width           =   4332
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   1812
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Please do not run them unless required."
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   2
      Left            =   600
      TabIndex        =   12
      Top             =   960
      Width           =   3972
   End
End
Attribute VB_Name = "frmRefreshData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ChangeEnum
    ceNotFound
    ceWikiLinkChanged
    ceDescriptionChanged
    ceInvalidPage
    ceInvalidTier
End Enum

Private Type ChangeType
    Change As ChangeEnum
    Page As String
    Item As String
    OldText As String
    NewText As String
End Type

Private mclsStatbar As clsStatusbar

Private mstrSpellPage() As String
Private mlngSpellPages As Long
Private mlngSpellPage As Long
Private mblnSortSpells As Boolean

Private mblnOverride As Boolean
Private mlngValue(3) As Long   ' This records the progress. (0) is total, 1/spells, 2/enh, 3/Dest
Private mlngTotal(3) As Long
Private mlngCurrent As Long
Private mblnCrawl(3) As Long  'This records if something is to be crawled or has been crawled

Private mtypChange() As ChangeType
Private mlngChanges As Long
Private mstrCurrentPage As String

' Used only for display purposes to identify what changed
Private mstrTree As String
Private mlngTier As Long

Private Sub Form_Load()
    cfg.RefreshColors Me
    Set mclsStatbar = New clsStatusbar
    mclsStatbar.Init Me.picStatbar
    mclsStatbar.AddPanel vbNullString, vbLeftJustify, pseSpring, 100, True
    mblnSortSpells = False
    ReDim mtypChange(127)
    mlngChanges = 0
    mstrCurrentPage = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xp.Mouse = msNormal
    Set mclsStatbar = Nothing
    If mblnSortSpells Then SortSpells
End Sub

Private Sub chkRefresh_Click(Index As Integer)
'Checkbox used as a push button.
    Dim blnAll As Boolean
    Dim i As Long
    
    'Index 0 is all, 1/spells, 2/enh, 3/dest
    'This sets our checkboxes that eventally set mblnCrawl()
    
    If UncheckButton(Me.chkRefresh(Index), mblnOverride) Then Exit Sub
    Erase mblnCrawl
    For i = 0 To 3
        Me.chkRefresh(i).Enabled = False
    Next
    ClearProgress
    If Index = 0 Then mlngCurrent = 2 Else mlngCurrent = Index
    If Index = 1 Then PrepSpells
    If Index = 0 Or Index = 2 Then PrepEnhancements
    If Index = 0 Or Index = 3 Then PrepDestinies
    
    'If Index = 0 Or Index = 4 Then PrepFeats
    
    ShowProgressIndex 0
    mclsStatbar.ProgressbarInit 1, mlngTotal(0)
    xp.Mouse = msAppWait
    Me.tmrCrawl.Enabled = True
End Sub

Private Sub ClearProgress()
    Dim i As Long
    
    Erase mlngValue, mlngTotal
    For i = 0 To 3
        Me.picProgress(i).Cls
    Next
End Sub

Private Sub ShowProgress()
    ShowProgressIndex 0
    ShowProgressIndex mlngCurrent
End Sub

Private Sub ShowProgressIndex(plngIndex As Long)
    Dim strDisplay As String
    
    Me.picProgress(plngIndex).Cls
    Me.picProgress(plngIndex).ForeColor = cfg.GetColor(cgeControls, cveText)
    If mlngTotal(plngIndex) = 0 Then Exit Sub
    strDisplay = mlngValue(plngIndex) & " of " & mlngTotal(plngIndex) & " (" & Int(100 * mlngValue(plngIndex) / mlngTotal(plngIndex) + 0.5) & "%)"
    With Me.picProgress(plngIndex)
        .CurrentX = (.ScaleWidth - .TextWidth(strDisplay)) \ 2
        .CurrentY = (.ScaleHeight - .TextHeight(strDisplay)) \ 2
    End With
    Me.picProgress(plngIndex).Print strDisplay
End Sub

Private Sub tmrCrawl_Timer()
    Me.tmrCrawl.Enabled = False
    
    If Increment() Then
        Select Case mlngCurrent
            Case 1: CrawlSpells mstrSpellPage(mlngValue(1))
            Case 2: CrawlTree db.Tree(mlngValue(2))
            Case 3: CrawlTree db.Destiny(mlngValue(3))
            'Case 4: CrawlFeat db.Feat(mlngValue(4))
        End Select
        Me.tmrCrawl.Interval = RandomDelay()
        Me.tmrCrawl.Enabled = True
    Else
        EndCrawl
    End If
End Sub

Private Sub EndCrawl()
    Dim i As Long
    
    xp.Mouse = msNormal
    mclsStatbar.ProgressbarRemove
    For i = 0 To 3
        Me.chkRefresh(i).Enabled = True
        Me.picProgress(i).Cls
    Next
    SaveChanges
End Sub

Private Function Increment() As Boolean
    Do
        If mlngValue(mlngCurrent) >= mlngTotal(mlngCurrent) Then
            mlngCurrent = mlngCurrent + 1
            If mlngCurrent > 1 Then mstrCurrentPage = vbNullString
            If mlngCurrent = 2 And mlngChanges > 0 Then SortChanges
        Else
            mlngValue(0) = mlngValue(0) + 1
            mlngValue(mlngCurrent) = mlngValue(mlngCurrent) + 1
            ShowProgress
            mclsStatbar.ProgressbarIncrement
            Increment = True
            Exit Function
        End If
    Loop Until mlngCurrent > 3
End Function

Private Function RandomDelay() As Long
    '2020.10.15 changed from 2000 to 500
    RandomDelay = Int(2001 * Rnd + 500)
End Function


' ************* SPELLS *************
Private Sub PrepSpells()
    If db.Spells = 0 Then Exit Sub
    SortSpellWiki
    mblnSortSpells = True
    ReDim mstrSpellPage(1 To 16)
    mlngSpellPages = 0
    InitSpellPage "https://ddowiki.com/page/Artificer_spells"
    InitSpellPage "https://ddowiki.com/page/Bard_spells"
    InitSpellPage "https://ddowiki.com/page/Cleric_/_Favored_Soul_spells"
    InitSpellPage "https://ddowiki.com/page/Druid_spells"
    InitSpellPage "https://ddowiki.com/page/Paladin_spells"
    InitSpellPage "https://ddowiki.com/page/Ranger_spells"
    InitSpellPage "https://ddowiki.com/page/Sorcerer_/_Wizard_spells"
    InitSpellPage "https://ddowiki.com/page/Warlock_spells"
    InitSpellPage "https://ddowiki.com/page/Alchemist_spells"
    ReDim Preserve mstrSpellPage(1 To mlngSpellPages)
    mlngTotal(0) = mlngTotal(0) + mlngSpellPages
    mlngTotal(1) = mlngSpellPages
    mblnCrawl(1) = True
    ShowProgressIndex 1
End Sub

Private Sub InitSpellPage(pstrPage As String)
    mlngSpellPages = mlngSpellPages + 1
    mstrSpellPage(mlngSpellPages) = pstrPage
End Sub

Private Sub CrawlSpells(pstrPage As String)
    Dim strRaw As String
    Dim strSpells() As String
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strSplit() As String
    Dim i As Long
    Dim j As Long
    
    mstrCurrentPage = UnWiki(pstrPage)
    mstrCurrentPage = Left(mstrCurrentPage, Len(mstrCurrentPage) - 7)
    strRaw = DownloadURL(pstrPage)
    lngStart = InStr(strRaw, "Metamagic Feats")
    lngEnd = InStrRev(strRaw, "</table>")
    strRaw = Mid$(strRaw, lngStart, lngEnd - lngStart)
    strRaw = Replace(strRaw, Chr(10), "")
    strSpells = Split(strRaw, "</tr>")
    For i = 0 To UBound(strSpells)
        If Left$(strSpells(i), 10) = "<tr style=" Then
            If j <> i Then
                strSpells(j) = strSpells(i)
                j = j + 1
            End If
        End If
    Next
    ReDim Preserve strSpells(j - 1)
    For i = 0 To UBound(strSpells)
        CleanSpell strSpells(i)
    Next
End Sub

Private Sub CleanSpell(pstrRaw As String)
    Const Quote As String = """"
    Dim strSpell As String
    Dim strLink As String
    Dim strDescrip As String
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngIndex As Long
    
    ' Link
    lngStart = InStr(pstrRaw, "<a href=") + 9
    lngEnd = InStr(lngStart, pstrRaw, Quote)
    If lngEnd > lngStart Then strLink = Mid$(pstrRaw, lngStart, lngEnd - lngStart)
    strLink = UnWiki(strLink)
    ' Spell
    lngStart = InStr(lngEnd, pstrRaw, ">") + 1
    lngEnd = InStr(lngStart, pstrRaw, "<")
    If lngEnd > lngStart Then strSpell = Mid$(pstrRaw, lngStart, lngEnd - lngStart)
    ' Descrip
    lngStart = InStr(lngEnd, pstrRaw, "<td style=")
    lngStart = InStr(lngStart, pstrRaw, ">") + 1
    lngEnd = InStr(lngStart, pstrRaw, "</td>")
    If lngEnd > lngStart Then strDescrip = Mid$(pstrRaw, lngStart, lngEnd - lngStart)
    ' Remove codes
    Do
        lngStart = InStr(strDescrip, "<")
        lngEnd = InStr(lngStart + 1, strDescrip, ">")
        If lngEnd = 0 Then Exit Do
        strDescrip = Left$(strDescrip, lngStart - 1) & Mid$(strDescrip, lngEnd + 1)
    Loop
    strDescrip = Trim$(strDescrip)
    ' Finish
    lngIndex = SeekSpellWiki(strLink)
    If lngIndex = 0 Then
        AddChange ceNotFound, "Spell not found: " & strSpell
    Else
        With db.Spell(lngIndex)
            If .Wiki <> strLink Then
                AddChange ceWikiLinkChanged, "Spell: & " & .SpellName, .Wiki, strLink
                .Wiki = strLink
            End If
            If .Descrip <> strDescrip Then
                AddChange ceDescriptionChanged, "Spell: " & .SpellName, .Descrip, strDescrip
                .Descrip = strDescrip
            End If
        End With
    End If
End Sub

Private Function ProcessSpell(pstrRaw As String) As String
    Dim strSpell As String
    
    strSpell = Split(pstrRaw, "<tr style=")
End Function

' ************* ENHANCEMENTS *************
Private Sub PrepEnhancements()
    mlngTotal(0) = mlngTotal(0) + db.Trees
    mlngTotal(2) = db.Trees
    mblnCrawl(2) = True
    ShowProgressIndex 2
End Sub


' ************* DESTINIES *************
Private Sub PrepDestinies()
    If db.Destinies = 0 Then Exit Sub
    mlngTotal(0) = mlngTotal(0) + db.Destinies
    mlngTotal(3) = db.Destinies
    mblnCrawl(3) = True
    ShowProgressIndex 3
End Sub

' ************* FEATS *************
Private Sub PrepFeats()
    mlngTotal(0) = mlngTotal(0) + db.Feats
    mlngTotal(4) = db.Feats
    mblnCrawl(4) = True
    ShowProgressIndex 4
End Sub

' ************* TREES *************
Private Sub CrawlTree(ptypTree As TreeType)
    Dim strPage As String
    Dim strRaw As String
    Dim strTier() As String
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim i As Long
    
    mstrTree = ptypTree.TreeName
    ' Range of Tiers
    Select Case ptypTree.TreeType
        Case tseClass, tseRaceClass, tseGlobal
            lngLast = 5
        Case tseRace
            lngLast = 4
        Case tseDestiny
            lngFirst = 1
            lngLast = 6
    End Select
    ' Wiki page
    lblCurrent.Caption = ptypTree.Wiki
    strPage = MakeWiki(ptypTree.Wiki)
    ' Process page
    strRaw = DownloadURL(strPage)
    'wikitable is a special class - trying to change to "wikitable enhancement"
    strTier = Split(strRaw, "wikitable")
    If UBound(strTier) <> lngLast + 1 Then
        AddChange ceInvalidPage, strPage
    Else
        'Parse each table as a tier.  lngLast = max tier
        For i = lngFirst To lngLast
            mlngTier = i
            SplitTierIntoAbilities strTier(i + 1), ptypTree.Tier(i), ptypTree.TreeName, i, ptypTree.TreeType
        Next
    End If
End Sub

Private Sub SplitTierIntoAbilities(pstrRaw As String, ptypTier As TierType, pstrTreeName As String, plngTier As Long, penTreeType As TreeStyleEnum)
    Dim strAbility() As String
    Dim blnSkip As Boolean
    Dim i As Long
    
    'Each row is an enhancement - Wood Elf has a |-Style so clear that
    pstrRaw = Replace(pstrRaw, "<tr style=""background-color:lightcyan;"">", "<tr>")
    strAbility = Split(pstrRaw, "<tr>")
    ' Special consideration for archmage, which has a unique SLA table between core and tier 1 on its wiki page
    If UBound(strAbility) = 11 Then  'Typically 5
        If pstrTreeName = "Archmage" And plngTier = 0 Then ReDim Preserve strAbility(6)
    End If
    ' Process tier
    If UBound(strAbility) <> ptypTier.Abilities Then
        AddChange ceInvalidTier, pstrTreeName & " Tier " & plngTier
    Else
        For i = 1 To ptypTier.Abilities
            blnSkip = False
            If i = ptypTier.Abilities Then
                If (penTreeType = tseClass Or penTreeType = tseGlobal Or penTreeType = tseRaceClass) And (plngTier = 3 Or plngTier = 4) And (pstrTreeName <> "Falconry") Then blnSkip = True
                If penTreeType = tseDestiny Then blnSkip = True
            End If
            If Not blnSkip Then CleanAbility strAbility(i), ptypTier.Ability(i)
        Next
    End If
End Sub

Private Sub CleanAbility(pstrRaw As String, ptypAbility As AbilityType)
    Dim strAbility As String
    Dim strDescrip As String
    Dim lngStart As Long
    Dim lngEnd As String
    
    'Name of ability is between the first <b>...</b>
    lngStart = InStr(pstrRaw, "<b>")
    If lngStart = 0 Then Exit Sub Else lngStart = lngStart + 3
    lngEnd = InStr(lngStart, pstrRaw, "</b>")
    If lngEnd < lngStart Then Exit Sub
    strAbility = Mid$(pstrRaw, lngStart, lngEnd - lngStart)
    'Description is 5 later until the first /DIV
    lngStart = lngEnd + 5
    lngEnd = InStr(lngStart, pstrRaw, "</div>")
    If lngEnd < lngStart Then Exit Sub
    strDescrip = Trim$(Mid$(pstrRaw, lngStart, lngEnd - lngStart))
    CleanCodes strAbility
    CleanCodes strDescrip
    With ptypAbility
        If .Descrip <> strDescrip Then
            AddChange ceDescriptionChanged, mstrTree & " Tier " & mlngTier & ": " & ptypAbility.AbilityName, .Descrip, strDescrip
            .Descrip = strDescrip
        End If
    End With
End Sub

Private Sub CleanCodes(pstrRaw As String)
    ' } = carriage return
    ' { = indent (trim left so that word wrapping lines up properly)
    ' { is followed by the middle character in the 3-space indentation for first line (typically a bullet point, dash or space)
    Dim strSplit() As String
    Dim lngPos As Long
    Dim i As Long
    
    ' Process line feeds and indentation
    If InStr(pstrRaw, vbLf) Then pstrRaw = Replace(pstrRaw, vbLf, "}")
    If InStr(pstrRaw, "<dl>") Then pstrRaw = Replace(pstrRaw, "<dl>", "}")
    If InStr(pstrRaw, "<dd>") Then pstrRaw = Replace(pstrRaw, "<dd>", "{ ")
    If InStr(pstrRaw, "<ul>") Then pstrRaw = Replace(pstrRaw, "<ul>", "")
    If InStr(pstrRaw, "<li>") Then
        pstrRaw = Replace(pstrRaw, "<li>", "}{-")
        pstrRaw = Replace(pstrRaw, "</li>", "")
    End If
    ' Remove all remaining html codes
    If InStr(pstrRaw, "<") Then
        strSplit = Split(pstrRaw, "<")
        For i = 0 To UBound(strSplit)
            lngPos = InStr(strSplit(i), ">")
            If lngPos Then
                strSplit(i) = Mid$(strSplit(i), lngPos + 1)
            ElseIf i > 0 Then
                strSplit(i) = "<" & strSplit(i)
            End If
        Next
        pstrRaw = Join(strSplit, vbNullString)
    End If
    ' Convert escape codes
    If InStr(pstrRaw, "&lt;") Then pstrRaw = Replace(pstrRaw, "&lt;", "<")
    If InStr(pstrRaw, "&gt;") Then pstrRaw = Replace(pstrRaw, "&gt;", ">")
    If InStr(pstrRaw, "&amp;") Then pstrRaw = Replace(pstrRaw, "&amp;", "&")
    ' Remove html codes that were "hidden" by being replaced with &lt; or &gt;
    If InStr(pstrRaw, "</div>") Then pstrRaw = Replace(pstrRaw, "</div>", "")
    If InStr(pstrRaw, "</s>") Then pstrRaw = Replace(pstrRaw, "</s>", "")
    If InStr(pstrRaw, "&#160;:") Then pstrRaw = Replace(pstrRaw, "&#160;:", "")
    ' Remove extraneous blank lines
    Do While InStr(pstrRaw, " }")
        pstrRaw = Replace(pstrRaw, " }", "}")
    Loop
    Do While InStr(pstrRaw, "}}")
        pstrRaw = Replace(pstrRaw, "}}", "}")
    Loop
    Do While InStr(pstrRaw, "{}")
        pstrRaw = Replace(pstrRaw, "{}", "")
    Loop
    ' Final cleanup
    Do
        pstrRaw = Trim$(pstrRaw)
        If Right$(pstrRaw, 1) = "}" Then pstrRaw = Left$(pstrRaw, Len(pstrRaw) - 1) Else Exit Do
    Loop
End Sub


' ************* FEATS *************
Private Sub CrawlFeat(ptypTree As TreeType)
    Dim strPage As String
    Dim strRaw As String
    Dim strTier() As String
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim i As Long
End Sub


' ************* MESSAGES *************
Private Sub AddChange(penChange As ChangeEnum, pstrItem As String, Optional pstrOld As String, Optional pstrNew As String)
    Dim lngMax As Long
    
    mlngChanges = mlngChanges + 1
    lngMax = UBound(mtypChange)
    If mlngChanges > lngMax Then
        lngMax = (lngMax * 3) \ 2
        ReDim Preserve mtypChange(lngMax)
    End If
    With mtypChange(mlngChanges)
        .Change = penChange
        .Page = mstrCurrentPage
        .Item = pstrItem
        .OldText = pstrOld
        .NewText = pstrNew
    End With
End Sub

Private Sub SaveChanges()
    Dim strFile As String
    Dim strLine() As String
    Dim lngLine As Long
    Dim lngCount As Long
    Dim strItem As String
    Dim i As Long
    
    If mlngChanges = 0 Then Exit Sub
    ' Create string array
    ReDim strLine(127)
    ' Descriptions changed
    lngCount = 0
    For i = 1 To mlngChanges
        If mtypChange(i).Change = ceDescriptionChanged Then lngCount = lngCount + 1
    Next
    If lngCount > 0 Then
        AddLine strLine, lngLine, lngCount & " descriptions changed:", True
        For i = 1 To mlngChanges
            With mtypChange(i)
                If .Change = ceDescriptionChanged Then
                    If Len(.Page) Then strItem = .Item & " (" & .Page & ")" Else strItem = .Item
                    AddLine strLine, lngLine, strItem
                    AddLine strLine, lngLine, "Old: " & .OldText
                    AddLine strLine, lngLine, "New: " & .NewText, True
                End If
            End With
        Next
    End If
    ' Miscellaneous errors
    lngCount = 0
    For i = 1 To mlngChanges
        If mtypChange(i).Change <> ceDescriptionChanged Then lngCount = lngCount + 1
    Next
    If lngCount > 0 Then
        AddLine strLine, lngLine, lngCount & " miscellaneous errors:", True
        For i = 1 To mlngChanges
            With mtypChange(i)
                Select Case .Change
                    Case ceInvalidPage: AddLine strLine, lngLine, "Invalid wiki link: " & .Item, True
                    Case ceInvalidTier: AddLine strLine, lngLine, "Invalid tier: " & .Item, True
                    Case ceNotFound: AddLine strLine, lngLine, .Item, True
                    Case ceWikiLinkChanged
                        AddLine strLine, lngLine, .Item & " link changed:"
                        AddLine strLine, lngLine, "Old: " & .OldText
                        AddLine strLine, lngLine, "New: " & .NewText, True
                End Select
            End With
        Next
    End If
    ' Change log
    lngCount = 0
    For i = 1 To mlngChanges
        If mtypChange(i).Change = ceDescriptionChanged Then lngCount = lngCount + 1
    Next
    If lngCount > 0 Then
        AddLine strLine, lngLine, lngCount & " changes:[list]", True
        For i = 1 To mlngChanges
            With mtypChange(i)
                If .Change = ceDescriptionChanged Then AddLine strLine, lngLine, "[*]" & .Item
            End With
        Next
    End If
    ReDim Preserve strLine(lngLine)
    strFile = App.Path & "\Crawl.txt"
    xp.File.SaveStringAs strFile, Join(strLine, vbNewLine)
    xp.File.Run strFile
    Erase mtypChange
    mlngChanges = 0
    If AskAlways("Commit changes?") Then
        If mblnCrawl(1) Then SaveSpellFile
        If mblnCrawl(2) Then SaveEnhancementsFile
        If mblnCrawl(3) Then SaveDestinyFile
        mblnSortSpells = False
    End If
End Sub

Private Sub AddLine(pstrLine() As String, plngLine As Long, pstrText As String, Optional pblnLineFeed As Boolean = False)
    Dim lngMax As Long
    
    lngMax = UBound(pstrLine)
    If plngLine > lngMax Then
        lngMax = (lngMax * 3) \ 2
        ReDim Preserve pstrLine(lngMax)
    End If
    pstrLine(plngLine) = pstrText
    If pblnLineFeed Then pstrLine(plngLine) = pstrLine(plngLine) & vbNewLine
    plngLine = plngLine + 1
End Sub

Private Function DownloadURL(pstrURL As String) As String
    Dim lngRandom As Long
    Dim strURL As String
    
    lngRandom = Int((9000) * Rnd + 1000)
    strURL = pstrURL & "?rnd=" & lngRandom
    DownloadURL = xp.DownloadURL(strURL)
End Function

Private Sub SortChanges()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As ChangeType
    
    Exit Sub
    iMin = LBound(mtypChange) + 1
    iMax = UBound(mtypChange)
    For i = iMin To iMax
        typSwap = mtypChange(i)
        For j = i To iMin Step -1
            If typSwap.Item < mtypChange(j - 1).Item Then mtypChange(j) = mtypChange(j - 1) Else Exit For
        Next j
        mtypChange(j) = typSwap
    Next i
End Sub
