Attribute VB_Name = "basOutput"
Option Explicit

Private Const ColorDim As String = "[color=silver]"
Private Const ColorCollectable As String = "[color=lightskyblue]"
Private Const ColorRed As String = "[color=lightpink]"
Private Const ColorOff As String = "[/color]"
Private Const BoldOn As String = "[b]"
Private Const BoldOff As String = "[/b]"

Public Type InfoFreqType
    Materials As Long
    Material() As String
End Type

Public Type InfoTierType
    Freq(1 To 3) As InfoFreqType
End Type

Public Type InfoSchoolType
    Tier(1 To 6) As InfoTierType
End Type

Public Type InfoType
    School(1 To 4) As InfoSchoolType
End Type

Private mtypInfo As InfoType
Private mstrLine() As String
Private mlngLines As Long

Public Sub LoadSchoolInfo(ptypInfo As InfoType)
    Dim typBlank As InfoType
    Dim i As Long
    
    ptypInfo = typBlank
    For i = 1 To db.Materials
        With db.Material(i)
            If .MatType = meCollectable Then
                With ptypInfo.School(.School).Tier(.Tier).Freq(.Frequency)
                    .Materials = .Materials + 1
                    ReDim Preserve .Material(1 To .Materials)
                    .Material(.Materials) = db.Material(i).Material
                End With
            End If
        End With
    Next
End Sub

Private Sub GetOutputFarms(ByVal penSchool As SchoolEnum, plngTier As Long, ptypFarm() As MaterialFarmType, plngFarms As Long)
    Dim i As Long
    
    plngFarms = 0
    Erase ptypFarm
    With db.School(penSchool).Tier(plngTier)
        For i = 1 To .TierFarms
            AddFarm penSchool, .TierFarm(i), ptypFarm, plngFarms
        Next
    End With
End Sub

Private Sub AddFarm(ByVal penSchool As SchoolEnum, ptypTierFarm As TierFarmType, ptypFarm() As MaterialFarmType, plngFarms As Long)
    Dim lngFarm As Long
    Dim dblDispensers As Double
    Dim dblSeconds As Double
    Dim dblRate As Double
    Dim typNew As MaterialFarmType
    Dim i As Long
    
    lngFarm = SeekFarm(ptypTierFarm.Farm)
    If lngFarm = 0 Then Exit Sub
    ' Count dispensers
    With db.Farm(lngFarm)
        dblSeconds = .Seconds
        Select Case penSchool
            Case seArcane: dblDispensers = .Arcane
            Case seLore: dblDispensers = .Lore
            Case seNatural: dblDispensers = .Natural
        End Select
        dblDispensers = dblDispensers + (.Any * db.Backpack(penSchool))
    End With
    ' Calculate seconds to farm up one collectable
    If dblSeconds <> 0 And dblDispensers <> 0 Then dblRate = dblSeconds / dblDispensers
    ' Load data into new tierfarm
    typNew.Farm = db.Farm(lngFarm)
    typNew.Difficulty = ptypTierFarm.Difficulty
    typNew.Dispensers = dblDispensers
    typNew.Rate = Int(dblRate + 0.5)
    ConvertCodes typNew.Farm.Notes
    ' Add this new tierfarm to the list in its proper sorted position
    plngFarms = plngFarms + 1
    ReDim Preserve ptypFarm(1 To plngFarms)
    ptypFarm(plngFarms) = typNew
    If typNew.Rate = 0 And typNew.Farm.TreasureBag = False Then Exit Sub
    For i = plngFarms To 2 Step -1
        If Not ptypFarm(i - 1).Farm.TreasureBag Then
            If ptypFarm(i - 1).Rate = 0 Or ptypFarm(i - 1).Rate > ptypFarm(i).Rate Or (ptypFarm(i - 1).Rate = ptypFarm(i).Rate And ptypFarm(i - 1).Dispensers < ptypFarm(i).Dispensers) Or (ptypFarm(i - 1).Farm.TreasureBag = False And ptypFarm(i).Farm.TreasureBag = True) Then
                typNew = ptypFarm(i - 1)
                ptypFarm(i - 1) = ptypFarm(i)
                ptypFarm(i) = typNew
            End If
        End If
    Next
End Sub

Public Sub ConvertCodes(pstrNotes As String)
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim strNew As String
    Dim strLink As String
    
    If InStr(pstrNotes, "{") = 0 Then Exit Sub
    pstrNotes = Replace(pstrNotes, "{", "[")
    pstrNotes = Replace(pstrNotes, "}", "]")
    lngLeft = 1
    Do
        lngLeft = InStr(lngLeft, pstrNotes, "]")
        If lngLeft = 0 Then Exit Do
        lngRight = InStr(lngLeft, pstrNotes, " ")
        If lngRight = 0 Then Exit Do
        strLink = Mid$(pstrNotes, lngLeft + 1, lngRight - lngLeft - 1)
        strLink = Replace(strLink, "_", " ") & "[/url]"
        strNew = strNew & Left$(pstrNotes, lngLeft) & strLink & Mid$(pstrNotes, lngRight)
        lngLeft = lngRight
    Loop
    pstrNotes = strNew
End Sub

Public Sub Output()
    Dim i As Long
    
    LoadSchoolInfo mtypInfo
    OutputCollectables
    For i = 1 To 6
        OutputTier i
    Next
End Sub

Private Sub OutputCollectables()
    Dim strFile As String
    Dim i As Long
    
    ReDim mstrLine(db.Materials)
    mlngLines = 0
    For i = 1 To db.Materials
        With db.Material(i)
            If .MatType = meCollectable Then
'                OutputLine "[td]" & .Material & "[/td][td]" & GetSchoolName(.School) & " Tier " & .Tier & "[/td]"
                OutputLine .Material & " " & ColorDim & GetSchoolName(.School) & " Tier " & .Tier & ColorOff
            End If
        End With
    Next
    strFile = App.Path & "\Output0.txt"
    mlngLines = mlngLines - 1
    If mlngLines < 0 Then Exit Sub
'    CompactList
'    For i = 0 To mlngLines
'        mstrLine(i) = "[tr]" & mstrLine(i) & "[/tr]"
'    Next
    ReDim Preserve mstrLine(mlngLines)
'    mstrLine(0) = "[table=" & Chr(34) & "width: 500" & Chr(34) & "]" & mstrLine(0)
'    mstrLine(mlngLines) = mstrLine(mlngLines) & "[/table]"
    xp.File.SaveStringAs strFile, Join(mstrLine, vbNewLine)
End Sub

Private Sub CompactList()
    Dim lngMax As Long
    Dim i As Long
    
    lngMax = 1 + mlngLines \ 3
    For i = 0 To lngMax - 1
        mstrLine(i) = mstrLine(i) & "[td][/td]" & mstrLine(i + lngMax)
        If i + lngMax * 2 <= mlngLines Then mstrLine(i) = "[tr]" & mstrLine(i) & "[td][/td]" & mstrLine(i + lngMax * 2)
    Next
    mlngLines = lngMax
End Sub

Private Sub OutputTier(plngTier As Long)
    Dim strFile As String
    Dim enSchool As SchoolEnum
    Dim lngFarm As Long
    Dim i As Long
    
    ReDim mstrLine(63)
    mlngLines = 0
    For enSchool = seArcane To seNatural
        OutputSchool enSchool, plngTier
        OutputFarms enSchool, plngTier
        OutputLine vbNullString, 2
    Next
    strFile = App.Path & "\Output" & plngTier & ".txt"
    ReDim Preserve mstrLine(mlngLines)
    xp.File.SaveStringAs strFile, Join(mstrLine, vbNewLine)
End Sub

Private Sub OutputSchool(penSchool As SchoolEnum, plngTier As Long)
    Dim strHeader As String
    Dim strLevel As String
    Dim enFreq As FrequencyEnum
    Dim blnEberron As Boolean
    Dim i As Long
    
    strHeader = BoldOn & UCase$(GetSchoolName(penSchool)) & " TIER " & plngTier & BoldOff
    If penSchool <> seCultural Or plngTier > 3 Then
        If plngTier = 6 Then strLevel = "26+" Else strLevel = (plngTier * 5) - 4 & " to " & plngTier * 5
        strHeader = strHeader & " " & ColorDim & "(Quest level " & strLevel & ")" & ColorOff
    End If
    OutputLine strHeader, 2
    For enFreq = feCommon To feRare
        If OutputFreq(penSchool, plngTier, enFreq) Then blnEberron = True
    Next
    OutputLine vbNullString
    If blnEberron Then OutputLine ColorRed & "(Eberron only. Not found in Forgotten Realms)" & ColorOff, 2
End Sub

Private Function OutputFreq(penSchool As SchoolEnum, plngTier As Long, penFreq As FrequencyEnum) As Boolean
    Dim strName As String
    Dim strList As String
    Dim lngMaterial As Long
    Dim i As Long
    
    With mtypInfo.School(penSchool).Tier(plngTier).Freq(penFreq)
        strName = ColorDim & GetFrequencyText(penFreq) & ": " & ColorOff
        For i = 1 To .Materials
            lngMaterial = SeekMaterial(.Material(i))
            If lngMaterial Then
                If Len(strList) Then strList = strList & ", "
                If db.Material(lngMaterial).Eberron Then
                    OutputFreq = True
                    strList = strList & ColorRed & .Material(i) & ColorOff
                Else
                    strList = strList & .Material(i)
                End If
            End If
        Next
    End With
    OutputLine strName & strList
End Function

Private Sub OutputFarms(penSchool As SchoolEnum, plngTier As Long)
    Dim typFarm() As MaterialFarmType
    Dim lngFarms As Long
    Dim strLine As String
    Dim i As Long
    
    GetOutputFarms penSchool, plngTier, typFarm, lngFarms
    For i = 1 To lngFarms
        With typFarm(i)
            strLine = "[url=" & .Farm.Wiki & "]" & .Farm.Farm & "[/url]"
            If .Farm.TreasureBag Then
                OutputLine strLine
                OutputCultureFarm .Farm.Notes
            Else
                strLine = strLine & " (" & .Difficulty & ")"
                If Len(.Farm.Video) Then strLine = strLine & " [url=" & .Farm.Video & "]Video[/url]"
                OutputLine strLine
                ShowFarmStats .Farm, penSchool
                OutputLine BoldOn & "Need: " & BoldOff & .Farm.Need
                OutputLine BoldOn & "Fight: " & BoldOff & .Farm.Fight
                OutputLine BoldOn & "Notes: " & BoldOff & .Farm.Notes, 2
            End If
        End With
    Next
End Sub

Private Sub OutputCultureFarm(pstrNotes As String)
    Dim strNote() As String
    Dim strToken() As String
    Dim lngPos As Long
    Dim strLine As String
    Dim strLink As String
    Dim strExtra As String
    Dim i As Long
    
    strNote = Split(pstrNotes, vbNewLine)
    For i = 0 To UBound(strNote)
        strToken = Split(strNote(i), "|")
        If UBound(strToken) <> 2 Then ReDim Preserve strToken(2)
        If Len(strToken(1)) = 0 Then strToken(1) = MakeWiki(strToken(0)) Else strToken(1) = MakeWiki(strToken(1))
        strLink = " ([url=" & strToken(1) & "]link[/url])"
        If Len(strToken(2)) Then strExtra = " " & strToken(2) Else strExtra = vbNullString
        strLine = "- " & strToken(0) & strLink & strExtra
        OutputLine strLine
    Next
    OutputLine vbNullString
End Sub

Private Sub ShowFarmStats(ptypFarm As FarmType, penSchool As SchoolEnum)
    Dim strTime As String
    Dim strNodes As String
    Dim dblNodes As Double
    Dim lngRate As Long
    Dim strRate As String
    Dim strClose As String
    
    Select Case penSchool
        Case seArcane
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seLore
            NodeCount strNodes, ptypFarm, seNatural
        Case seCultural, seAny
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seLore
            NodeCount strNodes, ptypFarm, seNatural
        Case seLore
            NodeCount strNodes, ptypFarm, seLore
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seNatural
        Case seNatural
            NodeCount strNodes, ptypFarm, seNatural
            NodeCount strNodes, ptypFarm, seAny
            NodeCount strNodes, ptypFarm, seArcane
            NodeCount strNodes, ptypFarm, seLore
    End Select
    With ptypFarm
        Select Case penSchool
            Case seArcane: dblNodes = .Arcane
            Case seLore: dblNodes = .Lore
            Case seNatural: dblNodes = .Natural
        End Select
        dblNodes = dblNodes + (.Any * db.Backpack(penSchool))
    End With
    If ptypFarm.Seconds > 0 Then strTime = " in " & SecondsToTime(ptypFarm.Seconds)
    If dblNodes > 0 Then
        lngRate = Int((ptypFarm.Seconds / dblNodes) + 0.5)
        strRate = SecondsToTime(lngRate) & " to farm one " & GetSchoolName(penSchool) & " ("
        strClose = ")"
    End If
    OutputLine strRate & strNodes & strTime & strClose
End Sub

Private Sub NodeCount(pstrNodes As String, ptypFarm As FarmType, penSchool As SchoolEnum)
    Dim strNew As String
    
    Select Case penSchool
        Case seAny: If ptypFarm.Any > 0 Then strNew = ptypFarm.Any & " Any"
        Case seArcane: If ptypFarm.Arcane > 0 Then strNew = ptypFarm.Arcane & " Arcane"
        Case seLore: If ptypFarm.Lore > 0 Then strNew = ptypFarm.Lore & " Lore"
        Case seNatural: If ptypFarm.Natural > 0 Then strNew = ptypFarm.Natural & " Natural"
    End Select
    If Len(strNew) = 0 Then Exit Sub
    If Len(pstrNodes) Then pstrNodes = pstrNodes & ", "
    pstrNodes = pstrNodes & strNew
End Sub

Private Sub OutputLine(pstrNew As String, Optional plngNewLines As Long = 1)
    If mlngLines > UBound(mstrLine) Then ReDim Preserve mstrLine((mlngLines * 3) \ 2)
    mstrLine(mlngLines) = pstrNew
    mlngLines = mlngLines + plngNewLines
End Sub

