Attribute VB_Name = "basTomes"
Option Explicit

Public Type TomeScheduleType
    Max As Long
    Level() As Long
End Type
    
' Controls
Public Type TomesType
    Stat As TomeScheduleType
    Skill As TomeScheduleType
    RacialAPMax As Long
    UniversalAPMax As Long
    PowerMax As Long
    RRMax As Long
    PermDestinyPointMin As Long  '12  https://ddowiki.com/page/Update_51_Release_Notes#Fate_Points
    PermDestinyPointMax As Long  '18
    FateTomeMax As Long
    DestinyTomeMax As Long
End Type

Public tomes As TomesType

Public Sub LoadTomeData()
    Dim strFile As String
    Dim strRaw As String
    Dim strLine() As String
    Dim strToken() As String
    Dim i As Long
    
    SetDefaults
    ' Load file
    strFile = App.Path & "\Data\Tomes.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strLine = Split(strRaw, vbNewLine)
    ' Parse file
    For i = 0 To UBound(strLine)
        If InStr(strLine(i), ":") Then
            strToken = Split(strLine(i), ":")
            With tomes
                Select Case LCase$(strToken(0))
                    Case "stat": ParseTomeSchedule .Stat, strToken(1)
                    Case "skill": ParseTomeSchedule .Skill, strToken(1)
                    Case "racialapmax": SetTomeMax .RacialAPMax, strToken(1)
                    Case "powermax": SetTomeMax .PowerMax, strToken(1)
                    Case "prr/mrrmax": SetTomeMax .RRMax, strToken(1)
                    Case "universalapmax": SetTomeMax .UniversalAPMax, strToken(1)
                    Case "permdestinypointmax": SetTomeMax .PermDestinyPointMax, strToken(1)
                    Case "permdestinypointmin": SetTomeMax .PermDestinyPointMin, strToken(1)
                    Case "fatetomemax": SetTomeMax .FateTomeMax, strToken(1)
                    Case "destinytomemax": SetTomeMax .DestinyTomeMax, strToken(1)
                End Select
            End With
        End If
    Next
End Sub

Private Sub SetDefaults()
    Dim typBlank As TomesType
    
    ' Set defaults
    tomes = typBlank
    With tomes
        ParseTomeSchedule .Stat, "8, 1, 1, 3, 7, 11, 15, 19, 22"
        ParseTomeSchedule .Skill, "5, 1, 1, 3, 7, 11"
        .RacialAPMax = 2
        .PowerMax = 4
        .RRMax = 4
        .UniversalAPMax = 1
        .DestinyTomeMax = 2
        .FateTomeMax = 3
        .PermDestinyPointMax = 18
    End With
End Sub

Private Sub ParseTomeSchedule(ptypTome As TomeScheduleType, pstrRaw As String)
    Dim strToken() As String
    Dim lngMax As Long
    Dim i As Long
    
    ' Format is
    '  <max tome level>
    '  list of level that you can use the tome.
    If InStr(pstrRaw, ",") = 0 Then Exit Sub
    strToken = Split(pstrRaw, ",")
    With ptypTome
        .Max = val(strToken(0))
        ReDim .Level(.Max)
        For i = 1 To UBound(strToken)
            If i <= .Max Then .Level(i) = val(strToken(i))
        Next
    End With
End Sub

Private Sub SetTomeMax(plngTome As Long, pstrRaw As String)
    Dim lngValue As Long
    
    lngValue = val(pstrRaw)
    If lngValue Then plngTome = lngValue
End Sub
