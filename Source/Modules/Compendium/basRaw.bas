Attribute VB_Name = "basRaw"
Option Explicit

Public Sub RawWilderness()
    Dim strFile As String
    Dim strLine() As String
    Dim lngLine As Long
    Dim strToken() As String
    Dim i As Long
    
    Erase db.Area
    strFile = DataPath() & "Wilderness(Raw).txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strLine = Split(xp.File.LoadToString(strFile), vbNewLine)
    ReDim db.Area(1 To UBound(strLine) + 1)
    db.Areas = 0
    For i = 0 To UBound(strLine)
        strToken = Split(strLine(i), vbTab)
        If UBound(strToken) = 3 Then
            db.Areas = db.Areas + 1
            With db.Area(db.Areas)
                .Area = Left$(strToken(0), Len(strToken(0)) - 6)
                ParseLevels strToken(1), .Lowest, .Highest
                .Explorer = ConvertNumber(strToken(2))
                .Pack = PackName(strToken(3))
            End With
        End If
    Next
    ReDim strLine(db.Areas * 6)
    ReDim Preserve db.Area(1 To db.Areas)
    For i = 1 To db.Areas
        With db.Area(i)
            strLine(lngLine) = "Area: " & .Area
            strLine(lngLine + 1) = "Low: " & .Lowest
            strLine(lngLine + 2) = "High: " & .Highest
            lngLine = lngLine + 3
            If .Explorer > 0 Then
                strLine(lngLine) = "Explorer: " & .Explorer
                lngLine = lngLine + 1
            End If
            If Len(.Pack) Then
                strLine(lngLine) = "Pack: " & .Pack
                lngLine = lngLine + 1
            End If
            lngLine = lngLine + 1
        End With
    Next
    ReDim Preserve strLine(lngLine)
    strFile = DataPath() & "Wilderness.txt"
    If xp.File.Exists(strFile) Then xp.File.Delete strFile
    xp.File.SaveStringAs strFile, Join(strLine, vbNewLine)
End Sub

Private Sub ParseLevels(ByVal pstrRaw As String, plngLow As Long, plngHigh As Long)
    Dim lngPos As Long
    Dim strToken() As String
    
    lngPos = InStr(pstrRaw, "(")
    pstrRaw = Mid$(pstrRaw, lngPos + 1)
    pstrRaw = Left$(pstrRaw, Len(pstrRaw) - 1)
    strToken = Split(pstrRaw, "-")
    If UBound(strToken) = 1 Then
        plngLow = Val(strToken(0))
        plngHigh = Val(strToken(1))
    End If
End Sub

Private Function ConvertNumber(ByVal pstrRaw As String) As Long
    pstrRaw = Replace(pstrRaw, Chr(34), vbNullString)
    pstrRaw = Replace(pstrRaw, ",", vbNullString)
    ConvertNumber = Val(pstrRaw)
End Function

Private Function PackName(pstrRaw As String) As String
    Dim lngPos As Long
    
    lngPos = InStr(pstrRaw, "(")
    If lngPos = 0 Then Exit Function
    PackName = Trim$(Left$(pstrRaw, lngPos - 1))
End Function
