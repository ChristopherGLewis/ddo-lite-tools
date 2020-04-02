Attribute VB_Name = "basFonts"
Option Explicit

Private Type FontType
    Loaded As Boolean
    Checked As Boolean
    RawFonts As Long
    Raw() As String
    Found As Long
    Valid() As String
    Sizes As Long
    Size() As String
End Type

Private fnt As FontType


' ************* FONT NAMES *************


Public Function GetFontList(pic As PictureBox) As String()
    If Not fnt.Loaded Then LoadFonts
    If Not fnt.Checked Then CheckFonts pic
    GetFontList = fnt.Valid
End Function

Private Sub LoadFonts()
    Dim typBlank As FontType
    Dim strFile As String
    Dim strFont() As String
    Dim lngFonts As Long
    Dim strTemp As String
    Dim i As Long
    
    fnt = typBlank
    strFile = App.Path & "\Data\Fonts.txt"
    If Not xp.File.Exists(strFile) Then Exit Sub
    strFont = Split(xp.File.LoadToString(strFile), vbNewLine)
    lngFonts = UBound(strFont)
    If lngFonts = 0 Then Exit Sub
    With fnt
        ReDim .Raw(1 To lngFonts + 1)
        For i = 0 To UBound(strFont)
            strTemp = Trim$(strFont(i))
            If Len(strTemp) Then
                If Left$(strTemp, 1) <> ";" Then
                    .RawFonts = .RawFonts + 1
                    .Raw(.RawFonts) = strTemp
                End If
            End If
        Next
        If .RawFonts = 0 Then
            Erase .Raw
        Else
            ReDim Preserve .Raw(1 To .RawFonts)
            .Loaded = True
        End If
    End With
End Sub

Private Sub CheckFonts(pic As PictureBox)
    Dim strList() As String
    Dim strFont As String
    Dim i As Long
    Dim j As Long
    
    If fnt.RawFonts = 0 Then Exit Sub
    With fnt
        .Found = 0
        ReDim .Valid(.RawFonts)
        For i = 1 To fnt.RawFonts
            strList = Split(fnt.Raw(i), ",")
            For j = 0 To UBound(strList)
                strFont = Trim$(strList(j))
                If AddFont(strFont, pic) Then Exit For
            Next
        Next
        ReDim Preserve .Valid(.Found)
        If .Found Then .Checked = True
    End With
End Sub

Private Function AddFont(pstrFontName As String, pic As PictureBox) As Boolean
    On Error Resume Next
    pic.FontName = pstrFontName
    If Err.Number Then Exit Function
    On Error GoTo 0
    With fnt
        .Found = .Found + 1
        .Valid(.Found) = pstrFontName
    End With
    AddFont = True
End Function


' ************* FONT SIZES *************


Public Function FontSizeToString(pdblFontSize As Double) As String
    Dim strReturn As String
    
    strReturn = Format(pdblFontSize, "0.0")
    If Right$(strReturn, 1) = "0" Then strReturn = Left$(strReturn, Len(strReturn) - 2)
    FontSizeToString = strReturn
End Function

Public Function GetFontSizes(pstrFontName As String, pic As PictureBox) As String()
    Dim dblCurrent As Double
    Dim s As Double
    
    ReDim fnt.Size(31)
    fnt.Sizes = 0
    On Error Resume Next
    pic.FontName = pstrFontName
    If Err.Number Then Exit Function
    On Error GoTo 0
    With fnt
        For s = 6 To 14.5 Step 0.25
            pic.FontSize = s
            If dblCurrent <> pic.FontSize Then
                dblCurrent = pic.FontSize
                .Sizes = .Sizes + 1
                .Size(.Sizes) = FontSizeToString(dblCurrent)
            End If
        Next
        ReDim Preserve .Size(.Sizes)
        GetFontSizes = .Size
        Erase .Size
        .Sizes = 0
    End With
End Function
