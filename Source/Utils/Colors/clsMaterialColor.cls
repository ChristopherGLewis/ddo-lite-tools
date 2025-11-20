VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaterialColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Material Design color schemes: https://www.google.com/design/spec/style/color.html#
' Key difference is that this module sorts them from darkest to lightest to keep consistent with adjusting RGB values
Option Explicit

' Color values, 0 to 18
Public Enum MaterialColorEnum
    mceAmber
    mceBlue
    mceBlueGrey
    mceBrown
    mceCyan
    mceDeepOrange
    mceDeepPurple
    mceGreen
    mceGrey
    mceIndigo
    mceLightBlue
    mceLightGreen
    mceLime
    mceOrange
    mcePink
    mcePurple
    mceRed
    mceTeal
    mceYellow
End Enum

Public Enum MaterialColorTextEnum
    mcteBlack = vbBlack
    mcteWhite = vbWhite
End Enum

Private Type MaterialColorPrimaryType
    Value As Long
    TextColor As MaterialColorTextEnum
End Type

Private Type MaterialColorType
    ColorName As String
    Color(9) As MaterialColorPrimaryType ' There are 10 different color gradients, 0-9
    FirstBlackText As Long ' Internal use only. For Color(Index), If Index>=FirstBlackText Then text= black Else text = white
End Type

Private mtypColor(18) As MaterialColorType

' Returns color value
Public Function GetColor(penMaterialColor As MaterialColorEnum, plngValue As Long) As Long
    GetColor = mtypColor(penMaterialColor).Color(plngValue).Value
End Function

' Returns color value 1-5 steps away from specified value in longest available direction (eg: 5 steps away from 9 is 4, from 3 is 8)
Public Function StepColor(penMaterialColor As MaterialColorEnum, plngValue As Long, plngStep As Long) As Long
    Dim lngComplement As Long
    
    If plngValue < 5 Then lngComplement = plngStep Else lngComplement = -plngStep
    StepColor = mtypColor(penMaterialColor).Color(plngValue + lngComplement).Value
End Function

Public Function GetTextColor(penMaterialColor As MaterialColorEnum, plngValue As Long) As MaterialColorTextEnum
    GetTextColor = mtypColor(penMaterialColor).Color(plngValue).TextColor
End Function

Public Function GetColorName(penMaterialColor As MaterialColorEnum) As String
    GetColorName = mtypColor(penMaterialColor).ColorName
End Function

Public Function GetColorIndex(pstrColorName As String) As MaterialColorEnum
    Dim i As Long
    
    For i = 0 To 18
        If mtypColor(i).ColorName = pstrColorName Then
            GetColorIndex = i
            Exit For
        End If
    Next
End Function

' 50, 100, 200, 300, ..., 900
Public Function GetValueName(plngValue As Long) As Long
    If plngValue = 0 Then GetValueName = 50 Else GetValueName = plngValue * 100
End Function

Public Function GetColorHex(penMaterialColor As MaterialColorEnum, plngValue As Long) As String
    GetColorHex = ColorToHex(GetColor(penMaterialColor, plngValue))
End Function


' ************* INITIALIZE *************


Private Sub Class_Initialize()
    Dim i As Long

    AddMaterialColor i, "Amber,FF6F00,FF8F00,FFA000,FFB300,FFC107,FFCA28,FFD54F,FFE082,FFECB3,FFF8E1,0"
    AddMaterialColor i, "Blue,0D47A1,1565C0,1976D2,1E88E5,2196F3,42A5F5,64B5F6,90CAF9,BBDEFB,E3F2FD,5"
    AddMaterialColor i, "Blue Grey,263238,37474F,455A64,546E7A,607D8B,78909C,90A4AE,B0BEC5,CFD8DC,ECEFF1,6"
    AddMaterialColor i, "Brown,3E2723,4E342E,5D4037,6D4C41,795548,8D6E63,A1887F,BCAAA4,D7CCC8,EFEBE9,7"
    AddMaterialColor i, "Cyan,006064,00838F,0097A7,00ACC1,00BCD4,26C6DA,4DD0E1,80DEEA,B2EBF2,E0F7FA,3"
    AddMaterialColor i, "Deep Orange,BF360C,D84315,E64A19,F4511E,FF5722,FF7043,FF8A65,FFAB91,FFCCBC,FBE9E7,5"
    AddMaterialColor i, "Deep Purple,311B92,4527A0,512DA8,5E35B1,673AB7,7E57C2,9575CD,B39DDB,D1C4E9,EDE7F6,7"
    AddMaterialColor i, "Green,1B5E20,2E7D32,388E3C,43A047,4CAF50,66BB6A,81C784,A5D6A7,C8E6C9,E8F5E9,4"
    AddMaterialColor i, "Grey,212121,424242,616161,757575,9E9E9E,BDBDBD,E0E0E0,EEEEEE,F5F5F5,FAFAFA,4"
    AddMaterialColor i, "Indigo,1A237E,283593,303F9F,3949AB,3F51B5,5C6BC0,7986CB,9FA8DA,C5CAE9,E8EAF6,7"
    AddMaterialColor i, "Light Blue,01579B,0277BD,0288D1,039BE5,03A9F4,29B6F6,4FC3F7,81D4FA,B3E5FC,E1F5FE,4"
    AddMaterialColor i, "Light Green,33691E,558B2F,689F38,7CB342,8BC34A,9CCC65,AED581,C5E1A5,DCEDC8,F1F8E9,3"
    AddMaterialColor i, "Lime,827717,9E9D24,AFB42B,C0CA33,CDDC39,D4E157,DCE775,E6EE9C,F0F4C3,F9FBE7,1"
    AddMaterialColor i, "Orange,E65100,EF6C00,F57C00,FB8C00,FF9800,FFA726,FFB74D,FFCC80,FFE0B2,FFF3E0,2"
    AddMaterialColor i, "Pink,880E4F,AD1457,C2185B,D81B60,E91E63,EC407A,F06292,F48FB1,F8BBD0,FCE4EC,7"
    AddMaterialColor i, "Purple,4A148C,6A1B9A,7B1FA2,8E24AA,9C27B0,AB47BC,BA68C8,CE93D8,E1BEE7,F3E5F5,7"
    AddMaterialColor i, "Red,B71C1C,C62828,D32F2F,E53935,F44336,EF5350,E57373,EF9A9A,FFCDD2,FFEBEE,6"
    AddMaterialColor i, "Teal,004D40,00695C,00796B,00897B,009688,26A69A,4DB6AC,80CBC4,B2DFDB,E0F2F1,5"
    AddMaterialColor i, "Yellow,F57F17,F9A825,FBC02D,FDD835,FFEB3B,FFEE58,FFF176,FFF59D,FFF9C4,FFFDE7,0"
End Sub

Private Sub AddMaterialColor(plngIndex As Long, pstrRaw As String)
    Dim strToken() As String
    Dim i As Long
    
    With mtypColor(plngIndex)
        strToken = Split(pstrRaw, ",")
        .ColorName = strToken(0)
        .FirstBlackText = CLng(strToken(11))
        For i = 0 To 9
            If i >= .FirstBlackText Then .Color(i).TextColor = mcteBlack Else .Color(i).TextColor = mcteWhite
            .Color(i).Value = HexToColor(strToken(i + 1))
        Next
    End With
    plngIndex = plngIndex + 1
End Sub


' ************* HEX *************


Private Function HexToColor(ByVal pstrHex As String) As Long
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    lngRed = CLng("&H" & Left$(pstrHex, 2))
    lngGreen = CLng("&H" & Mid$(pstrHex, 3, 2))
    lngBlue = CLng("&H" & Right$(pstrHex, 2))
    HexToColor = RGB(lngRed, lngGreen, lngBlue)
End Function

Private Function ColorToHex(ByVal plngColor As Long) As String
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    ColorToRGB plngColor, lngRed, lngGreen, lngBlue
    ColorToHex = HexColorValue(lngRed) & HexColorValue(lngGreen) & HexColorValue(lngBlue)
End Function

Private Sub ColorToRGB(plngColor As Long, plngRed As Long, plngGreen As Long, plngBlue As Long)
    plngRed = &HFF& And plngColor
    plngGreen = (&HFF00& And plngColor) \ 256
    plngBlue = (&HFF0000 And plngColor) \ 65536
End Sub

Private Function HexColorValue(plngValue As Long) As String
    Dim strReturn As String
    
    strReturn = Hex(plngValue)
    If Len(strReturn) < 2 Then strReturn = "0" & strReturn
    HexColorValue = strReturn
End Function

