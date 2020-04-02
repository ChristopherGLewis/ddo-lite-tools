Attribute VB_Name = "basResource"
Option Explicit

Public Enum ArrowEnum
    aeDelete
    aeUp
    aeDown
End Enum

Public Enum ArrowStateEnum
    aseEnabled
    aseDisabled
    asePressed
End Enum

Public Sub EnableArrow(pimg As Image, pblnEnabled As Boolean)
    Dim enState As ArrowStateEnum
    
    If pblnEnabled Then enState = aseEnabled Else enState = aseDisabled
    SetArrowIcon pimg, enState
    pimg.Enabled = pblnEnabled
End Sub

Public Sub SetArrowIcon(pimg As Image, penState As ArrowStateEnum)
    Dim strResourceID As String
    
    strResourceID = "ARW"
    Select Case pimg.Index
        Case aeUp: strResourceID = strResourceID & "UP"
        Case aeDown: strResourceID = strResourceID & "DN"
        Case aeDelete: strResourceID = strResourceID & "DEL"
    End Select
    Select Case penState
        Case aseEnabled
        Case aseDisabled: strResourceID = strResourceID & "DIS"
        Case asePressed: strResourceID = strResourceID & "PRESS"
    End Select
    pimg.Picture = LoadResPicture(strResourceID, vbResBitmap)
End Sub

