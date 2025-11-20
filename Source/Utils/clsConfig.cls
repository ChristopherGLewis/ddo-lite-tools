VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Written by Ellis Dee
' Super stripped down version for utils
Option Explicit

Public Enum ColorGroupEnum
    cgeWorkspace
    cgeControls
    cgeNavigation
    cgeDropSlots
    cgeOutput
End Enum

Public Enum ColorValueEnum
    cveBackground
    cveBackHighlight
    cveBackError
    cveBackRelated
    cveText
    cveTextError
    cveTextDim
    cveTextLink
    cveBorderInterior
    cveBorderExterior
    cveBorderHighlight
    cveLightGray
    cveRed
    cveYellow
    cveBlue
    cvePurple
    cveGreen
    cveOrange
    cveColorValues
End Enum

Private mlngColor(4, 17) As Long

Private Sub Class_Initialize()
    InitColors cveBackground, 0, 3947580, 6316128, 6316128, 2236962
    InitColors cveBackHighlight, 19532, 32125, 24672, 32125, 19275
    InitColors cveBackError, 76, 4934575, 96, 4934575, 75
    InitColors cveBackRelated, 4980736, 11488075, 6291456, 11488075, 4915200
    InitColors cveText, 12632256, 15461355, 14803425, 16777215, 16777215
    InitColors cveTextError, 10132207, 10132207, 12632319, 12632319, 11579647
    InitColors cveTextDim, 8553090, 11579568, 12895428, 11579568, 12632256
    InitColors cveTextLink, 16769480, 16769480, 16777215, 16762026, 16769480
    InitColors cveBorderInterior, 8421504, 7039851, 3947580, 2171169, 8421504
    InitColors cveBorderExterior, 12632256, 11579568, 3947580, 11513775, 12632256
    InitColors cveBorderHighlight, 10526880, 16777215, 3947580, 16777215, 16777215
    InitColors cveLightGray, 10395294, 12434877, 10395294, 10395294, 12434877
    InitColors cveRed, 10132207, 10132207, 10132207, 10132207, 10132207
    InitColors cveYellow, 7795199, 7795199, 3927039, 3927039, 7795199
    InitColors cveBlue, 16370320, 16370320, 16370320, 16370320, 16370320
    InitColors cvePurple, 14193614, 14193614, 14193614, 14193614, 14193614
    InitColors cveGreen, 8701825, 8701825, 10999461, 10999461, 8701825
    InitColors cveOrange, 2533375, 5093375, 2533375, 2533375, 5093375
End Sub

Private Sub InitColors(penColor As ColorValueEnum, plngWrk As Long, plngCtl As Long, plngNav As Long, plngDrp As Long, plngOut As Long)
    mlngColor(cgeWorkspace, penColor) = plngWrk
    mlngColor(cgeControls, penColor) = plngCtl
    mlngColor(cgeNavigation, penColor) = plngNav
    mlngColor(cgeDropSlots, penColor) = plngDrp
    mlngColor(cgeOutput, penColor) = plngOut
End Sub

Public Sub RefreshColors(pfrm As Form)
    Dim enGroup As ColorGroupEnum
    Dim ctl As Control

    Select Case pfrm.Name
        Case "frmColors"
            pfrm.RefreshColors
            Exit Sub
        Case "frmMain"
    End Select
    pfrm.BackColor = GetColor(cgeWorkspace, cveBackground)
    For Each ctl In pfrm.Controls
        Select Case ctl.Tag
            Case "nav": ApplyColors ctl, cgeNavigation
            Case "ctl": ApplyColors ctl, cgeControls
            Case "out": ApplyColors ctl, cgeOutput
            Case "wrk": ApplyColors ctl, cgeWorkspace
            Case "tip"
            Case Else
                Select Case TypeName(ctl)
                    Case "Frame", "Label", "Line", "Shape"
                        ApplyColors ctl, cgeWorkspace
                    Case "CheckBox", "ComboBox", "ListBox", "TextBox", "userSpinner", "userTextbox"
                        ApplyColors ctl, cgeControls
                    Case "userHeader"
                        ApplyColors ctl, cgeNavigation
                    Case "userCheckBox", "userDetails", "userList", "userQuests", "userStats"
                        ctl.RefreshColors
                    Case "userInfo"
                        ctl.TitleForeColor = GetColor(cgeWorkspace, cveText)
                        ctl.TitleBackColor = GetColor(cgeWorkspace, cveBackground)
                        ctl.TextColor = GetColor(cgeControls, cveText)
                        ctl.ErrorColor = GetColor(cgeControls, cveTextError)
                        ctl.LinkColor = GetColor(cgeControls, cveTextLink)
                        ctl.BackColor = GetColor(cgeControls, cveBackground)
                    Case "userTab"
                        ctl.BackColor = GetColor(cgeWorkspace, cveBackground)
                        ctl.TextActiveColor = GetColor(cgeControls, cveText)
                        ctl.TabActiveColor = GetColor(cgeControls, cveBackground)
                        ctl.TextInactiveColor = GetColor(cgeWorkspace, cveText)
                        ctl.TabInactiveColor = GetColor(cgeWorkspace, cveBackground)
                    Case "PictureBox"
                        If pfrm.Name = "frmColors" And ctl.Name = "picColor" Then
'                            ctl.BackColor = GetColor(frmColors.ColorGroup, ctl.Index)
                        ElseIf ctl.Name = "picGrid" Then
                            ctl.BackColor = GetColor(cgeWorkspace, cveBackground)
                            ctl.FillColor = GetColor(cgeControls, cveBackground)
                        ElseIf ctl.Name = "picPicker" Then
'                            ctl.BackColor = glngActiveColor
                        Else
                            ApplyColors ctl, cgeWorkspace
                        End If
                End Select
        End Select
        Set ctl = Nothing
    Next
    If App.Title = "Gear Planner" Then pfrm.RefreshColors
    Select Case pfrm.Name
        Case "frmSkills": pfrm.DrawGrid
        Case "frmHelp": pfrm.DrawText
        Case "frmOverview", "frmEnhancements", "frmFormat", "frmExport", "frmColorFile": pfrm.RefreshColors
    End Select
End Sub

Public Function GetColor(ByVal penGroup As ColorGroupEnum, ByVal penColor As ColorValueEnum) As Long
    If penColor = -1 Then GetColor = -1 Else GetColor = mlngColor(penGroup, penColor)
End Function

Public Function ApplyColors(pctl As Control, penGroup As ColorGroupEnum)
    Select Case TypeName(pctl)
        Case "Label"
            If Left$(pctl.Name, 3) = "lnk" Then
                pctl.ForeColor = mlngColor(penGroup, cveTextLink)
            Else
                pctl.ForeColor = mlngColor(penGroup, cveText)
            End If
            pctl.BackColor = mlngColor(penGroup, cveBackground)
        Case "TextBox", "ListBox", "ComboBox", "PictureBox", "Frame", "CheckBox", "OptionButton", "userTextbox"
            pctl.ForeColor = mlngColor(penGroup, cveText)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
        Case "userHeader"
            pctl.TextColor = mlngColor(penGroup, cveText)
            pctl.LinkColor = mlngColor(penGroup, cveTextLink)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.BorderColor = mlngColor(penGroup, cveBorderExterior)
        Case "userCheckBox"
            pctl.CustomColors penGroup, 0
        Case "userSpinner"
            pctl.ForeColor = mlngColor(penGroup, cveText)
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.DisabledColor = mlngColor(penGroup, cveTextDim)
            pctl.BorderColor = mlngColor(penGroup, cveBorderExterior)
            pctl.BorderInterior = mlngColor(penGroup, cveBorderInterior)
        Case "Shape"
            pctl.BackColor = mlngColor(penGroup, cveBackground)
            pctl.BorderColor = mlngColor(penGroup, cveBorderHighlight)
        Case "Line"
            If pctl.Name = "linHeader" Then pctl.BorderColor = GetColor(penGroup, cveText) Else pctl.BorderColor = GetColor(penGroup, cveBorderExterior)
    End Select
End Function
