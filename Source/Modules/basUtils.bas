Attribute VB_Name = "basUtils"
' Written by Ellis Dee
' Generic utility functions.
Option Explicit


' Remove "&" and "..." from menu captions for easier identification when a menu fires
Public Function StripMenuChars(ByVal pstrCaption As String) As String
    If Right(pstrCaption, 3) = "..." Then pstrCaption = Left(pstrCaption, Len(pstrCaption) - 3)
    If InStr(pstrCaption, "&&") Then pstrCaption = Replace(pstrCaption, "&&", "<Amp>")
    If InStr(pstrCaption, "&") Then pstrCaption = Replace(pstrCaption, "&", vbNullString)
    If InStr(pstrCaption, "<Amp>") Then pstrCaption = Replace(pstrCaption, "<Amp>", "&")
    StripMenuChars = pstrCaption
End Function

' Remove leading and trailing spaces and linefeeds
Public Function CleanText(pstrText As String)
    pstrText = Trim$(pstrText)
    Do While Left$(pstrText, 2) = vbNewLine
        pstrText = Trim$(Mid$(pstrText, 3))
    Loop
    Do While Right$(pstrText, 2) = vbNewLine
        pstrText = Trim$(Left$(pstrText, Len(pstrText) - 2))
    Loop
End Function

Public Function TrimAll(pstrText As String)
    TrimAll = Replace(Trim$(pstrText), vbTab, "")
End Function

' Set of functions to return a part of a filespec ("C:\Folder\File.ext")
Public Function GetPathFromFilespec(pstrFile As String) As String
    GetPathFromFilespec = Left(pstrFile, InStrRev(pstrFile, "\") - 1)
End Function

Public Function GetFileFromFilespec(pstrFile As String) As String
    GetFileFromFilespec = Mid(pstrFile, InStrRev(pstrFile, "\") + 1)
End Function

Public Function GetExtFromFilespec(pstrFile As String) As String
    GetExtFromFilespec = LCase(Mid(pstrFile, InStrRev(pstrFile, ".") + 1))
End Function

Public Function GetNameFromFilespec(ByVal pstrFile As String) As String
    If InStr(pstrFile, "\") Then pstrFile = Mid$(pstrFile, InStrRev(pstrFile, "\") + 1)
    If InStr(pstrFile, ".") Then pstrFile = Left$(pstrFile, InStrRev(pstrFile, ".") - 1)
    GetNameFromFilespec = pstrFile
End Function

Public Function GetForm(pfrm As Form, pstrFormName As String) As Boolean
    For Each pfrm In Forms
        If pfrm.Name = pstrFormName Then
            GetForm = True
            Exit Function
        End If
    Next
    Set pfrm = Nothing
End Function

Public Function Ask(pstrMessage As String, Optional pblnDefaultAnswer As Boolean = False) As Boolean
    Dim enStyle As VbMsgBoxStyle
    
    If cfg.Confirm Then
        enStyle = vbQuestion + vbYesNo
        If Not pblnDefaultAnswer Then enStyle = enStyle + vbDefaultButton2
        If MsgBox(pstrMessage, enStyle, "Notice") = vbYes Then Ask = True
    Else
        Ask = True
    End If
End Function

Public Function AskAlways(pstrMessage As String, Optional pblnDefaultAnswer As Boolean = False) As Boolean
    Dim enStyle As VbMsgBoxStyle
    
    enStyle = vbQuestion + vbYesNo
    If Not pblnDefaultAnswer Then enStyle = enStyle + vbDefaultButton2
    If MsgBox(pstrMessage, enStyle, "Notice") = vbYes Then AskAlways = True
End Function

Public Function AskCancel(pstrMessage As String, Optional pblnAlwaysAsk As Boolean = False) As VbMsgBoxResult
    Dim enStyle As VbMsgBoxStyle
    
    If cfg.Confirm Or pblnAlwaysAsk Then
        enStyle = vbQuestion + vbYesNoCancel + vbDefaultButton2
        AskCancel = MsgBox(pstrMessage, enStyle, "Notice")
    Else
        AskCancel = vbYes
    End If
End Function

Public Sub Notice(pstrMessage As String, Optional pblnExclamation As Boolean = False)
    Dim enStyle As VbMsgBoxStyle
    
    If pblnExclamation Then enStyle = vbExclamation Else enStyle = vbInformation
    MsgBox pstrMessage, enStyle, "Notice"
End Sub

Public Function UnderConstruction()
    MsgBox "Feature under construction.", vbInformation, "Sorry..."
End Function
