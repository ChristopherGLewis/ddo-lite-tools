Attribute VB_Name = "basHelp"
' Written by Ellis Dee
Option Explicit

Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Public Type WordType
    Text As String
    Code As String
    Bold As Boolean
    Italics As Boolean
    Underline As Boolean
    ErrorText As Boolean
End Type

Public Type ParagraphType
    Raw As String
    Word() As WordType
    Words As Long
    Indent As Long
End Type

Public Type HelpItemType
    HelpCode As String
    Paragraph() As ParagraphType
    Paragraphs As Long
End Type

Public Type HelpType
    HelpItem() As HelpItemType
    HelpItems As Long
    Current As Long
End Type

Public help As HelpType

Public Sub InitHelp()
    Dim strFile As String
    Dim strRaw As String
    Dim strHelp() As String
    Dim strParagraph() As String
    Dim lngPos As Long
    Dim i As Long
    Dim j As Long
    
    strFile = DataPath() & "Help.txt"
    If Not FileExists(strFile) Then Exit Sub
    strRaw = LoadToString(strFile)
    strHelp = Split(strRaw, "HelpTopic: ")
    With help
        .HelpItems = UBound(strHelp) - 1
        ReDim .HelpItem(1 To .HelpItems)
        For i = 2 To UBound(strHelp)
            With .HelpItem(i - 1)
                lngPos = InStr(strHelp(i), vbNewLine)
                .HelpCode = Left$(strHelp(i), lngPos - 1)
                strParagraph = Split(Mid$(strHelp(i), lngPos + 2), vbNewLine)
                .Paragraphs = UBound(strParagraph)
                ReDim .Paragraph(.Paragraphs)
                For j = 0 To .Paragraphs
                    ProcessParagraph .Paragraph(j), strParagraph(j)
                Next
            End With
        Next
    End With
End Sub

Private Sub ProcessParagraph(ptypParagraph As ParagraphType, pstrRaw As String)
    Dim strWord() As String
    Dim i As Long
    
    With ptypParagraph
        .Raw = pstrRaw
        If Left$(pstrRaw, 3) = " - " Then .Indent = 1
        If Left$(pstrRaw, 5) = "   - " Then .Indent = 2
        strWord = Split(pstrRaw, " ")
        .Words = UBound(strWord)
        If .Words <> -1 Then
            ReDim .Word(.Words)
            For i = 0 To .Words
                ProcessWord .Word(i), strWord(i)
            Next
        End If
    End With
End Sub

Private Sub ProcessWord(ptypWord As WordType, pstrRaw As String)
    Dim lngOpen As Long
    Dim lngClose As Long
    Dim strCode As String
    Dim strWord As String
    
    ptypWord.Text = pstrRaw
    lngOpen = InStr(pstrRaw, "{")
    If lngOpen = 0 Then Exit Sub
    lngClose = InStr(lngOpen, pstrRaw, "}")
    If lngClose = 0 Then Exit Sub
    strCode = Mid$(pstrRaw, lngOpen + 1, lngClose - lngOpen - 1)
    strWord = Mid$(pstrRaw, lngClose + 1)
    If InStr(strWord, "_") Then strWord = Replace(strWord, "_", " ")
    ptypWord.Text = strWord
    If InStr(strCode, "=") Then
        ptypWord.Code = strCode
    Else
        strCode = LCase$(strCode)
        Select Case strCode
            Case "error"
                ptypWord.ErrorText = True
            Case Else
                If InStr(strCode, "b") Then ptypWord.Bold = True
                If InStr(strCode, "i") Then ptypWord.Italics = True
                If InStr(strCode, "u") Then ptypWord.Underline = True
        End Select
    End If
End Sub

Public Function ShowHelp(pstrHelpCode As String)
    Dim frm As Form
    Dim i As Long
    
    help.Current = 0
    For i = 1 To help.HelpItems
        If help.HelpItem(i).HelpCode = pstrHelpCode Then
            help.Current = i
            Exit For
        End If
    Next
    If help.Current = 0 Then
        MsgBox "Help Code '" & pstrHelpCode & "' not found.", vbInformation, "Notice"
    Else
        If GetForm(frm, "frmMenuEditor") Or GetForm(frm, "frmSagaDetail") Then
            If GetForm(frm, "frmHelp") Then Unload frmHelp
            frmHelp.Show vbModal, Screen.ActiveForm
        Else
            If GetForm(frm, "frmHelp") Then
                frmHelp.DrawText
                If GetForm(frm, "frmMain") Then frm.UpdateWindowMenu
                On Error Resume Next
                frmHelp.SetFocus
            ElseIf GetForm(frm, "frmCompendium") Then
'                If frm.WindowState = vbMinimized Then
                    frmHelp.Show vbModeless, Screen.ActiveForm
'                Else
'                    frmHelp.Show vbModeless, frm
'                End If
            ElseIf pstrHelpCode = "Scaling" Then
                If GetForm(frm, "frmScaling") Then frmHelp.Show vbModeless, frm
            Else
                OpenForm "frmHelp"
            End If
        End If
    End If
End Function


' ************* UTILS *************


'Private Function DebugMode() As Boolean
'    DebugMode = (App.LogMode = 0)
'End Function
'
'Private Sub OpenURL(ByVal URL As String)
'    ShellExecute 0&, "OPEN", URL, vbNullString, vbNullString, vbNormalFocus
'End Sub

Private Function FileExists(ByVal File As String) As Boolean
    FileExists = (PathFileExists(File) = 1)
    If FileExists Then FileExists = (PathIsDirectory(File) = 0)
End Function
'
'Private Function Run(ByVal File As String, Optional ByVal WindowState As Long = 1, Optional ByVal DefaultFolder As String) As Long
'    Dim lngPos As Long
'    Dim lngDesktop As Long
'
'    If DefaultFolder = "" Then
'        lngPos = InStrRev(File, "\")
'        If lngPos > 0 Then DefaultFolder = Left$(File, lngPos - 1)
'    End If
'    lngDesktop = GetDesktopWindow()
'    Run = ShellExecute(lngDesktop, "Open", File, "", DefaultFolder, WindowState) 'SW_SHOW)
'End Function

Private Function LoadToString(File As String) As String
    Dim strReturn As String
    Dim FileNumber As Long

    FileNumber = FreeFile
    Open File For Binary Access Read As #FileNumber
    If LOF(FileNumber) > 0 Then
        strReturn = Space(LOF(FileNumber))
        Get #FileNumber, , strReturn
    End If
    Close #FileNumber
    LoadToString = strReturn
End Function

