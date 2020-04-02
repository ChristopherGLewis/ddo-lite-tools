VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Help"
   ClientHeight    =   7368
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   7512
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   614
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   626
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar scrollVertical 
      Height          =   2352
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2520
      ScaleHeight     =   252
      ScaleWidth      =   732
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4260
      Top             =   2640
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2652
      Left            =   60
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   326
      TabIndex        =   1
      Top             =   120
      Width           =   3912
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1812
         Left            =   240
         ScaleHeight     =   151
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   276
         TabIndex        =   2
         Top             =   180
         Width           =   3312
         Begin VB.Label lnkLink 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Link"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   720
            Visible         =   0   'False
            Width           =   408
         End
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Ellis Dee
' This module contains private versions of general helper functions for portability
Option Explicit

Private Const HAND As Long = 32649&

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private mblnLoaded As Boolean

Private mlngX As Long
Private mlngY As Long
Private mlngTextHeight As Long
Private mlngSpace As Long
Private mlngWidth As Long
Private mlngCode As Long
Private mlngMargin(2) As Long

Private mstrStack() As String
Private mlngStack As Long


' ************* GENERAL *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    mblnLoaded = False
    If Not DebugMode() Then Call WheelHook(Me.Hwnd)
End Sub

Private Sub Form_Resize()
    Me.tmrResize.Enabled = False
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleWidth < Me.TextWidth("X") * 10 Or Me.ScaleHeight < Me.TextHeight("X") * 3 Then
        Me.scrollVertical.Visible = False
        Exit Sub
    End If
    Me.picContainer.Move Me.scrollVertical.Width, 0, Me.ScaleWidth - Me.scrollVertical.Width * 2, Me.ScaleHeight
    Me.picClient.Move 0, 0, Me.picContainer.ScaleWidth
    Me.scrollVertical.Move Me.ScaleWidth - Me.scrollVertical.Width, 0, Me.scrollVertical.Width, Me.ScaleHeight
    If Not mblnLoaded Then
        DrawText
        mblnLoaded = True
    Else
        Me.tmrResize.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    If Not DebugMode() Then Call WheelUnHook(Me.Hwnd)
    Select Case App.Title
        Case "Character Builder Lite"
            If GetForm(frm, "frmCompendium") Then frm.UpdateWindowMenu
        Case "Cannith Crafting Builder Lite"
            CloseApp
    End Select
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    If Rotation < 0 Then
        KeyScroll 3
    Else
        KeyScroll -3
    End If
End Sub


' ************* DRAWING *************


Public Sub DrawText()
    Dim lngParagraph As Long
    Dim lngWord As Long
    Dim strCaption As String
    Dim i As Long
    
    ' Prep
    With Me.picClient
        .Visible = False
        .Top = 0
        .Cls
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        mlngTextHeight = .TextHeight("X")
        mlngSpace = .TextWidth(" ")
        mlngMargin(0) = 0
        mlngMargin(1) = .TextWidth(" - ")
        mlngMargin(2) = .TextWidth("   - ")
    End With
    mlngCode = 0
    mlngX = 0
    mlngY = 0
    ' Print words
    With help.HelpItem(help.Current)
        For lngParagraph = 0 To .Paragraphs
            With .Paragraph(lngParagraph)
                For lngWord = 0 To .Words
                    ApplyFormat .Word(lngWord), True
                    mlngWidth = Me.picClient.TextWidth(.Word(lngWord).Text)
                    If mlngX + mlngWidth + mlngSpace > Me.picClient.ScaleWidth Then
                        If mlngX > 0 Then NewLine .Indent
                    End If
                    PrintWord .Word(lngWord)
                    ApplyFormat .Word(lngWord), False
                Next
            End With
            NewLine 0
        Next
    End With
    ' Scrollbar
    mlngY = mlngY - mlngTextHeight
    If Me.picClient.ScaleHeight > mlngY Then Me.picClient.Height = mlngY
    If Me.picClient.Height > Me.picContainer.Height Then
        With Me.scrollVertical
            .Value = 0
            .Min = 0
            .Max = (picClient.Height - Me.ScaleHeight)
            .SmallChange = mlngTextHeight
            .LargeChange = ((.Height \ mlngTextHeight) * mlngTextHeight)
        End With
        Me.scrollVertical.Visible = True
    Else
        Me.scrollVertical.Visible = False
    End If
    ' Release unused links
    For i = Me.lnkLink.UBound To mlngCode + 1 Step -1
        Unload Me.lnkLink(i)
    Next
    Me.picClient.Visible = True
    DrawBack
    ' Form caption
    strCaption = Replace(help.HelpItem(help.Current).HelpCode, "_", " ")
    Select Case strCaption
        Case "What's New?", "Frequently Asked Questions", "Getting Started", "Table of Contents"
        Case Else: If Left$(strCaption, 23) <> "Save File Specification" Then strCaption = strCaption & " Help"
    End Select
    Me.Caption = strCaption
End Sub

Private Sub ApplyFormat(ptypWord As WordType, pblnApply As Boolean)
    Dim enColor As ColorValueEnum
    
    If ptypWord.Bold Then Me.picClient.FontBold = pblnApply
    If ptypWord.Italics Then Me.picClient.FontItalic = pblnApply
    If ptypWord.Underline Then Me.picClient.FontUnderline = pblnApply
    If ptypWord.ErrorText Then
        If pblnApply Then enColor = cveTextError Else enColor = cveText
        Me.picClient.ForeColor = cfg.GetColor(cgeWorkspace, enColor)
    End If
End Sub

Private Sub NewLine(plngIndent As Long)
    mlngX = mlngMargin(plngIndent)
    mlngY = mlngY + mlngTextHeight
    If mlngY + mlngTextHeight * 2 > Me.picClient.ScaleHeight Then Me.picClient.Height = mlngY + mlngTextHeight * 2
End Sub

Private Sub PrintWord(ptypWord As WordType)
    Me.picClient.CurrentX = mlngX
    Me.picClient.CurrentY = mlngY
    If Len(ptypWord.Code) Then
        mlngCode = mlngCode + 1
        If mlngCode > Me.lnkLink.UBound Then Load Me.lnkLink(mlngCode)
        With Me.lnkLink(mlngCode)
            .Caption = ptypWord.Text
            .Tag = ptypWord.Code
            .ToolTipText = Mid$(ptypWord.Code, 5)
            .Move mlngX, mlngY
            .Visible = True
            mlngX = .Left + .Width + mlngSpace
        End With
    Else
        Me.picClient.Print ptypWord.Text
        mlngX = mlngX + mlngWidth + mlngSpace
    End If
End Sub

Private Sub DrawBack()
    Dim strCaption As String
    Dim lngWidth As Long
    
    If mlngStack < 1 Then
        Me.picBack.Visible = False
        Exit Sub
    End If
    strCaption = " < Back  "
    With Me.picBack
        .ForeColor = cfg.GetColor(cgeWorkspace, cveTextLink)
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        .Width = Me.TextWidth(strCaption)
        .Height = Me.TextHeight(strCaption)
        .Cls
    End With
    Me.picBack.Print strCaption
    lngWidth = Me.ScaleWidth
    If Me.scrollVertical.Visible Then lngWidth = lngWidth - Me.scrollVertical.Width
    Me.picBack.Move lngWidth - Me.picBack.Width, 0
    Me.picBack.Visible = True
End Sub


' ************* LINKS *************


Private Sub lnkLink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, HAND)
End Sub

Private Sub lnkLink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTag As String
    Dim strCommand As String
    Dim strValue As String
    Dim lngPos As Long
    Dim strFile As String
    
    SetCursor LoadCursor(0, HAND)
    strTag = Me.lnkLink(Index).Tag
    lngPos = InStr(strTag, "=")
    If lngPos = 0 Then Exit Sub
    strCommand = Left(strTag, lngPos - 1)
    strValue = Mid(strTag, lngPos + 1)
    Select Case LCase(strCommand)
        Case "url"
            OpenURL strValue
        Case "run"
            strFile = strValue
            If Left(strFile, 1) = "\" Then strFile = App.Path & strFile
            If Not xp.File.Exists(strFile) Then
                If Not xp.Folder.Exists(strFile) Then
                    If strFile = App.Path & "\Error.log" Then
                        MsgBox "There were no errors on startup.", vbInformation, "Notice"
                    Else
                        MsgBox strFile & " not found", vbInformation, "Notice"
                    End If
                    Exit Sub
                End If
            End If
            Run strFile
        Case "hlp"
            If help.Current = 0 Then
                MsgBox "There was a problem and Help needs to close", vbInformation, "Notice"
                Unload Me
                Exit Sub
            End If
            mlngStack = mlngStack + 1
            ReDim Preserve mstrStack(1 To mlngStack)
            mstrStack(mlngStack) = help.HelpItem(help.Current).HelpCode
            ShowHelp strValue
        Case "frm"
            OpenForm strValue
    End Select
End Sub

Private Sub lnkLink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, HAND)
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, HAND)
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, HAND)
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTopic As String
    
    SetCursor LoadCursor(0, HAND)
    If mlngStack < 1 Then Exit Sub
    strTopic = mstrStack(mlngStack)
    mlngStack = mlngStack - 1
    If mlngStack < 1 Then
        mlngStack = 0
        Erase mstrStack
    Else
        ReDim Preserve mstrStack(1 To mlngStack)
    End If
    ShowHelp strTopic
End Sub


' ************* SCROLLBAR *************


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown: KeyScroll 1
        Case vbKeyUp: KeyScroll -1
        Case vbKeyPageUp: KeyScroll -2
        Case vbKeyPageDown: KeyScroll 2
        Case vbKeyHome: KeyScroll 0
        Case vbKeyEnd: KeyScroll 99
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub KeyScroll(plngIncrement As Long)
    Dim lngValue As Long
    
    If Not Me.scrollVertical.Visible Then Exit Sub
    With Me.scrollVertical
        Select Case plngIncrement
            Case -3, -1, 1, 3: lngValue = .Value + plngIncrement * .SmallChange
            Case -2, 2: lngValue = .Value + (plngIncrement \ 2) * .LargeChange
            Case 0: lngValue = 0
            Case 99: lngValue = .Max
        End Select
        If lngValue < 0 Then lngValue = 0
        If lngValue > .Max Then lngValue = .Max
        If .Value <> lngValue Then .Value = lngValue
    End With
End Sub

Private Sub scrollVertical_GotFocus()
    Me.picClient.SetFocus
End Sub

Private Sub scrollVertical_Change()
    Scroll
End Sub

Private Sub scrollVertical_Scroll()
    Scroll
End Sub

Private Sub Scroll()
    Me.picClient.Top = (0 - Me.scrollVertical.Value)
End Sub

Private Sub tmrResize_Timer()
    Me.tmrResize.Enabled = False
    DrawText
End Sub


' ************* UTILS *************


Private Function DebugMode() As Boolean
    DebugMode = (App.LogMode = 0)
End Function

Private Sub OpenURL(ByVal URL As String)
    ShellExecute 0&, "OPEN", URL, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Function FileExists(ByVal File As String) As Boolean
    FileExists = (PathFileExists(File) = 1)
    If FileExists Then FileExists = (PathIsDirectory(File) = 0)
End Function

Private Function Run(ByVal File As String, Optional ByVal WindowState As Long = 1, Optional ByVal DefaultFolder As String) As Long
    Dim lngPos As Long
    Dim lngDesktop As Long
    
    If DefaultFolder = "" Then
        lngPos = InStrRev(File, "\")
        If lngPos > 0 Then DefaultFolder = Left$(File, lngPos - 1)
    End If
    lngDesktop = GetDesktopWindow()
    Run = ShellExecute(lngDesktop, "Open", File, "", DefaultFolder, WindowState) 'SW_SHOW)
End Function

