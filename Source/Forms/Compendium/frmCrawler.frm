VERSION 5.00
Begin VB.Form frmCrawler 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crawler"
   ClientHeight    =   876
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   2028
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCrawler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   876
   ScaleWidth      =   2028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCrawl 
      Interval        =   1000
      Left            =   1620
      Top             =   240
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1152
   End
End
Attribute VB_Name = "frmCrawler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrQuest() As String
Private mlngQuest As Long

Private mblnOverride As Boolean

Private Sub Form_Load()
    cfg.RefreshColors Me
    LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRaw As String
    
    ReDim Preserve mstrQuest(mlngQuest)
    strRaw = Join(mstrQuest, vbNewLine)
    xp.File.SaveStringAs CrawlFile(), strRaw
    CompareQuests
End Sub

Private Function CrawlFile() As String
    CrawlFile = App.Path & "\QuestCrawl.txt"
End Function

Private Sub LoadData()
    Dim strFile As String
    Dim strRaw As String
    Dim strLine() As String
    Dim strToken() As String
    
    mlngQuest = 1
    ReDim mstrQuest(db.Quests)
    strFile = CrawlFile()
    If Not xp.File.Exists(strFile) Then Exit Sub
    strRaw = xp.File.LoadToString(strFile)
    strLine = Split(strRaw, vbNewLine)
    For mlngQuest = 1 To UBound(strLine)
        If Len(strLine(mlngQuest)) = 0 Then Exit For
        mstrQuest(mlngQuest) = strLine(mlngQuest)
        strToken = Split(strLine(mlngQuest), vbTab)
        If strToken(0) <> db.Quest(mlngQuest).Quest Then Exit For
    Next
    If mlngQuest <= UBound(strLine) Then mlngQuest = mlngQuest - 1
End Sub

Private Sub tmrCrawl_Timer()
    Me.tmrCrawl.Enabled = False
    If mlngQuest > db.Quests Then
        Unload Me
    Else
        CheckQuest
        mlngQuest = mlngQuest + 1
        Me.tmrCrawl.Interval = RandomNumber(4000, 2000)
        Me.tmrCrawl.Enabled = True
    End If
End Sub

Private Sub CheckQuest()
    Dim strFile As String
    Dim strRaw As String
    Dim lngFavor As Long
    Dim strPatron As String
    
    Me.lblProgress.Caption = mlngQuest & " of " & db.Quests
    strFile = CrawlFile()
    strRaw = xp.DownloadURL(MakeWiki(db.Quest(mlngQuest).Wiki))
    lngFavor = FindFavor(strRaw)
    strPatron = FindPatron(strRaw)
    mstrQuest(mlngQuest) = db.Quest(mlngQuest).Quest & vbTab & strPatron & vbTab & lngFavor
End Sub

Private Function FindFavor(pstrRaw As String) As Long
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    
    lngPos = InStr(pstrRaw, "/page/Base_favor")
    If lngPos = 0 Then
        Debug.Print "Error: " & db.Quest(mlngQuest).Quest
        Exit Function
    End If
    lngStart = InStr(lngPos, pstrRaw, "<td>") + 4
    lngEnd = InStr(lngStart, pstrRaw, "</td>")
    FindFavor = Val(Trim$(Mid$(pstrRaw, lngStart, lngEnd - lngStart)))
End Function

Private Function FindPatron(pstrRaw As String) As String
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    
    lngPos = InStr(pstrRaw, "/page/Patron")
    If lngPos = 0 Then
        Debug.Print "Error: " & db.Quest(mlngQuest).Quest
        Exit Function
    End If
    lngStart = InStr(lngPos + 6, pstrRaw, "/page/") + 6
    lngEnd = InStr(lngStart, pstrRaw, Chr(34))
    FindPatron = UnWiki(Trim$(Mid$(pstrRaw, lngStart, lngEnd - lngStart)))
End Function

Private Sub CompareQuests()
    Dim strToken() As String
    Dim i As Long
    
    For i = 1 To mlngQuest - 1
        With db.Quest(i)
            strToken = Split(mstrQuest(i), vbTab)
            If UBound(strToken) <> 2 Then
                Debug.Print "Quest " & .Quest & " has an error"
            ElseIf strToken(0) <> .Quest Then
                Debug.Print "Quest " & i & " " & strToken(0) & " does not match " & .Quest
            ElseIf strToken(1) <> .Patron Then
                Debug.Print "Patron mismatch: " & .Quest & " set to " & .Patron & " but wiki lists " & strToken(1)
            ElseIf Val(strToken(2)) <> .Favor Then
                Debug.Print "Favor mismatch: " & .Quest & " set to " & .Favor & " but wiki lists " & strToken(2)
            End If
        End With
    Next
End Sub
