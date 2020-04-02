VERSION 5.00
Begin VB.Form frmRelease 
   Appearance      =   0  'Flat
   Caption         =   "Create Release Images"
   ClientHeight    =   1692
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00EBEBEB&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1692
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStatbar 
      AutoRedraw      =   -1  'True
      Height          =   252
      Left            =   660
      ScaleHeight     =   204
      ScaleWidth      =   2784
      TabIndex        =   1
      Top             =   1020
      Width           =   2832
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "Copy Files"
      Height          =   492
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   420
      Width           =   1632
   End
End
Attribute VB_Name = "frmRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum DestinationEnum
    deBoth
    dePrograms
    deSource
End Enum

Private Type QueueType
    Folder As String ' Relative to root
    File As String
    Absolute As String
    Dest As DestinationEnum
End Type

Private Queue() As QueueType
Private mlngFiles As Long
Private mlngBuffer As Long

Private mstrRoot As String
Private mstrPrograms As String
Private mstrSource As String
Private mstrRelease As String

Private xp As clsWindowsXP

Private statbar As clsStatusbar
Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    Set xp = New clsWindowsXP
    Set statbar = New clsStatusbar
    statbar.Init Me.picStatbar
    statbar.AddPanel vbNullString, vbLeftJustify, pseFixed, Me.picStatbar.TextWidth("Copying Files...  "), False
    statbar.AddPanel vbNullString, vbCenter, pseSpring, 100, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set statbar = Nothing
    Set xp = Nothing
End Sub

Private Sub chkCopy_Click()
    If UncheckButton(Me.chkCopy, mblnOverride) Then Exit Sub
    If GatherFiles() Then CopyFiles
End Sub


' ************* GATHER *************


Private Function GatherFiles() As Boolean
    Dim strFile As String
    Dim strRaw As String
    Dim strLine() As String
    Dim i As Long
    
    ClearFiles
    strFile = App.Path & "\Files.txt"
    If Not xp.File.Exists(strFile) Then
        MsgBox "No action due to missing file:" & vbNewLine & vbNewLine & strFile, vbInformation, "Notice"
        Exit Function
    End If
    strRaw = xp.File.LoadToString(strFile)
    strLine = Split(strRaw, vbNewLine)
    For i = 0 To UBound(strLine)
        If GatherFile(strLine(i)) Then Exit Function
    Next
    GatherFiles = True
End Function

Private Function GatherFile(ByVal pstrRaw As String) As Boolean
    Dim strField As String
    Dim strValue As String
    Dim enDest As DestinationEnum
    Dim lngPos As Long
    Dim strRaw As String
    
    pstrRaw = Trim$(pstrRaw)
    If Len(pstrRaw) = 0 Then Exit Function
    If Asc(pstrRaw) = 39 Then Exit Function
    strRaw = pstrRaw
    lngPos = InStr(strRaw, "=")
    If lngPos < 2 Then Exit Function
    strField = Left$(strRaw, lngPos - 1)
    strRaw = Mid$(strRaw, lngPos + 1)
    lngPos = InStr(strRaw, vbTab)
    If lngPos = 0 Then
        strValue = strRaw
    Else
        strValue = Left$(strRaw, lngPos - 1)
        Select Case Mid$(strRaw, lngPos + 1)
            Case "Programs": enDest = dePrograms
            Case "Source": enDest = deSource
        End Select
    End If
    Select Case strField
        Case "Root"
            mstrRoot = xp.Folder.RelativeToAbsolute(App.Path & "\" & strValue)
        Case "Destination"
            mstrRelease = mstrRoot & strValue
            strValue = mstrRoot & strValue & "\" & Format(Date, "yyyy-mm-dd")
            mstrPrograms = strValue & "_DDOLiteTools\"
            mstrSource = strValue & "_SourceCode\"
            If xp.Folder.Exists(mstrPrograms) Then
                MsgBox "Folder already exists:" & vbNewLine & vbNewLine & mstrPrograms, vbInformation, "Notice"
            ElseIf xp.Folder.Exists(mstrSource) Then
                MsgBox "Folder already exists:" & vbNewLine & vbNewLine & mstrSource, vbInformation, "Notice"
            Else
                Exit Function
            End If
            GatherFile = True
        Case "File"
            GatherFile = AddFile(strValue, enDest)
        Case "Project"
            GatherFile = GatherProject(strValue)
        Case "Folder"
            CopyFolders strValue, enDest
        Case Else
            MsgBox "Invalid line:" & vbNewLine & vbNewLine & pstrRaw, vbInformation, "Notice"
            GatherFile = True
    End Select
End Function


' ************* FILES *************


Private Sub ClearFiles()
    mlngFiles = -1
    mlngBuffer = 255
    ReDim Queue(mlngBuffer)
End Sub

Private Function AddFile(ByVal pstrFile As String, penDest As DestinationEnum) As Boolean
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim lngMid As Long
    Dim strFile As String
    Dim strFolder As String
    Dim strAbsolute As String
    Dim lngPos As Long
    Dim i As Long
    
    ' Parse file, folder and absolute
    If Left$(pstrFile, Len(mstrRoot)) = mstrRoot Then pstrFile = Mid$(pstrFile, Len(mstrRoot) + 1)
    lngPos = InStrRev(pstrFile, "\")
    If lngPos Then
        strFolder = Left$(pstrFile, lngPos)
        strFile = Mid$(pstrFile, lngPos + 1)
    Else
        strFolder = vbNullString
        strFile = pstrFile
    End If
    strAbsolute = xp.Folder.RelativeToAbsolute(mstrRoot & pstrFile)
    ' File exists?
    If Not xp.File.Exists(strAbsolute) Then
        MsgBox "File not found:" & vbNewLine & vbNewLine & strAbsolute, vbInformation, "Notice"
        AddFile = True
        Exit Function
    End If
    ' Find insertion point
    lngFirst = 0
    lngLast = mlngFiles
    Do While lngFirst <= lngLast
        lngMid = (lngFirst + lngLast) \ 2
        Select Case CompareQueue(strFile, strFolder, lngMid)
            Case -1: lngLast = lngMid - 1
            Case 1: lngFirst = lngMid + 1
            Case 0: Exit Function
        End Select
    Loop
    ' Allocate space
    mlngFiles = mlngFiles + 1
    If mlngFiles > mlngBuffer Then
        mlngBuffer = (mlngBuffer * 3) \ 2
        ReDim Preserve Queue(mlngBuffer)
    End If
    ' Insert new file to index lngFirst
    For i = mlngFiles To lngFirst + 1 Step -1
        Queue(i) = Queue(i - 1)
    Next
    With Queue(lngFirst)
        .File = strFile
        .Folder = strFolder
        .Absolute = strAbsolute
        .Dest = penDest
    End With
End Function

Private Function CompareQueue(pstrFile As String, pstrFolder As String, plngIndex As Long) As Long
    If pstrFolder < Queue(plngIndex).Folder Then
        CompareQueue = -1
    ElseIf pstrFolder > Queue(plngIndex).Folder Then
        CompareQueue = 1
    ElseIf pstrFile < Queue(plngIndex).File Then
        CompareQueue = -1
    ElseIf pstrFile > Queue(plngIndex).File Then
        CompareQueue = 1
    End If
End Function


' ************* PROJECTS *************


Private Function GatherProject(pstrProject As String) As Boolean
    Dim strFile As String
    Dim strFolder As String
    Dim strRaw As String
    Dim strLine() As String
    Dim i As Long
    
    GatherProject = True
    If AddFile(mstrRoot & pstrProject & ".vbw", deSource) Then Exit Function
    strFile = mstrRoot & pstrProject & ".vbp"
    strFolder = Left$(strFile, InStrRev(strFile, "\") - 1) & "\"
    If AddFile(strFile, deSource) Then Exit Function
    strRaw = xp.File.LoadToString(strFile)
    strLine = Split(strRaw, vbNewLine)
    For i = 0 To UBound(strLine)
        If ParseProjectLine(strLine(i), strFolder) Then Exit Function
    Next
    GatherProject = False
End Function

Private Function ParseProjectLine(pstrLine As String, pstrFolder As String) As Boolean
    Dim strField As String
    Dim strValue As String
    Dim strFile As String
    Dim strFileX As String
    Dim lngPos As Long
    
    lngPos = InStr(pstrLine, "=")
    If lngPos < 2 Then Exit Function
    strField = Left$(pstrLine, lngPos - 1)
    strValue = Mid$(pstrLine, lngPos + 1)
    Select Case strField
        Case "Class", "Module"
            lngPos = InStr(strValue, "; ")
            If lngPos = 0 Then Exit Function
            strValue = Mid$(strValue, lngPos + 2)
            strFile = xp.Folder.RelativeToAbsolute(pstrFolder & strValue)
            ParseProjectLine = AddFile(strFile, deSource)
        Case "UserControl", "Form"
            strFile = xp.Folder.RelativeToAbsolute(pstrFolder & strValue)
            strFileX = Left$(strFile, Len(strFile) - 1) & "x"
            If xp.File.Exists(strFileX) Then AddFile strFileX, deSource
            ParseProjectLine = AddFile(strFile, deSource)
        Case "ResFile32"
            strFile = pstrFolder & Replace(strValue, Chr(34), vbNullString)
            ParseProjectLine = AddFile(strFile, deSource)
    End Select
End Function


' ************* COPY *************


Private Sub CopyFiles()
    Dim strProgramsFolder As String
    Dim strSourceFolder As String
    Dim i As Long
    
    statbar.SetPanelCaption 1, "Copying Files..."
    statbar.ProgressbarInit 2, mlngFiles
    For i = 0 To mlngFiles
        statbar.ProgressbarIncrement
        With Queue(i)
            If .Dest <> deSource Then
                If strProgramsFolder <> .Folder Then
                    xp.Folder.Create mstrPrograms & .Folder
                    strProgramsFolder = .Folder
                End If
            End If
            If .Dest <> dePrograms Then
                If strSourceFolder <> .Folder Then
                    xp.Folder.Create mstrSource & .Folder
                    strSourceFolder = .Folder
                End If
            End If
            If .Dest <> deSource Then xp.File.Copy .Absolute, mstrPrograms & .Folder & .File
            If .Dest <> dePrograms Then xp.File.Copy .Absolute, mstrSource & .Folder & .File
        End With
    Next
    MsgBox "Copies made successfully.", vbInformation, "Notice"
    xp.Folder.Explore mstrRelease
    statbar.ProgressbarRemove
    statbar.SetPanelCaption 1, vbNullString
End Sub

Private Sub CopyFolders(pstrFolder As String, penDest As DestinationEnum)
    Dim strFrom As String
    
    strFrom = mstrRoot & pstrFolder
    If penDest <> deSource Then CopyFolder strFrom, mstrPrograms & pstrFolder
    If penDest <> dePrograms Then CopyFolder strFrom, mstrSource & pstrFolder
End Sub

Private Sub CopyFolder(pstrFrom As String, ByVal pstrTo As String)
    Dim lngPos As Long
    
    lngPos = InStrRev(pstrTo, "\")
    If lngPos Then pstrTo = Left$(pstrTo, lngPos - 1)
    xp.Folder.Create pstrTo
    xp.Folder.Copy pstrFrom, pstrTo
End Sub

