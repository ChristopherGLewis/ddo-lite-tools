VERSION 5.00
Begin VB.Form frmValueChart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Collectable Value Chart"
   ClientHeight    =   9000
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4812
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmValueChart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   4812
   StartUpPosition =   3  'Windows Default
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   384
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4812
      _ExtentX        =   8488
      _ExtentY        =   677
      Spacing         =   264
      BorderColor     =   -2147483640
      LeftLinks       =   "by Value;by Collectable"
      RightLinks      =   "Help"
   End
   Begin VB.VScrollBar scrollVertical 
      Height          =   3552
      Left            =   4560
      TabIndex        =   4
      Top             =   372
      Width           =   252
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8592
      Left            =   0
      ScaleHeight     =   8592
      ScaleWidth      =   4572
      TabIndex        =   0
      Top             =   372
      Width           =   4572
      Begin VB.PictureBox picClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3132
         Left            =   0
         ScaleHeight     =   3132
         ScaleWidth      =   3012
         TabIndex        =   1
         Top             =   0
         Width           =   3012
         Begin VB.Label lnkCollectable 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Fragrant Drowshood"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   960
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   1848
         End
         Begin VB.Label lnkValue 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "9999"
            ForeColor       =   &H80000008&
            Height          =   216
            Index           =   0
            Left            =   312
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Popup"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Export to CSV"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmValueChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RowType
    Material As Long ' Index in db.Material()
    Value As Long
    Collectable As String
    SortName As String
End Type

Private mtypRow() As RowType
Private mlngRows As Long
Private mlngRowHeight As Long

Private mlngTab As Long


' ************* FORM *************


Private Sub Form_Load()
    cfg.RefreshColors Me
    mlngRowHeight = Me.TextHeight("Q") + PixelY * 3
    InitValues
    ShowValues
    If Not xp.DebugMode Then Call WheelHook(Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hwnd)
    CloseApp
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleWidth < Me.picClient.Width Or Me.ScaleHeight < mlngRowHeight * 6 Then Exit Sub
    Me.usrHeader.Width = Me.ScaleWidth
    Me.picContainer.Move 0, Me.usrHeader.Height + PixelY, Me.ScaleWidth - Me.scrollVertical.Width - PixelX, Me.ScaleHeight - Me.usrHeader.Height - PixelY
    Me.scrollVertical.Move Me.ScaleWidth - Me.scrollVertical.Width, Me.usrHeader.Height, Me.scrollVertical.Width, Me.ScaleHeight - Me.usrHeader.Height
    UpdateScrollbar
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    If Rotation < 0 Then
        KeyScroll Me.scrollVertical, 3
    Else
        KeyScroll Me.scrollVertical, -3
    End If
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Dim lngTab As Long
    
    Select Case pstrCaption
        Case "by Value": lngTab = 0
        Case "by Collectable": lngTab = 1
        Case "Help"
            ShowHelp "Value_Chart"
            Exit Sub
    End Select
    If mlngTab = lngTab Then Exit Sub
    mlngTab = lngTab
    ShowValues
End Sub


' ************* INITIALIZE *************


Private Sub InitValues()
    Dim i As Long
    
    mlngRows = 0
    ReDim mtypRow(1 To db.Materials)
    For i = 1 To db.Materials
        With db.Material(i)
            If .MatType = meCollectable And .Frequency <> feRare Then
                mlngRows = mlngRows + 1
                mtypRow(mlngRows).Material = i
                mtypRow(mlngRows).Collectable = .Material
                If InStr(.Material, "'") Then
                    mtypRow(mlngRows).SortName = Replace(.Material, "'", vbNullString)
                Else
                    mtypRow(mlngRows).SortName = .Material
                End If
                mtypRow(mlngRows).Value = .Value
            End If
        End With
    Next
    ReDim Preserve mtypRow(1 To mlngRows)
    Me.picClient.Height = mlngRowHeight * mlngRows
End Sub


' ************* CHART *************


Private Sub lnkCollectable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkCollectable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    OpenMaterial Me.lnkCollectable(Index).Caption
End Sub

Private Sub lnkValue_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
End Sub

Private Sub lnkValue_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    xp.SetMouseCursor mcHand
    OpenValue Me.lnkCollectable(Index).Caption
End Sub

Private Sub ShowValues()
    Dim lngTop As Long
    Dim i As Long
    
    Select Case mlngTab
        Case 0: SortByValue
        Case 1: SortByCollectable
    End Select
    For i = 1 To mlngRows
        If i > Me.lnkValue.UBound Then
            Load Me.lnkValue(i)
            With Me.lnkValue(i)
                .Move Me.lnkValue(0).Left, lngTop
                .Caption = mtypRow(i).Value
                .Visible = True
            End With
            Load Me.lnkCollectable(i)
            With Me.lnkCollectable(i)
                .Move Me.lnkCollectable(0).Left, lngTop
                .Caption = mtypRow(i).Collectable
                If Me.picClient.Width < .Left + .Width + PixelX Then Me.picClient.Width = .Left + .Width + PixelX
                .Visible = True
            End With
            lngTop = lngTop + mlngRowHeight
        Else
            Me.lnkValue(i).Caption = mtypRow(i).Value
            Me.lnkCollectable(i).Caption = mtypRow(i).Collectable
        End If
    Next
End Sub

' Insertion sort because it's stable and the list is small
Public Sub SortByCollectable()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As RowType

    iMin = LBound(mtypRow) + 1
    iMax = UBound(mtypRow)
    For i = iMin To iMax
        typSwap = mtypRow(i)
        For j = i To iMin Step -1
            If typSwap.SortName < mtypRow(j - 1).SortName Then mtypRow(j) = mtypRow(j - 1) Else Exit For
        Next j
        mtypRow(j) = typSwap
    Next i
End Sub

Public Sub SortByValue()
    Dim i As Long
    Dim j As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim typSwap As RowType

    iMin = LBound(mtypRow) + 1
    iMax = UBound(mtypRow)
    For i = iMin To iMax
        typSwap = mtypRow(i)
        For j = i To iMin Step -1
            If typSwap.Value > mtypRow(j - 1).Value Then mtypRow(j) = mtypRow(j - 1) Else Exit For
        Next j
        mtypRow(j) = typSwap
    Next i
End Sub


' ************* SCROLLING *************


Private Sub UpdateScrollbar()
    Dim dblValue As Double
    
    With Me.scrollVertical
        If .Value <> 0 And .Max <> 0 Then dblValue = .Value / .Max
        .Min = 0
        .Max = Me.picClient.Height - Me.picContainer.Height
        .SmallChange = mlngRowHeight
        .LargeChange = Me.picContainer.Height
       .Value = dblValue * .Max
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown: KeyScroll Me.scrollVertical, 1
        Case vbKeyUp: KeyScroll Me.scrollVertical, -1
        Case vbKeyPageUp: KeyScroll Me.scrollVertical, -2
        Case vbKeyPageDown: KeyScroll Me.scrollVertical, 2
        Case vbKeyHome: KeyScroll Me.scrollVertical, 0
        Case vbKeyEnd: KeyScroll Me.scrollVertical, 99
        Case vbKeyEscape: Unload Me
    End Select
End Sub

Private Sub KeyScroll(pctl As Control, plngIncrement As Long)
    Dim lngValue As Long
    
    If Not pctl.Visible Then Exit Sub
    Select Case plngIncrement
        Case -3, -1, 1, 3: lngValue = pctl.Value + (plngIncrement * pctl.SmallChange)
        Case -2, 2: lngValue = pctl.Value + (plngIncrement \ 2) * pctl.LargeChange
        Case 0: lngValue = 0
        Case 99: lngValue = pctl.Max
    End Select
    If lngValue < 0 Then lngValue = 0
    If lngValue > pctl.Max Then lngValue = pctl.Max
    If pctl.Value <> lngValue Then pctl.Value = lngValue
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
    Me.picClient.Top = -Me.scrollVertical.Value
End Sub


' ************* EXPORT *************


Private Sub picClient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu Me.mnuMain(0)
End Sub

Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu Me.mnuMain(0)
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Dim strFile As String
    
    strFile = xp.ShowSaveAsDialog(App.Path & "\Save", "Values.csv", "CSV Files|*.csv", "*.csv")
    If Len(strFile) = 0 Then Exit Sub
    If xp.File.Exists(strFile) Then
        If MsgBox("File exists. Overwrite?", vbQuestion + vbYesNo + vbDefaultButton2, "Notice") <> vbYes Then Exit Sub
        xp.File.Delete strFile
    End If
    ExportToCSV strFile
End Sub

Private Sub ExportToCSV(pstrFile As String)
    Dim strLine() As String
    Dim lngPos As Long
    Dim i As Long
    
    ReDim strLine(1 To mlngRows)
    For i = 1 To mlngRows
        With mtypRow(i)
            strLine(i) = .Value & ","
            lngPos = InStr(.SortName, ",")
            If lngPos Then
                strLine(i) = strLine(i) & Left$(.SortName, lngPos - 1)
            Else
                strLine(i) = strLine(i) & .SortName
            End If
        End With
    Next
    xp.File.SaveStringAs pstrFile, Join(strLine, vbNewLine)
End Sub
