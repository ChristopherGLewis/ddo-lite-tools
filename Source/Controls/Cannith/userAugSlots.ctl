VERSION 5.00
Begin VB.UserControl userAugSlots 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2592
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2928
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2592
   ScaleWidth      =   2928
   Begin CannithCrafting.userHeader usrHeader 
      Height          =   252
      Left            =   0
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   2928
      _ExtentX        =   5165
      _ExtentY        =   445
      Margin          =   0
      Spacing         =   0
      CaptionTop      =   0
      BorderColor     =   -2147483640
      RightLinks      =   "Close"
   End
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   1
      Left            =   12
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   300
      Width           =   408
   End
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   2
      Left            =   12
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   408
   End
   Begin VB.PictureBox picActive 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   3
      Left            =   12
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1140
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   1
      Left            =   12
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "Red"
      Top             =   1920
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   2
      Left            =   432
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "Orange"
      Top             =   1920
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   3
      Left            =   852
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "Purple"
      Top             =   1920
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   4
      Left            =   1272
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "Blue"
      Top             =   1920
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   5
      Left            =   1692
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Green"
      Top             =   1920
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   6
      Left            =   2112
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "Yellow"
      Top             =   1920
      Width           =   408
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   408
      Index           =   7
      Left            =   2532
      ScaleHeight     =   408
      ScaleWidth      =   408
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "Colorless"
      Top             =   1920
      Width           =   408
   End
   Begin CannithCrafting.userHeader usrFooter 
      Height          =   252
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2928
      _ExtentX        =   5165
      _ExtentY        =   445
      Margin          =   0
      Spacing         =   0
      CaptionTop      =   0
      BorderColor     =   -2147483640
      RightLinks      =   "Help"
   End
   Begin VB.Shape shpBorder 
      Height          =   492
      Left            =   2040
      Top             =   480
      Width           =   648
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Choose up to three slots:"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   60
      TabIndex        =   13
      Top             =   1620
      Width           =   2304
   End
   Begin VB.Label lblAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   540
      TabIndex        =   12
      Tag             =   "0"
      Top             =   396
      Width           =   60
   End
   Begin VB.Label lblAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   2
      Left            =   540
      TabIndex        =   11
      Tag             =   "0"
      Top             =   816
      Width           =   60
   End
   Begin VB.Label lblAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   3
      Left            =   540
      TabIndex        =   10
      Tag             =   "0"
      Top             =   1236
      Width           =   60
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Context"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuContext 
         Caption         =   "View Detail"
         Index           =   0
      End
      Begin VB.Menu mnuContext 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContext 
         Caption         =   "Clear Augment"
         Index           =   2
      End
   End
End
Attribute VB_Name = "userAugSlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event DataChanged(Slot As SlotEnum, Text As String)
Public Event ChooseAugment(Slot As SlotEnum, Color As AugmentColorEnum)
Public Event Hotkey(Slot As SlotEnum, KeyCode As Integer)
Public Event CloseControl()

Private menSlot As SlotEnum

Private mtypAugSlot(1 To 7) As AugmentSlotType
Private mlngPicking As Long


' ************* CONTROL *************


Public Sub Init()
    InitHeader UserControl.usrHeader
    InitHeader UserControl.usrFooter
    mlngPicking = 0
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .shpBorder.Move 0, 0, .ScaleWidth, .ScaleHeight
        .usrHeader.Width = .ScaleWidth
        .usrFooter.Width = .ScaleWidth
    End With
End Sub

Private Sub InitHeader(pusr As userHeader)
    pusr.CaptionTop = 0
    pusr.Margin = UserControl.TextWidth(" ")
End Sub

Public Sub RefreshColors()
    Dim i As Long
    
    With UserControl
        .BackColor = cfg.GetColor(cgeWorkspace, cveBackground)
        cfg.ApplyColors .lbl, cgeWorkspace
        cfg.ApplyColors .usrHeader, cgeNavigation
        cfg.ApplyColors .lblAugment(1), cgeWorkspace
        cfg.ApplyColors .lblAugment(2), cgeWorkspace
        cfg.ApplyColors .lblAugment(3), cgeWorkspace
        cfg.ApplyColors .usrFooter, cgeNavigation
        .shpBorder.BorderColor = cfg.GetColor(cgeNavigation, cveBorderExterior)
        For i = 1 To 7
            If i < 4 Then cfg.ApplyColors .picActive(i), cgeWorkspace
            cfg.ApplyColors .picColor(i), cgeWorkspace
        Next
    End With
    RefreshData
End Sub

Public Property Get IconTop() As Long
    IconTop = UserControl.picActive(1).Top
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
        Case vbKeyUp
        Case vbKeyPageDown
        Case vbKeyPageUp
        Case vbKeyHome
        Case vbKeyEnd
        Case vbKeyAdd
        Case vbKeySubtract
        Case Else: Exit Sub
    End Select
    RaiseEvent Hotkey(menSlot, KeyCode)
End Sub


' ************* NAVIGATION *************


Private Sub usrHeader_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Close"
            mlngPicking = 0
            RaiseEvent CloseControl
    End Select
End Sub

Private Sub usrFooter_Click(pstrCaption As String)
    Select Case pstrCaption
        Case "Help": ShowHelp "Augment_Picker"
    End Select
End Sub

Private Sub usrHeader_MouseOver()
    UserControl.usrFooter.Caption = vbNullString
End Sub

Private Sub usrFooter_MouseOver()
    UserControl.usrFooter.Caption = vbNullString
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.usrFooter.Caption = vbNullString
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.usrFooter.Caption = vbNullString
End Sub


' ************* DRAWING *************


Private Sub RefreshData()
    Dim lngCount As Long
    Dim strResourceID As String
    Dim strCaption As String
    Dim lngColor As Long
    Dim i As Long
    
    For i = 1 To 7
        strResourceID = "AUG" & GetAugmentColorName(i, True)
        With mtypAugSlot(i)
            If .Exists And lngCount < 3 Then
                lngCount = lngCount + 1
                If .Augment <> 0 And .Variation <> 0 Then
                    strCaption = db.Augment(.Augment).Variation(.Variation)
                Else
                    strCaption = "Empty " & GetAugmentColorName(i) & " Slot"
                End If
                DrawPicture UserControl.picActive(lngCount), strResourceID
                If mlngPicking = 0 Or mlngPicking = i Then
                    lngColor = cfg.GetColor(cgeWorkspace, cveTextLink)
                Else
                    GrayScale UserControl.picActive(lngCount)
                    lngColor = cfg.GetColor(cgeWorkspace, cveText)
                End If
                With UserControl.lblAugment(lngCount)
                    .Caption = strCaption
                    .ForeColor = lngColor
                    .Tag = i
                End With
            Else
                .Exists = False
            End If
            DrawPicture UserControl.picColor(i), strResourceID, .Exists
            If mlngPicking <> 0 And mlngPicking <> i Then GrayScale UserControl.picColor(i)
        End With
    Next
    lngColor = cfg.GetColor(cgeWorkspace, cveTextDim)
    For i = lngCount + 1 To 3
        DrawPicture UserControl.picActive(i), "MSCBARTER"
        With UserControl.lblAugment(i)
            .Caption = "Unused slot"
            .ForeColor = lngColor
            .Tag = 0
        End With
    Next
End Sub

Private Sub DrawPicture(ppic As PictureBox, pstrResourceID As String, Optional pblnActive As Boolean = False)
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    With ppic
        lngWidth = .ScaleWidth
        lngHeight = .ScaleHeight
        .Cls
        On Error Resume Next
        .PaintPicture LoadResPicture(pstrResourceID, vbResIcon), PixelX, PixelY, lngWidth - (PixelX * 2), lngHeight - (PixelY * 2)
        If Err.Number Then .PaintPicture LoadResPicture("UNKNOWN", vbResIcon), PixelX, PixelY, lngWidth - (PixelX * 2), lngHeight - (PixelY * 2)
        On Error GoTo 0
    End With
    If pblnActive Then DrawBorder ppic
End Sub

Private Sub DrawBorder(ppic As PictureBox)
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngColor As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    
    lngWidth = ppic.ScaleWidth
    lngHeight = ppic.ScaleHeight
    lngColor = RGB(60, 160, 60)
    ' Top and Bottom
    lngLeft = PixelX * 2
    lngRight = lngWidth - PixelX * 2
    lngTop = 0
    lngBottom = lngHeight - PixelY
    ppic.Line (lngLeft, lngTop)-(lngRight, lngTop), lngColor
    ppic.Line (lngLeft, lngBottom)-(lngRight, lngBottom), lngColor
    ' Left and Right
    lngLeft = 0
    lngRight = lngWidth - PixelX
    lngTop = PixelY * 2
    lngBottom = lngHeight - PixelY * 2
    ppic.Line (lngLeft, lngTop)-(lngLeft, lngBottom), lngColor
    ppic.Line (lngRight, lngTop)-(lngRight, lngBottom), lngColor
    ' Draw boxes
    lngLeft = PixelX
    lngRight = lngWidth - PixelX * 2
    lngTop = PixelX
    lngBottom = lngHeight - PixelY * 2
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(165, 240, 165), PixelX, PixelY
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(250, 255, 250), PixelX, PixelY
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(160, 225, 160), PixelX, PixelY
    DrawBox ppic, lngLeft, lngTop, lngRight, lngBottom, RGB(70, 185, 90), PixelX, PixelY
End Sub

Private Sub DrawBox(ppic As PictureBox, plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long, plngColor As Long, plngPixelX As Long, plngPixelY As Long)
    ppic.Line (plngLeft, plngTop)-(plngRight, plngBottom), plngColor, B
    ' Shrink next box one pixel on all sides
    plngLeft = plngLeft + plngPixelX
    plngTop = plngTop + plngPixelY
    plngRight = plngRight - plngPixelX
    plngBottom = plngBottom - plngPixelY
End Sub


' ************* SLOTS *************


Public Property Get Slot() As SlotEnum
    Slot = menSlot
End Property

Public Property Let Slot(penSlot As SlotEnum)
    menSlot = penSlot
    UserControl.usrHeader.Caption = GetSlotName(menSlot)
End Property

Public Property Get SlotData() As String
    SlotData = GearsetAugmentToString(mtypAugSlot)
End Property

Public Property Let SlotData(pstrSlots As String)
    StringToGearsetAugment mtypAugSlot, pstrSlots
    mlngPicking = 0
    RefreshData
End Property

Private Sub DataHasChanged()
    RaiseEvent DataChanged(menSlot, GearsetAugmentToString(mtypAugSlot))
End Sub

Private Sub picColor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngPicking = 0 Then xp.SetMouseCursor mcHand
    UserControl.usrFooter.Caption = UserControl.picColor(Index).Tag
End Sub

Private Sub picColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCount As Long
    Dim i As Long
    
    If mlngPicking Then Exit Sub
    xp.SetMouseCursor mcHand
    If mtypAugSlot(Index).Exists Then
        mtypAugSlot(Index).Exists = False
        RefreshData
        DataHasChanged
    Else
        For i = 1 To 7
            If mtypAugSlot(i).Exists Then lngCount = lngCount + 1
        Next
        If lngCount < 3 Then
            mtypAugSlot(Index).Exists = True
            RefreshData
            DataHasChanged
        Else
            Notice "Maximum 3 augment slots per item"
        End If
    End If
End Sub

Private Sub picColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlngPicking Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub lblAugment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.lblAugment(Index).Tag = 0 Or mlngPicking <> 0 Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub lblAugment_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngColor As Long
    
    If mlngPicking <> 0 Then Exit Sub
    lngColor = UserControl.lblAugment(Index).Tag
    If lngColor = 0 Then Exit Sub
    If Button = vbRightButton Then
        If mtypAugSlot(lngColor).Augment Then
            UserControl.mnuContext(0).Tag = lngColor
            PopupMenu UserControl.mnuMain(0)
        End If
    Else
        mlngPicking = lngColor
        If mlngPicking = 0 Then Exit Sub
        xp.SetMouseCursor mcHand
        RefreshData
        RaiseEvent ChooseAugment(menSlot, UserControl.lblAugment(Index).Tag)
    End If
End Sub

Private Sub lblAugment_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.lblAugment(Index).Tag = 0 Or mlngPicking <> 0 Then Exit Sub
    xp.SetMouseCursor mcHand
End Sub

Private Sub mnuContext_Click(Index As Integer)
    Dim lngColor As Long
    Dim lngAugment As Long
    Dim lngVariant As Long
    
    lngColor = UserControl.mnuContext(0).Tag
    If lngColor = 0 Then Exit Sub
    With mtypAugSlot(lngColor)
        lngAugment = .Augment
        lngVariant = .Variation
    End With
    Select Case UserControl.mnuContext(Index).Caption
        Case "View Detail"
            OpenAugment lngAugment, lngVariant, 0
        Case "Clear Augment"
            With mtypAugSlot(lngColor)
                .Augment = 0
                .Variation = 0
            End With
            RaiseEvent DataChanged(menSlot, GearsetAugmentToString(mtypAugSlot))
            RefreshData
    End Select
End Sub

