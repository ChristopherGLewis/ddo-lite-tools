VERSION 5.00
Begin VB.UserControl userItemSlot 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   996
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11076
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   996
   ScaleWidth      =   11076
   Begin VB.Timer tmrDeactivate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   360
   End
   Begin VB.PictureBox picEldritch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   9492
      ScaleHeight     =   384
      ScaleWidth      =   384
      TabIndex        =   21
      Top             =   252
      Visible         =   0   'False
      Width           =   384
   End
   Begin CannithCrafting.userCheckBox usrchkEldritch 
      Height          =   252
      Left            =   9120
      TabIndex        =   8
      Top             =   300
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin CannithCrafting.userSpinner usrspnML 
      Height          =   312
      Left            =   840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   204
      Visible         =   0   'False
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   550
      Max             =   34
      Value           =   34
      ShowZero        =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin CannithCrafting.userCheckBox usrchkAll 
      Height          =   252
      Left            =   10104
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Value           =   0   'False
      Caption         =   ""
   End
   Begin CannithCrafting.userCheckBox usrchkML 
      Height          =   252
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   564
      Visible         =   0   'False
      Width           =   912
      _ExtentX        =   1609
      _ExtentY        =   445
      Caption         =   "Done"
   End
   Begin CannithCrafting.userCheckBox usrchkEffect 
      Height          =   252
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin CannithCrafting.userCheckBox usrchkEffect 
      Height          =   252
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin CannithCrafting.userCheckBox usrchkEffect 
      Height          =   252
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Top             =   660
      Visible         =   0   'False
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin CannithCrafting.userIcon usrIcon 
      Height          =   408
      Left            =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Width           =   408
      _ExtentX        =   720
      _ExtentY        =   720
      AllowMenu       =   0   'False
      Style           =   1
   End
   Begin CannithCrafting.userCheckBox usrchkAugment 
      Height          =   252
      Index           =   0
      Left            =   5520
      TabIndex        =   5
      Top             =   60
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin CannithCrafting.userCheckBox usrchkAugment 
      Height          =   252
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   360
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin CannithCrafting.userCheckBox usrchkAugment 
      Height          =   252
      Index           =   2
      Left            =   5520
      TabIndex        =   7
      Top             =   660
      Width           =   312
      _ExtentX        =   550
      _ExtentY        =   445
      Caption         =   ""
   End
   Begin VB.Label lblBaseItem 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Armor / Shield"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   5856
      TabIndex        =   24
      Top             =   684
      Visible         =   0   'False
      Width           =   1284
   End
   Begin VB.Label lblDone 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Done"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   10440
      TabIndex        =   23
      Top             =   456
      Width           =   468
   End
   Begin VB.Label lblDone 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Item"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   10440
      TabIndex        =   22
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lnkAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Heavy Fortification"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   0
      Left            =   5856
      TabIndex        =   20
      Top             =   84
      Width           =   1668
   End
   Begin VB.Label lnkAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Diamond of Vitality +20"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   1
      Left            =   5856
      TabIndex        =   19
      Top             =   384
      Width           =   2124
   End
   Begin VB.Label lnkAugment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ruby Eye of the Glacier"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   2
      Left            =   5856
      TabIndex        =   18
      Top             =   684
      Width           =   2076
   End
   Begin VB.Line linBottom 
      X1              =   4500
      X2              =   6000
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line linRight 
      X1              =   10980
      X2              =   10980
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line linTop 
      X1              =   3900
      X2              =   6660
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Line linLeft 
      X1              =   60
      X2              =   60
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label lblNamed 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Named"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   840
      TabIndex        =   16
      Top             =   384
      Width           =   636
   End
   Begin VB.Label lblScaling 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "27"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   2
      Left            =   4020
      TabIndex        =   15
      Top             =   684
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblScaling 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "27"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   4320
      TabIndex        =   14
      Top             =   384
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblScaling 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "27"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   84
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lnkEffect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Ins. Enchant Resist"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   2
      Left            =   2256
      TabIndex        =   12
      Top             =   684
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lnkEffect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Monstrous Hum. Bane"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   1
      Left            =   2256
      TabIndex        =   11
      Top             =   384
      Visible         =   0   'False
      Width           =   1992
   End
   Begin VB.Label lnkEffect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Prefix"
      ForeColor       =   &H00FF0000&
      Height          =   216
      Index           =   0
      Left            =   2256
      TabIndex        =   10
      Top             =   84
      Visible         =   0   'False
      Width           =   492
   End
End
Attribute VB_Name = "userItemSlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event LevelChange(ML As Long)
Public Event MLDone(Value As Boolean)
Public Event EffectDone(ByVal Effect As Long, Value As Boolean)
Public Event AugmentDone(ByVal Color As AugmentColorEnum, Value As Boolean)
Public Event EldritchDone(Value As Boolean)
Public Event GearClick(Gear As GearEnum, ML As Long, Prefix As Long, Suffix As Long, Extra As Long, Augment As String, Eldritch As Long)

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hcursor As Long) As Long

Private Const CURSOR_HAND As Long = 32649&

Private Type UserEffectType
    Shard As Long
    Done As Boolean
    Visible As Boolean
End Type

Private Type UserAugmentType
    Color As AugmentColorEnum
    Augment As Long
    Variation As Long
    Scaling As Long
    Done As Boolean
    Visible As Boolean
    HasError As Boolean
End Type

Private Type UserItemSlotType
    Slot As SlotEnum
    Gear As GearEnum
    ItemStyle As String
    Crafted As Boolean
    Named As String
    ML As Long
    MLDone As Boolean
    MLVisible As Boolean
    Effect(2) As UserEffectType
    Augment(1 To 7) As AugmentSlotType
    AugSlot(2) As UserAugmentType
    Scales As Boolean
    Scaling() As String
    EldritchRitual As Long
    EldritchDone As Boolean
    AllDone As Boolean
End Type

Private itm As UserItemSlotType
Private mblnEldritch As Boolean ' Leave space for Eldritch Ritual?
Private mlngWidth As Long ' Width of labels (Effects and Augments)

Private mblnOverride As Boolean


' ************* CONTROL *************


Public Sub RefreshColors()
    Dim ctl As Control
    
    UserControl.BackColor = cfg.GetColor(cgeControls, cveBackground)
    For Each ctl In UserControl.Controls
        cfg.ApplyColors ctl, cgeControls
    Next
End Sub

Public Property Let EldritchLeaveSpace(pblnLeaveSpace As Boolean)
    mblnEldritch = pblnLeaveSpace
    SizeControls
End Property

Private Sub SizeControls()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngTextHeight As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim lngMargin As Long
    Dim lngLabelMargin As Long
    Dim lngPixelX As Long
    Dim lngPixelY As Long
    Dim i As Long

    lngPixelX = Screen.TwipsPerPixelX
    lngPixelY = Screen.TwipsPerPixelY
    With UserControl
        lngMargin = .usrIcon.Left
        lngLabelMargin = .lnkEffect(0).Left - .usrchkEffect(0).Left
        lngTextHeight = .TextHeight("Q")
        ' Icon
        .usrIcon.Top = (.ScaleHeight - .usrIcon.Height) \ 2
        ' ML
        lngLeft = .usrIcon.Left + .usrIcon.Height + lngMargin
        .usrspnML.Left = lngLeft
        .usrchkML.Left = lngLeft
        ' Named
        lngTop = (.ScaleHeight - lngTextHeight) \ 2
        .lblNamed.Move lngLeft, lngTop
        ' All Done
        lngLeft = .ScaleWidth - lngMargin - .TextWidth("Done")
        lngTop = (.ScaleHeight - (lngTextHeight * 2)) \ 2
        .lblDone(0).Move lngLeft, lngTop
        .lblDone(1).Move lngLeft, lngTop + lngTextHeight
        lngLeft = lngLeft - lngLabelMargin
        lngTop = (.ScaleHeight - .usrchkAll.Height) \ 2
        .usrchkAll.Move lngLeft, lngTop
        ' Eldritch
        If mblnEldritch Then
            lngLeft = lngLeft - lngMargin - .picEldritch.Width
            lngTop = (.ScaleHeight - .picEldritch.Height) \ 2
            .picEldritch.Move lngLeft, lngTop
            lngLeft = lngLeft - lngLabelMargin
            lngTop = (.ScaleHeight - .usrchkEldritch.Height) \ 2
            .usrchkEldritch.Move lngLeft, lngTop
        End If
        ' Effects
        mlngWidth = ((lngLeft - (.usrspnML.Left + .usrspnML.Width) - (lngMargin * 3)) \ 2) - lngLabelMargin
        lngLeft = .usrspnML.Left + .usrspnML.Width + lngMargin
        For i = 0 To 2
            .usrchkEffect(i).Left = lngLeft
            .lnkEffect(i).Left = lngLeft + lngLabelMargin
        Next
        ' Augments
        lngLeft = lngLeft + lngLabelMargin + mlngWidth + lngMargin
        For i = 0 To 2
            .usrchkAugment(i).Left = lngLeft
            .lnkAugment(i).Left = lngLeft + lngLabelMargin
        Next
        ' Base level
        .lblBaseItem.Move .lnkAugment(2).Left, .lnkAugment(2).Top
        ' Lines
        lngRight = .ScaleWidth - lngPixelX
        lngBottom = .ScaleHeight
        MoveLine .linLeft, 0, 0, 0, lngBottom
        MoveLine .linTop, lngPixelX, 0, lngRight, 0
        MoveLine .linRight, lngRight, 0, lngRight, lngBottom
        MoveLine .linBottom, lngPixelX, lngBottom - lngPixelY, lngRight, lngBottom - lngPixelY
    End With
End Sub

Private Sub MoveLine(plin As Line, plngLeft As Long, plngTop As Long, plngRight As Long, plngBottom As Long)
    plin.X1 = plngLeft
    plin.Y1 = plngTop
    plin.X2 = plngRight
    plin.Y2 = plngBottom
End Sub


' ************* SET DATA *************


Public Sub Clear()
    Dim typBlank As UserItemSlotType
    Dim i As Long
    
    itm = typBlank
    With UserControl
        .usrspnML.Visible = False
        .usrchkML.Visible = False
        .lblNamed.Visible = False
        For i = 0 To 2
            .usrchkEffect(i).Visible = False
            .lnkEffect(i).Visible = False
            .lblScaling(i).Visible = False
            .usrchkAugment(i).Visible = False
            .lnkAugment(i).Visible = False
        Next
        .lblBaseItem.Visible = False
        .usrchkEldritch.Visible = False
        .picEldritch.Visible = False
        .usrchkAll.Visible = False
        .lblDone(0).Visible = False
        .lblDone(1).Visible = False
    End With
End Sub

Public Property Get Slot() As SlotEnum
    Slot = itm.Slot
End Property

Public Property Let Slot(ByVal penSlot As SlotEnum)
    itm.Slot = penSlot
End Property

Public Property Let Gear(penGear As GearEnum)
    itm.Gear = penGear
End Property

Public Property Let ItemStyle(pstrItemStyle As String)
    Dim lngIndex As Long
    
    itm.ItemStyle = pstrItemStyle
    lngIndex = SeekItem(itm.ItemStyle)
    If lngIndex Then
        itm.Scales = db.Item(lngIndex).Scales
        itm.Scaling = db.Item(lngIndex).Scaling
    End If
End Property

Public Property Let Crafted(pblnCrafted As Boolean)
    itm.Crafted = pblnCrafted
End Property

Public Property Let Named(pstrNamed As String)
    itm.Named = pstrNamed
End Property

Public Sub SetML(plngML As Long, pblnDone As Boolean)
    itm.ML = plngML
    itm.MLDone = pblnDone
End Sub

Public Sub SetEffects(plngEffect() As Long, pblnDone() As Boolean)
    Dim i As Long
    
    For i = 0 To 2
        With itm.Effect(i)
            .Shard = plngEffect(i)
            .Done = pblnDone(i)
        End With
    Next
End Sub

Public Sub SetAugments(pstrText As String)
    StringToGearsetAugment itm.Augment, pstrText
End Sub

Public Sub SetEldritch(plngRitual As Long, pblnDone As Boolean)
    itm.EldritchRitual = plngRitual
    itm.EldritchDone = pblnDone
End Sub


' ************* SHOW DATA *************


Public Sub Refresh()
    Dim blnVisible As Boolean
    Dim i As Long
    
    mblnOverride = True
    With UserControl
        ' Gear icon
        With .usrIcon
            .Init itm.Gear, uiseLink, True
            .IconName = itm.ItemStyle
            .Active = False
        End With
        ' ML
        .usrspnML.Value = itm.ML
        .usrspnML.Visible = itm.Crafted
        .usrchkML.Value = itm.MLDone
        .usrchkML.Visible = itm.Crafted
        ' Named
        .lblNamed.Caption = itm.Named
        .lblNamed.Visible = Not itm.Crafted
        ' Effects
        ShowEffects
        ' Augments
        ShowAugments
        ' Base item
        ShowBaseItem
        ' Eldritch
        blnVisible = (itm.EldritchRitual <> 0)
        .usrchkEldritch.Value = itm.EldritchDone
        .usrchkEldritch.Visible = blnVisible
        If blnVisible Then .picEldritch.TooltipText = db.Ritual(itm.EldritchRitual).RitualName
        With .picEldritch
            If blnVisible Then .PaintPicture LoadResPicture("MSCELDRITCH", vbResBitmap), 0, 0, .Width, .Height
            .Visible = blnVisible
        End With
        ' All Done
        CheckAllDone
    End With
    mblnOverride = False
End Sub

Private Sub ShowEffects()
    Dim lngWidth As Long
    Dim blnVisible As Boolean
    Dim strEffect As String
    Dim strCaption As String
    Dim strScale As String
    Dim i As Long
    
    With UserControl
        lngWidth = .usrchkAugment(0).Left - .lnkEffect(0).Left - .TextWidth("  ")
        For i = 0 To 2
            blnVisible = (itm.Effect(i).Shard <> 0)
            If blnVisible Then blnVisible = (UserControl.usrspnML.Value >= db.Shard(itm.Effect(i).Shard).ML)
            If blnVisible = True And i = 2 And UserControl.usrspnML.Value < 10 Then blnVisible = False
            If blnVisible Then
                .usrchkEffect(i).Value = itm.Effect(i).Done
                strCaption = db.Shard(itm.Effect(i).Shard).ShardName
                strScale = GetScale(i)
                If .TextWidth(strCaption & " " & strScale) > lngWidth Then strCaption = db.Shard(itm.Effect(i).Shard).Abbreviation
                .lnkEffect(i).Caption = strCaption
                .lblScaling(i).Caption = strScale
                .lblScaling(i).Move .lnkEffect(i).Left + .lnkEffect(i).Width + .TextWidth(" ")
            End If
            .usrchkEffect(i).Visible = blnVisible
            .lnkEffect(i).Visible = blnVisible
            .lblScaling(i).Visible = blnVisible
            itm.Effect(i).Visible = blnVisible
        Next
    End With
End Sub

Private Function GetScale(penAffix As AffixEnum) As String
    Dim lngShard As Long
    Dim lngIndex As Long
    
    lngShard = itm.Effect(penAffix).Shard
    If lngShard = 0 Then Exit Function
    If db.Shard(lngShard).ScaleName = "None" Then Exit Function
    lngIndex = SeekScaling(db.Shard(lngShard).ScaleName)
    If lngIndex Then GetScale = db.Scaling(lngIndex).Table(itm.ML)
End Function

Private Sub ShowAugments()
    Dim lngIndex As Long
    Dim blnError As Boolean
    Dim lngWidth As Long
    Dim lngColor As Long
    Dim strCaption As String
    Dim i As Long
    
    Erase itm.AugSlot
    With UserControl
        For i = 1 To 7
            If itm.Augment(i).Exists Then
                If lngIndex < 3 Then
                    strCaption = ScaledAugmentName(itm.Augment(i), i, itm.ML, blnError)
                    If itm.Augment(i).Augment = 0 Or itm.Augment(i).Variation = 0 Then
                        lngColor = cfg.GetColor(cgeControls, cveText)
                        strCaption = "Empty " & GetAugmentColorName(i) & " slot"
                    ElseIf itm.Augment(i).Scaling = 0 Or blnError = True Then
                        lngColor = cfg.GetColor(cgeControls, cveTextError)
                    Else
                        lngColor = cfg.GetColor(cgeControls, cveTextLink)
                    End If
                    With .usrchkAugment(lngIndex)
                        .Style = i
                        .Value = (itm.Augment(i).Augment = 0 Or itm.Augment(i).Variation = 0 Or itm.Augment(i).Done = True)
                        .Visible = True
                    End With
                    With .lnkAugment(lngIndex)
                        .Caption = strCaption
                        .ForeColor = lngColor
                        .Visible = True
                    End With
                    With itm.AugSlot(lngIndex)
                        .Color = i
                        .Augment = itm.Augment(i).Augment
                        .Variation = itm.Augment(i).Variation
                        .Scaling = itm.Augment(i).Scaling
                        If .Augment = 0 Or .Variation = 0 Then .Done = True Else .Done = itm.Augment(i).Done
                        .HasError = blnError
                        .Visible = True
                    End With
                    lngIndex = lngIndex + 1
                End If
            End If
        Next
    End With
End Sub

Private Sub ShowBaseItem()
    Dim lngIndex As Long
    
    If Not (itm.Crafted And itm.Scales) Then Exit Sub
    Select Case itm.ML
        Case Is < 4: lngIndex = 0
        Case 4 To 9: lngIndex = 1
        Case 10 To 15: lngIndex = 2
        Case 16 To 21: lngIndex = 3
        Case Is > 21: lngIndex = 4
    End Select
    With UserControl.lblBaseItem
        .Caption = itm.Scaling(lngIndex)
        .Visible = True
    End With
End Sub


' ************* DONE? *************


Public Sub CheckAllDone()
    Dim blnDone As Boolean
    Dim blnVisible As Boolean
    Dim i As Long
    
    With UserControl
        Do
            If itm.Crafted Then
                blnVisible = True
                If Not itm.MLDone Then Exit Do
            End If
            For i = 0 To 2
                If itm.Effect(i).Visible Then
                    blnVisible = True
                    If Not itm.Effect(i).Done Then Exit Do
                End If
                If itm.AugSlot(i).Visible Then
                    blnVisible = True
                    If itm.AugSlot(i).HasError = False And Not itm.AugSlot(i).Done Then Exit Do
                End If
            Next
            If itm.EldritchRitual <> 0 Then
                blnVisible = True
                If itm.EldritchDone = False Then Exit Do
            End If
            blnDone = True
        Loop Until True
        itm.AllDone = blnDone
        .usrchkAll.Value = itm.AllDone
        .usrchkAll.Visible = blnVisible
        .lblDone(0).Visible = blnVisible
        .lblDone(1).Visible = blnVisible
    End With
End Sub

Private Sub usrchkML_UserChange()
    itm.MLDone = UserControl.usrchkML.Value
    CheckAllDone
    RaiseEvent MLDone(itm.MLDone)
End Sub

Private Sub usrchkEffect_UserChange(Index As Integer)
    itm.Effect(Index).Done = UserControl.usrchkEffect(Index).Value
    CheckAllDone
    RaiseEvent EffectDone(Index, itm.Effect(Index).Done)
End Sub

Private Sub usrchkAugment_UserChange(Index As Integer)
    Dim enColor As AugmentColorEnum
    Dim blnDone As Boolean
    
    If itm.AugSlot(Index).Augment = 0 Or itm.AugSlot(Index).Variation = 0 Then
        UserControl.usrchkAugment(Index).Value = True
    Else
        enColor = itm.AugSlot(Index).Color
        blnDone = UserControl.usrchkAugment(Index).Value
        itm.AugSlot(Index).Done = blnDone
        itm.Augment(enColor).Done = blnDone
        CheckAllDone
        RaiseEvent AugmentDone(enColor, blnDone)
    End If
End Sub

Private Sub usrchkEldritch_UserChange()
    itm.EldritchDone = UserControl.usrchkEldritch.Value
    CheckAllDone
    RaiseEvent EldritchDone(itm.EldritchDone)
End Sub

Private Sub lblDone_Click(Index As Integer)
    ToggleDone
End Sub

Private Sub lblDone_DblClick(Index As Integer)
    ToggleDone
End Sub

Private Sub usrchkAll_UserChange()
    AllDone
End Sub

Private Sub ToggleDone()
    With UserControl.usrchkAll
        .Value = Not .Value
    End With
    AllDone
End Sub

Private Sub AllDone()
    Dim blnValue As Boolean
    Dim enColor As AugmentColorEnum
    Dim i As Long
    
    With UserControl
        blnValue = .usrchkAll.Value
        If itm.Crafted = True And itm.MLDone <> blnValue Then
            itm.MLDone = blnValue
            .usrchkML.Value = blnValue
            RaiseEvent MLDone(blnValue)
        End If
        For i = 0 To 2
            If itm.Effect(i).Visible = True And itm.Effect(i).Done <> blnValue Then
                itm.Effect(i).Done = blnValue
                .usrchkEffect(i).Value = blnValue
                RaiseEvent EffectDone(i, blnValue)
            End If
            If itm.AugSlot(i).Visible = True And itm.AugSlot(i).Done <> blnValue Then
                If itm.AugSlot(i).Augment <> 0 And itm.AugSlot(i).Variation <> 0 Then
                    enColor = itm.AugSlot(i).Color
                    itm.AugSlot(i).Done = blnValue
                    itm.Augment(enColor).Done = blnValue
                    .usrchkAugment(i).Value = blnValue
                    RaiseEvent AugmentDone(enColor, blnValue)
                End If
            End If
        Next
        If itm.EldritchRitual <> 0 And itm.EldritchDone <> blnValue Then
            itm.EldritchDone = blnValue
            .usrchkEldritch.Value = blnValue
            RaiseEvent EldritchDone(blnValue)
        End If
    End With
End Sub


' ************* INTERFACE *************


Public Property Get ML() As Long
    ML = itm.ML
End Property

Public Property Let ML(plngML As Long)
    mblnOverride = True
    UserControl.usrspnML.Value = plngML
    mblnOverride = False
    ChangeML UserControl.usrspnML.Value
End Property

Private Sub usrspnML_Change()
    If mblnOverride Then Exit Sub
    ChangeML UserControl.usrspnML.Value
    RaiseEvent LevelChange(itm.ML)
End Sub

Private Sub ChangeML(plngML As Long)
    Dim enAffix As AffixEnum
    
    itm.ML = plngML
    ShowEffects
    ShowAugments
    ShowBaseItem
    CheckAllDone ' All Done can change if we scroll past ML10 and insightful/extra effects are the only ones not done
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl.usrIcon.Active = False
End Sub

Private Sub usrIcon_Click()
    Dim strAugment As String
    
    If itm.Crafted = False And Len(itm.Named) Then
        xp.OpenURL WikiSearch(itm.Named)
    Else
        strAugment = GearsetAugmentToString(itm.Augment)
        RaiseEvent GearClick(itm.Gear, itm.ML, itm.Effect(aePrefix).Shard, itm.Effect(aeSuffix).Shard, itm.Effect(aeExtra).Shard, strAugment, itm.EldritchRitual)
        UserControl.tmrDeactivate.Enabled = True
    End If
End Sub

Private Sub tmrDeactivate_Timer()
    UserControl.tmrDeactivate.Enabled = False
    UserControl.usrIcon.Active = False
End Sub

Private Sub lnkEffect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkEffect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
    If itm.Effect(Index).Shard Then OpenShard db.Shard(itm.Effect(Index).Shard).ShardName
End Sub

Private Sub lnkEffect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCursor LoadCursor(0, CURSOR_HAND)
End Sub

Private Sub lnkAugment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ValidAugment(Index) Then xp.SetMouseCursor mcHand
End Sub

Private Sub lnkAugment_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not ValidAugment(Index) Then Exit Sub
    xp.SetMouseCursor mcHand
    OpenAugment itm.AugSlot(Index).Augment, itm.AugSlot(Index).Variation, itm.AugSlot(Index).Scaling
End Sub

Private Sub lnkAugment_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ValidAugment(Index) Then xp.SetMouseCursor mcHand
End Sub

Private Function ValidAugment(Index As Integer) As Boolean
    If itm.AugSlot(Index).Augment = 0 Then Exit Function
    If itm.AugSlot(Index).Variation = 0 Then Exit Function
'    If itm.Augment(enColor).Scaling = 0 Then Exit Function
    ValidAugment = True
End Function

Public Property Get Spinhwnd() As Long
    Spinhwnd = UserControl.usrspnML.hwnd
End Property

Public Sub SpinWheel(plngValue As Long)
    UserControl.usrspnML.WheelScroll plngValue
End Sub

Private Sub picEldritch_Click()
    OpenForm "frmEldritch"
End Sub

