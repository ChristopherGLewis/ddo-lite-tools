VERSION 5.00
Begin VB.Form frmMenuEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Editor"
   ClientHeight    =   5940
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7956
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenuEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7956
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   180
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4632
      Left            =   60
      ScaleHeight     =   4632
      ScaleWidth      =   7212
      TabIndex        =   2
      Top             =   780
      Width           =   7212
      Begin VB.ListBox lstMenu 
         Height          =   2640
         ItemData        =   "frmMenuEditor.frx":000C
         Left            =   2532
         List            =   "frmMenuEditor.frx":000E
         TabIndex        =   8
         Top             =   300
         Width           =   2412
      End
      Begin VB.TextBox txtCaption 
         Enabled         =   0   'False
         Height          =   324
         Left            =   1380
         TabIndex        =   10
         Top             =   3180
         Width           =   2892
      End
      Begin VB.TextBox txtTarget 
         Enabled         =   0   'False
         Height          =   324
         Left            =   1380
         TabIndex        =   12
         Top             =   3600
         Width           =   4932
      End
      Begin VB.TextBox txtParam 
         Enabled         =   0   'False
         Height          =   324
         Left            =   1380
         TabIndex        =   15
         Top             =   4020
         Visible         =   0   'False
         Width           =   4932
      End
      Begin VB.CheckBox chkTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "..."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   324
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3600
         Width           =   432
      End
      Begin VB.Frame fraAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1932
         Left            =   360
         TabIndex        =   3
         Top             =   180
         Width           =   1872
         Begin VB.CheckBox chkNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Link"
            ForeColor       =   &H80000008&
            Height          =   372
            Index           =   0
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   420
            Width           =   1272
         End
         Begin VB.CheckBox chkNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Shortcut"
            ForeColor       =   &H80000008&
            Height          =   372
            Index           =   1
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   840
            Width           =   1272
         End
         Begin VB.CheckBox chkNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Separator"
            ForeColor       =   &H80000008&
            Height          =   372
            Index           =   2
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1260
            Width           =   1272
         End
         Begin VB.Label lblAddNew 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Add New"
            ForeColor       =   &H80000008&
            Height          =   216
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   804
         End
         Begin VB.Shape shpAdd 
            Height          =   1812
            Left            =   0
            Top             =   120
            Width           =   1872
         End
      End
      Begin VB.CheckBox chkButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   0
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   300
         Width           =   1212
      End
      Begin VB.CheckBox chkButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   1
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1212
      End
      Begin VB.CheckBox chkButton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Help"
         ForeColor       =   &H80000008&
         Height          =   372
         Index           =   2
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1260
         Width           =   1212
      End
      Begin VB.Label lblCaption 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Caption:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   420
         TabIndex        =   9
         Top             =   3216
         Width           =   828
      End
      Begin VB.Label lblTarget 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Target:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   588
         TabIndex        =   11
         Top             =   3636
         Width           =   660
      End
      Begin VB.Label lblParam 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Parameter:"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   216
         Left            =   240
         TabIndex        =   14
         Top             =   4056
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Image imgArrow 
         Height          =   312
         Index           =   1
         Left            =   4980
         Picture         =   "frmMenuEditor.frx":0010
         Stretch         =   -1  'True
         ToolTipText     =   "HIgher Priority"
         Top             =   360
         Width           =   300
      End
      Begin VB.Image imgArrow 
         Height          =   312
         Index           =   2
         Left            =   4980
         Picture         =   "frmMenuEditor.frx":080A
         Stretch         =   -1  'True
         ToolTipText     =   "Lower Priority"
         Top             =   720
         Width           =   300
      End
      Begin VB.Image imgArrow 
         Height          =   312
         Index           =   0
         Left            =   4980
         Picture         =   "frmMenuEditor.frx":1004
         Stretch         =   -1  'True
         ToolTipText     =   "Clear All"
         Top             =   1440
         Width           =   300
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   324
      Left            =   1440
      TabIndex        =   1
      Top             =   180
      Width           =   2892
   End
   Begin VB.Line linTitle 
      X1              =   -120
      X2              =   7500
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Title:"
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   864
      TabIndex        =   0
      Top             =   216
      Width           =   444
   End
End
Attribute VB_Name = "frmMenuEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mctlChild As Control

Private mblnOverride As Boolean


' ************* FORM *************


Private Sub Form_Load()
    cfg.Configure Me
    ShowTitle
    LoadData
    EnableArrows
    Select Case gtypMenu.Selected
        Case 0
        Case -1: 'If gtypMenu.Selected = -1 And Me.txtTitle.Visible = True Then me.txtTitle.SetFocus
        Case Else: Me.lstMenu.ListIndex = gtypMenu.Selected - 1
    End Select
    Me.tmrFocus.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cfg.SavePosition Me
End Sub

Private Sub ShowTitle()
    Me.lblTitle.Visible = gtypMenu.LinkList
    Me.txtTitle.Visible = gtypMenu.LinkList
    Me.linTitle.Visible = gtypMenu.LinkList
    If gtypMenu.LinkList Then
        Me.picMenu.Top = Me.linTitle.Y1 + PixelY
    Else
        Me.picMenu.Top = 0
        Me.Caption = gtypMenu.Title
    End If
    Me.Height = Me.picMenu.Top + Me.picMenu.Height + Me.Height - Me.ScaleHeight
    Me.Width = Me.picMenu.Width + Me.Width - Me.ScaleWidth
End Sub

Private Sub LoadData()
    Dim i As Long
    
    ListboxClear Me.lstMenu
    gtypMenu.Accepted = False
    Me.txtTitle.Text = gtypMenu.Title
    With gtypMenu
        For i = 1 To .Commands
            ListboxAddItem Me.lstMenu, .Command(i).Caption, i
        Next
    End With
End Sub

Private Sub tmrFocus_Timer()
    Me.tmrFocus.Enabled = False
    Select Case gtypMenu.Selected
        Case 0: Me.lstMenu.SetFocus
        Case -1: Me.txtTitle.SetFocus
        Case Else: Me.txtTarget.SetFocus
    End Select
End Sub


' ************* ADD *************


Private Sub chkNew_Click(Index As Integer)
    Dim typNew As MenuCommandType
    
    If UncheckButton(Me.chkNew(Index), mblnOverride) Then Exit Sub
    typNew.Style = Index
    Select Case Index
        Case 0: typNew.Caption = "New Link"
        Case 1: typNew.Caption = "New Shortcut"
        Case 2: typNew.Caption = "-"
    End Select
    With gtypMenu
        .Commands = .Commands + 1
        ReDim Preserve .Command(1 To .Commands)
        .Command(.Commands) = typNew
    End With
    ShowMenu
    Me.lstMenu.ListIndex = Me.lstMenu.ListCount - 1
    If Index < 2 Then
        With Me.txtCaption
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

Private Sub ShowMenu()
    Dim i As Long
    
    ListboxClear Me.lstMenu
    With gtypMenu
        For i = 1 To .Commands
            ListboxAddItem Me.lstMenu, .Command(i).Caption, i
        Next
    End With
    EnableControls
End Sub

Private Sub lstMenu_Click()
    EnableControls
End Sub


' ************* EDIT *************


Private Sub txtTitle_GotFocus()
    TextboxGotFocus Me.txtTitle
End Sub

Private Sub txtTitle_Change()
    gtypMenu.Title = Me.txtTitle.Text
End Sub

Private Sub EnableControls()
    Dim blnEnabled As Boolean
    Dim strCaption As String
    
    If Me.lstMenu.ListIndex = -1 Then
        Me.lblCaption.Enabled = False
        SetTextbox Me.txtCaption, "", False, True
        SetLabel Me.lblTarget, "Target:", False, True
        SetTextbox Me.txtTarget, "", False, True
        Me.chkTarget.Enabled = False
        Me.chkTarget.Visible = True
        SetLabel Me.lblParam, "", False, True
        SetTextbox Me.txtParam, "", False, True
    Else
        With gtypMenu.Command(Me.lstMenu.ListIndex + 1)
            ' Caption
            blnEnabled = (.Style <> mceSeparator)
            Me.lblCaption.Enabled = blnEnabled
            SetTextbox Me.txtCaption, .Caption, blnEnabled, True
            If .Style = mceLink Then strCaption = "URL:" Else strCaption = "Target:"
            SetLabel Me.lblTarget, strCaption, blnEnabled, blnEnabled
            SetTextbox Me.txtTarget, .Target, blnEnabled, blnEnabled
            blnEnabled = (.Style = mceShortcut)
            Me.chkTarget.Enabled = blnEnabled
            Me.chkTarget.Visible = blnEnabled
            SetLabel Me.lblParam, "", blnEnabled, blnEnabled
            SetTextbox Me.txtParam, .Param, blnEnabled, blnEnabled
        End With
    End If
    EnableArrows
End Sub

Private Sub EnableArrows()
    With Me.lstMenu
        EnableArrow Me.imgArrow(0), (.ListIndex <> -1)
        EnableArrow Me.imgArrow(1), (.ListIndex > 0)
        EnableArrow Me.imgArrow(2), (.ListIndex > -1 And .ListIndex < .ListCount - 1)
    End With
End Sub

Private Sub SetLabel(plbl As Label, pstrCaption As String, pblnEnabled As Boolean, pblnVisible As Boolean)
    If Len(pstrCaption) And plbl.Caption <> pstrCaption Then plbl.Caption = pstrCaption
    If plbl.Enabled <> pblnEnabled Then plbl.Enabled = pblnEnabled
    If plbl.Visible <> pblnVisible Then plbl.Visible = pblnVisible
End Sub

Private Sub SetTextbox(ptxt As TextBox, pstrText As String, pblnEnabled As Boolean, pblnVisible As Boolean)
    mblnOverride = True
    If ptxt.Text <> pstrText Then ptxt.Text = pstrText
    If ptxt.Enabled <> pblnEnabled Then ptxt.Enabled = pblnEnabled
    If ptxt.Visible <> pblnVisible Then ptxt.Visible = pblnVisible
    mblnOverride = False
End Sub

Private Sub txtCaption_GotFocus()
    TextboxGotFocus Me.txtCaption
End Sub

Private Sub txtCaption_Change()
    Dim lngIndex As Long
    
    If mblnOverride Then Exit Sub
    lngIndex = Me.lstMenu.ListIndex
    If lngIndex = -1 Then Exit Sub
    Me.lstMenu.List(lngIndex) = Me.txtCaption.Text
    gtypMenu.Command(lngIndex + 1).Caption = Me.txtCaption.Text
End Sub

Private Sub txtTarget_GotFocus()
    TextboxGotFocus Me.txtTarget
End Sub

Private Sub txtTarget_Change()
    Dim lngIndex As Long
    
    If mblnOverride Then Exit Sub
    lngIndex = Me.lstMenu.ListIndex
    If lngIndex = -1 Then Exit Sub
    gtypMenu.Command(lngIndex + 1).Target = Me.txtTarget.Text
End Sub

Private Sub chkTarget_Click()
    Dim strStrip As String
    Dim strFile As String
    
    If UncheckButton(Me.chkTarget, mblnOverride) Then Exit Sub
    strFile = XP.ShowOpenDialog(App.Path, "All Files (*.*)|*.*", "*.*")
    strStrip = App.Path & "\"
    If Left(strFile, Len(strStrip)) = strStrip Then strFile = Mid(strFile, Len(strStrip) + 1)
    Me.txtTarget.Text = strFile
End Sub

Private Sub txtParam_Change()
    Dim lngIndex As Long
    
    If mblnOverride Then Exit Sub
    lngIndex = Me.lstMenu.ListIndex
    If lngIndex = -1 Then Exit Sub
    gtypMenu.Command(lngIndex + 1).Param = Me.txtParam.Text
End Sub


' ************* ORDER *************


Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetArrowIcon Me.imgArrow(Index), asePressed
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetArrowIcon Me.imgArrow(Index), aseEnabled
    ButtonClick Index
End Sub

Private Sub imgArrow_DblClick(Index As Integer)
'    If Index > 0 Then ButtonClick Index
End Sub

Private Sub ButtonClick(Index As Integer)
    Select Case Index
        Case 0: DeleteRow
        Case 1: MoveRow -1
        Case 2: MoveRow 1
    End Select
End Sub

Private Sub MoveRow(plngIncrement As Long)
    Dim typSwap As MenuCommandType
    Dim strSwap As String
    Dim lngOld As Long
    Dim lngNew As Long
    
    lngOld = Me.lstMenu.ListIndex
    lngNew = lngOld + plngIncrement
    If lngOld < 0 Or lngOld > Me.lstMenu.ListCount - 1 Or lngNew < 0 Or lngNew > Me.lstMenu.ListCount - 1 Then Exit Sub
    With gtypMenu
        typSwap = .Command(lngOld + 1)
        .Command(lngOld + 1) = .Command(lngNew + 1)
        .Command(lngNew + 1) = typSwap
    End With
    With Me.lstMenu
        strSwap = .List(lngOld)
        .List(lngOld) = .List(lngNew)
        .List(lngNew) = strSwap
    End With
    Me.lstMenu.ListIndex = lngNew
End Sub

Private Sub DeleteRow()
    Dim lngIndex As Long
    Dim i As Long
    
    lngIndex = Me.lstMenu.ListIndex
    If lngIndex = -1 Then Exit Sub
    With gtypMenu
        For i = lngIndex + 1 To Me.lstMenu.ListCount - 1
            .Command(i) = .Command(i + 1)
        Next
        .Commands = .Commands - 1
        If .Commands = 0 Then Erase .Command Else ReDim Preserve .Command(1 To .Commands)
    End With
    With Me.lstMenu
        .RemoveItem lngIndex
        If lngIndex > .ListCount - 1 Then lngIndex = .ListCount - 1
        .ListIndex = lngIndex
        If .ListCount = 0 Then EnableControls
    End With
End Sub


' ************* OK *************


Private Sub chkButton_Click(Index As Integer)
    If UncheckButton(Me.chkButton(Index), mblnOverride) Then Exit Sub
    Select Case Me.chkButton(Index).Caption
        Case "OK"
            gtypMenu.Accepted = True
            Unload Me
        Case "Cancel"
            gtypMenu.Accepted = False
            Unload Me
        Case "Help"
            ShowHelp "Menu Editor"
    End Select
End Sub

