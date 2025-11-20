VERSION 5.00
Begin VB.Form frmColor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color"
   ClientHeight    =   3468
   ClientLeft      =   36
   ClientTop       =   300
   ClientWidth     =   5028
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3468
   ScaleWidth      =   5028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   1
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2880
      Width           =   1092
   End
   Begin VB.CheckBox chkButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   0
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2880
      Width           =   1092
   End
   Begin VB.ComboBox cboMaterial 
      Height          =   312
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   1632
   End
   Begin VB.PictureBox picPicker 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   324
      Left            =   2220
      ScaleHeight     =   300
      ScaleWidth      =   588
      TabIndex        =   0
      Top             =   1080
      Width           =   612
   End
   Begin VB.TextBox txtRGB 
      Appearance      =   0  'Flat
      Height          =   324
      Index           =   2
      Left            =   4020
      MaxLength       =   3
      TabIndex        =   13
      Top             =   2040
      Width           =   792
   End
   Begin VB.TextBox txtRGB 
      Appearance      =   0  'Flat
      Height          =   324
      Index           =   1
      Left            =   3120
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2040
      Width           =   792
   End
   Begin VB.TextBox txtRGB 
      Appearance      =   0  'Flat
      Height          =   324
      Index           =   0
      Left            =   2220
      MaxLength       =   3
      TabIndex        =   11
      Top             =   2040
      Width           =   792
   End
   Begin VB.TextBox txtHTML 
      Appearance      =   0  'Flat
      Height          =   324
      Left            =   2220
      MaxLength       =   7
      TabIndex        =   2
      Top             =   120
      Width           =   1572
   End
   Begin VB.ComboBox cboSystem 
      Height          =   312
      ItemData        =   "frmColor.frx":000C
      Left            =   2220
      List            =   "frmColor.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   2592
   End
   Begin Colors.userSpinner usrspnMaterial 
      Height          =   300
      Left            =   3900
      TabIndex        =   5
      Top             =   600
      Width           =   912
      _ExtentX        =   1609
      _ExtentY        =   529
      ShowZero        =   0   'False
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderColor     =   0
      BorderInterior  =   -2147483631
      Position        =   0
      Enabled         =   -1  'True
      DisabledColor   =   -2147483631
   End
   Begin VB.Label lblMaterial 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Material Color:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   180
      TabIndex        =   3
      Top             =   636
      Width           =   1932
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   4020
      TabIndex        =   16
      Top             =   2400
      Width           =   792
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   3120
      TabIndex        =   15
      Top             =   2400
      Width           =   792
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   2220
      TabIndex        =   14
      Top             =   2400
      Width           =   792
   End
   Begin VB.Label lblPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Color Picker:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   1104
      Width           =   1932
   End
   Begin VB.Label lblRGB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RGB color values:"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   10
      Top             =   2064
      Width           =   1932
   End
   Begin VB.Label lblHTML 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "HTML color code:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   180
      TabIndex        =   1
      Top             =   144
      Width           =   1932
   End
   Begin VB.Label lblColorPicker 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(Click box to open)"
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   2940
      TabIndex        =   7
      Top             =   1104
      Width           =   1728
   End
   Begin VB.Label lblSystem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "System Color:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   180
      TabIndex        =   8
      Top             =   1584
      Width           =   1932
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Written by Ellis Dee
' This is a generic form, not specific to this project.
' It requires a global variable (glngActiveColor) to communicate the color changes to the project
' It also makes extensive use of ComboAddItem(), found in basUtils (another generic module)
' I should probably add in checks to prevent RGB values going over 255, but RGB() doesn't seem to mind so whatever.
' Recent addition: Google Material Design color palettes, which now uses a spinner custom control.
Option Explicit

Private mlngMaterial(18, 9) As Long

Private mlngColor As Long

Private mblnOverride As Boolean

Private Sub Form_Load()
    cfg.RefreshColors Me
    LoadSystemColors
    LoadMaterialColors
    ShowColor glngActiveColor
    If Not xp.DebugMode Then Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not xp.DebugMode Then Call WheelUnHook(Me.hWnd)
End Sub

' Thanks to bushmobile of VBForums.com
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim lngValue As Long
    
    If Rotation < 0 Then lngValue = -3 Else lngValue = 3
    If IsOver(Me.usrspnMaterial.hWnd, Xpos, Ypos) Then Me.usrspnMaterial.WheelScroll lngValue
End Sub


' ************* CONTROLS *************


Private Sub chkButton_Click(Index As Integer)
    If mblnOverride Then
        mblnOverride = False
        If Index = 0 Then glngActiveColor = mlngColor
        Unload Me
    Else
        mblnOverride = True
        Me.chkButton(Index).Value = vbUnchecked
    End If
End Sub

Private Sub txtHTML_GotFocus()
    With Me.txtHTML
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtHTML_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57 ' 0-9
        Case 65 To 70 ' A-F
        Case 97 To 102: KeyAscii = KeyAscii - 32 ' a-f, convert to A-F
        Case 35 ' #
        Case 8 ' backspace
        Case 13 ' enter
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub txtHTML_Change()
    If mblnOverride Then Exit Sub
    If xp.HexColorIsValid(Me.txtHTML.Text) Then ShowColor xp.HexToColor(Me.txtHTML.Text)
End Sub

Private Sub cboMaterial_Click()
    If mblnOverride Then Exit Sub
    If Me.cboMaterial.ListIndex <> -1 Then ShowColor mlngMaterial(Me.cboMaterial.ListIndex, 10 - Me.usrspnMaterial.Value)
End Sub

Private Sub usrspnMaterial_Change()
    If mblnOverride Then Exit Sub
    If Me.cboMaterial.ListIndex <> -1 Then ShowColor mlngMaterial(Me.cboMaterial.ListIndex, 10 - Me.usrspnMaterial.Value)
End Sub

Private Sub txtRGB_GotFocus(Index As Integer)
    With Me.txtRGB(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtRGB_KeyPress(Index As Integer, KeyAscii As Integer)
    TextboxNumericKeypress KeyAscii, False, False, False
End Sub

Private Sub txtRGB_Change(Index As Integer)
    Dim lngColor As Long
    
    If mblnOverride Then Exit Sub
    lngColor = RGB(Val(Me.txtRGB(0).Text), Val(Me.txtRGB(1).Text), Val(Me.txtRGB(2).Text))
    ShowColor lngColor
End Sub

Private Sub cboSystem_Click()
    If mblnOverride Or Me.cboSystem.ListIndex < 1 Then Exit Sub
    ShowColor Me.cboSystem.ItemData(Me.cboSystem.ListIndex)
End Sub

Private Sub picPicker_Click()
    Dim lngColor As Long
    
    lngColor = xp.ShowColorDialog(Me.hWnd, True, lngColor)
    If lngColor <> -1 Then ShowColor lngColor
End Sub


' ************* GENERAL *************


Private Sub LoadSystemColors()
    Me.cboSystem.ListIndex = -1
    Me.cboSystem.Clear
    ComboAddItem Me.cboSystem, vbNullString, 0
    
    ComboAddItem Me.cboSystem, "Button Text", vbButtonText
    ComboAddItem Me.cboSystem, "Button Face", vbButtonFace
    ComboAddItem Me.cboSystem, "Button Shadow", vbButtonShadow
    
    ComboAddItem Me.cboSystem, "Window Text", vbWindowText
    ComboAddItem Me.cboSystem, "Window Background", vbWindowBackground
    ComboAddItem Me.cboSystem, "Window Frame", vbWindowFrame
    
    ComboAddItem Me.cboSystem, "Application Workspace", vbApplicationWorkspace
    
    ComboAddItem Me.cboSystem, "Gray Text", vbGrayText
    ComboAddItem Me.cboSystem, "Inactive Caption Text", vbInactiveCaptionText
    ComboAddItem Me.cboSystem, "Menu Text", vbMenuText
    
    ComboAddItem Me.cboSystem, "Highlight Text", vbHighlightText
    ComboAddItem Me.cboSystem, "Highlight", vbHighlight
    
    ComboAddItem Me.cboSystem, "Active Border", vbActiveBorder
    ComboAddItem Me.cboSystem, "Inactive Border", vbInactiveBorder
    
    ComboAddItem Me.cboSystem, "Info Text", vbInfoText
    ComboAddItem Me.cboSystem, "Info Background", vbInfoBackground
    
    ComboAddItem Me.cboSystem, "TitleBar Text", vbTitleBarText
    ComboAddItem Me.cboSystem, "Active Title Bar", vbActiveTitleBar
    ComboAddItem Me.cboSystem, "Inactive Title Bar", vbInactiveTitleBar
    
    ComboAddItem Me.cboSystem, "Menu Bar", vbMenuBar
    ComboAddItem Me.cboSystem, "Scroll Bars", vbScrollBars
    ComboAddItem Me.cboSystem, "Desktop", vbDesktop
    
    ComboAddItem Me.cboSystem, "3D Face", vb3DFace
    ComboAddItem Me.cboSystem, "3D Light", vb3DLight
    ComboAddItem Me.cboSystem, "3D Shadow", vb3DShadow
    ComboAddItem Me.cboSystem, "3D Dark Shadow", vb3DDKShadow
    ComboAddItem Me.cboSystem, "3D Highlight", vb3DHighlight
End Sub

Private Sub LoadMaterialColors()
    Dim i As Long
    
    Me.cboMaterial.Clear
    AddMaterialColor i, "Amber,FFF8E1,FFECB3,FFE082,FFD54F,FFCA28,FFC107,FFB300,FFA000,FF8F00,FF6F00"
    AddMaterialColor i, "Blue,E3F2FD,BBDEFB,90CAF9,64B5F6,42A5F5,2196F3,1E88E5,1976D2,1565C0,0D47A1"
    AddMaterialColor i, "Blue Grey,ECEFF1,CFD8DC,B0BEC5,90A4AE,78909C,607D8B,546E7A,455A64,37474F,263238"
    AddMaterialColor i, "Brown,EFEBE9,D7CCC8,BCAAA4,A1887F,8D6E63,795548,6D4C41,5D4037,4E342E,3E2723"
    AddMaterialColor i, "Cyan,E0F7FA,B2EBF2,80DEEA,4DD0E1,26C6DA,00BCD4,00ACC1,0097A7,00838F,006064"
    AddMaterialColor i, "Deep Orange,FBE9E7,FFCCBC,FFAB91,FF8A65,FF7043,FF5722,F4511E,E64A19,D84315,BF360C"
    AddMaterialColor i, "Deep Purple,EDE7F6,D1C4E9,B39DDB,9575CD,7E57C2,673AB7,5E35B1,512DA8,4527A0,311B92"
    AddMaterialColor i, "Green,E8F5E9,C8E6C9,A5D6A7,81C784,66BB6A,4CAF50,43A047,388E3C,2E7D32,1B5E20"
    AddMaterialColor i, "Grey,FAFAFA,F5F5F5,EEEEEE,E0E0E0,BDBDBD,9E9E9E,757575,616161,424242,212121"
    AddMaterialColor i, "Indigo,E8EAF6,C5CAE9,9FA8DA,7986CB,5C6BC0,3F51B5,3949AB,303F9F,283593,1A237E"
    AddMaterialColor i, "Light Blue,E1F5FE,B3E5FC,81D4FA,4FC3F7,29B6F6,03A9F4,039BE5,0288D1,0277BD,01579B"
    AddMaterialColor i, "Light Green,F1F8E9,DCEDC8,C5E1A5,AED581,9CCC65,8BC34A,7CB342,689F38,558B2F,33691E"
    AddMaterialColor i, "Lime,F9FBE7,F0F4C3,E6EE9C,DCE775,D4E157,CDDC39,C0CA33,AFB42B,9E9D24,827717"
    AddMaterialColor i, "Orange,FFF3E0,FFE0B2,FFCC80,FFB74D,FFA726,FF9800,FB8C00,F57C00,EF6C00,E65100"
    AddMaterialColor i, "Pink,FCE4EC,F8BBD0,F48FB1,F06292,EC407A,E91E63,D81B60,C2185B,AD1457,880E4F"
    AddMaterialColor i, "Purple,F3E5F5,E1BEE7,CE93D8,BA68C8,AB47BC,9C27B0,8E24AA,7B1FA2,6A1B9A,4A148C"
    AddMaterialColor i, "Red,FFEBEE,FFCDD2,EF9A9A,E57373,EF5350,F44336,E53935,D32F2F,C62828,B71C1C"
    AddMaterialColor i, "Teal,E0F2F1,B2DFDB,80CBC4,4DB6AC,26A69A,009688,00897B,00796B,00695C,004D40"
    AddMaterialColor i, "Yellow,FFFDE7,FFF9C4,FFF59D,FFF176,FFEE58,FFEB3B,FDD835,FBC02D,F9A825,F57F17"
End Sub

Private Sub AddMaterialColor(plngIndex As Long, pstrRaw As String)
    Dim strToken() As String
    Dim i As Long
    
    strToken = Split(pstrRaw, ",")
    Me.cboMaterial.AddItem strToken(0)
    For i = 0 To 9
        mlngMaterial(plngIndex, i) = xp.HexToColor(strToken(i + 1))
    Next
    plngIndex = plngIndex + 1
End Sub

Private Sub ShowColor(plngColor As Long)
    Dim lngColor As Long
    Dim strHex As String
    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    Dim i As Long
    
    mblnOverride = True
    mlngColor = plngColor
    If mlngColor < 0 Then
        ComboSetValue Me.cboSystem, plngColor
        lngColor = xp.SystemColorRGB(plngColor)
    Else
        Me.cboSystem.ListIndex = 0
        lngColor = mlngColor
    End If
    xp.ColorToRGB lngColor, lngRed, lngGreen, lngBlue
    Me.txtHTML.Text = "#" & xp.ColorToHex(lngColor)
    FindMaterialColor lngColor
    Me.txtRGB(0).Text = lngRed
    Me.txtRGB(1).Text = lngGreen
    Me.txtRGB(2).Text = lngBlue
    ShowPickerColor
    mblnOverride = False
    mlngColor = plngColor
End Sub

Private Sub ShowPickerColor()
    With Me.picPicker
        Me.picPicker.Line (0, 0)-(.ScaleWidth - PixelX, .ScaleHeight - PixelY), vbBlack, B
        Me.picPicker.Line (PixelX, PixelY)-(.ScaleWidth - PixelX * 2, .ScaleHeight - PixelY * 2), mlngColor, BF
    End With
End Sub

Private Sub FindMaterialColor(plngColor As Long)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To 18
        For j = 0 To 9
            If mlngMaterial(i, j) = plngColor Then
                Me.cboMaterial.ListIndex = i
                Me.usrspnMaterial.Value = 10 - j
                Exit Sub
            End If
        Next
    Next
    Me.cboMaterial.ListIndex = -1
End Sub
