VERSION 5.00
Begin VB.UserControl userTextbox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin VB.Timer tmrSaveTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1380
      Top             =   2280
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   1020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   1512
   End
   Begin VB.Shape shpBorder 
      Height          =   1692
      Left            =   420
      Top             =   360
      Visible         =   0   'False
      Width           =   2592
   End
End
Attribute VB_Name = "userTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngForeColor As Long
Private mlngBackColor As Long
Private mlngBorderColor As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private mblnOverride As Boolean


' ************* PROPERTY BAG *************


' The property bag object persists settings in design mode (as opposed to during runtime)
' This function only fires once ever per control, when you add it to a form in design mode
Private Sub UserControl_InitProperties()
    mlngForeColor = vbWindowText
    mlngBackColor = vbButtonFace
    UserControl.tmrSaveTimer.Interval = 2000
End Sub

' In design mode, when you close a form, the developer's property settings are stored to a mysterious "property bag"
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ForeColor", mlngForeColor, vbWindowText
    PropBag.WriteProperty "BackColor", mlngBackColor, vbButtonFace
    PropBag.WriteProperty "TimerInterval", UserControl.tmrSaveTimer.Interval, 2000
End Sub

' Next time you open the form in design view, each control instance retreives its developer-chosen properties from the property bag
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mlngForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    mlngBackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    UserControl.tmrSaveTimer.Interval = PropBag.ReadProperty("TimerInterval", 2000)
    DrawControl
End Sub

Private Sub UserControl_GotFocus()
On Error Resume Next
    UserControl.txt.SetFocus
End Sub

Private Sub UserControl_Resize()
    DrawControl
End Sub

Private Sub DrawControl()
    Dim PixelX As Long
    Dim PixelY As Long
    
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    With UserControl
        .BackColor = mlngBackColor
        .shpBorder.BorderColor = BorderColor()
        .shpBorder.Move 0, 0, .ScaleWidth, .ScaleHeight
        .txt.ForeColor = mlngForeColor
        .txt.BackColor = mlngBackColor
        .txt.Move PixelX * 4, PixelY * 4, .ScaleWidth - PixelX * 5, .ScaleHeight - PixelY * 5
    End With
End Sub

Private Sub DrawControlOld()
    Dim PixelX As Long
    Dim PixelY As Long
    
    PixelX = Screen.TwipsPerPixelX
    PixelY = Screen.TwipsPerPixelY
    With UserControl
        .BackColor = mlngBackColor
        .shpBorder.BorderColor = BorderColor()
        .shpBorder.Move 0, 0, .ScaleWidth, .ScaleHeight
        .txt.ForeColor = mlngForeColor
        .txt.BackColor = mlngBackColor
        .txt.Move PixelX * 4, PixelY * 4, .ScaleWidth - PixelX * 5, .ScaleHeight - PixelY * 5
    End With
End Sub

Public Property Get Text() As String
    Text = UserControl.txt.Text
End Property

Public Property Let Text(ByVal pstrText As String)
    mblnOverride = True
    UserControl.txt.Text = pstrText
    mblnOverride = False
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mlngForeColor
End Property

Public Property Let ForeColor(ByVal poleColor As OLE_COLOR)
    mlngForeColor = poleColor
    DrawControl
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mlngBackColor
End Property

Public Property Let BackColor(ByVal poleColor As OLE_COLOR)
    mlngBackColor = poleColor
    DrawControl
    PropertyChanged "BackColor"
End Property

Public Property Get TimerInterval() As Long
    TimerInterval = UserControl.tmrSaveTimer.Interval
End Property

Public Property Let TimerInterval(ByVal plngInterval As Long)
    UserControl.tmrSaveTimer.Interval = plngInterval
    PropertyChanged "TimerInterval"
End Property

Private Function BorderColor() As Long
    If GetBrightest() < 129 Then
        BorderColor = RGB(128, 128, 128)
    Else
        BorderColor = vbBlack
    End If
End Function

Private Function GetBrightest() As Long
    Dim lngColor As Long
    Dim lngHighest As Long
    Dim lngCheck As Long
    
    If mlngBackColor < 0 Then OleTranslateColor mlngBackColor, 0, lngColor Else lngColor = mlngBackColor
    lngHighest = &HFF& And lngColor
    lngCheck = (&HFF00& And lngColor) \ 256
    If lngHighest < lngCheck Then lngHighest = lngCheck
    lngCheck = (&HFF0000 And lngColor) \ 65536
    If lngHighest < lngCheck Then lngHighest = lngCheck
    GetBrightest = lngHighest
End Function

Private Sub txt_Change()
    If Not mblnOverride Then DirtyFlag dfeNotes
End Sub

